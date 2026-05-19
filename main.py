# -*- coding: utf-8 -*-
"""物流服务商对账服务 (WF-C 引擎). 由 n8n 定时 HTTP 触发 /recon/run.
逻辑: 扫物流对账单任务台 触发对账=true -> 下载服务商Excel附件 -> openpyxl解析按运单号聚合
 -> join 物流对账明细(服务商运单号) -> 我方侧按配置公式(关联发货任务.总计+明细字段) 逐票算
 -> 与服务商侧逐票比 -> 回写汇总/差异/结果/快照+关联明细 -> 飞书卡片通知物流(DRY_RUN只发Frankie)."""
import os, io, json, time, urllib.request, urllib.error
from fastapi import FastAPI, Header, HTTPException
from openpyxl import load_workbook

app = FastAPI()

E = os.environ.get
APP1 = (E("FEISHU_APP1_ID", ""), E("FEISHU_APP1_SECRET", ""))
APP3 = (E("FEISHU_APP3_ID", ""), E("FEISHU_APP3_SECRET", ""))
BT = E("BITABLE_APP_TOKEN", "")
T_RECON = E("TBL_RECON", "")
T_DETAIL = E("TBL_DETAIL", "")
T_TASK = E("TBL_TASK", "")
T_CFG = E("TBL_CFG", "")
FRANKIE = E("FRANKIE_OID", "")
LOGI_DEPT = E("LOGI_DEPT_ID", "")
LOGI_JOB = E("LOGI_JOB", "物流仓储主管")
DRY_RUN = E("DRY_RUN", "1") == "1"
BEARER = E("BEARER", "")
FB = "https://open.feishu.cn"


def _req(url, method="GET", body=None, headers=None, raw=False, timeout=120):
    data = None
    if body is not None:
        data = body if isinstance(body, (bytes, bytearray)) else json.dumps(body).encode("utf-8")
    h = {} if raw else {"Content-Type": "application/json"}
    if headers:
        h.update(headers)
    rq = urllib.request.Request(url, data=data, headers=h, method=method)
    try:
        with urllib.request.urlopen(rq, timeout=timeout) as r:
            c = r.read()
            return c if raw else json.loads(c.decode("utf-8"))
    except urllib.error.HTTPError as e:
        c = e.read()
        return c if raw else json.loads(c.decode("utf-8"))


def tok(app):
    r = _req(FB + "/open-apis/auth/v3/tenant_access_token/internal", "POST",
             {"app_id": app[0], "app_secret": app[1]})
    return r.get("tenant_access_token", "")


def nz(v):
    if v is None:
        return ""
    if isinstance(v, list):
        return "".join((x.get("text", "") if isinstance(x, dict) else str(x)) for x in v)
    if isinstance(v, dict):
        return v.get("text", "")
    return v


def num(v):
    s = nz(v)
    try:
        return float(str(s).replace(",", "").strip() or 0)
    except Exception:
        return 0.0


def lk(v):
    if not v:
        return []
    if isinstance(v, list):
        o = []
        for x in v:
            if isinstance(x, str):
                o.append(x)
            elif isinstance(x, dict):
                if x.get("record_ids"):
                    o += x["record_ids"]
                elif x.get("id"):
                    o.append(x["id"])
        return o
    if isinstance(v, dict):
        return v.get("link_record_ids") or v.get("record_ids") or []
    return []


def list_records(t1, table):
    out, pt = [], ""
    while True:
        u = (FB + "/open-apis/bitable/v1/apps/%s/tables/%s/records?page_size=200" % (BT, table)
             + (("&page_token=" + pt) if pt else ""))
        r = _req(u, "GET", None, {"Authorization": "Bearer " + t1})
        d = r.get("data", {}) or {}
        for it in d.get("items", []):
            out.append({"rid": it["record_id"], "f": it.get("fields", {})})
        if d.get("has_more") and d.get("page_token"):
            pt = d["page_token"]
        else:
            break
    return out


def batch_get(t1, table, ids):
    if not ids:
        return {}
    r = _req(FB + "/open-apis/bitable/v1/apps/%s/tables/%s/records/batch_get" % (BT, table),
             "POST", {"record_ids": ids}, {"Authorization": "Bearer " + t1})
    res = {}
    for x in (r.get("data", {}) or {}).get("records", []):
        res[x["record_id"]] = x.get("fields", {})
    return res


def put_rec(t1, table, rid, fields):
    return _req(FB + "/open-apis/bitable/v1/apps/%s/tables/%s/records/%s" % (BT, table, rid),
                "PUT", {"fields": fields}, {"Authorization": "Bearer " + t1})


def resolve_logi(t1):
    """物流仓储部 LOGI_JOB 职务 在职 open_id (铁律:职务实时查,不硬编码)."""
    out, pt = [], ""
    while True:
        u = (FB + "/open-apis/contact/v3/users?department_id=%s&department_id_type=department_id"
             "&user_id_type=open_id&page_size=50" % LOGI_DEPT + (("&page_token=" + pt) if pt else ""))
        r = _req(u, "GET", None, {"Authorization": "Bearer " + t1})
        d = r.get("data", {}) or {}
        for usr in d.get("items", []):
            act = (usr.get("status") or {}).get("is_activated", True)
            if act and usr.get("job_title") == LOGI_JOB:
                out.append((usr.get("open_id"), usr.get("name")))
        if d.get("has_more") and d.get("page_token"):
            pt = d["page_token"]
        else:
            break
    return out


def send_card(t3, oid, title, lines, warn):
    els = [{"tag": "div", "text": {"tag": "lark_md",
            "content": "\n".join(lines)}}]
    card = {"config": {"wide_screen_mode": True},
            "header": {"template": "red" if warn else "turquoise",
                       "title": {"tag": "plain_text", "content": title}},
            "elements": els}
    return _req(FB + "/open-apis/im/v1/messages?receive_id_type=open_id", "POST",
                {"receive_id": oid, "msg_type": "interactive",
                 "content": json.dumps(card, ensure_ascii=False)},
                {"Authorization": "Bearer " + t3})


def download_media(t1, file_token):
    return _req(FB + "/open-apis/drive/v1/medias/%s/download" % file_token, "GET", None,
                {"Authorization": "Bearer " + t1}, raw=True)


def parse_supplier_excel(content, cfg):
    """按配置解析 xlsx, 返回 {运单号: {amount: sum折合人民币, fees:[(费用类型,金额)]}}."""
    wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    hr = int(cfg["hr"])
    ds = int(cfg["ds"])
    rows = list(ws.iter_rows(values_only=True))
    if hr - 1 >= len(rows):
        return {}, []
    header = ["" if c is None else str(c).strip() for c in rows[hr - 1]]
    jc = cfg["jcol"]
    ac = cfg["acol"]
    fc = cfg.get("fcol", "")
    ji = header.index(jc) if jc in header else -1
    ai = header.index(ac) if ac in header else -1
    fi = header.index(fc) if fc in header else -1
    if ji < 0 or ai < 0:
        return {}, ["表头缺列: join=%s amount=%s header=%s" % (jc, ac, header)]
    agg = {}
    warns = []
    for r in rows[ds - 1:]:
        if r is None:
            continue
        key = "" if ji >= len(r) or r[ji] is None else str(r[ji]).strip()
        if not key:
            continue
        amt = 0.0
        if ai < len(r) and r[ai] is not None:
            try:
                amt = float(str(r[ai]).replace(",", "").strip() or 0)
            except Exception:
                amt = 0.0
        ft = ""
        if 0 <= fi < len(r) and r[fi] is not None:
            ft = str(r[fi]).strip()
        a = agg.setdefault(key, {"amount": 0.0, "fees": []})
        a["amount"] += amt
        if ft:
            a["fees"].append((ft, amt))
    return agg, warns


@app.get("/health")
def health():
    return {"ok": True, "dry_run": DRY_RUN, "service": "recon-wf-c"}


@app.post("/recon/run")
def recon_run(authorization: str = Header(default="")):
    if BEARER and authorization != "Bearer " + BEARER:
        raise HTTPException(status_code=401, detail="unauthorized")
    t1 = tok(APP1)
    t3 = tok(APP3) if APP3[0] else t1
    if not t1:
        raise HTTPException(status_code=500, detail="feishu token fail")

    # 1. 服务商映射配置
    cfgmap = {}
    for it in list_records(t1, T_CFG):
        f = it["f"]
        sup = nz(f.get("服务商"))
        if not sup:
            continue
        en = f.get("启用")
        if en is False:
            continue
        cfgmap[sup] = {
            "hr": int(num(f.get("Excel表头行")) or 1),
            "ds": int(num(f.get("Excel数据起始行")) or 2),
            "jcol": nz(f.get("Excel_join列名")),
            "acol": nz(f.get("Excel_金额列名")),
            "fcol": nz(f.get("Excel_费用类型列名")),
            "myjoin": nz(f.get("我方join字段")),
            "formula": nz(f.get("我方汇总公式")),
            "thr": num(f.get("差异阈值百分比")) or 5.0,
        }

    # 2. 待对账的对账单 (触发对账=true)
    recons = [it for it in list_records(t1, T_RECON) if it["f"].get("触发对账") is True]
    logi = resolve_logi(t1)
    logi_oids = [o for o, _ in logi] or []
    logi_names = "、".join(n for _, n in logi) or "-"

    # 3. 预读全部物流对账明细 (含关联发货任务 -> 取 TASK 字段)
    details = list_records(t1, T_DETAIL)
    task_ids = []
    for d in details:
        task_ids += lk(d["f"].get("关联发货任务"))
    tasks = batch_get(t1, T_TASK, list(set(task_ids))) if task_ids else {}

    out = []
    for rec in recons:
        rid = rec["rid"]
        f = rec["f"]
        sup = nz(f.get("服务商"))
        c = cfgmap.get(sup)
        if not c:
            put_rec(t1, T_RECON, rid, {"对账明细快照": "无映射配置, 请先在『服务商对账映射配置』加 %s 并启用" % sup,
                                       "触发对账": False})
            put_rec(t1, T_RECON, rid, {"对账结果": "待对账"})
            out.append({"rid": rid, "sup": sup, "act": "no-config"})
            continue
        atts = f.get("服务商对账单附件") or []
        ft = atts[0].get("file_token") if atts and isinstance(atts[0], dict) else None
        if not ft:
            put_rec(t1, T_RECON, rid, {"对账明细快照": "未上传服务商对账单附件", "触发对账": False})
            out.append({"rid": rid, "sup": sup, "act": "no-attach"})
            continue
        content = download_media(t1, ft)
        try:
            agg, warns = parse_supplier_excel(content, c)
        except Exception as ex:
            put_rec(t1, T_RECON, rid, {"对账明细快照": "Excel解析失败: %s" % str(ex)[:200],
                                       "触发对账": False})
            out.append({"rid": rid, "sup": sup, "act": "parse-fail", "err": str(ex)[:200]})
            continue

        # 4. join 我方物流对账明细 (服务商运单号 == Excel join key)
        toks = [x.strip() for x in str(c["formula"]).split("+") if x.strip()]
        sup_total = sum(v["amount"] for v in agg.values())
        my_total = 0.0
        matched_dids = []
        snap = []
        n_diff = 0
        n_match = 0
        for wb_no, sv in agg.items():
            md = None
            for d in details:
                if nz(d["f"].get(c["myjoin"])).strip() == wb_no:
                    md = d
                    break
            if md is None:
                snap.append("%s 服务商¥%.2f 我方<无明细>" % (wb_no, sv["amount"]))
                continue
            n_match += 1
            matched_dids.append(md["rid"])
            myv = 0.0
            tf = {}
            tlinks = lk(md["f"].get("关联发货任务"))
            if tlinks:
                tf = tasks.get(tlinks[0], {})
            for tk in toks:
                if "." in tk:
                    sc, fld = tk.split(".", 1)
                    if sc == "TASK":
                        myv += num(tf.get(fld))
                    elif sc == "DETAIL":
                        myv += num(md["f"].get(fld))
            my_total += myv
            diff = myv - sv["amount"]
            pct = (abs(diff) / sv["amount"] * 100) if sv["amount"] else (100 if myv else 0)
            ok = pct <= c["thr"]
            if not ok:
                n_diff += 1
            snap.append("%s 服务商¥%.2f 我方¥%.2f 差%.2f(%.1f%%) %s"
                        % (wb_no, sv["amount"], myv, diff, pct, "平" if ok else "差异"))
            put_rec(t1, T_DETAIL, md["rid"],
                    {"对账状态": "对账平" if ok else "对账有差异"})

        result = "对账平" if (n_diff == 0 and n_match == len(agg) and n_match > 0) else (
            "有差异" if n_diff > 0 else "部分匹配")
        diff_total = my_total - sup_total
        pct_total = (abs(diff_total) / sup_total * 100) if sup_total else 0
        snaptxt = ("服务商账单¥%.2f / 我方汇总¥%.2f / 差异¥%.2f (%.1f%%) / 命中%d票 / 服务商总票%d\n"
                   % (sup_total, my_total, diff_total, pct_total, n_match, len(agg))
                   ) + "\n".join(snap[:60])
        put_rec(t1, T_RECON, rid, {
            "服务商账单金额": round(sup_total, 2),
            "我方汇总金额": round(my_total, 2),
            "命中票数": n_match,
            "差异金额": round(diff_total, 2),
            "差异百分比": round(pct_total, 2),
            "对账明细快照": snaptxt[:9000],
            "关联物流对账明细": matched_dids,
            "已对账": True,
            "触发对账": False,
        })
        put_rec(t1, T_RECON, rid, {"对账结果": result})

        warn = (result != "对账平")
        period = ""
        try:
            ps, pe = f.get("对账周期起"), f.get("对账周期止")
            period = "%s ~ %s" % (
                time.strftime("%Y-%m-%d", time.gmtime(ps / 1000)) if ps else "?",
                time.strftime("%Y-%m-%d", time.gmtime(pe / 1000)) if pe else "?")
        except Exception:
            period = "-"
        lines = [
            "**服务商**：%s" % sup,
            "**对账周期**：%s" % period,
            "**服务商账单**：¥%.2f" % sup_total,
            "**我方汇总**：¥%.2f（命中 %d / 服务商 %d 票）" % (my_total, n_match, len(agg)),
            "**差异**：¥%.2f（%.1f%%）" % (diff_total, pct_total),
            "**结果**：%s" % result,
            "差异票/快照见对账单『对账明细快照』，请核对后改『处理状态』。",
        ]
        if DRY_RUN:
            lines = ["⚠ **DRY-RUN** 真实应发：物流仓储主管（%s）" % logi_names] + lines
            targets = [FRANKIE]
        else:
            targets = logi_oids or [FRANKIE]
        mids = []
        for o in targets:
            # 用聪哥1号发: resolve_logi/FRANKIE 的 open_id 均属聪哥1号命名空间
            # (聪哥3号发会 99992361 open_id cross app)
            rr = send_card(t1, o, "物流对账 · %s %s" % (sup, result), lines, warn)
            mids.append((rr.get("data", {}) or {}).get("message_id")
                        or ("ERR:" + json.dumps(rr, ensure_ascii=False)[:120]))
        out.append({"rid": rid, "sup": sup, "result": result, "sup_total": round(sup_total, 2),
                    "my_total": round(my_total, 2), "matched": n_match, "diff_rows": n_diff,
                    "mids": mids})

    return {"dry_run": DRY_RUN, "processed": len(out), "detail": out}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", "8080")))
