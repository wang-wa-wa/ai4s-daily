#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AI4S 小分子日报

从 arXiv 抓取最近几天「小分子 / 药物发现 + AI」方向的新论文，
生成中英结合的要点摘要（分类标签、中文一句话看点、英文关键摘录、链接）并发送到邮箱。
不发送论文全文，只发关键信息。

用法:
    python arxiv_ai4s_digest.py               正常运行并发送邮件
    python arxiv_ai4s_digest.py --dry-run      只生成 digest_preview.html 预览，不发送
    python arxiv_ai4s_digest.py --days 7       自定义回看天数

首次使用:
    1. 编辑同目录 config.json，填写 smtp.key（163 邮箱的 SMTP 授权码，不是登录密码）
    2. 可选: 把 llm.enabled 设为 true 并填写 api_key（智谱开放平台或任意 OpenAI 兼容接口），
       中文「一句话看点」会生成得更准；不填则使用内置的规则摘要

Windows 每日定时运行（PowerShell）:
    schtasks /create /tn "AI4S日报" /tr "python C:\\Users\\31517\\Downloads\\ai4s-digest\\arxiv_ai4s_digest.py" /sc daily /st 09:00

云端定时（GitHub Actions，推荐）:
    仓库含 .github/workflows/daily.yml，每天 01:10 UTC（北京 09:10）自动运行
    密钥一律通过仓库 Secrets 环境变量注入，不写进任何被提交的文件:
        AI4S_SMTP_KEY      163 邮箱 SMTP 授权码
        AI4S_LLM_API_KEY   智谱 API Key（可选）
        AI4S_LLM_MODEL     可选，覆盖 llm.model
        AI4S_RECIPIENT / AI4S_SMTP_USER / AI4S_DAYS  可选
    已发论文记录在 state.json（30 天内不重复发送），工作流会自动提交回仓库
"""

import argparse
import json
import os
import re
import smtplib
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ATOM = "{http://www.w3.org/2005/Atom}"
ARXIV_API = "http://export.arxiv.org/api/query"

DEFAULT_CONFIG = {
    "recipient": "rjwang22@163.com",
    "days_back": 3,
    "max_papers": 12,
    "fetch_per_query": 60,
    "send_when_empty": False,
    "smtp": {
        "host": "smtp.163.com",
        "port": 465,
        "user": "rjwang22@163.com",
        "key": "在这里填163邮箱的SMTP授权码"
    },
    "llm": {
        "enabled": False,
        "api_base": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
        "api_key": "",
        "model": "glm-5.3-flash"
    },
    "keywords": {
        "molecule": [
            "small molecule", "drug discovery", "drug design", "de novo design",
            "molecular generation", "molecule generation", "molecular property",
            "property prediction", "virtual screening", "docking", "retrosynthesis",
            "synthesis planning", "reaction prediction", "chemical space",
            "molecular dynamics", "molecular conformation", "scaffold", "SMILES",
            "molecular graph", "chemical reaction", "ligand", "binding affinity",
            "ADMET", "solubility", "toxicity"
        ],
        "ai": [
            "machine learning", "deep learning", "neural network",
            "artificial intelligence", "language model", "foundation model",
            "generative model", "graph neural", "reinforcement learning",
            "transformer", "diffusion model", "self-supervised", "contrastive learning",
            "active learning", "AI", "LLM", "GNN"
        ]
    }
}

KEYWORD_TAGS = [
    ("small molecule", "小分子 Small Molecule"),
    ("drug discovery", "药物发现 Drug Discovery"),
    ("drug design", "药物设计 Drug Design"),
    ("de novo", "从头设计 De Novo Design"),
    ("molecular generation", "分子生成 Molecule Generation"),
    ("molecule generation", "分子生成 Molecule Generation"),
    ("property prediction", "性质预测 Property Prediction"),
    ("virtual screening", "虚拟筛选 Virtual Screening"),
    ("docking", "分子对接 Docking"),
    ("retrosynthesis", "逆合成 Retrosynthesis"),
    ("synthesis planning", "合成规划 Synthesis Planning"),
    ("reaction prediction", "反应预测 Reaction Prediction"),
    ("chemical space", "化学空间 Chemical Space"),
    ("molecular dynamics", "分子动力学 MD"),
    ("conformation", "构象生成 Conformation"),
    ("SMILES", "SMILES 建模"),
    ("molecular graph", "分子图 Graph"),
    ("graph neural", "分子图 Graph"),
    ("scaffold", "骨架 Scaffold"),
    ("ligand", "配体 Ligand"),
    ("binding affinity", "亲和力预测 Binding Affinity"),
    ("admet", "ADMET 性质"),
    ("toxicity", "毒性预测 Toxicity"),
    ("solubility", "溶解度 Solubility"),
    ("foundation model", "基础模型 Foundation Model"),
    ("language model", "大语言模型 LLM"),
    ("diffusion", "扩散模型 Diffusion"),
    ("reinforcement learning", "强化学习 RL"),
    ("active learning", "主动学习 Active Learning"),
    ("interpretability", "可解释性 Interpretability"),
]

TASK_CN = [
    ("molecular generation", "分子生成"), ("molecule generation", "分子生成"),
    ("de novo", "分子设计"), ("property prediction", "分子性质预测"),
    ("virtual screening", "虚拟筛选"), ("docking", "分子对接"),
    ("retrosynthesis", "逆合成"), ("synthesis planning", "合成规划"),
    ("reaction prediction", "反应预测"), ("chemical space", "化学空间探索"),
    ("molecular dynamics", "分子动力学模拟"), ("conformation", "分子构象生成"),
    ("scaffold", "分子骨架设计"), ("binding affinity", "结合亲和力预测"),
    ("admet", "ADMET 性质预测"), ("toxicity", "毒性预测"),
    ("solubility", "溶解度预测"), ("drug discovery", "药物发现"),
    ("drug design", "药物设计"),
]

BENCHMARKS = ["QM9", "QM8", "PCQM4M", "MoleculeNet", "GuacaMol", "ZINC", "ChEMBL",
              "Tox21", "BBBP", "ESOL", "FreeSolv", "Lipophilicity", "USPTO",
              "PDBbind", "DUD-E", "MD17", "rMD17", "ChEBI", "BACE", "ClinTox",
              "SIDER", "HIV", "MUV", "ToxCast"]


def norm_ws(s):
    return re.sub(r"\s+", " ", s).strip()


def load_config(path):
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        print(f"[提示] 未找到配置文件，已生成默认配置: {path}")
        print("       请先填写 smtp.key（163 邮箱 SMTP 授权码），可先用 --dry-run 预览效果。")
        return None
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    merged = json.loads(json.dumps(DEFAULT_CONFIG))
    for k, v in cfg.items():
        if k == "smtp" or k == "llm" or k == "keywords":
            merged[k].update(v)
        else:
            merged[k] = v
    return merged


def apply_env_overrides(cfg):
    env = os.environ
    if env.get("AI4S_RECIPIENT"):
        cfg["recipient"] = env["AI4S_RECIPIENT"]
    if env.get("AI4S_SMTP_USER"):
        cfg["smtp"]["user"] = env["AI4S_SMTP_USER"]
    if env.get("AI4S_SMTP_KEY"):
        cfg["smtp"]["key"] = env["AI4S_SMTP_KEY"]
    if env.get("AI4S_LLM_API_KEY"):
        cfg["llm"]["api_key"] = env["AI4S_LLM_API_KEY"]
        cfg["llm"]["enabled"] = True
    if env.get("AI4S_LLM_MODEL"):
        cfg["llm"]["model"] = env["AI4S_LLM_MODEL"]
    if env.get("AI4S_DAYS"):
        try:
            cfg["days_back"] = int(env["AI4S_DAYS"])
        except ValueError:
            pass


STATE_FILE = os.path.join(SCRIPT_DIR, "state.json")


def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f).get("sent", {})
    except Exception:
        return {}


def save_state(sent):
    cutoff = (datetime.now(timezone.utc) - timedelta(days=30)).strftime("%Y-%m-%d")
    sent = {k: v for k, v in sent.items() if v >= cutoff}
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({"sent": sent}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[警告] 已发记录 state.json 写入失败: {e}")


def term_found(term, text):
    t = term.lower()
    if term.isupper() or len(t) <= 3:
        return re.search(r"\b" + re.escape(t) + r"\b", text) is not None
    return t in text


def fetch_papers(cfg, days):
    mol = cfg["keywords"]["molecule"]
    ai_part = " OR ".join('abs:"%s"' % k for k in cfg["keywords"]["ai"])
    max_results = int(cfg.get("fetch_per_query", 60))
    papers, seen = [], set()
    for i in range(0, len(mol), 10):
        chunk = mol[i:i + 10]
        mol_part = " OR ".join('abs:"%s"' % k for k in chunk)
        q = f"({mol_part}) AND ({ai_part})"
        url = ARXIV_API + "?" + urllib.parse.urlencode({
            "search_query": q, "sortBy": "submittedDate",
            "sortOrder": "descending", "max_results": str(max_results)})
        try:
            with urllib.request.urlopen(url, timeout=60) as r:
                root = ET.fromstring(r.read().decode("utf-8"))
        except Exception as e:
            print(f"[警告] arXiv 请求失败: {e}")
            continue
        for en in root.findall(ATOM + "entry"):
            try:
                raw_id = en.findtext(ATOM + "id") or ""
                base = re.sub(r"v\d+$", "", raw_id.split("/abs/")[-1])
                if not base or base in seen:
                    continue
                pub = datetime.strptime(en.findtext(ATOM + "published"),
                                        "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                if pub < datetime.now(timezone.utc) - timedelta(days=days):
                    continue
                seen.add(base)
                papers.append({
                    "arxiv_id": base,
                    "url": "https://arxiv.org/abs/" + base,
                    "title": norm_ws(en.findtext(ATOM + "title") or ""),
                    "abstract": norm_ws(en.findtext(ATOM + "summary") or ""),
                    "authors": [a.findtext(ATOM + "name") for a in en.findall(ATOM + "author")],
                    "published": pub,
                    "cats": [c.get("term") for c in en.findall(ATOM + "category")],
                })
            except Exception:
                continue
        time.sleep(3)
    return papers


def score_paper(p, mol_terms, ai_terms):
    title, abstract = p["title"].lower(), p["abstract"].lower()
    mol_hits = [t for t in mol_terms if term_found(t, title) or term_found(t, abstract)]
    ai_hits = [t for t in ai_terms if term_found(t, title) or term_found(t, abstract)]
    if not mol_hits or not ai_hits:
        return -1
    score = 0
    for t in mol_hits:
        score += 4 if term_found(t, title) else 1
    for t in ai_hits:
        score += 2 if term_found(t, title) else 1
    return min(score, 40)


def select_papers(papers, cfg):
    mol, ai = cfg["keywords"]["molecule"], cfg["keywords"]["ai"]
    scored = [(score_paper(p, mol, ai), p) for p in papers]
    scored = [(s, p) for s, p in scored if s > 0]
    scored.sort(key=lambda x: (-x[0], -x[1]["published"].timestamp()))
    limit = int(cfg.get("max_papers", 12))
    return [p for _, p in scored[:limit]], len(scored)


def build_tags(title, abstract):
    text = (title + " " + abstract).lower()
    tags = []
    for kw, tag in KEYWORD_TAGS:
        if term_found(kw, text) and tag not in tags:
            tags.append(tag)
    return tags[:4]


def heuristic_point(title, abstract):
    text = (title + " " + abstract).lower()
    m = re.search(r"\b(?:propose|present|introduce|develop|build|design)\s+([^.,;]{5,110})",
                  abstract, re.I)
    method = ""
    if m:
        method = re.sub(r"^(a|an|the)\s+", "", m.group(1).strip(), flags=re.I)[:90]
    tasks = [cn for kw, cn in TASK_CN if kw in text]
    bench = [b for b in BENCHMARKS
             if re.search(r"\b" + re.escape(b) + r"\b", title + " " + abstract, re.I)]
    s = "本文" + (f"提出 {method}" if method else "提出一种新方法")
    if tasks:
        s += "，用于" + "、".join(list(dict.fromkeys(tasks))[:3])
    if bench:
        s += "；在 " + ", ".join(bench[:3]) + " 等基准上验证"
    return s + "。"


def key_excerpt(abstract, max_sents=3):
    sents = [s.strip() for s in re.split(r"(?<=[.!?])\s+", abstract) if s.strip()]
    if not sents:
        return ""
    picked = [sents[0]]
    for pat in (r"\bwe\s+(?:propose|present|introduce|develop|build|design)\b",
                r"\b(?:outperform|state-of-the-art|SOTA|achiev\w*|improv\w*|superior)\b",
                r"\b(?:demonstrat\w*|show|experiment\w*)\b"):
        for s in sents:
            if re.search(pat, s, re.I) and s not in picked:
                picked.append(s)
                break
    out = []
    for s in picked[:max_sents]:
        out.append(s if len(s) <= 340 else s[:340].rstrip() + "...")
    return " ".join(out)


_LLM_DISABLED = False


def llm_summarize(llm_cfg, title, abstract):
    global _LLM_DISABLED
    if _LLM_DISABLED:
        return None, None
    prompt = (
        "你是 AI4S（小分子药物发现）领域的论文摘要助手。根据下面的标题与摘要，"
        "严格只输出一行 JSON，不要输出任何其他文字："
        '{"point": "中文一句话看点，60字以内，概括方法与用途，可保留 diffusion、GNN 等英文术语", '
        '"tags": ["中英文标签，如: 分子生成 Molecular Generation", "最多3个"]}\n'
        f"标题: {title}\n摘要: {abstract[:2500]}"
    )
    payload = json.dumps({
        "model": llm_cfg.get("model", "glm-5.3-flash"),
        "temperature": 0.2,
        "messages": [{"role": "user", "content": prompt}],
    }).encode("utf-8")
    req = urllib.request.Request(
        llm_cfg.get("api_base", "https://open.bigmodel.cn/api/paas/v4/chat/completions"),
        data=payload,
        headers={"Content-Type": "application/json",
                 "Authorization": "Bearer " + llm_cfg.get("api_key", "")})
    try:
        with urllib.request.urlopen(req, timeout=45) as r:
            data = json.loads(r.read().decode("utf-8"))
        content = data["choices"][0]["message"]["content"]
        m = re.search(r"\{.*\}", content, re.S)
        obj = json.loads(m.group(0))
        point = str(obj.get("point", "")).strip()
        tags = [str(t).strip() for t in obj.get("tags", []) if str(t).strip()][:4]
        return (point or None), (tags or None)
    except urllib.error.HTTPError as e:
        if e.code == 401:
            _LLM_DISABLED = True
            print("[LLM] 401 认证失败: llm.api_key 无效，已自动停用 LLM，本次改用本地规则摘要")
            print("       常见原因: ①把 163 邮箱 SMTP 授权码误填进 llm.api_key（它应填在 smtp.key）")
            print("                 ②Key 复制不完整或已被删除  ③用了其他平台但没改 llm.api_base / llm.model")
            print("       智谱 Key 获取: open.bigmodel.cn 登录 → 右上角「API Keys」→ 新建并完整复制")
        else:
            print(f"[LLM] 调用失败(HTTP {e.code})，回退到本地规则摘要: {e}")
        return None, None
    except Exception as e:
        print(f"[LLM] 调用失败，回退到本地规则摘要: {e}")
        return None, None


def build_entries(papers, cfg):
    llm_cfg = cfg.get("llm", {})
    use_llm = llm_cfg.get("enabled") and llm_cfg.get("api_key")
    entries = []
    for i, p in enumerate(papers, 1):
        point, tags = (None, None)
        if use_llm and not _LLM_DISABLED:
            point, tags = llm_summarize(llm_cfg, p["title"], p["abstract"])
            time.sleep(0.3)
        if not point:
            point = heuristic_point(p["title"], p["abstract"])
        if not tags:
            tags = build_tags(p["title"], p["abstract"])
        authors = ", ".join([a for a in p["authors"][:3] if a])
        if len(p["authors"]) > 3:
            authors += " et al."
        entries.append({
            "no": i, "arxiv_id": p["arxiv_id"], "title": p["title"],
            "authors": authors, "date": p["published"].strftime("%Y-%m-%d"),
            "cats": ", ".join(p["cats"][:4]), "tags": tags, "point": point,
            "excerpt": key_excerpt(p["abstract"]), "url": p["url"],
            "pdf": "https://arxiv.org/pdf/" + p["arxiv_id"],
        })
    return entries


HTML_HEAD = """<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
 body{font-family:'Segoe UI',Arial,'Microsoft YaHei',sans-serif;max-width:860px;
      margin:0 auto;padding:20px;color:#1a202c;line-height:1.65;background:#fff}
 h2{color:#2b6cb0;border-bottom:2px solid #2b6cb0;padding-bottom:8px;margin-bottom:6px}
 .summary{color:#718096;font-size:13px;margin-bottom:18px}
 .paper{border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin:16px 0}
 .paper h3{margin:0 0 4px;font-size:16px}
 .paper h3 a{color:#1a202c;text-decoration:none}
 .meta{color:#718096;font-size:13px;margin:2px 0 8px}
 .tags span{background:#ebf8ff;color:#2b6cb0;border-radius:4px;padding:1px 8px;
            margin-right:4px;font-size:12px;display:inline-block;margin-bottom:2px}
 .point{background:#f0fff4;border-left:3px solid #38a169;padding:7px 10px;
        margin:8px 0;border-radius:0 4px 4px 0;font-size:14px}
 .excerpt{color:#4a5568;font-size:14px;margin:8px 0}
 .links{font-size:13px}
 .links a{color:#2b6cb0}
</style></head><body>
"""


def build_html(entries, days, total_rel):
    gen = datetime.now().strftime("%Y-%m-%d %H:%M")
    parts = [HTML_HEAD, "<h2>AI4S 小分子日报</h2>",
             f'<p class="summary">检索到 {total_rel} 篇相关论文，精选 {len(entries)} 篇 · '
             f"回看 {days} 天 · 生成于 {gen}</p>"]
    for e in entries:
        tag_html = "".join(f"<span>{escape(t)}</span>" for t in e["tags"])
        parts.append(f"""
<div class="paper">
  <h3>{e['no']}. <a href="{e['url']}">{escape(e['title'])}</a></h3>
  <p class="meta">{e['date']} | {escape(e['authors'])} | {escape(e['cats'])}</p>
  <p class="tags">{tag_html}</p>
  <p class="point"><b>一句话看点</b>：{escape(e['point'])}</p>
  <p class="excerpt"><b>关键摘录 Key excerpt</b>：<br>{escape(e['excerpt'])}</p>
  <p class="links">摘要页: <a href="{e['url']}">{e['url']}</a>
     | PDF: <a href="{e['pdf']}">arxiv.org/pdf/{escape(e['arxiv_id'])}</a></p>
</div>""")
    parts.append("</body></html>")
    return "".join(parts)


def build_text(entries, days, total_rel):
    lines = [f"AI4S 小分子日报（精选 {len(entries)} 篇 / 相关 {total_rel} 篇，回看 {days} 天）", ""]
    for e in entries:
        lines += [
            f"{e['no']}. {e['title']}",
            f"   {e['date']} | {e['authors']} | {e['cats']}",
            f"   标签: {' | '.join(e['tags'])}",
            f"   看点: {e['point']}",
            f"   摘录: {e['excerpt']}",
            f"   {e['url']}", "",
        ]
    return "\n".join(lines)


def send_email(cfg, subject, html, text):
    s = cfg["smtp"]
    key = s.get("key", "")
    if not key or "在这里" in key:
        print("[错误] config.json 的 smtp.key 还没填写（163 邮箱 SMTP 授权码），本次未发送。")
        print("       获取方式: 网页版 163 邮箱 → 设置 → POP3/SMTP/IMAP → 开启服务 → 新增授权码")
        return False
    msg = MIMEMultipart("alternative")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = s.get("user", "")
    msg["To"] = cfg.get("recipient", "")
    msg.attach(MIMEText(text, "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP_SSL(s.get("host", "smtp.163.com"),
                              int(s.get("port", 465)), timeout=60) as smtp:
            smtp.login(s.get("user", ""), key)
            smtp.sendmail(msg["From"], [msg["To"]], msg.as_string())
        return True
    except Exception as e:
        print(f"[错误] 邮件发送失败: {e}")
        return False


def log_line(status, detail=""):
    try:
        with open(os.path.join(SCRIPT_DIR, "run.log"), "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\t{status}\t{detail}\n")
    except Exception:
        pass


def main():
    ap = argparse.ArgumentParser(description="AI4S 小分子 arXiv 日报")
    ap.add_argument("--dry-run", action="store_true", help="只预览不发送")
    ap.add_argument("--days", type=int, help="回看天数，覆盖配置")
    ap.add_argument("--config", default=os.path.join(SCRIPT_DIR, "config.json"))
    args = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    cfg = load_config(args.config)
    if not cfg:
        sys.exit(1)
    apply_env_overrides(cfg)
    days = args.days or int(cfg.get("days_back", 3))

    print(f"[1/3] 抓取 arXiv 最近 {days} 天的论文 ...")
    papers = fetch_papers(cfg, days)
    print(f"      共抓到 {len(papers)} 篇")

    selected, total_rel = select_papers(papers, cfg)
    print(f"[2/3] 相关性筛选: 相关 {total_rel} 篇，精选 {len(selected)} 篇")

    sent = load_state()
    fresh = [p for p in selected if p["arxiv_id"] not in sent]
    dup = len(selected) - len(fresh)
    if dup:
        print(f"      其中 {dup} 篇此前已发过，本次跳过（30 天内不重复）")
    selected = fresh

    if not selected:
        print("      近期没有符合条件的新论文，跳过发送。")
        log_line("ok", "no papers")
        if not cfg.get("send_when_empty") and not args.dry_run:
            return

    entries = build_entries(selected, cfg)
    today = datetime.now().strftime("%Y-%m-%d")
    subject = f"【AI4S小分子日报】{today} · {len(entries)} 篇新论文"
    html = build_html(entries, days, total_rel)
    text = build_text(entries, days, total_rel)

    preview = os.path.join(SCRIPT_DIR, "digest_preview.html")
    with open(preview, "w", encoding="utf-8") as f:
        f.write(html)

    if args.dry_run:
        print(f"[3/3] dry-run 预览已生成: {preview}")
        for e in entries:
            print(f"\n{e['no']}. {e['title']}")
            print(f"    {e['date']} | {e['authors']} | {e['cats']}")
            print(f"    标签: {' | '.join(e['tags'])}")
            print(f"    看点: {e['point']}")
            print(f"    摘录: {e['excerpt'][:200]}")
            print(f"    {e['url']}")
        log_line("dry-run", f"{len(entries)} papers")
        return

    print(f"[3/3] 发送邮件到 {cfg.get('recipient')} ...")
    ok = send_email(cfg, subject, html, text)
    log_line("sent" if ok else "failed", f"{len(entries)} papers")
    if ok:
        for e in entries:
            sent[e["arxiv_id"]] = today
        save_state(sent)
        print("      发送成功")
    else:
        print("      发送失败")
        sys.exit(1)


if __name__ == "__main__":
    main()
