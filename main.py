import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

import docx
import requests
from dotenv import load_dotenv

INPUT_DIR = Path("Input")
OUTPUT_DIR = Path("output")
PROMPT_FILE = Path("skill/prompt.txt")

MODEL = "deepseek-chat"
BASE_URL = "https://api.deepseek.com/chat/completions"
BATCH_SIZE = 8
TEMPERATURE = 1.2
TIMEOUT = 90
MAX_RETRIES = 3

# Action verbs by UK RQF / FHEQ level. A curriculum at level N should open its
# learning outcomes and assessment criteria with a verb from bank[N].
VERB_BANKS = {
    3: {
        "apply", "calculate", "carryout", "classify", "combine", "compare",
        "complete", "compute", "construct", "convert", "demonstrate", "describe",
        "design", "determine", "develop", "display", "employ", "estimate",
        "explain", "extract", "identify", "illustrate", "interpret", "list",
        "locate", "maintain", "measure", "name", "operate", "organise",
        "organize", "outline", "perform", "prepare", "present", "produce",
        "record", "select", "show", "state", "summarise", "summarize", "use",
        "utilise", "utilize", "calculate", "post", "enter", "process", "set",
        "setup", "draw", "plot", "write", "drawup",
    },
    4: {
        "analyse", "analyze", "apply", "appraise", "assess", "calculate",
        "categorise", "categorize", "classify", "combine", "compare", "compile",
        "conduct", "construct", "contrast", "demonstrate", "describe", "design",
        "determine", "develop", "diagnose", "differentiate", "distinguish",
        "employ", "estimate", "evaluate", "examine", "explain", "explore",
        "extract", "formulate", "generate", "identify", "illustrate",
        "implement", "interpret", "investigate", "justify", "maintain",
        "measure", "operate", "organise", "organize", "perform", "plan",
        "prepare", "present", "produce", "record", "report", "review",
        "select", "solve", "summarise", "summarize", "test", "use", "utilise",
        "utilize", "verify", "write", "configure", "integrate", "monitor",
        "predict", "process", "reconcile", "trace", "setup", "carryout",
        "post", "enter", "allocate", "apportion", "absorb", "budget", "cost",
        "audit", "forecast", "record", "measure", "set",
    },
    5: {
        "analyse", "analyze", "appraise", "assess", "compare", "construct",
        "debate", "demonstrate", "design", "determine", "develop", "devise",
        "differentiate", "discriminate", "distinguish", "employ", "evaluate",
        "examine", "explore", "formulate", "generate", "identify", "implement",
        "initiate", "interpret", "investigate", "judge", "justify", "operate",
        "perform", "plan", "present", "prioritise", "prioritize", "produce",
        "propose", "question", "report", "review", "scrutinise", "scrutinize",
        "solve", "summarise", "summarize", "synthesise", "synthesize",
        "translate", "validate", "weigh", "apply", "articulate", "communicate",
        "classify", "undertake", "compose", "prepare", "calculate",
        "allocate", "apportion", "absorb", "reconcile", "compute", "record",
        "process", "measure", "cost", "budget", "forecast", "audit",
        "recommend", "organise", "organize", "conduct", "establish",
        "configure", "enter", "post", "integrate", "select", "choose",
        "explain", "describe", "contrast", "critique", "test", "verify",
        "trace", "use", "utilise", "utilize", "spreadsheet", "predict",
        "estimate", "analysis", "set", "maintain", "monitor", "control",
        "setup", "carryout", "takeon", "makeuse", "justify", "critically",
    },
    7: {
        "analyse", "analyze", "appraise", "argue", "assess", "challenge",
        "conceptualise", "conceptualize", "construct", "contextualise",
        "contextualize", "create", "criticallyanalyse", "criticallyevaluate",
        "criticallyexamine", "criticallyreview", "critique", "debate",
        "deconstruct", "demonstrate", "design", "determine", "develop",
        "devise", "differentiate", "discriminate", "evaluate", "examine",
        "explore", "formulate", "generate", "hypothesise", "hypothesize",
        "initiate", "innovate", "integrate", "interrogate", "investigate",
        "judge", "justify", "model", "negotiate", "originate", "plan",
        "postulate", "predict", "problematise", "problematize", "produce",
        "propose", "recommend", "reconcile", "reflect", "reframe", "review",
        "scrutinise", "scrutinize", "speculate", "substantiate", "synthesise",
        "synthesize", "theorise", "theorize", "validate", "weigh",
    },
}

LEVEL_VERBS = VERB_BANKS[5]


def detect_level(filename):
    """Pick the RQF level from the file name, e.g. 'Level 4 Diploma ...' -> 4."""
    name = Path(filename).name.lower()
    matches = re.findall(r"\blevel[\s\-_]*([3457])\b", name)
    if matches:
        return int(matches[-1])
    matches = re.findall(r"\b([3457])[a-z]{0,3}\s*(?:extended\s+)?(?:diploma|certificate|award|dip)\b", name)
    if matches:
        return int(matches[0])
    return 5


def level_verbs(level):
    return VERB_BANKS.get(level, VERB_BANKS[5])

BANNED_WORDS = {
    "delve", "leverage", "robust", "holistic", "seamless", "tapestry", "realm",
    "foster", "underscore", "pivotal", "multifaceted", "myriad", "nuanced",
    "crucial", "comprehensive", "meticulous", "intricate",
}

BANNED_PHRASES = [
    "it is important to note", "it is worth noting", "plays a crucial role",
    "plays a vital role", "in today's rapidly evolving", "in today's rapidly changing",
    "not only", "a wide range of",
]


def read_prompts():
    text = PROMPT_FILE.read_text(encoding="utf-8")
    aim = text.split("[AIM_SYSTEM]", 1)[1].split("[TITLE_SYSTEM]", 1)[0].strip()
    title = text.split("[TITLE_SYSTEM]", 1)[1].split("[LO_SYSTEM]", 1)[0].strip()
    lo = text.split("[LO_SYSTEM]", 1)[1].strip()
    return aim, title, lo


def apply_level(template, level):
    verbs = sorted(level_verbs(level))
    return (template
            .replace("{LEVEL}", str(level))
            .replace("{VERBS}", ", ".join(verbs)))


def chat(system, user, api_key):
    body = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "temperature": TEMPERATURE,
        "max_tokens": 2048,
        "response_format": {"type": "json_object"},
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    delay = 2
    last = None
    for _ in range(MAX_RETRIES):
        try:
            r = requests.post(BASE_URL, headers=headers, json=body, timeout=TIMEOUT)
            if r.status_code == 429 or r.status_code >= 500:
                last = f"HTTP {r.status_code}: {r.text[:200]}"
                time.sleep(delay)
                delay *= 2.5
                continue
            r.raise_for_status()
            return r.json()["choices"][0]["message"]["content"]
        except (requests.Timeout, requests.ConnectionError) as e:
            last = str(e)
            time.sleep(delay)
            delay *= 2.5
        except requests.HTTPError as e:
            last = f"HTTPError: {e}"
            break
    raise RuntimeError(last or "unknown API error")


def is_unit_table(table):
    joined = " ".join(c.text for row in table.rows for c in row.cells)
    return "Learning Outcomes" in joined and "Assessment Criteria" in joined


def unique_cells(cells):
    seen, out = set(), []
    for c in cells:
        key = id(c._tc)
        if key not in seen:
            seen.add(key)
            out.append(c)
    return out


def collect_targets(table):
    rows = table.rows
    targets = []
    if len(rows) < 6:
        return targets

    name_cell = rows[1].cells[1] if len(rows[1].cells) > 1 else None
    if name_cell is not None and name_cell.text.strip():
        targets.append({"kind": "title", "cell": name_cell, "text": name_cell.text.strip()})

    aim_cell = rows[2].cells[0]
    if aim_cell.text.strip():
        targets.append({"kind": "aim", "cell": aim_cell, "text": aim_cell.text.strip()})

    body = rows[4:]
    for cell in unique_cells([r.cells[0] for r in body if len(r.cells) > 0]):
        txt = cell.text.strip()
        if txt:
            targets.append({"kind": "lo", "cell": cell, "text": txt})

    for row in body:
        if len(row.cells) > 3:
            cell = row.cells[3]
            txt = cell.text.strip()
            if txt:
                targets.append({"kind": "ac", "cell": cell, "text": txt})

    return targets


def digits(text):
    return re.findall(r"\d+", text)


def first_verb(text):
    stripped = re.sub(r"^\s*\d+(\.\d+)*\.?\s*", "", text).strip()
    m = re.match(r"([A-Za-z]+)", stripped)
    return m.group(1).lower() if m else ""


def guard(original, rewritten, kind, level=5):
    if not rewritten:
        return original
    new = rewritten.strip()
    old = original.strip()
    if new.lower() == old.lower():
        return original
    if "\u2014" in new:
        return original
    if any(d not in new for d in digits(old)):
        return original
    low = new.lower()
    if any(w in low.split() for w in BANNED_WORDS):
        return original
    if any(p in low for p in BANNED_PHRASES):
        return original

    if kind == "title":
        if len(new.split()) > 7:
            return original
        if len(new) > 55:
            return original
        return new

    ratio = len(new.split()) / max(1, len(old.split()))
    if ratio < 0.70 or ratio > 1.60:
        return original
    if kind in ("lo", "ac"):
        verb = first_verb(new)
        if verb in ("set", "carry", "take", "make", "draw"):
            m = re.match(r"(?:\d+(?:\.\d+)*\.?\s*)?([A-Za-z]+(?:\s+[a-z]+)?)", new.strip())
            verb = "".join(m.group(1).lower().split()) if m else verb
        bank = level_verbs(level)
        if verb not in bank:
            return original
    return new


def humanize(items, system, api_key, log):
    out = {}
    for start in range(0, len(items), BATCH_SIZE):
        chunk = items[start:start + BATCH_SIZE]
        payload = {"items": [{"id": i, "text": t} for i, t in chunk]}
        try:
            raw = chat(system, json.dumps(payload, ensure_ascii=False), api_key)
            results = json.loads(raw).get("results", [])
            if len(results) != len(chunk):
                raise ValueError("item count mismatch")
            for item in results:
                out[int(item["id"])] = item["text"]
        except Exception as e:
            log.append(f"batch {start}-{start + len(chunk) - 1} failed ({e}); retrying individually")
            for i, t in chunk:
                one = json.dumps({"items": [{"id": i, "text": t}]}, ensure_ascii=False)
                try:
                    raw = chat(system, one, api_key)
                    out[i] = json.loads(raw)["results"][0]["text"]
                except Exception as e2:
                    log.append(f"item {i} failed ({e2}); keeping original")
    return out


def apply_text(target, new_text):
    cell = target["cell"]
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    para = cell.paragraphs[0]
    runs = para.runs
    if runs:
        runs[0].text = new_text
        for r in runs[1:]:
            r.text = ""
    else:
        para.add_run(new_text)


def process(src, dest, aim_system, title_system, lo_system, api_key, dry_run, limit, level):
    print(f"\n=== {src.name} ===")
    print(f"RQF level detected: {level}")
    doc = docx.Document(str(src))
    tables = [t for t in doc.tables if is_unit_table(t)]
    print(f"unit tables found: {len(tables)}")

    plan = []
    for ti, table in enumerate(tables):
        targets = collect_targets(table)
        counts = {k: sum(1 for x in targets if x["kind"] == k) for k in ("title", "aim", "lo", "ac")}
        print(f"  unit {ti + 1}: {len(targets)} targets "
              f"({counts['title']} title, {counts['aim']} aim, {counts['lo']} lo, {counts['ac']} ac)")
        for t in targets:
            plan.append((ti + 1, t))

    if limit:
        plan = plan[:limit]

    if dry_run:
        for unit, t in plan:
            print(f"  [unit {unit} {t['kind']}] {t['text'][:90]}")
        return

    log = []
    results = {}
    for kind, system in (("title", title_system), ("aim", aim_system),
                         ("lo", lo_system), ("ac", lo_system)):
        batch_items = [(n, t["text"]) for n, (_, t) in enumerate(plan) if t["kind"] == kind]
        if not batch_items:
            print(f"  WARNING: no {kind} targets found")
            continue
        print(f"  rewriting {len(batch_items)} {kind} item(s)...")
        got = humanize(batch_items, system, api_key, log)
        if kind == "title":
            unchanged = [(n, txt) for n, txt in batch_items
                         if n in got and got[n].strip().lower() == txt.strip().lower()]
            if unchanged:
                print(f"  retrying {len(unchanged)} unchanged title(s)...")
                retry = humanize(unchanged, system, api_key, log)
                for n, txt in retry.items():
                    if retry[n].strip().lower() != dict(batch_items)[n].strip().lower():
                        got[n] = retry[n]
        results.update(got)

    rows_log = []
    kept = 0
    for n, (unit, t) in enumerate(plan):
        if n not in results:
            kept += 1
            continue
        final = guard(t["text"], results[n], t["kind"], level)
        if final == t["text"]:
            if results[n].strip() != t["text"].strip():
                log.append(f"guard rejected: unit {unit} {t['kind']}: {results[n][:160]}")
            kept += 1
            continue
        rows_log.append((unit, t["kind"], t["text"], final))
        apply_text(t, final)

    doc.save(str(dest))

    log_path = dest.with_suffix(".log.txt")
    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(log))
        fh.write("\n\n")
        for unit, kind, before, after in rows_log:
            fh.write(f"unit {unit} | {kind}\nBEFORE: {before}\nAFTER : {after}\n\n")

    print(f"  rewritten: {len(rows_log)} | kept original: {kept}")
    print(f"  saved: {dest}")
    print(f"  log:   {log_path}")


def main():
    parser = argparse.ArgumentParser(description="De-AI a curriculum .docx (UK RQF Level 5).")
    parser.add_argument("--file", help="single file name inside Input/")
    parser.add_argument("--dry-run", action="store_true", help="list targets, no API calls")
    parser.add_argument("--limit", type=int, help="process only the first N items")
    parser.add_argument("--overwrite", action="store_true", help="replace files in output/")
    parser.add_argument("--level", type=int, choices=sorted(VERB_BANKS),
                        help="override the RQF level (default: read from the file name)")
    args = parser.parse_args()

    load_dotenv()
    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not args.dry_run:
        if not api_key or api_key == "paste_your_key_here":
            print("ERROR: set a real DEEPSEEK_API_KEY in .env (project root).")
            return 1

    OUTPUT_DIR.mkdir(exist_ok=True)

    if args.file:
        files = [INPUT_DIR / args.file]
    else:
        files = sorted(p for p in INPUT_DIR.glob("*.docx") if not p.name.startswith("~$"))

    if not files:
        print("No .docx files found in Input/")
        return 1

    aim_tpl, title_tpl, lo_tpl = read_prompts()

    for src in files:
        if not src.exists():
            print(f"Missing: {src}")
            continue
        dest = OUTPUT_DIR / src.name
        if dest.exists() and not args.overwrite and not args.dry_run:
            print(f"Skipping {src.name}: {dest} exists (use --overwrite)")
            continue

        level = args.level or detect_level(src.name)
        if level not in VERB_BANKS:
            print(f"WARNING: unsupported level {level} for {src.name}; using 5")
            level = 5

        aim_system = apply_level(aim_tpl, level)
        title_system = apply_level(title_tpl, level)
        lo_system = apply_level(lo_tpl, level)

        process(src, dest, aim_system, title_system, lo_system, api_key,
                args.dry_run, args.limit, level)

    return 0


if __name__ == "__main__":
    sys.exit(main())
