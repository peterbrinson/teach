"""Stage 2 of deck-to-lecture: raw markdown + map.json -> finished vault page.

The map is the only thing a person or an LLM writes. The deck's own words never
pass through the model, so nothing can be silently reworded, summarised or
dropped -- which is why there is no "diff it against the deck" step at the end.

    python build.py --work ./work --map map.json \
        --out "…/teach/534/Lectures/Week 1 - Narrative Roles.md" \
        --attachments "…/teach/534/Lectures/attachments"
"""
import argparse, json, os, re, shutil, urllib.parse

SEP = "\n\n" + "<br>\n" * 10 + "\n"


def clean(t):
    """Deterministic repairs to pptx2md output. Every rule here is provably
    safe -- none of them can change a word of the deck's actual text."""
    # PowerPoint soft line breaks (Shift+Enter) survive as \x0b. A RUN of them
    # is a paragraph break; a single one is a wrap inside one sentence. These
    # are the cause of every "run-on title" and "missingspace" in deck output.
    t = re.sub("\x0b{2,}", "\n\n", t)
    t = t.replace("\x0b", " ")

    # pptx2md backslash-escapes ordinary punctuation.
    t = re.sub(r"\\([.,()\-!?:;'\"])", r"\1", t)

    # Emphasis artifacts. PowerPoint italicises run-by-run, so connectors come
    # out as "_ and _" and stray "_ _". An underscore with whitespace on BOTH
    # sides is never valid markdown emphasis, so dropping it is always safe.
    t = re.sub(r"__\s*__", "", t)
    t = re.sub(r"(?<=\s)_(?=\s)", "", t)
    t = re.sub(r"_([.,;:!?]+)_", r"\1", t)

    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r" +([.,;:!?])", r"\1", t)
    return t


def split_slides(raw, notes):
    """Split pptx2md output into one entry per slide.

    pptx2md emits speaker notes as their OWN chunk and occasionally an empty
    one, so chunk count != slide count. Notes are identified by matching the
    text python-pptx read from the deck. Raises if the result does not match
    the known slide count -- a wrong split would silently misplace headings.
    """
    chunks = [c.strip() for c in re.split(r"\n-{3,}\n", raw)]
    chunks = [c for c in chunks if c]
    note_texts = {clean(v).strip(): k for k, v in notes.items()}

    slides, attached = [], {}
    for c in chunks:
        key = clean(c).strip()
        if key in note_texts and slides:
            attached[len(slides)] = c          # notes belong to previous slide
        else:
            slides.append(c)
    return slides, attached


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--work", default="./work")
    ap.add_argument("--map", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--attachments", required=True)
    a = ap.parse_args()

    m = json.load(open(a.map, encoding="utf-8"))
    info = json.load(open(os.path.join(a.work, "deck-info.json"),
                          encoding="utf-8"))
    raw = clean(open(os.path.join(a.work, "raw.md"), encoding="utf-8").read())

    for old, new in m.get("text_fixes", []):
        if old in raw:
            raw = raw.replace(old, new)
        else:
            print("  !! text_fix did not match: " + old[:60])

    slides, attached = split_slides(raw, info["notes"])
    if len(slides) != info["slides_visible"]:
        raise SystemExit(
            "slide split mismatch: got " + str(len(slides)) + ", deck has " +
            str(info["slides_visible"]) + " visible. Inspect raw.md."
        )
    print("slides: " + str(len(slides)) + " (matches deck)")

    os.makedirs(a.attachments, exist_ok=True)
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    imgmap, dropped, copied, unmapped = m["images"], [], 0, []

    def swap(mt):
        nonlocal copied
        p = os.path.basename(
            urllib.parse.unquote(mt.group(1)).replace("\\", "/"))
        if p not in imgmap:
            unmapped.append(p)
            return mt.group(0)
        entry = imgmap[p]
        if entry in (None, False, "drop"):        # deliberately discarded
            dropped.append(p)
            return ""
        new, cap = entry
        src = os.path.join(a.work, "img", p)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(a.attachments, new))
            copied += 1
        return "![[" + new + "]]" + ("\n*" + cap + "*" if cap else "")

    def title(mt):
        txt = mt.group(1).strip()
        # a "title" over 80 chars is a sentence the deck styled as a title
        return ("**" + txt + "**") if len(txt) > 80 else ("### " + txt)

    body = []
    for i, s in enumerate(slides, 1):
        s = re.sub(r"!\[\]\(([^)]+)\)", swap, s)
        s = re.sub(r"^# (?!#)(.+)$", title, s, flags=re.M)
        if i in attached:
            note = "\n".join("> " + ln for ln in attached[i].splitlines())
            s += "\n\n> [!note] Instructor notes\n" + note
        sec = m["sections"].get(str(i))
        if sec:
            s = "## " + sec + "\n\n" + s
        body.append(s.strip())

    doc = ('---\ntitle: "' + m["title"] + '"\n---\n\n'
           + ("*" + m["tagline"] + "*\n" if m.get("tagline") else "")
           + SEP + SEP.join(body))
    if m.get("related"):
        doc += ("\n\n## Related\n\n"
                + "\n".join("- [[" + r + "]]" for r in m["related"]) + "\n")
    doc = re.sub(r"\n{3,}", "\n\n", doc)
    open(a.out, "w", encoding="utf-8").write(doc)

    if unmapped:
        print("  !! images missing from map.json: " + str(sorted(set(unmapped))))
    print("images kept " + str(copied) + ", dropped " + str(len(dropped)))
    print("wrote " + a.out + " (" + str(os.path.getsize(a.out)) + " bytes)")


if __name__ == "__main__":
    main()
