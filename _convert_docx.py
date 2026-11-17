# -*- coding: utf-8 -*-
"""Convert 《AI组织项目管理升级》docx → HTML book site."""
from __future__ import annotations

import html
import re
import shutil
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

DOCX = Path("/Users/jj/大模型训练/项目管理/AI 组织中项目管理升级.docx")
OUT = Path("/Users/jj/大模型训练/项目管理")
IMG = OUT / "images"
IMG.mkdir(parents=True, exist_ok=True)

STYLES_SRC = Path("/Users/jj/大模型训练/学习笔记/styles.css")

BOOK_TITLE = "AI组织项目管理升级"
BOOK_SUBTITLE = "人机协同时代的交付、协同与治理"
BOOK_AUTHOR = "赵登琴"
BOOK_AUTHOR_BIO = (
    "赵登琴 · 希望学产研团队项目经理专家，清华大学本科。"
    "聚焦 AI 组织中的项目战略、交付节奏与人机协同治理。"
)
BRAND_MARK = "PM"

NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

PAGE_CSS_EXTRA = """
:root {
  --ha-bg: #f5f7fa;
  --ha-sidebar: #f5f7fa;
  --ha-ink: #0f1c2e;
  --ha-muted: #6b7a8d;
  --ha-line: #dde3eb;
  --ha-active-bg: #e8eef5;
  --ha-accent: #f5a623;
  --ha-accent-ink: #1c1c1a;
}
body.hermes-book {
  background: var(--ha-bg) !important;
  color: var(--ha-ink);
  font-size: 17px;
  line-height: 1.75;
}
body.hermes-book a { color: #9a5b00; }
body.hermes-book a:hover { color: var(--ha-ink); }
.ha-shell {
  display: grid;
  grid-template-columns: 270px minmax(0, 1fr);
  min-height: 100vh;
}
.ha-sidebar {
  position: sticky;
  top: 0;
  align-self: start;
  height: 100vh;
  overflow-y: auto;
  padding: 1.25rem 1rem 1.5rem;
  border-right: 1px solid var(--ha-line);
  background: var(--ha-sidebar);
  display: flex;
  flex-direction: column;
  gap: 0.85rem;
}
.ha-back {
  display: inline-flex;
  align-items: center;
  gap: 0.35rem;
  color: var(--ha-muted) !important;
  text-decoration: none;
  font-size: 0.88rem;
}
.ha-back:hover { color: var(--ha-ink) !important; }
.ha-toc-label {
  font-size: 0.78rem;
  font-weight: 600;
  color: var(--ha-muted);
  letter-spacing: 0.06em;
  margin: 0.35rem 0 0.15rem;
}
.ha-chapters {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 0.15rem;
}
.ha-chapters > li {
  border-radius: 10px;
}
.ha-chapters .chap-toggle {
  display: flex;
  align-items: flex-start;
  gap: 0.35rem;
  width: 100%;
  padding: 0.55rem 0.7rem;
  border: 0;
  border-radius: 10px;
  border-left: 3px solid transparent;
  background: transparent;
  color: var(--ha-muted);
  font: inherit;
  font-size: 0.9rem;
  line-height: 1.35;
  text-align: left;
  cursor: pointer;
}
.ha-chapters .chap-toggle:hover {
  background: rgba(15,28,46,0.04);
  color: var(--ha-ink);
}
.ha-chapters > li.open > .chap-toggle,
.ha-chapters > li.active > .chap-toggle {
  background: var(--ha-active-bg);
  color: var(--ha-ink);
  font-weight: 600;
  border-left-color: var(--ha-accent);
}
.ha-chapters .chap-caret {
  flex: 0 0 auto;
  width: 0.9rem;
  margin-top: 0.15rem;
  color: inherit;
  opacity: 0.7;
  transition: transform 0.15s ease;
}
.ha-chapters > li.open > .chap-toggle .chap-caret {
  transform: rotate(90deg);
}
.ha-chapters .chap-subs {
  list-style: none;
  margin: 0 0 0.35rem;
  padding: 0 0.4rem 0.25rem 1.35rem;
  display: none;
}
.ha-chapters > li.open > .chap-subs {
  display: grid;
  gap: 0.12rem;
}
.ha-chapters .chap-subs a {
  display: block;
  padding: 0.32rem 0.55rem;
  border-radius: 8px;
  text-decoration: none;
  color: var(--ha-muted);
  font-size: 0.82rem;
  line-height: 1.35;
}
.ha-chapters .chap-subs a:hover {
  background: rgba(15,28,46,0.04);
  color: var(--ha-ink);
}
.ha-chapters .chap-subs a.active {
  color: var(--ha-ink);
  font-weight: 600;
  background: rgba(31,111,235,0.08);
}
.ha-side-foot {
  margin-top: auto;
  padding-top: 1rem;
  border-top: 1px solid var(--ha-line);
  display: flex;
  flex-direction: column;
  gap: 0.55rem;
}
.ha-side-foot a {
  color: var(--ha-muted);
  text-decoration: none;
  font-size: 0.88rem;
}
.ha-side-foot a:hover { color: var(--ha-ink); }
.ha-main {
  min-width: 0;
  display: flex;
  flex-direction: column;
}
.ha-topbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1rem;
  padding: 0.95rem 2rem;
  border-bottom: 1px solid var(--ha-line);
  background: rgba(245,247,250,0.92);
  backdrop-filter: blur(8px);
  position: sticky;
  top: 0;
  z-index: 5;
}
.ha-brand {
  display: inline-flex;
  align-items: center;
  gap: 0.55rem;
  text-decoration: none;
  color: var(--ha-ink);
  font-weight: 650;
  font-size: 0.98rem;
  letter-spacing: -0.02em;
}
.ha-brand-mark {
  width: 1.55rem;
  height: 1.55rem;
  border-radius: 6px;
  background: #0f1c2e;
  color: #fff;
  display: grid;
  place-items: center;
  font-size: 0.62rem;
  font-weight: 700;
  font-family: var(--font-mono);
}
.ha-top-links {
  display: flex;
  gap: 1.1rem;
  flex-wrap: wrap;
}
.ha-top-links a {
  color: var(--ha-muted);
  text-decoration: none;
  font-size: 0.9rem;
}
.ha-top-links a:hover { color: var(--ha-ink); }
.ha-article {
  width: min(760px, calc(100% - 3rem));
  margin: 0 auto;
  padding: 1.75rem 0 4rem;
}
.ha-meta {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 1rem;
  color: var(--ha-muted);
  font-size: 0.88rem;
  margin-bottom: 1.25rem;
}
.ha-meta a { color: inherit; text-decoration: none; }
.ha-meta a:hover { color: var(--ha-ink); }
.ha-chapter-head {
  display: flex;
  align-items: flex-start;
  gap: 0.9rem;
  margin-bottom: 1.5rem;
}
.ha-chapter-num {
  flex: 0 0 auto;
  min-width: 2.4rem;
  height: 2.4rem;
  padding: 0 0.55rem;
  border-radius: 8px;
  background: var(--ha-accent);
  color: var(--ha-accent-ink);
  display: grid;
  place-items: center;
  font-family: var(--font-mono);
  font-weight: 700;
  font-size: 0.95rem;
}
.ha-chapter-head h1 {
  margin: 0;
  font-size: clamp(1.55rem, 3vw, 2.15rem);
  border: none;
  padding: 0;
  letter-spacing: -0.03em;
}
.ha-section-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 0.05rem;
}
.ha-section-list a {
  display: block;
  padding: 0.35rem 0.55rem;
  border-radius: 8px;
  text-decoration: none;
  color: var(--ha-muted);
  font-size: 0.82rem;
  line-height: 1.35;
}
.ha-section-list a:hover {
  background: rgba(15,28,46,0.04);
  color: var(--ha-ink);
}
.ha-content h2 {
  border-top: none;
  padding-top: 0.2rem;
  margin-top: 2.2rem;
}
.ha-content h2::before {
  content: "";
  display: block;
  width: 2rem;
  height: 3px;
  background: var(--ha-accent);
  margin-bottom: 0.75rem;
  border-radius: 2px;
}
.ha-content p { margin: 0 0 1rem; text-indent: 2em; }
.ha-sec-box {
  margin: 0 0 1.75rem;
  padding: 1rem 1.15rem 1.05rem;
  border-radius: 12px;
  background: #fff;
  border: 1px solid var(--ha-line);
}
.ha-sec-box .label {
  margin: 0 0 0.55rem;
  font-size: 0.82rem;
  font-weight: 700;
  color: var(--ha-muted);
  letter-spacing: 0.04em;
}
.ha-sec-box ul {
  list-style: none;
  margin: 0;
  padding: 0;
  display: grid;
  gap: 0.35rem;
}
.ha-sec-box a {
  text-decoration: none;
  color: var(--ha-ink);
  font-size: 0.95rem;
}
.ha-sec-box a:hover { color: #1554b5; }
.book-block {
  margin: 0 0 1rem;
  padding: 1rem 1.1rem 1.05rem;
  border-radius: 12px;
  background: #fff;
  border: 1px solid var(--ha-line);
}
.book-block h2 {
  margin: 0 0 0.55rem;
  font-size: 1.05rem;
  border: none;
  padding: 0;
  display: flex;
  gap: 0.65rem;
  align-items: baseline;
}
.book-block h2 a { text-decoration: none; color: var(--ha-ink); }
.book-block h2 .num {
  display: inline-grid;
  place-items: center;
  width: 1.85rem;
  height: 1.85rem;
  border-radius: 7px;
  background: var(--ha-accent);
  color: var(--ha-accent-ink);
  font-family: var(--font-mono);
  font-size: 0.82rem;
  font-weight: 700;
}
.book-block ul.secs {
  list-style: none;
  margin: 0;
  padding: 0;
  color: #3a4652;
  font-size: 0.92rem;
  display: grid;
  gap: 0.28rem;
}
.book-block ul.secs a { text-decoration: none; color: #3a4652; }
.book-block ul.secs a:hover { color: #1554b5; }
.figcap { color: var(--ha-muted); font-size: 0.9rem; text-align: center; margin: 0.25rem 0 1rem; }
.table-wrap { overflow-x: auto; margin: 1rem 0; }
.table-wrap table { width: 100%; border-collapse: collapse; font-size: 0.92rem; }
.table-wrap th, .table-wrap td { border: 1px solid var(--ha-line); padding: 0.45rem 0.6rem; text-align: left; vertical-align: top; }
.table-wrap th { background: #e8eef5; }
.figure img { max-width: 100%; height: auto; border-radius: 10px; border: 1px solid var(--ha-line); }
.ha-footer-nav {
  display: flex;
  justify-content: space-between;
  gap: 1rem;
  margin-top: 2.5rem;
  padding-top: 1.25rem;
  border-top: 1px solid var(--ha-line);
}
.ha-footer-nav a {
  text-decoration: none;
  color: var(--ha-ink);
  font-weight: 600;
}
.ha-index-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 0.45rem;
}
.ha-index-list a {
  display: grid;
  grid-template-columns: 2.6rem 1fr;
  gap: 0.85rem;
  align-items: center;
  padding: 0.95rem 1rem;
  border-radius: 12px;
  text-decoration: none;
  color: var(--ha-ink);
  background: #fff;
  border: 1px solid var(--ha-line);
}
.ha-index-list a:hover { border-color: #c5d0de; background: #f8fafc; }
.ha-index-list .num {
  width: 2.4rem;
  height: 2.4rem;
  border-radius: 8px;
  background: var(--ha-accent);
  color: var(--ha-accent-ink);
  display: grid;
  place-items: center;
  font-family: var(--font-mono);
  font-weight: 700;
}
.ha-author {
  margin: 0 0 1.5rem;
  padding: 1.1rem 1.2rem;
  border-radius: 12px;
  background: #fff;
  border: 1px solid var(--ha-line);
}
.ha-author h2 {
  margin: 0 0 0.45rem;
  font-size: 1.05rem;
  border: none;
  padding: 0;
}
.ha-author p { margin: 0; color: #3a4652; font-size: 0.98rem; }
.ha-cover-thumb {
  display: inline-block;
  line-height: 0;
  border-radius: 8px;
  overflow: hidden;
  border: 1px solid var(--ha-line);
  margin: 0 0 1.25rem;
}
.ha-cover-thumb img {
  width: min(240px, 100%);
  height: auto;
  display: block;
}
@media (max-width: 920px) {
  .ha-shell { grid-template-columns: 1fr; }
  .ha-sidebar {
    position: static;
    height: auto;
    max-height: none;
    border-right: none;
    border-bottom: 1px solid var(--ha-line);
  }
  .ha-article { width: min(760px, calc(100% - 2rem)); }
  .ha-topbar { padding: 0.85rem 1rem; }
}
"""

CHAPTER_RE = re.compile(r"^第\s*([一二三四五六七八九十百零\d]+)\s*章")
CN_SECTION_RE = re.compile(r"^[一二三四五六七八九十]+[、．.]")
SECTION_RE = re.compile(r"^(\d+(?:\.\d+)+)\.?\s*\S")
CN_NUM = {
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
}


def esc(s: str) -> str:
    return html.escape(s or "", quote=False)


def cn_chapter_num(text: str) -> int | None:
    m = CHAPTER_RE.match(text)
    if not m:
        return None
    raw = m.group(1)
    if raw.isdigit():
        return int(raw)
    if raw in CN_NUM:
        return CN_NUM[raw]
    if raw.startswith("十") and len(raw) == 2:
        return 10 + CN_NUM.get(raw[1], 0)
    return None


def main() -> None:
    if STYLES_SRC.exists():
        shutil.copy2(STYLES_SRC, OUT / "styles.css")

    doc = Document(str(DOCX))

    image_map: dict[str, str] = {}
    for rel in doc.part.rels.values():
        if "image" not in rel.reltype:
            continue
        blob = rel.target_part.blob
        ext = Path(rel.target_ref).suffix.lower() or ".png"
        if ext not in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg"}:
            ext = ".png"
        name = f"img_{rel.rId}{ext}"
        (IMG / name).write_bytes(blob)
        image_map[rel.rId] = f"images/{name}"

    def para_images(p: Paragraph) -> list[str]:
        urls = []
        for blip in p._element.findall(".//a:blip", NSMAP):
            rid = blip.get(qn("r:embed"))
            if rid and rid in image_map:
                urls.append(image_map[rid])
        return urls

    def style_level(style_name: str, text: str) -> str | None:
        s = (style_name or "").strip()
        if s in {"Heading 1", "标题 1", "1"}:
            return "h1"
        if s in {"Heading 2", "标题 2", "2"}:
            return "h2"
        if s in {"Heading 3", "标题 3", "3"}:
            return "h3"
        if s in {"Heading 4", "标题 4", "4"}:
            return "h4"
        if text == "附录" or (text.startswith("附录") and len(text) <= 12):
            return "h1"
        if CHAPTER_RE.match(text) and len(text) < 60:
            return "h1"
        if CN_SECTION_RE.match(text) and len(text) <= 80:
            return "h2"
        if len(text) <= 100:
            m = SECTION_RE.match(text)
            if m:
                depth = m.group(1).count(".")
                if depth == 1:
                    return "h2"
                if depth == 2:
                    return "h3"
                return "h4"
        return None

    def is_code_style(style_name: str) -> bool:
        return "代码" in (style_name or "")

    def is_list_style(style_name: str) -> bool:
        s = style_name or ""
        return s.startswith("List") or "列表" in s

    blocks: list[dict] = []

    def add_para(p: Paragraph) -> None:
        text = (p.text or "").replace("\u00a0", " ").strip()
        style = p.style.name if p.style else "Normal"
        imgs = para_images(p)
        level = style_level(style, text)
        if level and text:
            blocks.append({"type": level, "text": text})
            for u in imgs:
                blocks.append({"type": "img", "src": u})
            return
        if imgs and not text:
            for u in imgs:
                blocks.append({"type": "img", "src": u})
            return
        if not text and not imgs:
            return
        if is_code_style(style):
            blocks.append({"type": "code", "text": text})
            for u in imgs:
                blocks.append({"type": "img", "src": u})
            return
        if is_list_style(style):
            blocks.append({"type": "li", "text": text})
            for u in imgs:
                blocks.append({"type": "img", "src": u})
            return
        if text.startswith("图") and len(text) < 80:
            blocks.append({"type": "figcap", "text": text})
        else:
            blocks.append({"type": "p", "text": text})
        for u in imgs:
            blocks.append({"type": "img", "src": u})

    def table_to_html(table: Table) -> str:
        rows = []
        for i, row in enumerate(table.rows):
            cells = [esc((c.text or "").replace("\n", " ").strip()) for c in row.cells]
            tag = "th" if i == 0 else "td"
            rows.append("<tr>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</tr>")
        return '<div class="table-wrap"><table>\n' + "\n".join(rows) + "\n</table></div>"

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            add_para(Paragraph(child, doc))
        elif child.tag == qn("w:tbl"):
            blocks.append({"type": "table", "html": table_to_html(Table(child, doc))})

    # Drop placeholder / untitled front matter noise
    cleaned: list[dict] = []
    for b in blocks:
        if b["type"] == "p" and b["text"] in {"Untitled", "（注：部分内容由豆包工作 AI 生成）"}:
            continue
        cleaned.append(b)
    blocks = cleaned

    # merge consecutive code / list items
    merged: list[dict] = []
    code_buf: list[str] = []
    list_buf: list[str] = []

    def flush_code() -> None:
        nonlocal code_buf
        if code_buf:
            merged.append({"type": "pre", "text": "\n".join(code_buf)})
            code_buf = []

    def flush_list() -> None:
        nonlocal list_buf
        if list_buf:
            merged.append({"type": "ul", "items": list_buf[:]})
            list_buf = []

    i = 0
    while i < len(blocks):
        b = blocks[i]
        if b["type"] == "code":
            flush_list()
            code_buf.append(b["text"])
            i += 1
            while i < len(blocks) and blocks[i]["type"] == "code":
                code_buf.append(blocks[i]["text"])
                i += 1
            flush_code()
            continue
        flush_code()
        if b["type"] == "li":
            list_buf.append(b["text"])
            i += 1
            continue
        flush_list()
        merged.append(b)
        i += 1
    flush_code()
    flush_list()
    blocks = merged

    chapters: list[dict] = []
    cur = {"num": 0, "title": "前言", "blocks": []}
    for b in blocks:
        if b["type"] == "h1":
            if cur["blocks"] or cur["num"] > 0:
                chapters.append(cur)
            n = cn_chapter_num(b["text"])
            if n is None:
                if b["text"].startswith("附录"):
                    n = 11
                else:
                    n = cur["num"] + 1
            cur = {"num": n, "title": b["text"], "blocks": [b]}
        else:
            cur["blocks"].append(b)
    if cur["blocks"]:
        chapters.append(cur)
    chapters = [c for c in chapters if c["blocks"]]

    files: list[tuple[str, dict]] = []
    for c in chapters:
        if c["num"] == 0:
            fname = "ch00.html"
        elif c["title"].startswith("附录"):
            fname = "appendix.html"
        else:
            fname = f"ch{c['num']:02d}.html"
        files.append((fname, c))

    def short_title(c: dict) -> str:
        t = c["title"]
        m = re.match(r"^第\s*[一二三四五六七八九十百零\d]+\s*章\s*", t)
        if m:
            rest = t[m.end() :].strip()
            return rest or t
        if t.startswith("附录"):
            return "附录"
        return t

    def chapter_num_label(c: dict) -> str:
        if c["title"].startswith("附录"):
            return "附"
        if c["num"] == 0:
            return "·"
        return f"{c['num']:02d}"

    def display_chapter_title(c: dict) -> str:
        """目录展示：保留「第一章 / 第二章」完整标题。"""
        return c["title"]

    chapter_sections: dict[str, list[dict]] = {}

    def book_chapter_items(active: str) -> str:
        items = []
        for fname, c in files:
            secs = chapter_sections.get(fname, [])
            is_open = fname == active
            li_cls = []
            if is_open:
                li_cls.append("open")
                li_cls.append("active")
            cls_attr = f' class="{" ".join(li_cls)}"' if li_cls else ""
            sub_items = "".join(
                f'<li><a href="{fname}#{esc(s["id"])}">{esc(s["label"])}</a></li>'
                for s in secs
            )
            items.append(
                f'<li{cls_attr} data-chap="{fname}">'
                f'<button type="button" class="chap-toggle" aria-expanded="'
                f'{"true" if is_open else "false"}">'
                f'<span class="chap-caret" aria-hidden="true">▸</span>'
                f'<span>{esc(display_chapter_title(c))}</span>'
                f"</button>"
                f'<ul class="chap-subs">{sub_items}</ul>'
                f"</li>"
            )
        return "".join(items)

    SIDEBAR_JS = """
<script>
(function () {
  var list = document.querySelector(".ha-chapters");
  if (!list) return;
  list.addEventListener("click", function (e) {
    var btn = e.target.closest(".chap-toggle");
    if (!btn || !list.contains(btn)) return;
    var li = btn.closest("li[data-chap]");
    if (!li) return;
    var willOpen = !li.classList.contains("open");
    list.querySelectorAll("li[data-chap]").forEach(function (item) {
      item.classList.remove("open");
      var b = item.querySelector(".chap-toggle");
      if (b) b.setAttribute("aria-expanded", "false");
    });
    if (willOpen) {
      li.classList.add("open");
      btn.setAttribute("aria-expanded", "true");
      // 第一次展开时进入该章正文；已在本章则只展开小节
      var href = li.getAttribute("data-chap");
      var file = (location.pathname.split("/").pop() || "");
      if (href && file !== href && !file.endsWith("/" + href)) {
        location.href = href;
      }
    }
  });
})();
</script>
"""

    def sidebar_html(active: str) -> str:
        back = '<a class="ha-back" href="index.html">← 返回目录</a>'
        return f"""<aside class="ha-sidebar" aria-label="目录">
      {back}
      <div class="ha-toc-label">目录</div>
      <ul class="ha-chapters">
        {book_chapter_items(active)}
      </ul>
      <div class="ha-side-foot">
        <a href="cover.html">封面</a>
        <a href="index.html">目录</a>
      </div>
    </aside>"""

    def shell(
        title: str,
        active: str,
        article_inner: str,
    ) -> str:
        return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{esc(title)}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,650&family=IBM+Plex+Mono:wght@400;500&family=Noto+Sans+SC:wght@400;500;700&family=Source+Sans+3:wght@400;600;700&display=swap" rel="stylesheet" />
  <link rel="stylesheet" href="styles.css" />
  <style>{PAGE_CSS_EXTRA}</style>
</head>
<body class="hermes-book">
  <div class="ha-shell">
    {sidebar_html(active)}
    <div class="ha-main">
      <header class="ha-topbar">
        <a class="ha-brand" href="index.html">
          <span class="ha-brand-mark">{BRAND_MARK}</span>
          {esc(BOOK_TITLE)}
        </a>
        <div class="ha-top-links">
          <a href="cover.html">封面</a>
          <a href="index.html">目录</a>
        </div>
      </header>
{article_inner}
    </div>
  </div>
  {SIDEBAR_JS}
</body>
</html>
"""

    def heading_id(text: str, used: dict[str, int]) -> str:
        base = re.sub(r"[^\w\u4e00-\u9fff\-]+", "-", text, flags=re.UNICODE)
        base = re.sub(r"-+", "-", base).strip("-")[:80] or "section"
        n = used.get(base, 0)
        used[base] = n + 1
        return base if n == 0 else f"{base}-{n}"

    def strip_cn_section_prefix(text: str) -> str:
        return re.sub(r"^[一二三四五六七八九十]+[、．.]\s*", "", text).strip()

    def section_label(chap_no: int | str, idx: int, text: str) -> str:
        prefix = "A" if str(chap_no) in {"11", "附"} or str(chap_no) == "附" else str(chap_no)
        if prefix == "11":
            prefix = "A"
        title = strip_cn_section_prefix(text)
        return f"{prefix}.{idx} {title}"

    def collect_sections(bs: list[dict], chap_no) -> list[dict]:
        secs = []
        i = 0
        for b in bs:
            if b["type"] == "h2":
                i += 1
                label = section_label(chap_no, i, b["text"])
                secs.append({"text": b["text"], "label": label, "idx": i})
        return secs

    def render_blocks(bs: list[dict], chap_no) -> tuple[str, str, list[dict]]:
        used_ids: dict[str, int] = {}
        sections = collect_sections(bs, chap_no)
        sec_map = {s["text"]: s for s in sections}

        for b in bs:
            if b["type"] in {"h2", "h3", "h4"}:
                if b["type"] == "h2" and b["text"] in sec_map:
                    label = sec_map[b["text"]]["label"]
                    b["_label"] = label
                    b["_id"] = heading_id(label, used_ids)
                    sec_map[b["text"]]["id"] = b["_id"]
                else:
                    b["_label"] = b["text"]
                    b["_id"] = heading_id(b["text"], used_ids)

        toc_items = []
        for s in sections:
            toc_items.append(
                f'<li style="padding-left:0">'
                f'<a href="#{esc(s["id"])}">{esc(s["label"])}</a></li>'
            )
        for b in bs:
            if b["type"] in {"h3", "h4"}:
                level = int(b["type"][1])
                pad = (level - 2) * 0.75
                toc_items.append(
                    f'<li style="padding-left:{pad}rem">'
                    f'<a href="#{esc(b["_id"])}">{esc(b["text"])}</a></li>'
                )
        section_toc_items = "".join(toc_items)

        out: list[str] = []
        for b in bs:
            t = b["type"]
            if t in {"h1", "h2", "h3", "h4"}:
                aid = b.get("_id") or heading_id(b["text"], used_ids)
                label = b.get("_label") or b["text"]
                out.append(f'<{t} id="{esc(aid)}">{esc(label)}</{t}>')
            elif t == "p":
                out.append(f"<p>{esc(b['text'])}</p>")
            elif t == "figcap":
                out.append(f'<p class="figcap">{esc(b["text"])}</p>')
            elif t == "pre":
                out.append(f"<pre><code>{esc(b['text'])}</code></pre>")
            elif t == "ul":
                items = "".join(f"<li>{esc(x)}</li>" for x in b["items"])
                out.append(f"<ul>{items}</ul>")
            elif t == "img":
                out.append(
                    f'<figure class="figure"><img src="{esc(b["src"])}" alt="" loading="lazy" /></figure>'
                )
            elif t == "table":
                out.append(b["html"])
        return section_toc_items, "\n".join(out), sections

    def page(
        title: str,
        active: str,
        chap_num: str,
        crumb: str,
        h1: str,
        section_toc_items: str,
        sections: list[dict],
        body: str,
        prev_h: str,
        next_h: str,
    ) -> str:
        sec_box = ""
        if sections:
            items = "".join(
                f'<li><a href="#{esc(s["id"])}">{esc(s["label"])}</a></li>'
                for s in sections
            )
            sec_box = f"""        <nav class="ha-sec-box" aria-label="本章小节">
          <div class="label">本章小节</div>
          <ul>
{items}
          </ul>
        </nav>
"""
        article = f"""      <article class="ha-article">
        <div class="ha-meta">
          <div><a href="index.html">目录</a> / {esc(crumb)}</div>
          <div>{esc(BOOK_AUTHOR)} 著</div>
        </div>
        <div class="ha-chapter-head">
          <div class="ha-chapter-num">{esc(chap_num)}</div>
          <h1>{esc(h1)}</h1>
        </div>
{sec_box}        <div class="ha-content">
{body}
        </div>
        <div class="ha-footer-nav">
          <a href="{prev_h}">← 上一页</a>
          <a href="{next_h}">下一页 →</a>
        </div>
      </article>"""
        return shell(title, active, article)

    print("chapters:", [(c["num"], c["title"], len(c["blocks"])) for c in chapters])
    print("images:", len(image_map))

    # 先算齐各章小节，再写页面，保证每页侧栏都能展开任意章
    chapter_bodies: dict[str, tuple[str, str, list[dict]]] = {}
    for fname, c in files:
        body_blocks = c["blocks"]
        if body_blocks and body_blocks[0]["type"] == "h1":
            body_blocks = body_blocks[1:]
        if c["title"].startswith("附录"):
            chap_no = "A"
        elif c["num"] == 0:
            chap_no = 0
        else:
            chap_no = c["num"]
        section_toc_items, body_html, sections = render_blocks(body_blocks, chap_no)
        chapter_sections[fname] = sections
        chapter_bodies[fname] = (section_toc_items, body_html, sections)

    for idx, (fname, c) in enumerate(files):
        prev_h = "index.html" if idx == 0 else files[idx - 1][0]
        next_h = "index.html" if idx == len(files) - 1 else files[idx + 1][0]
        if c["title"].startswith("附录"):
            crumb = "附录"
        elif c["num"] == 0:
            crumb = "前言"
        else:
            crumb = f"第{c['num']:02d}章"
        chap_num = chapter_num_label(c)
        section_toc_items, body_html, sections = chapter_bodies[fname]
        html_page = page(
            title=f"{c['title']} · {BOOK_TITLE}",
            active=fname,
            chap_num=chap_num,
            crumb=crumb,
            h1=c["title"],
            section_toc_items=section_toc_items,
            sections=sections,
            body=body_html,
            prev_h=prev_h,
            next_h=next_h,
        )
        (OUT / fname).write_text(html_page, encoding="utf-8")
        print("wrote", fname, (OUT / fname).stat().st_size, "sections:", len(sections))

    blocks_html = []
    for fname, c in files:
        secs = chapter_sections.get(fname, [])
        sec_lis = "".join(
            f'<li><a href="{fname}#{esc(s["id"])}">{esc(s["label"])}</a></li>'
            for s in secs
        )
        blocks_html.append(
            f"""        <section class="book-block">
          <h2><a href="{fname}">{esc(display_chapter_title(c))}</a></h2>
          <ul class="secs">
{sec_lis}
          </ul>
        </section>"""
        )

    index_article = f"""      <article class="ha-article">
        <div class="ha-meta">
          <div>{esc(BOOK_SUBTITLE)}</div>
          <div>{esc(BOOK_AUTHOR)} 著</div>
        </div>
        <div class="ha-chapter-head">
          <div class="ha-chapter-num">{BRAND_MARK}</div>
          <h1>{esc(BOOK_TITLE)}</h1>
        </div>
        <p style="margin:0 0 1rem;text-indent:0"><a href="cover.html">查看书籍封面 →</a></p>
        <a class="ha-cover-thumb" href="cover.html">
          <img src="images/cover.jpg" alt="{esc(BOOK_TITLE)} 封面" />
        </a>
        <section class="ha-author" aria-label="作者">
          <h2>作者介绍</h2>
          <p style="text-indent:0">{esc(BOOK_AUTHOR_BIO)}</p>
        </section>
{chr(10).join(blocks_html)}
        <div class="ha-footer-nav">
          <a href="cover.html">← 封面</a>
          <a href="{files[0][0]}">开始阅读 →</a>
        </div>
      </article>"""
    (OUT / "index.html").write_text(
        shell(BOOK_TITLE, "index.html", index_article),
        encoding="utf-8",
    )
    print("wrote index.html")

    # 封面页与二维码由人工维护，避免覆盖 images/cover.* / cover.html
    if not (OUT / "cover.html").exists():
        print("skip cover.html (missing; keep manual cover assets)")
    else:
        print("keep existing cover.html / cover images")


if __name__ == "__main__":
    main()
