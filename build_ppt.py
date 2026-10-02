#!/usr/bin/env python3
# Generator: "Cara Menggunakan AI Agent untuk Pekerjaan"
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------------- palette
INK      = RGBColor(0x0E, 0x17, 0x26)
NAVY     = RGBColor(0x10, 0x1A, 0x33)
NAVY_2   = RGBColor(0x18, 0x25, 0x47)
INDIGO   = RGBColor(0x4F, 0x46, 0xE5)
CYAN     = RGBColor(0x06, 0xB6, 0xD4)
AMBER    = RGBColor(0xF5, 0x9E, 0x0B)
GREEN    = RGBColor(0x10, 0xB9, 0x81)
RED      = RGBColor(0xEF, 0x44, 0x44)
MUTED    = RGBColor(0x66, 0x70, 0x85)
LIGHT    = RGBColor(0xF4, 0xF6, 0xFB)
BORDER   = RGBColor(0xE3, 0xE8, 0xF0)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
FONT     = "Calibri"

SW, SH = 13.333, 7.5
ML, MR = 0.75, 0.75
CW = SW - ML - MR  # 11.833

DECK_NAME = "AI Agent untuk Pekerjaan"

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]

_page = {"n": 0}


# ---------------------------------------------------------------- helpers
def rect(slide, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            s.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
    s.shadow.inherit = False
    return s


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


def para(tf, text, size=16, bold=False, color=INK, align=PP_ALIGN.LEFT,
         first=False, space_before=0, space_after=6, line_spacing=1.05,
         italic=False, font=FONT):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    runs = text if isinstance(text, list) else [(text, {})]
    for content, kw in runs:
        r = p.add_run()
        r.text = content
        f = r.font
        f.name = kw.get("font", font)
        f.size = Pt(kw.get("size", size))
        f.bold = kw.get("bold", bold)
        f.italic = kw.get("italic", italic)
        f.color.rgb = kw.get("color", color)
    return p


def bullets(tf, items, size=15, color=RGBColor(0x33, 0x3D, 0x51),
            bullet_color=INDIGO, gap=8, first=True, line_spacing=1.08,
            bold_color=INK):
    """items: list of str, or (str, level), or (str, level, dict)"""
    fp = True
    for it in items:
        lvl, opts = 0, {}
        if isinstance(it, tuple):
            if len(it) == 2:
                it, lvl = it
            else:
                it, lvl, opts = it
        p = tf.paragraphs[0] if fp and first else tf.add_paragraph()
        fp = False
        p.space_before = Pt(0)
        p.space_after = Pt(opts.get("gap", gap))
        p.line_spacing = line_spacing
        pPr = p._p.get_or_add_pPr()
        ind = Inches(0.26 + 0.24 * lvl)
        pPr.set("marL", str(int(ind)))
        pPr.set("indent", str(-int(ind)))
        bc = opts.get("bullet_color", bullet_color)
        marks = ["•", "–", "›"]
        r0 = p.add_run()
        r0.text = marks[min(lvl, 2)] + "    " if opts.get("bullet", True) else ""
        r0.font.name = FONT
        r0.font.size = Pt(opts.get("size", size))
        r0.font.bold = True
        r0.font.color.rgb = bc
        body = it if isinstance(it, str) else str(it)
        # bold segments marked with **...**
        segs = body.split("**")
        for i, seg in enumerate(segs):
            if not seg:
                continue
            r = p.add_run()
            r.text = seg
            r.font.name = FONT
            r.font.size = Pt(opts.get("size", size))
            r.font.bold = (i % 2 == 1) or opts.get("bold", False)
            r.font.color.rgb = opts.get("color", color if i % 2 == 0 else bold_color)
    return tf


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def content_slide(eyebrow, title, subtitle=None, kicker_color=INDIGO):
    _page["n"] += 1
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=WHITE)
    rect(s, 0, 0, SW, 0.13, fill=INDIGO)
    rect(s, 0, 0, 3.1, 0.13, fill=CYAN)

    tf = textbox(s, ML, 0.36, CW, 0.28)
    para(tf, eyebrow.upper(), size=10.5, bold=True, color=kicker_color,
         first=True, space_after=0)

    tf = textbox(s, ML, 0.62, CW, 0.62)
    para(tf, title, size=29, bold=True, color=INK, first=True, space_after=0)

    rect(s, ML, 1.30, 1.15, 0.055, fill=AMBER)

    y = 1.52
    if subtitle:
        tf = textbox(s, ML, 1.46, CW, 0.4)
        para(tf, subtitle, size=13.5, color=MUTED, first=True, space_after=0)
        y = 1.95

    # footer
    rect(s, ML, 6.97, CW, 0.014, fill=BORDER)
    tf = textbox(s, ML, 7.06, 6.0, 0.25)
    para(tf, DECK_NAME, size=9, color=MUTED, first=True, space_after=0)
    tf = textbox(s, SW - MR - 3.0, 7.06, 3.0, 0.25)
    para(tf, f"{_page['n']:02d}", size=9, bold=True, color=MUTED,
         align=PP_ALIGN.RIGHT, first=True, space_after=0)
    return s, y


def card(slide, x, y, w, h, title, body=None, accent=INDIGO, fill=LIGHT,
         title_size=14.5, body_size=11.5, num=None, items=None, item_size=11.5,
         title_color=INK, body_color=RGBColor(0x44, 0x4E, 0x63), gap=5,
         title_lines=1):
    rect(slide, x, y, w, h, fill=fill, line=BORDER, lw=0.75,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.055)
    px, pw = x + 0.26, w - 0.52
    ty = y + 0.24
    if num:
        rect(slide, x + 0.26, y + 0.24, 0.42, 0.42, fill=accent,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.28)
        ntf = textbox(slide, x + 0.26, y + 0.28, 0.42, 0.34,
                      anchor=MSO_ANCHOR.MIDDLE)
        para(ntf, str(num), size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, first=True, space_after=0)
        ttf = textbox(slide, px + 0.56, y + 0.28, pw - 0.56, 0.4,
                      anchor=MSO_ANCHOR.MIDDLE)
        para(ttf, title, size=title_size, bold=True, color=title_color,
             first=True, space_after=0, line_spacing=1.0)
        ty = y + 0.78
    else:
        th = 0.34 + 0.28 * (title_lines - 1)
        ttf = textbox(slide, px, ty, pw, th + 0.06)
        para(ttf, title, size=title_size, bold=True, color=title_color,
             first=True, space_after=0, line_spacing=1.05)
        ty += th
        rect(slide, px, ty, 0.46, 0.05, fill=accent)
        ty += 0.20

    if body:
        btf = textbox(slide, px, ty, pw, h - (ty - y) - 0.2)
        para(btf, body, size=body_size, color=body_color, first=True,
             space_after=0, line_spacing=1.18)
        ty = None
    if items:
        iy = ty if ty is not None else y + 0.9
        itf = textbox(slide, px, iy, pw, h - (iy - y) - 0.18)
        bullets(itf, items, size=item_size, gap=gap)


def table(slide, x, y, w, col_w, headers, rows, row_h=0.42, head_h=0.44,
          font_size=11.5, head_size=11.5, head_fill=NAVY, zebra=True,
          bold_first_col=False):
    nr, nc = len(rows) + 1, len(headers)
    g = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w),
                               Inches(head_h + row_h * len(rows)))
    t = g.table
    t.first_row = True
    t.horz_banding = False
    for i, cw in enumerate(col_w):
        t.columns[i].width = Inches(cw)
    t.rows[0].height = Inches(head_h)
    for i in range(1, nr):
        t.rows[i].height = Inches(row_h)

    def fill_cell(cell, text, bold, size, color, bg, align=PP_ALIGN.LEFT):
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.13)
        cell.margin_right = Inches(0.13)
        cell.margin_top = Inches(0.05)
        cell.margin_bottom = Inches(0.05)
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        p.line_spacing = 1.05
        p.space_after = Pt(0)
        for i, seg in enumerate(str(text).split("**")):
            if not seg:
                continue
            r = p.add_run()
            r.text = seg
            r.font.name = FONT
            r.font.size = Pt(size)
            r.font.bold = bold or (i % 2 == 1)
            r.font.color.rgb = color

    for c, h in enumerate(headers):
        fill_cell(t.cell(0, c), h, True, head_size, WHITE, head_fill)
    for r, row in enumerate(rows):
        bg = WHITE if (r % 2 == 0 or not zebra) else LIGHT
        for c, val in enumerate(row):
            fill_cell(t.cell(r + 1, c), val,
                      bold=(bold_first_col and c == 0), size=font_size,
                      color=INK if c == 0 else RGBColor(0x44, 0x4E, 0x63),
                      bg=bg)
    return t


def dark_slide():
    _page["n"] += 1
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=NAVY)
    return s


# ================================================================ 01 COVER
s = dark_slide()
rect(s, 0, 0, 0.32, SH, fill=INDIGO)
rect(s, 0, 0, 0.32, 2.5, fill=CYAN)
# decoration
rect(s, 8.55, 0.9, 4.1, 5.7, fill=NAVY_2, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
rect(s, 8.95, 1.35, 3.3, 0.9, fill=RGBColor(0x22, 0x30, 0x57),
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.16)
rect(s, 8.95, 2.45, 2.4, 0.5, fill=INDIGO, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
rect(s, 8.95, 3.10, 3.3, 0.42, fill=RGBColor(0x2A, 0x39, 0x63),
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
rect(s, 8.95, 3.66, 2.9, 0.42, fill=RGBColor(0x2A, 0x39, 0x63),
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
rect(s, 8.95, 4.35, 1.5, 1.5, fill=CYAN, shape=MSO_SHAPE.OVAL)
rect(s, 10.6, 4.65, 1.0, 1.0, fill=AMBER, shape=MSO_SHAPE.OVAL)
rect(s, 10.6, 5.8, 1.65, 0.42, fill=RGBColor(0x2A, 0x39, 0x63),
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)

tf = textbox(s, 1.0, 1.55, 7.1, 0.3)
para(tf, "PANDUAN PRAKTIS  •  PRODUCTIVITY WITH AI", size=11.5, bold=True,
     color=CYAN, first=True, space_after=0)

tf = textbox(s, 1.0, 2.05, 7.3, 2.1)
para(tf, "Cara Menggunakan", size=44, bold=True, color=WHITE, first=True,
     space_after=2, line_spacing=1.0)
para(tf, [("AI Agent", {"color": AMBER}), (" untuk Pekerjaan", {})],
     size=44, bold=True, color=WHITE, space_after=0, line_spacing=1.0)

rect(s, 1.0, 4.25, 1.6, 0.06, fill=CYAN)

tf = textbox(s, 1.0, 4.55, 6.9, 1.1)
para(tf, "Bagaimana memakai agen AI untuk mengotomasi tugas, mempercepat "
         "pengambilan keputusan, dan menaikkan produktivitas tim — dari "
         "memilih use case hingga monitoring dan skalasi.",
     size=15, color=RGBColor(0xB9, 0xC2, 0xD6), first=True, space_after=0,
     line_spacing=1.3)

rect(s, 1.0, 6.15, 3.9, 0.014, fill=RGBColor(0x2A, 0x39, 0x63))
tf = textbox(s, 1.0, 6.3, 7.0, 0.35)
para(tf, "2026  •  Deck internal untuk tim & pimpinan", size=11.5,
     color=MUTED, first=True, space_after=0)
notes(s, "Pembuka: definisikan AI agent sebagai 'pekerja digital' yang "
         "menjalankan tugas, bukan sekadar chatbot yang menjawab pertanyaan.")

# ================================================================ 02 AGENDA
s, y = content_slide("Agenda", "Apa yang Akan Kita Bahas",
                     "Alur dari konsep → implementasi → tata kelola → rencana aksi")
items = [
    ("01", "Apa itu AI Agent & cara kerjanya", INDIGO),
    ("02", "Perbandingan dengan chatbot & otomasi biasa", CYAN),
    ("03", "Kapan layak dipakai + use case per fungsi", AMBER),
    ("04", "6 langkah implementasi end-to-end", INDIGO),
    ("05", "Tools, framework & pola desain agent", CYAN),
    ("06", "Best practice, risiko & mitigasi", AMBER),
    ("07", "Studi kasus: laporan mingguan otomatis", INDIGO),
    ("08", "Roadmap 30-60-90 hari & langkah berikutnya", CYAN),
]
cw = (CW - 0.4) / 2
for i, (num, txt, col) in enumerate(items):
    cx = ML + (cw + 0.4) * (i % 2)
    cy = y + (i // 2) * 1.20
    rect(s, cx, cy, cw, 1.0, fill=LIGHT, line=BORDER, lw=0.75,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.09)
    rect(s, cx + 0.24, cy + 0.24, 0.52, 0.52, fill=col,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
    tf = textbox(s, cx + 0.24, cy + 0.3, 0.52, 0.42, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, num, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         first=True, space_after=0)
    tf = textbox(s, cx + 0.95, cy + 0.2, cw - 1.2, 0.65, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, txt, size=14.5, bold=True, color=INK, first=True, space_after=0,
         line_spacing=1.12)
notes(s, "Tekankan urutan logis: pahami dulu konsepnya, baru implementasi, "
         "baru tata kelola risiko, baru roadmap.")

# ================================================================ 03 DEFINISI
s, y = content_slide("Konsep Dasar", "Apa Itu AI Agent?",
                     "Sistem yang mengejar sebuah tujuan secara mandiri: merencanakan, "
                     "memanggil tools, lalu evaluasi hasilnya")
rect(s, ML, y, CW, 1.28, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.09)
rect(s, ML, y, 0.09, 1.28, fill=AMBER)
tf = textbox(s, ML + 0.45, y + 0.22, CW - 0.9, 0.9, anchor=MSO_ANCHOR.MIDDLE)
para(tf, [("“", {"color": CYAN}),
          ("Diberi ", {}), ("tujuan", {"bold": True, "color": WHITE}),
          (", bukan sekadar perintah — agent memecah tugas, mengambil tindakan "
           "dengan tools, membaca hasilnya, lalu memperbaiki diri sampai target tercapai.",
           {}), ("”", {"color": CYAN})],
     size=17, color=RGBColor(0xD5, 0xDC, 0xEC), first=True, space_after=0,
     line_spacing=1.25)

y2 = y + 1.55
tf = textbox(s, ML, y2, CW, 0.3)
para(tf, "LIMA CIRI UTAMA", size=11, bold=True, color=INDIGO, first=True,
     space_after=0)
y2 += 0.36
cir = [
    ("Goal-oriented", "Fokus pada hasil akhir, bukan satu giliran chat.", INDIGO),
    ("Autonomous", "Menjalankan banyak langkah tanpa diawasi terus-menerus.", CYAN),
    ("Tool-using", "Akses API, database, browser, file, dan code.", AMBER),
    ("Feedback loop", "Menilai output sendiri dan mengulang bila perlu.", GREEN),
    ("Memory", "Mengingat konteks, riwayat, dan pengetahuan perusahaan.", INDIGO),
]
cwid = (CW - 4 * 0.22) / 5
for i, (t, b, col) in enumerate(cir):
    cx = ML + i * (cwid + 0.22)
    card(s, cx, y2, cwid, 1.75, t, b, accent=col, title_size=13.5,
         body_size=11.5)

tf = textbox(s, ML, y2 + 1.95, CW, 0.4)
para(tf, "Ringkasnya: model AI adalah otaknya, agent-lah yang memberinya tangan, "
         "ingatan, dan prosedur kerja.", size=13, italic=True, color=MUTED,
     first=True, space_after=0)
notes(s, "Analogi: chatbot = resepsionis yang menjawab; agent = staf yang "
         "diberi brief, lalu mengerjakan sampai selesai.")

# ================================================================ 04 PERBANDINGAN
s, y = content_slide("Perbandingan", "AI Agent vs Chatbot vs Otomasi Biasa",
                     "Tiga pendekatan berbeda — pilih sesuai kompleksitas pekerjaan")
table(s, ML, y, CW, [2.7, 3.05, 3.05, 3.03],
      ["Aspek", "Chatbot", "Otomasi / RPA", "AI Agent"],
      [
          ["Input", "Pertanyaan pengguna", "Trigger & aturan tetap", "Tujuan / brief terbuka"],
          ["Langkah", "Satu giliran (Q&A)", "Alur sudah dipetakan", "Merencanakan multi-langkah"],
          ["Fleksibilitas", "Rendah–sedang", "Rendah (if-then)", "Tinggi, adaptif pada konteks"],
          ["Pemakaian tools", "Tidak / terbatas", "Ya, sesuai skrip", "Ya, memilih tools sendiri"],
          ["Menangani kasus baru", "Serba salah", "Gagal / perlu kode baru", "Bisa menyusun pendekatan"],
          ["Butuh intervensi", "Setiap giliran", "Saat exception", "Saat melewati batas aman"],
          ["Cocok untuk", "FAQ & dukungan", "Proses berulang & terstruktur", "Pekerjaan butuh penalaran"],
      ], row_h=0.44, head_h=0.46, font_size=11.5, bold_first_col=True)

yy = y + 0.46 + 7 * 0.44 + 0.28
rect(s, ML, yy, CW, 0.72, fill=RGBColor(0xEC, 0xEE, 0xFC),
     line=INDIGO, lw=0.75, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.14)
tf = textbox(s, ML + 0.3, yy + 0.1, CW - 0.6, 0.55, anchor=MSO_ANCHOR.MIDDLE)
para(tf, [("Kesimpulan:  ", {"bold": True, "color": INDIGO}),
          ("AI Agent = otak (LLM) + tangan (tools) + ingatan (memory) + prosedur "
           "(guardrail). Gunakan chatbot untuk tanya-jawab, RPA untuk alur kaku, "
           "agent untuk pekerjaan yang butuh penalaran dan adaptasi.", {})],
     size=13, color=RGBColor(0x33, 0x3D, 0x51), first=True, space_after=0,
     line_spacing=1.2)
notes(s, "Bila prosesnya bisa dideskripsikan dalam if-then, RPA cukup. Agent "
         "mulai unggul saat ada banyak kemungkinan jalur dan sumber data.")

# ================================================================ 05 SIKLUS
s, y = content_slide("Cara Kerja", "Siklus Kerja Sebuah AI Agent",
                     "Loop yang berulang: rencana → tindakan → hasil → evaluasi")
steps = [
    ("1", "PERCEIVE", "Baca input, konteks, dan data dari sumber terhubung.", INDIGO),
    ("2", "PLAN", "Pecah tujuan menjadi sub-tugas dan urutan langkah.", CYAN),
    ("3", "ACT", "Panggil tools: API, database, email, browser, code.", AMBER),
    ("4", "OBSERVE", "Baca hasil eksekusi dan error yang muncul.", GREEN),
    ("5", "REFLECT", "Bandingkan dengan kriteria; iterasi bila belum sesuai.", RED),
]
n = len(steps)
gap = 0.30
bw = (CW - gap * (n - 1)) / n
for i, (num, t, b, col) in enumerate(steps):
    bx = ML + i * (bw + gap)
    rect(s, bx, y, bw, 2.35, fill=WHITE, line=BORDER, lw=1,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    rect(s, bx, y, bw, 0.09, fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
    rect(s, bx + 0.24, y + 0.32, 0.5, 0.5, fill=col, shape=MSO_SHAPE.OVAL)
    tf = textbox(s, bx + 0.24, y + 0.37, 0.5, 0.4, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, num, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         first=True, space_after=0)
    tf = textbox(s, bx + 0.24, y + 0.95, bw - 0.48, 0.3)
    para(tf, t, size=13, bold=True, color=col, first=True, space_after=0)
    tf = textbox(s, bx + 0.24, y + 1.3, bw - 0.48, 0.95)
    para(tf, b, size=11.5, color=RGBColor(0x44, 0x4E, 0x63), first=True,
         space_after=0, line_spacing=1.18)
    if i < n - 1:
        tf = textbox(s, bx + bw + 0.03, y + 0.95, gap, 0.4,
                     anchor=MSO_ANCHOR.MIDDLE)
        para(tf, "›", size=22, bold=True, color=MUTED, align=PP_ALIGN.CENTER,
             first=True, space_after=0)

y2 = y + 2.65
sup = [
    ("Memory", "Menyimpan riwayat percakapan, keputusan, dan pengetahuan perusahaan agar agent tidak mulai dari nol.", INDIGO),
    ("Guardrail", "Aturan batas: data apa yang boleh diakses, aksi mana yang butuh persetujuan manusia, format output yang valid.", CYAN),
]
sw_ = (CW - 0.35) / 2
for i, (t, b, col) in enumerate(sup):
    card(s, ML + i * (sw_ + 0.35), y2, sw_, 1.5, t, b, accent=col,
         title_size=14, body_size=12)

tf = textbox(s, ML, y2 + 1.68, CW, 0.35)
para(tf, "Loop inilah yang membuat agent mengerjakan tugas sampai selesai — "
         "bukan berhenti setelah satu jawaban.", size=12.5, italic=True,
     color=MUTED, first=True, space_after=0)
notes(s, "Tekankan langkah REFLECT: di situlah letak kualitas. Tanpa evaluasi, "
         "agent hanya menebak.")

# ================================================================ 06 KOMPONEN
s, y = content_slide("Arsitektur", "Komponen Sistem AI Agent",
                     "Lima lapisan yang harus disiapkan sebelum agent diproduksi")
comp = [
    ("1", "Model (LLM)", "Otak penalaran. Pilih berdasarkan kemampuan reasoning, konteks, biaya, dan privasi data.", INDIGO),
    ("2", "Tools & Integrations", "Tangan digital: Gmail, Slack, Notion, Jira, CRM, database, browser, code runner.", CYAN),
    ("3", "Memory", "Konteks jangka pendek + pengetahuan jangka panjang (knowledge base / RAG).", AMBER),
    ("4", "Orchestration", "Pengatur alur: satu agent atau banyak agent bertugas (router → spesialis).", GREEN),
    ("5", "Guardrail & Monitoring", "Validasi output, kontrol akses, log, metrik, dan peringatan bila menyimpang.", RED),
]
ch = 0.88
for i, (num, t, b, col) in enumerate(comp):
    cy = y + i * (ch + 0.13)
    rect(s, ML, cy, CW, ch, fill=LIGHT, line=BORDER, lw=0.75,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.11)
    rect(s, ML, cy, 0.08, ch, fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    rect(s, ML + 0.34, cy + 0.24, 0.44, 0.44, fill=col,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.28)
    tf = textbox(s, ML + 0.34, cy + 0.28, 0.44, 0.38, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, num, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         first=True, space_after=0)
    tf = textbox(s, ML + 1.0, cy + 0.16, 3.0, 0.62, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, t, size=14.5, bold=True, color=INK, first=True, space_after=0,
         line_spacing=1.05)
    tf = textbox(s, ML + 4.15, cy + 0.16, CW - 4.5, 0.62, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, b, size=12.5, color=RGBColor(0x44, 0x4E, 0x63), first=True,
         space_after=0, line_spacing=1.18)
notes(s, "Kesalahan umum: hanya menyiapkan model, lalu lupa memory, guardrail, "
         "dan monitoring. Agent mentah tanpa ini sulit dipercaya.")

# ================================================================ 07 KAPAN PAKAI
s, y = content_slide("Kelayakan", "Kapan AI Agent Layak Dipakai?",
                     "Uji cepat sebelum berinvestasi: cocokkan dengan profil pekerjaan")
col_w = (CW - 0.35) / 2
rect(s, ML, y, col_w, 2.85, fill=RGBColor(0xEC, 0xFB, 0xF6),
     line=GREEN, lw=1, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
tf = textbox(s, ML + 0.3, y + 0.24, col_w - 0.6, 0.32)
para(tf, "✓  COCOK", size=13, bold=True, color=GREEN, first=True, space_after=0)
tf = textbox(s, ML + 0.3, y + 0.66, col_w - 0.6, 2.0)
bullets(tf, [
    "Tugas **berulang** harian/mingguan dengan volume tinggi",
    "Melibatkan **banyak sumber data** (file, API, chat, database)",
    "Aturan mainnya bisa **didokumentasikan**",
    "Ada cara **mengukur** benar atau salahnya",
    "Biaya kalau salah masih **bisa ditoleransi**",
], size=12.5, gap=9)

rect(s, ML + col_w + 0.35, y, col_w, 2.85, fill=RGBColor(0xFE, 0xF2, 0xF2),
     line=RED, lw=1, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
tf = textbox(s, ML + col_w + 0.65, y + 0.24, col_w - 0.6, 0.32)
para(tf, "✕  BELUM COCOK", size=13, bold=True, color=RED, first=True, space_after=0)
tf = textbox(s, ML + col_w + 0.65, y + 0.66, col_w - 0.6, 2.0)
bullets(tf, [
    "Keputusan **kritis** tanpa pengawasan manusia",
    "Data **sangat sensitif** tanpa kontrol akses yang jelas",
    "Proses **belum stabil** / belum pernah dipetakan",
    "Hasilnya **subyektif** dan tidak bisa dievaluasi",
    "Satu kesalahan berdampak **hukum atau finansial besar**",
], size=12.5, gap=9, bullet_color=RED)

y2 = y + 3.1
tf = textbox(s, ML, y2, CW, 0.3)
para(tf, "TIGA PERTANYAAN SEBELUM MULAI", size=11, bold=True, color=INDIGO,
     first=True, space_after=0)
qs = [
    ("Bisa ditulis langkahnya?", "Jika alur kerjanya belum bisa dijelaskan ke orang baru, agent juga tidak akan bisa."),
    ("Ada sumber data yang jelas?", "Agent hanya sebaik data dan akses yang diberikan padanya."),
    ("Ada cara mengukur berhasil?", "Tanpa baseline & metrik, kita tidak tahu apakah ini benar-benar membantu."),
]
qw = (CW - 0.5) / 3
for i, (t, b) in enumerate(qs):
    card(s, ML + i * (qw + 0.25), y2 + 0.35, qw, 1.5, t, b, accent=INDIGO,
         title_size=13.5, body_size=11.5)
notes(s, "Saran: mulai dari use case dengan 'biaya salah rendah' seperti "
         "drafting, ringkasan, dan klasifikasi.")

# ================================================================ 08 USE CASE
s, y = content_slide("Penerapan", "Contoh Pekerjaan per Fungsi",
                     "Dari pekerjaan berulang menjadi hasil terukur dalam hitungan menit")
uc = [
    ("Marketing", ["Riset kata kunci & pesaing", "Draf konten multi-kanal", "Laporan performa kampanye"], INDIGO),
    ("Sales", ["Enrichment & skoring lead", "Personalisasi email tindak lanjut", "Ringkasan hasil panggilan"], CYAN),
    ("HR / People", ["Screening awal CV", "Alur onboarding otomatis", "FAQ kebijakan perusahaan"], AMBER),
    ("Finance", ["Rekonsiliasi transaksi", "Deteksi anomali & flag", "Draf laporan berkala"], GREEN),
    ("IT / Engineering", ["Triage insiden & tiket", "Code review & dokumentasi", "Pembaruan status teknis"], RED),
    ("Operations & CX", ["Klasifikasi tiket masuk", "Jawaban tier-1 24/7", "QC dokumen & invoice"], INDIGO),
]
uw = (CW - 0.6) / 3
uh = 2.05
for i, (t, its, col) in enumerate(uc):
    cx = ML + (i % 3) * (uw + 0.30)
    cy = y + (i // 3) * (uh + 0.22)
    card(s, cx, cy, uw, uh, t, accent=col, title_size=15,
         items=[(x, 0) for x in its], item_size=11.5, gap=7)

yy = y + 2 * uh + 0.22 + 0.15
rect(s, ML, yy, CW, 0.5, fill=LIGHT, line=BORDER, lw=0.75,
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.18)
tf = textbox(s, ML + 0.3, yy + 0.04, CW - 0.6, 0.42, anchor=MSO_ANCHOR.MIDDLE)
para(tf, [("Aturan praktis:  ", {"bold": True, "color": INDIGO}),
          ("mulai dari pekerjaan yang menghabiskan waktu, bukan yang menghasilkan "
           "keputusan. Hemat waktu dulu, otonomi belakangan.", {})],
     size=12.5, color=RGBColor(0x33, 0x3D, 0x51), first=True, space_after=0)
notes(s, "Minta peserta memilih satu fungsi mereka sendiri sambil melihat "
         "slide ini — biarkan konteksnya relevan.")

# ================================================================ 09 6 LANGKAH
s, y = content_slide("Kerangka Kerja", "6 Langkah Implementasi",
                     "Urutan yang terbukti aman dan cepat menghasilkan nilai")
steps6 = [
    ("01", "Pilih use case berdampak", "Prioritaskan dampak × kemudahan; mulai dari quick win.", INDIGO),
    ("02", "Petakan proses & data", "Alur as-is, sumber data, pemilik data, dan pengetahuan yang dibutuhkan.", CYAN),
    ("03", "Desain prompt & guardrail", "Struktur instruksi, format output, batasan, dan titik persetujuan manusia.", AMBER),
    ("04", "Integrasi tools", "Sambungkan sistem; mulai read-only, lalu write dengan approval.", GREEN),
    ("05", "Uji & ukur", "Golden dataset, evaluasi otomatis, bandingkan dengan baseline.", RED),
    ("06", "Deploy, monitor, iterasi", "Pilot → bertahap → penuh; dashboard, feedback, dan perbaikan rutin.", INDIGO),
]
sw2 = (CW - 0.6) / 3
sh2 = 2.3
for i, (num, t, b, col) in enumerate(steps6):
    cx = ML + (i % 3) * (sw2 + 0.30)
    cy = y + (i // 3) * (sh2 + 0.3)
    card(s, cx, cy, sw2, sh2, t, b, accent=col, title_size=15, body_size=12,
         num=num)
notes(s, "Enam langkah ini dijabarkan pada slide berikutnya satu per satu.")

# ================================================================ 10-15 DETAIL LANGKAH
def step_slide(idx, title, subtitle, goal, items, right_title, right_items,
               color=INDIGO, tip=None):
    s, y = content_slide(f"Langkah {idx} dari 6", title, subtitle, kicker_color=color)
    lw_ = CW * 0.585
    rw_ = CW - lw_ - 0.38
    PH = 4.15
    rect(s, ML, y, lw_, PH, fill=WHITE, line=BORDER, lw=1,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    rect(s, ML, y, lw_, 0.09, fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
    tf = textbox(s, ML + 0.32, y + 0.3, lw_ - 0.64, 0.3)
    para(tf, "YANG DILAKUKAN", size=10.5, bold=True, color=color, first=True,
         space_after=0)
    tf = textbox(s, ML + 0.32, y + 0.66, lw_ - 0.64, PH - 0.85)
    bullets(tf, items, size=13, gap=11, bullet_color=color)

    rx = ML + lw_ + 0.38
    rect(s, rx, y, rw_, PH, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    rect(s, rx, y, 0.09, PH, fill=AMBER, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    tf = textbox(s, rx + 0.34, y + 0.3, rw_ - 0.68, 0.3)
    para(tf, right_title.upper(), size=10.5, bold=True, color=CYAN, first=True,
         space_after=0)
    tf = textbox(s, rx + 0.34, y + 0.66, rw_ - 0.68, PH - 0.85)
    bullets(tf, right_items, size=12.5, gap=10, color=RGBColor(0xD5, 0xDC, 0xEC),
            bullet_color=AMBER, bold_color=WHITE)

    if tip:
        yy = y + PH + 0.16
        rect(s, ML, yy, CW, 0.6, fill=RGBColor(0xFF, 0xF7, 0xE8),
             line=AMBER, lw=0.75, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.16)
        tf = textbox(s, ML + 0.3, yy + 0.06, CW - 0.6, 0.48, anchor=MSO_ANCHOR.MIDDLE)
        para(tf, [("Tips:  ", {"bold": True, "color": RGBColor(0xB4, 0x73, 0x06)}),
                  (tip, {})], size=12, color=RGBColor(0x5A, 0x44, 0x14),
             first=True, space_after=0, line_spacing=1.15)
    notes(s, "Langkah %s dari 6 - %s. Pokok pembahasan: %s. "
             "Panel gelap di kanan bisa dijadikan checklist implementasi. %s"
          % (idx, title, right_title,
             ("Tips penyampaian: " + tip) if tip else ""))
    return s


step_slide(
    "1", "Pilih Use Case yang Berdampak",
    "Pekerjaan yang tepat membuat hasil cepat terlihat dan mudah mendapat dukungan",
    None,
    [
        "Petakan daftar tugas tim: **frekuensi × waktu × kesulitan**",
        "Pilih yang **berulang harian/mingguan** dan berbasis data",
        "Cari *quick win*: hasil terlihat dalam 2–4 minggu",
        "Hindari menggantikan **seluruh proses** sekaligus di awal",
        "Pastikan ada **pemilik use case** yang bertanggung jawab",
    ],
    "Kriteria pemilihan",
    [
        "Volume tinggi",
        "Langkah bisa didokumentasikan",
        "Biaya per kesalahan rendah",
        "Data tersedia & legal diakses",
        "Mudah diukur hasilnya",
    ],
    color=INDIGO,
    tip="Mulai dari satu use case yang paling menyakitkan, bukan yang paling keren.",
)

step_slide(
    "2", "Petakan Proses & Siapkan Data",
    "Agent hanya sebaik konteks dan data yang Anda berikan kepadanya",
    None,
    [
        "Tulis alur **as-is** per langkah, termasuk pengecualian",
        "Tandai setiap **input, keputusan, dan output**",
        "Inventaris sumber data: dokumen, DB, Slack, email, CRM",
        "Bersihkan & rapikan **knowledge base** (dokumen usang = jawaban usang)",
        "Tetapkan **data owner** dan aturan akses per sumber",
    ],
    "Checklist data",
    [
        "Format konsisten & mudah dibaca",
        "Pembaruan berkala terjadwal",
        "Tidak ada data pribadi tanpa dasar",
        "Versioning dokumen penting",
        "Akses minimal, sesuai kebutuhan",
    ],
    color=CYAN,
    tip="Sering kali 60% pekerjaan implementasi ada di sini — bukan di model AI-nya.",
)

step_slide(
    "3", "Desain Prompt & Guardrail",
    "Instruksi yang jelas dan batasan yang tegas adalah kualitas agent",
    None,
    [
        "Susun prompt dari 5 blok: **role → konteks → instruksi → format → batasan**",
        "Wajibkan **format output terstruktur** (JSON/templat) agar bisa divalidasi",
        "Tetapkan **ambang keyakinan**: bila ragu, agent harus bertanya, bukan mengarang",
        "Tandai aksi **berisiko** yang butuh persetujuan manusia",
        "Simpan prompt sebagai **aset berversi**, bukan tulisan lepas",
    ],
    "Guardrail wajib",
    [
        "Read-only sebagai default",
        "Approval manusia untuk aksi tulis/eksternal",
        "Validasi skema output",
        "Blokir akses di luar daftar sumber",
        "Log prompt & hasil untuk audit",
    ],
    color=AMBER,
    tip="Uji prompt dengan kasus terburuk (edge case), bukan hanya contoh yang mudah.",
)

step_slide(
    "4", "Integrasi dengan Tools & Sistem",
    "Sambungkan agent ke tempat pekerjaan benar-benar terjadi",
    None,
    [
        "Mulai dengan **read-only**: baca email, dokumen, laporan, tiket",
        "Naikkan ke **write access** bertahap, dengan persetujuan manusia",
        "Pisahkan kredensial per fungsi — terapkan **least privilege**",
        "Catat setiap pemanggilan tool: siapa, kapan, apa, hasil apa",
        "Siapkan **retry, timeout, dan idempotensi** agar aman diulang",
    ],
    "Contoh tools",
    [
        "Komunikasi: Slack, Gmail, Meet",
        "Dokumen: Notion, Drive, Confluence",
        "Proyek: Jira, Linear, Asana",
        "Data: Sheets, BigQuery, API internal",
        "Browser & code execution (terbatas)",
    ],
    color=GREEN,
    tip="Pisahkan akses baca dan tulis. Perubahan izin = perubahan risiko.",
)

step_slide(
    "5", "Uji, Evaluasi & Ukur",
    "Tanpa baseline dan metrik, tidak ada cara membedakan membantu vs mengarang",
    None,
    [
        "Kumpulkan **golden dataset**: input beserta jawaban benar",
        "Jalankan **evaluasi otomatis** pada setiap perubahan prompt/model",
        "Uji **regresi** — perbaikan A boleh jadi merusak B",
        "Bandingkan hasil agent vs cara kerja lama (**A/B**)",
        "Ukur dari sisi **kualitas, waktu, dan biaya** sekaligus",
    ],
    "Metrik utama",
    [
        "Task success rate",
        "Akurasi / tingkat halusinasi",
        "Waktu hemat per tugas",
        "Biaya per tugas (token)",
        "Tingkat eskalasi ke manusia",
    ],
    color=RED,
    tip="Tetapkan angka baseline dulu sebelum menargetkan improvement.",
)

step_slide(
    "6", "Deploy, Monitor & Iterasi",
    "Agent bukan proyek sekali jalan — ia hidup dan perlu dirawat",
    None,
    [
        "Rilis bertahap: **pilot kecil → satu tim → seluruh organisasi**",
        "Pasang dashboard: sukses, latensi, biaya, dan error",
        "Bangun **feedback loop**: penilaian pengguna + koreksi mereka",
        "**Versioning** prompt, konfigurasi, dan pengetahuan agent",
        "Siapkan **rencana rollback** bila hasil memburuk",
    ],
    "Ritme perawatan",
    [
        "Harian: pantau error & biaya",
        "Mingguan: tinjau kasus gagal",
        "Bulanan: audit akurasi & akses",
        "Quarterly: evaluasi ROI & model",
        "Berkala: perbarui knowledge base",
    ],
    color=INDIGO,
    tip="Jadwalkan pembaruan knowledge base — konteks yang basi adalah penyebab umum kegagalan.",
)

# ================================================================ 16 TOOLS
s, y = content_slide("Peralatan", "Tools & Framework Populer",
                     "Pilih berdasarkan kebutuhan integrasi, keamanan, dan biaya — bukan tren")
table(s, ML, y, CW, [2.9, 5.5, 3.43],
      ["Kategori", "Pilihan Umum", "Cocok Untuk"],
      [
          ["Coding agent", "OpenCode, Claude Code, Cursor, GitHub Copilot", "Developer & data task"],
          ["Agent framework", "OpenAI Agents SDK, Anthropic Agent SDK, LangGraph", "Tim teknis membangun sendiri"],
          ["Multi-agent", "CrewAI, AutoGen, LangGraph", "Alur butuh banyak peran"],
          ["No-code / otomasi", "n8n, Zapier Agents, Make, Copilot Studio", "Tim non-teknis, cepat coba"],
          ["Kolaborasi & kantor", "Microsoft 365 Copilot, Google Agentspace", "Dokumen, email, meeting"],
          ["Bisnis / CRM", "Salesforce Agentforce, ServiceNow, HubSpot Breeze", "Sales, support, IT service"],
      ], row_h=0.42, head_h=0.46, font_size=11.5, bold_first_col=True)

yy = y + 0.46 + 6 * 0.42 + 0.26
crit = [
    ("Integrasi", "Apakah terhubung ke sistem yang tim Anda pakai sehari-hari?"),
    ("Keamanan", "Kontrol akses, audit log, opsi data residency / on-prem."),
    ("Biaya", "Model per token, lisensi, dan biaya operasional per tugas."),
    ("Kendali", "Bisa evaluasi, versioning, dan rollback sendiri?"),
]
cw3 = (CW - 0.6) / 4
for i, (t, b) in enumerate(crit):
    card(s, ML + i * (cw3 + 0.2), yy, cw3, 1.55, t, b, accent=CYAN,
         title_size=13, body_size=11)
notes(s, "Tidak ada tool 'terbaik' universal — yang cocok bergantung pada "
         "tumpukan teknologi dan tingkat kontrol keamanan yang dibutuhkan.")

# ================================================================ 17 POLA DESAIN
s, y = content_slide("Pola Desain", "Pola yang Sering Dipakai dalam Agent",
                     "Pilih pola sesuai tingkat kerumitan, bukan selalu yang paling canggih")
patterns = [
    ("ReAct", "Reasoning + acting bergantian: berpikir, bertindak, membaca hasil, berpikir lagi.", "Default untuk kebanyakan tugas", INDIGO),
    ("Plan & Execute", "Rencana lengkap disusun dulu, baru dieksekuti langkah demi langkah.", "Tugas panjang & terstruktur", CYAN),
    ("Multi-agent", "Router menugaskan spesialis: riset, penulisan, analisis, QA.", "Pekerjaan kompleks & beragam", AMBER),
    ("Human-in-the-loop", "Agent berhenti di titik persetujuan sebelum aksi berisiko.", "Aksi tulis/eksternal penting", GREEN),
    ("Reflection", "Agent menilai dan menyunting hasilnya sendiri sebelum dikirim.", "Output yang butuh kualitas tinggi", RED),
    ("RAG + Agent", "Menarik pengetahuan relevan dari basis data perusahaan saat dibutuhkan.", "Jawaban berbasis dokumen internal", INDIGO),
]
pw2 = (CW - 0.6) / 3
ph2 = 2.28
for i, (t, b, tag, col) in enumerate(patterns):
    cx = ML + (i % 3) * (pw2 + 0.30)
    cy = y + (i // 3) * (ph2 + 0.28)
    card(s, cx, cy, pw2, ph2, t, b, accent=col, title_size=15.5, body_size=12)
    rect(s, cx + 0.26, cy + ph2 - 0.62, pw2 - 0.52, 0.4, fill=WHITE,
         line=BORDER, lw=0.75, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
    tf = textbox(s, cx + 0.4, cy + ph2 - 0.58, pw2 - 0.7, 0.32,
                 anchor=MSO_ANCHOR.MIDDLE)
    para(tf, tag, size=10, bold=True, color=col, first=True, space_after=0)
notes(s, "Saran: 80% use case kantor cukup dengan ReAct + RAG + "
         "human-in-the-loop.")

# ================================================================ 18 BEST PRACTICE
s, y = content_slide("Praktik Terbaik", "8 Prinsip yang Tidak Boleh Dilewatkan",
                     "Dari pengalaman implementasi di dunia nyata")
bp = [
    ("Mulai kecil, skala bertahap", "Bukti nilai lebih dulu, baru perluas. Hindari proyek 'transformasi total' di awal.", INDIGO),
    ("Manusia tetap di kursi kemudi", "Aksi berisiko selalu butuh persetujuan, terutama yang bersifat eksternal atau finansial.", CYAN),
    ("Konteks > prompt panjang", "Rapikan data dan knowledge base. Kualitas input menentukan kualitas hasil.", AMBER),
    ("Instrumentasi sejak hari pertama", "Log, metrik, dan biaya dipantau sejak pilot agar keputusan berbasis data.", GREEN),
    ("Versioning prompt & konfigurasi", "Perlakukan prompt seperti kode: review, uji regresi, dan riwayat perubahan.", RED),
    ("Libatkan tim pendukung sejak awal", "Security, legal, dan data owner terlibat di muka - lebih murah mencegah daripada memperbaiki.", INDIGO),
    ("Batasi akses (least privilege)", "Agent hanya memegang izin minimum yang diperlukan untuk tugasnya.", CYAN),
    ("Latih orang, bukan hanya tool", "Agent mengubah cara kerja. Tanpa pelatihan, adopsi dan manfaatnya mandek.", AMBER),
]
bw2 = (CW - 0.6) / 4
bh2 = 2.3
for i, (t, b, col) in enumerate(bp):
    cx = ML + (i % 4) * (bw2 + 0.2)
    cy = y + (i // 4) * (bh2 + 0.3)
    card(s, cx, cy, bw2, bh2, t, b, accent=col, title_size=13.5, body_size=11.5,
         title_lines=2)
notes(s, "Poin 7 dan 8 sering dilupakan: kontrol akses dan kesiapan tim.")

# ================================================================ 19 RISIKO
s, y = content_slide("Tata Kelola", "Risiko & Cara Mitigasinya",
                     "Waspadai sejak awal agar kepercayaan tidak hilang di tengah jalan")
table(s, ML, y, CW, [2.85, 4.5, 4.48],
      ["Risiko", "Dampak", "Mitigasi"],
      [
          ["Halusinasi", "Informasi keliru dikira fakta", "RAG, validasi skema output, evaluasi rutin, wajib verifikasi sumber"],
          ["Kebocoran data", "Data sensitif keluar dari lingkungan aman", "Masking data, kontrol akses, pilihan model/private deployment"],
          ["Izin berlebihan", "Agent menulis/menghapus di luar wewenang", "Least privilege, mulai read-only, approval untuk aksi tulis"],
          ["Prompt injection", "Instruksi jahat membelokkan agent", "Isolasi konteks teks eksternal, filter input, batasi tool"],
          ["Shadow AI", "Alat tak resmi dipakai tanpa pengawasan", "Kebijakan & daftar tool resmi, pelatihan, opsi aman yang mudah"],
          ["Ketergantungan", "Keterampilan & kontrol manusia menurun", "Dokumentasi proses, pelatihan, manusia tetap pengambil keputusan"],
      ], row_h=0.56, head_h=0.46, font_size=11.5, bold_first_col=True)

yy = y + 0.46 + 6 * 0.56 + 0.24
rect(s, ML, yy, CW, 0.62, fill=RGBColor(0xFE, 0xF2, 0xF2), line=RED, lw=0.75,
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.16)
tf = textbox(s, ML + 0.3, yy + 0.06, CW - 0.6, 0.5, anchor=MSO_ANCHOR.MIDDLE)
para(tf, [("Prinsip utama:  ", {"bold": True, "color": RED}),
          ("otonomi diberikan bertahap seiring bukti keandalan — bukan sebaliknya. "
           "Mulai dengan mengawasi, baru kurangi pengawasan.", {})],
     size=12.5, color=RGBColor(0x5A, 0x1A, 0x1A), first=True, space_after=0)
notes(s, "Bahaskan minimal tiga risiko paling relevan untuk konteks "
         "organisasi Anda.")

# ================================================================ 20 STUDI KASUS
s, y = content_slide("Studi Kasus", "Laporan Mingguan Otomatis",
                     "Contoh nyata penerapan agent pada pekerjaan berulang tim")
# BEFORE
bwf = (CW - 0.35) / 2
rect(s, ML, y, bwf, 1.35, fill=RGBColor(0xF9, 0xFA, 0xFC), line=BORDER, lw=1,
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
tf = textbox(s, ML + 0.3, y + 0.2, bwf - 0.6, 0.3)
para(tf, "SEBELUM", size=11, bold=True, color=MUTED, first=True, space_after=0)
tf = textbox(s, ML + 0.3, y + 0.56, bwf - 0.6, 0.7)
bullets(tf, ["±4 jam/minggu untuk **copy-paste** dari 3 sistem",
             "Rawan salah ketik & data basi", "Insight datang terlambat"],
        size=11.5, gap=3, bullet_color=MUTED)

rect(s, ML + bwf + 0.35, y, bwf, 1.35, fill=RGBColor(0xEC, 0xFB, 0xF6),
     line=GREEN, lw=1, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
tf = textbox(s, ML + bwf + 0.65, y + 0.2, bwf - 0.6, 0.3)
para(tf, "SESUDAH", size=11, bold=True, color=GREEN, first=True, space_after=0)
tf = textbox(s, ML + bwf + 0.65, y + 0.56, bwf - 0.6, 0.7)
bullets(tf, ["**20 menit** review & approve oleh lead",
             "Anomali terdeteksi otomatis (deviasi >10%)", "Kirim tepat waktu, setiap Jumat"],
        size=11.5, gap=3, bullet_color=GREEN)

y2 = y + 1.62
flow = [
    ("Jumat 16.00", "Agent menarik data dari CRM, Analytics, dan Tiket"),
    ("Analisis", "Bandingkan target vs realisasi, deteksi anomali"),
    ("Draf", "Susun laporan + 3 rekomendasi prioritas"),
    ("Review", "Kirim ke Slack; lead menyetujui/mengedit"),
    ("Arsip", "Simpan ke Notion + catat metrik penggunaan"),
]
fw = (CW - 4 * 0.26) / 5
for i, (t, b) in enumerate(flow):
    fx = ML + i * (fw + 0.26)
    rect(s, fx, y2, fw, 1.75, fill=WHITE, line=BORDER, lw=1,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    rect(s, fx, y2, fw, 0.08, fill=[INDIGO, CYAN, AMBER, GREEN, INDIGO][i],
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
    tf = textbox(s, fx + 0.22, y2 + 0.3, fw - 0.44, 0.3)
    para(tf, t, size=12.5, bold=True, color=INK, first=True, space_after=0)
    tf = textbox(s, fx + 0.22, y2 + 0.68, fw - 0.44, 1.0)
    para(tf, b, size=11, color=RGBColor(0x44, 0x4E, 0x63), first=True,
         space_after=0, line_spacing=1.18)
    if i < 4:
        tf = textbox(s, fx + fw - 0.02, y2 + 0.75, 0.3, 0.3,
                     anchor=MSO_ANCHOR.MIDDLE)
        para(tf, "›", size=18, bold=True, color=MUTED, align=PP_ALIGN.CENTER,
             first=True, space_after=0)

y3 = y2 + 2.02
res = [("70%", "waktu kerja berkurang", INDIGO),
       ("0", "laporan terlambat", CYAN),
       ("100%", "diperiksa manusia", AMBER),
       ("3 jam", "hemat per orang/minggu", GREEN)]
rw3 = (CW - 0.6) / 4
for i, (big, lab, col) in enumerate(res):
    rx = ML + i * (rw3 + 0.2)
    rect(s, rx, y3, rw3, 0.95, fill=LIGHT, line=BORDER, lw=0.75,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
    tf = textbox(s, rx + 0.2, y3 + 0.13, rw3 - 0.4, 0.45, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, big, size=24, bold=True, color=col, first=True, space_after=0)
    tf = textbox(s, rx + 0.2, y3 + 0.58, rw3 - 0.4, 0.3)
    para(tf, lab, size=11, color=MUTED, first=True, space_after=0)
notes(s, "Kunci sukses: manusia tetap menyetujui sebelum laporan dikirim, "
         "dan ada metrik pemakaian yang dicatat tiap pekan.")

# ================================================================ 21 ROADMAP
s, y = content_slide("Rencana Aksi", "Roadmap 30-60-90 Hari",
                     "Tiga fase: bukti nilai → kematangan → skala dan ROI")
phases = [
    ("30 HARI", "Fondasi & Pilot", INDIGO, [
        "Pilih 1–2 use case quick win",
        "Petakan proses & sumber data",
        "Siapkan baseline metrik (waktu, kualitas, biaya)",
        "Tetapkan kebijakan penggunaan & akses data",
        "Jalankan pilot terbatas pada 1 tim",
    ]),
    ("60 HARI", "Integrasi & Evaluasi", CYAN, [
        "Sambungkan tools inti (read → write + approval)",
        "Bangun golden dataset & evaluasi otomatis",
        "Latih 1 tim lengkap + runbook operasional",
        "Pasang logging, dashboard, dan peringatan",
        "Perbaiki prompt berdasarkan kasus gagal",
    ]),
    ("90 HARI", "Skala & ROI", AMBER, [
        "Perluas ke 3–5 use case lintas fungsi",
        "Tetapkan pemilik agent & ritme perawatan",
        "Hitung ROI: jam hemat vs biaya operasional",
        "Bentuk komunitas praktik / center of excellence",
        "Susun rencana tahun berikutnya",
    ]),
]
pw3 = (CW - 0.6) / 3
ph3 = 4.35
for i, (t, sub, col, its) in enumerate(phases):
    cx = ML + i * (pw3 + 0.30)
    rect(s, cx, y, pw3, ph3, fill=WHITE, line=BORDER, lw=1,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    rect(s, cx, y, pw3, 0.85, fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.13)
    rect(s, cx, y + 0.5, pw3, 0.35, fill=col)
    tf = textbox(s, cx + 0.3, y + 0.14, pw3 - 0.6, 0.6, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, t, size=17, bold=True, color=WHITE, first=True, space_after=0)
    para(tf, sub, size=11.5, color=RGBColor(0xF0, 0xF2, 0xFF), space_after=0)
    tf = textbox(s, cx + 0.3, y + 1.08, pw3 - 0.6, 3.1)
    bullets(tf, its, size=12.5, gap=11, bullet_color=col)
notes(s, "Sesuaikan target dengan kapasitas tim — lebih baik tuntas tiga "
         "langkah daripada setengah jalan sepuluh langkah.")

# ================================================================ 22 KESIMPULAN
s, y = content_slide("Penutup", "Kesimpulan & Langkah Berikutnya",
                     "Empat hal yang perlu diingat setelah sesi ini")
kc = [
    ("Agent mengerjakan, bukan sekadar menjawab", "Nilainya ada pada tugas yang selesai — bukan pada jawaban yang terdengar pintar.", INDIGO),
    ("Kunci: use case, data, guardrail, ukuran", "Empat hal ini menentukan sukses atau gagalnya jauh lebih besar daripada pilihan model.", CYAN),
    ("Mulai kecil, ukur, lalu skalakan", "Quick win membangun kepercayaan; metrik membenarkan investasi berikutnya.", AMBER),
    ("Manusia tetap pengambil keputusan akhir", "Otonomi diberikan bertahap seiring bukti keandalan, bukan sebaliknya.", GREEN),
]
kw2 = (CW - 0.4) / 2
kh2 = 1.6
for i, (t, b, col) in enumerate(kc):
    cx = ML + (i % 2) * (kw2 + 0.4)
    cy = y + (i // 2) * (kh2 + 0.3)
    card(s, cx, cy, kw2, kh2, t, b, accent=col, title_size=15, body_size=12)

yy = y + 2 * kh2 + 0.3 + 0.2
rect(s, ML, yy, CW, 1.2, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
rect(s, ML, yy, 0.09, 1.2, fill=AMBER, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
tf = textbox(s, ML + 0.42, yy + 0.2, CW - 0.85, 0.3)
para(tf, "MULAI MINGGU INI", size=10.5, bold=True, color=CYAN, first=True,
     space_after=0)
tf = textbox(s, ML + 0.42, yy + 0.55, CW - 0.85, 0.5)
para(tf, [("1. ", {"bold": True, "color": AMBER}),
          ("Tuliskan 3 tugas paling menyakitkan di tim Anda.   ", {}),
          ("2. ", {"bold": True, "color": AMBER}),
          ("Pilih satu yang paling mudah diukur.   ", {}),
          ("3. ", {"bold": True, "color": AMBER}),
          ("Jalankan pilot 2 minggu dan bandingkan dengan baseline hari ini.", {})],
     size=14, color=WHITE, first=True, space_after=0, line_spacing=1.2)
notes(s, "Tutup dengan komitmen konkret: satu use case, dua minggu, satu "
         "metrik.")

# ================================================================ 23 Q&A
s = dark_slide()
rect(s, 0, 0, 0.32, SH, fill=INDIGO)
rect(s, 0, SH - 2.5, 0.32, 2.5, fill=CYAN)
rect(s, 8.9, 1.5, 3.6, 4.5, fill=NAVY_2, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
rect(s, 9.3, 2.0, 2.8, 0.7, fill=INDIGO, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.22)
rect(s, 9.3, 2.9, 2.1, 0.5, fill=RGBColor(0x2A, 0x39, 0x63),
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
rect(s, 9.3, 3.6, 2.8, 0.5, fill=RGBColor(0x2A, 0x39, 0x63),
     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
rect(s, 9.3, 4.3, 1.3, 1.3, fill=AMBER, shape=MSO_SHAPE.OVAL)
rect(s, 10.9, 4.7, 1.0, 1.0, fill=CYAN, shape=MSO_SHAPE.OVAL)

tf = textbox(s, 1.0, 2.3, 7.4, 0.3)
para(tf, "TERIMA KASIH", size=12, bold=True, color=CYAN, first=True, space_after=0)
tf = textbox(s, 1.0, 2.75, 7.4, 1.3)
para(tf, "Pertanyaan &", size=46, bold=True, color=WHITE, first=True,
     space_after=0, line_spacing=1.0)
para(tf, "Diskusi", size=46, bold=True, color=AMBER, space_after=0, line_spacing=1.0)
rect(s, 1.0, 4.55, 1.6, 0.06, fill=CYAN)
tf = textbox(s, 1.0, 4.85, 7.2, 0.9)
para(tf, "Mari tentukan satu use case yang akan kita coba dalam dua minggu ke depan.",
     size=15, color=RGBColor(0xB9, 0xC2, 0xD6), first=True, space_after=0,
     line_spacing=1.3)
tf = textbox(s, 1.0, 6.3, 7.2, 0.35)
para(tf, "Sesi tanya jawab • 2026", size=11.5, color=MUTED, first=True,
     space_after=0)
notes(s, "Buka sesi diskusi; kumpulkan ide use case peserta di papan.")

# ---------------------------------------------------------------- props
prs.core_properties.title = "Cara Menggunakan AI Agent untuk Pekerjaan"
prs.core_properties.author = "AI Assistant"
prs.core_properties.subject = "Panduan praktis implementasi AI agent di tempat kerja"
prs.core_properties.keywords = "AI agent, produktivitas, otomasi, implementasi"

OUT = "/workspaces/My-Project/Cara-Menggunakan-AI-Agent-untuk-Pekerjaan.pptx"
prs.save(OUT)
print("OK ->", OUT, "| slides:", len(prs.slides))
