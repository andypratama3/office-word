---
name: office-word-mcp
description: Create, read, and edit Microsoft Word (.docx) files via the Office-Word MCP server and smart_write helper. Use when the user mentions Word documents, DOCX read/write, live Word editing, reloading Word, tables, headings, or document automation.
license: MIT
compatibility: opencode
metadata:
  audience: developers
  format: docx
---

# SKILL: Office-Word MCP — Baca/Tulis Word Manual & Live

Skill ini mengajarkan agent cara membuat, membaca, dan mengedit dokumen Word (`.docx`)
lewat **Office-Word-MCP-Server** (protocol MCP) dan helper **`smart_write.py`**.

Repo: hasil fork dari `gongrzhe/office-word-mcp-server` + helper live-edit macOS.

---

## 0. Cara kerja (pahami dulu)

| Jalur | Mekanisme | Cocok untuk |
|-------|-----------|-------------|
| **MCP tools** (`word-document-server`) | `python-docx`: tulis langsung ke file di disk | Semua operasi, termasuk tabel & format |
| **Live mode** (`live_write.py` via AppleScript) | Sisip teks ke dokumen yang **sedang terbuka** di Word for Mac | Append paragraf/heading di ujung dokumen, terlihat mengetik |
| **smart_write.py** | Otomatis pilih jalur + jaga sinkronisasi | **Pakai ini sebagai default** |

> Batasan fundamental: `python-docx` tidak berkomunikasi dengan aplikasi Word.
> Word memuat dokumen ke memori — tulisan via file baru terlihat setelah **reload**.
> Live mode menulis ke memori Word langsung — terlihat seketika, tanpa tutup-buka.

---

## 1. Perintah utama — `smart_write.py`

```bash
export PYTHONPATH=/path/ke/repo-office
VENV=/path/ke/repo-office/.venv-mcp/bin/python

# Tulis live (kalau dokumen terbuka) atau via file (kalau tertutup)
$VENV smart_write.py append FILE.docx "teks paragraf" [--style "Heading 2"] [--delay 1.5] [--watch]
$VENV smart_write.py heading FILE.docx "Judul Bab" [--level 1] [--delay 1] [--watch]
$VENV smart_write.py pagebreak FILE.docx [--watch]

# Operasi file (otomatis save → tulis → reload bila dokumen terbuka)
$VENV smart_write.py table FILE.docx ROWS COLS '[["a","b"],["1","2"]]' [--watch] [--no-reload]
$VENV smart_write.py replace FILE.docx "lama" "baru" [--watch] [--no-reload]
$VENV smart_write.py sync FILE.docx [--watch]  # save + satu reload (penutup batch)

# FILE boleh "auto" = dokumen AKTIF di Word (deteksi otomatis, tanpa hardcode)
$VENV smart_write.py info auto
$VENV reload_word.py            # reload dokumen aktif, background
$VENV reload_word.py FILE.docx --watch   # reload + bawa Word ke depan

# Baca (tidak mengubah apa pun)
$VENV smart_write.py read FILE.docx
$VENV smart_write.py info FILE.docx
```

Batch hemat reload (3 operasi file = 1 reload):

```bash
$VENV smart_write.py table F.docx 2 2 '[["a","b"],["1","2"]]' --no-reload
$VENV smart_write.py replace F.docx "lama" "baru" --no-reload
$VENV smart_write.py sync F.docx
```

Reload itu background murni: Word disembunyikan sesaat selama tukar dokumen
(tanpa kedip bila di background), posisi scroll dikembalikan, fokus aplikasi
tidak dicuri. `--watch` hanya untuk demo yang memang ditonton.

Aturan sinkronisasi (jangan dilawan):

1. Habis **live-write** → `save` dulu sebelum operasi file.
2. Habis **operasi file** saat dokumen terbuka → **reload** dulu sebelum live-write.
3. `smart_write.py` sudah melakukan 1–2 otomatis. Jangan campur pemanggilan
   langsung ke MCP tools + AppleScript manual tanpa mengikuti urutan ini.

Kemampuan live mode: **append di ujung dokumen saja**, style per paragraf
(`Normal`, `Heading 1`–`Heading 9`), page break. Tabel, replace tengah,
format per-kata, dan hapus → lewat jalur file.

---

## 2. MCP server langsung (tanpa smart_write)

Berguna untuk operasi batch/skrip, atau saat Word tidak terlibat.

```bash
PYTHONPATH=/path/ke/repo-office $VENV -c "
import asyncio
from word_document_server.tools import document_tools, content_tools, format_tools
async def main():
    print(await document_tools.create_document('lap.docx', title='Judul', author='Nama'))
    print(await content_tools.add_heading('lap.docx', 'Pendahuluan', level=1))
    print(await content_tools.add_paragraph('lap.docx', 'Isi...'))
    print(await content_tools.add_table('lap.docx', 2, 2, [['A','B'],['1','2']]))
    print(await format_tools.highlight_table_header('lap.docx', 0))
    print(await document_tools.get_document_text('lap.docx'))
asyncio.run(main())
"
```

Semua tool MCP bersifat `async` — bungkus dengan `asyncio.run`.

---

## 3. Mendaftarkan MCP ke OpenCode

`install.sh` melakukannya otomatis. Manual:

```json
{
  "mcp": {
    "word-document-server": {
      "type": "local",
      "command": "/path/ke/repo-office/.venv-mcp/bin/python",
      "args": ["/path/ke/repo-office/word_mcp_server.py"],
      "env": {
        "PYTHONPATH": "/path/ke/repo-office",
        "MCP_TRANSPORT": "stdio"
      },
      "enabled": true
    }
  }
}
```

Restart OpenCode setelah mengubah config.

---

## 4. Anti-patterns

| ❌ Jangan | ✅ Lakukan |
|-----------|-----------|
| Harap tulisan file langsung terlihat di Word yang terbuka | `reload_word.py` / tutup-buka ulang / pakai live mode |
| Mengetik manual di Word saat live-write berjalan | Diam dulu — reload memakai `saving no`, ketikan belum-save hilang |
| Ganti isi tengah dokumen via live mode | `replace` (jalur file) |
| Klaim tabel/format aman tanpa cek | `info` → `table_count`, `read` → verifikasi teks |
| Commit `.venv-mcp/` ke git | Sudah di `.gitignore` — tiap device bikin venv sendiri via `install.sh` |
| Resize gambar asal (distorsi / luber halaman) | Fit-box proporsional (§5.3) |
| Ganti nomor caption tanpa perbaiki acuan di teks | Renumber + ref selalu satu paket (§5.4) |
| Edit XML langsung untuk ganti teks | Replace run-preserving (§5.5) |

---

## 5. Membangun dokumen jurnal dari Markdown + gambar (resep terbukti)

Pola yang dipakai untuk `Jurnal_Andy_Pratama.docx` (4000 kata, 6 tabel, 8 gambar).
Prinsip: **satu skrip build → verifikasi → satu reload di akhir.**

### 5.1 Parse Markdown → python-docx

| Sintaks md | Hasil docx |
|------------|-----------|
| `# ` | Judul 16pt bold tengah |
| `## ` / `### ` | Heading 1 (14pt) / Heading 2 (12pt), Times New Roman |
| `**tebal**`, `*miring*`, `` `kode` `` | Run bold / italic / Consolas 10pt (regex `(\*\*.+?\*\*|\*[^*\n]+?\*|`+"`[^`\n]+?`"+`)`) |
| Tabel `\|` + baris pemisah | `Table Grid`, header shading `4472C4` + teks putih bold |
| Blok ```` ```mermaid ```` | Ekstrak isi `[...]` → paragraf style `List Number` |
| Rumus `$...$` | Ganti jadi teks polos (`x̄ = Σxi / n`) |
| `> ` | Catatan italic 11pt |

Daftar Pustaka: deteksi paragraf pola `Nama. (Tahun).` → *hanging indent*
(`left_indent` 1 cm + `first_line_indent` −1 cm).

### 5.2 Gambar: ekstrak, petakan, verifikasi visual

```python
from docx import Document
d = Document("skripsi.docx")
# 1) симpan semua part gambar
for rel in d.part.rels.values():
    if "image" in rel.target_ref:
        open("fig/" + rel.target_ref.split("/")[-1], "wb").write(rel.target_part.blob)
# 2) urutan kemunculan di dokumen = urutan r:embed di document.xml
import re
embeds = re.findall(r'r:embed="(rId\d+)"', d.part.blob.decode("utf-8"))
```

- Gambar kecil (1–8 KB) biasanya ikon/logo — abaikan; yang besar berurutan
  biasanya cocok dengan urutan caption (Gambar 1–N). **Selalu verifikasi visual**
  minimal gambar kunci sebelum dipakai (salah label = fatal).
- Hanya sisipkan gambar yang **dirujuk di teks**; tiap penyisipan wajib ditemani caption.

### 5.3 Ukuran gambar: fit-box proporsional (A4, margin 3 cm)

Baca dimensi PNG dari header (byte 16–24, tanpa dependensi tambahan),
lalu muatkan ke kotak **maks 5,8 inci × 7,5 inci** (tinggi dibatasi agar
caption tetap sehalaman):

- Aspek ≤ 1,2 (screenshot, grafik): kunci **lebar** 5,8 inci.
- Aspek > 1,2 (flowchart tinggi, ERD): kunci **tinggi** 6,8 inci.
- Jangan set lebar+tinggi manual dari tebakan — selalu hitung dari dimensi asli.

### 5.4 Caption + acuan teks = satu paket

- Caption di bawah gambar, tengah, 11pt, label bold (`Gambar N.` + teks biasa).
- Urutan caption di dokumen **harus** 1–N berurutan — cek via
  `get_document_outline` / baca semua paragraf `Gambar \d+\.`.
- Tiap renumber caption, perbaiki **semua** acuan `(Gambar N)` di teks pada
  skrip yang sama. Sumber salah acuan umum: teks masih memakai nomor gambar
  dokumen asal (skripsi) padahal caption sudah dinomori ulang untuk jurnal.

### 5.5 Ganti teks tanpa menghancurkan format

Fungsi `replace_in_para(p, old, new)`: bila string target dalam **satu run**,
ganti di run itu saja. Bila terpecah antar-run, rebuild paragraf via parser
§5.1 (aman bila paragraf hanya berisi markdown inline). Jangan rebuild
paragraf yang berisi *field*, *hyperlink*, atau format manual kompleks.

### 5.6 Hardening layout (satu skrip, sebelum reload final)

- Body rata kanan-kiri, TNR 12, `space_after` 6pt, spasi tunggal, widow control.
- Heading 1/2: `keep_with_next` + `space_before` 12/6pt (heading tak boleh
  yatim di bawah halaman).
- Paragraf gambar: `keep_with_next` (gambar menempel caption);
  caption `space_before` 6pt + `space_after` 12pt.
- Header tabel: `w:tblHeader` = true (berulang tiap halaman).
- Blok kode: `keep_together` + shading `F2F2F2`.
- Blok judul/penulis dipadatkan (afiliasi 11pt).
- Footer: nomor halaman tengah via field `PAGE`.
- Hapus spasi ganda sisa (kecuali blok kode).
