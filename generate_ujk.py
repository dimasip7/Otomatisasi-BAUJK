import math
import os
from copy import deepcopy
from datetime import datetime

import pandas as pd
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.table import _Cell


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATASET_FILE_ASESOR = os.path.join(BASE_DIR, "dataset_asesor.xlsx")
DATASET_FILE_TUK = os.path.join(BASE_DIR, "dataset_tuk.xlsx")
FOLDER_TEMPLATE = os.path.join(BASE_DIR, "templates")
FOLDER_OUTPUT = os.path.join(BASE_DIR, "output_ujk")

DEFAULT_ROWS = 8
MAX_ASESI_PER_HARI = 12

BULAN_INDO = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]
ASESOR_TAGS = {
    "nama": "NAMA_ASESOR",
    "noreg": "NO_REG_ASESOR",
    "alamat": "ALAMAT_ASESOR",
    "nohp": "NOHP_ASESOR",
}
JUDUL_LIST_DOKUMEN = "LIST DOKUMEN UJI"
PREFIX_NAMA_UMUM = ["i", "ni", "m.", "dr.", "drs.", "ir.", "prof.", "prof.dr.", "prof.dr.ir.", "muhammad", "muhamad", "mohammad", "mohamad"]

TEMPLATE_UMUM_BERPASANGAN = [
    "01. PERNYATAAN KERAHASIAAN.docx",
    "03. SURAT PENYATAAN KOMITMEN ASESOR.docx",
    "04. SURAT TUGAS ASESOR KOMPETENSI.docx",
    "05. DAFTAR HADIR ASESOR.docx",
    "06. DAFTAR HADIR ASESI.docx",
    "07. BERITA ACARA ASESMEN.docx",
]


def _load_dataset(path, columns):
    if os.path.exists(path):
        return pd.read_excel(path)
    df = pd.DataFrame(columns=columns)
    df.to_excel(path, index=False)
    return df


def _append_and_save(df, entry, sort_col, path):
    df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
    df = df.sort_values(by=sort_col, key=lambda col: col.str.lower()).reset_index(drop=True)
    df["No"] = range(1, len(df) + 1)
    df.to_excel(path, index=False)


def load_dataset_asesor():
    return _load_dataset(DATASET_FILE_ASESOR, ["No", "Nama", "No Reg", "Alamat", "No HP"])


def save_asesor_baru(nama, no_reg, alamat, nohp):
    entry = {"No": 0, "Nama": nama, "No Reg": no_reg, "Alamat": alamat, "No HP": nohp}
    _append_and_save(load_dataset_asesor(), entry, "Nama", DATASET_FILE_ASESOR)
    print(f"[+] Data asesor '{nama}' berhasil disimpan ke dataset '{os.path.basename(DATASET_FILE_ASESOR)}'.")


def pilih_atau_input_asesor():
    df = load_dataset_asesor()

    print("\n--- PILIH / INPUT DATA ASESOR ---")
    if not df.empty:
        print("\nDaftar Asesor yang Tersedia:")
        for idx, row in df.iterrows():
            print(f" {idx + 1}. {row['Nama']} (No Reg: {row['No Reg']})")
        print(f" {len(df) + 1}. Input Data Asesor Baru")

        pilihan = input(f"\nPilih nomor asesor (1-{len(df) + 1}): ").strip()
        if pilihan.isdigit() and 1 <= int(pilihan) <= len(df):
            row = df.iloc[int(pilihan) - 1]
            return {
                "nama": str(row["Nama"]),
                "noreg": str(row["No Reg"]),
                "alamat": str(row["Alamat"]),
                "nohp": str(row["No HP"]),
            }

    print("\n[Info] Input Data Asesor Baru:")
    nama = input("Nama Asesor: ").strip()
    noreg = input("No. Reg Asesor: ").strip()
    alamat = input("Alamat Asesor: ").strip()
    nohp = input("No. HP Asesor: ").strip()

    save_asesor_baru(nama, noreg, alamat, nohp)
    return {"nama": nama, "noreg": noreg, "alamat": alamat, "nohp": nohp}


def input_daftar_asesor():
    """Tanya jumlah asesor (1/2), lalu pilih/input data tiap asesor. Mengembalikan list asesor."""
    while True:
        jumlah = input("\nJumlah Asesor (1/2) [default 1]: ").strip() or "1"
        if jumlah in ("1", "2"):
            break
        print("[!] Masukkan 1 atau 2.")

    daftar = []
    for n in range(1, int(jumlah) + 1):
        if jumlah == "2":
            print(f"\n>>> ASESOR {n} dari 2")
        while True:
            asesor = pilih_atau_input_asesor()
            if asesor["nama"] not in (a["nama"] for a in daftar):
                break
            print("[!] Asesor ini sudah dipilih. Pilih asesor yang berbeda.")
        daftar.append(asesor)
    return daftar


def load_dataset_tuk():
    return _load_dataset(DATASET_FILE_TUK, ["No", "Nama TUK", "Alamat TUK", "Kota TUK", "Zona Waktu"])


def save_tuk_baru(nama_tuk, alamat_tuk, kota_tuk, zona_waktu):
    entry = {
        "No": 0,
        "Nama TUK": nama_tuk,
        "Alamat TUK": alamat_tuk,
        "Kota TUK": kota_tuk,
        "Zona Waktu": zona_waktu,
    }
    _append_and_save(load_dataset_tuk(), entry, "Zona Waktu", DATASET_FILE_TUK)
    print(f"[+] Data TUK '{nama_tuk}' berhasil disimpan ke dataset '{os.path.basename(DATASET_FILE_TUK)}'.")


def pilih_atau_input_tuk():
    df = load_dataset_tuk()

    print("\n--- PILIH / INPUT DATA TUK ---")
    if not df.empty:
        print("\nDaftar TUK yang Tersedia:")
        for idx, row in df.iterrows():
            print(f" {idx + 1}. {row['Nama TUK']}")
        print(f" {len(df) + 1}. Input Data TUK Baru")

        pilihan = input(f"\nPilih nomor TUK (1-{len(df) + 1}): ").strip()
        if pilihan.isdigit() and 1 <= int(pilihan) <= len(df):
            row = df.iloc[int(pilihan) - 1]
            return {
                "nama_tuk": str(row["Nama TUK"]),
                "alamat_tuk": str(row["Alamat TUK"]),
                "kota_tuk": str(row["Kota TUK"]),
                "zona_waktu": str(row["Zona Waktu"]),
            }

    print("\n[Info] Input Data TUK Baru:")
    nama_tuk = input("Nama TUK: ").strip()
    alamat_tuk = input("Alamat TUK: ").strip()
    kota_tuk = input("Kota TUK: ").strip()
    zona_waktu = input("Zona Waktu (WIB/WITA): ").strip().upper()

    save_tuk_baru(nama_tuk, alamat_tuk, kota_tuk, zona_waktu)
    return {
        "nama_tuk": nama_tuk,
        "alamat_tuk": alamat_tuk,
        "kota_tuk": kota_tuk,
        "zona_waktu": zona_waktu,
    }


def format_tanggal_indo(dt):
    return f"{dt.day:02d} {BULAN_INDO[dt.month - 1]} {dt.year}"


def split_asesi_balanced(list_asesi, force=False):
    """Bagi asesi menjadi 2 kelompok bila force=True (2 asesor) atau jumlah melebihi MAX_ASESI_PER_HARI."""
    if not force and len(list_asesi) <= MAX_ASESI_PER_HARI:
        return list_asesi, []
    half = math.ceil(len(list_asesi) / 2)
    return list_asesi[:half], list_asesi[half:]


def asesor_placeholders(asesor, suffix=""):
    return {f"{{{{ {tag}{suffix} }}}}": asesor[key] for key, tag in ASESOR_TAGS.items()}


def upper_values(data):
    return {k: (v.upper() if isinstance(v, str) else v) for k, v in data.items()}


def build_asesi_data(base_data, hari, tanggal, asesi_list):
    data = base_data.copy()
    data["{{ HARI_UJK }}"] = hari
    data["{{ TANGGAL_UJK }}"] = tanggal
    for idx in range(1, max(DEFAULT_ROWS, len(asesi_list)) + 1):
        data[f"{{{{ ASESI_{idx} }}}}"] = asesi_list[idx - 1] if idx <= len(asesi_list) else "-"
    return data


def buat_nama_file_asesor(nama_asesor):
    """Ambil nama (maksimal 3 kata jika diawali gelar) yang aman untuk nama file Windows."""
    words = nama_asesor.strip().split()
    if len(words) > 1 and words[0].lower() in PREFIX_NAMA_UMUM:
        nama = " ".join(words[:3])
    else:
        nama = words[0] if words else "ASESOR"
    return "".join(c for c in nama if c.isalnum() or c in (" ", "_", "-")).strip()


def set_cell_background(cell, fill_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def replace_placeholder_in_paragraph(paragraph, data_dict):
    """Ganti placeholder sambil mempertahankan format run bawaan template."""
    keys = [k for k in data_dict if k in paragraph.text]
    if not keys:
        return

    for run in paragraph.runs:
        for key in keys:
            if key in run.text:
                run.text = run.text.replace(key, str(data_dict[key]))

    # Gabungkan ke run pertama bila tag placeholder terpecah antar run oleh Word.
    if paragraph.runs and any(k in paragraph.text for k in keys):
        new_text = paragraph.text
        for key, val in data_dict.items():
            new_text = new_text.replace(key, str(val))
        for run in paragraph.runs:
            run.text = ""
        paragraph.runs[0].text = new_text


def _gray_out_empty_row(row):
    """Kosongkan penanda '-' dan beri latar abu-abu kecuali pada sel vMerge restart."""
    for cell in row.cells:
        vmerge = cell._tc.get_or_add_tcPr().find(qn("w:vMerge"))
        if vmerge is not None and vmerge.get(qn("w:val")) == "restart":
            continue

        for p in cell.paragraphs:
            if p.text.strip() == "-":
                p.text = ""
        set_cell_background(cell, "D3D3D3")


def replace_placeholders_in_doc(doc, data_dict):
    for p in doc.paragraphs:
        replace_placeholder_in_paragraph(p, data_dict)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholder_in_paragraph(p, data_dict)

            if any(cell.text.strip() == "-" for cell in row.cells):
                _gray_out_empty_row(row)


def _elem_text(el):
    """Baca teks elemen XML secara langsung tanpa menelusuri sel merge."""
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _find_row_idx(table, marker):
    for i, tr in enumerate(table._tbl.tr_lst):
        if marker in _elem_text(tr):
            return i
    return None


def _replace_in_tr(tr, old, new):
    for p in tr.iter(qn("w:p")):
        ts = list(p.iter(qn("w:t")))
        if not ts:
            continue
        full = "".join(t.text or "" for t in ts)
        if old not in full:
            continue
        ts[0].text = full.replace(old, new)
        for t in ts[1:]:
            t.text = ""


def _is_ttd_table(table):
    return any("tanda tangan" in _elem_text(tr).lower() for tr in table._tbl.tr_lst[:3])


def _set_vmerge(tc, mode):
    tcPr = tc.get_or_add_tcPr()
    for el in tcPr.findall(qn("w:vMerge")):
        tcPr.remove(el)
    if mode is None:
        return

    vm = OxmlElement("w:vMerge")
    if mode == "restart":
        vm.set(qn("w:val"), "restart")

    # Urutan skema OOXML mengharuskan vMerge ditempatkan setelah tcW, gridSpan, atau hMerge.
    anchor = next(
        (el for el in (tcPr.find(qn(tag)) for tag in ("w:hMerge", "w:gridSpan", "w:tcW")) if el is not None),
        None,
    )
    if anchor is not None:
        anchor.addnext(vm)
    else:
        tcPr.insert(0, vm)


def _find_ref_rpr(trs, ttd_cols=2):
    for tr in trs:
        for tc in tr.tc_lst[-ttd_cols:]:
            for r in tc.iter(qn("w:r")):
                rpr = r.find(qn("w:rPr"))
                if rpr is not None:
                    return deepcopy(rpr)

    for tr in trs:
        for tc in tr.tc_lst[-ttd_cols:]:
            for p in tc.iter(qn("w:p")):
                ppr = p.find(qn("w:pPr"))
                rpr = ppr.find(qn("w:rPr")) if ppr is not None else None
                if rpr is not None:
                    return deepcopy(rpr)
    return None


def _set_cell_text(tc, table, text, ref_rpr=None):
    cell = _Cell(tc, table)
    for p in cell.paragraphs[1:]:
        p._p.getparent().remove(p._p)

    para = cell.paragraphs[0]
    if para.runs:
        para.runs[0].text = text
        for r in para.runs[1:]:
            r._r.getparent().remove(r._r)
        return

    run = para.add_run(text)
    if ref_rpr is not None:
        old = run._r.find(qn("w:rPr"))
        if old is not None:
            run._r.remove(old)
        run._r.insert(0, deepcopy(ref_rpr))


def _write_ttd_cell(tc, table, number, real_count, align, ref_rpr):
    text = f"{number}." if number <= real_count else ""
    _set_cell_text(tc, table, text, ref_rpr)
    _Cell(tc, table).vertical_alignment = align


def rebuild_ttd_columns(table, first_row_idx, count, ttd_cols=2, real_count=None):
    """Gabungkan vertikal tiap pasangan 2 baris TTD (ganjil di kiri atas, genap di kanan bawah)."""
    real_count = count if real_count is None else real_count

    trs = table._tbl.tr_lst
    rows = [trs[first_row_idx + i] for i in range(count)]
    ref_rpr = _find_ref_rpr(rows, ttd_cols)
    ttd = [tr.tc_lst[-ttd_cols:] for tr in rows]

    for cells in ttd:
        for tc in cells:
            _set_vmerge(tc, None)
            _set_cell_text(tc, table, "", ref_rpr)

    for start in range(0, count, 2):
        (left_top, right_top), (left_bot, right_bot) = ttd[start], ttd[start + 1]

        _write_ttd_cell(left_top, table, start + 1, real_count, WD_ALIGN_VERTICAL.TOP, ref_rpr)
        _write_ttd_cell(right_top, table, start + 2, real_count, WD_ALIGN_VERTICAL.BOTTOM, ref_rpr)

        for top, bottom in ((left_top, left_bot), (right_top, right_bot)):
            _set_vmerge(top, "restart")
            _set_vmerge(bottom, "continue")


def expand_rows(doc, asesi_list, with_ttd=None):
    """Gandakan baris terakhir tabel asesi bila jumlah peserta melebihi DEFAULT_ROWS."""
    target_count = len(asesi_list)
    if target_count <= DEFAULT_ROWS:
        return

    for table in doc.tables:
        tmpl_idx = _find_row_idx(table, f"{{{{ ASESI_{DEFAULT_ROWS} }}}}")
        first_idx = _find_row_idx(table, "{{ ASESI_1 }}")
        if tmpl_idx is None or first_idx is None:
            continue

        use_ttd = _is_ttd_table(table) if with_ttd is None else with_ttd
        row_count = target_count + (target_count % 2) if use_ttd else target_count

        template_tr = table._tbl.tr_lst[tmpl_idx]
        template_marker = f"{{{{ ASESI_{DEFAULT_ROWS} }}}}"
        insert_after = template_tr

        for num in range(DEFAULT_ROWS + 1, row_count + 1):
            new_tr = deepcopy(template_tr)
            insert_after.addnext(new_tr)
            insert_after = new_tr

            replacement = "-" if num > target_count else f"{{{{ ASESI_{num} }}}}"
            _replace_in_tr(new_tr, template_marker, replacement)

        if use_ttd:
            rebuild_ttd_columns(table, first_idx, row_count, ttd_cols=2, real_count=target_count)


def save_document(doc, path_out):
    try:
        if os.path.exists(path_out):
            try:
                os.remove(path_out)
            except PermissionError:
                print(f"  [!] File sedang terbuka, pastikan ditutup: {os.path.basename(path_out)}")
                return False
        doc.save(path_out)
        return True
    except PermissionError:
        print(f"  [!] Gagal menyimpan (Permission Denied): {os.path.basename(path_out)}")
        print("      Kemungkinan file sedang terbuka di program lain. Tutup file terlebih dahulu.")
        return False
    except Exception as e:
        print(f"  [!] Error saat menyimpan: {e}")
        return False


def hapus_halaman_2(doc):
    """Hapus halaman kedua LIST DOKUMEN UJI beserta page break pemisahnya."""
    body = doc.element.body
    children = [el for el in body if el.tag != qn("w:sectPr")]
    judul = [
        i for i, el in enumerate(children)
        if el.tag == qn("w:p") and _elem_text(el).strip().upper().startswith(JUDUL_LIST_DOKUMEN)
    ]
    if len(judul) < 2:
        print("  [!] Judul halaman 2 tidak ditemukan di template, halaman 2 tidak dihapus.")
        return

    start = judul[1]
    while start > 0 and children[start - 1].tag == qn("w:p") and not _elem_text(children[start - 1]).strip():
        start -= 1

    for el in children[start:]:
        body.remove(el)

    last = children[start - 1] if start else None
    if last is not None and last.tag == qn("w:p"):
        for br in [b for b in last.iter(qn("w:br")) if b.get(qn("w:type")) == "page"]:
            br.getparent().remove(br)
        ppr = last.find(qn("w:pPr"))
        sect = ppr.find(qn("w:sectPr")) if ppr is not None else None
        if sect is not None:
            ppr.remove(sect)
    else:
        # OOXML mewajibkan elemen paragraf setelah tabel sebelum sectPr penutup.
        spacer = parse_xml(
            f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
            '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
        )
        sect_body = body.find(qn("w:sectPr"))
        if sect_body is not None:
            sect_body.addprevious(spacer)
        else:
            body.append(spacer)


def _generate_document(path_in, data, asesi_list, output_filename, halaman_2=True):
    doc = Document(path_in)
    if not halaman_2:
        hapus_halaman_2(doc)
    if asesi_list:
        expand_rows(doc, asesi_list)
    replace_placeholders_in_doc(doc, data)

    if save_document(doc, os.path.join(FOLDER_OUTPUT, output_filename)):
        print(f"  [+] Berhasil dibuat: {output_filename}")


def process_template_file(tmpl_filename, data_dict, asesi_list_for_check=None,
                          is_paired=False, nama_file_asesor=("", ""), halaman_2=True):
    """Render satu template tunggal atau sepasang dokumen per kelompok asesi."""
    path_in = os.path.join(FOLDER_TEMPLATE, tmpl_filename)
    if not os.path.exists(path_in):
        print(f"  [!] Template tidak ditemukan: {path_in}")
        return

    base_name, ext = os.path.splitext(os.path.basename(tmpl_filename))

    if not is_paired:
        _generate_document(path_in, data_dict, asesi_list_for_check, f"{base_name}{ext}", halaman_2)
        return

    data_1, data_2, buat_kelompok_2 = data_dict
    asesi_1, asesi_2 = asesi_list_for_check or (None, None)
    asesor_1, asesor_2 = nama_file_asesor

    _generate_document(path_in, data_1, asesi_1, f"{base_name} - {asesor_1} 1{ext}")
    if buat_kelompok_2:
        _generate_document(path_in, data_2, asesi_2, f"{base_name} - {asesor_2} 2{ext}")


def process_skema_templates(kode_skema, data_1, data_2, buat_kelompok_2,
                            asesi_1, asesi_2, nama_file_asesor):
    """Render berkas .docx di folder skema (berkas rekap dibuat per kelompok, lainnya sekali)."""
    path_folder_skema = os.path.join(FOLDER_TEMPLATE, "skema", kode_skema)
    if not os.path.isdir(path_folder_skema):
        return

    templates_skema = sorted(
        f for f in os.listdir(path_folder_skema) if f.lower().endswith(".docx")
    )
    if not templates_skema:
        print(f"[!] Warning: Tidak ada file .docx di folder skema '{kode_skema}'.")

    for file_name in templates_skema:
        relative_template = os.path.join("skema", kode_skema, file_name)
        if "rekap" in file_name.lower():
            process_template_file(
                relative_template,
                (data_1, data_2, buat_kelompok_2),
                asesi_list_for_check=(asesi_1, asesi_2),
                is_paired=True,
                nama_file_asesor=nama_file_asesor,
            )
        else:
            process_template_file(relative_template, data_1)


def input_daftar_asesi():
    print("\n--- [3/3] INPUT DAFTAR ASESI ---")
    print("Masukkan nama asesi satu per satu. Tekan ENTER tanpa mengetik nama untuk SELESAI.")

    list_asesi = []
    while True:
        nama = input(f"Nama Asesi ke-{len(list_asesi) + 1}: ").strip()
        if not nama:
            return list_asesi
        list_asesi.append(nama)


def input_kode_skema():
    folder_skema_base = os.path.join(FOLDER_TEMPLATE, "skema")
    skema_tersedia = [
        f for f in os.listdir(folder_skema_base)
        if os.path.isdir(os.path.join(folder_skema_base, f))
    ]

    print("\nPilihan Kode Skema yang Tersedia:")
    for idx, s in enumerate(skema_tersedia, 1):
        print(f" {idx}. {s}")

    kode_skema = input("\nMasukkan Kode/Folder Skema (contoh: AC / SSMI): ").strip().upper()
    if not os.path.exists(os.path.join(folder_skema_base, kode_skema)):
        print(f"[!] Warning: Folder skema '{kode_skema}' tidak ditemukan. "
              "Hanya template umum yang akan diproses.")
    return kode_skema


def main():
    print("==================================================")
    print("      OTOMATISASI PENGISIAN BERITA ACARA UJK      ")
    print("==================================================\n")

    print("--- [1/3] INPUT DATA UMUM, TUK & ASESOR ---")
    kode_skema = input_kode_skema()
    nama_skema = input("Nama Skema : ").strip()
    data_tuk = pilih_atau_input_tuk()
    nama_kejuruan = input("Nama Kejuruan (contoh: Las): ").strip()
    daftar_asesor = input_daftar_asesor()
    dua_asesor = len(daftar_asesor) == 2
    asesor_1 = daftar_asesor[0]
    asesor_2 = daftar_asesor[1] if dua_asesor else asesor_1
    nama_file_asesor = (buat_nama_file_asesor(asesor_1["nama"]),
                        buat_nama_file_asesor(asesor_2["nama"]))

    print("\n--- [2/3] INPUT TANGGAL PELAKSANAAN & SURAT TUGAS ---")
    tgl_hari_ini = format_tanggal_indo(datetime.now())
    tgl_surat_tugas = input(
        f"Tanggal Surat Tugas dibuat (Default Hari Ini: {tgl_hari_ini}): "
    ).strip() or tgl_hari_ini
    nost_1 = input("Nomor Surat Tugas: ").strip()

    hari_1 = input("Hari UJK - Hari 1 (contoh: Selasa): ").strip()
    tgl_1 = input("Tanggal UJK - Hari 1 (contoh: 15 September 2026): ").strip()

    list_asesi = input_daftar_asesi()
    if not list_asesi:
        print("\n[!] Error: Daftar asesi tidak boleh kosong.")
        return

    asesi_1, asesi_2 = split_asesi_balanced(list_asesi, force=dua_asesor)
    buat_kelompok_2 = bool(asesi_2)
    nost_2 = nost_1

    if dua_asesor:
        print(f"\n[Info] 2 asesor: UJK 1 hari, total asesi ({len(list_asesi)} orang) "
              "dibagi per asesor secara seimbang:")
        for asesor, kelompok in zip(daftar_asesor, (asesi_1, asesi_2)):
            print(f"  - {asesor['nama']}: {len(kelompok)} orang")
        hari_2, tgl_2 = hari_1, tgl_1
        nost_2 = input("Nomor Surat Tugas - Asesor 2: ").strip()
    elif buat_kelompok_2:
        print(f"\n[Info] Total asesi ({len(list_asesi)} orang) > {MAX_ASESI_PER_HARI}, "
              "dibagi 2 hari secara seimbang:")
        print(f"  - Hari 1: {len(asesi_1)} orang")
        print(f"  - Hari 2: {len(asesi_2)} orang")
        nost_2 = input("Nomor Surat Tugas - Hari 2: ").strip()
        hari_2 = input("Hari UJK - Hari 2 (contoh: Rabu): ").strip()
        tgl_2 = input("Tanggal UJK - Hari 2 (contoh: 16 September 2026): ").strip()
    else:
        print(f"\n[Info] Total asesi ({len(list_asesi)} orang) <= {MAX_ASESI_PER_HARI}, "
              "hanya diproses untuk 1 HARI UJK.")
        hari_2, tgl_2 = hari_1, tgl_1

    common_data = {
        "{{ NAMA_SKEMA }}": nama_skema,
        "{{ NAMA_KEJURUAN }}": nama_kejuruan,
        "{{ NAMA_TUK }}": data_tuk["nama_tuk"],
        "{{ ALAMAT_TUK }}": data_tuk["alamat_tuk"],
        "{{ KOTA_TUK }}": data_tuk["kota_tuk"],
        "{{ ZONA_WAKTU }}": data_tuk["zona_waktu"],
        "{{ TANGGAL_SURAT_TUGAS }}": tgl_surat_tugas,
    }

    data_1 = build_asesi_data({**common_data, **asesor_placeholders(asesor_1), "{{ NO_SURAT_TUGAS }}": nost_1}, hari_1, tgl_1, asesi_1)
    data_2 = build_asesi_data({**common_data, **asesor_placeholders(asesor_2), "{{ NO_SURAT_TUGAS }}": nost_2}, hari_2, tgl_2, asesi_2)

    # LIST DOKUMEN UJI memakai huruf kapital dengan tag asesor berindeks per halaman (_1, _2).
    data_combined_upper = upper_values({**common_data, **asesor_placeholders(asesor_1)})
    for n, asesor in enumerate((asesor_1, asesor_2), 1):
        data_combined_upper.update(upper_values(asesor_placeholders(asesor, f"_{n}")))
    data_combined_upper.update(upper_values({
        "{{ HARI_UJK_1 }}": hari_1,
        "{{ TANGGAL_UJK_1 }}": tgl_1,
        "{{ HARI_UJK_2 }}": hari_2,
        "{{ TANGGAL_UJK_2 }}": tgl_2,
    }))

    os.makedirs(FOLDER_OUTPUT, exist_ok=True)
    print("\nMemproses pembuatan dokumen...")

    for tmpl in TEMPLATE_UMUM_BERPASANGAN:
        process_template_file(
            os.path.join("umum", tmpl),
            (data_1, data_2, buat_kelompok_2),
            asesi_list_for_check=(asesi_1, asesi_2),
            is_paired=True,
            nama_file_asesor=nama_file_asesor,
        )

    process_skema_templates(kode_skema, data_1, data_2, buat_kelompok_2,
                            asesi_1, asesi_2, nama_file_asesor)

    process_template_file(os.path.join("umum", "LIST DOKUMEN UJI.docx"), data_combined_upper,
                          halaman_2=buat_kelompok_2)

    print("\n==================================================")
    print(f" SELESAI! Seluruh file telah disimpan di folder '{FOLDER_OUTPUT}'")
    print("==================================================")


if __name__ == "__main__":
    main()