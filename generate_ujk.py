import os
import math
import pandas as pd
from datetime import datetime
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATASET_FILE = os.path.join(BASE_DIR, "dataset_asesor.xlsx")


TIMEZONE_MAP = {
    "WIB": [
        "Banyuwangi", "Bondowoso", "Jember", "Situbondo", "Pasuruan", "Probolinggo"
    ],
    "WITA": [
        "Denpasar", "Gianyar", "Jembrana", "Karangasem", "Buleleng"
    ]
}

# 1. PENGELOLAAN DATASET ASESOR
def load_dataset():
    if os.path.exists(DATASET_FILE):
        df = pd.read_excel(DATASET_FILE)
    else:
        df = pd.DataFrame(columns=["No", "Nama", "No Reg", "Alamat", "No HP"])
        df.to_excel(DATASET_FILE, index=False)
    return df

def save_asesor_baru(nama, no_reg, alamat, nohp):
    df = load_dataset()

    new_entry = {
        "No": len(df) + 1,
        "Nama": nama,
        "No Reg": no_reg,
        "Alamat": alamat,
        "No HP": nohp
    }

    df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)

    df.to_excel(DATASET_FILE, index=False)
    print(f"[+] Data asesor '{nama}' berhasil disimpan ke dataset '{os.path.basename(DATASET_FILE)}'.")
    return new_entry

def pilih_atau_input_asesor():
    df = load_dataset()

    print ("\n--- PILIH / INPUT DATA ASESOR ---")
    if not df.empty:
        print("\nDaftar Asesor yang Tersedia:")
        for idx, row in df.iterrows():
            print(f" {idx + 1}. {row['Nama']} (No Reg: {row['No Reg']})")
        print(f" {len(df) + 1}. Input Data Asesor Baru")

        pilihan = input(f"\nPilih nomor asesor (1-{len(df) + 1}): ").strip()

        if pilihan.isdigit() and 1 <= int(pilihan) <= len(df):
            selected_row = df.iloc[int(pilihan) - 1]
            return {
                "nama": str(selected_row["Nama"]),
                "noreg": str(selected_row["No Reg"]),
                "alamat": str(selected_row["Alamat"]),
                "nohp": str(selected_row["No HP"])
            }

    print("\n[Info] Input Data Asesor Baru:")
    nama = input("Nama Asesor: ").strip()  
    noreg = input("No. Reg Asesor: ").strip()
    alamat = input("Alamat Asesor: ").strip()
    nohp = input("No. HP Asesor: ").strip()

    save_asesor_baru(nama, noreg, alamat, nohp)
    return {
        "nama": nama,
        "noreg": noreg,
        "alamat": alamat,
        "nohp": nohp
    }


# 2. HELPER UTILS & FORMATTING
def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def detect_timezone(text_location):
    text_clean = text_location.strip().lower()
    for tz, locations in TIMEZONE_MAP.items():
        for loc in locations:
            if loc.lower() in text_clean:
                return tz
    return "WIB"  # Default timezone if not found

def format_tanggal_indo(dt):
    """
    Mengubah datetime menjadi format bahasa Indonesia: DD MMMM YYYY
    Contoh: 30 Agustus 2026
    """
    bulan_indo = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    return f"{dt.day:02d} {bulan_indo[dt.month - 1]} {dt.year}"

def replace_placeholder_in_paragraph(paragraph, data_dict):
    """
    Mengganti tag placeholder dengan mempertahankan 100% format bawaan template Word 
    (termasuk ALL CAPS, Bold, Underline, Ukuran Font 18pt/12pt, dsb).
    """
    full_text = paragraph.text
    # Cek apakah paragraf ini mengandung salah satu key placeholder
    keys_to_replace = [k for k in data_dict.keys() if k in full_text]
    if not keys_to_replace:
        return

    # Jika placeholder berada dalam satu 'run' tunggal
    for run in paragraph.runs:
        for key in keys_to_replace:
            if key in run.text:
                run.text = run.text.replace(key, str(data_dict[key]))

    # Jika placeholder terpecah di antara beberapa 'run'
    # Periksa kembali apakah masih ada tag yang tersisa di paragraph.text
    if any(key in paragraph.text for key in keys_to_replace):
        # Ambil format dan style dari run pertama
        if paragraph.runs:
            first_run = paragraph.runs[0]
            font_name = first_run.font.name
            font_size = first_run.font.size
            is_bold = first_run.bold
            is_italic = first_run.italic
            is_underline = first_run.underline
            is_all_caps = first_run.font.all_caps

            # Lakukan replace pada teks utuh
            new_text = paragraph.text
            for key, val in data_dict.items():
                if key in new_text:
                    new_text = new_text.replace(key, str(val))

            # Kosongkan semua run dan tumpahkan teks baru ke run pertama
            for run in paragraph.runs:
                run.text = ""

            first_run.text = new_text
            if font_name: first_run.font.name = font_name
            if font_size: first_run.font.size = font_size
            if is_bold is not None: first_run.bold = is_bold
            if is_italic is not None: first_run.italic = is_italic
            if is_underline is not None: first_run.underline = is_underline
            if is_all_caps is not None: first_run.font.all_caps = is_all_caps

def replace_placeholders_in_doc(doc, data_dict):
    """
    Mengganti seluruh placeholder {{ TAG }} di paragraf maupun tabel Word.
    """
    # 1. Paragraf Utama
    for p in doc.paragraphs:
        replace_placeholder_in_paragraph(p, data_dict)

    # 2. Tabel
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholder_in_paragraph(p, data_dict)

            has_empty_marker = any(cell.text.strip() == "-" for cell in row.cells)

            if has_empty_marker:
                for cell in row.cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    vMerge = tcPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}vMerge')

                    vMerge_val = vMerge.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if vMerge is not None else None

                    if vMerge_val == "restart":
                        continue
                    
                    for p in cell.paragraphs:
                        if p.text.strip() == "-":
                            p.text = ""  # Kosongkan teks
                        # Jika cell kosong, beri background abu-abu muda
                    set_cell_background(cell, "D3D3D3")  # Hex untuk abu-abu muda
                

def split_asesi_balanced(list_asesi):
    """
    Membagi daftar asesi menjadi 2 hari secara seimbang (maksimal 8 per hari).
    """
    total = len(list_asesi)
    if total <= 8:
        return list_asesi, []
    
    half = math.ceil(total / 2)
    hari1 = list_asesi[:half]
    hari2 = list_asesi[half:]
    return hari1, hari2

def extract_default_city(nama_tuk):
    """
    Mengambil kata terakhir dari nama TUK sebagai perkiraan nama kota.
    """
    words = nama_tuk.strip().split()
    return words[-1] if words else ""

def build_asesi_data(base_data, hari, tanggal, asesi_list):
    """
    Membangun dictionary data dengan hari, tanggal, dan daftar asesi.
    """
    data = base_data.copy()
    data["{{ HARI_UJK }}"] = hari
    data["{{ TANGGAL_UJK }}"] = tanggal
    for idx in range(1, 9):
        val = asesi_list[idx-1] if idx <= len(asesi_list) else "-"
        data[f"{{{{ ASESI_{idx} }}}}"] = val
    return data


# 3. MAIN FUNCTION
def main():
    print("==================================================")
    print("      OTOMATISASI PENGISIAN BERITA ACARA UJK      ")
    print("==================================================\n")

    # --- 1. INPUT DATA UMUM & TUK ---
    print("--- [1/3] INPUT DATA UMUM, TUK & ASESOR ---")
    # Tampilkan daftar skema yang tersedia berdasarkan folder di 'templates/skema'
    folder_skema_base = os.path.join(BASE_DIR, "templates", "skema")
    skema_tersedia = [f for f in os.listdir(folder_skema_base) if os.path.isdir(os.path.join(folder_skema_base, f))]
    
    print("\nPilihan Kode Skema yang Tersedia:")
    for idx, s in enumerate(skema_tersedia, 1):
        print(f" {idx}. {s}")
        
    kode_skema = input("\nMasukkan Kode/Folder Skema (contoh: SSMI / LAS): ").strip().upper()
    
    # Validasi jika folder skema tidak ada
    dir_skema_terpilih = os.path.join(folder_skema_base, kode_skema)
    if not os.path.exists(dir_skema_terpilih):
        print(f"[!] Warning: Folder skema '{kode_skema}' tidak ditemukan. Menggunakan template default.")

    nama_skema = input("Nama Skema : ").strip()
    nama_tuk = input("Nama TUK (contoh: UPTD BLK Jember): ").strip()
    alamat_tuk = input("Alamat TUK: ").strip()
    
    default_kota = extract_default_city(nama_tuk)
    kota_input = input(f"Kota Lokasi TUK (Default: {default_kota}): ").strip()
    kota_tuk = kota_input if kota_input else default_kota

    detected_timezone = detect_timezone(f"{kota_tuk}")
    tz_input = input(f"Zona Waktu (Default: {detected_timezone}): ").strip().upper()
    zona_waktu = tz_input if tz_input in TIMEZONE_MAP else detected_timezone

    nama_kejuruan = input("Nama Kejuruan (contoh: Las): ").strip()

    data_asesor = pilih_atau_input_asesor()
    nama_asesor = data_asesor["nama"]
    no_reg_asesor = data_asesor["noreg"]
    alamat_asesor = data_asesor["alamat"]
    nohp_asesor = data_asesor["nohp"]

    # Sanitasi Nama Asesor untuk Penamaan File (Menghilangkan Karakter Ilegal File Windows)
    nama_depan_asesor = nama_asesor.strip().split()[0] if nama_asesor.strip() else "ASESOR"
    clean_nama_asesor = "".join([c for c in nama_depan_asesor if c.isalnum() or c in (' ', '_', '-')]).strip()

    # --- 2. INPUT DATA HARI & TANGGAL UJK ---
    print("\n--- [2/3] INPUT TANGGAL PELAKSANAAN & SURAT TUGAS ---")
    
    tgl_hari_ini = format_tanggal_indo(datetime.now())
    tgl_surat_tugas_input = input(f"Tanggal Surat Tugas dibuat (Default Hari Ini: {tgl_hari_ini}): ").strip()
    tgl_surat_tugas = tgl_surat_tugas_input if tgl_surat_tugas_input else tgl_hari_ini

    hari_1 = input("Hari UJK - Hari 1 (contoh: Selasa): ").strip()
    tgl_1 = input("Tanggal UJK - Hari 1 (contoh: 15 September 2026): ").strip()
    
    # --- 3. INPUT DAFTAR ASESI ---
    print("\n--- [3/3] INPUT DAFTAR ASESI ---")
    print("Masukkan nama asesi satu per satu. Tekan ENTER tanpa mengetik nama untuk SELESAI.")
    
    list_asesi = []
    i = 1
    while True:
        nama = input(f"Nama Asesi ke-{i}: ").strip()
        if not nama:
            break
        list_asesi.append(nama)
        i += 1

    if not list_asesi:
        print("\n[!] Error: Daftar asesi tidak boleh kosong.")
        return

    # Pembagian Asesi Seimbang
    asesi_hari1, asesi_hari2 = split_asesi_balanced(list_asesi)
    
    is_two_days = len(asesi_hari2) > 0
    hari_2, tgl_2 = "", ""
    if is_two_days:
        print(f"\n[Info] Total asesi ({len(list_asesi)} orang) > 8, dibagi 2 hari secara seimbang:")
        print(f"  - Hari 1: {len(asesi_hari1)} orang")
        print(f"  - Hari 2: {len(asesi_hari2)} orang")
        hari_2 = input("Hari UJK - Hari 2 (contoh: Rabu): ").strip()
        tgl_2 = input("Tanggal UJK - Hari 2 (contoh: 16 September 2026): ").strip()
    else:
        print(f"\n[Info] Total asesi ({len(list_asesi)} orang) <= 8, hanya diproses untuk 1 HARI UJK.")

    # Dictionary Data Umum
    common_data = {
        "{{ NAMA_SKEMA }}": nama_skema,
        "{{ NAMA_KEJURUAN }}": nama_kejuruan,
        "{{ NAMA_TUK }}": nama_tuk,
        "{{ ALAMAT_TUK }}": alamat_tuk,
        "{{ KOTA_TUK }}": kota_tuk,
        "{{ ZONA_WAKTU }}": zona_waktu,
        "{{ NAMA_ASESOR }}": nama_asesor,
        "{{ NO_REG_ASESOR }}": no_reg_asesor,
        "{{ ALAMAT_ASESOR }}": alamat_asesor,
        "{{ NOHP_ASESOR }}": nohp_asesor,
        "{{ TANGGAL_SURAT_TUGAS }}": tgl_surat_tugas,
    }

    # Data Hari 1 dan Hari 2 dengan daftar asesi
    data_h1 = build_asesi_data(common_data, hari_1, tgl_1, asesi_hari1)
    data_h2 = build_asesi_data(common_data, hari_2 if is_two_days else hari_1, 
                               tgl_2 if is_two_days else tgl_1, asesi_hari2 if is_two_days else [])

    # Data untuk LIST DOKUMEN UJI (dengan uppercase untuk template khusus)
    common_data_upper = {
        k: (v.upper() if isinstance(v, str) else v) 
        for k, v in common_data.items()
    }
    
    data_combined_upper = common_data_upper.copy()
    data_combined_upper["{{ HARI_UJK_1 }}"] = hari_1
    data_combined_upper["{{ TANGGAL_UJK_1 }}"] = tgl_1
    data_combined_upper["{{ HARI_UJK_2 }}"] = hari_2 if is_two_days else hari_1
    data_combined_upper["{{ TANGGAL_UJK_2 }}"] = tgl_2 if is_two_days else tgl_1

    # --- 4. PEMROSESAN FILE TEMPLATE MASTER ---
    folder_template = os.path.join(BASE_DIR, "templates")
    folder_output = os.path.join(BASE_DIR, "output_ujk")

    if not os.path.exists(folder_output):
        os.makedirs(folder_output)
    
    # Function untuk menghapus file jika sudah ada (mengatasi file yang terkunci)
    def save_document(doc, path_out):
        """
        Simpan dokumen dengan menangani file yang sudah ada/terkunci.
        """
        try:
            # Jika file sudah ada, coba hapus terlebih dahulu
            if os.path.exists(path_out):
                try:
                    os.remove(path_out)
                except PermissionError:
                    print(f"  [!] File sedang terbuka, pastikan ditutup: {os.path.basename(path_out)}")
                    return False
            
            # Simpan dokumen baru
            doc.save(path_out)
            return True
        except PermissionError as e:
            print(f"  [!] Gagal menyimpan (Permission Denied): {os.path.basename(path_out)}")
            print(f"      Kemungkinan file sedang terbuka di program lain. Tutup file terlebih dahulu.")
            return False
        except Exception as e:
            print(f"  [!] Error saat menyimpan: {str(e)}")
            return False

    def process_template_file(tmpl_filename, data_dict, is_paired=False, asesor_name=""):
        """
        Memproses satu template file dan menyimpan hasilnya.
        
        Args:
            tmpl_filename: Nama file template
            data_dict: Dictionary data untuk replacement (atau tuple untuk paired)
            is_paired: True jika dokumen berpasangan (hari 1 & 2)
            asesor_name: Nama asesor untuk penamaan file
        """
        path_in = os.path.join(folder_template, tmpl_filename)
        
        if not os.path.exists(path_in):
            print(f"  [!] Template tidak ditemukan: {path_in}")
            return

        only_filename = os.path.basename(tmpl_filename)
        base_name, ext = os.path.splitext(only_filename)
        
        if is_paired:
            data_h1, data_h2, create_day2 = data_dict
            # Dokumen Hari 1
            doc1 = Document(path_in)
            replace_placeholders_in_doc(doc1, data_h1)
            output_filename_1 = f"{base_name} - {asesor_name} 1{ext}"
            path_out_1 = os.path.join(folder_output, output_filename_1)
            if save_document(doc1, path_out_1):
                print(f"  [+] Berhasil dibuat: {output_filename_1}")
            
            # Dokumen Hari 2 (jika ada)
            if create_day2:
                doc2 = Document(path_in)
                replace_placeholders_in_doc(doc2, data_h2)
                output_filename_2 = f"{base_name} - {asesor_name} 2{ext}"
                path_out_2 = os.path.join(folder_output, output_filename_2)
                if save_document(doc2, path_out_2):
                    print(f"  [+] Berhasil dibuat: {output_filename_2}")
        else:
            doc = Document(path_in)
            replace_placeholders_in_doc(doc, data_dict)
            output_filename = f"{base_name}{ext}"
            path_out = os.path.join(folder_output, output_filename)
            if save_document(doc, path_out):
                print(f"  [+] Berhasil dibuat: {output_filename}")


    # Template Master Berpasangan
    templates_master_paired_umum = [
        "01. PERNYATAAN KERAHASIAAN.docx",
        "03. SURAT PENYATAAN KOMITMEN ASESOR.docx",
        "04. SURAT TUGAS ASESOR KOMPETENSI.docx",
        "05. DAFTAR HADIR ASESOR.docx",
        "06. DAFTAR HADIR ASESI.docx",
        "07. BERITA ACARA ASESMEN.docx",
    ]

    print("\nMemproses pembuatan dokumen...")

    # --- Proses Template Umum Berpasangan ---
    for tmpl in templates_master_paired_umum:
        path_template = os.path.join("umum", tmpl)
        process_template_file(path_template, (data_h1, data_h2, is_two_days), 
                              is_paired=True, asesor_name=clean_nama_asesor)

    # Proses semua dokumen .docx di folder skema yang dipilih.
    # Rekapitulasi dibuat per hari; dokumen skema lainnya dibuat sekali.
    path_folder_skema = os.path.join(folder_template, "skema", kode_skema)
    if os.path.isdir(path_folder_skema):
        templates_skema = sorted(
            file_name for file_name in os.listdir(path_folder_skema)
            if file_name.lower().endswith(".docx")
        )

        if not templates_skema:
            print(f"[!] Warning: Tidak ada file .docx di folder skema '{kode_skema}'.")

        for file_name in templates_skema:
            relative_template = os.path.join("skema", kode_skema, file_name)
            is_rekap = "rekap" in file_name.lower()
            process_template_file(
                relative_template,
                (data_h1, data_h2, is_two_days) if is_rekap else data_h1,
                is_paired=is_rekap,
                asesor_name=f"{clean_nama_asesor}" if is_rekap else "",
            )
    else:
        print(f"[!] Warning: Folder skema '{kode_skema}' tidak ditemukan.")

    # --- Proses List Dokumen Uji (Multi-Halaman Umum) ---
    process_template_file(os.path.join("umum", "LIST DOKUMEN UJI.docx"), data_combined_upper, is_paired=False)

    print("\n==================================================")
    print(f" SELESAI! Seluruh file telah disimpan di folder '{folder_output}'")
    print("==================================================")

if __name__ == "__main__":
    main()