import re
from io import BytesIO

import pandas as pd
import streamlit as st


# =========================================================
# KONFIGURASI HALAMAN
# =========================================================
st.set_page_config(
    page_title="Laporan APBN",
    layout="wide",
)


# =========================================================
# MEMBACA DATA EXCEL DAN MENENTUKAN TAHUN OTOMATIS
# =========================================================
@st.cache_data
def baca_data(file_bytes: bytes):
    # Membaca judul laporan pada sel A1
    judul_laporan = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name="Sheet1",
        header=None,
        nrows=1,
        usecols="A",
    ).iloc[0, 0]

    judul_laporan = str(judul_laporan).strip()

    # Mengambil tanggal awal dan akhir dari judul laporan
    tanggal_ditemukan = re.findall(
        r"\b\d{1,2}/\d{1,2}/\d{4}\b",
        judul_laporan,
    )

    tanggal_awal = None
    tanggal_akhir = None

    if len(tanggal_ditemukan) >= 2:
        tanggal_awal = pd.to_datetime(
            tanggal_ditemukan[0],
            dayfirst=True,
            errors="coerce",
        )

        tanggal_akhir = pd.to_datetime(
            tanggal_ditemukan[1],
            dayfirst=True,
            errors="coerce",
        )

    # Membaca tabel utama
    data = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name="Sheet1",
        header=1,
        usecols="A:I",
    )

    if data.shape[1] < 9:
        raise ValueError(
            "Struktur file tidak sesuai. Tabel harus memiliki minimal 9 kolom "
            "dari kolom A sampai I."
        )

    # Mencari tahun yang tertulis pada header Excel sebagai cadangan
    tahun_dari_header = sorted(
        {
            int(tahun)
            for nama_kolom in data.columns
            for tahun in re.findall(
                r"\b20\d{2}\b",
                str(nama_kolom),
            )
        }
    )

    # Tahun berjalan diprioritaskan dari tanggal akhir laporan
    if tanggal_akhir is not None and not pd.isna(tanggal_akhir):
        tahun_berjalan = int(tanggal_akhir.year)
    elif tahun_dari_header:
        tahun_berjalan = max(tahun_dari_header)
    else:
        raise ValueError(
            "Tahun laporan tidak dapat ditemukan dari judul maupun header Excel."
        )

    tahun_sebelumnya = tahun_berjalan - 1

    # Nama internal dibuat generik agar tidak perlu diubah tiap tahun
    data = data.iloc[:, :9].copy()

    data.columns = [
        "SEQ",
        "URAIAN",
        "DIPA_SEBELUMNYA",
        "REALISASI_SEBELUMNYA",
        "PERSEN_SEBELUMNYA",
        "DIPA_BERJALAN",
        "REALISASI_BERJALAN",
        "PERSEN_BERJALAN",
        "INDIKATOR",
    ]

    # Menghapus baris tanpa uraian
    data = data.dropna(
        subset=["URAIAN"]
    ).reset_index(drop=True)

    # Membersihkan kolom teks
    data["URAIAN"] = (
        data["URAIAN"]
        .astype(str)
        .str.strip()
    )

    # Mengubah kolom angka menjadi numerik
    kolom_angka = [
        "SEQ",
        "DIPA_SEBELUMNYA",
        "REALISASI_SEBELUMNYA",
        "PERSEN_SEBELUMNYA",
        "DIPA_BERJALAN",
        "REALISASI_BERJALAN",
        "PERSEN_BERJALAN",
    ]

    for kolom in kolom_angka:
        data[kolom] = pd.to_numeric(
            data[kolom],
            errors="coerce",
        ).fillna(0)

    return (
        data,
        tanggal_awal,
        tanggal_akhir,
        judul_laporan,
        tahun_sebelumnya,
        tahun_berjalan,
    )


# =========================================================
# FORMAT TANGGAL INDONESIA
# =========================================================
def format_tanggal_indonesia(tanggal) -> str:
    if tanggal is None or pd.isna(tanggal):
        return "-"

    nama_bulan = {
        1: "Januari",
        2: "Februari",
        3: "Maret",
        4: "April",
        5: "Mei",
        6: "Juni",
        7: "Juli",
        8: "Agustus",
        9: "September",
        10: "Oktober",
        11: "November",
        12: "Desember",
    }

    return (
        f"{tanggal.day} "
        f"{nama_bulan[tanggal.month]} "
        f"{tanggal.year}"
    )


# =========================================================
# FORMAT ANGKA INDONESIA
# =========================================================
def format_angka_indonesia(
    nilai: float,
    jumlah_desimal: int = 2,
) -> str:
    if pd.isna(nilai):
        return "-"

    hasil = f"{nilai:,.{jumlah_desimal}f}"

    return (
        hasil
        .replace(",", "_")
        .replace(".", ",")
        .replace("_", ".")
    )


def format_persen_indonesia(nilai) -> str:
    if pd.isna(nilai):
        return "-"

    return (
        f"{nilai:.2f}%"
        .replace(".", ",")
    )


# =========================================================
# MEMOTONG DATA BERDASARKAN URAIAN
# =========================================================
def ambil_bagian(
    data: pd.DataFrame,
    uraian_awal: str,
    uraian_akhir: str,
) -> pd.DataFrame:
    posisi_awal = data.index[
        data["URAIAN"].eq(uraian_awal)
    ].tolist()

    posisi_akhir = data.index[
        data["URAIAN"].eq(uraian_akhir)
    ].tolist()

    if not posisi_awal:
        raise ValueError(
            f'"{uraian_awal}" tidak ditemukan dalam file.'
        )

    if not posisi_akhir:
        raise ValueError(
            f'"{uraian_akhir}" tidak ditemukan dalam file.'
        )

    indeks_awal = posisi_awal[0]
    indeks_akhir = posisi_akhir[0]

    return (
        data.iloc[indeks_awal:indeks_akhir]
        .copy()
        .reset_index(drop=True)
    )


# =========================================================
# MENYIAPKAN TABEL NUMERIK
# =========================================================
def siapkan_tabel(
    data: pd.DataFrame,
    tampilkan_persentase: bool = False,
) -> pd.DataFrame:
    tabel = data[
        [
            "URAIAN",
            "DIPA_SEBELUMNYA",
            "REALISASI_SEBELUMNYA",
            "DIPA_BERJALAN",
            "REALISASI_BERJALAN",
        ]
    ].copy()

    # Persentase realisasi terhadap pagu masing-masing tahun.
    # Kolom ini hanya ditampilkan pada tabel Belanja Negara.
    if tampilkan_persentase:
        tabel["PERSEN_REALISASI_SEBELUMNYA"] = (
            tabel["REALISASI_SEBELUMNYA"]
            / tabel["DIPA_SEBELUMNYA"]
            * 100
        )

        tabel["PERSEN_REALISASI_SEBELUMNYA"] = (
            tabel["PERSEN_REALISASI_SEBELUMNYA"].where(
                tabel["DIPA_SEBELUMNYA"] != 0
            )
        )

        tabel["PERSEN_REALISASI_BERJALAN"] = (
            tabel["REALISASI_BERJALAN"]
            / tabel["DIPA_BERJALAN"]
            * 100
        )

        tabel["PERSEN_REALISASI_BERJALAN"] = (
            tabel["PERSEN_REALISASI_BERJALAN"].where(
                tabel["DIPA_BERJALAN"] != 0
            )
        )

    # YoY realisasi tahun berjalan terhadap tahun sebelumnya
    tabel["YoY"] = (
        (
            tabel["REALISASI_BERJALAN"]
            - tabel["REALISASI_SEBELUMNYA"]
        )
        / tabel["REALISASI_SEBELUMNYA"]
        * 100
    )

    # Menghindari pembagian dengan nol
    tabel["YoY"] = tabel["YoY"].where(
        tabel["REALISASI_SEBELUMNYA"] != 0
    )

    # Mengubah rupiah menjadi miliar
    kolom_rupiah = [
        "DIPA_SEBELUMNYA",
        "REALISASI_SEBELUMNYA",
        "DIPA_BERJALAN",
        "REALISASI_BERJALAN",
    ]

    tabel[kolom_rupiah] = (
        tabel[kolom_rupiah]
        / 1_000_000_000
    )

    # Menentukan urutan kolom agar persentase muncul tepat
    # setelah kolom realisasi masing-masing tahun.
    if tampilkan_persentase:
        tabel = tabel[
            [
                "URAIAN",
                "DIPA_SEBELUMNYA",
                "REALISASI_SEBELUMNYA",
                "PERSEN_REALISASI_SEBELUMNYA",
                "DIPA_BERJALAN",
                "REALISASI_BERJALAN",
                "PERSEN_REALISASI_BERJALAN",
                "YoY",
            ]
        ]
    else:
        tabel = tabel[
            [
                "URAIAN",
                "DIPA_SEBELUMNYA",
                "REALISASI_SEBELUMNYA",
                "DIPA_BERJALAN",
                "REALISASI_BERJALAN",
                "YoY",
            ]
        ]

    return tabel


# =========================================================
# MEMBUAT NAMA KOLOM DINAMIS
# =========================================================
def nama_kolom_tampilan(
    tahun_sebelumnya: int,
    tahun_berjalan: int,
    tampilkan_persentase: bool = False,
) -> dict:
    nama_kolom = {
        "DIPA_SEBELUMNYA": (
            f"DIPA {tahun_sebelumnya} (Rp Miliar)"
        ),
        "REALISASI_SEBELUMNYA": (
            f"REALISASI {tahun_sebelumnya} (Rp Miliar)"
        ),
        "DIPA_BERJALAN": (
            f"DIPA {tahun_berjalan} (Rp Miliar)"
        ),
        "REALISASI_BERJALAN": (
            f"REALISASI {tahun_berjalan} (Rp Miliar)"
        ),
    }

    if tampilkan_persentase:
        nama_kolom.update(
            {
                "PERSEN_REALISASI_SEBELUMNYA": (
                    f"% REALISASI {tahun_sebelumnya}"
                ),
                "PERSEN_REALISASI_BERJALAN": (
                    f"% REALISASI {tahun_berjalan}"
                ),
            }
        )

    return nama_kolom


def buat_tabel_tampilan(
    data: pd.DataFrame,
    tahun_sebelumnya: int,
    tahun_berjalan: int,
    tampilkan_persentase: bool = False,
) -> pd.DataFrame:
    tabel = data.copy()

    nama_kolom = nama_kolom_tampilan(
        tahun_sebelumnya,
        tahun_berjalan,
        tampilkan_persentase,
    )

    tabel = tabel.rename(
        columns=nama_kolom
    )

    kolom_miliar = [
        nama_kolom["DIPA_SEBELUMNYA"],
        nama_kolom["REALISASI_SEBELUMNYA"],
        nama_kolom["DIPA_BERJALAN"],
        nama_kolom["REALISASI_BERJALAN"],
    ]

    for kolom in kolom_miliar:
        tabel[kolom] = tabel[kolom].apply(
            format_angka_indonesia
        )

    if tampilkan_persentase:
        kolom_persen = [
            nama_kolom["PERSEN_REALISASI_SEBELUMNYA"],
            nama_kolom["PERSEN_REALISASI_BERJALAN"],
        ]

        for kolom in kolom_persen:
            tabel[kolom] = tabel[kolom].apply(
                format_persen_indonesia
            )

    tabel["YoY"] = tabel["YoY"].apply(
        format_persen_indonesia
    )

    return tabel


def buat_konfigurasi_kolom(
    tahun_sebelumnya: int,
    tahun_berjalan: int,
    tampilkan_persentase: bool = False,
) -> dict:
    nama_kolom = nama_kolom_tampilan(
        tahun_sebelumnya,
        tahun_berjalan,
        tampilkan_persentase,
    )

    konfigurasi = {
        "URAIAN": st.column_config.TextColumn(
            label="URAIAN",
            width=350,
            pinned=True,
        ),
        "YoY": st.column_config.TextColumn(
            label=(
                f"YoY {tahun_berjalan} thd "
                f"{tahun_sebelumnya}"
            ),
            width="small",
            alignment="right",
        ),
    }

    kolom_miliar_internal = [
        "DIPA_SEBELUMNYA",
        "REALISASI_SEBELUMNYA",
        "DIPA_BERJALAN",
        "REALISASI_BERJALAN",
    ]

    for nama_internal in kolom_miliar_internal:
        nama_tampilan = nama_kolom[nama_internal]

        konfigurasi[nama_tampilan] = (
            st.column_config.TextColumn(
                label=nama_tampilan,
                width="medium",
                alignment="right",
            )
        )

    if tampilkan_persentase:
        kolom_persen_internal = [
            "PERSEN_REALISASI_SEBELUMNYA",
            "PERSEN_REALISASI_BERJALAN",
        ]

        for nama_internal in kolom_persen_internal:
            nama_tampilan = nama_kolom[nama_internal]

            konfigurasi[nama_tampilan] = (
                st.column_config.TextColumn(
                    label=nama_tampilan,
                    width="small",
                    alignment="right",
                )
            )

    return konfigurasi


# =========================================================
# FUNGSI RINGKASAN NARATIF
# =========================================================
def ambil_baris(
    data: pd.DataFrame,
    uraian: str,
) -> pd.Series:
    hasil = data.loc[
        data["URAIAN"].eq(uraian)
    ]

    if hasil.empty:
        raise ValueError(
            f'Uraian "{uraian}" tidak ditemukan.'
        )

    return hasil.iloc[0]


def hitung_yoy(
    realisasi_sekarang: float,
    realisasi_sebelumnya: float,
):
    if (
        pd.isna(realisasi_sebelumnya)
        or realisasi_sebelumnya == 0
    ):
        return None

    return (
        (
            realisasi_sekarang
            - realisasi_sebelumnya
        )
        / realisasi_sebelumnya
        * 100
    )


def format_realisasi_miliar(
    baris: pd.Series,
) -> str:
    nilai_miliar = (
        baris["REALISASI_BERJALAN"]
        / 1_000_000_000
    )

    return format_angka_indonesia(
        nilai_miliar,
        jumlah_desimal=2,
    )


def dapatkan_yoy_baris(
    baris: pd.Series,
):
    return hitung_yoy(
        realisasi_sekarang=baris["REALISASI_BERJALAN"],
        realisasi_sebelumnya=baris["REALISASI_SEBELUMNYA"],
    )


def frasa_perubahan(yoy) -> str:
    if yoy is None:
        return "tidak dapat dibandingkan"

    persen = format_angka_indonesia(
        abs(yoy),
        jumlah_desimal=2,
    )

    if yoy > 0:
        return f"naik {persen}%"

    if yoy < 0:
        return f"turun {persen}%"

    return "tidak berubah (0,00%)"


def tampilkan_ringkasan_pendapatan(
    data: pd.DataFrame,
):
    pendapatan_negara = ambil_baris(
        data,
        "A. Pendapatan dan Hibah",
    )

    penerimaan_perpajakan = ambil_baris(
        data,
        "1. Penerimaan Perpajakan",
    )

    pnbp = ambil_baris(
        data,
        "2. Penerimaan Negara Bukan Pajak",
    )

    st.markdown("#### Ringkasan Pendapatan Negara")

    st.markdown(
        f"""
- Realisasi Pendapatan Negara sebesar **Rp{format_realisasi_miliar(pendapatan_negara)} miliar**, **{frasa_perubahan(dapatkan_yoy_baris(pendapatan_negara))}** dibandingkan realisasi tahun sebelumnya.
- Realisasi Penerimaan Perpajakan sebesar **Rp{format_realisasi_miliar(penerimaan_perpajakan)} miliar**, **{frasa_perubahan(dapatkan_yoy_baris(penerimaan_perpajakan))}** dibandingkan realisasi tahun sebelumnya.
- Realisasi Penerimaan Negara Bukan Pajak (PNBP) sebesar **Rp{format_realisasi_miliar(pnbp)} miliar**, **{frasa_perubahan(dapatkan_yoy_baris(pnbp))}** dibandingkan realisasi tahun sebelumnya.
        """
    )


def hitung_persen_pagu(
    baris: pd.Series,
):
    """
    Menghitung persentase realisasi terhadap pagu tahun berjalan.
    """

    pagu = baris["DIPA_BERJALAN"]

    if pd.isna(pagu) or pagu == 0:
        return None

    return (
        baris["REALISASI_BERJALAN"]
        / pagu
        * 100
    )


def frasa_persen_pagu(
    baris: pd.Series,
) -> str:
    """
    Menghasilkan keterangan seperti:
    45,25% dari Pagu
    """

    persen_pagu = hitung_persen_pagu(
        baris
    )

    if persen_pagu is None:
        return "persentase terhadap Pagu tidak dapat dihitung"

    persen = format_angka_indonesia(
        persen_pagu,
        jumlah_desimal=2,
    )

    return f"{persen}% dari Pagu"


def hapus_nomor_uraian(
    uraian: str,
) -> str:
    """
    Menghapus nomor di awal uraian.

    Contoh:
    1. Belanja Pegawai -> Belanja Pegawai
    """

    return re.sub(
        r"^\d+\.\s*",
        "",
        str(uraian),
    ).strip()


def ambil_komponen_bernomor(
    data: pd.DataFrame,
    uraian_awal: str,
    uraian_akhir=None,
) -> pd.DataFrame:
    """
    Mengambil seluruh komponen langsung bernomor 1., 2., 3., dan
    seterusnya di dalam suatu kelompok.

    Hanya komponen dengan pagu tahun berjalan yang tidak sama
    dengan nol yang diambil.
    """

    posisi_awal = data.index[
        data["URAIAN"].eq(uraian_awal)
    ].tolist()

    if not posisi_awal:
        raise ValueError(
            f'Uraian "{uraian_awal}" tidak ditemukan.'
        )

    indeks_awal = posisi_awal[0] + 1

    if uraian_akhir is not None:
        posisi_akhir = data.index[
            data["URAIAN"].eq(uraian_akhir)
        ].tolist()

        if not posisi_akhir:
            raise ValueError(
                f'Uraian "{uraian_akhir}" tidak ditemukan.'
            )

        indeks_akhir = posisi_akhir[0]
    else:
        indeks_akhir = len(data)

    bagian = data.iloc[
        indeks_awal:indeks_akhir
    ].copy()

    # Hanya mengambil baris langsung bernomor 1., 2., 3., dst.
    bagian = bagian.loc[
        bagian["URAIAN"].str.match(
            r"^\d+\.\s+",
            na=False,
        )
    ]

    # Hanya komponen yang memiliki pagu tahun berjalan
    bagian = bagian.loc[
        bagian["DIPA_BERJALAN"] != 0
    ]

    return bagian.reset_index(drop=True)


def buat_rincian_komponen(
    komponen: pd.DataFrame,
) -> str:
    """
    Menyusun narasi seluruh komponen beserta realisasi dan
    persentase terhadap pagu.
    """

    if komponen.empty:
        return (
            "Tidak terdapat komponen dengan pagu "
            "tahun berjalan yang tidak sama dengan nol."
        )

    kalimat = []

    for _, baris in komponen.iterrows():
        nama_komponen = hapus_nomor_uraian(
            baris["URAIAN"]
        )

        kalimat.append(
            f"Realisasi {nama_komponen} sebesar "
            f"Rp{format_realisasi_miliar(baris)} miliar "
            f"({frasa_persen_pagu(baris)})."
        )

    return " ".join(kalimat)


def tampilkan_ringkasan_belanja(
    data: pd.DataFrame,
):
    belanja_negara = ambil_baris(
        data,
        "B. Belanja Negara",
    )

    belanja_pemerintah_pusat = ambil_baris(
        data,
        "I. Belanja Pemerintah Pusat",
    )

    transfer_ke_daerah = ambil_baris(
        data,
        "II. Transfer Ke Daerah",
    )

    # Mengambil seluruh jenis BPP bernomor 1., 2., 3., dst.
    komponen_bpp = ambil_komponen_bernomor(
        data=data,
        uraian_awal="I. Belanja Pemerintah Pusat",
        uraian_akhir="II. Transfer Ke Daerah",
    )

    # Mengambil seluruh jenis TKD bernomor 1., 2., 3., dst.
    komponen_tkd = ambil_komponen_bernomor(
        data=data,
        uraian_awal="II. Transfer Ke Daerah",
    )

    rincian_bpp = buat_rincian_komponen(
        komponen_bpp
    )

    rincian_tkd = buat_rincian_komponen(
        komponen_tkd
    )

    st.markdown("#### Ringkasan Belanja Negara")

    st.markdown(
        f"""
- Realisasi Belanja Negara sebesar **Rp{format_realisasi_miliar(belanja_negara)} miliar ({frasa_persen_pagu(belanja_negara)})**, **{frasa_perubahan(dapatkan_yoy_baris(belanja_negara))}** dibandingkan realisasi tahun sebelumnya.
- Realisasi Belanja Pemerintah Pusat sebesar **Rp{format_realisasi_miliar(belanja_pemerintah_pusat)} miliar ({frasa_persen_pagu(belanja_pemerintah_pusat)})**, **{frasa_perubahan(dapatkan_yoy_baris(belanja_pemerintah_pusat))}** dibandingkan realisasi tahun sebelumnya. {rincian_bpp}
- Realisasi Transfer ke Daerah sebesar **Rp{format_realisasi_miliar(transfer_ke_daerah)} miliar ({frasa_persen_pagu(transfer_ke_daerah)})**, **{frasa_perubahan(dapatkan_yoy_baris(transfer_ke_daerah))}** dibandingkan realisasi tahun sebelumnya. {rincian_tkd}
        """
    )


# =========================================================
# BAGIAN 2: BELANJA PER BAGIAN ANGGARAN
# =========================================================
@st.cache_data
def baca_data_bagian_anggaran(
    file_bytes: bytes,
) -> pd.DataFrame:
    """
    Membaca file 'Realisasi Belanja Per Bagian Anggaran'.

    Struktur file:
    - Sheet: Laporan
    - Data mulai pada baris keempat
    - Kolom B  : Bagian Anggaran
    - Kolom AM : Total Pagu
    - Kolom AN : Total Realisasi
    """

    data = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name="Laporan",
        header=None,
        skiprows=3,
        usecols=[1, 38, 39],
    )

    data.columns = [
        "BAGIAN_ANGGARAN",
        "PAGU",
        "REALISASI",
    ]

    # Hanya mengambil baris yang berbentuk:
    # 022 | KEMENTERIAN PERHUBUNGAN
    data = data.dropna(
        subset=["BAGIAN_ANGGARAN"]
    ).copy()

    data["BAGIAN_ANGGARAN"] = (
        data["BAGIAN_ANGGARAN"]
        .astype(str)
        .str.strip()
    )

    data["KODE_BA"] = (
        data["BAGIAN_ANGGARAN"]
        .str.extract(
            r"^\s*(\d{3})\s*\|",
            expand=False,
        )
    )

    data = data.dropna(
        subset=["KODE_BA"]
    ).copy()

    for kolom in ["PAGU", "REALISASI"]:
        data[kolom] = pd.to_numeric(
            data[kolom],
            errors="coerce",
        ).fillna(0)

    # Antisipasi apabila satu kode BA muncul lebih dari sekali.
    # Nama BA tetap memakai nama yang terdapat pada file tersebut.
    data = (
        data.groupby(
            "KODE_BA",
            as_index=False,
        )
        .agg(
            BAGIAN_ANGGARAN=(
                "BAGIAN_ANGGARAN",
                "first",
            ),
            PAGU=(
                "PAGU",
                "sum",
            ),
            REALISASI=(
                "REALISASI",
                "sum",
            ),
        )
    )

    return data


def ambil_tahun_dari_nama_file(
    nama_file: str,
):
    """
    Mengambil tahun dari nama file, misalnya:
    Realisasi_Belanja_..._30 Juni 2026.xlsx
    """

    tahun = re.findall(
        r"\b(20\d{2})\b",
        str(nama_file),
    )

    if not tahun:
        return None

    return int(tahun[-1])


def ambil_tanggal_dari_nama_file(
    nama_file: str,
) -> str:
    """
    Mengambil label tanggal dari nama file, misalnya:
    30 Juni 2026 -> 30 Juni
    """

    nama_bulan = (
        "Januari|Februari|Maret|April|Mei|Juni|"
        "Juli|Agustus|September|Oktober|November|Desember"
    )

    hasil = re.search(
        rf"\b(\d{{1,2}})\s+({nama_bulan})\s+20\d{{2}}\b",
        str(nama_file),
        flags=re.IGNORECASE,
    )

    if hasil is None:
        return "tanggal laporan"

    tanggal = hasil.group(1)
    bulan = hasil.group(2).title()

    return f"{tanggal} {bulan}"


def hitung_growth_ba(
    realisasi_berjalan: float,
    realisasi_sebelumnya: float,
):
    """
    Growth realisasi tahun berjalan terhadap tahun sebelumnya.
    """

    if (
        pd.isna(realisasi_sebelumnya)
        or realisasi_sebelumnya == 0
    ):
        return None

    return (
        (
            realisasi_berjalan
            - realisasi_sebelumnya
        )
        / realisasi_sebelumnya
        * 100
    )


def gabungkan_bagian_anggaran(
    data_berjalan: pd.DataFrame,
    data_sebelumnya: pd.DataFrame,
) -> pd.DataFrame:
    """
    Baseline memakai kode dan nama Bagian Anggaran tahun berjalan.

    Kode yang hanya terdapat pada tahun sebelumnya tidak ikut
    ditampilkan.
    """

    pembanding = data_sebelumnya[
        [
            "KODE_BA",
            "REALISASI",
        ]
    ].rename(
        columns={
            "REALISASI": "REALISASI_SEBELUMNYA",
        }
    )

    hasil = data_berjalan.merge(
        pembanding,
        on="KODE_BA",
        how="left",
    )

    hasil["REALISASI_SEBELUMNYA"] = (
        hasil["REALISASI_SEBELUMNYA"]
        .fillna(0)
    )

    hasil = hasil.rename(
        columns={
            "PAGU": "PAGU_BERJALAN",
            "REALISASI": "REALISASI_BERJALAN",
        }
    )

    hasil["PERSEN_TERHADAP_PAGU"] = (
        hasil["REALISASI_BERJALAN"]
        / hasil["PAGU_BERJALAN"]
        * 100
    ).where(
        hasil["PAGU_BERJALAN"] != 0
    )

    hasil["GROWTH"] = (
        (
            hasil["REALISASI_BERJALAN"]
            - hasil["REALISASI_SEBELUMNYA"]
        )
        / hasil["REALISASI_SEBELUMNYA"]
        * 100
    ).where(
        hasil["REALISASI_SEBELUMNYA"] != 0
    )

    total_realisasi = (
        hasil["REALISASI_BERJALAN"].sum()
    )

    if total_realisasi != 0:
        hasil["PERSEN_TOTAL_REALISASI"] = (
            hasil["REALISASI_BERJALAN"]
            / total_realisasi
            * 100
        )
    else:
        hasil["PERSEN_TOTAL_REALISASI"] = pd.NA

    return hasil


def buat_baris_ringkasan_ba(
    data: pd.DataFrame,
    nama_baris: str,
    total_realisasi: float,
) -> dict:
    """
    Menyusun baris agregat:
    - 10 K/L Pagu Terbesar
    - K/L Lainnya + TKD
    - Total
    """

    pagu = data["PAGU_BERJALAN"].sum()
    realisasi = data["REALISASI_BERJALAN"].sum()
    realisasi_sebelumnya = (
        data["REALISASI_SEBELUMNYA"].sum()
    )

    persen_pagu = None
    if pagu != 0:
        persen_pagu = (
            realisasi
            / pagu
            * 100
        )

    growth = hitung_growth_ba(
        realisasi,
        realisasi_sebelumnya,
    )

    persen_total = None
    if total_realisasi != 0:
        persen_total = (
            realisasi
            / total_realisasi
            * 100
        )

    return {
        "BAGIAN_ANGGARAN": nama_baris,
        "PAGU_BERJALAN": pagu,
        "REALISASI_BERJALAN": realisasi,
        "REALISASI_SEBELUMNYA": realisasi_sebelumnya,
        "PERSEN_TERHADAP_PAGU": persen_pagu,
        "GROWTH": growth,
        "PERSEN_TOTAL_REALISASI": persen_total,
    }


def siapkan_tabel_bagian_anggaran(
    data_gabungan: pd.DataFrame,
) -> pd.DataFrame:
    """
    Menampilkan:
    1. Sepuluh K/L dengan pagu terbesar.
    2. Ringkasan 10 K/L Pagu Terbesar.
    3. K/L Lainnya.
    4. Transfer ke Daerah (BA 999).
    5. Total.

    BA 999 tidak dipilih sebagai K/L Top 10 dan ditampilkan
    sebagai baris agregat Transfer ke Daerah yang terpisah.
    """

    total_realisasi = (
        data_gabungan["REALISASI_BERJALAN"].sum()
    )

    # Seluruh K/L selain BA 999
    data_kl = data_gabungan.loc[
        data_gabungan["KODE_BA"] != "999"
    ].copy()

    # BA 999 diperlakukan sebagai Transfer ke Daerah
    data_tkd = data_gabungan.loc[
        data_gabungan["KODE_BA"] == "999"
    ].copy()

    top_10 = (
        data_kl.sort_values(
            "PAGU_BERJALAN",
            ascending=False,
        )
        .head(10)
        .copy()
    )

    kode_top_10 = set(
        top_10["KODE_BA"]
    )

    # K/L lainnya hanya berisi BA selain 999
    # yang tidak masuk dalam Top 10.
    kl_lainnya = data_kl.loc[
        ~data_kl["KODE_BA"].isin(
            kode_top_10
        )
    ].copy()

    kolom_tabel = [
        "BAGIAN_ANGGARAN",
        "PAGU_BERJALAN",
        "REALISASI_BERJALAN",
        "REALISASI_SEBELUMNYA",
        "PERSEN_TERHADAP_PAGU",
        "GROWTH",
        "PERSEN_TOTAL_REALISASI",
    ]

    tabel_detail = top_10[
        kolom_tabel
    ].copy()

    baris_top_10 = buat_baris_ringkasan_ba(
        top_10,
        "10 K/L Pagu Terbesar",
        total_realisasi,
    )

    baris_lainnya = buat_baris_ringkasan_ba(
        kl_lainnya,
        "K/L Lainnya",
        total_realisasi,
    )

    baris_tkd = buat_baris_ringkasan_ba(
        data_tkd,
        "Transfer ke Daerah",
        total_realisasi,
    )

    baris_total = buat_baris_ringkasan_ba(
        data_gabungan,
        "Total",
        total_realisasi,
    )

    tabel_ringkasan = pd.DataFrame(
        [
            baris_top_10,
            baris_lainnya,
            baris_tkd,
            baris_total,
        ]
    )

    tabel = pd.concat(
        [
            tabel_detail,
            tabel_ringkasan,
        ],
        ignore_index=True,
    )

    return tabel


def format_growth_ba(nilai) -> str:
    """
    Memformat nilai growth yang sudah berhasil dihitung.
    """

    if nilai is None or pd.isna(nilai):
        return "-"

    return format_persen_indonesia(
        nilai
    )


def format_growth_ba_baris(
    baris: pd.Series,
) -> str:
    """
    Menentukan tampilan growth berdasarkan realisasi dua tahun.

    - Tahun sebelumnya 0 dan tahun berjalan > 0:
      "Baru terealisasi"
    - Kedua tahun 0:
      "-"
    - Selain itu:
      persentase growth
    """

    realisasi_berjalan = baris["REALISASI_BERJALAN"]
    realisasi_sebelumnya = baris["REALISASI_SEBELUMNYA"]

    if (
        pd.isna(realisasi_sebelumnya)
        or realisasi_sebelumnya == 0
    ):
        if (
            not pd.isna(realisasi_berjalan)
            and realisasi_berjalan > 0
        ):
            return "Baru terealisasi"

        return "-"

    return format_growth_ba(
        baris["GROWTH"]
    )


def buat_tabel_bagian_anggaran_tampilan(
    data: pd.DataFrame,
    tahun_berjalan: int,
    tanggal_realisasi: str,
) -> pd.DataFrame:
    """
    Mengubah tabel numerik menjadi tabel teks agar hasil copy
    sesuai dengan tampilan.
    """

    tabel = data.copy()

    # Growth harus ditentukan selagi nilai realisasi masih numerik.
    tabel["GROWTH"] = tabel.apply(
        format_growth_ba_baris,
        axis=1,
    )

    tabel["PAGU_BERJALAN"] = (
        tabel["PAGU_BERJALAN"]
        / 1_000_000_000
    )

    tabel["REALISASI_BERJALAN"] = (
        tabel["REALISASI_BERJALAN"]
        / 1_000_000_000
    )

    tabel["PAGU_BERJALAN"] = (
        tabel["PAGU_BERJALAN"]
        .apply(format_angka_indonesia)
    )

    tabel["REALISASI_BERJALAN"] = (
        tabel["REALISASI_BERJALAN"]
        .apply(format_angka_indonesia)
    )

    tabel["PERSEN_TERHADAP_PAGU"] = (
        tabel["PERSEN_TERHADAP_PAGU"]
        .apply(format_persen_indonesia)
    )

    tabel["PERSEN_TOTAL_REALISASI"] = (
        tabel["PERSEN_TOTAL_REALISASI"]
        .apply(format_persen_indonesia)
    )

    tabel = tabel.drop(
        columns=["REALISASI_SEBELUMNYA"],
        errors="ignore",
    )

    tabel = tabel.rename(
        columns={
            "BAGIAN_ANGGARAN": "Kementerian/Lembaga",
            "PAGU_BERJALAN": (
                f"Pagu {tahun_berjalan} (Rp Miliar)"
            ),
            "REALISASI_BERJALAN": (
                f"Real. s.d {tanggal_realisasi} "
                f"(Rp Miliar)"
            ),
            "PERSEN_TERHADAP_PAGU": (
                "% terhadap Pagu"
            ),
            "GROWTH": "Growth (%)",
            "PERSEN_TOTAL_REALISASI": (
                "% terhadap Total Realisasi"
            ),
        }
    )

    return tabel


def buat_konfigurasi_bagian_anggaran(
    tahun_berjalan: int,
    tanggal_realisasi: str,
) -> dict:
    return {
        "Kementerian/Lembaga": (
            st.column_config.TextColumn(
                label="Kementerian/Lembaga",
                width=360,
                pinned=True,
            )
        ),
        f"Pagu {tahun_berjalan} (Rp Miliar)": (
            st.column_config.TextColumn(
                label=(
                    f"Pagu {tahun_berjalan} "
                    f"(Rp Miliar)"
                ),
                width="medium",
                alignment="right",
            )
        ),
        (
            f"Real. s.d {tanggal_realisasi} "
            f"(Rp Miliar)"
        ): (
            st.column_config.TextColumn(
                label=(
                    f"Real. s.d {tanggal_realisasi} "
                    f"(Rp Miliar)"
                ),
                width="medium",
                alignment="right",
            )
        ),
        "% terhadap Pagu": (
            st.column_config.TextColumn(
                label="% terhadap Pagu",
                width="small",
                alignment="right",
            )
        ),
        "Growth (%)": (
            st.column_config.TextColumn(
                label="Growth (%)",
                width="small",
                alignment="right",
            )
        ),
        "% terhadap Total Realisasi": (
            st.column_config.TextColumn(
                label="% terhadap Total Realisasi",
                width="small",
                alignment="right",
            )
        ),
    }


# =========================================================
# RINGKASAN 10 K/L DENGAN PAGU TERBESAR
# =========================================================
def nama_ba_tanpa_kode(
    nama_ba: str,
) -> str:
    """
    Menghapus kode BA di depan nama K/L.

    Contoh:
    022 | KEMENTERIAN PERHUBUNGAN
    menjadi:
    KEMENTERIAN PERHUBUNGAN
    """

    return re.sub(
        r"^\s*\d{3}\s*\|\s*",
        "",
        str(nama_ba),
    ).strip()


def frasa_growth_ba(
    growth,
    realisasi_berjalan=None,
    realisasi_sebelumnya=None,
) -> str:
    """
    Mengubah growth menjadi frasa naratif.

    Jika tahun sebelumnya belum memiliki realisasi tetapi
    tahun berjalan sudah memiliki realisasi, growth tidak
    dinyatakan sebagai persentase.
    """

    if (
        realisasi_sebelumnya is not None
        and not pd.isna(realisasi_sebelumnya)
        and realisasi_sebelumnya == 0
    ):
        if (
            realisasi_berjalan is not None
            and not pd.isna(realisasi_berjalan)
            and realisasi_berjalan > 0
        ):
            return (
                "belum dapat dibandingkan karena pada tahun "
                "sebelumnya belum terdapat realisasi"
            )

        return "tidak dapat dibandingkan"

    if growth is None or pd.isna(growth):
        return "tidak dapat dibandingkan"

    persen = format_angka_indonesia(
        abs(growth),
        jumlah_desimal=2,
    )

    if growth > 0:
        return f"naik {persen}%"

    if growth < 0:
        return f"turun {persen}%"

    return "tidak berubah (0,00%)"


def tampilkan_ringkasan_10_kl(
    data_gabungan: pd.DataFrame,
    tahun_berjalan: int,
    tahun_sebelumnya: int,
):
    """
    Menampilkan ringkasan agregat 10 K/L dengan pagu terbesar.

    Ringkasan tidak menampilkan K/L dengan pagu terbesar,
    persentase realisasi tertinggi, maupun growth individual.
    """

    data_kl = data_gabungan.loc[
        data_gabungan["KODE_BA"] != "999"
    ].copy()

    top_10 = (
        data_kl.sort_values(
            "PAGU_BERJALAN",
            ascending=False,
        )
        .head(10)
        .copy()
    )

    if top_10.empty:
        st.warning(
            "Data 10 K/L dengan pagu terbesar tidak tersedia."
        )
        return

    total_pagu_top_10 = (
        top_10["PAGU_BERJALAN"].sum()
    )

    total_realisasi_top_10 = (
        top_10["REALISASI_BERJALAN"].sum()
    )

    total_realisasi_sebelumnya_top_10 = (
        top_10["REALISASI_SEBELUMNYA"].sum()
    )

    total_pagu_seluruh_ba = (
        data_gabungan["PAGU_BERJALAN"].sum()
    )

    total_realisasi_seluruh_ba = (
        data_gabungan["REALISASI_BERJALAN"].sum()
    )

    persen_realisasi_top_10 = None

    if total_pagu_top_10 != 0:
        persen_realisasi_top_10 = (
            total_realisasi_top_10
            / total_pagu_top_10
            * 100
        )

    share_pagu_top_10 = None

    if total_pagu_seluruh_ba != 0:
        share_pagu_top_10 = (
            total_pagu_top_10
            / total_pagu_seluruh_ba
            * 100
        )

    share_realisasi_top_10 = None

    if total_realisasi_seluruh_ba != 0:
        share_realisasi_top_10 = (
            total_realisasi_top_10
            / total_realisasi_seluruh_ba
            * 100
        )

    growth_top_10 = hitung_growth_ba(
        total_realisasi_top_10,
        total_realisasi_sebelumnya_top_10,
    )

    def format_persen_ringkasan(nilai):
        if nilai is None or pd.isna(nilai):
            return "-"

        return (
            f"{format_angka_indonesia(nilai)}%"
        )

    st.markdown("#### Ringkasan 10 K/L")

    st.markdown(
        f"""
- Total pagu 10 K/L terbesar tahun {tahun_berjalan} mencapai **Rp{format_angka_indonesia(total_pagu_top_10 / 1_000_000_000)} miliar**, atau **{format_persen_ringkasan(share_pagu_top_10)}** dari total pagu.
- Total realisasi 10 K/L terbesar sebesar **Rp{format_angka_indonesia(total_realisasi_top_10 / 1_000_000_000)} miliar**, atau **{format_persen_ringkasan(persen_realisasi_top_10)}** dari pagu 10 K/L terbesar.
- Dibandingkan tahun {tahun_sebelumnya}, realisasi 10 K/L terbesar **{frasa_growth_ba(
    growth_top_10,
    realisasi_berjalan=total_realisasi_top_10,
    realisasi_sebelumnya=total_realisasi_sebelumnya_top_10,
)}**.
- Kontribusi realisasi 10 K/L terbesar mencapai **{format_persen_ringkasan(share_realisasi_top_10)}** dari total realisasi tahun {tahun_berjalan}.
        """
    )


# =========================================================
# 5 K/L DENGAN PERSENTASE REALISASI TERTINGGI/TERENDAH
# =========================================================
def pilih_5_kl_persentase_realisasi(
    data_gabungan: pd.DataFrame,
    urutan: str,
) -> pd.DataFrame:
    """
    Memilih 5 K/L berdasarkan persentase realisasi terhadap pagu.

    - BA 999 tidak disertakan karena merupakan Transfer ke Daerah.
    - Hanya K/L dengan pagu tahun berjalan lebih dari nol.
    - urutan: "tertinggi" atau "terendah".
    """

    data_kl = data_gabungan.loc[
        (data_gabungan["KODE_BA"] != "999")
        & (data_gabungan["PAGU_BERJALAN"] > 0)
        & (data_gabungan["PERSEN_TERHADAP_PAGU"].notna())
    ].copy()

    if urutan == "tertinggi":
        ascending = False
    elif urutan == "terendah":
        ascending = True
    else:
        raise ValueError(
            'Parameter urutan harus "tertinggi" atau "terendah".'
        )

    hasil = (
        data_kl.sort_values(
            [
                "PERSEN_TERHADAP_PAGU",
                "PAGU_BERJALAN",
            ],
            ascending=[
                ascending,
                False,
            ],
        )
        .head(5)
        .reset_index(drop=True)
    )

    hasil.insert(
        0,
        "NO",
        range(1, len(hasil) + 1),
    )

    return hasil


def buat_tabel_5_kl_tampilan(
    data: pd.DataFrame,
    tahun_berjalan: int,
    tanggal_realisasi: str,
) -> pd.DataFrame:
    """
    Menyiapkan tabel teks agar hasil copy sesuai tampilan.
    """

    tabel = data[
        [
            "NO",
            "BAGIAN_ANGGARAN",
            "PAGU_BERJALAN",
            "REALISASI_BERJALAN",
            "REALISASI_SEBELUMNYA",
            "PERSEN_TERHADAP_PAGU",
            "GROWTH",
            "PERSEN_TOTAL_REALISASI",
        ]
    ].copy()

    # Growth harus ditentukan selagi nilai realisasi masih numerik.
    tabel["GROWTH"] = tabel.apply(
        format_growth_ba_baris,
        axis=1,
    )

    tabel["PAGU_BERJALAN"] = (
        tabel["PAGU_BERJALAN"]
        / 1_000_000_000
    )

    tabel["REALISASI_BERJALAN"] = (
        tabel["REALISASI_BERJALAN"]
        / 1_000_000_000
    )

    tabel["PAGU_BERJALAN"] = (
        tabel["PAGU_BERJALAN"]
        .apply(format_angka_indonesia)
    )

    tabel["REALISASI_BERJALAN"] = (
        tabel["REALISASI_BERJALAN"]
        .apply(format_angka_indonesia)
    )

    tabel["PERSEN_TERHADAP_PAGU"] = (
        tabel["PERSEN_TERHADAP_PAGU"]
        .apply(format_persen_indonesia)
    )

    tabel["PERSEN_TOTAL_REALISASI"] = (
        tabel["PERSEN_TOTAL_REALISASI"]
        .apply(format_persen_indonesia)
    )

    # Kolom ini hanya dipakai untuk menentukan status Growth.
    tabel = tabel.drop(
        columns=["REALISASI_SEBELUMNYA"],
        errors="ignore",
    )

    tabel = tabel.rename(
        columns={
            "NO": "No",
            "BAGIAN_ANGGARAN": "Kementerian/Lembaga",
            "PAGU_BERJALAN": (
                f"Pagu {tahun_berjalan} (Rp Miliar)"
            ),
            "REALISASI_BERJALAN": (
                f"Real. s.d {tanggal_realisasi} "
                f"(Rp Miliar)"
            ),
            "PERSEN_TERHADAP_PAGU": (
                "% terhadap Pagu"
            ),
            "GROWTH": "Growth (%)",
            "PERSEN_TOTAL_REALISASI": (
                "% terhadap Total Realisasi"
            ),
        }
    )

    return tabel


def buat_konfigurasi_5_kl(
    tahun_berjalan: int,
    tanggal_realisasi: str,
) -> dict:
    return {
        "No": st.column_config.NumberColumn(
            label="No",
            width="small",
            format="%d",
        ),
        "Kementerian/Lembaga": (
            st.column_config.TextColumn(
                label="Kementerian/Lembaga",
                width=360,
                pinned=True,
            )
        ),
        f"Pagu {tahun_berjalan} (Rp Miliar)": (
            st.column_config.TextColumn(
                label=(
                    f"Pagu {tahun_berjalan} "
                    f"(Rp Miliar)"
                ),
                width="medium",
                alignment="right",
            )
        ),
        (
            f"Real. s.d {tanggal_realisasi} "
            f"(Rp Miliar)"
        ): (
            st.column_config.TextColumn(
                label=(
                    f"Real. s.d {tanggal_realisasi} "
                    f"(Rp Miliar)"
                ),
                width="medium",
                alignment="right",
            )
        ),
        "% terhadap Pagu": (
            st.column_config.TextColumn(
                label="% terhadap Pagu",
                width="small",
                alignment="right",
            )
        ),
        "Growth (%)": (
            st.column_config.TextColumn(
                label="Growth (%)",
                width="small",
                alignment="right",
            )
        ),
        "% terhadap Total Realisasi": (
            st.column_config.TextColumn(
                label="% terhadap Total Realisasi",
                width="small",
                alignment="right",
            )
        ),
    }


def tampilkan_ringkasan_5_kl(
    data_5_kl: pd.DataFrame,
    data_gabungan: pd.DataFrame,
    tahun_berjalan: int,
    tahun_sebelumnya: int,
    jenis: str,
):
    """
    Menampilkan hanya satu K/L dengan persentase realisasi
    paling tinggi atau paling rendah pada masing-masing tab.
    """

    if data_5_kl.empty:
        st.warning(
            f"Data K/L dengan persentase realisasi {jenis} "
            "tidak tersedia."
        )
        return

    # Data pada tab tertinggi sudah diurutkan menurun,
    # sedangkan tab terendah sudah diurutkan menaik.
    baris = data_5_kl.iloc[0]

    nama_kl = nama_ba_tanpa_kode(
        baris["BAGIAN_ANGGARAN"]
    )

    pagu_miliar = (
        baris["PAGU_BERJALAN"]
        / 1_000_000_000
    )

    realisasi_miliar = (
        baris["REALISASI_BERJALAN"]
        / 1_000_000_000
    )

    persen_realisasi = (
        baris["PERSEN_TERHADAP_PAGU"]
    )

    growth = baris["GROWTH"]

    kontribusi = (
        baris["PERSEN_TOTAL_REALISASI"]
    )

    if jenis == "tertinggi":
        judul = "Realisasi Tertinggi"
        kalimat_utama = (
            "K/L dengan persentase realisasi tertinggi"
        )
    else:
        judul = "Realisasi Terendah"
        kalimat_utama = (
            "K/L dengan persentase realisasi terendah"
        )

    st.markdown(
        f"#### {judul}"
    )

    st.markdown(
        f"""
**{nama_kl}** merupakan {kalimat_utama}, dengan realisasi sebesar **Rp{format_angka_indonesia(realisasi_miliar)} miliar** atau **{format_persen_indonesia(persen_realisasi)}** dari pagu sebesar **Rp{format_angka_indonesia(pagu_miliar)} miliar**.

Dibandingkan tahun {tahun_sebelumnya}, realisasinya **{frasa_growth_ba(
    growth,
    realisasi_berjalan=baris["REALISASI_BERJALAN"],
    realisasi_sebelumnya=baris["REALISASI_SEBELUMNYA"],
)}** dan memberikan kontribusi sebesar **{format_persen_indonesia(kontribusi)}** terhadap total realisasi tahun {tahun_berjalan}.
        """
    )


# =========================================================
# TAMPILAN UTAMA
# =========================================================
st.title("Laporan Mingguan Kinerja APBN")

# Menyimpan tahun hasil pembacaan file I-Account agar dapat
# dipakai oleh bagian Belanja per Bagian Anggaran.
if "tahun_iaccount_sebelumnya" not in st.session_state:
    st.session_state.tahun_iaccount_sebelumnya = None

if "tahun_iaccount_berjalan" not in st.session_state:
    st.session_state.tahun_iaccount_berjalan = None


# =========================================================
# BAGIAN 1: LAPORAN I-ACCOUNT
# =========================================================
st.header("Laporan I-Account DIPA Year on Year")

file_excel = st.file_uploader(
    "Unggah file Laporan I-Account DIPA",
    type=["xlsx"],
    key="file_i_account",
    help=(
        "Gunakan file dengan struktur yang sama: judul periode pada sel A1 "
        "dan tabel utama pada baris kedua."
    ),
)

if file_excel is None:
    st.info(
        "File I-Account belum diunggah. "
        "Bagian Anggaran di bawah tetap dapat digunakan."
    )
else:
    try:
        file_bytes = file_excel.getvalue()

        (
            data,
            tanggal_awal,
            tanggal_akhir,
            judul_laporan,
            tahun_sebelumnya,
            tahun_berjalan,
        ) = baca_data(file_bytes)

        # Tahun untuk Bagian Anggaran selalu mengikuti
        # tahun yang terbaca dari laporan I-Account.
        st.session_state.tahun_iaccount_sebelumnya = (
            tahun_sebelumnya
        )
        st.session_state.tahun_iaccount_berjalan = (
            tahun_berjalan
        )

        st.subheader(
            f"Perbandingan Tahun {tahun_sebelumnya} "
            f"dan {tahun_berjalan}"
        )

        if (
            tanggal_awal is not None
            and tanggal_akhir is not None
            and not pd.isna(tanggal_awal)
            and not pd.isna(tanggal_akhir)
        ):
            st.caption(
                f"Periode "
                f"{format_tanggal_indonesia(tanggal_awal)} "
                f"sampai dengan "
                f"{format_tanggal_indonesia(tanggal_akhir)}"
            )
        else:
            st.warning(
                "Tanggal periode tidak berhasil dibaca dari judul Excel."
            )
            st.caption(judul_laporan)

        # Membagi data pendapatan dan belanja
        pendapatan_negara = ambil_bagian(
            data=data,
            uraian_awal="A. Pendapatan dan Hibah",
            uraian_akhir="B. Belanja Negara",
        )

        belanja_negara = ambil_bagian(
            data=data,
            uraian_awal="B. Belanja Negara",
            uraian_akhir="C. Surplus Defisit",
        )

        # Tabel numerik untuk perhitungan
        tabel_pendapatan_numerik = siapkan_tabel(
            pendapatan_negara,
            tampilkan_persentase=False,
        )

        tabel_belanja_numerik = siapkan_tabel(
            belanja_negara,
            tampilkan_persentase=True,
        )

        # Tabel teks untuk tampilan, copy, dan download
        tabel_pendapatan = buat_tabel_tampilan(
            tabel_pendapatan_numerik,
            tahun_sebelumnya,
            tahun_berjalan,
            tampilkan_persentase=False,
        )

        tabel_belanja = buat_tabel_tampilan(
            tabel_belanja_numerik,
            tahun_sebelumnya,
            tahun_berjalan,
            tampilkan_persentase=True,
        )

        konfigurasi_kolom_pendapatan = buat_konfigurasi_kolom(
            tahun_sebelumnya,
            tahun_berjalan,
            tampilkan_persentase=False,
        )

        konfigurasi_kolom_belanja = buat_konfigurasi_kolom(
            tahun_sebelumnya,
            tahun_berjalan,
            tampilkan_persentase=True,
        )

        tab_pendapatan, tab_belanja = st.tabs(
            [
                f"Pendapatan Negara {tahun_berjalan}",
                f"Belanja Negara {tahun_berjalan}",
            ]
        )

        # =====================================================
        # TAB PENDAPATAN NEGARA
        # =====================================================
        with tab_pendapatan:
            st.subheader("Pendapatan Negara")

            (
                col_tabel_pendapatan,
                col_ringkasan_pendapatan,
            ) = st.columns(
                [3, 1],
                gap="small",
                border=True,
            )

            with col_tabel_pendapatan:
                st.dataframe(
                    tabel_pendapatan,
                    width="stretch",
                    height="content",
                    hide_index=True,
                    column_config=konfigurasi_kolom_pendapatan,
                )

            with col_ringkasan_pendapatan:
                tampilkan_ringkasan_pendapatan(
                    pendapatan_negara
                )

        # =====================================================
        # TAB BELANJA NEGARA
        # =====================================================
        with tab_belanja:
            st.subheader("Belanja Negara")

            (
                col_tabel_belanja,
                col_ringkasan_belanja,
            ) = st.columns(
                [3, 1],
                gap="small",
                border=True,
            )

            with col_tabel_belanja:
                st.dataframe(
                    tabel_belanja,
                    width="stretch",
                    height="content",
                    hide_index=True,
                    column_config=konfigurasi_kolom_belanja,
                )

            with col_ringkasan_belanja:
                tampilkan_ringkasan_belanja(
                    belanja_negara
                )

    except Exception as error_i_account:
        st.error(
            "Gagal membaca atau mengolah file I-Account: "
            f"{error_i_account}"
        )


# =========================================================
# BAGIAN 2: BELANJA PER BAGIAN ANGGARAN
# =========================================================
st.divider()

st.header("Belanja per Bagian Anggaran")

st.caption(
    "Bagian ini berdiri sendiri. Nama dan kode Kementerian/Lembaga "
    "menggunakan baseline file tahun berjalan."
)

(
    kolom_upload_berjalan,
    kolom_upload_sebelumnya,
) = st.columns(
    2,
    gap="medium",
    border=True,
)

with kolom_upload_berjalan:
    file_ba_berjalan = st.file_uploader(
        "Unggah Bagian Anggaran Tahun Berjalan",
        type=["xlsx"],
        key="file_ba_berjalan",
        help=(
            "Tahun file mengikuti periode pada laporan I-Account. "
            "Nama file bebas dan tidak perlu diubah."
        ),
    )

with kolom_upload_sebelumnya:
    file_ba_sebelumnya = st.file_uploader(
        "Unggah Bagian Anggaran Tahun Sebelumnya",
        type=["xlsx"],
        key="file_ba_sebelumnya",
        help=(
            "Tahun file mengikuti periode pembanding pada laporan "
            "I-Account. Nama file bebas dan tidak perlu diubah."
        ),
    )

if (
    file_ba_berjalan is None
    or file_ba_sebelumnya is None
):
    st.info(
        "Unggah kedua file Bagian Anggaran untuk menampilkan "
        "perbandingan pagu dan realisasi. Tahun tabel akan mengikuti "
        "periode pada file I-Account."
    )
else:
    try:
        # Tahun Bagian Anggaran wajib mengikuti tahun yang
        # terbaca dari laporan I-Account, bukan nama file.
        tahun_ba_berjalan = (
            st.session_state.tahun_iaccount_berjalan
        )
        tahun_ba_sebelumnya = (
            st.session_state.tahun_iaccount_sebelumnya
        )

        if (
            tahun_ba_berjalan is None
            or tahun_ba_sebelumnya is None
        ):
            raise ValueError(
                "Unggah dan proses file I-Account terlebih dahulu "
                "agar tahun berjalan dan tahun sebelumnya dapat ditentukan."
            )

        data_ba_berjalan = baca_data_bagian_anggaran(
            file_ba_berjalan.getvalue()
        )

        data_ba_sebelumnya = baca_data_bagian_anggaran(
            file_ba_sebelumnya.getvalue()
        )

        data_ba_gabungan = gabungkan_bagian_anggaran(
            data_ba_berjalan,
            data_ba_sebelumnya,
        )

        tabel_ba_numerik = siapkan_tabel_bagian_anggaran(
            data_ba_gabungan
        )

        # Label tanggal tabel Bagian Anggaran mengikuti
        # tanggal akhir pada laporan I-Account.
        if (
            "tanggal_akhir" in locals()
            and tanggal_akhir is not None
            and not pd.isna(tanggal_akhir)
        ):
            tanggal_realisasi_ba = (
                f"{tanggal_akhir.day} "
                f"{format_tanggal_indonesia(tanggal_akhir).split(' ', 1)[1]}"
            )
        else:
            tanggal_realisasi_ba = "tanggal laporan"

        tabel_ba = buat_tabel_bagian_anggaran_tampilan(
            tabel_ba_numerik,
            tahun_ba_berjalan,
            tanggal_realisasi_ba,
        )

        konfigurasi_ba = buat_konfigurasi_bagian_anggaran(
            tahun_ba_berjalan,
            tanggal_realisasi_ba,
        )

        st.subheader(
            f"10 K/L dengan Pagu Terbesar Tahun "
            f"{tahun_ba_berjalan}"
        )

        st.caption(
            f"Perbandingan realisasi tahun {tahun_ba_berjalan} "
            f"terhadap tahun {tahun_ba_sebelumnya}."
        )

        (
            kolom_tabel_ba,
            kolom_ringkasan_ba,
        ) = st.columns(
            [4, 1.35],
            gap="small",
            border=True,
        )

        with kolom_tabel_ba:
            st.dataframe(
                tabel_ba,
                width="stretch",
                height="content",
                hide_index=True,
                column_config=konfigurasi_ba,
            )

        with kolom_ringkasan_ba:
            tampilkan_ringkasan_10_kl(
                data_ba_gabungan,
                tahun_ba_berjalan,
                tahun_ba_sebelumnya,
            )

        # =====================================================
        # SUBBAGIAN 5 K/L TERTINGGI DAN TERENDAH
        # =====================================================
        st.subheader(
            "Peringkat Persentase Realisasi Kementerian/Lembaga"
        )

        (
            tab_5_tertinggi,
            tab_5_terendah,
        ) = st.tabs(
            [
                "5 K/L Realisasi Tertinggi",
                "5 K/L Realisasi Terendah",
            ]
        )

        data_5_tertinggi = (
            pilih_5_kl_persentase_realisasi(
                data_ba_gabungan,
                urutan="tertinggi",
            )
        )

        data_5_terendah = (
            pilih_5_kl_persentase_realisasi(
                data_ba_gabungan,
                urutan="terendah",
            )
        )

        konfigurasi_5_kl = buat_konfigurasi_5_kl(
            tahun_ba_berjalan,
            tanggal_realisasi_ba,
        )

        with tab_5_tertinggi:
            tabel_5_tertinggi = (
                buat_tabel_5_kl_tampilan(
                    data_5_tertinggi,
                    tahun_ba_berjalan,
                    tanggal_realisasi_ba,
                )
            )

            (
                kolom_tabel_5_tertinggi,
                kolom_ringkasan_5_tertinggi,
            ) = st.columns(
                [4, 1.35],
                gap="small",
                border=True,
            )

            with kolom_tabel_5_tertinggi:
                st.dataframe(
                    tabel_5_tertinggi,
                    width="stretch",
                    height="content",
                    hide_index=True,
                    column_config=konfigurasi_5_kl,
                )

            with kolom_ringkasan_5_tertinggi:
                tampilkan_ringkasan_5_kl(
                    data_5_tertinggi,
                    data_ba_gabungan,
                    tahun_ba_berjalan,
                    tahun_ba_sebelumnya,
                    jenis="tertinggi",
                )

        with tab_5_terendah:
            tabel_5_terendah = (
                buat_tabel_5_kl_tampilan(
                    data_5_terendah,
                    tahun_ba_berjalan,
                    tanggal_realisasi_ba,
                )
            )

            (
                kolom_tabel_5_terendah,
                kolom_ringkasan_5_terendah,
            ) = st.columns(
                [4, 1.35],
                gap="small",
                border=True,
            )

            with kolom_tabel_5_terendah:
                st.dataframe(
                    tabel_5_terendah,
                    width="stretch",
                    height="content",
                    hide_index=True,
                    column_config=konfigurasi_5_kl,
                )

            with kolom_ringkasan_5_terendah:
                tampilkan_ringkasan_5_kl(
                    data_5_terendah,
                    data_ba_gabungan,
                    tahun_ba_berjalan,
                    tahun_ba_sebelumnya,
                    jenis="terendah",
                )

    except Exception as error_ba:
        st.error(
            "Gagal mengolah data Bagian Anggaran: "
            f"{error_ba}"
        )
