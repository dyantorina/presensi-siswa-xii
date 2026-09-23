import os
from datetime import datetime, timezone, timedelta
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st

# 1. Konfigurasi Halaman
st.set_page_config(
    page_title="Presensi Siswa Kelas XII",
    page_icon="📝",
    layout="centered"
)

# Zona Waktu Indonesia Barat (WIB = UTC+7) bawaan Python
WIB = timezone(timedelta(hours=7))
now = datetime.now(WIB)
today_str = now.strftime("%Y-%m-%d")
time_str = now.strftime("%H:%M:%S")

# Jadwal Mapel Tetap (Senin - Jumat)
MAPEL_HARIAN = {
    0: "Bahasa Indonesia",  # Senin
    1: "Matematika",        # Selasa
    2: "Bahasa Inggris",    # Rabu
    3: "KIK",               # Kamis
    4: "Kejuruan",          # Jumat
}

# Pilihan Mapel Khusus Sabtu & Minggu (5 Mapel Utama)
PILIHAN_MAPEL_AKHIR_PEKAN = [
    "-- Pilih Mata Pelajaran --",
    "Bahasa Indonesia",
    "Matematika",
    "Bahasa Inggris",
    "KIK",
    "Kejuruan",
]

# Deteksi Hari: Sabtu (5) atau Minggu (6)
is_akhir_pekan = now.weekday() in [5, 6]
mapel_terpilih = None

if not is_akhir_pekan:
    mapel_terpilih = MAPEL_HARIAN.get(now.weekday())

# 2. Master Data Siswa Kelas XII (Hasil Ekstrak File Excel)
DATA_SISWA_PER_KELAS = {
    "XII GEOMATIKA": [
        "ADITYA ARDIYANSAH",
        "AKBAR EKA PRATAMA",
        "ALFARIJI",
        "ALFINO TEGAR RADITYA",
        "ALIF AKMAL AMANULLAH",
        "ANGGUN APRIANI",
        "ARIF MAULIDANI",
        "BIMO PRASTYO",
        "DAFFA' AMIRUL MUKMININ",
        "DESTY AIDHA CAHYANI",
        "ENDI JULIAN SAPUTRA",
        "FABIANSYAH HAIDAR ARASH",
        "FADLI HARDIYANSYAH",
        "FAJAR RAJAFLI",
        "IBRAHIM KAMIL",
        "KEISHA AKMAL MULA RISQI",
        "KHAYLA ALTAFUNNISA DEWI",
        "KHENZO VIGO HENDRIAN",
        "MUHAMAD ALI IMBRON ABDULLAH",
        "MUHAMMAD ALDIANSYAH",
        "NADYA D SASTALIMAH",
        "NOVAL MARDIANSYAH",
        "RACHMI NUR KOMALASARI",
        "REFANA SIFA",
        "RIDWAN KURNIAWAN",
        "WASILAH RAHMAH",
        "ZAZKIA PUTRI RAHAYU",
    ],
    "XII DPIB 1": [
        "ABRAM JOSEPH MOLO",
        "AFIDA KHOIRUN NISA",
        "AHMAD SATRIO",
        "AKMAL MAULANA",
        "ALISA CAHYANI",
        "AZ ALIYY RIZQI",
        "AZRIL ALFARIZ VALENCIA",
        "CHANTYQ ANGGRAENI DYAH PUSPO ARUM",
        "DAFFA RAKHA AR-RAYYAN",
        "DHIMAS RIZKY ADITYA",
        "DHINA SYAFIRA",
        "DIMAS SULTAN ADITYA",
        "FAZAR ADHI PERMANA",
        "FERRY SETYAWAN",
        "GALIH DIMAS PRAYOGA",
        "M. NATHAN AZHARI",
        "MAITSA'A HANAAN",
        "MALIK KUS SOLEH",
        "MOHAMMAD DIMAS PRAKASA",
        "MUHAMAD ALFARISY",
        "MUHAMAD NUR SULTAN",
        "MUHAMAD RIKI SAPUTRA",
        "MUHAMMAD NAUFAL GHAISAN DANISH",
        "PANJI FAJRIANA",
        "PITRIYANI",
        "RAISSA AULIA HAFIZ",
        "RATNA GALIH",
        "RAYHAN YANUAR HIDAYAT",
        "RISKA MAULIDA",
        "RULI EVANZA",
        "SABRINA NASYA SYAHIRA",
        "SOPIAN PALRIJI",
        "VALDIO YUDHAR PRANANTA",
        "WIDIA MAULANI",
    ],
    "XII DPIB 2": [
        "ADITYA PRATAMA",
        "AL IKHRAM RAKASIWI",
        "ANGGRAENI ANINDA PUTRI",
        "ARGA IRAWAN",
        "DELISA ISWADI PUTRI",
        "DELLU JUAN IRSYAD",
        "DEWI MA'RIFATUS SA'ADAH",
        "DHIAJANNAH ZAKIRA SETIAWAN",
        "ERLANGGA BASTIYAN",
        "EVAN ANDIKA CHANDRA",
        "FACHRY CESSA PRADANA",
        "HAIDIR ABDI RASYID",
        "HILMY ABDULHASIB",
        "INDAH PUSPITA SARI",
        "INTAN SABITA AZHAR",
        "KEVIN ARDYANI",
        "LUTHFI MUZAKKY RAMADHANY",
        "M ZAHRON ABRAR",
        "MUHAMAD AZRIL NUR SYAFIQ",
        "MUHAMAD FAKHRI",
        "MUHAMMAD ALI RIDHO",
        "MUHAMMAD FAIZ HABIBI",
        "MUHAMMAD FARIZ AFFAN ILYAS",
        "MUHAMMAD MAHESA AL HABSY",
        "MUHAMMAD RESTU",
        "MUHAMMAD SYADDAD NURALLAM",
        "NEIL IZZI HAFI ZULFA",
        "NUR RIZQULLAH KANAZAWA",
        "RADEN MALKAN AONILLAH FRINGGA KUSUMA",
        "RICKY ALFARIZKY",
        "RISTI MULYANA",
        "RIZKY RAMADHAN",
        "ROSDIANA RIZKY",
        "SILVIA ANGGRAINI",
        "WILDAN SUWARSO",
    ],
    "XII TKP": [
        "ABI MANYU BADRAYANA",
        "ABID FAUZAN PUTRA RIANSYAH",
        "AGNE AURA PERTIWI",
        "AKHMALUL RAPHY",
        "ALFATHIR KAUTSAR",
        "ARIQ DHAMIRI SALIM",
        "AZKA CINTA INDATA ATHHAR",
        "AZKIA ALTHAFUNNISA ODYSA",
        "BOAS HASUDUNGAN HUTAURUK",
        "FIRZA ALVA WIRAWAN",
        "FITRIA AZAHRA",
        "HANAS FAUZAN FATURRAHMAN",
        "JOSHUA TONIE WIJAYA",
        "KRESNA SAPUTRA",
        "LAILA LINDA SARI",
        "M. FITRAH RAMDANI",
        "MALIA NAYA SASTRI",
        "MOHAMMAD RIDHO SYAHRAFI",
        "MUHAMAD ALFAT KURNIAWAN",
        "MUHAMAD REZA FAHLEVI",
        "MUHAMMAD AZZAM DWI CAHYO",
        "MUHAMMAD NUR AL-FATI",
        "NABIL ARIEF PERDANA",
        "NURUL AINI ASSYIFFA",
        "RADITYA GALANG PRAMUDIA",
        "RANGGA SAPUTRA",
        "RENDY ALDIANSYAH",
        "REVANI NAZWA ANGGRAENI",
        "RIDHO SAPUTRA",
        "SELINA RAHAYU",
        "SHELLY PUTRI FADILLAH",
        "SUKMA AYU LESTARI",
    ],
    "XII TPG": [
        "ACHMAD BAHATIAR ASSIRADJ",
        "ACHMAD SAFIKI",
        "AHMAD AL FARABY",
        "AJI ROSYADI",
        "ALKAUSAR RAJA LANGIT BUANA",
        "ALLYTHEA RACHMAN",
        "ANANDA GUSTI FIRMANSYAH",
        "BIMA SANDI PRASETYO",
        "DESFERADO MARCIANO ANAKOTTAPARY",
        "DIMAS ANGGA RIYADI",
        "FADLAN",
        "IKHSAN DE BARLI",
        "IKRAM PURNOMO",
        "IPAN PEBRIAN",
        "JUAN FERNANDES SIMANGUNSONG",
        "KEYSYAR AHMAD YUSUF",
        "KHAERUL ILMAN",
        "LALU SANDI SAPUTRA",
        "MANZILIRRAHMAH",
        "MOHAMAD NUR HAKIM",
        "MUHAMAD ALI SODIKIN",
        "MUHAMAD RIZKY ADITYA DARMA",
        "MUHAMAD SAEFULLOH",
        "MUHAMMAD KHADAFFI AL FARUQ",
        "NAZMI ALFAJRI",
        "NAZUAN ALIF IZHARDI",
        "NUKE DINATA",
        "PUTRI OKTAVIANI",
        "RADITIA MAULANA",
        "REYHAN JAINI",
        "SATRIA AL FARAZH",
        "WILY ANDRIYAN",
    ],
    "XII TITL 1": [
        "ABIMANYU SATRIA WICAKSANA",
        "ABU YASID",
        "ACHMAD NAQA",
        "ALI SHOVIAN",
        "ALIF ZIYAD CAHYADI",
        "ANDIKA EKA PRATAMA",
        "ANDRIANSYAH PRATAMA",
        "BAGUS HERMAWAN",
        "BAYU TRI ANTONI",
        "DAFFA ALBUKHORI",
        "DANENDRA HAFIZH AR ROYYAN",
        "DANIEL SATRIO WIBISONO",
        "ERYKO MUBAROK",
        "FADHIL RAFI ANWAR",
        "FATIFAR AL AZIZ",
        "FIKRI ALDIANSYAH",
        "HAIKAL NUR HAKIKI",
        "IBNU AL FATHAR KURNIAWAN",
        "KHOIRUL APRILLANA YUSUF",
        "MUHAMMAD AFITD ZAKUAN ARDAN",
        "MUHAMMAD FAHMI NUR HAFIIZHU",
        "MUHAMMAD FAHRUR RAZI",
        "MUHAMMAD RIZKI ABDILAH",
        "MUHAMMAD ZAKI RAMADHAN",
        "NAUFAL KAAMIL",
        "RENDI JUANTA IRAWAN",
        "RHEVAN ADITYA PRATAMA",
        "RIDO SAPUTRA",
        "RIZKY PASHA RAMADHAN",
        "SATRIA MUHAMMAD FATIH",
        "VERNANDA RAFKA ERAL PRASSETYA",
        "YUSUF HAMDANY",
    ],
    "XII TITL 2": [
        "AL FATHIR KHOIRULLAH",
        "ALIF AL FAZRI",
        "ALIP SURAHIYAN",
        "ARYA TRI CAHYONO",
        "AZKA FAIZ ALFATIH",
        "BINTANG WURI SANDI",
        "CALVIN ARVIANSYAH",
        "DAFFA RAJAULAKBAR",
        "DEWO WICAKSONO WIDODO",
        "ENGGAR GALEH WIJAYANTO",
        "EQ LESMANA",
        "FATUR ROHMAN AL-FAUZI",
        "FEBBY MUHAMAD RIZKI",
        "GILANG RAMADHAN",
        "HAICAL JULIANSYAH",
        "HALIFA MUSTAQIM",
        "IBNU SENO",
        "ILHAM EGY SAPUTRA",
        "IRSYAD FAIZ FATURRAHMAN",
        "MAS RUQY HADI MAULANA",
        "MUHAMAD ALVA FACHREZA",
        "MUHAMAD DAFA ISNANTO",
        "MUHAMAD FIQRI",
        "MUHAMAD HOIDIR",
        "MUHAMAD ZIDAN FEBRIANO",
        "MUHAMMAD IKHLAS",
        "MUHAMMAD RHEFAN KHADAFI",
        "NAUFAL FATHONI",
        "RANGGA SAPTAGUSTIADI",
        "RIFQI TAUFIK HAKIM",
        "RIKO RISKIANO",
        "YUAN AKBAR PUTRANTO",
        "YUSUF PRASETYO",
    ],
    "XII TITL 3": [
        "AKMAL ZIDA",
        "ALOYSIUS WIDIYA HARTANTO",
        "ALWAN FA'IQ HIDAYAT",
        "ANDIKA MUHAMAD AKBAR",
        "ANDRA FAUZI ESSA",
        "AZRIL FAJAR RISMANA",
        "BARRY NAUFAL ALKAHFY",
        "BILLY ANGKA JAYA PUTRA",
        "DANISH HANA NURCAHYA",
        "DWI FAJRI WAHYUDI",
        "EVAN HENDARTO",
        "FAJRI FADILLAH SABIL",
        "FALDAN BASIR",
        "FARROS ALFRIDHO",
        "FATAN FATIHUDIN",
        "GERALDI MAULANA PRADITA",
        "KAYRUL AZZAM",
        "KEVIN ALAMSYAH",
        "MARVEL KRISTIAWAN",
        "MAULANA MALIK ZHAIRI",
        "MOCHAMAD DANU SAPUTRA",
        "MUCHLIS ZAMZAMI",
        "MUHAMAD RIZKY",
        "MUHAMAD RIZKY ADHITYA",
        "MUHAMMAD DZAKIY MAULIDY",
        "MUHAMMAD FERDIANSYAH HARTADI",
        "MUHAMMAD RAYHAN PURNAMA",
        "MUHAMMAD RIANDRA GUNAWAN",
        "MUHAMMAD RIFKY ALAMSYAH",
        "RADHITYA EKA SYAPUTRA",
        "RAFAEL MATHEW ALEXIS",
        "RAFI PANGESTU",
        "RIZQI DIAZ KURNIAWAN",
        "SUHERMAWAN",
    ],
    "XII TOI": [
        "ACHMAD MALVINO JULIAN SALIM",
        "AFIF RAFI SALMAN",
        "AHMAD FAUZAN PRATAMA",
        "AKBAR ZAELANI",
        "DEVAN KARNALIM",
        "FAHRI RAMADHAN",
        "FARREL RALDI RUS SYAMSI",
        "FAUZAN ANSHARY",
        "FAUZI PUTRA ARYA PRATAMA",
        "INDRA MAULANA GUNAWAN",
        "LUKAS FERNANDO HUTABARAT",
        "MIFTAKUSSURURI",
        "MOCHAMMAD AZZAM SYAHPUTERA",
        "MOH DOFAN",
        "MOHAMAD RAFIKHSASUIATUL NUGROHO",
        "MOHAMMAD RENDY SAPUTRA",
        "MUHAMAD RIZKI GURBI",
        "MUHAMMAD CHELVIN MAULANA",
        "MUHAMMAD JAVIER HAFIZH",
        "MUHAMMAD LANGIT RAMADHAN",
        "MUHAMMAD MUSA AFRIYANSYAH",
        "NAKULA PUTRA HARYANTO",
        "RADITH PRASTYO",
        "RADITTYA AL FARIS",
        "RAFA KARUNIA SUPARNO",
        "RAFI ADRIYAN TRI FADHLA",
        "RAIHAN AKBAR KEVIANSYAH",
        "RAJIV RUKI RAMJANI",
        "RANGGA BAYU PERMANA",
        "RESSYA RABBANY ALFARIS",
        "SATRIO MAULANA MALIK",
        "SUCIWATI NURUL RAHAYU",
        "SYAHDAN ALI RAFSANJANI",
        "TOMI ROMANTIMO TSE",
        "ZULVIAN ADI VILANO",
    ],
    "XII PEMESINAN 1": [
        "AHMAD RAFFI PRATAMA",
        "AL DAFI ADRATAMA",
        "ANDIKA ANINDYA GUNA ISWANTO",
        "AUREL FERLY PRATAMA",
        "DAFFA ABID RIVALDI",
        "DAVE BES HALIEL GINTING",
        "DAVI RIANDRA YUSUF",
        "DHIKA NUR FAHRI",
        "DZAKIY ALHAFIIZH ELDIANSYAH",
        "FERAINSEL MARSHALL SIDABUTAR",
        "FINO MARVEL AZERA",
        "LUQMAN ABDUL MALIK",
        "M. AFGAN MEILANDRI",
        "MARSEL NURDIANSYAH",
        "MUHAMAD ADITYA",
        "MUHAMAD AL FAHRI",
        "MUHAMAD FARDAN SYAH REZA",
        "MUHAMAD FAUZAN",
        "MUHAMAD RAFI AULIA",
        "MUHAMAD SABIL CHOIRIE",
        "MUHAMAD TRI SURYA PERKASA",
        "MUHAMMAD ALFI",
        "MUHAMMAD FAKHRI UBAIDILLAH",
        "MUHAMMAD FAREL",
        "MUHAMMAD RAIHAN",
        "RAEFAN BRILIANT UTAMA",
        "RAIHAN HANNIP AL MUSYAFFA",
        "RIDWAN DHUTA PALEMA",
        "RIFQIY NAJIYULLAH",
        "RIZKY AGUNG SETIAWAN",
        "SATRIA ADI PRATAMA",
        "TB. SAGIL WARIDI",
        "YONATHAN KURNIA",
        "ZAKY ANANDA ZEIN",
    ],
    "XII PEMESINAN 2": [
        "ACHMAD IKHSAN",
        "ADHITRI NOVIANTO",
        "AHMAD AL RASIQ",
        "ANDIKA FEBRIANSYAH",
        "ANDIKA PRATAMA",
        "APRIFAIZ PUJI NAWWAF ELFARRASI",
        "BANGKIT REZA PRASTIA",
        "DIMAS TUNGGUL WICAKSONO",
        "DODY RAHMAT DARMAWAN",
        "FAHRI NAFI NASYID NUR ILHAM",
        "FAJAR RASYIID",
        "FIKRI RAMADANI",
        "GIAN RADITYA AKMAL",
        "GILANG TRI AHMADANI",
        "HAFIZ ZULFIKAR",
        "HERU VERNANDO",
        "IBNU SENA",
        "INDRA PUTRA RAMADAN",
        "JULIAN RIZKY ADITYA",
        "MUHAMAD DZIKRI APRIANSYAH",
        "MUHAMAD KHOERI ROFIQ",
        "MUHAMMAD AFIF SUDRAJAT",
        "MUHAMMAD DRIAN FATHURRAHMAN",
        "MUHAMMAD IBNU ABBAS",
        "MUHAMMAD IKRAM AZ ZAHRAN",
        "PRANAYA ADISYA PUTRA",
        "RADITHIA DANIAR PUTRA",
        "RADITYA SYIFA AJI WICAKSANA",
        "RASYA ERLANGGA PRATAMA",
        "RENOVAN ADHITAMA",
        "RICKY MAULANA",
        "SAMUEL ADRIAN ALVINO",
        "SUKOCO SURYO PRABOWO",
        "ZIMMY ROMADHONA",
    ],
    "XII TMI": [
        "ACHMAD RAFIQIN",
        "ADHE SETYONO PAMUNGKAS",
        "AFDAL ANDRIANTO",
        "ALFARO GINZA ALANA",
        "ASHIL NUR KASFIL AZIZ",
        "BINTANG RAMADHAN CATUR YOGA",
        "CIK ALIM ARYA PUTRA",
        "DAVIDIO RAFFA APRILIANO PUTRA",
        "DWI ANDIKA NUR ROHMAN",
        "DYLAN PATTO FERDIAN",
        "FAIZ AQILA SAPUTRA",
        "FREDRICH ADRI SALOMO MARBUN",
        "GALANG RAMBU ANARKI",
        "KEIVARO MAYNARD ARDHANA",
        "LINGGARJATI",
        "MARCELINO SUOTH PANGGEY",
        "MOHAMMAD IMAM'MUL MUTTAQIN",
        "MUHAMAD AHDA RAMDANI",
        "MUHAMAD IRGI",
        "MUHAMAD RAGIL",
        "MUHAMAD SYUKUR PRATAMA",
        "MUHAMMAD DENI GUNAWAN",
        "MUHAMMAD ILHAM",
        "MUHAMMAD RICKY FIRMANSYAH",
        "MUHAMMAD RIDWAN",
        "MUHAMMAD RIPAL PAHREZI NASUTION",
        "PANJI HERMANSAH",
        "RADIT DWI PUTRA",
        "RADITYA DWI BAGASKARA",
        "RAMDAN SAFAWI",
        "RAYA SEGARA",
        "RIDHO FAHMI AZZAM",
        "SAFYAN HUGENG INDRIYANO",
        "SHENO DWI AJI",
        "SYAHREZA WAHYU GUMMELAR",
        "WILDAN CAKRA AULIA",
    ],
    "XII DGM": [
        "ABDUL FAHRI",
        "ABDUL MUNTHOLIB",
        "ADRYAN MUTI",
        "AKBAR KURNIAWAN",
        "CATUR RIFQHI RADITYA",
        "DAVID SUSANDY",
        "ELZYA SANSETIA AYU",
        "ERLANGGA PRASETYO",
        "FAHMI ADLIY NURFAUZI",
        "FARAEL YAS",
        "FATHIR FATHURRAHMAN",
        "FATHURROZAAQ ALIANSYAH",
        "FIERO APREZA KARBUY",
        "GUNTUR TETUKO WIBOWO",
        "IRSYAD FAIRUZ SUDRAJAT",
        "IYAN ANDRIYANSYAH",
        "MAULANA SAPUTRA",
        "MUHAMAD ALFIANSYAH PRIANGGA",
        "MUHAMAD BADAR TAMIM AL HUSNI",
        "MUHAMAD FATHIR UTOMO",
        "MUHAMAD PAHARUL ROSI",
        "MUHAMAD RIDHO",
        "MUHAMMAD RIZKY",
        "MUHAMAD ROZAKI",
        "MUHAMMAD DIMAS HERNANDA",
        "MUHAMMAD DZAKY",
        "MUHAMMAD IHSAN NURAHMAN",
        "MUHAMMAD ILYAS",
        "PRAMADHANU FRITZI AKBAR",
        "RAFFA ADHITYA AL FISYAR",
        "ROMI CAHYA KUSNADI",
        "RULI ALPIANSYAH",
        "SOFIYAH RAHMAH",
        "VICKY FAKHIR HIBATULLAH",
        "ZAHRATUSITA",
    ],
    "XII RPL 1": [
        "ABILLISHA FACHRY",
        "ACHMAD ZAELANI PUTRA",
        "ADITYA ALFAIZ",
        "ADRIAN TIKTA ADITYA",
        "AHMAD RAFI FADHILAH",
        "ALEX NICHOLAS RAJA RATU",
        "ALMAS FILZA",
        "ARIA ATTALA AL SYAMFATHDRIANDRYA SIGIT",
        "ARVARIZZY CHANDRAWINATA",
        "BAMBANG WIRAYUDA",
        "DESY JULIATI",
        "FAHRI KURNIAWAN",
        "FATAH RIZKI GUNAWAN",
        "FIKRI AZIZ MUDZAKIR",
        "FIRRAS RADIY HUTOMO",
        "GIBRAN RASYAD ANDI SURYA",
        "HESSA NADIRA SYAFA",
        "IBAS SINATRYA GUNARDI",
        "LING MAYDARLING",
        "MUHAMAD AFIF AL GHOZALI",
        "MUHAMMAD AKBAR HEZRA",
        "MUHAMMAD ALFI RIZKI",
        "MUHAMMAD ENDRA AL-FATEH",
        "MUHAMMAD PRAYUDA AL-FATIH",
        "NOVIANA ANUGRAH RIZKY CARENEA",
        "OVAN SEPTY RAMADHAN",
        "RAIHANA KHAIROTUNNISA",
        "RAKA RAISSA PUTRA",
        "RANI PERMATASARI",
        "RAYHAN HARI KARTIKO",
        "ROIS ALFIAN",
        "TAJMADHAN FIRAS AKENO",
        "VINA BAASITHUNISA",
        "YUDA HADI PERMANA",
        "ZEIN ARYO ABDILLAH",
    ],
    "XII RPL 2": [
        "AHMAD FAISAL",
        "ANDIKA DWI PRATAMA",
        "ARDIKA DWI PRATAMA",
        "ARJUNA KENSHIMADA SURAHMAN",
        "ATTAR DIANDRA AL - FATHIR",
        "AYUNINGTIAS FEBRIANTI SUBARTHA",
        "AZIZAH MUTYA TSANNY",
        "BIMO SETYO NUGROHO",
        "FABIAN KHOERUL UMMAM",
        "FADHILAH HARYADI",
        "FAHRI REVAN AHMAD",
        "FAUZAN SAMMY MOCHTAR",
        "FINA NAILATUL IZZAH",
        "KARTIKA CAHYANING RATRI",
        "KHAYLA KEYSHA AFRINA",
        "KYARA APRILIA KHAIRUNISA",
        "MUHAMAD AL FAUZAN",
        "MUHAMAD RIDHO",
        "MUHAMAD SIDIQ FADILA",
        "MUHAMAT RAIHAN",
        "MUHAMMAD ABYAN FARRAS",
        "MUHAMMAD ADIL FARRAS ALFAT",
        "MUHAMMAD BAGUS FEBRIAN",
        "MUHAMMAD ZIYADDINO PUTRA HERNOWO",
        "NICODEMUS KAISER",
        "RASSYA ATHALLA RIZQI",
        "RAYHAN AULIA GUNAWAN",
        "REZA FADILLAH",
        "SHENDIANSYAH",
        "SURAHMAN",
        "SYAFIRA RAHMATARI",
        "SYARIF HIDAYATULLAH",
        "YUDDI CHRISNALDI",
        "ZIVARA AURELYTA AHMAD",
    ],
}


# 3. Koneksi Google Sheets (Fleksibel: Lokal kunci.json & Cloud st.secrets)


@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    # Cek apakah ada file kunci.json di folder proyek
    if os.path.exists("kunci.json"):
        credentials = Credentials.from_service_account_file(
            "kunci.json", scopes=scopes)
    elif os.path.exists("../kunci.json"):
        credentials = Credentials.from_service_account_file(
            "../kunci.json", scopes=scopes)
    elif "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(
            creds_dict, scopes=scopes)
    else:
        st.error(
            "Konfigurasi kunci Google Sheets (kunci.json atau st.secrets) belum ditemukan.")
        st.stop()

    return gspread.authorize(credentials)


def init_worksheet(client, spreadsheet_title, sheet_name):
    try:
        sh = client.open(spreadsheet_title)
    except Exception:
        st.error(
            f"Gagal membuka spreadsheet '{spreadsheet_title}'. Pastikan nama persis sama dan bot sudah dijadikan Editor."
        )
        return None

    try:
        ws = sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        # Otomatis buat worksheet per kelas jika belum ada
        ws = sh.add_worksheet(title=sheet_name, rows="500", cols="10")
        ws.append_row(
            ["Tanggal", "Waktu", "Nama Siswa", "Status Kehadiran", "Mata Pelajaran"]
        )
    return ws


def rekap_otomatis_tidak_hadir(client, spreadsheet_title, kelas, daftar_siswa_kelas, tanggal_kemarin, mapel_kemarin):
    """
    Mengecek siswa yang belum absen pada hari sebelumnya dan mencatatnya sebagai 'Tidak Hadir'.
    """
    try:
        sh = client.open(spreadsheet_title)
        ws = sh.worksheet(kelas)
    except Exception:
        return

    records = ws.get_all_records()
    if not records:
        return

    df_sheet = pd.DataFrame(records)
    if "Tanggal" not in df_sheet.columns or "Nama Siswa" not in df_sheet.columns:
        return

    # Ambil siswa yang sudah tercatat pada tanggal kemarin
    df_kemarin = df_sheet[df_sheet["Tanggal"] == tanggal_kemarin]
    siswa_sudah_tercatat = set(df_kemarin["Nama Siswa"].tolist())

    # Cari siswa yang BELUM ada datanya sama sekali di tanggal kemarin
    baris_tidak_hadir = []
    for nama in daftar_siswa_kelas:
        if nama not in siswa_sudah_tercatat:
            baris_tidak_hadir.append([
                tanggal_kemarin,
                "23:59:59",
                nama,
                "Tidak Hadir",
                mapel_kemarin
            ])

    # Tambahkan sekaligus ke Google Sheet
    if baris_tidak_hadir:
        ws.append_rows(baris_tidak_hadir)


SPREADSHEET_NAME = "Database_Presensi_Siswa"

try:
    gc = get_gspread_client()
except Exception:
    st.error("Konfigurasi Google Secrets belum lengkap.")
    st.stop()

# 4. Tampilan Antarmuka
st.title("📝 Presensi Siswa Mandiri")
st.caption(
    f"📅 **{now.strftime('%A, %d %B %Y')}** | ⏰ **{time_str} WIB**"
)
# Tampilan Informasi Mapel Hari Ini
if not is_akhir_pekan:
    st.info(f"📚 **Mata Pelajaran Hari Ini:** {mapel_terpilih}")
else:
    st.warning(
        "📅 **Hari Akhir Pekan (Sabtu/Minggu):** Silakan pilih mata pelajaran yang diikuti.")

st.markdown("---")

# Dropdown Kelas
pilihan_kelas = list(DATA_SISWA_PER_KELAS.keys())
kelas_terpilih = st.selectbox(
    "1. Pilih Kelas:",
    options=["-- Pilih Kelas --"] + pilihan_kelas,
    index=0
)

if kelas_terpilih != "-- Pilih Kelas --":
    daftar_nama = sorted(DATA_SISWA_PER_KELAS[kelas_terpilih])
# Cek & rekap otomatis siswa yang tidak hadir di hari KBM sebelumnya
    kemarin = now - timedelta(days=1)
    kemarin_str = kemarin.strftime("%Y-%m-%d")

    if kemarin.weekday() in MAPEL_HARIAN:
        mapel_kemarin = MAPEL_HARIAN[kemarin.weekday()]
        rekap_otomatis_tidak_hadir(
            gc,
            SPREADSHEET_NAME,
            kelas_terpilih,
            daftar_nama,
            kemarin_str,
            mapel_kemarin
        )
    nama_terpilih = st.selectbox(
        f"2. Pilih Nama Siswa ({kelas_terpilih}):",
        options=["-- Pilih Nama --"] + daftar_nama,
        index=0
    )

    # Dropdown hanya muncul di hari Sabtu dan Minggu
    if is_akhir_pekan:
        mapel_input = st.selectbox(
            "3. Pilih Mata Pelajaran:",
            options=PILIHAN_MAPEL_AKHIR_PEKAN,
            index=0
        )
        if mapel_input != "-- Pilih Mata Pelajaran --":
            mapel_terpilih = mapel_input
        else:
            mapel_terpilih = None

    st.write("")

    if st.button("🚀 Kirim Presensi Hadir", use_container_width=True):
        if nama_terpilih == "-- Pilih Nama --":
            st.error("Silakan pilih nama Anda terlebih dahulu!")
        elif is_akhir_pekan and not mapel_terpilih:
            st.error("Silakan pilih mata pelajaran terlebih dahulu!")
        else:
            with st.spinner("Mencatat ke Google Sheets..."):
                ws = init_worksheet(gc, SPREADSHEET_NAME, kelas_terpilih)

                if ws:
                    records = ws.get_all_records()
                    sudah_absen = False

                    if records:
                        df_sheet = pd.DataFrame(records)
                        if "Tanggal" in df_sheet.columns and "Nama Siswa" in df_sheet.columns:
                            cek = df_sheet[
                                (df_sheet["Tanggal"] == today_str) &
                                (df_sheet["Nama Siswa"] == nama_terpilih)
                            ]
                            if not cek.empty:
                                sudah_absen = True

                    if sudah_absen:
                        st.warning(
                            f"⚠️ **{nama_terpilih}** ({kelas_terpilih}) sudah absen hari ini ({today_str}). Sampai jumpa besok!")
                    else:
                        ws.append_row([
                            today_str,
                            time_str,
                            nama_terpilih,
                            "Hadir",
                            mapel_terpilih
                        ])
                        st.balloons()
                        st.success(
                            f"✅ Berhasil! **{nama_terpilih}** tercatat **Hadir** untuk mapel **{mapel_terpilih}**.")
else:
    st.info("👆 Silakan pilih kelas terlebih dahulu.")
