"""Isi Templat_Brief_dan_Storyboard.docx -> brief/Brief_dan_Storyboard.docx."""
import copy
from pathlib import Path

import docx
from docx.shared import Cm

ROOT = Path(__file__).resolve().parent.parent
d = docx.Document(ROOT / "Templat_Brief_dan_Storyboard.docx")


def isi(cell, teks):
    """Ganti isi sel dengan teks; tiap '\n' jadi paragraf baru dengan format paragraf pertama."""
    p0 = cell.paragraphs[0]
    rpr = next((copy.deepcopy(r._r.rPr) for p in cell.paragraphs for r in p.runs if r._r.rPr is not None), None)
    for p in cell.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    for i, baris in enumerate(teks.split("\n")):
        p = p0 if i == 0 else cell.add_paragraph()
        if i and p0._p.pPr is not None:
            p._p.insert(0, copy.deepcopy(p0._p.pPr))
        run = p.add_run(baris)
        if rpr is not None:
            run._r.insert(0, copy.deepcopy(rpr))


identitas, brief, storyboard, sketsa = d.tables[:4]

isi(identitas.rows[0].cells[1], "[Nama Anggota 1] ([NIM])")
isi(identitas.rows[1].cells[1], "[Nama Anggota 2] ([NIM])")
isi(identitas.rows[2].cells[1], "Selalu Lebih Mahal: Harga Pangan Pokok di Tanah Papua, 2021–2025")

isi(brief.rows[0].cells[1],
    "Badan Pangan Nasional (Panel Harga Pangan), Rata-rata Harga Pangan Bulanan Tingkat Konsumen Provinsi.\n"
    "https://data.badanpangan.go.id/datasetpublications/5up/rata-rata-harga-pangan-bulanan-konsumen-provinsi, "
    "diunduh 31 Oktober 2026. 36.348 baris × 6 kolom: 38 provinsi, 27 komoditas, Januari 2021–Januari 2026.")
isi(brief.rows[1].cells[1],
    "Dengan harga bulanan per provinsi selama lima tahun, kita bisa menguji apakah selisih harga antarwilayah "
    "hanya lonjakan sesaat atau menetap. Angka inflasi nasional tidak memperlihatkan hal ini. Sejak 2024 data "
    "juga memuat empat provinsi hasil pemekaran Papua, jadi ketimpangan di dalam Papua bisa dibandingkan untuk "
    "pertama kalinya.")
isi(brief.rows[2].cells[1],
    "Direktorat Distribusi dan Cadangan Pangan, Deputi Bidang Ketersediaan dan Stabilisasi Pangan, Badan Pangan "
    "Nasional. Unit ini menentukan alokasi program Fasilitasi Distribusi Pangan (subsidi ongkos angkut) antarprovinsi.\n"
    "Mereka sudah tahu harga pangan di Papua mahal. Mereka belum punya angka seberapa besar selisihnya, apakah "
    "selisih itu menyempit dari tahun ke tahun, dan provinsi serta komoditas mana di Papua yang paling berat.")
isi(brief.rows[3].cells[1],
    "\"Pangan pokok di Tanah Papua konsisten sekitar 25% lebih mahal dari median nasional sejak 2021, "
    "dan di pegunungan selisihnya hampir 70%.\"")
isi(brief.rows[4].cells[1],
    "Pada 2024, lima provinsi dengan harga pangan pokok tertinggi relatif terhadap median nasional semuanya "
    "ada di Tanah Papua (Papua Pegunungan 1,67×, Papua Tengah 1,48×, Papua Barat 1,29×, Papua Barat Daya 1,26×, "
    "Papua 1,25×). Data: rasio harga provinsi terhadap median nasional bulanan, rata-rata 12 komoditas yang lengkap.")
isi(brief.rows[5].cells[1],
    "Selisih itu menetap: rasio tahunan Papua dan Papua Barat ada di 1,19–1,34× setiap tahun 2021–2025 tanpa "
    "tren menurun. Data: deret bulanan rasio kedua provinsi yang tercatat sejak 2021.")
isi(brief.rows[6].cells[1],
    "Bebannya tidak merata. Papua Pegunungan (1,67×) dan Papua Tengah (1,48×) jauh di atas provinsi Papua lain, "
    "dengan selisih terbesar pada bawang merah, cabai, dan minyak goreng (1,6–2,0×). Beras di Papua Selatan "
    "hampir setara harga nasional (1,05×). Data: matriks rasio 6 provinsi × 12 komoditas, 2024.")
isi(brief.rows[7].cells[1],
    "1) Memprioritaskan alokasi Fasilitasi Distribusi Pangan ke Papua Pegunungan dan Papua Tengah untuk bawang, "
    "cabai, dan minyak goreng sebagai program rutin, karena selisihnya sudah bertahan lima tahun.\n"
    "2) Menambah titik pantau harga di Papua Pegunungan. Data 2025 provinsi itu 62% kosong, dan angka yang ada "
    "tampak berasal dari sedikit pengamatan.")

baris = [
    ("Setup",
     "Lima provinsi dengan pangan pokok termahal di Indonesia semuanya berada di Tanah Papua.",
     "Dot plot horizontal 38 provinsi, diurutkan dari rasio tertinggi, dengan garis acuan 1,0× (median nasional). "
     "Yang dibandingkan adalah satu nilai antar 38 kategori. Titik lebih ringkas dari batang, dan yang dibaca "
     "adalah jarak dari garis 1,0×, jadi sumbu tidak perlu mulai dari nol.",
     "Enam titik Papua berwarna oranye, 32 provinsi lain abu-abu. Anotasi pada lima titik teratas; label nilai hanya "
     "untuk provinsi Papua, nilai provinsi lain lewat hover."),
    ("Konflik",
     "Selama lima tahun harga di Papua tidak pernah mendekati median nasional.",
     "Grafik garis bulanan Jan 2021–Jan 2026, rasio Papua dan Papua Barat terhadap median nasional. Garis paling "
     "jelas memperlihatkan ada tidaknya tren pada deret waktu. Hanya dua provinsi yang punya deret lima tahun, "
     "jadi hanya ada dua garis.",
     "Pita arsir rentang rasio dan garis putus-putus 1,0×; anotasi nilai terendah selama lima tahun, dan penanda "
     "Januari 2024 saat provinsi pemekaran mulai tercatat."),
    ("Resolusi",
     "Di dalam Papua bebannya tidak sama: pegunungan menanggung paling berat, terutama bumbu dan minyak goreng.",
     "Heatmap 6 provinsi × 12 komoditas, rasio 2024, baris dan kolom diurutkan dari rasio tertinggi. Ada dua dimensi "
     "kategorikal dengan satu nilai per sel, sehingga pola provinsi (baris) dan komoditas (kolom) terbaca sekaligus. "
     "Skala warna satu hue mulai 1,0×: makin gelap makin mahal.",
     "Bingkai pada baris Papua Pegunungan & Papua Tengah dan kolom bawang/cabai/minyak; catatan pada sel beras "
     "Papua Selatan (1,05×) sebagai pembanding. Angka rasio tercetak di tiap sel."),
    ("Resolusi",
     "Angka paling ekstrem datang dari data paling tipis, jadi titik pantau di Papua Pegunungan perlu ditambah.",
     "Strip ketersediaan data: baris 6 provinsi Papua, kolom bulan Jan 2024–Jan 2026, warna = persentase komoditas "
     "yang tercatat. Yang ditanyakan hanya kelengkapan data, jadi encoding sederhana sudah cukup.",
     "Blok kosong Papua Pegunungan Agustus 2025–Januari 2026 dianotasi dan dihubungkan ke rekomendasi kedua."),
    ("Dibuang",
     "(Peta choropleth rasio 2024 per provinsi.)",
     "Dipertimbangkan untuk pertanyaan \"di mana\". Dibuang karena daratan Papua yang luas memenuhi tampilan, "
     "sehingga kesan \"parah\" lebih banyak datang dari ukuran wilayah daripada nilainya. Provinsi kecil (DKI, DIY, "
     "Bali) nyaris tak terlihat, dan peringkat sulit dibaca dari gradasi warna. Dot plot di Gambar 1 menjawab "
     "pertanyaan yang sama.",
     "Ditampilkan di akhir cerita pada bagian \"Gambar yang tidak kami pakai\"."),
]
for row, (bagian, kalimat, bentuk, perhatian) in zip(storyboard.rows[1:], baris):
    for cell, teks in zip(row.cells[1:], (bagian, kalimat, bentuk, perhatian)):
        isi(cell, teks)
for row in storyboard.rows[len(baris) + 1:]:
    row._tr.getparent().remove(row._tr)

sel = sketsa.rows[0].cells[0]
isi(sel, "")
sel.paragraphs[0].add_run().add_picture(str(ROOT / "brief" / "sketsa_storyboard.png"), width=Cm(15))

# Hapus halaman contoh: dari paragraf "CONTOH TERISI" sampai sebelum sectPr
body = d.element.body
anak = list(body.iterchildren())
mulai = next(i for i, el in enumerate(anak) if "CONTOH TERISI" in "".join(el.itertext()))
for el in anak[mulai:-1]:
    body.remove(el)

d.save(ROOT / "brief" / "Brief_dan_Storyboard.docx")
