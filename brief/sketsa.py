"""Sketsa wireframe storyboard (dibuat sebelum gambar final, tanpa data sungguhan)."""
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

plt.xkcd(scale=1, length=120, randomness=2)
AKSEN, ABU = "#c0392b", "#9a9a9a"

fig = plt.figure(figsize=(11, 15))

# Tata letak halaman: urutan dari atas ke bawah
ax = fig.add_axes([0.03, 0.05, 0.27, 0.9]); ax.set_xlim(0, 10); ax.set_ylim(0, 100); ax.axis("off")
ax.set_title("Tata letak halaman", fontsize=13)
blok = [("Judul + PESAN UTAMA (1 kalimat)", 92, 6), ("Pengantar singkat", 85, 4),
        ("Gbr 1  SETUP\ndot plot 38 prov", 66, 16), ("Gbr 2  KONFLIK\ngaris 2021-2026", 46, 16),
        ("Gbr 3  RESOLUSI\nheatmap 6 x 12", 26, 16), ("Gbr 4  data bolong\n-> titik pantau", 12, 10),
        ("Rekomendasi | Sumber | Limitasi\nGambar yang dibuang", 1, 8)]
for teks, y, h in blok:
    ax.add_patch(Rectangle((0.5, y), 9, h, fill=False, lw=1.5))
    ax.text(5, y + h / 2, teks, ha="center", va="center", fontsize=9)

def kotak(rect, judul):
    a = fig.add_axes(rect); a.set_xticks([]); a.set_yticks([])
    a.set_title(judul, fontsize=10, loc="left"); return a

# Gbr 1: dot plot provinsi, garis acuan 1,0x, titik Papua beraksen
a = kotak([0.38, 0.74, 0.58, 0.2], "Gbr 1 (Setup): '5 provinsi termahal semuanya di Papua'")
vals = [1.67, 1.5, 1.3, 1.28, 1.26, 1.2, 1.18, 1.16, 1.12, 1.1, 1.05, 1.0, .97, .95, .92]
for i, v in enumerate(vals):
    a.plot(v, len(vals) - i, "o", color=AKSEN if i < 5 else ABU, ms=7)
a.axvline(1.0, ls="--", color="k", lw=1); a.text(1.01, 0.3, "median\nnasional = 1,0x", fontsize=8)
a.annotate("5 teratas = Papua", xy=(1.45, 13), xytext=(1.2, 7), arrowprops=dict(arrowstyle="->"), fontsize=9)

# Gbr 2: garis rasio dari waktu ke waktu, pita 1,2-1,35x
a = kotak([0.38, 0.5, 0.58, 0.19], "Gbr 2 (Konflik): 'tidak pernah turun di bawah 1,2x'")
x = range(60)
a.axhspan(1.2, 1.35, color=AKSEN, alpha=0.15)
a.plot(x, [1.27 + 0.04 * ((i % 7) - 3) / 3 for i in x], color=AKSEN)
a.plot(x, [1.0] * 60, "--", color="k", lw=1); a.set_ylim(0.9, 1.45)
a.text(2, 1.02, "median nasional", fontsize=8); a.text(30, 1.37, "pita 1,2-1,35x selama 5 tahun", fontsize=8)

# Gbr 3: heatmap provinsi x komoditas, sorot baris & kolom
a = kotak([0.38, 0.26, 0.58, 0.19], "Gbr 3 (Resolusi): 'terberat di pegunungan: bawang, cabai, minyak'")
a.set_xlim(0, 12); a.set_ylim(0, 6)
for r in range(6):
    for c in range(12):
        a.add_patch(Rectangle((c, r), 1, 1, fill=True, alpha=0.08 + 0.12 * (r > 3) + 0.1 * (c < 3), color=AKSEN))
a.add_patch(Rectangle((0, 4), 12, 2, fill=False, lw=2.5))
a.add_patch(Rectangle((0, 0), 3, 6, fill=False, lw=2.5, ls="--"))
a.text(6, 5.4, "baris Pegunungan & Tengah", fontsize=8, ha="center"); a.text(0.2, -0.7, "bawang, cabai, minyak", fontsize=8)

# Gbr 4: strip ketersediaan data per bulan
a = kotak([0.38, 0.06, 0.58, 0.15], "Gbr 4: 'Pegunungan 2025: separuh data kosong'")
a.set_xlim(0, 25); a.set_ylim(0, 6)
for r in range(6):
    for c in range(25):
        kosong = r == 5 and 19 <= c < 24
        a.add_patch(Rectangle((c, r), 0.9, 0.8, color="white" if kosong else ABU, ec="k", lw=0.5))
a.annotate("-> tambah titik pantau", xy=(21, 5.4), xytext=(8, 2.5), arrowprops=dict(arrowstyle="->"), fontsize=9)

fig.suptitle("Storyboard: Harga pangan di Tanah Papua (sketsa kasar, tanpa data)", fontsize=14)
fig.savefig(__file__.replace("sketsa.py", "sketsa_storyboard.png"), dpi=130)
