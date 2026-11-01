# Selalu Lebih Mahal: Harga Pangan Pokok di Tanah Papua, 2021–2025

Midterm Project Visualisasi Data Lanjut (CAK7AGB4), Magister Informatika, Telkom University.
Erlangga Dewa Sakti (203012510035) · Alya Nabila Muliani (203012510013)

> Pangan pokok di Tanah Papua konsisten sekitar 25% lebih mahal dari median nasional sejak 2021, dan di pegunungan selisihnya hampir 70%.

- **Data story:** `docs/index.html` (GitHub Pages: `https://<username>.github.io/<repo>/`)
- **Message Brief & Storyboard:** `brief/Brief_dan_Storyboard.pdf`
- **Notebook:** `analisis.ipynb` menghasilkan seluruh angka dan gambar

## Data

Badan Pangan Nasional, *Rata-rata Harga Pangan Bulanan Tingkat Konsumen Provinsi*,
https://data.badanpangan.go.id/datasetpublications/5up/rata-rata-harga-pangan-bulanan-konsumen-provinsi,
diunduh 31 Oktober 2026. Salinannya: `data/harga_pangan_konsumen_provinsi.csv`.
Batas provinsi untuk peta (gambar yang dibuang): `data/geo/38_provinsi.json`, dari
[ardian28/GeoJson-Indonesia-38-Provinsi](https://github.com/ardian28/GeoJson-Indonesia-38-Provinsi).

## Menjalankan ulang

```bash
pip install -r requirements.txt
jupyter nbconvert --to notebook --execute --inplace analisis.ipynb
```

Keluaran: `docs/index.html` (teks dari `templat_cerita.html`, angka dan gambar dari notebook) dan `docs/img/*.png`.

Brief dibuat dari templat mata kuliah: `python brief/sketsa.py` lalu `python brief/isi_templat.py`
(butuh `Templat_Brief_dan_Storyboard.docx` di folder ini), lalu diekspor ke PDF lewat Word.
