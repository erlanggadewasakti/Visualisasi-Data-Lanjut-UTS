// Membuat presentasi/Kesenjangan_yang_Menetap.pptx dari PNG di docs/img/slide (hasil analisis.ipynb).
// Jalankan dari folder presentasi:  node buat_slide.js
const path = require("path");
const pptxgen = require("pptxgenjs");

const IMG = path.join(__dirname, "..", "docs", "img", "slide");
const OUT = path.join(__dirname, "Kesenjangan_yang_Menetap.pptx");
// apply_theme.js dari skill pptx; kalau tidak ada, warna tema tetap bawaan Office
const APPLY_THEME = process.env.APPLY_THEME;

const THEME = {
  name: "Kesenjangan yang Menetap",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1A1D1C", lt1: "FFFFFF", dk2: "0D5C54", lt2: "F8F9F7",
    accent1: "1C8A7E", accent2: "0D5C54", accent3: "525A57", accent4: "A3D5C9",
    accent5: "E1F0EC", accent6: "C3C7C4", hlink: "0D5C54", folHlink: "525A57",
  },
};
// Skala heatmap Gambar 3, sama dengan SKALA di notebook
const SKALA = [[0, "EEF7F4"], [0.3, "A3D5C9"], [0.6, "2A8A7E"], [1, "083F39"]];
const SUMBER = "Sumber: Badan Pangan Nasional, Rata-rata Harga Pangan Bulanan Tingkat Konsumen Provinsi, diolah";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13,333 x 7,5 inci
pres.title = "Kesenjangan yang Menetap";
pres.author = "Erlangga Dewa Sakti, Alya Nabila Muliani";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const JUDUL = "+mj-lt"; // huruf judul tema (Cambria) untuk angka besar

// ---------- Layout ----------
const M = 0.6;
const label = (w, warna) => ({ placeholder: { options: { name: "label", type: "body", x: M, y: 0.45, w, h: 0.4, fontSize: 16, bold: true, color: warna, margin: 0, valign: "top" }, text: "" } });
const judul = (w, h, warna) => ({ placeholder: { options: { name: "title", type: "title", x: M, y: 0.85, w, h, fontSize: 34, bold: true, color: warna, margin: 0, valign: "top", align: "left" }, text: "" } });
const nomor = (warna) => ({ x: 12.18, y: 7.0, w: 0.55, h: 0.3, fontSize: 11, color: warna, align: "right" });
const sumber = (w, warna) => ({ text: { text: SUMBER, options: { x: M, y: 7.0, w, h: 0.3, fontSize: 11, color: warna, margin: 0 } } });

pres.defineSlideMaster({ title: "SAMPUL", background: { color: C.text2 }, objects: [] });
pres.defineSlideMaster({ title: "ISI", background: { color: C.background2 },
  objects: [label(12.1, C.accent2), judul(12.1, 1.3, C.text1)], slideNumber: nomor(C.accent3) });
pres.defineSlideMaster({ title: "GRAFIK", background: { color: C.background2 },
  objects: [label(9.0, C.accent2), judul(9.0, 1.0, C.text1), sumber(10.5, C.accent3)], slideNumber: nomor(C.accent3) });
pres.defineSlideMaster({ title: "GRAFIK_KANAN", background: { color: C.background2 },
  objects: [label(6.6, C.accent2), judul(6.6, 1.0, C.text1), sumber(6.6, C.accent3)], slideNumber: nomor(C.accent3) });
pres.defineSlideMaster({ title: "PENUTUP", background: { color: C.text2 },
  objects: [label(12.1, C.accent4), judul(12.1, 1.0, C.background1), sumber(10.5, C.accent4)], slideNumber: nomor(C.accent4) });
pres.defineSlideMaster({ title: "CADANGAN", background: { color: C.background1 },
  objects: [{ text: { text: "Cadangan · tanya jawab", options: { x: M, y: 0.45, w: 8, h: 0.4, fontSize: 16, bold: true, color: C.accent3, margin: 0, valign: "top" } } },
    judul(12.1, 1.3, C.text1)], slideNumber: nomor(C.accent3) });

// ---------- Komponen ----------
function teks(s, isi, o) {
  s.addText(isi, Object.assign({ isTextBox: true, margin: 0, valign: "top", color: C.text1, fontSize: 17 }, o));
}

// Angka besar di kanan atas slide grafik: motif "×" yang sama dengan kolom kiri di web
function angkaKunci(s, angka, ket) {
  teks(s, angka, { objectName: "angka kunci", x: 9.9, y: 0.4, w: 2.83, h: 0.75, fontFace: JUDUL, fontSize: 44, bold: true, color: C.accent2, align: "right" });
  teks(s, ket, { objectName: "keterangan angka kunci", x: 9.9, y: 1.2, w: 2.83, h: 0.6, fontSize: 13, color: C.accent3, align: "right" });
}

function campur(a, b, t) {
  const h = (x) => parseInt(x, 16);
  return [0, 2, 4].map((i) => Math.round(h(a.slice(i, i + 2)) * (1 - t) + h(b.slice(i, i + 2)) * t).toString(16).padStart(2, "0")).join("").toUpperCase();
}
function warnaSkala(t) {
  for (let i = 0; i < SKALA.length - 1; i++) {
    const [p0, c0] = SKALA[i], [p1, c1] = SKALA[i + 1];
    if (t <= p1) return campur(c0, c1, (t - p0) / (p1 - p0));
  }
  return SKALA[SKALA.length - 1][1];
}

// Legenda satu baris; ikon dibuat dari bentuk PowerPoint agar sama dengan legenda web
function legenda(s, x, y, isi) {
  const S = pres.ShapeType;
  for (const [jenis, warna, label] of isi) {
    const cy = y + 0.11;
    let lebar = 0.3;
    if (jenis === "titik") {
      s.addShape(S.ellipse, { x, y: cy - 0.08, w: 0.16, h: 0.16, fill: { color: warna }, line: { color: warna, width: 0 } });
      lebar = 0.16;
    } else if (jenis === "garis") {
      s.addShape(S.line, { x, y: cy, w: 0.32, h: 0, line: { color: warna, width: 1.5, dashType: "sysDot" } });
      lebar = 0.32;
    } else if (jenis === "pita") {
      s.addShape(S.rect, { x, y: cy - 0.08, w: 0.32, h: 0.16, fill: { color: warna, transparency: 85 }, line: { color: warna, width: 0.75, transparency: 55 } });
      lebar = 0.32;
    } else if (jenis === "skala") {
      const n = 8;
      for (let i = 0; i < n; i++) {
        const c = warnaSkala((i + 0.5) / n);
        s.addShape(S.rect, { x: x + i * 0.11, y: cy - 0.08, w: 0.11, h: 0.16, fill: { color: c }, line: { color: c, width: 0 } });
      }
      lebar = 0.88;
    } else if (jenis === "bingkai" || jenis === "bingkai-putus") {
      s.addShape(S.rect, { x, y: cy - 0.08, w: 0.24, h: 0.16, fill: { color: "FFFFFF", transparency: 100 },
        line: { color: warna, width: jenis === "bingkai" ? 2.25 : 1.25, dashType: jenis === "bingkai" ? "solid" : "dash" } });
      lebar = 0.24;
    } else if (jenis === "isi" || jenis === "kosong") {
      s.addShape(S.rect, { x, y: cy - 0.08, w: 0.24, h: 0.16, fill: { color: warna }, line: { color: jenis === "kosong" ? "B6BCB9" : warna, width: 0.75 } });
      lebar = 0.24;
    }
    const w = label.length * 0.078 + 0.1;
    teks(s, label, { x: x + lebar + 0.1, y: y, w, h: 0.25, fontSize: 13, color: C.accent3, valign: "middle" });
    x += lebar + 0.1 + w + 0.3;
  }
}

function gambar(s, nama, x, y, w, h, alt) {
  s.addImage({ path: path.join(IMG, nama), x, y, w, h, altText: alt, objectName: nama });
}

function kartu(s, x, y, w, h, fill, garis) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: garis ? { color: garis, width: 0.75 } : { type: "none" } });
}

function bulatNomor(s, x, y, n, isi, latar) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: 0.55, h: 0.55, fill: { color: latar }, line: { type: "none" } });
  teks(s, String(n), { x, y, w: 0.55, h: 0.55, fontFace: JUDUL, fontSize: 20, bold: true, color: isi, align: "center", valign: "middle" });
}

function tabel(s, baris, x, y, w, kolomW) {
  const kepala = baris[0].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: THEME.colors.dk2 } } }));
  const isi = baris.slice(1).map((r, i) => r.map((t, j) => ({ text: t, options: { bold: j === 0, fill: { color: i % 2 ? "FFFFFF" : THEME.colors.lt2 } } })));
  s.addTable([kepala, ...isi], { x, y, w, colW: kolomW, fontSize: 16, color: C.text1, rowH: 0.5, valign: "middle",
    margin: [0, 0.12, 0, 0.12], border: { type: "solid", color: "DFE3DF", pt: 0.75 } });
}

// ---------- Slide ----------
let s;

pres.addSection({ title: "Pembuka" });
s = pres.addSlide({ masterName: "SAMPUL", sectionTitle: "Pembuka" });
teks(s, "Data story · Ketimpangan Harga Pangan di Tanah Papua, 2021 sampai 2025", { x: M, y: 0.9, w: 8, h: 0.4, fontSize: 16, bold: true, color: C.accent4 });
teks(s, "Kesenjangan yang Menetap", { x: M, y: 1.45, w: 8.0, h: 1.75, fontFace: JUDUL, fontSize: 54, bold: true, color: C.background1 });
teks(s, [
  { text: "Pada 2024, " },
  { text: "lima provinsi termahal seluruhnya berada di Tanah Papua", options: { bold: true, color: C.accent4 } },
  { text: "; Papua dan Papua Barat tetap " },
  { text: "sekitar 19–34% di atas median nasional", options: { bold: true, color: C.accent4 } },
  { text: " setiap tahun sejak 2021." },
], { x: M, y: 3.45, w: 7.4, h: 2.0, fontSize: 26, color: C.background1, lineSpacingMultiple: 1.15 });
teks(s, "5 dari 5", { x: 9.0, y: 1.45, w: 3.7, h: 1.0, fontFace: JUDUL, fontSize: 66, bold: true, color: C.accent4 });
teks(s, "provinsi termahal tahun 2024 ada di Tanah Papua", { x: 9.0, y: 2.5, w: 3.7, h: 0.75, fontSize: 15, color: C.background1 });
teks(s, "19–34%", { x: 9.0, y: 3.55, w: 3.7, h: 1.0, fontFace: JUDUL, fontSize: 66, bold: true, color: C.accent4 });
teks(s, "di atas median nasional, Papua dan Papua Barat setiap tahun 2021 sampai 2025", { x: 9.0, y: 4.6, w: 3.7, h: 0.75, fontSize: 15, color: C.background1 });
teks(s, "Erlangga Dewa Sakti (203012510035) · Alya Nabila Muliani (203012510013)", { x: M, y: 6.25, w: 12, h: 0.35, fontSize: 16, bold: true, color: C.background1 });
teks(s, "Visualisasi Data Lanjut (CAK7AGB4) · Magister Informatika, Telkom University", { x: M, y: 6.65, w: 12, h: 0.35, fontSize: 14, color: C.accent4 });
s.addNotes("Erlangga, ±30 detik.\n- Buka dengan pesan utama, satu kalimat utuh.\n- 5 dari 5: lima provinsi termahal 2024 semuanya di Tanah Papua.\n- 19–34%: rasio tahunan Papua dan Papua Barat 1,19 sampai 1,34× median nasional, 2021 sampai 2025. Hanya dua provinsi ini yang punya data sejak 2021.\n- Data: Badan Pangan Nasional, 38 provinsi, Januari 2021 sampai Januari 2026.");

s = pres.addSlide({ masterName: "ISI", sectionTitle: "Pembuka" });
s.addText("Mengapa audiens ini", { placeholder: "label" });
s.addText("Direktorat Stabilisasi Pasokan dan Harga Pangan, Badan Pangan Nasional", { placeholder: "title" });
[
  ["Pekerjaan mereka", "Memantau dan menganalisis harga pangan untuk stabilisasi harga."],
  ["Yang sudah mereka tahu", "Harga pangan di Papua mahal."],
  ["Yang belum mereka pegang", "Seberapa besar selisihnya, apakah bertahan, dan di mana paling berat."],
].forEach(([h, b], i) => {
  const x = M + i * 4.13;
  kartu(s, x, 2.55, 3.83, 2.75, "FFFFFF", "DFE3DF");
  bulatNomor(s, x + 0.3, 2.85, i + 1, C.background1, THEME.colors.accent1);
  teks(s, h, { x: x + 0.3, y: 3.6, w: 3.25, h: 0.4, fontSize: 19, bold: true, color: C.accent2 });
  teks(s, b, { x: x + 0.3, y: 4.05, w: 3.25, h: 1.1, fontSize: 17 });
});
teks(s, "Karena itu cerita ini membandingkan provinsi dengan provinsi.", { x: M, y: 5.75, w: 12.1, h: 0.4, fontSize: 18, italic: true, color: C.accent3 });
s.addNotes("Erlangga, ±40 detik.\n- Direktorat ini, di bawah Deputi Bidang Ketersediaan dan Stabilisasi Pangan, memantau dan menganalisis harga pangan di tingkat nasional. Cerita kami membandingkan provinsi dengan provinsi, jadi cocok dengan cakupan kerja mereka.\n- Mereka sudah tahu Papua mahal. Kami menjawab tiga hal yang belum mereka pegang: seberapa besar, apakah bertahan, provinsi dan komoditas mana yang paling berat.");

pres.addSection({ title: "Cerita" });
s = pres.addSlide({ masterName: "GRAFIK_KANAN", sectionTitle: "Cerita" });
s.addText("Setup · Gambar 1", { placeholder: "label" });
s.addText("Lima teratas semuanya Papua", { placeholder: "title" });
teks(s, "5 dari 5", { x: M, y: 1.95, w: 6.6, h: 0.95, fontFace: JUDUL, fontSize: 60, bold: true, color: C.accent2 });
teks(s, "provinsi termahal tahun 2024 ada di Tanah Papua", { x: M, y: 2.95, w: 6.6, h: 0.4, fontSize: 16, color: C.accent3 });
teks(s, "Harga tiap provinsi dibagi median provinsi nasional yang tersedia untuk komoditas dan bulan yang sama. 1,0× berarti setara provinsi yang di tengah.", { x: M, y: 3.7, w: 6.6, h: 1.0, fontSize: 18 });
teks(s, [
  { text: "Papua Pegunungan " }, { text: "1,67×", options: { bold: true } },
  { text: " · Papua Tengah " }, { text: "1,48×", options: { bold: true } },
], { x: M, y: 4.9, w: 6.6, h: 0.45, fontSize: 20 });
legenda(s, M, 5.9, [["titik", THEME.colors.accent1, "Provinsi di Tanah Papua"], ["titik", THEME.colors.accent6, "Provinsi lain"]]);
gambar(s, "fig1.png", 7.9, 0.4, 4.8, 6.6, "Dot plot 38 provinsi tahun 2024: lima provinsi termahal semuanya di Tanah Papua, Papua Pegunungan tertinggi 1,67 kali median nasional");
s.addNotes("Erlangga, ±55 detik.\n- Cara membaca: tiap titik satu provinsi, garis putus-putus median nasional 1,0×, makin ke kanan makin mahal.\n- Arahkan mata ke titik hijau toska: lima teratas semuanya Papua.\n- Papua Pegunungan 1,67× = 67% lebih mahal dari provinsi yang di tengah.\n- Kalau ditanya kenapa 2024: tahun lengkap pertama untuk empat provinsi pemekaran.");

s = pres.addSlide({ masterName: "GRAFIK", sectionTitle: "Cerita" });
s.addText("Konflik · Gambar 2", { placeholder: "label" });
s.addText("Lima tahun tetap di atas 1,0×", { placeholder: "title" });
angkaKunci(s, "1,12×", "bulan termurah dalam lima tahun, April 2025");
legenda(s, M, 1.9, [["pita", THEME.colors.accent1, "Rentang kedua garis selama lima tahun"], ["garis", THEME.colors.accent3, "Median nasional = 1,0×"]]);
gambar(s, "fig2.png", 1.15, 2.35, 11.0, 4.4, "Garis rasio harga bulanan Papua dan Papua Barat, Januari 2021 sampai Januari 2026, selalu di atas 1,12 kali median nasional");
s.addNotes("Erlangga, ±55 detik.\n- Kalau mahalnya sementara (cuaca, kapal telat), garis akan turun lagi ke 1,0×. Ternyata tidak.\n- Pita hijau toska: rentang kedua garis selama lima tahun. Tahunan 1,19 sampai 1,34×, rata-rata 27% lebih mahal. Besarnya naik-turun, tapi kesenjangannya tidak pernah hilang.\n- Titik terendah pun 1,12×, April 2025.\n- Garis vertikal Januari 2024: wilayah kedua provinsi menyusut karena pemekaran, jadi sebelum dan sesudahnya bukan wilayah yang persis sama.\n- Serahkan ke Alya.");

s = pres.addSlide({ masterName: "GRAFIK", sectionTitle: "Cerita" });
s.addText("Konflik · Gambar 3", { placeholder: "label" });
s.addText("Kesenjangan terbesar di pegunungan", { placeholder: "title" });
angkaKunci(s, "1,67×", "Papua Pegunungan, provinsi termahal 2024");
legenda(s, M, 1.9, [["skala", "", "1,0× sampai 2,0×: makin gelap makin mahal"], ["bingkai", THEME.colors.dk1, "Dua provinsi terberat"], ["bingkai-putus", THEME.colors.dk1, "Empat komoditas terberat"]]);
gambar(s, "fig3.png", 1.57, 2.35, 10.2, 4.39, "Heatmap enam provinsi Papua dan dua belas komoditas, 2024: Papua Pegunungan dan Papua Tengah paling gelap, terutama cabai, bawang merah, dan minyak goreng");
s.addNotes("Alya, ±55 detik.\n- Kalau dirinci, kesenjangannya tidak merata.\n- Bingkai tebal: Papua Pegunungan dan Papua Tengah jauh di atas provinsi Papua lain.\n- Garis putus: empat komoditas terberat, yaitu cabai rawit, bawang merah, cabai merah, minyak goreng; 1,63 sampai 2,01×.\n- Kontras: beras di Papua Selatan hampir setara harga nasional, 1,05×.\n- Ini menunjukkan di mana telaah lanjutan sebaiknya difokuskan: provinsi mana dan komoditas apa.");

s = pres.addSlide({ masterName: "GRAFIK", sectionTitle: "Cerita" });
s.addText("Resolusi · Gambar 4", { placeholder: "label" });
s.addText("Termahal, tapi datanya paling tipis", { placeholder: "title" });
angkaKunci(s, "62%", "data Papua Pegunungan tahun 2025 kosong");
legenda(s, M, 1.9, [["isi", "8A918E", "Semua 12 komoditas tercatat"], ["isi", "B6BCB9", "Sebagian tercatat"], ["kosong", "FFFFFF", "Tidak ada data"]]);
gambar(s, "fig4.png", M, 2.35, 12.1, 3.54, "Kelengkapan data enam provinsi Papua per bulan sejak Januari 2024: Papua Pegunungan tanpa data sejak Agustus 2025");
teks(s, [
  { text: "86%", options: { bold: true, color: C.accent2 } },
  { text: " harga Papua Pegunungan 2025 yang tercatat berupa kelipatan Rp500, dibanding " },
  { text: "2%", options: { bold: true, color: C.accent2 } },
  { text: " di provinsi lain. Angka sebulat itu menandakan pengamatannya sangat sedikit." },
], { x: M, y: 6.15, w: 12.1, h: 0.7, fontSize: 18 });
s.addNotes("Alya, ±45 detik.\n- Kami ingin melihat apakah 2025 membaik di Papua Pegunungan. Ternyata 62% datanya kosong, dan tidak ada data sama sekali sejak Agustus 2025.\n- Yang tercatat pun 86% kelipatan Rp500, padahal di provinsi lain hanya 2%.\n- Jadi kami tidak menafsirkan 2025 untuk Papua Pegunungan. Angka paling ekstrem harus dibaca bersama kualitas datanya: cakupan pencatatan perlu dievaluasi dan diperkuat sebelum dipakai untuk keputusan operasional. Ini dasar usulan kedua.");

pres.addSection({ title: "Desain" });
s = pres.addSlide({ masterName: "ISI", sectionTitle: "Desain" });
s.addText("Keputusan desain paling menentukan", { placeholder: "label" });
s.addText("Semua angka adalah rasio terhadap median nasional", { placeholder: "title" });
kartu(s, M, 2.2, 5.0, 4.15, THEME.colors.dk2);
teks(s, "Rasio", { x: M + 0.4, y: 2.5, w: 4.2, h: 0.4, fontSize: 18, bold: true, color: C.accent4 });
teks(s, "harga provinsi", { x: M + 0.4, y: 3.15, w: 4.2, h: 0.6, fontFace: JUDUL, fontSize: 30, bold: true, color: C.background1, align: "center" });
teks(s, "÷", { x: M + 0.4, y: 3.75, w: 4.2, h: 0.6, fontFace: JUDUL, fontSize: 30, color: C.accent4, align: "center" });
teks(s, "median provinsi", { x: M + 0.4, y: 4.35, w: 4.2, h: 0.6, fontFace: JUDUL, fontSize: 30, bold: true, color: C.background1, align: "center" });
teks(s, "yang tersedia, untuk komoditas dan bulan yang sama", { x: M + 0.4, y: 5.4, w: 4.2, h: 0.4, fontSize: 15, color: C.accent4, align: "center" });
[
  ["Satu sumbu untuk 12 komoditas", "Beras (±Rp13 ribu) dan daging sapi (±Rp135 ribu) terbaca di skala yang sama."],
  ["Median, bukan rata-rata", "Harga Papua tidak ikut menaikkan pembandingnya. Dengan rata-rata, Papua Pegunungan hanya terbaca 1,58×."],
  ["Dot plot, bukan peta", "Posisi titik pada satu sumbu lebih mudah dibandingkan daripada gradasi warna. Petanya kami buat lalu buang."],
].forEach(([h, b], i) => {
  const y = 2.2 + i * 1.42;
  bulatNomor(s, 6.1, y, i + 1, C.background1, THEME.colors.accent1);
  teks(s, h, { x: 6.9, y: y - 0.02, w: 5.8, h: 0.4, fontSize: 19, bold: true, color: C.accent2 });
  teks(s, b, { x: 6.9, y: y + 0.42, w: 5.8, h: 0.85, fontSize: 16 });
});
s.addNotes("Alya, ±55 detik.\n- Keputusan paling menentukan: semua angka dijadikan rasio terhadap median provinsi nasional yang tersedia (38 provinsi sejak 2024).\n- (1) dua belas komoditas dengan harga berbeda sepuluh kali lipat bisa dibaca di satu sumbu.\n- (2) median, supaya provinsi Papua yang ekstrem tidak menaikkan pembandingnya sendiri.\n- (3) yang dibaca jarak dari 1,0×, jadi dot plot lebih tepat daripada peta. Peta dibuang: Papua yang luas mendominasi, DKI dan DIY nyaris tak terlihat, Papua Barat 1,29× dan Papua 1,25× tampak sama warnanya.");

s = pres.addSlide({ masterName: "ISI", sectionTitle: "Desain" });
s.addText("Satu hal yang berubah selama pengerjaan", { placeholder: "label" });
s.addText("Data 2025 yang rapuh mengubah fokus cerita", { placeholder: "title" });
kartu(s, M, 2.3, 5.55, 3.4, "FFFFFF", "C3C7C4");
teks(s, "Rencana awal", { x: M + 0.4, y: 2.65, w: 4.75, h: 0.35, fontSize: 15, bold: true, color: C.accent3 });
teks(s, "Fokus Papua Pegunungan saja", { x: M + 0.4, y: 3.1, w: 4.75, h: 1.0, fontFace: JUDUL, fontSize: 26, bold: true });
teks(s, "Telur tampak 2,75× median nasional pada 2025.", { x: M + 0.4, y: 4.25, w: 4.75, h: 1.0, fontSize: 17, color: C.accent3 });
s.addShape(pres.ShapeType.rightArrow, { x: 6.33, y: 3.7, w: 0.67, h: 0.6, fill: { color: THEME.colors.accent1 }, line: { type: "none" } });
kartu(s, 7.15, 2.3, 5.55, 3.4, THEME.colors.accent5);
teks(s, "Yang kami lakukan", { x: 7.55, y: 2.65, w: 4.75, h: 0.35, fontSize: 15, bold: true, color: C.accent2 });
teks(s, "Diperluas ke enam provinsi Papua", { x: 7.55, y: 3.1, w: 4.75, h: 1.0, fontFace: JUDUL, fontSize: 26, bold: true });
teks(s, "Tahun pembanding diganti ke 2024 yang lengkap. Data 2025 yang rapuh dijadikan temuan (Gambar 4).", { x: 7.55, y: 4.25, w: 4.75, h: 1.2, fontSize: 17 });
s.addNotes("Erlangga, ±40 detik.\n- Awalnya kami ingin fokus ke Papua Pegunungan saja karena angka telurnya mencolok, 2,75× median pada 2025.\n- Dicek per bulan: telur melompat dari Rp117.500 ke Rp47.500 dalam sebulan, dan hampir semua harga 2025 berupa angka bulat.\n- Jadi fokus diperluas ke enam provinsi Papua, tahun pembanding 2024. Kerapuhan datanya kami tampilkan sebagai Gambar 4.");

pres.addSection({ title: "Usulan" });
s = pres.addSlide({ masterName: "PENUTUP", sectionTitle: "Usulan" });
s.addText("Yang kami usulkan", { placeholder: "label" });
s.addText("Dua usulan untuk pemantauan dan evaluasi", { placeholder: "title" });
[
  [{ text: "Jadikan Papua Pegunungan dan Papua Tengah prioritas telaah lebih lanjut", options: { bold: true } },
   { text: ", khususnya bawang, cabai, dan minyak goreng, dalam pemantauan dan evaluasi program stabilisasi harga." }],
  [{ text: "Evaluasi dan perkuat cakupan pencatatan harga di Papua Pegunungan", options: { bold: true } },
   { text: ": 62% data 2025 kosong, jadi kualitas datanya perlu diperbaiki sebelum dipakai untuk keputusan operasional." }],
].forEach((isi, i) => {
  const y = 2.3 + i * 1.65;
  teks(s, String(i + 1), { x: M, y, w: 0.8, h: 1.0, fontFace: JUDUL, fontSize: 60, bold: true, color: C.accent4 });
  teks(s, isi, { x: 1.6, y: y + 0.1, w: 11.1, h: 1.3, fontSize: 24, color: C.background1 });
});
teks(s, "Pada 2024, lima provinsi termahal seluruhnya berada di Tanah Papua; Papua dan Papua Barat tetap sekitar 19–34% di atas median nasional setiap tahun sejak 2021.",
  { x: M, y: 5.95, w: 12.1, h: 0.75, fontSize: 17, italic: true, color: C.accent4 });
s.addNotes("Alya, ±35 detik, lalu tanya jawab.\n- Usulan 1: Papua Pegunungan dan Papua Tengah paling tinggi, terutama empat komoditas terberat, jadi layak jadi prioritas telaah. Data harga saja belum cukup untuk menentukan bentuk intervensinya.\n- Usulan 2: 62% data Papua Pegunungan 2025 kosong. Cakupan pencatatan perlu dievaluasi dan diperkuat sebelum datanya dipakai untuk keputusan operasional yang lebih spesifik.\n- Tutup dengan membaca ulang pesan utama di bawah, lalu: \"Terima kasih, kami siap menjawab pertanyaan.\"");

pres.addSection({ title: "Cadangan" });
s = pres.addSlide({ masterName: "CADANGAN", sectionTitle: "Cadangan" });
s.addText("Kenapa median, bukan rata-rata", { placeholder: "title" });
tabel(s, [
  ["Provinsi (2024)", "Terhadap median", "Terhadap rata-rata"],
  ["Papua Pegunungan", "1,67×", "1,58×"],
  ["Papua Tengah", "1,48×", "1,40×"],
  ["Papua Barat", "1,29×", "1,22×"],
  ["Papua", "1,25×", "1,18×"],
  ["Papua Selatan", "1,17×", "1,11×"],
], M, 2.15, 7.4, [3.2, 2.1, 2.1]);
teks(s, "2 sampai 7%", { x: 8.6, y: 2.1, w: 4.1, h: 0.85, fontFace: JUDUL, fontSize: 40, bold: true, color: C.accent2 });
teks(s, "rata-rata nasional lebih tinggi dari median di setiap komoditas (2024), karena harga Papua menariknya ke atas.", { x: 8.6, y: 2.95, w: 4.1, h: 1.3, fontSize: 16, color: C.accent3 });
teks(s, "Dengan rata-rata sebagai pembanding, selisih Papua terlihat lebih kecil dari kenyataannya.", { x: 8.6, y: 4.45, w: 4.1, h: 1.2, fontSize: 17, bold: true });
s.addNotes("Jawaban singkat: harga antarprovinsi condong ke kanan karena Papua. Kalau pembandingnya rata-rata, Papua ikut menaikkan pembandingnya sendiri dan selisihnya tampak lebih kecil. Median tidak terpengaruh ekor itu, dan 1,0× langsung berarti provinsi yang di tengah. Kelemahan yang perlu diakui: median tidak ditimbang jumlah penduduk.");

s = pres.addSlide({ masterName: "CADANGAN", sectionTitle: "Cadangan" });
s.addText("Apakah 1,67× bergantung pada data yang meragukan?", { placeholder: "title" });
[
  ["1,67×", "Setahun penuh 2024, angka yang dipakai di cerita"],
  ["1,64×", "Hanya Januari sampai Juni 2024, saat porsi angka bulat 0 sampai 10%"],
  ["1,69×", "Semua harga kelipatan Rp500 dibuang"],
].forEach(([a, k], i) => {
  const x = M + i * 4.13;
  kartu(s, x, 2.2, 3.83, 2.5, i === 0 ? THEME.colors.accent5 : THEME.colors.lt2, "DFE3DF");
  teks(s, a, { x: x + 0.35, y: 2.45, w: 3.1, h: 0.95, fontFace: JUDUL, fontSize: 48, bold: true, color: C.accent2 });
  teks(s, k, { x: x + 0.35, y: 3.5, w: 3.1, h: 1.0, fontSize: 16, color: C.accent3 });
});
teks(s, "Dengan ketiga cara, Papua Pegunungan tetap termahal dan lima besar tetap semuanya di Tanah Papua. Porsi angka bulat di Papua Pegunungan naik bertahap: 0% di awal 2024, 100% sejak Mei 2025.",
  { x: M, y: 5.1, w: 12.1, h: 1.0, fontSize: 18 });
s.addNotes("Kalau ditanya soal kualitas data 2024: penurunan kualitas mulai pertengahan 2024, tapi kesimpulan tidak bergantung pada bulan yang meragukan. Hindari kata \"malas\": data hanya menunjukkan angka makin bulat dan makin jarang; penyebabnya bisa jumlah pasar yang dipantau, pergantian petugas, atau akses.");

s = pres.addSlide({ masterName: "CADANGAN", sectionTitle: "Cadangan" });
s.addText("Kenapa data kosong tidak diinterpolasi", { placeholder: "title" });
[
  ["Setelah Juli 2025 tidak ada titik ujung", "Mengisinya berarti ekstrapolasi, menebak harga yang tidak pernah diamati."],
  ["Titik yang tercatat pun lemah", "86% harga Papua Pegunungan 2025 berupa kelipatan Rp500."],
  ["Kekosongan itu sendiri temuannya", "Provinsi termahal paling jarang dipantau; itu dasar usulan kedua."],
  ["Klaim utama tidak memakainya", "Gambar 1 dan 3 memakai 2024 yang hampir lengkap. Celah Papua Selatan (Januari sampai Maret 2025) jatuh di Ramadan dan Lebaran, saat harga biasanya melonjak."],
].forEach(([h, b], i) => {
  const x = i % 2 ? 6.8 : M, y = 2.15 + Math.floor(i / 2) * 2.1;
  bulatNomor(s, x, y, i + 1, C.background1, THEME.colors.accent1);
  teks(s, h, { x: x + 0.8, y: y - 0.02, w: 5.1, h: 0.4, fontSize: 19, bold: true, color: C.accent2 });
  teks(s, b, { x: x + 0.8, y: y + 0.45, w: 5.1, h: 1.3, fontSize: 16 });
});
s.addNotes("Jawaban singkat: data setelah Juli 2025 tidak ada sama sekali, jadi mengisinya berarti mengarang. Data yang ada pun kebanyakan angka bulat. Kekosongan ini justru temuan kami, dan klaim utama memakai data 2024 yang hampir lengkap.");

s = pres.addSlide({ masterName: "CADANGAN", sectionTitle: "Cadangan" });
s.addText("Kenapa Gambar 2 hanya Papua dan Papua Barat", { placeholder: "title" });
tabel(s, [
  ["Provinsi", "Data mulai"],
  ["Papua", "Januari 2021"],
  ["Papua Barat", "Januari 2021"],
  ["Papua Tengah", "Januari 2024"],
  ["Papua Pegunungan", "Januari 2024 (berhenti Juli 2025)"],
  ["Papua Selatan", "Januari 2024"],
  ["Papua Barat Daya", "Januari 2024"],
], M, 2.15, 6.9, [2.8, 4.1]);
teks(s, "Empat provinsi hasil pemekaran 2022 baru dicatat terpisah mulai Januari 2024. Klaim \"lima tahun\" hanya bisa diuji dengan provinsi yang tercatat sejak 2021.",
  { x: 8.1, y: 2.15, w: 4.6, h: 1.9, fontSize: 18 });
teks(s, "Sebelum 2024, garis Papua masih mencakup Pegunungan, Tengah, dan Selatan. Turunnya garis setelah 2024 sebagian karena wilayahnya menyusut.",
  { x: 8.1, y: 4.2, w: 4.6, h: 1.7, fontSize: 16, color: C.accent3 });
s.addNotes("Keempat provinsi lain tetap muncul di Gambar 1, 3, dan 4. Kalau ditanya: mereka bisa ditambahkan sebagai garis pendek mulai 2024, tapi judul \"lima tahun\" hanya berlaku untuk dua garis.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  if (APPLY_THEME) {
    const { applyTheme } = require(APPLY_THEME);
    await applyTheme(OUT, THEME);
  }
  console.log("Tersimpan:", OUT);
})();
