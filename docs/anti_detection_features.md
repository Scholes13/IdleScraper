# Fitur Anti-Deteksi Werkudara Scraper

Dokumentasi ini menjelaskan fitur-fitur anti-deteksi yang diimplementasikan dalam Werkudara Scraper untuk meningkatkan ketahanan scraping dan mencegah pemblokiran dari Google Maps.

## 1. Rotasi User-Agent

Fitur ini secara otomatis merotasi user-agent yang digunakan dalam permintaan scraping untuk meniru berbagai browser dan perangkat.

### Implementasi:
- Daftar 12+ user-agent modern (desktop & mobile) tersedia
- Rotasi otomatis pada setiap permintaan baru
- Rotasi tambahan saat terjadi error

### Konfigurasi:
- Aktifkan/nonaktifkan di UI dengan opsi "Use rotating User-Agents"
- Default: Aktif

### Keuntungan:
- Menyamarkan permintaan sebagai pengguna normal, bukan bot
- Mencegah deteksi pola permintaan dari user-agent yang sama
- Meningkatkan keberhasilan scraping di sesi panjang

## 2. Delay Dinamis

Sistem delay adaptif yang secara otomatis menyesuaikan waktu tunggu berdasarkan respons dan pola error.

### Implementasi:
- Sistem delay dasar dengan jitter acak
- Peningkatan otomatis delay saat terjadi error (backoff eksponensial)
- Pengurangan gradual delay saat permintaan sukses
- Pelacakan waktu respons

### Parameter Konfigurasi:
- Base delay: 2 detik (default)
- Max delay: 15 detik
- Delay factor: Menyesuaikan secara otomatis (1.0 - 5.0)

### Keuntungan:
- Meniru perilaku pengguna manusia
- Menghindari pola delay yang konsisten/mudah dideteksi
- Mengurangi risiko throttling/pembatasan permintaan
- Secara otomatis beradaptasi dengan kebijakan throttling Google

## 3. Rotasi Proxy (Eksperimental)

Menggunakan daftar proxy gratis yang diperbarui secara otomatis untuk merotasi alamat IP.

### Implementasi:
- Pengambilan otomatis daftar proxy gratis dari sumber publik
- Pengujian koneksi proxy untuk memverifikasi ketersediaan
- Rotasi otomatis saat terjadi kegagalan
- Caching daftar proxy valid

### Konfigurasi:
- Aktifkan/nonaktifkan di UI dengan opsi "Use free proxy rotation (experimental)"
- Default: Nonaktif

### Keuntungan:
- Memvariasikan alamat IP untuk menghindari pembatasan per IP
- Meningkatkan anonimitas dan ketahanan scraping
- Mengurangi risiko pemblokiran pada alamat IP tetap
- Bekerja dengan proxy gratis tanpa biaya tambahan

### Keterbatasan:
- Proxy gratis dapat tidak stabil/lambat
- Ketersediaan proxy berkualitas terbatas
- Dapat memperlambat proses scraping

## Praktik Terbaik

1. **Kombinasi Fitur:**
   - Untuk scraping skala kecil (< 50 perusahaan): Gunakan rotasi user-agent dan delay dinamis
   - Untuk scraping skala besar (> 50 perusahaan): Aktifkan semua fitur anti-deteksi

2. **Konfigurasi yang Disarankan:**
   - Scraping umum: Rotasi user-agent aktif, delay dinamis aktif
   - Saat menghadapi pembatasan/throttling: Aktifkan rotasi proxy

3. **Pertimbangan Kinerja:**
   - Rotasi proxy dapat memperlambat proses scraping
   - Delay dinamis meningkat saat terjadi error, namun juga meningkatkan keberhasilan jangka panjang

## Tabel Perbandingan

| Fitur | Kecepatan | Ketahanan | Kompleksitas | Default |
|-------|-----------|-----------|--------------|---------|
| Rotasi User-Agent | Cepat | Sedang | Rendah | Aktif |
| Delay Dinamis | Sedang | Tinggi | Sedang | Aktif |
| Rotasi Proxy | Lambat | Sangat Tinggi | Tinggi | Nonaktif |

## Troubleshooting

**Jika Anda mengalami pemblokiran:**

1. Pastikan rotasi user-agent aktif
2. Tingkatkan nilai base delay (perubahan di kode)
3. Aktifkan rotasi proxy untuk sesi berikutnya
4. Gunakan caching untuk mengurangi jumlah permintaan 