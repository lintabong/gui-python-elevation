# MeasurementApp
MeasurementApp adalah aplikasi berbasis GUI yang dirancang untuk memproses data pengukuran pasang surut air dari file Excel, menampilkan hasil dalam bentuk tabel dan grafik, serta menyimpan hasil analisis ke file Excel.

## Fitur Utama:
Memuat file Excel untuk diproses.
Menampilkan hasil pengolahan data, termasuk tinggi air maksimum, minimum, tunggang pasut, formzahl, dan tipe pasut.
Menampilkan grafik pasang surut air.
Menyimpan hasil pengolahan dalam format Excel.
Fitur dropdown untuk memilih dan menampilkan plot per hari.
Teknologi yang Digunakan:
Python untuk pemrosesan data dan pengembangan GUI.
Tkinter untuk antarmuka pengguna.
Matplotlib untuk menampilkan grafik data.
Pandas untuk manipulasi data.
Openpyxl untuk membaca dan menulis file Excel.

## Cara Instalasi:
### Prasyarat:
- Python 3.x harus terinstal di komputer Anda. Anda dapat mengunduhnya di python.org.
- File requirement.txt tersedia untuk menginstall semua dependencies yang dibutuhkan.


### Langkah-langkah Instalasi:
1. Buat virtual environment (opsional, tapi direkomendasikan)
```sh
python -m venv venv
```
2. Aktifkan virtual environtment
```sh
venv\Scripts\activate
```
3. Install dependencies yang diperlukan
```sh
pip install -r requirements.txt
```
4. Jalankan Aplikasi
```sh
python main.py
```

5. Export Aplikasi ke *.exe:
```sh
pyinstaller --onefile --windowed main.py
```