Cara deploy EXAMUNA ke Google Apps Script

Prasyarat
- Node.js + npm (opsional untuk `clasp`)
- akun Google yang punya akses ke Script ID yang diberikan

1) Pasang clasp (jika belum):
```bash
npm install -g @google/clasp
```

2) Login ke akun Google:
```bash
clasp login
```

3) Pilihan: clone remote atau kaitkan lokal ke Script ID
- Clone remote (ambil kode dari GAS):
```bash
clasp clone 1li0ycYSk1Y2d19bMfx9KUqJrUnUDszoPU3iCi1T9IhY
```
- Atau jika Anda sudah berada di folder project lokal (`f:\EXAMUNA`), cukup pastikan `.clasp.json` berisi `scriptId` (sudah dibuat) lalu:
```bash
clasp push
```

4) Buat deployment (web app)
- Menggunakan Apps Script Editor (direkomendasikan):
  - Buka https://script.google.com dan buka project terkait
  - Menu Deploy → New deployment → Pilih "Web app"
  - Atur "Execute as" = "Me (user deploying)" dan "Who has access" sesuai kebutuhan (mis. Anyone)
  - Klik Deploy → salin Deployment URL

- Atau via CLI:
```bash
clasp deploy --description "Initial deploy"
```

5) Update `index.html` jika perlu
- Jika ada placeholder untuk `SCRIPT_URL` atau endpoint, ganti dengan Deployment URL yang didapat.

Catatan
- `.clasp.json` sudah dibuat di repository dengan `scriptId` yang Anda berikan.
- Saya tidak akan mem-push perubahan ke Apps Script tanpa izin Anda; jika mau saya bantu menjalankan perintah `clasp push`/`clasp deploy`, beri tahu saya.
