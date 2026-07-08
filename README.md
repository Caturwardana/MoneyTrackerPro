# 💰 Money Tracker Pro

> Web-based personal finance tracker built with **Google Apps Script** and **Google Sheets**.

Money Tracker Pro adalah aplikasi pencatatan keuangan harian berbasis web yang berjalan sepenuhnya di ekosistem Google.
Aplikasi ini dirancang sebagai solusi manajemen keuangan yang:

* ✅ Gratis
* 🔒 Privat
* ⚡ Cepat
* 📱 Mobile-first
* ☁️ Tanpa server eksternal

Database tersimpan langsung di Google Drive milik pengguna.

---

## ✨ Features

### 🔐 Session-Based Auto Login
Login sekali, session tersimpan 24 jam. Buka-tutup aplikasi tanpa perlu masukkan PIN lagi.

### ⚡ Quick Transaction Input
Catat pemasukan, pengeluaran, dan transfer antar wallet dengan UI responsif.

### 💳 Multi Wallet Support
Kelola saldo dari berbagai sumber (Tunai, Bank, E-Wallet).

### 👉 Slide-to-Action UI
Gesture geser untuk edit atau hapus transaksi seperti aplikasi mobile native.

### 📊 Real-time Dashboard
Visualisasi pengeluaran menggunakan interactive doughnut chart, plus perbandingan Pemasukan vs Pengeluaran (VS mode).

### 📈 Monthly Recap Chart
Rekap pemasukan & pengeluaran per bulan dalam bentuk bar chart. Navigasi antar tahun dengan tombol ◀ ▶.

### 🔍 Advanced Filtering
Filter riwayat berdasarkan periode:

* Mingguan
* Bulanan
* Tahunan
* Custom date range

### 🔎 Live Search
Cari transaksi real-time berdasarkan kategori, catatan, atau nominal.

### 🧑‍🤝‍🧑 Multi-User Support
Admin dapat menambahkan akun pengguna lain. Setiap pengguna punya data terpisah.

### 🗑️ Smart Delete
Tombol hapus wallet/kategori hanya muncul jika item tersebut belum pernah dipakai di transaksi, mencegah kehilangan data.

### 🌙 Dark / Light Theme
Toggle tema gelap-terang yang tersimpan di localStorage.

### 📥 CSV Export
Ekspor laporan keuangan kompatibel dengan Microsoft Excel.

### 🔒 100% Private
Data tersimpan di Google Drive pengguna tanpa server pihak ketiga.

---

## 🏗 Architecture

Aplikasi menggunakan **N-Tier Thin Client Architecture** dengan **Session-Based Authentication**.

### Tech Stack

| Component            | Technology                             |
| -------------------- | -------------------------------------- |
| Backend Runtime      | Google Apps Script (V8 Engine)         |
| Database             | Google Sheets API                      |
| Frontend             | Bootstrap 5                            |
| Chart Engine         | Chart.js                               |
| Typography           | Plus Jakarta Sans                      |
| Interaction          | Native JavaScript Touch Events         |
| Theme                | CSS Custom Properties (Dark/Light)     |
| Session Cache        | Google Apps Script CacheService (24h)  |
| Frontend Cache       | In-memory dictionary with TTL          |
| Security             | Session Token + Input Sanitization     |

### Security Model

```
Frontend                          Server (GAS)
─────────────────────────────────────────────────
Login (PIN) ──────────────────► loginSession()
                                    │
◄──── { sessionToken, userData } ───┘
         │
Simpan token di localStorage
         │
Setiap request:
  payload._token = sessionToken
         │
         ▼
getTransactions() ────┬──► _validateAndGetUserId()
                      │       │
                      │   Cek CacheService
                      │   Token valid? → userId
                      │   Kadaluarsa? → throw Error
                      │
                      └──► _getTransactions()
                           (hanya untuk userId terverifikasi)
```

Fungsi internal (`_ensureAppSetup`, `_resolveUserId`, dll) menggunakan prefix `_` sehingga tidak bisa dipanggil langsung dari `google.script.run` via browser console.

---

## 📁 Project Structure

```
/
├── backend/
│   └── Code.gs           # Server-side logic (CRUD, auth, session)
├── frontend/
│   └── Index.html        # SPA UI (Bootstrap 5 + Chart.js)
├── index.html            # Root deployment (redirect + placeholder DEPLOYMENT_ID)
├── screenshots/          # Application preview images
└── README.md
```

### Server Functions (Code.gs)

| Kategori    | Public (dipanggil frontend)       | Internal (prefix _)                  |
|------------|-----------------------------------|--------------------------------------|
| Auth       | `loginSession`, `verifyAutoLogin`, `destroySession` | `_validateAndGetUserId`, `_loginInit`, `_resolveUserId` |
| CRUD       | `simpanTransaksi`, `hapusTransaksi` | `_validateTransaksi`, `_updateExistingTransaction` |
| Wallet     | `getWalletBalances`, `saveWallet`, `editWallet`, `deleteWallet` | `_updateBalance`, `_revertBalance` |
| Kategori   | `getCategories`, `saveCategory`, `editCategory`, `deleteCategory` | — |
| Data       | `getMasterData`, `getTransactions`, `getDashboardData`, `getMonthlyRecap` | `_getUsageInfo` |
| Export     | `getExportData`                    | `_escapeCsvField` |
| Utility    | `doGet`, `include`                 | `_ensureAppSetup`, `_getSheetValues`, `_getUsers`, `_generateId` dll |

---

## ⚙️ Installation Guide

Tidak perlu instal software tambahan.
Cukup gunakan akun Google.

---

### 1️⃣ Setup Database (Google Sheets)

Buat spreadsheet baru dengan struktur berikut.

#### 📄 Sheet: `Transaksi`

| Tanggal | Tipe | Kategori | Wallet | Nominal | Catatan | TRX ID | User ID |
|---|---|---|---|---|---|---|---|

---

#### 📄 Sheet: `Wallets`

| Nama Wallet | Saldo | User ID |
|---|---|---|

---

#### 📄 Sheet: `Users`

| ID | Nama | PIN | Created At | isAdmin |
|---|---|---|---|---|

> 💡 PIN default untuk admin pertama: `1234`

---

#### 📄 Sheet: `Kategori`

| Nama | Tipe | User ID |
|---|---|---|

---

### 2️⃣ Setup Google Apps Script

1. Google Sheets → **Extensions → Apps Script**
2. Hapus kode default:

```
function myFunction() {}
```

3. Buat file baru bernama `Code.gs`, copy isi:
```
backend/Code.gs
```

4. Buat file HTML bernama `Index`, copy isi:
```
frontend/Index.html
```

5. (Opsional) Buat file `Stylesheet.html` untuk kustomisasi CSS.

6. Save project.

---

### 3️⃣ Deploy Web App

1. Deploy → New Deployment
2. Pilih Web App
3. Configure:

* Execute as → Me
* Who has access → sesuai kebutuhan

4. Deploy
5. Gunakan URL Web App yang diberikan.

---

## 🔐 Security & Privacy

* **Session Token** — Setiap request dari frontend diverifikasi dengan token unik (24 jam TTL)
* **Internal Function Protection** — Fungsi internal (`_*`) tidak bisa dipanggil via `google.script.run`
* **XSS Protection** — Semua output HTML melalui fungsi `escapeHtml()`
* **Input Validation** — Setiap input dari frontend divalidasi di server sebelum diproses
* **No External Server** — Tidak ada data dikirim ke pihak ketiga
* **Data Ownership** — Data tersimpan di Google Drive pengguna sendiri
* **Auditable** — Source code dapat diaudit kapan saja

---

## ⭐ Support

Jika project ini membantu, jangan lupa ⭐ repository ini.

---
