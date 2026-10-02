# HPS Progress Dashboard — GitHub Pages

Versi ini **bukan Next.js**. Ini adalah website HTML/CSS/JavaScript statis sehingga bisa langsung dipasang dengan:

**GitHub → Settings → Pages → Deploy from a branch → `main` → `/ (root)` → Save**

## Struktur

- `index.html` — halaman utama GitHub Pages
- `app.js` — dashboard dan pemrosesan Excel di browser
- `style.css` — tampilan
- `data/db.json` — data awal

## Penting

Jangan hapus `index.html` dari root repository. GitHub Pages akan membuka file itu.

Upload Excel diproses di browser dan disimpan di `localStorage` browser tersebut. Upload tidak mengubah file `data/db.json` di GitHub dan tidak otomatis mengubah data di perangkat lain.

Jika hanya ingin memperbarui data bawaan untuk semua pengguna, ganti `data/db.json` di repository dengan file JSON hasil import yang sesuai.
