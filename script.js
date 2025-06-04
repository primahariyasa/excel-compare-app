let dataReferensi = [];

async function loadReferensi() {
    try {
        const response = await fetch('https://yourdomain.com/export.php'); // ganti dengan URL Netlify/backend kamu
        if (!response.ok) throw new Error('Gagal fetch data referensi');
        dataReferensi = await response.json();
        console.log('✅ Data referensi berhasil dimuat:', dataReferensi);
    } catch (err) {
        console.error(err);
    }
}

document.getElementById('inputExcel').addEventListener('change', handleFile, false);

async function handleFile(e) {
    await loadReferensi(); // pastikan data referensi sudah dimuat

    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const newSheet = { ...sheet };

        // mulai dari baris ke-7, karena C7 dst
        for (let i = 0; i < dataReferensi.length; i++) {
            const row = i + 7; // Excel row (1-based)
            const ref = dataReferensi[i];

            // Mapping sesuai permintaan
            newSheet[`C${row}`] = { t: 's', v: ref.nama_barang };   // C7 → B2 (nama)
            newSheet[`D${row}`] = { t: 's', v: ref.deskripsi };     // D7 → G2 (deskripsi)
            newSheet[`E${row}`] = { t: 's', v: ref.url_gambar };    // E7 → H2 (url_gambar)
            newSheet[`X${row}`] = { t: 'n', v: Number(ref.harga) }; // X7 → C2 (harga)
            newSheet[`Y${row}`] = { t: 's', v: ref.warna };         // Y7 → F2 (warna)
            newSheet[`Z${row}`] = { t: 's', v: ref.sku };           // Z7 → A2 (SKU)
        }

        // Buat file baru dan download
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
        XLSX.writeFile(newWorkbook, 'hasil_tiktokshop.xlsx');
    };
    reader.readAsArrayBuffer(file);
}
