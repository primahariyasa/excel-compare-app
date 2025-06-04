function prosesKomparasi() {
  const templateInput = document.getElementById('templateFile').files[0];
  const referensiInput = document.getElementById('referensiFile').files[0];
  const status = document.getElementById('status');

  if (!templateInput || !referensiInput) {
    alert('Harap upload kedua file: template dan referensi.');
    return;
  }

  status.textContent = '📤 Memproses file...';

  const readerReferensi = new FileReader();
  readerReferensi.onload = function (e) {
    const dataReferensi = new Uint8Array(e.target.result);
    const wbRef = XLSX.read(dataReferensi, { type: 'array' });
    const refSheet = wbRef.Sheets[wbRef.SheetNames[0]];
    const refData = XLSX.utils.sheet_to_json(refSheet);

    const readerTemplate = new FileReader();
    readerTemplate.onload = function (evt) {
      const dataTemplate = new Uint8Array(evt.target.result);
      const wbTemplate = XLSX.read(dataTemplate, { type: 'array' });
      const sheetName = wbTemplate.SheetNames[0];
      const sheet = wbTemplate.Sheets[sheetName];

      // Mulai isi dari baris ke-7
      for (let i = 0; i < refData.length; i++) {
        const row = i + 7;
        const ref = refData[i];

        sheet[`C${row}`] = { t: 's', v: ref.nama_barang || '' };   // C7 = B2 (nama barang)
        sheet[`D${row}`] = { t: 's', v: ref.deskripsi || '' };     // D7 = G2
        sheet[`E${row}`] = { t: 's', v: ref.url_gambar || '' };    // E7 = H2
        sheet[`X${row}`] = { t: 'n', v: Number(ref.harga) || 0 };  // X7 = C2
        sheet[`Y${row}`] = { t: 's', v: ref.warna || '' };         // Y7 = F2
        sheet[`Z${row}`] = { t: 's', v: ref.sku || '' };           // Z7 = A2
      }

      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, sheet, sheetName);
      XLSX.writeFile(newWorkbook, 'hasil_tiktokshop.xlsx');

      status.textContent = '✅ Proses selesai. File berhasil diunduh.';
    };
    readerTemplate.readAsArrayBuffer(templateInput);
  };
  readerReferensi.readAsArrayBuffer(referensiInput);
}
