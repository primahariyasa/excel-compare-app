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
    const refData = XLSX.utils.sheet_to_json(refSheet, { defval: "" });

    const readerTemplate = new FileReader();
    readerTemplate.onload = function (evt) {
      const dataTemplate = new Uint8Array(evt.target.result);
      const wbTemplate = XLSX.read(dataTemplate, { type: 'array', cellStyles: true });

      const sheetName = wbTemplate.SheetNames[0]; // Nama sheet asli dipertahankan
      const sheet = wbTemplate.Sheets[sheetName];

      // Loop untuk isi hanya kolom yang diminta
      for (let i = 0; i < refData.length; i++) {
        const row = i + 7; // mulai dari baris 7
        const ref = refData[i];

        // Hanya overwrite kolom target
        if (ref.nama_barang) sheet[`C${row}`] = { ...sheet[`C${row}`], t: 's', v: ref.nama_barang };
        if (ref.deskripsi) sheet[`D${row}`] = { ...sheet[`D${row}`], t: 's', v: ref.deskripsi };
        if (ref.url_gambar) sheet[`E${row}`] = { ...sheet[`E${row}`], t: 's', v: ref.url_gambar };
        if (ref.harga) sheet[`X${row}`] = { ...sheet[`X${row}`], t: 'n', v: Number(ref.harga) };
        if (ref.warna) sheet[`Y${row}`] = { ...sheet[`Y${row}`], t: 's', v: ref.warna };
        if (ref.sku) sheet[`Z${row}`] = { ...sheet[`Z${row}`], t: 's', v: ref.sku };
      }

      // Ekspor workbook, tidak ubah sheet name atau format apapun
      XLSX.writeFile(wbTemplate, 'hasil_tiktokshop.xlsx', {
        bookType: 'xlsx',
        compression: true,
        cellStyles: true,
      });

      status.textContent = '✅ Proses selesai. Format, rumus, dan nama file tetap utuh.';
    };

    readerTemplate.readAsArrayBuffer(templateInput);
  };

  readerReferensi.readAsArrayBuffer(referensiInput);
}
