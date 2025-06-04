let workbookA, workbookB;

document.getElementById('fileA').addEventListener('change', handleFileA);
document.getElementById('fileB').addEventListener('change', handleFileB);
document.getElementById('processBtn').addEventListener('click', processFiles);

function handleFileA(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  reader.onload = (e) => {
    workbookA = XLSX.read(e.target.result, { type: 'binary' });
  };
  reader.readAsBinaryString(file);
}

function handleFileB(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  reader.onload = (e) => {
    workbookB = XLSX.read(e.target.result, { type: 'binary' });
  };
  reader.readAsBinaryString(file);
}

function processFiles() {
  if (!workbookA || !workbookB) {
    alert("Pastikan kedua file sudah di-upload.");
    return;
  }

  const sheetA = workbookA.Sheets[workbookA.SheetNames[0]];
  const sheetB = workbookB.Sheets[workbookB.SheetNames[0]];

  const dataA = XLSX.utils.sheet_to_json(sheetA);
  const dataB = XLSX.utils.sheet_to_json(sheetB);

  // Contoh: cocokkan berdasarkan "ID", dan ambil "Nama" dari Excel A ke Excel B
  const updatedDataB = dataB.map(rowB => {
    const match = dataA.find(rowA => rowA.ID === rowB.ID);
    return {
      ...rowB,
      Nama: match ? match.Nama : rowB.Nama || ""
    };
  });

  const newSheet = XLSX.utils.json_to_sheet(updatedDataB);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Hasil");

  const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'binary' });

  // Convert to Blob for download
  const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
  const link = document.getElementById('downloadLink');
  link.href = URL.createObjectURL(blob);
  link.download = "hasil.xlsx";
  link.style.display = 'inline';
  link.textContent = 'Download Hasil';
}

function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}
