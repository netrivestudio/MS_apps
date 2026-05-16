// =====================================
// KONSTANTA BISNIS
// =====================================
const MODAL_PER_GALON = 4500;
const BIAYA_OPERASIONAL_PER_GALON = 500;

// =====================================
// LOAD DATA DARI LOCAL STORAGE
// =====================================
let data = JSON.parse(localStorage.getItem("minumoData")) || [];

// =====================================
// HELPER AMAN SET TEXT
// =====================================
function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

// =====================================
// SIMPAN DATA
// =====================================
function simpanData() {
  localStorage.setItem("minumoData", JSON.stringify(data));
}

// =====================================
// TAMBAH DATA
// =====================================
function tambahData() {
  const nama = document.getElementById("namaPelanggan").value.trim();
  const tanggal = document.getElementById("tanggal").value;
  const jenis = document.getElementById("jenis").value;
  const jumlah = parseInt(document.getElementById("jumlah").value) || 0;
  const harga = parseInt(document.getElementById("harga").value) || 0;

  // 🔥 INFAQ MANUAL (NOMINAL BEBAS)
  const infaqNominal =
    parseInt(document.getElementById("infaq").value) || 0;

  if (!nama || !tanggal || jumlah <= 0 || harga <= 0) {
    alert("Isi semua data dengan benar!");
    return;
  }

  // ✅ Infaq HANYA dari input, tidak dikali galon
  const potonganInfaq = jenis === "Pemasukan" ? infaqNominal : 0;

  data.push({
    nama,
    tanggal,
    jenis,
    jumlah,
    harga,
    potongan: potonganInfaq
  });

  simpanData();
  renderTable();
  updateInfo();

  // reset input
  document.getElementById("namaPelanggan").value = "";
  document.getElementById("jumlah").value = "";
  document.getElementById("harga").value = "";
  document.getElementById("infaq").value = "";
}

// =====================================
// RENDER TABEL
// =====================================
function renderTable() {
  const tbody = document.querySelector("#dataTable tbody");
  if (!tbody) return;

  tbody.innerHTML = "";

  data.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${item.nama}</td>
      <td>${item.tanggal}</td>
      <td>${item.jenis}</td>
      <td>${item.jumlah}</td>
      <td>Rp ${item.harga.toLocaleString("id-ID")}</td>
      <td>Rp ${item.potongan.toLocaleString("id-ID")}</td>
      <td><button onclick="hapusBaris(${index})">Hapus</button></td>
    `;
    tbody.appendChild(tr);
  });
}

// =====================================
// HAPUS BARIS
// =====================================
function hapusBaris(index) {
  data.splice(index, 1);
  simpanData();
  renderTable();
  updateInfo();
}

// =====================================
// HAPUS SEMUA
// =====================================
function hapusSemua() {
  if (!confirm("Yakin mau hapus semua data?")) return;
  data = [];
  simpanData();
  renderTable();
  updateInfo();
}

// =====================================
// HITUNG TOTAL & PROFIT
// =====================================
function updateInfo() {
  let totalPenjualan = 0;
  let totalPengeluaran = 0;
  let totalInfaq = 0;
  let totalGalon = 0;
  let totalBiayaOperasional = 0;

  data.forEach(item => {
    const omzet = item.jumlah * item.harga;

    if (item.jenis === "Pemasukan") {
      totalPenjualan += omzet;
      totalGalon += item.jumlah;
      totalInfaq += item.potongan;
      totalBiayaOperasional +=
        item.jumlah * BIAYA_OPERASIONAL_PER_GALON;
    } else {
      totalPengeluaran += omzet;
    }
  });

  const totalModal = totalGalon * MODAL_PER_GALON;

  const profitOperasional =
    totalPenjualan
    - totalModal
    - totalBiayaOperasional
    - totalInfaq
    - totalPengeluaran;

  // TAMPILKAN KE UI
  setText("totalPenjualan", totalPenjualan.toLocaleString("id-ID"));
  setText("totalPengeluaran", totalPengeluaran.toLocaleString("id-ID"));
  setText(
    "saldoAkhir",
    (totalPenjualan - totalPengeluaran).toLocaleString("id-ID")
  );
  setText("totalInfaq", totalInfaq.toLocaleString("id-ID"));
  setText("totalGalon", totalGalon.toLocaleString("id-ID"));
  setText("totalModal", totalModal.toLocaleString("id-ID"));
  setText(
    "totalBiayaOperasional",
    totalBiayaOperasional.toLocaleString("id-ID")
  );
  setText(
    "profitOperasional",
    profitOperasional.toLocaleString("id-ID")
  );
}

// =====================================
// EXPORT EXCEL
// =====================================
function exportExcel() {
  if (data.length === 0) {
    alert("Data masih kosong!");
    return;
  }

  const rows = data.map((item, i) => ({
    No: i + 1,
    Nama: item.nama,
    Tanggal: item.tanggal,
    Jenis: item.jenis,
    Galon: item.jumlah,
    Harga: item.harga,
    Omzet: item.jumlah * item.harga,
    Infaq: item.potongan
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "MINUMO");

  XLSX.writeFile(wb, "Pembukuan_MINUMO.xlsx");
}

// =====================================
// EXPORT PDF
// =====================================
function exportPDF() {
  if (data.length === 0) {
    alert("Data masih kosong!");
    return;
  }

  const doc = new window.jspdf.jsPDF("p", "mm", "a4");

  let totalPenjualan = 0;
  let totalPengeluaran = 0;
  let totalInfaq = 0;
  let totalGalon = 0;
  let totalBiayaOperasional = 0;

  data.forEach(item => {
    const omzet = item.jumlah * item.harga;
    if (item.jenis === "Pemasukan") {
      totalPenjualan += omzet;
      totalGalon += item.jumlah;
      totalInfaq += item.potongan;
      totalBiayaOperasional +=
        item.jumlah * BIAYA_OPERASIONAL_PER_GALON;
    } else {
      totalPengeluaran += omzet;
    }
  });

  const totalModal = totalGalon * MODAL_PER_GALON;
  const profitOperasional =
    totalPenjualan
    - totalModal
    - totalBiayaOperasional
    - totalInfaq
    - totalPengeluaran;

  doc.setFontSize(16);
  doc.text('MAKMUR SENTOSA "Agen Air Mineral Minumo"', 14, 15);
  doc.setFontSize(10);
  doc.text(`Export: ${new Date().toLocaleString("id-ID")}`, 14, 22);

  doc.autoTable({
    startY: 30,
    head: [[
      "No", "Nama", "Tanggal", "Jenis", "Galon", "Harga", "Omzet", "Infaq"
    ]],
    body: data.map((item, i) => ([
      i + 1,
      item.nama,
      item.tanggal,
      item.jenis,
      item.jumlah,
      `Rp ${item.harga.toLocaleString("id-ID")}`,
      `Rp ${(item.jumlah * item.harga).toLocaleString("id-ID")}`,
      `Rp ${item.potongan.toLocaleString("id-ID")}`
    ])),
    theme: "grid",
    headStyles: { fillColor: [33, 150, 243], textColor: 255 }
  });

  doc.autoTable({
    startY: doc.lastAutoTable.finalY + 4,
    theme: "plain",
    styles: { fontSize: 11, cellPadding: 1 },
    columnStyles: {
      0: { cellWidth: 60, fontStyle: "bold" },
      1: { cellWidth: 5, halign: "center" },
      2: { cellWidth: 60 }
    },
    body: [
      ["Total Penjualan", ":", `Rp ${totalPenjualan.toLocaleString("id-ID")}`],
      ["Total Pengeluaran", ":", `Rp ${totalPengeluaran.toLocaleString("id-ID")}`],
      ["Total Galon", ":", totalGalon],
      ["Total Modal", ":", `Rp ${totalModal.toLocaleString("id-ID")}`],
      ["Biaya Operasional", ":", `Rp ${totalBiayaOperasional.toLocaleString("id-ID")}`],
      ["Total Infaq", ":", `Rp ${totalInfaq.toLocaleString("id-ID")}`],
      ["Profit Operasional", ":", `Rp ${profitOperasional.toLocaleString("id-ID")}`]
    ]
  });

  doc.save("Pembukuan_Makmur Sentosa.pdf");
}

// =====================================
// LOAD AWAL
// =====================================
renderTable();
updateInfo();
