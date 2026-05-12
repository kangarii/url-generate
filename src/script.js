/* ================================================================
       STATE — data yang dikelola secara global
    ================================================================ */
let allNames = []; // Array semua nama hasil parse
let allLinks = []; // Array objek {no, name, link}
let baseUrlVal = ""; // Base URL yang diinput user

/* ================================================================
       DRAG & DROP — setup event listener pada dropzone
    ================================================================ */
const dropzone = document.getElementById("dropzone");

dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.classList.add("drag-over");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("drag-over");
});

dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.classList.remove("drag-over");
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

/* File input change event */
document.getElementById("fileInput").addEventListener("change", (e) => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
});

/* ================================================================
       HANDLE FILE — deteksi ekstensi dan route ke parser yang tepat
    ================================================================ */
function handleFile(file) {
  hideError();
  const ext = file.name.split(".").pop().toLowerCase();

  if (!["xlsx", "xls", "csv", "txt"].includes(ext)) {
    showError("Format file tidak didukung. Gunakan .xlsx, .csv, atau .txt");
    return;
  }

  // Tampilkan info file
  document.getElementById("fileInfo").classList.remove("hidden");
  document.getElementById("fileName").textContent =
    `${file.name} (${formatBytes(file.size)})`;

  const reader = new FileReader();

  if (ext === "xlsx" || ext === "xls") {
    // Excel: baca sebagai ArrayBuffer lalu parse dengan SheetJS
    reader.onload = (e) => parseExcel(e.target.result);
    reader.readAsArrayBuffer(file);
  } else {
    // CSV / TXT: baca sebagai teks biasa
    reader.onload = (e) => parseText(e.target.result, ext);
    reader.readAsText(file, "UTF-8");
  }
}

/* ================================================================
       PARSER EXCEL — menggunakan SheetJS
    ================================================================ */
function parseExcel(buffer) {
  try {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Kumpulkan semua sel dari setiap baris, skip yang kosong
    const names = [];
    rows.forEach((row) => {
      row.forEach((cell) => {
        const name = String(cell ?? "").trim();
        if (name) names.push(name);
      });
    });

    validateAndSetNames(names);
  } catch (err) {
    showError("Gagal membaca file Excel. Pastikan file tidak rusak.");
    console.error(err);
  }
}

/* ================================================================
       PARSER TXT / CSV — split per baris, pisahkan dengan koma/titik koma
    ================================================================ */
function parseText(text, ext) {
  const lines = text.split(/\r?\n/);
  const names = [];

  lines.forEach((line) => {
    // CSV: pisahkan dengan koma atau titik koma
    const parts = ext === "csv" ? line.split(/[,;]/) : [line];
    parts.forEach((part) => {
      const name = part.trim().replace(/^["']|["']$/g, ""); // strip quotes
      if (name) names.push(name);
    });
  });

  validateAndSetNames(names);
}

/* ================================================================
       VALIDASI & SIMPAN NAMA
    ================================================================ */
function validateAndSetNames(names) {
  if (names.length === 0) {
    showError(
      "File tidak berisi nama yang valid. Pastikan minimal ada 1 nama.",
    );
    return;
  }
  allNames = names;
  showToast(`✅ ${names.length} nama berhasil dimuat!`, "📂");
}

/* ================================================================
       GENERATE LINKS — dipanggil saat tombol Generate diklik
    ================================================================ */
function generateLinks() {
  // Validasi: harus sudah ada nama
  if (allNames.length === 0) {
    showError("Upload file daftar nama terlebih dahulu.");
    return;
  }

  // Validasi: base URL tidak boleh kosong
  const rawUrl = document.getElementById("baseUrl").value.trim();
  if (!rawUrl) {
    showError("Masukkan Base URL undangan terlebih dahulu.");
    return;
  }

  hideError();
  baseUrlVal = rawUrl;

  // Buat array link: encode nama dengan encodeURIComponent
  allLinks = allNames.map((name, idx) => ({
    no: idx + 1,
    name: name,
    link: rawUrl + encodeURIComponent(name),
  }));

  // Update statistik
  document.getElementById("statTotal").textContent = allNames.length;
  document.getElementById("statLinks").textContent = allLinks.length;
  document.getElementById("statUrl").textContent = rawUrl;

  // Tampilkan section statistik & hasil
  document.getElementById("statsSection").classList.remove("hidden");
  document.getElementById("statsSection").classList.add("grid");
  document.getElementById("resultSection").classList.remove("hidden");

  // Reset search field
  document.getElementById("searchInput").value = "";
  renderTable(allLinks);

  // Scroll ke hasil
  document
    .getElementById("resultSection")
    .scrollIntoView({ behavior: "smooth", block: "start" });
  showToast(`🎉 ${allLinks.length} link berhasil di-generate!`, "🎉");
}

/* ================================================================
       RENDER TABLE — render baris tabel dari array data
    ================================================================ */
function renderTable(data) {
  const tbody = document.getElementById("resultTableBody");
  tbody.innerHTML = "";

  // Update filter count info
  document.getElementById("filterCount").textContent =
    data.length < allLinks.length
      ? `Menampilkan ${data.length} dari ${allLinks.length} nama`
      : `Menampilkan semua ${data.length} nama`;

  if (data.length === 0) {
    document.getElementById("noResult").classList.remove("hidden");
    return;
  }
  document.getElementById("noResult").classList.add("hidden");

  // Buat baris untuk setiap item
  data.forEach((item) => {
    const tr = document.createElement("tr");
    tr.className = "tbl-row transition-colors";
    tr.innerHTML = `
          <td class="px-4 py-3 text-gray-400 font-mono text-xs">${item.no}</td>
          <td class="px-4 py-3 font-semibold text-hitam">${escapeHtml(item.name)}</td>
          <td class="px-4 py-3">
            <a href="${escapeHtml(item.link)}" target="_blank"
               class="text-blue-600 hover:text-blue-800 hover:underline break-all text-xs font-mono">
              ${escapeHtml(item.link)}
            </a>
          </td>
          <td class="px-4 py-3 text-center">
            <button
              onclick="copyLink(this, '${escapeAttr(item.link)}')"
              title="Salin link"
              class="w-8 h-8 rounded-xl bg-kuning-100 hover:bg-kuning-400 transition flex items-center justify-center mx-auto text-base btn-press"
            >📋</button>
          </td>
        `;
    tbody.appendChild(tr);
  });
}

/* ================================================================
       FILTER / SEARCH — filter berdasarkan input pencarian
    ================================================================ */
function filterTable() {
  const q = document.getElementById("searchInput").value.toLowerCase().trim();
  if (!q) {
    renderTable(allLinks);
    return;
  }
  const filtered = allLinks.filter((item) =>
    item.name.toLowerCase().includes(q),
  );
  renderTable(filtered);
}

/* ================================================================
       COPY SINGLE LINK — salin satu link ke clipboard
    ================================================================ */
function copyLink(btn, link) {
  navigator.clipboard
    .writeText(link)
    .then(() => {
      const original = btn.textContent;
      btn.textContent = "✅";
      btn.classList.add("copy-flash", "bg-kuning-400");
      setTimeout(() => {
        btn.textContent = original;
        btn.classList.remove("copy-flash", "bg-kuning-400");
      }, 1200);
      showToast("Link disalin!", "📋");
    })
    .catch(() => fallbackCopy(link));
}

/* ================================================================
       COPY ALL LINKS — salin semua link ke clipboard
    ================================================================ */
function copyAllLinks() {
  if (allLinks.length === 0) return;
  const text = allLinks.map((item) => `${item.name}\t${item.link}`).join("\n");
  navigator.clipboard
    .writeText(text)
    .then(() => {
      showToast(`${allLinks.length} link disalin ke clipboard!`, "📋");
    })
    .catch(() => fallbackCopy(text));
}

/* Fallback copy untuk browser lama */
function fallbackCopy(text) {
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.position = "fixed";
  ta.style.opacity = "0";
  document.body.appendChild(ta);
  ta.select();
  document.execCommand("copy");
  document.body.removeChild(ta);
  showToast("Link disalin!", "📋");
}

/* ================================================================
       EXPORT EXCEL — export hasil ke file .xlsx menggunakan SheetJS
    ================================================================ */
function exportExcel() {
  if (allLinks.length === 0) return;

  // Siapkan data untuk SheetJS (header + rows)
  const rows = [["No", "Nama", "Link Undangan"]];
  allLinks.forEach((item) => rows.push([item.no, item.name, item.link]));

  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Atur lebar kolom agar lebih rapi
  ws["!cols"] = [{ wch: 5 }, { wch: 30 }, { wch: 80 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Link Undangan");

  XLSX.writeFile(wb, "link-undangan.xlsx");
  showToast("File Excel berhasil diunduh!", "📊");
}

/* ================================================================
       EXPORT TXT — export hasil ke file .txt
    ================================================================ */
function exportTxt() {
  if (allLinks.length === 0) return;

  const lines = allLinks.map((item) => `${item.name}\t${item.link}`);
  const content = `Link Undangan\n${"=".repeat(60)}\n${lines.join("\n")}`;

  const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "link-undangan.txt";
  a.click();
  URL.revokeObjectURL(url);
  showToast("File TXT berhasil diunduh!", "📝");
}

/* ================================================================
       RESET UPLOAD — hapus file yang sudah diupload
    ================================================================ */
function resetUpload() {
  allNames = [];
  document.getElementById("fileInfo").classList.add("hidden");
  document.getElementById("fileInput").value = "";
  document.getElementById("errorMsg").classList.add("hidden");
  hideError();
}

/* ================================================================
       UTILITY — helper functions
    ================================================================ */

/** Format bytes ke KB/MB */
function formatBytes(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

/** Escape HTML untuk mencegah XSS di innerHTML */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Escape untuk atribut HTML */
function escapeAttr(str) {
  return String(str).replace(/'/g, "&#39;").replace(/"/g, "&quot;");
}

/** Tampilkan pesan error */
function showError(msg) {
  const el = document.getElementById("errorMsg");
  document.getElementById("errorText").textContent = msg;
  el.classList.remove("hidden");
}

/** Sembunyikan pesan error */
function hideError() {
  document.getElementById("errorMsg").classList.add("hidden");
}

/* ================================================================
       TOAST NOTIFICATION — tampilkan notifikasi sementara
    ================================================================ */
let toastTimer = null;

function showToast(msg, emoji = "✅") {
  const toast = document.getElementById("toast");
  document.getElementById("toastText").textContent = msg;
  document.getElementById("toastEmoji").textContent = emoji;

  toast.classList.remove("hide");
  toast.classList.add("show");

  // Auto-hide setelah 2.5 detik
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.remove("show");
    toast.classList.add("hide");
  }, 2500);
}
