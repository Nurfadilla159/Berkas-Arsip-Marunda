const STORAGE_KEY = "berkasArsipMarunda";

const state = {
  rows: [],
  filtered: [],
  currentPage: 1,
  pageSize: 15,
  workbookName: "",
  importedAt: null
};

async function checkAuthentication() {

    const {
        data: { session },
        error
    } = await supabaseClient.auth.getSession();

    if (error) {
        console.error("Auth error:", error);
        window.location.href = "login.html";
        return null;
    }

    if (!session) {
        window.location.href = "login.html";
        return null;
    }

    return session;
}

checkAuthentication();


async function loadDataFromSupabase() {
    try {
        console.log("Mengambil seluruh data dari Supabase...");

        const semuaData = [];
        const ukuranBatch = 500;
        let mulai = 0;

        while (true) {
            const { data, error } = await supabaseClient
                .from("berkas")
                .select("*")
                .order("id", { ascending: true })
                .range(mulai, mulai + ukuranBatch - 1);

            if (error) {
                console.error(
                    "Gagal mengambil data dari Supabase:",
                    error
                );

                showToast(
                    "Gagal memuat data: " + error.message,
                    "error"
                );

                return;
            }

            if (!data || data.length === 0) {
                break;
            }

            semuaData.push(...data);

            console.log(
                `Sudah mengambil ${semuaData.length} data...`
            );

            if (data.length < ukuranBatch) {
                break;
            }

            mulai += ukuranBatch;
        }

        console.log("Total data diterima:", semuaData.length);

        state.rows = semuaData.map((row, index) => ({
            no: index + 1,
            id: row.id,
            date: row.tanggal_dokumen
                ? new Date(row.tanggal_dokumen + "T12:00:00")
                : null,
            contract: row.nomor_kontrak || "",
            name: row.nama_dokumen || "",
            vendor: row.vendor || "",
            value: Number(row.nilai) || 0,
            box: row.nama_box || "",
            type: row.jenis_berkas || "Umum"
        }));

        state.filtered = state.rows.slice();
        state.page = 1;

        populateYearFilter();
        populateTypeFilter();

        applyFilters(false);

        console.log(
            `Berhasil memuat ${state.rows.length} data dari Supabase.`
        );

    } catch (err) {
        console.error("ERROR LOAD SUPABASE:", err);

        showToast(
            "Terjadi kesalahan saat memuat data.",
            "error"
        );
    }
}

let realtimeChannel = null;

function setupRealtime() {

  console.log("Memulai koneksi Realtime...");

  if (realtimeChannel) {
    console.log("Realtime sudah aktif.");
    return;
  }

  realtimeChannel = supabaseClient
    .channel("berkas-realtime")

    .on(
      "postgres_changes",
      {
        event: "*",
        schema: "public",
        table: "berkas"
      },
      
      async (payload) => {

        console.log("=================================");
        console.log("REALTIME EVENT DITERIMA");
        console.log("Event:", payload.eventType);
        console.log("Data:", payload);
        console.log("=================================");

        await loadDataFromSupabase();
      }
    )

    .subscribe((status) => {

      console.log("REALTIME STATUS:", status);

    });
}

function $(id) {
  return document.getElementById(id);
}

function escapeHTML(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function parseNumber(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return 0;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  let text = String(value)
    .trim()
    .replace(/\s/g, "")
    .replace(/Rp/gi, "")
    .replace(/IDR/gi, "");

  if (!text) {
    return 0;
  }

 
  if (
    text.includes(".") &&
    text.includes(",")
  ) {

    text = text
      .replace(/\./g, "")
      .replace(",", ".");

    const result = Number(text);

    return Number.isFinite(result)
      ? result
      : 0;
  }

  if (
    text.includes(".") &&
    !text.includes(",")
  ) {

    const parts = text.split(".");

    if (
      parts.length > 1 &&
      parts.every(part => /^\d+$/.test(part))
    ) {

      const result = Number(
        parts.join("")
      );

      return Number.isFinite(result)
        ? result
        : 0;
    }
  }

  if (
    text.includes(",") &&
    !text.includes(".")
  ) {

    const result = Number(
      text.replace(",", ".")
    );

    return Number.isFinite(result)
      ? result
      : 0;
  }

  const result = Number(
    text.replace(/[^\d.-]/g, "")
  );

  return Number.isFinite(result)
    ? result
    : 0;
}

function formatRupiah(value) {

  const number = parseNumber(value);

  return new Intl.NumberFormat(
    "id-ID",
    {
      style: "currency",
      currency: "IDR",
      maximumFractionDigits: 0
    }
  ).format(number);
}

function excelSerialToDate(serial) {
  if (
    typeof serial !== "number" ||
    !Number.isFinite(serial)
  ) {
    return null;
  }

  const date = XLSX.SSF.parse_date_code(serial);

  if (!date) {
    return null;
  }

  return new Date(
    date.y,
    date.m - 1,
    date.d,
    12,
    0,
    0
  );
}

function parseDate(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return null;
  }

  if (typeof value === "number") {
    return excelSerialToDate(value);
  }

  if (value instanceof Date) {

    if (Number.isNaN(value.getTime())) {
      return null;
    }

    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate(),
      12,
      0,
      0
    );
  }

  const text =
    String(value).trim();

  if (!text) {
    return null;
  }

  let match =
    text.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})$/
    );

  if (match) {

    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      12,
      0,
      0
    );
  }

  match =
    text.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})T/
    );

  if (match) {

    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      12,
      0,
      0
    );
  }

  match =
    text.match(
      /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/
    );

  if (match) {

    return new Date(
      Number(match[3]),
      Number(match[2]) - 1,
      Number(match[1]),
      12,
      0,
      0
    );
  }

  return null;
}

function formatDate(value) {

  const date = parseDate(value);

  if (!date) {
    return "-";
  }

  const day = String(
    date.getDate()
  ).padStart(2, "0");

  const month = String(
    date.getMonth() + 1
  ).padStart(2, "0");

  const year = date.getFullYear();

  return `${day}/${month}/${year}`;
}

function formatDateInput(value) {

  const date = parseDate(value);

  if (!date) {
    return "";
  }

  const month = String(
    date.getMonth() + 1
  ).padStart(2, "0");

  const day = String(
    date.getDate()
  ).padStart(2, "0");

  return `${date.getFullYear()}-${month}-${day}`;
}


function normalizeHeader(value) {

  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/_/g, " ");
}

function findColumn(headers, candidates) {

  const normalizedHeaders = headers.map(
    normalizeHeader
  );

  for (const candidate of candidates) {

    const target = normalizeHeader(
      candidate
    );

    const exactIndex =
      normalizedHeaders.indexOf(target);

    if (exactIndex !== -1) {
      return exactIndex;
    }
  }

  for (let i = 0; i < normalizedHeaders.length; i++) {

    const header = normalizedHeaders[i];

    for (const candidate of candidates) {

      const target = normalizeHeader(
        candidate
      );

      if (
        header.includes(target) ||
        target.includes(header)
      ) {
        return i;
      }
    }
  }

  return -1;
}

function getMonthYearFromSheetName(sheetName) {

  const text = String(
    sheetName || ""
  ).trim();


  let match = text.match(
    /^(0?[1-9]|1[0-2])[\s._/-](20\d{2})$/
  );

  if (match) {

    return {
      month: Number(match[1]),
      year: Number(match[2])
    };
  }

  match = text.match(
    /^(20\d{2})[\s._/-](0?[1-9]|1[0-2])$/
  );

  if (match) {

    return {
      month: Number(match[2]),
      year: Number(match[1])
    };
  }

  return {
    month: null,
    year: null
  };
}

function convertSheet(worksheet, sheetName) {

  const raw = XLSX.utils.sheet_to_json(
    worksheet,
    {
      header: 1,
      defval: "",
      blankrows: false
    }
  );

  if (
    !raw ||
    raw.length === 0
  ) {
    return [];
  }

  let headerIndex = -1;

  for (
    let i = 0;
    i < Math.min(raw.length, 20);
    i++
  ) {

    const row = raw[i];

    const text = row
      .map(value =>
        String(value ?? "")
          .trim()
          .toLowerCase()
      )
      .join(" | ");

    const hasTanggal =
      text.includes("tanggal");

    const hasKontrak =
      text.includes("kontrak");

    const hasDokumen =
      text.includes("dokumen");

    const hasVendor =
      text.includes("vendor");

    const hits = [
      hasTanggal,
      hasKontrak,
      hasDokumen,
      hasVendor
    ].filter(Boolean).length;

    if (hits >= 2) {

      headerIndex = i;

      break;
    }
  }

  if (headerIndex === -1) {

    console.warn(
      `Header tidak ditemukan pada sheet: ${sheetName}`
    );

    return [];
  }

  const headers =
    raw[headerIndex].map(
      value =>
        String(value ?? "").trim()
    );

  const dateIndex = findColumn(
    headers,
    [
      "tanggal dokumen",
      "tanggal",
      "tgl dokumen",
      "tgl"
    ]
  );

  const contractIndex = findColumn(
    headers,
    [
      "nomor kontrak",
      "no kontrak",
      "nomor kontrak/spk",
      "kontrak"
    ]
  );

  const nameIndex = findColumn(
    headers,
    [
      "nama dokumen",
      "nama berkas",
      "dokumen"
    ]
  );

  const vendorIndex = findColumn(
    headers,
    [
      "vendor",
      "nama vendor",
      "mitra"
    ]
  );

  const valueIndex = findColumn(
    headers,
    [
      "nilai",
      "nilai tagihan",
      "nilai kontrak",
      "total",
      "jumlah"
    ]
  );

  const boxIndex = findColumn(
    headers,
    [
      "nama box",
      "box",
      "no box",
      "nomor box"
    ]
  );


  const typeIndex = findColumn(
    headers,
    [
      "jenis berkas",
      "jenis",
      "berkas",
      "kategori"
    ]
  );


  const noIndex = findColumn(
    headers,
    [
      "no.",
      "no",
      "nomor"
    ]
  );

  console.log(
    `Sheet ${sheetName}`,
    {
      headerIndex,
      headers,
      dateIndex,
      contractIndex,
      nameIndex,
      vendorIndex,
      valueIndex,
      boxIndex,
      typeIndex,
      noIndex
    }
  );

  const sheetInfo =
    getMonthYearFromSheetName(
      sheetName
    );

  const rows = [];

  for (
    let i = headerIndex + 1;
    i < raw.length;
    i++
  ) {

    const row = raw[i];


    if (!Array.isArray(row)) {
      continue;
    }


    const values =
      row.map(value =>
        String(value ?? "").trim()
      );


    if (
      values.every(
        value => value === ""
      )
    ) {
      continue;
    }


    const rawDate =
      dateIndex >= 0
        ? row[dateIndex]
        : null;


    const date =
      parseDate(rawDate);


    const contract =
      contractIndex >= 0
        ? String(
            row[contractIndex] ?? ""
          ).trim()
        : "";


    const name =
      nameIndex >= 0
        ? String(
            row[nameIndex] ?? ""
          ).trim()
        : "";


    const vendor =
      vendorIndex >= 0
        ? String(
            row[vendorIndex] ?? ""
          ).trim()
        : "";


    const value =
      valueIndex >= 0
        ? parseNumber(
            row[valueIndex]
          )
        : 0;

    const box =
      boxIndex >= 0
        ? String(
            row[boxIndex] ?? ""
          ).trim()
        : "";

    const type =
      typeIndex >= 0
        ? String(
            row[typeIndex] ?? ""
          ).trim()
        : "Umum";

    const hasMainData =
      name ||
      contract ||
      vendor ||
      box ||
      date;

    if (!hasMainData) {
      continue;
    }

    let finalDate = date;

    if (
      !finalDate &&
      sheetInfo.year &&
      sheetInfo.month
    ) {

      finalDate = new Date(
        sheetInfo.year,
        sheetInfo.month - 1,
        1,
        12
      );
    }

    rows.push({

      no:
        noIndex >= 0
          ? row[noIndex]
          : rows.length + 1,

      date:
        finalDate,

      contract,

      name,

      vendor,

      value,

      box,

      type:
        type || "Umum"

    });
  }

  console.log(
    `Sheet "${sheetName}" menghasilkan ${rows.length} data`
  );

  return rows;
}

function convertRowForSupabase(row) {
  let tanggalDokumen = null;

  if (row.date instanceof Date && !isNaN(row.date.getTime())) {
    const year = row.date.getFullYear();
    const month = String(row.date.getMonth() + 1).padStart(2, "0");
    const day = String(row.date.getDate()).padStart(2, "0");

    tanggalDokumen = `${year}-${month}-${day}`;
  } 
  else if (row.date) {
    const value = String(row.date).trim();

    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      tanggalDokumen = value;
    }

    else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(value)) {
      const parts = value.split("/");
      const day = parts[0].padStart(2, "0");
      const month = parts[1].padStart(2, "0");
      const year = parts[2];

      tanggalDokumen = `${year}-${month}-${day}`;
    }
  }

  return {
    tanggal_dokumen: tanggalDokumen,
    nomor_kontrak: row.contract || "",
    nama_dokumen: row.name || "",
    vendor: row.vendor || "",
    nilai: Number(row.value) || 0,
    nama_box: row.box || "",
    jenis_berkas: row.type || "Umum"
  };
}

async function handleExcel(files) {

  if (!files || files.length === 0) {
    return;
  }

  try {

    const importedRows = [];
    const fileNames = [];

    for (const file of files) {

      fileNames.push(file.name);

      const buffer = await file.arrayBuffer();

      const workbook = XLSX.read(
        buffer,
        {
          type: "array",
          cellDates: false,
          raw: true
        }
      );

      for (const sheetName of workbook.SheetNames) {

        const worksheet = workbook.Sheets[sheetName];

        const rows = convertSheet(
          worksheet,
          sheetName
        );

        importedRows.push(...rows);
      }
    }

    if (importedRows.length === 0) {

      showToast(
        "Tidak ada data yang berhasil dibaca.",
        "error"
      );

      return;
    }


    console.log(
      "Data hasil pembacaan Excel:",
      importedRows
    );

    const supabaseRows = importedRows.map(
      convertRowForSupabase
    );


    console.log(
      "Data yang akan dikirim ke Supabase:",
      supabaseRows
    );

    showToast(
      `Menyimpan ${supabaseRows.length} data ke database...`,
      "success"
    );

    const {
      data,
      error
    } = await supabaseClient
      .from("berkas")
      .insert(supabaseRows)
      .select();


    if (error) {
      console.error(
        "Supabase INSERT ERROR:",
        error
      );

      showToast(
        "Data gagal disimpan ke Supabase: " +
        error.message,
        "error"
      );

      return;
    }


    console.log(
      "Data berhasil masuk Supabase:",
      data
    );


    await loadDataFromSupabase();

    state.workbookName =
      fileNames.join(" + ");

    state.importedAt =
      new Date();


    populateFilters();

    applyFilters();

    updateLastUpdate();


    const preview =
      $("importPreview");

    const previewText =
      $("previewText");


    if (preview) {

      preview.classList.remove(
        "hidden"
      );
    }


    if (previewText) {

      previewText.textContent =
        `${importedRows.length} data berhasil diimpor dari ${fileNames.length} file. Total data sekarang ${state.rows.length} berkas.`;
    }


    showToast(
      `${importedRows.length} data berhasil diimport ke database.`,
      "success"
    );


    const input =
      $("excelInput");

    if (input) {
      input.value = "";
    }


  } catch (error) {

    console.error(
      "Import Excel error:",
      error
    );

    showToast(
      "Terjadi kesalahan saat mengimpor data Excel.",
      "error"
    );
  }
}

function updateLastUpdate() {

  const element =
    $("lastUpdate");


  if (!element) {
    return;
  }


  if (
    !state.importedAt
  ) {

    element.textContent =
      "Belum ada data diimpor";

    return;
  }


  const date =
    parseDate(
      state.importedAt
    );


  if (!date) {

    element.textContent =
      "Data tersimpan";

    return;
  }


  element.textContent =
    `Update terakhir: ${date.toLocaleString(
      "id-ID",
      {
        dateStyle: "medium",
        timeStyle: "short"
      }
    )}`;
}


function populateFilters() {

  populateYearFilter();

  populateTypeFilter();
}


function populateYearFilter() {

  const select =
    $("yearFilter");


  if (!select) {
    return;
  }


  const currentValue =
    select.value;


  const years =
    new Set();


  state.rows.forEach(
    row => {

      const date =
        parseDate(row.date);


      if (date) {

        years.add(
          date.getFullYear()
        );
      }

    }
  );


  const sortedYears =
    Array.from(years)
      .sort(
        (a, b) => b - a
      );


  select.innerHTML =
    `<option value="">Semua Tahun</option>`;


  sortedYears.forEach(
    year => {

      const option =
        document.createElement(
          "option"
        );

      option.value =
        String(year);

      option.textContent =
        String(year);

      select.appendChild(
        option
      );
    }
  );


  if (
    sortedYears
      .map(String)
      .includes(currentValue)
  ) {

    select.value =
      currentValue;
  }
}


function populateTypeFilter() {

  const select =
    $("typeFilter");


  if (!select) {
    return;
  }


  const currentValue =
    select.value;


  const types =
    new Set();


  state.rows.forEach(
    row => {

      const type =
        String(
          row.type || "Umum"
        ).trim();


      if (type) {
        types.add(type);
      }
    }
  );

  const sortedTypes =
    Array.from(types)
      .sort(
        (a, b) =>
          a.localeCompare(
            b,
            "id"
          )
      );


  select.innerHTML =
    `<option value="">Semua Berkas</option>`;


  sortedTypes.forEach(
    type => {

      const option =
        document.createElement(
          "option"
        );

      option.value =
        type;

      option.textContent =
        type;

      select.appendChild(
        option
      );
    }
  );


  if (
    sortedTypes.includes(
      currentValue
    )
  ) {

    select.value =
      currentValue;
  }
}

function applyFilters(
  resetPage = true
) {

  const search =
    String(
      $("searchInput")?.value || ""
    )
      .trim()
      .toLowerCase();


  const year =
    String(
      $("yearFilter")?.value || ""
    );


  const month =
    String(
      $("monthFilter")?.value || ""
    );


  const type =
    String(
      $("typeFilter")?.value || ""
    );


  state.filtered =
    state.rows.filter(
      row => {

        const date =
          parseDate(row.date);


        const searchable = [
          row.contract,
          row.name,
          row.vendor,
          row.box,
          row.type
        ]
          .join(" ")
          .toLowerCase();


        if (
          search &&
          !searchable.includes(
            search
          )
        ) {
          return false;
        }


        if (year) {

          if (
            !date ||
            String(
              date.getFullYear()
            ) !== year
          ) {
            return false;
          }
        }


        if (month) {

          if (
            !date ||
            String(
              date.getMonth() + 1
            ) !== month
          ) {
            return false;
          }
        }


        if (type) {

          if (
            String(
              row.type || "Umum"
            ) !== type
          ) {
            return false;
          }
        }


        return true;
      }
    );


  if (resetPage) {
    state.currentPage = 1;
  }

  renderTable();

  updateStats();

  updateResultInfo();

  renderPagination();
}

function renderTable() {

  const tbody =
    $("dataBody");


  if (!tbody) {
    return;
  }

  if (
    state.filtered.length === 0
  ) {

    tbody.innerHTML = `

      <tr class="empty-row">

        <td colspan="9">

          <div class="empty-state">
                    
          <div class="stat-icon blue">
             <i class="fa-solid fa-folder-open"></i>
          </div>

            <strong>
              Belum ada data
            </strong>

            <span>
              Tidak ada data yang sesuai dengan filter.
            </span>

            <button
              class="btn btn-primary"
              type="button"
              id="emptyImportBtn"
            >
              Import Excel
            </button>

          </div>

        </td>

      </tr>

    `;


    bindEmptyImport();

    return;
  }


  const start =
    (
      state.currentPage - 1
    ) * state.pageSize;


  const end =
    start +
    state.pageSize;


  const pageRows =
    state.filtered.slice(
      start,
      end
    );


  tbody.innerHTML =
    pageRows.map(
      row => {

        const originalIndex =
          state.rows.indexOf(
            row
          );


        return `

          <tr>

            <td>
              ${row.no}
            </td>

            <td>
              ${escapeHTML(
                formatDate(
                  row.date
                )
              )}
            </td>

            <td>
              ${escapeHTML(
                row.contract
              )}
            </td>

            <td>
              ${escapeHTML(
                row.name
              )}
            </td>

            <td>
              ${escapeHTML(
                row.vendor
              )}
            </td>

            <td class="numeric">
              ${formatRupiah(
                row.value
              )}
            </td>

            <td>
              ${escapeHTML(
                row.box
              )}
            </td>

            <td>
              ${escapeHTML(
                row.type || "Umum"
              )}
            </td>

            <td class="action-cell">

              <div class="action-buttons">

                <button
                  type="button"
                  class="action-btn edit-btn"
                  data-index="${originalIndex}"
                  title="Edit data"
                >
                  <i class="fa-solid fa-pen"></i>
                </button>

                <button
                  type="button"
                  class="action-btn delete-btn"
                  data-id="${row.id}"
                  title="Hapus data"
                >
                  <i class="fa-solid fa-trash"></i>
                </button>

              </div>

            </td>

          </tr>

        `;
      }
    )
    .join("");
}


function updateResultInfo() {

  const element =
    $("resultInfo");


  if (!element) {
    return;
  }


  const total =
    state.filtered.length;


  if (state.rows.length === 0) {

    element.textContent =
      "Belum ada data. Silakan import file Excel.";

    return;
  }


  element.textContent =
    `Menampilkan ${total} dari ${state.rows.length} berkas.`;
}


function updateStats() {

  const rows =
    state.filtered;


  const totalFiles =
    $("totalFiles");


  const totalValue =
    $("totalValue");


  const totalVendors =
    $("totalVendors");


  if (totalFiles) {

    totalFiles.textContent =
      rows.length.toLocaleString(
        "id-ID"
      );
  }

  const total =
    rows.reduce(
      (
        sum,
        row
      ) =>
        sum +
        parseNumber(
          row.value
        ),
      0
    );


  if (totalValue) {

    totalValue.textContent =
      formatRupiah(total);
  }


  const vendors =
    new Set();


  rows.forEach(
    row => {

      const vendor =
        String(
          row.vendor || ""
        ).trim();


      if (vendor) {
        vendors.add(
          vendor.toLowerCase()
        );
      }

    }
  );


  if (totalVendors) {

    totalVendors.textContent =
      vendors.size.toLocaleString(
        "id-ID"
      );
  }


  const filesHint =
    $("filesHint");

  if (filesHint) {

    filesHint.textContent =
      "Data yang tampil";
  }


  const valueHint =
    $("valueHint");

  if (valueHint) {

    valueHint.textContent =
      "Nilai dari data yang tampil";
  }


  const vendorHint =
    $("vendorHint");

  if (vendorHint) {

    vendorHint.textContent =
      "Vendor yang tampil";
  }
}

function renderPagination() {

  const container =
    $("pagination");


  if (!container) {
    return;
  }


  const totalPages =
    Math.ceil(
      state.filtered.length /
      state.pageSize
    );


  if (
    totalPages <= 1
  ) {

    container.innerHTML =
      "";

    return;
  }


  let html = "";


  html += `

    <button
      type="button"
      class="page-btn"
      data-page="${state.currentPage - 1}"
      ${state.currentPage === 1 ? "disabled" : ""}
    >
      ‹
    </button>

  `;


  const maxVisible = 5;


  let startPage =
    Math.max(
      1,
      state.currentPage -
      Math.floor(
        maxVisible / 2
      )
    );


  let endPage =
    Math.min(
      totalPages,
      startPage +
      maxVisible -
      1
    );


  if (
    endPage - startPage <
    maxVisible - 1
  ) {

    startPage =
      Math.max(
        1,
        endPage -
        maxVisible +
        1
      );
  }


  for (
    let page = startPage;
    page <= endPage;
    page++
  ) {

    html += `

      <button
        type="button"
        class="page-btn ${
          page === state.currentPage
            ? "active"
            : ""
        }"
        data-page="${page}"
      >
        ${page}
      </button>

    `;
  }


  html += `

    <button
      type="button"
      class="page-btn"
      data-page="${state.currentPage + 1}"
      ${
        state.currentPage === totalPages
          ? "disabled"
          : ""
      }
    >
      ›
    </button>

  `;

  container.innerHTML =
    html;


  container
    .querySelectorAll(
      ".page-btn"
    )
    .forEach(
      button => {

        button.addEventListener(
          "click",
          () => {

            const page =
              Number(
                button.dataset.page
              );

            if (
              page < 1 ||
              page > totalPages
            ) {
              return;
            }

            state.currentPage =
              page;


            renderTable();

            renderPagination();

            window.scrollTo({
              top: 0,
              behavior: "smooth"
            });

          }
        );

      }
    );
}


function editRow(index) {

  const row =
    state.rows[index];


  if (!row) {

    showToast(
      "Data tidak ditemukan.",
      "error"
    );

    return;
  }

  const modal =
    $("editModal");


  if (!modal) {
    return;
  }

  $("editIndex").value =
    index;


  $("editDate").value =
    formatDateInput(
      row.date
    );


  $("editContract").value =
    row.contract || "";


  $("editName").value =
    row.name || "";


  $("editVendor").value =
    row.vendor || "";


  $("editValue").value =
    parseNumber(
      row.value
    );


  $("editBox").value =
    row.box || "";


  $("editType").value =
    row.type || "Umum";


  modal.classList.remove(
    "hidden"
  );
}

function closeEditModal() {

  const modal =
    $("editModal");


  if (!modal) {
    return;
  }


  modal.classList.add(
    "hidden"
  );
}

async function saveEditedRow() {

  const index =
    Number(
      $("editIndex").value
    );


  if (
    !Number.isInteger(index) ||
    !state.rows[index]
  ) {

    showToast(
      "Data yang diedit tidak ditemukan.",
      "error"
    );

    return;
  }


  const row =
    state.rows[index];

  const id =
    row.id;


  if (!id) {

    showToast(
      "ID data tidak ditemukan.",
      "error"
    );

    return;
  }

  const dateValue =
    $("editDate").value;


  let tanggalDokumen = null;

  if (dateValue) {

    const parsedDate =
      parseDate(dateValue);


    if (
      parsedDate instanceof Date &&
      !isNaN(parsedDate.getTime())
    ) {

      const year =
        parsedDate.getFullYear();

      const month =
        String(
          parsedDate.getMonth() + 1
        ).padStart(2, "0");

      const day =
        String(
          parsedDate.getDate()
        ).padStart(2, "0");


      tanggalDokumen =
        `${year}-${month}-${day}`;
    }
  }


  const updatedData = {

    tanggal_dokumen:
      tanggalDokumen,

    nomor_kontrak:
      $("editContract").value.trim(),

    nama_dokumen:
      $("editName").value.trim(),

    vendor:
      $("editVendor").value.trim(),

    nilai:
      parseNumber(
        $("editValue").value
      ),

    nama_box:
      $("editBox").value.trim(),

    jenis_berkas:
      $("editType").value.trim() ||
      "Umum",

    updated_at:
      new Date().toISOString()

  };


  try {

    showToast(
      "Menyimpan perubahan...",
      "success"
    );


    const { error } =
      await supabaseClient
        .from("berkas")
        .update(updatedData)
        .eq("id", id);


    if (error) {

      console.error(
        "Gagal update data:",
        error
      );


      showToast(
        "Data gagal diperbarui: " +
        error.message,
        "error"
      );

      return;
    }

    await loadDataFromSupabase();

    closeEditModal();

    showToast(
      "Data berhasil diperbarui dan disimpan.",
      "success"
    );


  } catch (error) {

    console.error(
      "ERROR UPDATE:",
      error
    );


    showToast(
      "Terjadi kesalahan saat menyimpan perubahan.",
      "error"
    );
  }
}

function closeManualAddModal() {
  const modal = document.getElementById("manualAddModal");
  if (modal) modal.classList.add("hidden");
}

function setupManualAdd() {
  const modal = document.getElementById("manualAddModal");
  const form = document.getElementById("manualAddForm");
  const openButton = document.getElementById("openManualAdd");
  const closeButton = document.getElementById("closeManualAdd");
  const cancelButton = document.getElementById("cancelManualAdd");
  const overlay = document.getElementById("manualAddOverlay");
  const saveButton = document.getElementById("saveManualAdd");

  if (!modal || !form || !openButton || !saveButton) {
    console.error("Elemen Tambah Berkas tidak lengkap.", {
      modal: !!modal,
      form: !!form,
      openButton: !!openButton,
      saveButton: !!saveButton
    });
    return;
  }

  if (form.dataset.manualAddReady === "true") {
    return;
  }
  form.dataset.manualAddReady = "true";

  openButton.addEventListener("click", function () {
    form.reset();

    const typeInput = document.getElementById("manualType");
    const valueInput = document.getElementById("manualValue");
    const dateInput = document.getElementById("manualDate");

    if (typeInput) typeInput.value = "Umum";
    if (valueInput) valueInput.value = "0";

    modal.classList.remove("hidden");

    if (dateInput) dateInput.focus();
  });

  closeButton?.addEventListener("click", closeManualAddModal);
  cancelButton?.addEventListener("click", closeManualAddModal);
  overlay?.addEventListener("click", closeManualAddModal);

  form.addEventListener("submit", async function (event) {
    event.preventDefault();

    if (saveButton.disabled) return;

    if (
      typeof supabaseClient === "undefined" ||
      !supabaseClient
    ) {
      showToast("Koneksi Supabase belum tersedia.", "error");
      console.error("supabaseClient tidak tersedia.");
      return;
    }

    const getValue = function (id) {
      const element = document.getElementById(id);
      return element ? element.value.trim() : "";
    };

    const namaDokumen = getValue("manualName");
    const nilaiInput = getValue("manualValue");
    const nilai = Number(nilaiInput);

    if (!namaDokumen) {
      showToast("Nama dokumen wajib diisi.", "error");
      document.getElementById("manualName")?.focus();
      return;
    }

    if (
      nilaiInput === "" ||
      !Number.isFinite(nilai) ||
      nilai < 0
    ) {
      showToast("Nilai berkas harus berupa angka 0 atau lebih.", "error");
      document.getElementById("manualValue")?.focus();
      return;
    }

    const dataBaru = {
      tanggal_dokumen: getValue("manualDate") || null,
      nomor_kontrak: getValue("manualContract"),
      nama_dokumen: namaDokumen,
      vendor: getValue("manualVendor"),
      nilai: nilai,
      nama_box: getValue("manualBox"),
      jenis_berkas: getValue("manualType") || "Umum"
    };

    const teksTombolAwal = saveButton.innerHTML;

    try {
      saveButton.disabled = true;
      saveButton.innerHTML =
        '<i class="fa-solid fa-spinner fa-spin"></i> Menyimpan...';

      console.log("Mengirim data berkas ke Supabase:", dataBaru);

      const { data, error } = await supabaseClient
        .from("berkas")
        .insert([dataBaru])
        .select();

      if (error) {
        console.error("GAGAL INSERT BERKAS:", error);

        showToast(
          "Gagal menyimpan berkas: " + error.message,
          "error"
        );
        return;
      }

      if (!data || data.length === 0) {
        console.error(
          "Supabase tidak mengembalikan data setelah insert.",
          data
        );

        showToast(
          "Data belum dapat dikonfirmasi. Periksa izin akses Supabase.",
          "error"
        );
        return;
      }

      console.log("BERKAS BERHASIL DISIMPAN:", data[0]);

      if (typeof loadDataFromSupabase !== "function") {
        console.error("Fungsi loadDataFromSupabase tidak ditemukan.");

        showToast(
          "Data tersimpan, tetapi fungsi pemuatan tabel tidak ditemukan.",
          "error"
        );
        return;
      }

      await loadDataFromSupabase();

      closeManualAddModal();
      form.reset();

      showToast("Berkas berhasil disimpan dan data diperbarui.", "success");

    } catch (error) {
      console.error("KESALAHAN TAMBAH BERKAS:", error);

      showToast(
        "Terjadi kesalahan: " +
        (error.message || "Tidak diketahui"),
        "error"
      );

    } finally {
      saveButton.disabled = false;
      saveButton.innerHTML = teksTombolAwal;
    }
  });

  document.addEventListener("keydown", function (event) {
    if (
      event.key === "Escape" &&
      !modal.classList.contains("hidden")
    ) {
      closeManualAddModal();
    }
  });

  console.log("Fitur Tambah Berkas siap digunakan.");
}

if (document.readyState === "loading") {
  document.addEventListener(
    "DOMContentLoaded",
    setupManualAdd,
    { once: true }
  );
} else {
  setupManualAdd();
}

function deleteRow(id) {

  if (!id) {

    showToast(
      "ID data tidak ditemukan.",
      "error"
    );

    return;
  }

  const row = state.rows.find(
    item => Number(item.id) === Number(id)
  );

  if (!row) {

    showToast(
      "Data tidak ditemukan.",
      "error"
    );

    return;
  }

  const modal = $("singleDeleteModal");
  const text = $("singleDeleteText");

  if (!modal) {

    console.error(
      "singleDeleteModal tidak ditemukan."
    );

    return;
  }

  modal.dataset.deleteId = id;

  const documentName =
    row.name || "data ini";

  text.textContent =
    `Apakah kamu yakin ingin menghapus "${documentName}"?`;

  modal.classList.remove("hidden");
}

function closeSingleDeleteModal() {

  const modal =
    $("singleDeleteModal");

  if (!modal) return;

  modal.classList.add("hidden");

  modal.dataset.deleteId = "";
}

async function confirmSingleDelete() {

  const modal =
    $("singleDeleteModal");

  const confirmButton =
    $("confirmSingleDelete");

  if (!modal) return;


  const id =
    Number(modal.dataset.deleteId);


  if (!id) {

    showToast(
      "ID data tidak ditemukan.",
      "error"
    );

    closeSingleDeleteModal();

    return;
  }


  try {

    if (confirmButton) {

      confirmButton.disabled = true;

      confirmButton.innerHTML = `
        <i class="fa-solid fa-spinner fa-spin"></i>
        Menghapus...
      `;
    }

    const { error } =
      await supabaseClient
        .from("berkas")
        .delete()
        .eq("id", id);


    if (error) {

      console.error(
        "Gagal menghapus data:",
        error
      );

      closeSingleDeleteModal();

      showToast(
        "Data gagal dihapus: " +
        error.message,
        "error"
      );

      return;
    }


    closeSingleDeleteModal();

    await loadDataFromSupabase();

    setTimeout(() => {

      showToast(
        "Data berhasil dihapus.",
        "success"
      );

    }, 150);


  } catch (error) {

    console.error(
      "ERROR DELETE:",
      error
    );

    closeSingleDeleteModal();

    showToast(
      "Terjadi kesalahan saat menghapus data.",
      "error"
    );


  } finally {

    if (confirmButton) {

      confirmButton.disabled = false;

      confirmButton.innerHTML = `
        <i class="fa-solid fa-trash-can"></i>
        Hapus Data
      `;

    }

  }
}

function setupSingleDeleteModal() {

  const modal =
    $("singleDeleteModal");

  const closeButton =
    $("closeSingleDeleteModal");

  const cancelButton =
    $("cancelSingleDelete");

  const confirmButton =
    $("confirmSingleDelete");

  const overlay =
    document.querySelector(
      ".single-delete-overlay"
    );


  if (!modal) {

    console.error(
      "singleDeleteModal tidak ditemukan."
    );

    return;
  }

  closeButton?.addEventListener(
    "click",
    () => {

      closeSingleDeleteModal();

    }
  );


  cancelButton?.addEventListener(
    "click",
    () => {

      closeSingleDeleteModal();

    }
  );

  overlay?.addEventListener(
    "click",
    () => {

      closeSingleDeleteModal();

    }
  );

  confirmButton?.addEventListener(
    "click",
    async () => {

      await confirmSingleDelete();

    }
  );


  document.addEventListener(
    "keydown",
    event => {

      if (
        event.key === "Escape" &&
        !modal.classList.contains("hidden")
      ) {

        closeSingleDeleteModal();

      }

    }
  );

}

async function clearData() {

  if (state.rows.length === 0) {

    showToast(
      "Belum ada data untuk dihapus.",
      "error"
    );

    return;
  }

  const modal = $("deleteModal");
  const modalText = $("deleteModalText");

  if (!modal) {

    console.error(
      "deleteModal tidak ditemukan di HTML."
    );

    return;
  }

  modalText.textContent =
    `Semua ${state.rows.length} data berkas akan dihapus secara permanen.`;

  modal.classList.remove("hidden");
}

async function confirmClearData() {

  const confirmButton = $("confirmDelete");

  try {

    if (confirmButton) {

      confirmButton.disabled = true;

      confirmButton.innerHTML = `
        <i class="fa-solid fa-spinner fa-spin"></i>
        Menghapus...
      `;
    }

    const { error } = await supabaseClient
      .from("berkas")
      .delete()
      .gt("id", 0);


    if (error) {

      console.error(
        "Gagal menghapus semua data:",
        error
      );

      closeDeleteModal();

      showToast(
        "Data gagal dihapus: " + error.message,
        "error"
      );

      return;
    }

    state.rows = [];

    state.filtered = [];

    state.currentPage = 1;

    state.workbookName = "";

    state.importedAt = null;

    localStorage.removeItem(
      STORAGE_KEY
    );


    populateFilters();

    applyFilters();

    updateLastUpdate();

    const preview = $("importPreview");

    if (preview) {

      preview.classList.add(
        "hidden"
      );

    }

    closeDeleteModal();


    setTimeout(() => {

      showToast(
        "Semua data berhasil dihapus.",
        "success"
      );

    }, 150);


  } catch (error) {

    console.error(
      "ERROR CLEAR DATA:",
      error
    );

    closeDeleteModal();

    showToast(
      "Terjadi kesalahan saat menghapus semua data.",
      "error"
    );


  } finally {

    if (confirmButton) {

      confirmButton.disabled = false;

      confirmButton.innerHTML = `
        <i class="fa-solid fa-trash-can"></i>
        Hapus Semua
      `;

    }

  }
}

function setupDeleteModal() {

  const modal = $("deleteModal");

  const closeButton =
    $("closeDeleteModal");

  const cancelButton =
    $("cancelDelete");

  const confirmButton =
    $("confirmDelete");

  const overlay =
    document.querySelector(
      ".delete-modal-overlay"
    );


  if (!modal) {

    console.error(
      "deleteModal tidak ditemukan."
    );

    return;
  }


  closeButton?.addEventListener(
    "click",
    () => {

      closeDeleteModal();

    }
  );


  cancelButton?.addEventListener(
    "click",
    () => {

      closeDeleteModal();

    }
  );


  overlay?.addEventListener(
    "click",
    () => {

      closeDeleteModal();

    }
  );


  confirmButton?.addEventListener(
    "click",
    async () => {

      await confirmClearData();

    }
  );


  document.addEventListener(
    "keydown",
    (event) => {

      if (event.key === "Escape") {

        closeDeleteModal();

      }

    }
  );

}


function exportData() {

  if (
    state.filtered.length === 0
  ) {

    showToast(
      "Tidak ada data untuk diekspor.",
      "error"
    );

    return;
  }

  const exportRows =
    state.filtered.map(
      (row, index) => ({

        NO:
          index + 1,

        "TANGGAL DOKUMEN":
          formatDate(
            row.date
          ),

        "NOMOR KONTRAK":
          row.contract,

        "NAMA DOKUMEN":
          row.name,

        VENDOR:
          row.vendor,

        NILAI:
          parseNumber(
            row.value
          ),

        "NAMA BOX":
          row.box,

        "JENIS BERKAS":
          row.type || "Umum"

      })
    );

  const worksheet =
    XLSX.utils.json_to_sheet(
      exportRows
    );


  const workbook =
    XLSX.utils.book_new();


  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Data Arsip"
  );

  const date =
    new Date();


  const dateText =
    `${date.getFullYear()}-${String(
      date.getMonth() + 1
    ).padStart(2, "0")}-${String(
      date.getDate()
    ).padStart(2, "0")}`;


  XLSX.writeFile(
    workbook,
    `Berkas_Arsip_Marunda_${dateText}.xlsx`
  );

  showToast(
    "Data berhasil diekspor.",
    "success"
  );
}

function setupNavigation() {
    const navItems = document.querySelectorAll(".nav-item");

    const dashboard = $("dashboardSection");
    const importSection = $("importSection");

    navItems.forEach(item => {
        item.addEventListener("click", () => {
            const section = item.dataset.section;

            navItems.forEach(nav => {
                nav.classList.remove("active");
            });

            item.classList.add("active");

            if (section === "import") {
                dashboard?.classList.add("hidden");
                importSection?.classList.remove("hidden");
            } else {
                dashboard?.classList.remove("hidden");
                importSection?.classList.add("hidden");
            }

            if (section === "dashboard") {
                setTimeout(() => {
                    window.scrollTo({
                        top: 0,
                        behavior: "smooth"
                    });
                }, 50);
            }

            if (section === "data") {
                setTimeout(() => {
                    $("dataPanel")?.scrollIntoView({
                        behavior: "smooth",
                        block: "start"
                    });
                }, 50);
            }

            $("sidebar")?.classList.remove("open");
        });
    });

    const mobileMenu = $("mobileMenu");

    if (mobileMenu) {
        mobileMenu.addEventListener("click", () => {
            $("sidebar")?.classList.toggle("open");
        });
    }
}

function setupImport() {

  const input =
    $("excelInput");


  const importBtn =
    $("importBtn");


  const emptyImportBtn =
    $("emptyImportBtn");


  const dropzone =
    $("dropzone");


  if (input) {

    input.addEventListener(
      "change",
      event => {

        const files =
          Array.from(
            event.target.files || []
          );

        handleExcel(files);

      }
    );
  }


  if (importBtn) {

    importBtn.addEventListener(
      "click",
      () => {

        document
          .querySelectorAll(
            ".nav-item"
          )
          .forEach(
            item => {

              item.classList.toggle(
                "active",
                item.dataset.section === "import"
              );

            }
          );


        $("dashboardSection")
          ?.classList.add(
            "hidden"
          );


        $("importSection")
          ?.classList.remove(
            "hidden"
          );


        setTimeout(
          () => {

            input?.click();

          },
          100
        );

      }
    );
  }


  if (emptyImportBtn) {

    emptyImportBtn.addEventListener(
      "click",
      () => {

        input?.click();

      }
    );
  }

  if (dropzone) {

    [
      "dragenter",
      "dragover"
    ].forEach(
      eventName => {

        dropzone.addEventListener(
          eventName,
          event => {

            event.preventDefault();

            event.stopPropagation();

            dropzone.classList.add(
              "dragover"
            );

          }
        );

      }
    );

    [
      "dragleave",
      "drop"
    ].forEach(
      eventName => {

        dropzone.addEventListener(
          eventName,
          event => {

            event.preventDefault();

            event.stopPropagation();

            dropzone.classList.remove(
              "dragover"
            );

          }
        );

      }
    );

    dropzone.addEventListener(
      "drop",
      event => {

        const files =
          Array.from(
            event.dataTransfer.files || []
          );


        const excelFiles =
          files.filter(
            file =>
              /\.(xlsx|xls)$/i.test(
                file.name
              )
          );


        if (
          excelFiles.length === 0
        ) {

          showToast(
            "File harus berupa Excel (.xlsx atau .xls).",
            "error"
          );

          return;
        }


        handleExcel(
          excelFiles
        );

      }
    );

  }
}

function setupFilters() {

  const search =
    $("searchInput");


  const year =
    $("yearFilter");


  const month =
    $("monthFilter");


  const type =
    $("typeFilter");


  const reset =
    $("resetFilters");


  if (search) {

    search.addEventListener(
      "input",
      () => {

        applyFilters();

      }
    );
  }


  if (year) {

    year.addEventListener(
      "change",
      () => {

        applyFilters();

      }
    );
  }


  if (month) {

    month.addEventListener(
      "change",
      () => {

        applyFilters();

      }
    );
  }


  if (type) {

    type.addEventListener(
      "change",
      () => {

        applyFilters();

      }
    );
  }


  if (reset) {

    reset.addEventListener(
      "click",
      () => {

        if (search) {
          search.value = "";
        }

        if (year) {
          year.value = "";
        }

        if (month) {
          month.value = "";
        }

        if (type) {
          type.value = "";
        }


        applyFilters();

        showToast(
          "Filter berhasil direset.",
          "success"
        );

      }
    );
  }
}


function setupTableActions() {

  const tbody =
    $("dataBody");


  if (tbody) {

    tbody.addEventListener(
      "click",
      event => {

        const editButton =
          event.target.closest(
            ".edit-btn"
          );


        const deleteButton =
          event.target.closest(
            ".delete-btn"
          );


        if (editButton) {

          const index =
            Number(
              editButton.dataset.index
            );


          editRow(index);

          return;
        }


        if (deleteButton) {

          const id =
            Number(
              deleteButton.dataset.id
            );

          deleteRow(id);

          return;
        }


        const importButton =
          event.target.closest(
            "#emptyImportBtn"
          );


        if (importButton) {

          $("excelInput")?.click();

        }

      }
    );
  }


  const form =
    $("editForm");


  if (form) {

    form.addEventListener(
      "submit",
      event => {

        event.preventDefault();

        saveEditedRow();

      }
    );
  }


  const closeButton =
    $("closeEditModal");


  const cancelButton =
    $("cancelEdit");


  closeButton?.addEventListener(
    "click",
    closeEditModal
  );


  cancelButton?.addEventListener(
    "click",
    closeEditModal
  );


  const modal =
    $("editModal");


  if (modal) {

    modal.addEventListener(
      "click",
      event => {

        if (
          event.target === modal
        ) {

          closeEditModal();

        }

      }
    );
  }


  document.addEventListener(
    "keydown",
    event => {

      if (
        event.key === "Escape"
      ) {

        closeEditModal();

      }

    }
  );
}

function setupLogo() {

  const logoInput =
    $("logoInput");


  if (!logoInput) {
    return;
  }


  logoInput.addEventListener(
    "change",
    event => {

      const file =
        event.target.files?.[0];


      if (!file) {
        return;
      }


      const reader =
        new FileReader();


      reader.onload =
        event => {

          const logo =
            event.target.result;


          localStorage.setItem(
            "berkasArsipLogo",
            logo
          );


          document
            .querySelectorAll(
              ".sidebar-logo img, .logo-wrap img"
            )
            .forEach(
              img => {

                img.src =
                  logo;

              }
            );

        };


      reader.readAsDataURL(
        file
      );

    }
  );
}

function bindEmptyImport() {

  const button =
    $("emptyImportBtn");


  if (!button) {
    return;
  }


  button.addEventListener(
    "click",
    () => {

      $("excelInput")?.click();

    }
  );
}

function showToast(
  message,
  type = "success"
) {

  const toast =
    $("toast");


  if (!toast) {
    return;
  }


  toast.textContent =
    message;


  toast.className =
    `toast ${type}`;


  toast.classList.add(
    "show"
  );


  clearTimeout(
    showToast.timer
  );


  showToast.timer =
    setTimeout(
      () => {

        toast.classList.remove(
          "show"
        );

      },
      3000
    );
}

function setupButtons() {

  const clearButton =
    $("clearDataBtn");


  const exportButton =
    $("exportBtn");


  clearButton?.addEventListener(
    "click",
    clearData
  );


  exportButton?.addEventListener(
    "click",
    exportData
  );
}

function loadSavedLogo() {

  const savedLogo =
    localStorage.getItem(
      "berkasArsipLogo"
    );


  if (!savedLogo) {
    return;
  }


  document
    .querySelectorAll(
      ".sidebar-logo img, .logo-wrap img"
    )
    .forEach(
      img => {

        img.src =
          savedLogo;

      }
    );
}

async function logout() {
  try {
    const { error } = await supabaseClient.auth.signOut();

    if (error) {
      console.error("Logout error:", error);
      showToast("Gagal logout: " + error.message, "error");
      return;
    }

    window.location.href = "login.html";

  } catch (error) {
    console.error("ERROR LOGOUT:", error);
    showToast("Terjadi kesalahan saat logout.", "error");
  }
}

function setupLogout() {
  const logoutBtn = $("logoutBtn");
  const logoutModal = $("logoutModal");
  const closeLogoutModal = $("closeLogoutModal");
  const cancelLogout = $("cancelLogout");
  const confirmLogout = $("confirmLogout");
  const overlay = document.querySelector(".logout-modal-overlay");

  if (!logoutBtn || !logoutModal) return;

  logoutBtn.addEventListener("click", () => {
    logoutModal.classList.remove("hidden");
  });

  closeLogoutModal?.addEventListener("click", () => {
    logoutModal.classList.add("hidden");
  });

  cancelLogout?.addEventListener("click", () => {
    logoutModal.classList.add("hidden");
  });

  overlay?.addEventListener("click", () => {
    logoutModal.classList.add("hidden");
  });

  confirmLogout?.addEventListener("click", async () => {
    await logout();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      logoutModal.classList.add("hidden");
    }
  });
}

async function initApp() {

  const session = await checkAuthentication();

  if (!session) return;

  setupManualAdd();
  setupLogout();
  setupDeleteModal();
  setupSingleDeleteModal();

  await loadDataFromSupabase();

  setupRealtime();
}

initApp();

function closeDeleteModal() {

  const modal = $("deleteModal");

  if (!modal) return;

  modal.classList.add("hidden");
}

document.addEventListener(
  "DOMContentLoaded",
  () => {

    console.log(
      "Berkas Arsip Marunda dimulai..."
    );

    setupNavigation();

    setupImport();

    setupFilters();

    setupTableActions();

    setupButtons();

    setupLogo();

    loadSavedLogo();

    updateStats();

    updateResultInfo();

    renderPagination();

    updateLastUpdate();

    console.log(
      "Aplikasi siap digunakan."
    );

  }
);