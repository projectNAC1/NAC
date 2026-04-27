import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { getMasterData, exportMasterData } from "../api/api";

export default function Filter() {
  const [data, setData] = useState<any[]>([]);
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [page, setPage] = useState(1);
  const [totalPages, setTotalPages] = useState(1);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    setPage(1);
  }, [startDate, endDate]);

  // 🔥 FETCH DATA
  useEffect(() => {
  setLoading(true);

  getMasterData(page, startDate, endDate)
    .then((res) => {
      setData(res.data);
      setTotalPages(res.totalPages);
    })
    .catch((err) => {
      console.error("FETCH ERROR:", err);
      setData([]);
    })
    .finally(() => setLoading(false));
}, [page, startDate, endDate]);

  // 🔥 EXPORT EXCEL
  const exportToExcel = async () => {
  try {
    const allData = await exportMasterData(startDate, endDate);

    if (allData.length === 0) {
      alert("Tidak ada data untuk di export");
      return;
    }

    const exportData = allData.map((item: any, i: number) => ({
      No: i + 1,
      Fullname: item.fullname,
      PN: item.pn,
      Divisi: item.divisi,
      Lokasi: item.lokasi,
      Lantai: item.lantai,
      Host: item.host_name,
      SN: item.sn,
      MAC: item.mac,
      Device: item.device_type,
      Boarding: item.boarding,
      Manual: item.boarding_manual,
      Posturing: item.posturing_boarding,
      Group: item.div_by_group,
      XDR: item.xdr,
      Reason: item.reason,
      Date: formatDate(item.date),
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Master Data");

    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    const file = new Blob([excelBuffer], {
      type: "application/octet-stream",
    });

    saveAs(
      file,
      `master_data_${startDate || "all"}_${endDate || "all"}.xlsx`
    );

  } catch (err) {
    console.error("EXPORT ERROR:", err);
  }
};

  // 🔥 FILTER
  const filteredData = data.filter((item) => {
    if (!startDate && !endDate) return true;

    const itemDate = new Date(item.date);
    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;

    if (start && itemDate < start) return false;
    if (end && itemDate > end) return false;

    return true;
  });

  return (
    <div style={styles.container}>
      <h1 style={styles.title}>📊 Master Data</h1>

      {/* 🔍 SEARCH + EXPORT */}
      <div style={styles.topBar}>
        <input
          type="date"
          value={startDate}
          onChange={(e) => setStartDate(e.target.value)}
          style={styles.input}
        />

        <input
          type="date"
          value={endDate}
          onChange={(e) => setEndDate(e.target.value)}
          style={styles.input}
        />

        <button
          onClick={exportToExcel}
          disabled={loading}
          style={{
            ...styles.exportBtn,
            background: filteredData.length === 0 ? "#999" : "#16a34a",
            cursor: filteredData.length === 0 ? "not-allowed" : "pointer"
          }}
        >
          ⬇ Export Excel
        </button>
      </div>

      {/* LOADING */}
      {loading && <p>Loading data...</p>}

      {/* TABLE */}
      <div style={styles.tableWrapper}>
        <table style={styles.table}>
          <thead style={styles.thead}>
            <tr>
              {headers.map((h) => (
                <th key={h} style={styles.cellHeader}>{h}</th>
              ))}
            </tr>
          </thead>

          <tbody>
            {filteredData.length > 0 ? (
              filteredData.map((item, index) => (
                <tr
                  key={item.no}
                  style={index % 2 === 0 ? styles.row : styles.rowAlt}
                >
                  <td style={styles.cell}>{index + 1}</td>
                  <td style={styles.cell}>{item.fullname || "-"}</td>
                  <td style={styles.cell}>{item.pn || "-"}</td>
                  <td style={styles.cell}>{item.divisi || "-"}</td>
                  <td style={styles.cell}>{item.lokasi || "-"}</td>
                  <td style={styles.cell}>{item.lantai || "-"}</td>
                  <td style={styles.cell}>{item.host_name || "-"}</td>
                  <td style={styles.cell}>{item.sn || "-"}</td>
                  <td style={styles.cell}>{item.mac || "-"}</td>
                  <td style={styles.cell}>{item.device_type || "-"}</td>
                  <td style={styles.cell}>{item.boarding || "-"}</td>
                  <td style={styles.cell}>{item.boarding_manual || "-"}</td>

                  <td style={{
                    ...styles.cell,
                    ...getStatusStyle(item.posturing_boarding)
                  }}>
                    {item.posturing_boarding || "-"}
                  </td>

                  <td style={styles.cell}>{item.div_by_group || "-"}</td>
                  <td style={styles.cell}>{item.xdr || "-"}</td>
                  <td style={styles.cell}>{item.reason || "-"}</td>
                  <td style={styles.cell}>{formatDate(item.date)}</td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan={17} style={{ textAlign: "center", padding: 20 }}>
                  Tidak ada data
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* PAGINATION */}
      <div style={styles.pagination}>
        <button
          disabled={page <= 1}
          onClick={() => setPage(prev => prev - 1)}
          style={styles.button}
        >
          ⬅ Prev
        </button>

        <span>Page {page} / {totalPages}</span>

        <button
          disabled={page >= totalPages}
          onClick={() => setPage(prev => prev + 1)}
          style={styles.button}
        >
          Next ➡
        </button>
      </div>
    </div>
  );
}

// 🔥 HEADER LIST
const headers = [
  "No","Fullname","PN","Divisi","Lokasi","Lantai","Host","SN","MAC",
  "Device","Boarding","Manual","Posturing","Group","XDR","Reason","Date"
];

// 🔥 FORMAT DATE
function formatDate(date: any) {
  if (!date) return "-";
  try {
    return new Date(date).toLocaleDateString("id-ID");
  } catch {
    return date;
  }
}

// 🔥 STATUS COLOR
function getStatusStyle(status: string) {
  return {
    color: status === "No Boarding" ? "red" : "green",
    fontWeight: "bold"
  };
}

// 🎨 STYLE
const styles = {
  container: {
    padding: "30px",
    fontFamily: "Arial",
    background: "#f3f4f6",
    minHeight: "100vh"
  },
  title: {
    marginBottom: "20px",
  },
  topBar: {
    display: "flex",
    gap: "10px",
    marginBottom: "20px"
  },
  input: {
    padding: "10px",
    width: "350px",
    borderRadius: "8px",
    border: "1px solid #ccc",
  },
  exportBtn: {
    color: "#fff",
    border: "none",
    padding: "10px 16px",
    borderRadius: "8px",
  },
  tableWrapper: {
    width: "100%",
    overflowX: "auto" as const,
    background: "#fff",
    borderRadius: "10px",
    padding: "10px",
    boxShadow: "0 4px 10px rgba(0,0,0,0.05)"
  },
  table: {
    width: "100%",
    borderCollapse: "collapse" as const,
    fontSize: "14px",
  },
  thead: {
    position: "sticky" as const,
    top: 0,
    background: "#3b82f6",
  },
  cellHeader: {
    padding: "12px",
    color: "#fff",
    textAlign: "center" as const,
  },
  cell: {
    padding: "10px 14px",
    textAlign: "center" as const,
    borderBottom: "1px solid #eee",
    whiteSpace: "nowrap" as const,
  },
  row: {
    backgroundColor: "#fff",
  },
  rowAlt: {
    backgroundColor: "#f9fafb",
  },
  pagination: {
    marginTop: "20px",
    display: "flex",
    justifyContent: "center",
    gap: "10px",
  },
  button: {
    padding: "8px 16px",
    borderRadius: "6px",
    border: "none",
    background: "#3b82f6",
    color: "#fff",
    cursor: "pointer",
  },
};