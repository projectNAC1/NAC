import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

export default function Filter() {
  const [data, setData] = useState<any[]>([]);
  const [date, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [totalPages, setTotalPages] = useState(1);
  const [loading, setLoading] = useState(false);

  // 🔥 EXPORT EXCEL
  const exportToExcel = () => {
    if (filteredData.length === 0) {
      alert("Tidak ada data untuk di export");
      return;
    }

    const exportData = filteredData.map((item) => ({
      No: item.no,
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

    saveAs(file, `master_data_${date || "all"}.xlsx`);
  };

  // 🔥 FETCH DATA
  useEffect(() => {
    setLoading(true);

    fetch(`http://localhost:8000/api/master-data?page=${page}`)
      .then(res => res.json())
      .then(res => {
        if (Array.isArray(res)) {
          setData(res);
          setTotalPages(1);
        } else if (res.data) {
          setData(res.data);
          setTotalPages(res.totalPages || 1);
        } else {
          setData([]);
        }
      })
      .catch(err => {
        console.error("FETCH ERROR:", err);
        setData([]);
      })
      .finally(() => setLoading(false));
  }, [page]);

  // 🔥 FILTER
  const filteredData = data.filter(item => {
    const keyword = date.toLowerCase();

    return (
      item.fullname?.toLowerCase().includes(keyword) ||
      item.pn?.toLowerCase().includes(keyword) ||
      item.divisi?.toLowerCase().includes(keyword) ||
      item.lokasi?.toLowerCase().includes(keyword) ||
      item.device_type?.toLowerCase().includes(keyword) ||
      formatDate(item.date).toLowerCase().includes(keyword)
    );
  });

  return (
    <div style={styles.container}>
      <h1 style={styles.title}>📊 Master Data</h1>

      {/* 🔍 SEARCH + EXPORT */}
      <div style={{ display: "flex", gap: "10px", marginBottom: "20px" }}>
        <input
          style={styles.input}
          placeholder="Cari nama / PN / divisi / tanggal..."
          value={date}
          onChange={(e) => setSearch(e.target.value)}
        />

        <button
          onClick={exportToExcel}
          disabled={filteredData.length === 0}
          style={{
            background: filteredData.length === 0 ? "#999" : "green",
            color: "#fff",
            border: "none",
            padding: "10px 16px",
            borderRadius: "8px",
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
              <th>No</th>
              <th>Fullname</th>
              <th>PN</th>
              <th>Divisi</th>
              <th>Lokasi</th>
              <th>Lantai</th>
              <th>Host</th>
              <th>SN</th>
              <th>MAC</th>
              <th>Device</th>
              <th>Boarding</th>
              <th>Manual</th>
              <th>Posturing</th>
              <th>Group</th>
              <th>XDR</th>
              <th>Reason</th>
              <th>Date</th>
            </tr>
          </thead>

          <tbody>
            {filteredData.length > 0 ? (
              filteredData.map((item) => (
                <tr key={item.no}>
                  <td>{item.no}</td>
                  <td>{item.fullname || "-"}</td>
                  <td>{item.pn || "-"}</td>
                  <td>{item.divisi || "-"}</td>
                  <td>{item.lokasi || "-"}</td>
                  <td>{item.lantai || "-"}</td>
                  <td>{item.host_name || "-"}</td>
                  <td>{item.sn || "-"}</td>
                  <td>{item.mac || "-"}</td>
                  <td>{item.device_type || "-"}</td>
                  <td>{item.boarding || "-"}</td>
                  <td>{item.boarding_manual || "-"}</td>

                  <td style={getStatusStyle(item.posturing_boarding)}>
                    {item.posturing_boarding || "-"}
                  </td>

                  <td>{item.div_by_group || "-"}</td>
                  <td>{item.xdr || "-"}</td>
                  <td>{item.reason || "-"}</td>
                  <td>{formatDate(item.date)}</td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan={17} style={{ textAlign: "center" }}>
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

        <span>
          Page {page} / {totalPages}
        </span>

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
  },
  title: {
    marginBottom: "20px",
  },
  input: {
    padding: "10px",
    width: "350px",
    borderRadius: "8px",
    border: "1px solid #ccc",
  },
  tableWrapper: {
    width: "100%",
    overflowX: "scroll" as const,
    background: "#fff",
    borderRadius: "10px",
    padding: "10px",
  },
  table: {
    width: "max-content",
    minWidth: "100%",
    borderCollapse: "collapse" as const,
  },
  thead: {
    position: "sticky" as const,
    top: 0,
    background: "#eee",
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