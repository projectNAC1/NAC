const express = require("express");
const { Pool } = require("pg");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.json());

const pool = new Pool({
  host: "10.10.90.10",
  user: "nac",
  password: "P@ssw0rd!",
  database: "postgres",
  port: 5432,
});

// 🔥 test koneksi
pool.query("SELECT NOW()")
  .then(() => console.log("✅ DB Connected"))
  .catch(err => console.error("❌ DB Error:", err));

// 🔥 helper normalize key (ANTI mismatch)
function normalizeRow(row) {
  const newRow = {};
  Object.keys(row).forEach(key => {
    const cleanKey = key
      .toLowerCase()
      .replace(/\s+/g, "");
    newRow[cleanKey] = row[key];
  });
  return newRow;
}

// 🔥 route test
app.get("/", (req, res) => {
  res.send("API is running 🚀");
});

//count total
app.get("/api/summary", async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT 
        COUNT(*) AS total,

        SUM(CASE 
          WHEN LOWER(reason) = 'comply' THEN 1 ELSE 0 
        END) AS comply,

        SUM(CASE 
          WHEN LOWER(reason) != 'comply' THEN 1 ELSE 0 
        END) AS not_comply,

        SUM(CASE 
          WHEN LOWER(reason) LIKE '%xdr%' THEN 1 ELSE 0 
        END) AS no_xdr,

        SUM(CASE 
          WHEN LOWER(reason) LIKE '%boarding%' THEN 1 ELSE 0 
        END) AS no_boarding

      FROM master
    `);

    res.json(result.rows[0]);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to fetch summary" });
  }
});

app.post("/api/master-data", async (req, res) => {
  const rows = req.body.data;

  if (!rows || !Array.isArray(rows)) {
    return res.status(400).json({ error: "Invalid data" });
  }

  const client = await pool.connect();

  try {
    await client.query("BEGIN");

    const values = [];
    const placeholders = [];
    const COLS = 16;

    rows.forEach((row, i) => {
      const clean = normalizeRow(row); // 🔥 NORMALIZE

      if (i === 0) {
        console.log("==== RAW ROW ====");
        console.log(row);

        console.log("==== NORMALIZED ====");
        console.log(clean);
      }

      const base = i * COLS;

      placeholders.push(`(
        $${base + 1}, $${base + 2}, $${base + 3}, $${base + 4}, $${base + 5},
        $${base + 6}, $${base + 7}, $${base + 8}, $${base + 9},
        $${base + 10}, $${base + 11}, $${base + 12},
        $${base + 13}, $${base + 14}, $${base + 15}, $${base + 16}
      )`);
console.log("POSTURING:", row["Posturing+Boarding"]);
        values.push(
        // row ["no"] ?? null,
        row["Fullname"] ?? null,
        getValue(row, ["PN"]),
        getValue(row, ["Divisi"]),
        getValue(row, ["Lokasi"]),
        getValue(row, ["Lantai"]),
        row["Host Name"] ?? null,
        getValue(row, ["Serial Number", "SN"]),
        getValue(row, ["MAC Address", "MAC"]),
        getValue(row, ["Device Type"]),
        getValue(row, ["Check Boarding"]),
        getValue(row, ["Check Boarding on Manual"]),
        getValue(row, ["Posturing+Boarding"]),
        // row["Posturing+Boarding"] ?? null,
        getValue(row, ["Divisi by Grouping"]),
        getValue(row, ["XDR"]),
        getValue(row, ["Reason Not Comply"]),

        new Date()
        );
    });

    // 🔥 VALIDASI
    console.log("VALUES SAMPLE:", values.slice(0, 16));
    console.log("TOTAL VALUES:", values.length);
    console.log("EXPECTED:", rows.length * COLS);

    await client.query(`
      INSERT INTO master (
        fullname, pn, divisi, lokasi, lantai,
        host_name, sn, mac, device_type,
        boarding, boarding_manual, posturing_boarding,
        div_by_group, xdr, reason, date
      ) VALUES ${placeholders.join(",")}
    `, values);

    await client.query("COMMIT");

    res.json({
      status: "success",
      inserted: rows.length,
    });

  } catch (error) {
    await client.query("ROLLBACK");
    console.error("❌ INSERT ERROR:", error);
    res.status(500).json({ error: "Insert failed" });
  } finally {
    client.release();
  }
});

app.get("/api/master-data", async (req, res) => {
  const page = parseInt(req.query.page) || 1;
  const limit = 100;
  const offset = (page - 1) * limit;

  const { start, end } = req.query;

  let where = [];
  let values = [];

  // 🔥 FILTER DATE
  if (start) {
    values.push(start);
    where.push(`date >= $${values.length}`);
  }

  if (end) {
    values.push(end);
    where.push(`date <= $${values.length}`);
  }

  const whereClause = where.length ? `WHERE ${where.join(" AND ")}` : "";

  try {
    // 🔥 QUERY DATA
    const dataQuery = `
      SELECT * FROM master
      ${whereClause}
      ORDER BY no ASC
      LIMIT $${values.length + 1} OFFSET $${values.length + 2}
    `;

    // 🔥 QUERY TOTAL (BIAR PAGE SESUAI FILTER)
    const countQuery = `
      SELECT COUNT(*) FROM master ${whereClause}
    `;

    const dataValues = [...values, limit, offset];

    const dataResult = await pool.query(dataQuery, dataValues);
    const countResult = await pool.query(countQuery, values);

    const total = parseInt(countResult.rows[0].count);
    const totalPages = Math.ceil(total / limit);

    res.json({
      data: dataResult.rows,
      totalPages,
    });

  } catch (err) {
    console.error("ERROR:", err);
    res.status(500).json({ error: "Failed to fetch data" });
  }
});

function getValue(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== "") {
      return row[key];
    }
  }
  return null;
}

app.get("/api/export", async (req, res) => {
  const { start, end } = req.query;

  let where = [];
  let values = [];

  if (start) {
    values.push(start);
    where.push(`date >= $${values.length}`);
  }

  if (end) {
    values.push(end);
    where.push(`date <= $${values.length}`);
  }

  const whereClause = where.length ? `WHERE ${where.join(" AND ")}` : "";

  try {
    const result = await pool.query(`
      SELECT * FROM master
      ${whereClause}
      ORDER BY no ASC
    `, values);

    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Export failed" });
  }
});

app.listen(8000, "0.0.0.0", () => {
  console.log("🚀 Server running on http://0.0.0.0:8000");
});