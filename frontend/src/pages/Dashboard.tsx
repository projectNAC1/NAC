import { useEffect, useState } from "react";
import { getSummary } from "../api/api";
import {
  PieChart,
  Pie,
  Cell,
  Tooltip,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  ResponsiveContainer,
} from "recharts";

type Summary = {
  total: number;
  comply: number;
  not_comply: number;
  no_xdr: number;
  no_boarding: number;
};

export default function Dashboard() {
  const [summary, setSummary] = useState<Summary | null>(null);

  useEffect(() => {
  getSummary()
    .then((data) => setSummary(data))
    .catch((err) => console.error(err));
}, []);

  if (!summary) return <div style={container}>Loading...</div>;

  // 🔥 DATA CHART
  const pieData = [
    { name: "Comply", value: Number(summary.comply) },
    { name: "Not Comply", value: Number(summary.not_comply) },
  ];

  const barData = [
    { name: "No XDR", value: Number(summary.no_xdr) },
    { name: "No Boarding", value: Number(summary.no_boarding) },
  ];

  return (
    <div style={container}>
      <h1 style={title}>📊 Dashboard Monitoring</h1>

      {/* KPI */}
      <div style={kpiContainer}>
        <Card title="Total Data" value={summary.total} color="#3b82f6" />
        <Card title="Comply" value={summary.comply} color="#22c55e" />
        <Card title="Not Comply" value={summary.not_comply} color="#ef4444" />
        <Card title="No Install XDR" value={summary.no_xdr} color="#f97316" />
        <Card title="No Boarding" value={summary.no_boarding} color="#eab308" />
      </div>

      {/* CHART */}
      <div style={chartGrid}>
        {/* PIE */}
        <div style={chartCard}>
          <h3>Compliance</h3>
          <ResponsiveContainer width="100%" height={250}>
            <PieChart>
              <Pie
                data={pieData}
                dataKey="value"
                outerRadius={90}
                label
              >
                <Cell fill="#22c55e" />
                <Cell fill="#ef4444" />
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </div>

        {/* BAR */}
        <div style={chartCard}>
          <h3>Issues</h3>
          <ResponsiveContainer width="100%" height={250}>
            <BarChart data={barData}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" stroke="#fff" />
              <YAxis stroke="#fff" />
              <Tooltip />
              <Bar dataKey="value">
                <Cell fill="#f97316" />
                <Cell fill="#eab308" />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ================= COMPONENT ================= */

function Card({ title, value, color }: any) {
  return (
    <div style={{ ...card, borderLeft: `5px solid ${color}` }}>
      <p style={{ opacity: 0.7 }}>{title}</p>
      <h2>{value}</h2>
    </div>
  );
}

/* ================= STYLE ================= */

const container: React.CSSProperties = {
  padding: 20,
  background: "#0f172a",
  minHeight: "100vh",
  color: "#fff",
};

const title: React.CSSProperties = {
  marginBottom: 20,
};

const kpiContainer: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
  gap: 20,
  marginBottom: 30,
};

const card: React.CSSProperties = {
  background: "#1e293b",
  padding: 20,
  borderRadius: 12,
};

const chartGrid: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))",
  gap: 20,
};

const chartCard: React.CSSProperties = {
  background: "#1e293b",
  padding: 20,
  borderRadius: 12,
};