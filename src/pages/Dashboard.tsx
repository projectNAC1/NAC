export default function Dashboard() {
  return (
    <div>
      <h1>Dashboard</h1>

      <div style={{ display: "flex", gap: 20 }}>
        <div style={card}>Total Data</div>
        <div style={card}>Success</div>
        <div style={card}>Checkleak</div>
      </div>
    </div>
  );
}

const card = {
  padding: 20,
  background: "#fff",
  borderRadius: 10,
};