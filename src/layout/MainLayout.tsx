import Sidebar from "../components/Sidebar";

export default function MainLayout({ children }: any) {
  return (
    <div style={{ display: "flex" }}>
      {/* SIDEBAR FIX WIDTH */}
      <div style={{ width: 240, flexShrink: 0 }}>
        <Sidebar />
      </div>

      {/* CONTENT */}
      <div
        style={{
          flex: 1,
          padding: 20,
          overflowX: "hidden", // 🔥 penting
        }}
      >
        {children}
      </div>
    </div>
  );
}