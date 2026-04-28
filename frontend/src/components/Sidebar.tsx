import { Link, useLocation } from "react-router-dom";

export default function Sidebar() {
  const location = useLocation();

  const menu = [
    { name: "Dashboard", path: "/" },
    { name: "Generate", path: "/generate" },
    { name: "Filter", path: "/filter" },
  ];

  return (
    <div style={styles.sidebar}>
      <h2 style={{ color: "#fff" }}>NAC</h2>

      {menu.map((item) => (
        <Link
          key={item.path}
          to={item.path}
          style={{
            ...styles.link,
            background:
              location.pathname === item.path ? "#2563eb" : "transparent",
          }}
        >
          {item.name}
        </Link>
      ))}
    </div>
  );
}

const styles: any = {
  sidebar: {
  width: 240,
  height: "100vh",
  background: "#0f172a",
  padding: 20,
  position: "fixed", // 🔥 biar ga ketiban
  left: 0,
  top: 0,
},
  link: {
    display: "block",
    padding: 12,
    marginTop: 10,
    borderRadius: 8,
    color: "#fff",
    textDecoration: "none",
  },
};