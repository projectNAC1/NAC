import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { BrowserRouter, Routes, Route } from "react-router-dom";

import "./index.css";
import App from "./App"; // ⬅️ ini tetap

import MainLayout from "./layout/MainLayout";
import Dashboard from "./pages/Dashboard";
import Filter from "./pages/Filter";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <BrowserRouter>
      <MainLayout>
        <Routes>
          <Route path="/" element={<Dashboard />} />

          {/* 🔥 INI PENTING */}
          <Route path="/generate" element={<App />} />

          <Route path="/filter" element={<Filter />} />
        </Routes>
      </MainLayout>
    </BrowserRouter>
  </StrictMode>
);