const BASE_URL = "http://localhost:8000/api";

// 🔥 helper fetch
async function fetchAPI(endpoint: string) {
  const res = await fetch(`${BASE_URL}${endpoint}`);

  if (!res.ok) {
    throw new Error(`API error: ${res.status}`);
  }

  return res.json();
}

// ================= ENDPOINT =================

// 🔥 summary
export async function getSummary() {
  const data = await fetchAPI("/summary");

  return {
    total: Number(data.total),
    comply: Number(data.comply),
    not_comply: Number(data.not_comply),
    no_xdr: Number(data.no_xdr),
    no_boarding: Number(data.no_boarding),
  };
}

// 🔥 master data (pagination)
export async function getMasterData(
  page: number,
  startDate?: string,
  endDate?: string
) {
  let url = `/master-data?page=${page}`;

  if (startDate) {
    url += `&start=${startDate}`;
  }

  if (endDate) {
    url += `&end=${endDate}`;
  }

  return fetchAPI(url);
}

export async function exportMasterData(startDate?: string, endDate?: string) {
  let url = `/export`;

  if (startDate) url += `?start=${startDate}`;
  if (endDate) url += `${startDate ? "&" : "?"}end=${endDate}`;

  return fetchAPI(url);
}