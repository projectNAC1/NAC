import { useMemo, useState } from "react";
import type { CSSProperties, ChangeEvent } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

type UploadKey =
  | "rollout"
  | "boarding"
  | "software"
  | "checkleak"
  | "deviceList";

type FileState = Record<UploadKey, File | null>;
type GenericRow = Record<string, any>;

type SheetRows = {
  sheetName: string;
  rows: GenericRow[];
};

type MasterKeyRecord = {
  key: string;
  serial: string;
  mac: string; // compare format: AABBCCDDEEFF
};

type CheckleakStatus = {
  xdrInstalled: boolean;
  kaseyaInstalled: boolean;
};

type MergeContext = {
  rowNumber: number;
  keyRecord: MasterKeyRecord;
  rolloutRow: GenericRow | null;
  boardingMatch: GenericRow | null;
  softwareMatch: GenericRow | null;
  checkleakMatch: GenericRow | null;
  deviceMatchByMac: GenericRow | null;
  matchedDeviceBySerial: GenericRow | null;
  serialValue: string;
  candidateMacs: string[];
  boardingUsernameByMac: Map<string, string>;
  checkleakStatusByMac: Map<string, CheckleakStatus>;
  rolloutPnValue: string;
};

const FILE_KEYS = {
  rollout: "rollout",
  boarding: "boarding",
  software: "software",
  checkleak: "checkleak",
  deviceList: "deviceList",
} as const;

const REQUIRED_FILES: Array<{ key: UploadKey; label: string }> = [
  { key: FILE_KEYS.rollout, label: "Rollout File (.xls/.xlsx)" },
  { key: FILE_KEYS.boarding, label: "Boarding Apply Logs (.xls/.xlsx)" },
  {
    key: FILE_KEYS.software,
    label: "Software Installation Report (.xls/.xlsx)",
  },
  {
    key: FILE_KEYS.checkleak,
    label: "System Checkleak Report (.xls/.xlsx, multi-sheet)",
  },
  { key: FILE_KEYS.deviceList, label: "Device List (.xls/.xlsx)" },
];

const EMPTY_FILES: FileState = {
  rollout: null,
  boarding: null,
  software: null,
  checkleak: null,
  deviceList: null,
};

export default function App() {
  const [files, setFiles] = useState<FileState>(EMPTY_FILES);
  const [status, setStatus] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);

  const readyCount = useMemo(
    () => Object.values(files).filter(Boolean).length,
    [files],
  );

  const handleFileChange = (
    key: UploadKey,
    event: ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.target.files?.[0] ?? null;
    setFiles((prev) => ({ ...prev, [key]: file }));
  };

  const generateMaster = async () => {
    if (isProcessing) return;

    try {
      setIsProcessing(true);
      setStatus("Loading built-in master template...");

      for (const item of REQUIRED_FILES) {
        if (!files[item.key]) {
          throw new Error(`Missing file: ${item.label}`);
        }
      }

      const masterTemplateWb = await loadTemplateWorkbook(
        "/template/master.xlsx",
      );

      setStatus("Reading uploaded files...");
      await pauseUi();

      const [
        rolloutSheets,
        boardingSheets,
        softwareSheets,
        checkleakSheets,
        deviceListSheets,
      ] = await Promise.all([
        readAllSheets(files.rollout!),
        readAllSheets(files.boarding!),
        readAllSheets(files.software!),
        readAllSheets(files.checkleak!),
        readAllSheets(files.deviceList!),
      ]);

      const masterSheetName = masterTemplateWb.SheetNames[0];
      const masterSheet = masterTemplateWb.Sheets[masterSheetName];
      const masterRowsRaw = XLSX.utils.sheet_to_json(masterSheet, {
        defval: "",
      });

      if (!masterRowsRaw.length) {
        throw new Error("Built-in master template is empty.");
      }

      const masterHeaders = extractHeaders(masterRowsRaw);

      setStatus("Preparing source data...");
      await pauseUi();

      const rolloutRows = flattenSheetRows(rolloutSheets);
      const boardingRows = flattenSheetRows(boardingSheets);
      const softwareRows = flattenSheetRows(softwareSheets);
      const checkleakRows = flattenSheetRows(checkleakSheets);
      const deviceListRows = flattenSheetRows(deviceListSheets);

      const rolloutSerialKey = detectColumn(rolloutRows, [
        "serial number",
        "serialnumber",
        "serial number laptop / merk handphone",
        "sn",
        "service tag",
        "asset serial",
        "bios serial number",
      ]);

      const rolloutMacKey = detectColumn(rolloutRows, [
        "mac address",
        "mac address laptop / handphone",
        "mac",
        "wifi mac",
        "wireless mac",
        "lan mac",
      ]);

      const rolloutPnKey = detectColumn(rolloutRows, [
        "pn",
        "personal number",
        "personnel number",
        "nik",
        "employee id",
        "employee number",
      ]);

      const boardingMacKey = detectColumn(boardingRows, [
        "mac address of the terminal",
        "mac address",
        "mac",
        "wifi mac",
        "wireless mac",
        "lan mac",
      ]);

      const softwareMacKey = detectColumn(softwareRows, [
        "mac address",
        "mac",
        "wifi mac",
        "wireless mac",
        "lan mac",
      ]);

      const checkleakMacKey = detectColumn(checkleakRows, [
        "mac address",
        "mac",
        "wifi mac",
        "wireless mac",
        "lan mac",
      ]);

      const deviceListMacKey = detectColumn(deviceListRows, [
        "mac address",
        "mac",
        "wifi mac",
        "wireless mac",
        "lan mac",
      ]);

      const deviceListSerialKey = detectColumn(deviceListRows, [
        "bios serial number",
        "serial number",
        "serialnumber",
        "sn",
        "service tag",
        "asset serial",
      ]);

      setStatus("Indexing rows for matching...");
      await pauseUi();

      const rolloutIndexBySerial = indexRowsByKeys(
        rolloutRows,
        [rolloutSerialKey],
        {
          normalizeMacKeys: false,
        },
      );

      const rolloutIndexByMac = indexRowsByKeys(rolloutRows, [rolloutMacKey], {
        normalizeMacKeys: true,
      });

      const boardingIndexByMac = indexRowsByKeys(
        boardingRows,
        [boardingMacKey],
        {
          normalizeMacKeys: true,
        },
      );

      const softwareIndexByMac = indexRowsByKeys(
        softwareRows,
        [softwareMacKey],
        {
          normalizeMacKeys: true,
        },
      );

      const checkleakIndexByMac = indexRowsByKeys(
        checkleakRows,
        [checkleakMacKey],
        {
          normalizeMacKeys: true,
        },
      );

      const checkleakStatusByMac = buildCheckleakStatusIndex(
        checkleakRows,
        checkleakMacKey,
      );

      const deviceListIndexByMac = indexRowsByKeys(
        deviceListRows,
        [deviceListMacKey],
        {
          normalizeMacKeys: true,
        },
      );

      const deviceListIndexBySerial = indexRowsByKeys(
        deviceListRows,
        [deviceListSerialKey],
        {
          normalizeMacKeys: false,
        },
      );

      const boardingUsernameByMac = buildBoardingUsernameIndex(boardingRows);

      setStatus("Building union of all keys...");
      await pauseUi();

      const masterKeys = buildMasterKeys({
        rolloutRows,
        boardingRows,
        softwareRows,
        checkleakRows,
        deviceListRows,
        rolloutSerialKey,
        rolloutMacKey,
        boardingMacKey,
        softwareMacKey,
        checkleakMacKey,
        deviceListSerialKey,
        deviceListMacKey,
      });

      if (!masterKeys.length) {
        throw new Error("No usable rows found in uploaded files.");
      }

      setStatus(`Combining ${masterKeys.length} rows into master.xlsx...`);
      await pauseUi();

      const combinedRows: GenericRow[] = [];
      const chunkSize = 200;

      for (let start = 0; start < masterKeys.length; start += chunkSize) {
        const end = Math.min(start + chunkSize, masterKeys.length);

        for (let index = start; index < end; index += 1) {
          const keyRecord = masterKeys[index];

          const matchedDeviceBySerial =
            keyRecord.serial && deviceListIndexBySerial.has(keyRecord.serial)
              ? (deviceListIndexBySerial.get(keyRecord.serial) ?? null)
              : null;

          const rolloutBySerial =
            keyRecord.serial && rolloutIndexBySerial.has(keyRecord.serial)
              ? (rolloutIndexBySerial.get(keyRecord.serial) ?? null)
              : null;

          const rolloutByMac =
            keyRecord.mac && rolloutIndexByMac.has(keyRecord.mac)
              ? (rolloutIndexByMac.get(keyRecord.mac) ?? null)
              : null;

          const rolloutRow = rolloutBySerial || rolloutByMac || null;

          const candidateMacs = dedupeStrings([
            keyRecord.mac,
            ...collectPossibleMacs([
              rolloutRow,
              matchedDeviceBySerial,
              rolloutByMac,
              rolloutBySerial,
            ]),
          ]);

          const boardingMatch = findFirstMacMatch(
            candidateMacs,
            boardingIndexByMac,
          );
          const softwareMatch = findFirstMacMatch(
            candidateMacs,
            softwareIndexByMac,
          );
          const checkleakMatch = findFirstMacMatch(
            candidateMacs,
            checkleakIndexByMac,
          );
          const deviceMatchByMac = findFirstMacMatch(
            candidateMacs,
            deviceListIndexByMac,
          );

          const serialValue =
            keyRecord.serial ||
            normalizeText(
              getCell(rolloutRow, rolloutSerialKey) ||
                getCell(matchedDeviceBySerial, deviceListSerialKey),
            );

          const rolloutPnValue = normalizeText(
            getValueFromRow(rolloutRow, [
              rolloutPnKey || "",
              "pn",
              "personal number",
              "personnel number",
              "nik",
              "employee id",
              "employee number",
            ]),
          );

          const merged: GenericRow = {};

          for (const header of masterHeaders) {
            merged[header] = pickValueForMasterHeader(header, {
              rowNumber: index + 1,
              keyRecord,
              rolloutRow,
              boardingMatch,
              softwareMatch,
              checkleakMatch,
              deviceMatchByMac,
              matchedDeviceBySerial,
              serialValue,
              candidateMacs,
              boardingUsernameByMac,
              checkleakStatusByMac,
              rolloutPnValue,
            });
          }

          for (const key of Object.keys(merged)) {
            merged[key] = withNA(merged[key]);
          }

          combinedRows.push(merged);
        }

        setStatus(`Combining rows... ${end}/${masterKeys.length}`);
        await pauseUi();
      }

      setStatus("Sorting rows...");
      await pauseUi();

      const sortedRows = sortRowsByQuality(combinedRows);
      sortedRows.forEach((row, index) => {
        row["No"] = index + 1;
      });

      setStatus("Writing Excel file...");
      await pauseUi();

      const outputWb = XLSX.utils.book_new();
      const outputWs = XLSX.utils.json_to_sheet(sortedRows, {
        header: masterHeaders,
      });

      XLSX.utils.book_append_sheet(outputWb, outputWs, "Master");

      const buffer = XLSX.write(outputWb, {
        bookType: "xlsx",
        type: "array",
      });

      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      saveAs(blob, "master.xlsx");
      setStatus(`Done. Exported ${combinedRows.length} rows to master.xlsx`);
    } catch (error) {
      console.error(error);
      setStatus(
        error instanceof Error
          ? error.message
          : "Failed to generate master file.",
      );
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        <h1 style={styles.title}>XLSX Master Merger</h1>

        <p style={styles.subtitle}>
          Built-in template version.
          <br />
          User only uploads 5 source files.
          <br />
          Priority:
          <br />
          1. Device List
          <br />
          2. Software Information
          <br />
          3. System Checkleak
          <br />
          4. Boarding
          <br />
          5. Rollout
          <br />
          <br />
          Rules:
          <br />
          - Master rows built from all files
          <br />
          - Serial Number first, MAC fallback
          <br />
          - MAC matching ignores ":" and "-"
          <br />
          - Output MAC uses ":"
          <br />
          - PN priority: Boarding username first, Rollout personal number
          fallback
          <br />
          - Check Boarding from Boarding MAC
          <br />
          - Check Boarding on Manual from Rollout PN
          <br />
          - Posturing+Boarding depends on those 2 columns
          <br />
          - Divisi by Grouping from Software Installation Department Name
          <br />
          - XDR / KASEYA from Checkleak Vulnerability description
          <br />- Empty values become #N/A
        </p>

        <div style={styles.notice}>
          Built-in template path: <code>public/template/master.xlsx</code>
        </div>

        <div style={styles.grid}>
          {REQUIRED_FILES.map((item) => (
            <div key={item.key} style={styles.fileBox}>
              <label style={styles.label}>{item.label}</label>
              <input
                type="file"
                accept=".xls,.xlsx"
                onChange={(e) => handleFileChange(item.key, e)}
                disabled={isProcessing}
              />
              <div style={styles.fileName}>
                {files[item.key] ? files[item.key]!.name : "No file selected"}
              </div>
            </div>
          ))}
        </div>

        <div style={styles.actions}>
          <button
            style={styles.button}
            onClick={generateMaster}
            disabled={isProcessing}>
            {isProcessing ? "Processing..." : "Generate master.xlsx"}
          </button>
          <div style={styles.counter}>
            {readyCount} / {REQUIRED_FILES.length} files selected
          </div>
        </div>

        {status && <div style={styles.status}>{status}</div>}
      </div>
    </div>
  );
}

async function pauseUi() {
  await new Promise<void>((resolve) => {
    setTimeout(() => resolve(), 0);
  });
}

async function loadTemplateWorkbook(templatePath: string) {
  const response = await fetch(templatePath);

  if (!response.ok) {
    throw new Error(
      `Cannot load template file at ${templatePath}. Make sure it exists in public/template/master.xlsx`,
    );
  }

  const arrayBuffer = await response.arrayBuffer();

  if (!arrayBuffer || arrayBuffer.byteLength === 0) {
    throw new Error(`Template file is empty: ${templatePath}`);
  }

  try {
    return XLSX.read(arrayBuffer, { type: "array" });
  } catch {
    throw new Error(`Failed to parse template file at ${templatePath}`);
  }
}

async function readWorkbook(file: File) {
  const buffer = await file.arrayBuffer();

  if (!buffer || buffer.byteLength === 0) {
    throw new Error(`File is empty: ${file.name}`);
  }

  try {
    return XLSX.read(buffer, { type: "array" });
  } catch {
    throw new Error(
      `Cannot read file "${file.name}". It may not be a valid Excel file.`,
    );
  }
}

async function readAllSheets(file: File): Promise<SheetRows[]> {
  const workbook = await readWorkbook(file);

  return workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];

    let rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (!looksLikeRealData(rows)) {
      rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 1 });
    }

    if (!looksLikeRealData(rows)) {
      rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 2 });
    }

    if (!looksLikeRealData(rows)) {
      rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 3 });
    }

    return {
      sheetName,
      rows: cleanupRows(rows),
    };
  });
}

function looksLikeRealData(rows: GenericRow[]) {
  if (!rows?.length) return false;

  const keys = Object.keys(rows[0] || {}).map((k) => normalizeHeader(k));

  return keys.some((k) =>
    [
      "mac",
      "macaddress",
      "username",
      "serialnumber",
      "sn",
      "personalnumber",
      "vulnerabilitydescription",
      "device",
      "departmentname",
    ].some((x) => k.includes(x)),
  );
}

function cleanupRows(rows: GenericRow[]) {
  return rows
    .map((row) => normalizeRowKeys(row))
    .filter((row) =>
      Object.values(row).some((v) => String(v ?? "").trim() !== ""),
    );
}

function normalizeRowKeys(row: GenericRow) {
  const normalized: GenericRow = {};
  for (const [key, value] of Object.entries(row)) {
    const cleanKey = String(key ?? "").trim();
    normalized[cleanKey] = value ?? "";
  }
  return normalized;
}

function flattenSheetRows(sheets: SheetRows[]) {
  return sheets.flatMap((sheet) =>
    sheet.rows.map((row) => ({
      ...row,
      __sheetName: sheet.sheetName,
    })),
  );
}

function extractHeaders(rows: GenericRow[]) {
  const headers = Object.keys(rows[0] || {});
  return headers.length ? headers : ["Serial Number", "MAC Address"];
}

function normalizeHeader(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function detectColumn(rows: GenericRow[], possibleNames: string[]) {
  if (!rows.length) return null;

  const allKeys = new Set<string>();
  for (const row of rows.slice(0, 100)) {
    Object.keys(row).forEach((k) => allKeys.add(k));
  }

  const keys = [...allKeys];
  const normalizedMap = new Map(keys.map((k) => [normalizeHeader(k), k]));

  for (const name of possibleNames) {
    const found = normalizedMap.get(normalizeHeader(name));
    if (found) return found;
  }

  for (const key of keys) {
    const nk = normalizeHeader(key);
    if (possibleNames.some((name) => nk.includes(normalizeHeader(name)))) {
      return key;
    }
  }

  return null;
}

function normalizeText(value: unknown) {
  return String(value ?? "")
    .trim()
    .toUpperCase();
}

function normalizeMacForCompare(value: unknown) {
  if (!value) return "";

  return String(value)
    .trim()
    .replace(/[^a-fA-F0-9]/g, "")
    .toUpperCase();
}

function formatMacForOutput(value: unknown) {
  const cleaned = normalizeMacForCompare(value);

  if (cleaned.length !== 12) return "";

  const parts = cleaned.match(/.{1,2}/g);
  return parts ? parts.join(":") : "";
}

function dedupeStrings(values: Array<string | null | undefined>) {
  const set = new Set<string>();

  for (const value of values) {
    const clean = String(value ?? "").trim();
    if (clean) set.add(clean);
  }

  return [...set];
}

function getCell(row: GenericRow | null, key: string | null) {
  if (!row || !key) return "";
  return row[key] ?? "";
}

function countFilledValues(row: GenericRow) {
  return Object.values(row).filter((v) => {
    const s = String(v ?? "").trim();
    return s !== "" && s.toUpperCase() !== "#N/A";
  }).length;
}

function indexRowsByKeys(
  rows: GenericRow[],
  candidateKeys: Array<string | null>,
  options: { normalizeMacKeys?: boolean } = {},
) {
  const map = new Map<string, GenericRow>();
  const usableKeys = (candidateKeys || []).filter(Boolean) as string[];

  for (const row of rows) {
    for (const key of usableKeys) {
      const raw = getCell(row, key);
      const normalized = options.normalizeMacKeys
        ? normalizeMacForCompare(raw)
        : normalizeText(raw);

      if (!normalized) continue;

      const existing = map.get(normalized);

      if (!existing) {
        map.set(normalized, row);
        continue;
      }

      const existingFilled = countFilledValues(existing);
      const currentFilled = countFilledValues(row);

      if (currentFilled >= existingFilled) {
        map.set(normalized, row);
      }
    }
  }

  return map;
}

function collectPossibleMacs(rows: Array<GenericRow | null>) {
  const found = new Set<string>();

  for (const row of rows) {
    if (!row) continue;

    for (const [key, value] of Object.entries(row)) {
      if (normalizeHeader(key).includes("mac")) {
        const mac = normalizeMacForCompare(value);
        if (mac) found.add(mac);
      }
    }
  }

  return [...found];
}

function findFirstMacMatch(
  candidateMacs: string[],
  indexMap: Map<string, GenericRow>,
) {
  for (const mac of candidateMacs) {
    const row = indexMap.get(mac);
    if (row) return row;
  }
  return null;
}

function buildBoardingUsernameIndex(boardingRows: GenericRow[]) {
  const map = new Map<string, string>();

  for (const row of boardingRows) {
    const username = getValueFromRow(row, ["user name", "username", "name"]);
    if (!username) continue;

    const macs = collectPossibleMacs([row]);
    for (const mac of macs) {
      if (mac && !map.has(mac)) {
        map.set(mac, String(username));
      }
    }
  }

  return map;
}

function buildCheckleakStatusIndex(rows: GenericRow[], macKey: string | null) {
  const map = new Map<string, CheckleakStatus>();

  if (!macKey) return map;

  for (const row of rows) {
    const mac = normalizeMacForCompare(getCell(row, macKey));
    if (!mac) continue;

    const vuln = String(
      getValueFromRow(row, [
        "vulnerability description",
        "vulnerabilitydescription",
      ]) ?? "",
    )
      .trim()
      .toUpperCase();

    if (!map.has(mac)) {
      map.set(mac, {
        xdrInstalled: false,
        kaseyaInstalled: false,
      });
    }

    const status = map.get(mac)!;

    if (
      vuln.includes("CORTEX") &&
      vuln.includes("XDR") &&
      vuln.includes("ALREADY INSTALLED")
    ) {
      status.xdrInstalled = true;
    }

    if (vuln.includes("KASEYA") && vuln.includes("ALREADY INSTALLED")) {
      status.kaseyaInstalled = true;
    }
  }

  return map;
}

function withNA(value: unknown) {
  if (value === null || value === undefined) return "#N/A";

  const str = String(value).trim();

  if (
    str === "" ||
    str.toLowerCase() === "null" ||
    str.toLowerCase() === "undefined"
  ) {
    return "#N/A";
  }

  return value;
}

function buildMasterKeys(params: {
  rolloutRows: GenericRow[];
  boardingRows: GenericRow[];
  softwareRows: GenericRow[];
  checkleakRows: GenericRow[];
  deviceListRows: GenericRow[];
  rolloutSerialKey: string | null;
  rolloutMacKey: string | null;
  boardingMacKey: string | null;
  softwareMacKey: string | null;
  checkleakMacKey: string | null;
  deviceListSerialKey: string | null;
  deviceListMacKey: string | null;
}) {
  const map = new Map<string, MasterKeyRecord>();

  const addRecord = (serial: string, mac: string) => {
    const normalizedSerial = normalizeText(serial);
    const normalizedMac = normalizeMacForCompare(mac);
    const key = normalizedSerial || normalizedMac;

    if (!key) return;

    const existing = map.get(key);

    if (!existing) {
      map.set(key, {
        key,
        serial: normalizedSerial,
        mac: normalizedMac,
      });
      return;
    }

    if (!existing.serial && normalizedSerial) {
      existing.serial = normalizedSerial;
    }
    if (!existing.mac && normalizedMac) {
      existing.mac = normalizedMac;
    }
  };

  for (const row of params.deviceListRows) {
    addRecord(
      getCell(row, params.deviceListSerialKey),
      getCell(row, params.deviceListMacKey),
    );
  }

  for (const row of params.softwareRows) {
    addRecord("", getCell(row, params.softwareMacKey));
  }

  for (const row of params.checkleakRows) {
    addRecord("", getCell(row, params.checkleakMacKey));
  }

  for (const row of params.boardingRows) {
    addRecord("", getCell(row, params.boardingMacKey));
  }

  for (const row of params.rolloutRows) {
    addRecord(
      getCell(row, params.rolloutSerialKey),
      getCell(row, params.rolloutMacKey),
    );
  }

  return [...map.values()];
}

function pickValueForMasterHeader(header: string, context: MergeContext) {
  const headerNorm = normalizeHeader(header);

  const sourcePriority: Array<GenericRow | null> = [
    context.deviceMatchByMac,
    context.matchedDeviceBySerial,
    context.softwareMatch,
    context.checkleakMatch,
    context.boardingMatch,
    context.rolloutRow,
  ];

  const mergedCheckleakStatus = {
    xdrInstalled: false,
    kaseyaInstalled: false,
  };

  for (const mac of context.candidateMacs || []) {
    const status = context.checkleakStatusByMac.get(mac);
    if (!status) continue;

    if (status.xdrInstalled) mergedCheckleakStatus.xdrInstalled = true;
    if (status.kaseyaInstalled) mergedCheckleakStatus.kaseyaInstalled = true;
  }

  if (headerNorm === "no" || headerNorm === "number") {
    return context.rowNumber;
  }

  if (
    headerNorm.includes("serialnumber") ||
    headerNorm === "sn" ||
    headerNorm.includes("servicetag")
  ) {
    const serialCandidate =
      String(context.serialValue ?? "").trim() !== ""
        ? context.serialValue
        : context.keyRecord.serial ||
          findAnyValue(sourcePriority, [
            "serial number",
            "serialnumber",
            "sn",
            "service tag",
            "bios serial number",
          ]);

    return serialCandidate;
  }

  if (headerNorm.includes("mac")) {
    const macCandidate =
      context.candidateMacs[0] ||
      context.keyRecord.mac ||
      findAnyValue(sourcePriority, [
        "mac address",
        "mac address laptop / handphone",
        "mac address of the terminal",
        "mac",
        "wifi mac",
        "wireless mac",
        "lan mac",
      ]);

    return formatMacForOutput(macCandidate);
  }

  if (headerNorm === "pn") {
    for (const mac of context.candidateMacs || []) {
      const username = sanitizePnValue(context.boardingUsernameByMac.get(mac));

      if (username) return username;
    }

    const rolloutPn = sanitizePnValue(context.rolloutPnValue);
    if (rolloutPn) {
      return rolloutPn;
    }

    return "";
  }

  if (headerNorm === "checkboarding") {
    return context.boardingMatch ? "Posturing+Boarding" : "No Boarding";
  }

  if (headerNorm === "checkboardingonmanual") {
    return String(context.rolloutPnValue ?? "").trim() !== ""
      ? "Posturing+Boarding"
      : "No Boarding";
  }

  if (headerNorm === "posturingboarding") {
    const checkBoardingValue = context.boardingMatch
      ? "Posturing+Boarding"
      : "No Boarding";

    const checkBoardingManualValue =
      String(context.rolloutPnValue ?? "").trim() !== ""
        ? "Posturing+Boarding"
        : "No Boarding";

    if (
      checkBoardingValue === "No Boarding" &&
      checkBoardingManualValue === "No Boarding"
    ) {
      return "No Boarding";
    }

    return "Posturing+Boarding";
  }

  if (headerNorm === "divisibygrouping") {
    return findAnyValue(
      [context.softwareMatch],
      ["department name", "departmentname", "department"],
    );
  }

  if (headerNorm === "xdr") {
    return mergedCheckleakStatus.xdrInstalled
      ? "Cortex XDR already installed"
      : "Not install";
  }

  if (headerNorm === "kaseya") {
    return mergedCheckleakStatus.kaseyaInstalled
      ? "Kaseya already installed"
      : "Not install";
  }

  // if (headerNorm === "reasonnotcomply" || headerNorm === "comply") {
  //   const xdrValue = mergedCheckleakStatus.xdrInstalled
  //     ? "Cortex XDR already installed"
  //     : "Not install";

  //   const kaseyaValue = mergedCheckleakStatus.kaseyaInstalled
  //     ? "Kaseya already installed"
  //     : "Not install";

  //   if (
  //     xdrValue === "Cortex XDR already installed" &&
  //     kaseyaValue === "Kaseya already installed"
  //   ) {
  //     return "Comply";
  //   }

  //   if (
  //     xdrValue === "Cortex XDR already installed" &&
  //     kaseyaValue === "Not install"
  //   ) {
  //     return "Not install Kaseya";
  //   }

  //   if (
  //     xdrValue === "Not install" &&
  //     kaseyaValue === "Kaseya already installed"
  //   ) {
  //     return "Not install XDR";
  //   }

  //   if (xdrValue === "Not install" && kaseyaValue === "Not install") {
  //     return "Not Comply";
  //   }

  //   return "Not Comply";
  // }

  if (headerNorm === "reasonnotcomply") {
    const xdrValue = mergedCheckleakStatus.xdrInstalled
      ? "Cortex XDR already installed"
      : "Not install";

    const kaseyaValue = mergedCheckleakStatus.kaseyaInstalled
      ? "Kaseya already installed"
      : "Not install";

    const posturingBoardingValue =
      context.boardingMatch ||
      String(context.rolloutPnValue ?? "").trim() !== ""
        ? "Boarding"
        : "No Boarding";

    if (
      xdrValue === "Cortex XDR already installed" &&
      kaseyaValue === "Kaseya already installed"
    ) {
      return posturingBoardingValue === "No Boarding"
        ? "Not Boarding"
        : "Comply";
    }

    if (
      xdrValue === "Cortex XDR already installed" &&
      kaseyaValue === "Not install"
    ) {
      return "Not install Kaseya";
    }

    if (
      xdrValue === "Not install" &&
      kaseyaValue === "Kaseya already installed"
    ) {
      return "Not install XDR";
    }

    if (xdrValue === "Not install" && kaseyaValue === "Not install") {
      return posturingBoardingValue === "No Boarding"
        ? "Not Boarding"
        : "Not install XDR and KASEYA";
    }

    return "Not Comply";
  }

  const exactValue = findMatchingHeaderValue(header, sourcePriority);
  if (String(exactValue ?? "").trim() !== "") return exactValue;

  const looseValue = findLooseHeaderValue(header, sourcePriority);
  if (String(looseValue ?? "").trim() !== "") return looseValue;

  return "";
}

function findMatchingHeaderValue(
  targetHeader: string,
  rows: Array<GenericRow | null>,
) {
  const targetNorm = normalizeHeader(targetHeader);

  for (const row of rows) {
    if (!row) continue;

    for (const [key, value] of Object.entries(row)) {
      if (
        normalizeHeader(key) === targetNorm &&
        String(value ?? "").trim() !== ""
      ) {
        return value;
      }
    }
  }

  return "";
}

function getValueFromRow(row: GenericRow | null, possibleHeaders: string[]) {
  if (!row) return "";

  for (const [key, value] of Object.entries(row)) {
    const keyNorm = normalizeHeader(key);

    for (const header of possibleHeaders) {
      const headerNorm = normalizeHeader(header);

      if (!headerNorm) continue;

      if (
        keyNorm === headerNorm ||
        keyNorm.includes(headerNorm) ||
        headerNorm.includes(keyNorm)
      ) {
        if (String(value ?? "").trim() !== "") {
          return value;
        }
      }
    }
  }

  return "";
}

function findLooseHeaderValue(
  targetHeader: string,
  rows: Array<GenericRow | null>,
) {
  const targetNorm = normalizeHeader(targetHeader);

  for (const row of rows) {
    if (!row) continue;

    for (const [key, value] of Object.entries(row)) {
      const keyNorm = normalizeHeader(key);
      if (
        String(value ?? "").trim() !== "" &&
        (keyNorm.includes(targetNorm) || targetNorm.includes(keyNorm))
      ) {
        return value;
      }
    }
  }

  return "";
}

function findAnyValue(
  rows: Array<GenericRow | null>,
  possibleHeaders: string[],
) {
  for (const row of rows) {
    if (!row) continue;

    for (const [key, value] of Object.entries(row)) {
      const keyNorm = normalizeHeader(key);

      for (const header of possibleHeaders) {
        const headerNorm = normalizeHeader(header);

        if (
          keyNorm === headerNorm ||
          keyNorm.includes(headerNorm) ||
          headerNorm.includes(keyNorm)
        ) {
          if (String(value ?? "").trim() !== "") {
            return value;
          }
        }
      }
    }
  }

  return "";
}

function sortRowsByQuality(rows: GenericRow[]) {
  return rows.sort((a, b) => {
    const score = (row: GenericRow) =>
      Object.values(row).filter(
        (v) => String(v).trim() !== "" && String(v).toUpperCase() !== "#N/A",
      ).length;

    return score(b) - score(a);
  });
}

function sanitizePnValue(value: unknown) {
  const text = String(value ?? "").trim();

  if (!text) return "";

  // if contains email, take part before "@"
  if (text.includes("@")) {
    return text.split("@")[0];
  }

  return text;
}
const styles: Record<string, CSSProperties> = {
  page: {
    minHeight: "100vh",
    padding: 24,
    background: "#f5f7fb",
  },
  card: {
    maxWidth: 1200,
    margin: "0 auto",
    background: "#fff",
    padding: 24,
    borderRadius: 16,
    boxShadow: "0 8px 30px rgba(0,0,0,0.08)",
  },
  title: {
    margin: 0,
    marginBottom: 8,
    fontSize: 28,
  },
  subtitle: {
    marginTop: 0,
    color: "#4b5563",
  },
  notice: {
    marginTop: 12,
    background: "#eef2ff",
    border: "1px solid #c7d2fe",
    color: "#3730a3",
    padding: 12,
    borderRadius: 10,
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))",
    gap: 16,
    marginTop: 24,
  },
  fileBox: {
    border: "1px solid #e5e7eb",
    borderRadius: 12,
    padding: 16,
    background: "#fafafa",
  },
  label: {
    display: "block",
    fontWeight: 700,
    marginBottom: 8,
  },
  fileName: {
    marginTop: 8,
    fontSize: 13,
    color: "#6b7280",
    wordBreak: "break-word",
  },
  actions: {
    display: "flex",
    alignItems: "center",
    gap: 16,
    marginTop: 24,
    flexWrap: "wrap",
  },
  button: {
    border: 0,
    borderRadius: 10,
    background: "#2563eb",
    color: "#fff",
    padding: "12px 18px",
    cursor: "pointer",
    fontWeight: 700,
  },
  counter: {
    color: "#374151",
  },
  status: {
    marginTop: 18,
    padding: 12,
    background: "#eff6ff",
    border: "1px solid #bfdbfe",
    borderRadius: 10,
    color: "#1d4ed8",
  },
};
