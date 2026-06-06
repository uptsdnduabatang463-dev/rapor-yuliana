import React, { useState, useEffect, useRef } from "react";
import { unstable_batchedUpdates } from "react-dom";
import {
  BrowserRouter as Router,
  Route,
  Routes,
  Link,
  useNavigate,
} from "react-router-dom";
import jsPDF from "jspdf"; // Tambahkan impor ini
import autoTable from "jspdf-autotable";
import SignatureCanvas from "react-signature-canvas";
import * as XLSX from "xlsx";

// Extend jsPDF type untuk mendukung lastAutoTable dari plugin autotable
declare module "jspdf" {
  interface jsPDF {
    lastAutoTable: {
      finalY: number;
    };
  }
}

interface RowData {
  [key: string]: string;
}

interface SheetInfo {
  sheetName: string;
  mapel: string;
  semester: string;
  kelas: string;
}

interface SchoolData {
  namaSekolah: string;
  npsn: string;
  alamatSekolah: string;
  kodePos: string;
  desaKelurahan: string;
  kabKota: string;
  provinsi: string;
  tahunPelajaran: string; // ✅ TAMBAHAN BARU
  tanggalRapor: string; // ✅ TAMBAHAN BARU
  nilaiKKM: string;
  kelas: string;
  rombel: string;
  namaKepsek: string;
  nipKepsek: string;
  ttdKepsek: string;
  namaGuru: string;
  nipGuru: string;
  ttdGuru: string;
}

interface KehadiranData {
  [key: string]: string;
}

const endpoint =
  "https://script.google.com/macros/s/AKfycbzQfl3bj32WODDuVZ2G53navKZVejAJxa_XWPenjjniu04XWtlcpvEJzRzXY7cVB_bV/exec";

const throttle = (func: Function, delay: number) => {
  let timeoutId: ReturnType<typeof setTimeout> | null = null;
  let lastRan: number = 0;

  return function (this: any, ...args: any[]) {
    const now = Date.now();

    if (now - lastRan >= delay) {
      func.apply(this, args);
      lastRan = now;
    } else {
      if (timeoutId) clearTimeout(timeoutId);
      timeoutId = setTimeout(() => {
        func.apply(this, args);
        lastRan = Date.now();
      }, delay - (now - lastRan));
    }
  };
};

// ============================================================
// ✅ INDEXEDDB HELPER - Penyimpanan permanen untuk data TP
// ============================================================
const DB_NAME = "RaporDB";
const DB_VERSION = 3;
const STORE_TP = "tpData";
const STORE_MAPEL = "mapelData";

const openDB = (): Promise<IDBDatabase> => {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = (e.target as IDBOpenDBRequest).result;
      if (!db.objectStoreNames.contains(STORE_TP)) {
        db.createObjectStore(STORE_TP, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(STORE_MAPEL)) {
        db.createObjectStore(STORE_MAPEL, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains("kelasData")) {
        db.createObjectStore("kelasData", { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains("sekolahData")) {
        db.createObjectStore("sekolahData", { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains("siswaData")) {
        db.createObjectStore("siswaData", { keyPath: "id" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
};

const idbSave = async (storeName: string, data: any): Promise<void> => {
  try {
    const db = await openDB();
    const tx = db.transaction(storeName, "readwrite");
    const store = tx.objectStore(storeName);
    store.put({ id: "main", data, savedAt: Date.now() });
    return new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (e) {
    console.warn(`idbSave(${storeName}) gagal:`, e);
  }
};

const idbLoad = async (storeName: string): Promise<any | null> => {
  try {
    const db = await openDB();
    const tx = db.transaction(storeName, "readonly");
    const store = tx.objectStore(storeName);
    const req = store.get("main");
    return new Promise((resolve, reject) => {
      req.onsuccess = () => resolve(req.result ? req.result.data : null);
      req.onerror = () => reject(req.error);
    });
  } catch (e) {
    console.warn(`idbLoad(${storeName}) gagal:`, e);
    return null;
  }
};
// ============================================================

// Context untuk pre-loading data rekap
interface RekapContextType {
  availableSheets: SheetInfo[];
  schoolData: SchoolData | null;
  kehadiranData: KehadiranData[];
  allMapelData: { [sheetName: string]: RowData[] };
  tpData: RowData[];
  mapelListData: RowData[];
  kokurikulerData: RowData[];
  ekstrakurikulerData: RowData[];
  siswaData: RowData[];
  rekapNilaiData: RowData[];
  rekapNilai2Data: RowData[];
  loading: boolean;
  error: string | null;
  refreshRekapData: (silent?: boolean) => Promise<void>;
  updateLocalData: (sheetType: string, newData: any) => void;
  refreshMapelSheet: (mapelName: string) => Promise<void>;
  addSiswaToMapelSheets: (
    namaSiswa: string,
    kelas: string,
    nis: string,
    nisn: string
  ) => number | undefined; // ✅ TAMBAH
  removeSiswaFromMapelSheets: (
    namaSiswa: string,
    kelas: string
  ) => number | undefined; // ✅ TAMBAH
}

const RekapContext = React.createContext<RekapContextType | null>(null);

const RekapProvider: React.FC<{ children: React.ReactNode }> = ({
  children,
}) => {
  const [availableSheets, setAvailableSheets] = useState<SheetInfo[]>([]);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [kehadiranData, setKehadiranData] = useState<KehadiranData[]>([]);
  const [allMapelData, setAllMapelData] = useState<{
    [sheetName: string]: RowData[];
  }>({});
  const [tpData, setTPData] = useState<RowData[]>([]);
  const [mapelListData, setMapelListData] = useState<RowData[]>([]);
  const [kokurikulerData, setKokurikulerData] = useState<RowData[]>([]);
  const [ekstrakurikulerData, setEkstrakurikulerData] = useState<RowData[]>([]);
  const [siswaData, setSiswaData] = useState<RowData[]>([]);
  const [rekapNilaiData, setRekapNilaiData] = useState<RowData[]>([]);
  const [rekapNilai2Data, setRekapNilai2Data] = useState<RowData[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // ← TAMBAH: load cache localStorage saat pertama kali
  useEffect(() => {
    try {
      const cachedTP = localStorage.getItem("cache_tpData");
      const cachedMapel = localStorage.getItem("cache_mapelData");
      const cachedAllMapel = localStorage.getItem("cache_allMapelData");
      const cachedSiswa = localStorage.getItem("cache_siswaData");

      if (cachedTP) setTPData(JSON.parse(cachedTP));
      if (cachedMapel) setMapelListData(JSON.parse(cachedMapel));
      if (cachedAllMapel) setAllMapelData(JSON.parse(cachedAllMapel));
      if (cachedSiswa) setSiswaData(JSON.parse(cachedSiswa));
    } catch (e) {
      console.warn("Gagal load cache:", e);
    }
  }, []);

  const refreshRekapData = async (silent: boolean = false) => {
    if (!silent) {
      setLoading(true);
    }
    setError(null);
    try {
      const [
        schoolResponse,
        kehadiranResponse,
        sheetsResponse,
        tpResponse,
        mapelResponse,
        kokurikulerResponse,
        ekstrakurikulerResponse,
        siswaResponse,
        rekapNilaiResponse,
        rekapNilai2Response,
      ] = await Promise.all([
        fetch(`${endpoint}?action=schoolData`),
        fetch(`${endpoint}?sheet=DataKehadiran`),
        fetch(`${endpoint}?action=listSheets`),
        fetch(`${endpoint}?sheet=DataTP`),
        fetch(`${endpoint}?sheet=DataMapel`),
        fetch(`${endpoint}?sheet=DataKokurikuler`),
        fetch(`${endpoint}?sheet=DataEkstrakurikuler`),
        fetch(`${endpoint}?sheet=DataSiswa`),
        fetch(`${endpoint}?sheet=RekapNilai1`), // ✅ UBAH dari RekapNilai
        fetch(`${endpoint}?sheet=RekapNilai2`),
      ]);

      if (kehadiranResponse.ok) {
        const kehadiranJson = await kehadiranResponse.json();
        if (Array.isArray(kehadiranJson)) {
          // ← TAMBAH pengecekan
          setKehadiranData(kehadiranJson.slice(1));
        }
      }
      let tpJson: any[] = []; // ← TAMBAH deklarasi di luar
      let mapelJson: any[] = []; // ← TAMBAH deklarasi di luar
      let siswaJson: any[] = []; // ← TAMBAH deklarasi di luar

      if (tpResponse.ok) {
        tpJson = await tpResponse.json(); // ← hapus const
        setTPData(tpJson);
      }

      if (mapelResponse.ok) {
        mapelJson = await mapelResponse.json(); // ← hapus const
        setMapelListData(mapelJson);
      }

      if (kokurikulerResponse.ok) {
        const kokurikulerJson = await kokurikulerResponse.json();
        setKokurikulerData(kokurikulerJson);
      }

      if (ekstrakurikulerResponse.ok) {
        const ekstrakurikulerJson = await ekstrakurikulerResponse.json();
        setEkstrakurikulerData(ekstrakurikulerJson);
      }

      if (siswaResponse.ok) {
        siswaJson = await siswaResponse.json(); // ← hapus const
        setSiswaData(siswaJson);
      }

      if (rekapNilaiResponse.ok) {
        const rekapNilaiJson = await rekapNilaiResponse.json();
        setRekapNilaiData(rekapNilaiJson);
      }

      if (rekapNilai2Response.ok) {
        const rekapNilai2Json = await rekapNilai2Response.json();
        setRekapNilai2Data(rekapNilai2Json);
      }

      if (schoolResponse.ok) {
        const schoolJson = await schoolResponse.json();
        if (schoolJson.success && schoolJson.data?.length > 0) {
          setSchoolData(schoolJson.data[0]);
        }
      }

      if (!sheetsResponse.ok) throw new Error("Gagal mengambil daftar sheet");
      const sheetsRaw: SheetInfo[] = await sheetsResponse.json();

      // Filter sheet yang tidak valid (REF error, N/A, kosong)
      const sheets = sheetsRaw.filter((sheet) => {
        const mapel = sheet.mapel || "";
        return (
          mapel.trim() !== "" &&
          !mapel.includes("#REF!") &&
          !mapel.includes("#N/A") &&
          !mapel.includes("N/A") &&
          !mapel.toUpperCase().includes("ERROR")
        );
      });

      setAvailableSheets(sheets);

      // ✅ PRE-LOAD SEMUA DATA SHEET MAPEL SEKALIGUS (ini yang kamu inginkan!)
      const allDataPromises = sheets.map(async (sheet) => {
        try {
          const response = await fetch(`${endpoint}?sheet=${sheet.sheetName}`);
          if (!response.ok) {
            console.warn(`Gagal load sheet ${sheet.sheetName}`);
            return null;
          }
          const jsonData = await response.json();

          // Filter hanya siswa yang punya nama
          const headers = jsonData[0];
          const filteredData = jsonData.slice(1).filter((row: any) => {
            const nama = row.Data4;
            return nama && typeof nama === "string" && nama.trim() !== "";
          });

          const cleanedData = [headers, ...filteredData];

          return { sheetName: sheet.sheetName, data: cleanedData };
        } catch (err) {
          console.error(`Error loading ${sheet.sheetName}:`, err);
          return null;
        }
      });

      const results = await Promise.all(allDataPromises);

      // Simpan semua data ke state
      const mapelDataMap: { [sheetName: string]: RowData[] } = {};
      results.forEach((result) => {
        if (result) {
          mapelDataMap[result.sheetName] = result.data;
        }
      });

      setAllMapelData(mapelDataMap);

      // ← TAMBAH: simpan ke localStorage sebagai cache
      try {
        localStorage.setItem("cache_tpData", JSON.stringify(tpJson));
        localStorage.setItem("cache_mapelData", JSON.stringify(mapelJson));
        localStorage.setItem(
          "cache_allMapelData",
          JSON.stringify(mapelDataMap)
        );
        localStorage.setItem("cache_siswaData", JSON.stringify(siswaJson));
      } catch (e) {
        console.warn("Gagal simpan cache:", e);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unknown error");
      console.error(err);
    } finally {
      if (!silent) {
        setLoading(false);
      }
    }
  };

  const updateLocalData = (sheetType: string, newData: any) => {
    switch (sheetType) {
      case "siswa":
        setSiswaData(newData);
        break;
      case "tp":
        setTPData(newData);
        break;
      case "mapel":
        setMapelListData(newData);
        break;
      case "kehadiran":
        setKehadiranData(newData);
        break;
      case "kokurikuler":
        setKokurikulerData(newData);
        break;
      case "ekstrakurikuler":
        setEkstrakurikulerData(newData);
        break;
      case "rekapNilai":
        setRekapNilaiData(newData);
        break;
      case "rekapNilai2": // ✅ TAMBAH INI
        setRekapNilai2Data(newData);
        break;
      case "sheets":
        setAvailableSheets(newData);
        break;
    }
  };

  const refreshMapelSheet = async (mapelName: string) => {
    try {
      // Cari sheet yang sesuai dengan mapel
      const targetSheets = availableSheets.filter(
        (sheet) => sheet.mapel === mapelName
      );

      if (targetSheets.length === 0) {
        console.log(`No sheets found for mapel: ${mapelName}`);
        return;
      }

      // Reload semua sheet untuk mapel ini (bisa ada lebih dari 1 semester/kelas)
      const reloadPromises = targetSheets.map(async (sheet) => {
        try {
          const response = await fetch(`${endpoint}?sheet=${sheet.sheetName}`);
          if (!response.ok) {
            console.warn(`Failed to reload sheet ${sheet.sheetName}`);
            return null;
          }
          const jsonData = await response.json();

          // Filter hanya siswa yang punya nama
          const headers = jsonData[0];
          const filteredData = jsonData.slice(1).filter((row: any) => {
            const nama = row.Data4;
            return nama && typeof nama === "string" && nama.trim() !== "";
          });

          const cleanedData = [headers, ...filteredData];

          return { sheetName: sheet.sheetName, data: cleanedData };
        } catch (err) {
          console.error(`Error loading ${sheet.sheetName}:`, err);
          return null;
        }
      });

      const results = await Promise.all(reloadPromises);

      // Update allMapelData dengan data yang baru
      const updatedMapelData = { ...allMapelData };
      results.forEach((result) => {
        if (result) {
          updatedMapelData[result.sheetName] = result.data;
        }
      });

      setAllMapelData(updatedMapelData);

      console.log(`✅ Sheet mapel ${mapelName} berhasil di-refresh`);
    } catch (err) {
      console.error(`Error refreshing mapel sheet ${mapelName}:`, err);
    }
  };

  // ✅ TAMBAH FUNGSI BARU - Optimistic update untuk siswa baru
  const addSiswaToMapelSheets = (
    namaSiswa: string,
    kelas: string,
    nis: string,
    nisn: string
  ) => {
    console.log(`🚀 Optimistic update: Adding ${namaSiswa} to mapel sheets`);

    // Filter sheet yang sesuai dengan kelas
    const targetSheets = availableSheets.filter(
      (sheet) => sheet.kelas === kelas
    );

    if (targetSheets.length === 0) {
      console.log(`No sheets found for class ${kelas}`);
      return;
    }

    // Clone allMapelData untuk update
    const updatedMapelData = { ...allMapelData };

    targetSheets.forEach((sheet) => {
      const sheetData = updatedMapelData[sheet.sheetName];
      if (!sheetData || sheetData.length === 0) return;

      // Ambil header
      const headers = sheetData[0];

      // Buat row baru untuk siswa (semua kolom nilai kosong)
      const newRow: RowData = {};
      Object.keys(headers).forEach((key) => {
        if (key === "Data1") newRow[key] = sheet.mapel; // Mapel
        else if (key === "Data2") newRow[key] = sheet.semester; // Semester
        else if (key === "Data3") newRow[key] = kelas; // Kelas
        else if (key === "Data4") newRow[key] = namaSiswa; // Nama
        else if (key === "Data26") newRow[key] = nisn; // NISN
        else if (key === "Data27") newRow[key] = nis; // NIS
        else newRow[key] = ""; // Kolom lain kosong
      });

      // Tambahkan row baru ke data
      const updatedSheetData = [...sheetData, newRow];
      updatedMapelData[sheet.sheetName] = updatedSheetData;

      console.log(`✅ Added ${namaSiswa} to ${sheet.sheetName} (optimistic)`);
    });

    // Update state dan localStorage
    setAllMapelData(updatedMapelData);

    return targetSheets.length;
  };

  // ✅ TAMBAH FUNGSI BARU - Optimistic delete untuk siswa
  const removeSiswaFromMapelSheets = (namaSiswa: string, kelas: string) => {
    console.log(
      `🗑️ Optimistic update: Removing ${namaSiswa} from mapel sheets`
    );

    // Filter sheet yang sesuai dengan kelas
    const targetSheets = availableSheets.filter(
      (sheet) => sheet.kelas === kelas
    );

    if (targetSheets.length === 0) {
      console.log(`No sheets found for class ${kelas}`);
      return;
    }

    // Clone allMapelData untuk update
    const updatedMapelData = { ...allMapelData };

    targetSheets.forEach((sheet) => {
      const sheetData = updatedMapelData[sheet.sheetName];
      if (!sheetData || sheetData.length === 0) return;

      // Filter out siswa yang dihapus
      const headers = sheetData[0];
      const filteredData = sheetData.slice(1).filter((row) => {
        return row.Data4 !== namaSiswa;
      });

      const updatedSheetData = [headers, ...filteredData];
      updatedMapelData[sheet.sheetName] = updatedSheetData;

      console.log(
        `✅ Removed ${namaSiswa} from ${sheet.sheetName} (optimistic)`
      );
    });

    return targetSheets.length;
  };

  useEffect(() => {
    const init = async () => {
      // Fetch listSheets duluan agar dropdown langsung tersedia
      try {
        const sheetsResponse = await fetch(`${endpoint}?action=listSheets`);
        if (sheetsResponse.ok) {
          const sheetsRaw: SheetInfo[] = await sheetsResponse.json();
          const sheets = sheetsRaw.filter((sheet) => {
            const mapel = sheet.mapel || "";
            return (
              mapel.trim() !== "" &&
              !mapel.includes("#REF!") &&
              !mapel.includes("#N/A") &&
              !mapel.includes("N/A") &&
              !mapel.toUpperCase().includes("ERROR")
            );
          });
          setAvailableSheets(sheets);
        }
      } catch (err) {
        console.error("Error fetching sheets list:", err);
      }
      // Baru fetch semua data lengkap
      refreshRekapData();
    };
    init();
  }, []);

  return (
    <RekapContext.Provider
      value={{
        availableSheets,
        schoolData,
        kehadiranData,
        allMapelData,
        tpData,
        mapelListData,
        kokurikulerData,
        ekstrakurikulerData,
        siswaData,
        rekapNilaiData,
        rekapNilai2Data,
        loading,
        error,
        refreshRekapData,
        updateLocalData,
        refreshMapelSheet,
        addSiswaToMapelSheets, // ✅ TAMBAH INI
        removeSiswaFromMapelSheets,
      }}
    >
      {children}
    </RekapContext.Provider>
  );
};

// Hook untuk menggunakan context
const useRekapData = () => {
  const context = React.useContext(RekapContext);
  if (!context) {
    throw new Error("useRekapData must be used within RekapProvider");
  }
  return context;
};

const InputNilai = () => {
  const { availableSheets, refreshRekapData } = useRekapData();
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [selectedSheet, setSelectedSheet] = useState<string>("MAPEL101");
  const [showTPPopup, setShowTPPopup] = useState(false);
  const [selectedTP, setSelectedTP] = useState<string>("");
  const [tpDetails, setTPDetails] = useState<any>(null);
  const [loadingTP, setLoadingTP] = useState(false);
  const [showFloatingButton, setShowFloatingButton] = useState(false);
  const [floatingButtonPosition, setFloatingButtonPosition] = useState({
    top: 0,
    left: 0,
    visible: true,
  });
  const [activeInput, setActiveInput] = useState<{
    rowIndex: number;
    colIndex: number;
  } | null>(null);
  const [isProcessingClick, setIsProcessingClick] = useState(false);
  const [showDescPopup, setShowDescPopup] = useState(false);
  const [selectedStudentDesc, setSelectedStudentDesc] = useState<{
    nama: string;
    descMin: string;
    descMax: string;
    tpMin: string;
    tpMax: string;
    nilaiMin: string;
    nilaiMax: string;
  } | null>(null);
  const [isLoadingDesc, setIsLoadingDesc] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);
  const [importResult, setImportResult] = useState<{
    matched: number;
    notFound: string[];
    targetColumn: string;
  } | null>(null);
  const [isEditMode, setIsEditMode] = useState(false);
  const [originalData, setOriginalData] = useState<RowData[]>([]);
  // ✅ TAMBAH: Menyimpan mapping row index asli (di sheet) vs row index yang terfilter (di UI)
  const [rowMapping, setRowMapping] = useState<number[]>([]);

  const loadSheetData = async (sheetName: string) => {
    setLoading(true);
    setError(null);

    try {
      const startTime = performance.now();

      const response = await fetch(`${endpoint}?sheet=${sheetName}`);

      if (!response.ok) {
        throw new Error(`Failed to load ${sheetName}`);
      }

      const jsonData = await response.json();

      // ✅ PERBAIKAN: Filter siswa dengan nama DAN simpan mapping row asli
      if (jsonData.length > 0) {
        const headers = jsonData[0];
        const allRows = jsonData.slice(1);

        const filteredDataWithMapping: { row: any; originalIndex: number }[] =
          [];

        allRows.forEach((row: any, index: number) => {
          const nama = row.Data4;
          // Filter hanya row yang punya nama
          if (nama && typeof nama === "string" && nama.trim() !== "") {
            filteredDataWithMapping.push({
              row: row,
              originalIndex: index + 2, // +2 karena: +1 untuk header, +1 untuk 0-based index
            });
          }
        });

        // Ekstrak data yang sudah difilter
        const filteredData = filteredDataWithMapping.map((item) => item.row);

        // Simpan mapping row index asli
        const mapping = filteredDataWithMapping.map(
          (item) => item.originalIndex
        );
        setRowMapping(mapping);

        const cleanedData = [headers, ...filteredData];
        setData(cleanedData);

        const endTime = performance.now();
        console.log(
          `✅ Loaded ${sheetName} in ${((endTime - startTime) / 1000).toFixed(
            2
          )}s`
        );
        console.log(
          `📊 Filtered ${filteredData.length} rows with valid names from ${allRows.length} total rows`
        );
        console.log(`🗺️ Row mapping:`, mapping);
      } else {
        setData(jsonData);
        setRowMapping([]);
      }

      setLoading(false);
    } catch (err) {
      console.error(`Error loading ${sheetName}:`, err);
      setError(err instanceof Error ? err.message : "Unknown error");
      setLoading(false);
    }
  };

  useEffect(() => {
    if (availableSheets.length > 0 && !selectedSheet) {
      setSelectedSheet(availableSheets[0].sheetName);
    }
  }, [availableSheets]);

  useEffect(() => {
    if (selectedSheet) {
      loadSheetData(selectedSheet);
    }
  }, [selectedSheet]);

  useEffect(() => {
    const updateButtonPosition = () => {
      if (showFloatingButton && activeInput) {
        const { rowIndex, colIndex } = activeInput;
        const input = document.getElementById(
          `input-${rowIndex}-${colIndex}`
        ) as HTMLInputElement;

        if (input) {
          const rect = input.getBoundingClientRect();
          const tableContainer = document.getElementById(
            "table-scroll-container"
          );

          if (tableContainer) {
            const containerRect = tableContainer.getBoundingClientRect();
            const thead = tableContainer.querySelector("thead");
            const headerHeight = thead ? thead.offsetHeight : 40;

            const inputTopInContainer = rect.top - containerRect.top;
            const inputBottomInContainer = rect.bottom - containerRect.top;

            const isVisibleInContainer =
              inputTopInContainer >= headerHeight &&
              inputBottomInContainer > headerHeight &&
              rect.bottom <= containerRect.bottom &&
              rect.left >= containerRect.left - 100 &&
              rect.right <= window.innerWidth + 100;

            setFloatingButtonPosition({
              top: rect.top + rect.height / 2 - 28,
              left: rect.right + 10,
              visible: isVisibleInContainer,
            });
          }
        }
      }
    };

    const handleScroll = throttle(updateButtonPosition, 16);
    const tableContainer = document.getElementById("table-scroll-container");

    if (tableContainer) {
      tableContainer.addEventListener("scroll", handleScroll as any, {
        passive: true,
      });
    }

    window.addEventListener("scroll", handleScroll as any, { passive: true });

    return () => {
      if (tableContainer) {
        tableContainer.removeEventListener("scroll", handleScroll as any);
      }
      window.removeEventListener("scroll", handleScroll as any);
    };
  }, [showFloatingButton, activeInput]);

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleKeyDown = (
    e: React.KeyboardEvent<HTMLInputElement>,
    rowIndex: number,
    colIndex: number,
    actualDataLength: number
  ) => {
    if (e.key === "Enter") {
      e.preventDefault();
      const nextRow = rowIndex + 1;
      if (nextRow < actualDataLength) {
        const nextInput = document.getElementById(
          `input-${nextRow}-${colIndex}`
        ) as HTMLInputElement | null;
        if (nextInput) {
          nextInput.focus();
          nextInput.select();
        }
      }
    }
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("No changes to save!");
      return;
    }

    setIsSaving(true);

    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");

      // ✅ PERBAIKAN: Gunakan rowMapping untuk mendapatkan row index asli di sheet
      const originalRowIndex = rowMapping[rowIndex];

      console.log(
        `📝 Saving row ${rowIndex} (UI) -> Row ${originalRowIndex} (Sheet):`,
        values
      );

      updates.push({
        rowIndex: originalRowIndex + 1, // +1 karena sheet row dimulai dari 1, bukan 0
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: selectedSheet,
        updates: updates,
      };

      console.log("📤 Sending updates:", requestBody);

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const responseText = await response.text();

      try {
        const responseJson = JSON.parse(responseText);
        if (responseJson.error) {
          throw new Error(responseJson.error);
        }
        console.log("✅ Server response:", responseJson);
      } catch (parseError) {
        console.log(
          "Could not parse response, but request might be successful"
        );
      }

      alert("All changes saved successfully!");
      setChangedRows(new Set());
      setShowFloatingButton(false);
      setActiveInput(null);
      setIsEditMode(false);
      setOriginalData([]);

      // Reload data dari server
      await loadSheetData(selectedSheet);

      setIsSaving(false);

      // Background refresh untuk RekapNilai
      setTimeout(() => {
        console.log("🔄 Background refresh RekapNilai...");
        refreshRekapData(true);
      }, 2000);
    } catch (err) {
      console.error("=== ERROR DETAILS ===");
      console.error(err);
      alert(
        "Error updating rows: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  const handleClearAllValues = async () => {
    const confirmation = window.confirm(
      `⚠️ HAPUS SEMUA NILAI?\n\nIni akan mengosongkan kolom Data5-Data19 untuk SEMUA baris di sheet "${selectedSheet}".\n\nLanjutkan?`
    );
    if (!confirmation) return;

    const doubleConfirm = window.confirm(
      "⚠️ KONFIRMASI TERAKHIR!\n\nSemua nilai akan dihapus permanen!\n\nYakin lanjutkan?"
    );
    if (!doubleConfirm) return;

    setIsSaving(true);

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "clear_nilai_range",
          sheetName: selectedSheet,
        }),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      alert("✅ Semua nilai berhasil dihapus!");
      setChangedRows(new Set());
      await loadSheetData(selectedSheet);
      setTimeout(() => refreshRekapData(true), 2000);
    } catch (err) {
      alert(
        "❌ Error: " + (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const handleSheetChange = (newSheet: string) => {
    setSelectedSheet(newSheet);
    setChangedRows(new Set());
    setIsEditMode(false);
    setOriginalData([]);
  };

  const reloadDataKehadiran = async () => {
    try {
      const response = await fetch(`${endpoint}?sheet=${selectedSheet}`);
      if (!response.ok) {
        throw new Error("Failed to reload data");
      }
      const jsonData = await response.json();

      if (jsonData.length > 0) {
        const headers = jsonData[0];
        const allRows = jsonData.slice(1);

        const filteredDataWithMapping: { row: any; originalIndex: number }[] =
          [];

        allRows.forEach((row: any, index: number) => {
          const nama = row.Data4;
          if (nama && typeof nama === "string" && nama.trim() !== "") {
            filteredDataWithMapping.push({
              row: row,
              originalIndex: index + 2,
            });
          }
        });

        const filteredData = filteredDataWithMapping.map((item) => item.row);
        const mapping = filteredDataWithMapping.map(
          (item) => item.originalIndex
        );

        setRowMapping(mapping);
        const cleanedData = [headers, ...filteredData];
        setData(cleanedData);

        return cleanedData;
      } else {
        setData(jsonData);
        setRowMapping([]);
        return jsonData;
      }
    } catch (err) {
      console.error("Error reloading data:", err);
      return null;
    }
  };

  const fetchTPDetails = async (
    tpCode: string,
    mapel: string,
    rowIndex: number,
    kelas: string = "",
    semester: string = ""
  ) => {
    console.log(
      "Fetching TP:",
      tpCode,
      "Mapel:",
      mapel,
      "Kelas:",
      kelas,
      "Semester:",
      semester
    );

    setLoadingTP(true);
    setShowTPPopup(true);
    setSelectedTP(tpCode);

    try {
      const url = `${endpoint}?sheet=DataTP&tp=${encodeURIComponent(
        tpCode
      )}&mapel=${encodeURIComponent(mapel)}&kelas=${encodeURIComponent(
        kelas
      )}&semester=${encodeURIComponent(semester)}`;

      const response = await fetch(url);
      if (!response.ok) throw new Error("Failed to fetch TP details");

      const tpData = await response.json();
      setTPDetails(tpData);
    } catch (err) {
      console.error("Error fetching TP details:", err);
      setTPDetails({ error: "Gagal memuat data" });
    } finally {
      setLoadingTP(false);
    }
  };

  const handleFloatingArrowClick = () => {
    if (isProcessingClick) return;

    setIsProcessingClick(true);

    if (activeInput) {
      const { rowIndex, colIndex } = activeInput;
      const nextRow = rowIndex + 1;
      if (nextRow < actualData.length) {
        const nextInput = document.getElementById(
          `input-${nextRow}-${colIndex}`
        ) as HTMLInputElement | null;
        if (nextInput) {
          nextInput.focus();
          nextInput.select();
        }
      }
    }

    setTimeout(() => {
      setIsProcessingClick(false);
    }, 300);
  };

  const updateFloatingButtonPosition = (
    element: HTMLInputElement,
    rowIndex: number,
    colIndex: number,
    forceShow: boolean = true
  ) => {
    const rect = element.getBoundingClientRect();

    setFloatingButtonPosition({
      top: rect.top + rect.height / 2 - 28,
      left: rect.right + 10,
      visible: true,
    });
    setActiveInput({ rowIndex, colIndex });

    if (forceShow) {
      setShowFloatingButton(rowIndex < actualData.length - 1);
    }
  };

  const IMPORTABLE_COLUMNS = [
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
    "Data11",
    "Data12",
    "Data13",
    "Data14",
    "Data15",
    "Data16",
    "Data17",
    "Data18",
    "Data19",
    "Data22",
  ];

  const handleDownloadPDF = async () => {
    if (data.length === 0) return;
    setIsSaving(true);

    try {
      // Ambil data sekolah terbaru
      let schoolData: SchoolData | null = null;
      try {
        const schoolRes = await fetch(`${endpoint}?action=schoolData`);
        if (schoolRes.ok) {
          const schoolJson = await schoolRes.json();
          if (schoolJson.success && schoolJson.data?.length > 0) {
            schoolData = schoolJson.data[0];
          }
        }
      } catch (e) {
        console.warn("Gagal fetch schoolData:", e);
      }

      const doc = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "a4",
      });

      const pageW = 297;
      const margin = 10;

      const mapel = actualData[0]?.Data1 || "N/A";
      const kelas = actualData[0]?.Data3 || "N/A";
      const semester = actualData[0]?.Data2 || "N/A";

      // ─── JUDUL ───
      doc.setFont("helvetica", "bold");
      doc.setFontSize(13);
      doc.text("DAFTAR NILAI SISWA", pageW / 2, margin + 5, {
        align: "center",
      });

      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(
        `Mata Pelajaran: ${mapel}   |   Kelas: ${kelas}   |   Semester: ${semester}`,
        pageW / 2,
        margin + 11,
        { align: "center" }
      );

      // ─── KATEGORIKAN KOLOM ───
      // Cek apakah display header adalah kode TP (format angka.angka seperti "6.1")
      const isTPHeader = (headerDisplay: string): boolean => {
        // Hapus Zero Width Space (\u200B) dan karakter tersembunyi lainnya
        const cleaned = headerDisplay.replace(
          /[\u200B\u200C\u200D\uFEFF\s]/g,
          ""
        );
        return /^\d+\.\d+$/.test(cleaned);
      };

      // Cek apakah kolom adalah Sumatif Lingkup Materi (Data22)
      const isSLMHeader = (headerDisplay: string): boolean => {
        return (
          /^(BAB|Bab|bab)\s*\d+$/i.test(headerDisplay.trim()) ||
          /^\d+$/.test(headerDisplay.trim())
        );
      };

      // ─── HITUNG SPAN UNTUK MERGED HEADER ───
      // Kelompokkan kolom: no, nama, [TP group], [SLM group], [sisanya]
      let tpCount = 0;
      let slmCount = 0;
      let otherCount = 0; // kolom selain nama, TP, SLM

      const namaColIdx = 1; // index 0 = No, index 1 = Nama

      // Filter Data32 agar tidak muncul di PDF
      const pdfHeaders = visibleHeaders.filter(
        (h) => h !== "Data32" && h !== "Data20" && h !== "Data21"
      );

      const colCategories = pdfHeaders.map((header, idx) => {
        const dispHeader = data[0][header] || "";
        if (isTPHeader(dispHeader)) return "TP";
        if (isSLMHeader(dispHeader)) return "SLM";
        return "OTHER";
      });

      tpCount = colCategories.filter((c) => c === "TP").length;
      slmCount = colCategories.filter((c) => c === "SLM").length;
      otherCount = colCategories.filter((c) => c === "OTHER").length;

      // ─── HITUNG LEBAR KOLOM ───
      const noColW = 8;
      const namaColW = 38;
      const fixedColW = 16; // lebar fixed untuk Data20, Data21, Data22, Data23

      // Helper: lebar minimum kolom TP/SLM berdasarkan teks header
      const getMinTPColW = (headerText: string): number => {
        const cleanText = headerText
          .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
          .trim();
        doc.setFontSize(6.5);
        return doc.getTextWidth(cleanText) + 4;
      };

      // Lebar FIXED untuk kolom OTHER (tidak berubah)
      const otherFixedW: { [key: string]: number } = {
        Data22: fixedColW, // BAB/SLM fixed
        Data23: fixedColW, // kolom fixed
        Data24: 16, // NILAI_SAS
        Data25: 16, // NILAI_MURNI
      };
      const defaultOtherW = 22; // NILAI_RAPOR, RANKING, dst

      // Pisahkan kolom berdasarkan kategori
      const tpslmHeaders: string[] = [];
      const otherDynHeaders: string[] = [];

      pdfHeaders.forEach((header, idx) => {
        if (header === "Data4") return; // Nama sudah fixed
        const cat = colCategories[idx];
        if (cat === "TP" || cat === "SLM") {
          tpslmHeaders.push(header);
        } else {
          otherDynHeaders.push(header);
        }
      });

      // Hitung total lebar terpakai oleh kolom FIXED (No + Nama + OTHER)
      let totalFixedUsed = margin * 2 + noColW + namaColW;
      otherDynHeaders.forEach((header) => {
        totalFixedUsed += otherFixedW[header] ?? defaultOtherW;
      });

      // Sisa ruang PENUH untuk kolom TP/SLM
      const remainingForTPSLM = pageW - totalFixedUsed;

      // Hitung lebar minimum total TP/SLM
      let totalTPSLMMin = 0;
      tpslmHeaders.forEach((header) => {
        const dispH = (data[0][header] || "")
          .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
          .trim();
        totalTPSLMMin += getMinTPColW(dispH);
      });

      // Extra lebar dibagi rata ke semua kolom TP/SLM
      const extraPerTPSLM =
        tpslmHeaders.length > 0
          ? Math.max(
              0,
              (remainingForTPSLM - totalTPSLMMin) / tpslmHeaders.length
            )
          : 0;

      // Fungsi lebar per kolom
      const getColW = (header: string): number => {
        if (header === "Data4") return namaColW;
        const idx = pdfHeaders.indexOf(header);
        const cat = idx >= 0 ? colCategories[idx] : "OTHER";
        if (cat === "TP" || cat === "SLM") {
          const dispH = (data[0][header] || "")
            .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
            .trim();
          return getMinTPColW(dispH) + extraPerTPSLM; // ✅ melebar mengisi sisa ruang
        }
        return otherFixedW[header] ?? defaultOtherW; // ✅ OTHER tetap fixed
      };

      const baseColW = defaultOtherW;

      // ─── BUAT HEADER ROW 1 (merged group header) ───
      // Menggunakan pendekatan manual dengan jsPDF karena autoTable
      // tidak support rowspan/colspan secara native.
      // Kita gambar 2 baris header manual, lalu body dengan autoTable.

      const startY = margin + 15;
      const headerRow1H = 8; // tinggi baris group header
      const headerRow2H = 10; // tinggi baris sub header (TP codes)
      const bodyRowH = 5;

      // Hitung posisi X setiap kolom
      const colPositions: number[] = [];
      const colWidths: number[] = [];

      // Kolom No
      colPositions.push(margin);
      colWidths.push(noColW);

      // Kolom dari pdfHeaders
      let currentX = margin + noColW;
      pdfHeaders.forEach((header) => {
        colPositions.push(currentX);
        const w = getColW(header);
        colWidths.push(w);
        currentX += w;
      });

      // ✅ lastX dihitung dari currentX setelah semua kolom ditambahkan
      const computedLastX = currentX;

      // ─── GAMBAR HEADER ROW 1 (group labels) ───
      doc.setFontSize(7);
      doc.setFont("helvetica", "bold");

      // Hitung X dan W untuk group TP
      let tpStartX = 0;
      let tpTotalW = 0;
      let slmStartX = 0;
      let slmTotalW = 0;

      pdfHeaders.forEach((header, idx) => {
        const cat = colCategories[idx];
        const x = colPositions[idx + 1];
        const w = colWidths[idx + 1];

        if (cat === "TP") {
          if (tpStartX === 0) tpStartX = x;
          tpTotalW += w;
        }
        if (cat === "SLM") {
          if (slmStartX === 0) slmStartX = x;
          slmTotalW += w;
        }
      });

      // Kolom "No" - rowspan 2 (gambar manual)
      doc.setFillColor(41, 128, 185); // biru tua
      doc.setTextColor(255, 255, 255);
      doc.rect(margin, startY, noColW, headerRow1H + headerRow2H, "FD");
      doc.text(
        "NO",
        margin + noColW / 2,
        startY + (headerRow1H + headerRow2H) / 2 + 2,
        {
          align: "center",
        }
      );

      // Kolom "Nama Siswa" - rowspan 2
      const namaX = colPositions[1];
      doc.setFillColor(41, 128, 185);
      doc.rect(namaX, startY, namaColW, headerRow1H + headerRow2H, "FD");
      doc.text(
        "NAMA SISWA",
        namaX + namaColW / 2,
        startY + (headerRow1H + headerRow2H) / 2 + 2,
        {
          align: "center",
        }
      );

      // Group "SUMATIF TUJUAN PEMBELAJARAN" - hanya row 1
      if (tpCount > 0) {
        doc.setFillColor(0, 176, 240);
        doc.setTextColor(255, 0, 0);
        doc.rect(tpStartX, startY, tpTotalW, headerRow1H, "FD");
        doc.setFont("helvetica", "bold");

        // Auto-fit font size agar teks tidak melebihi lebar kolom
        const tpText = "SUMATIF TUJUAN PEMBELAJARAN";
        let tpFontSize = 7;
        doc.setFontSize(tpFontSize);
        while (doc.getTextWidth(tpText) > tpTotalW - 2 && tpFontSize > 3) {
          tpFontSize -= 0.5;
          doc.setFontSize(tpFontSize);
        }

        doc.text(
          tpText,
          tpStartX + tpTotalW / 2,
          startY + headerRow1H / 2 + 2,
          {
            align: "center",
          }
        );
      }

      // Group "SUMATIF LINGKUP MATERI" - hanya row 1
      if (slmCount > 0) {
        doc.setFillColor(0, 176, 240);
        doc.setTextColor(255, 0, 0);
        doc.rect(slmStartX, startY, slmTotalW, headerRow1H, "FD");
        doc.setFont("helvetica", "bold");

        // Auto-fit font size agar teks tidak melebihi lebar kolom
        const slmText = "SUMATIF LINGKUP MATERI";
        let slmFontSize = 7;
        doc.setFontSize(slmFontSize);
        while (doc.getTextWidth(slmText) > slmTotalW - 2 && slmFontSize > 3) {
          slmFontSize -= 0.5;
          doc.setFontSize(slmFontSize);
        }

        doc.text(
          slmText,
          slmStartX + slmTotalW / 2,
          startY + headerRow1H / 2 + 2,
          { align: "center" }
        );
      }

      // Kolom OTHER selain nama: rowspan 2
      pdfHeaders.forEach((header, idx) => {
        const cat = colCategories[idx];
        if (cat === "OTHER") {
          const x = colPositions[idx + 1];
          const w = colWidths[idx + 1];
          const dispH = data[0][header] || header;
          doc.setFillColor(41, 128, 185);
          doc.setTextColor(255, 255, 255);
          doc.rect(x, startY, w, headerRow1H + headerRow2H, "FD");
          doc.setFontSize(6);
          doc.setFont("helvetica", "bold");
          // Split teks jika panjang
          const lines = doc.splitTextToSize(dispH, w - 1);
          const textY =
            startY +
            (headerRow1H + headerRow2H) / 2 -
            ((lines.length - 1) * 3) / 2 +
            2;
          lines.forEach((line: string, li: number) => {
            doc.text(line, x + w / 2, textY + li * 3.5, { align: "center" });
          });
        }
      });

      // ─── GAMBAR HEADER ROW 2 (sub header: kode TP) ───
      const row2Y = startY + headerRow1H;

      pdfHeaders.forEach((header, idx) => {
        const cat = colCategories[idx];
        if (cat === "TP" || cat === "SLM") {
          const x = colPositions[idx + 1];
          const w = colWidths[idx + 1];
          // ✅ Bersihkan \u200B dari display header
          const dispH = (data[0][header] || header)
            .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
            .trim();

          // Warna bergantian untuk TP
          const colors: [number, number, number][] = [
            [255, 153, 204], // pink
            [255, 204, 153], // orange
            [153, 255, 153], // hijau
            [153, 204, 255], // biru
            [204, 153, 255], // ungu
            [255, 255, 153], // kuning
          ];
          const colorIdx = idx % colors.length;
          doc.setFillColor(...colors[colorIdx]);
          doc.setTextColor(0, 0, 0);
          doc.rect(x, row2Y, w, headerRow2H, "FD");
          doc.setFontSize(6.5);
          doc.setFont("helvetica", "bold");

          // Rotasi teks untuk header sempit
          doc.text(dispH, x + w / 2, row2Y + headerRow2H - 2, {
            align: "center",
          });
        }
      });

      // ─── GAMBAR GARIS BORDER HEADER ───
      doc.setDrawColor(100, 100, 100);
      doc.setLineWidth(0.3);

      const lastX = computedLastX;

      // Garis horizontal luar
      doc.line(margin, startY, lastX, startY);
      doc.line(
        margin,
        startY + headerRow1H + headerRow2H,
        lastX,
        startY + headerRow1H + headerRow2H
      );

      // Garis horizontal tengah (pemisah row1 dan row2)
      // Hanya digambar di area yang BUKAN merged (bukan area TP dan SLM di row1,
      // dan bukan area No/Nama yang rowspan 2)
      // Kita gambar per segmen:
      pdfHeaders.forEach((header, idx) => {
        const cat = colCategories[idx];
        const x = colPositions[idx + 1];
        const w = colWidths[idx + 1];
        // Garis tengah hanya untuk kolom TP dan SLM (karena kolom OTHER & No & Nama = rowspan 2)
        if (cat === "TP" || cat === "SLM") {
          doc.line(x, startY + headerRow1H, x + w, startY + headerRow1H);
        }
      });
      // Garis tengah untuk area No dan Nama tidak digambar (rowspan)

      // Garis vertikal - digambar per baris header agar tidak menembus merge
      // Baris header 1 (startY sampai startY + headerRow1H):
      // Gambar garis vertikal hanya di batas antar GROUP dan di tepi No/Nama

      // Kiri luar
      doc.line(margin, startY, margin, startY + headerRow1H + headerRow2H);
      // Kanan luar
      doc.line(lastX, startY, lastX, startY + headerRow1H + headerRow2H);

      // Garis kanan kolom No (full height karena rowspan)
      const noRightX = margin + noColW;
      doc.line(noRightX, startY, noRightX, startY + headerRow1H + headerRow2H);

      // Garis kanan kolom Nama (full height karena rowspan)
      const namaRightX = colPositions[1] + namaColW;
      doc.line(
        namaRightX,
        startY,
        namaRightX,
        startY + headerRow1H + headerRow2H
      );

      // Garis kanan grup TP (full height)
      if (tpCount > 0) {
        doc.line(
          tpStartX + tpTotalW,
          startY,
          tpStartX + tpTotalW,
          startY + headerRow1H + headerRow2H
        );
        // Garis kiri grup TP (full height)
        doc.line(
          tpStartX,
          startY,
          tpStartX,
          startY + headerRow1H + headerRow2H
        );
      }

      // Garis kanan grup SLM (full height)
      if (slmCount > 0) {
        doc.line(
          slmStartX + slmTotalW,
          startY,
          slmStartX + slmTotalW,
          startY + headerRow1H + headerRow2H
        );
      }

      // Garis vertikal antar kolom DALAM grup TP - hanya di baris row2
      pdfHeaders.forEach((header, idx) => {
        const cat = colCategories[idx];
        if (cat === "TP") {
          const x = colPositions[idx + 1];
          // Jangan gambar garis kiri paling awal grup (sudah digambar di atas)
          if (x > tpStartX) {
            doc.line(
              x,
              startY + headerRow1H,
              x,
              startY + headerRow1H + headerRow2H
            );
          }
        }
        if (cat === "SLM") {
          const x = colPositions[idx + 1];
          if (x > slmStartX) {
            doc.line(
              x,
              startY + headerRow1H,
              x,
              startY + headerRow1H + headerRow2H
            );
          }
        }
      });

      // Garis vertikal kolom OTHER - full height (karena rowspan 2)
      pdfHeaders.forEach((header, idx) => {
        const cat = colCategories[idx];
        if (cat === "OTHER") {
          const x = colPositions[idx + 1];
          doc.line(x, startY, x, startY + headerRow1H + headerRow2H);
          const xRight = x + colWidths[idx + 1];
          doc.line(xRight, startY, xRight, startY + headerRow1H + headerRow2H);
        }
      });

      // ─── GAMBAR BODY DATA ───
      doc.setTextColor(0, 0, 0);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7);

      const bodyStartY = startY + headerRow1H + headerRow2H;
      let currentY = bodyStartY;

      actualData.forEach((row, rowIndex) => {
        // Cek apakah perlu halaman baru
        if (currentY + bodyRowH > 210 - margin) {
          doc.addPage();
          currentY = margin;
        }

        const bgColor: [number, number, number] =
          rowIndex % 2 === 0 ? [224, 255, 255] : [255, 255, 255];

        // Gambar background baris
        doc.setFillColor(...bgColor);
        doc.rect(margin, currentY, lastX - margin, bodyRowH, "F");

        // No
        doc.text(
          String(rowIndex + 1),
          margin + noColW / 2,
          currentY + bodyRowH / 2 + doc.getFontSize() / 6,
          { align: "center" }
        );

        // Data kolom
        pdfHeaders.forEach((header, idx) => {
          const x = colPositions[idx + 1];
          const w = colWidths[idx + 1];
          const value =
            row[header] !== undefined && row[header] !== null
              ? String(row[header])
              : "";
          const isNama = header === "Data4";

          doc.text(
            value,
            isNama ? x + 1 : x + w / 2,
            currentY + bodyRowH / 2 + doc.getFontSize() / 6,
            { align: isNama ? "left" : "center", maxWidth: w - 1 }
          );
        });

        // Garis horizontal baris
        doc.setDrawColor(180, 180, 180);
        doc.setLineWidth(0.2);
        doc.line(margin, currentY + bodyRowH, lastX, currentY + bodyRowH);

        // Garis vertikal
        colPositions.forEach((x) => {
          doc.line(x, currentY, x, currentY + bodyRowH);
        });
        doc.line(lastX, currentY, lastX, currentY + bodyRowH);

        currentY += bodyRowH;
      });

      // ─── FOOTER ───
      doc.setFont("helvetica", "italic");
      doc.setFontSize(8);
      doc.setTextColor(100, 100, 100);
      doc.text(`Total siswa: ${actualData.length}`, margin, currentY + 5);

      // ─── AMBIL DATA TP DARI INDEXEDDB ───
      let tpTableData: { tp: string; rincian: string; bab: string }[] = [];
      try {
        const mapelName = actualData[0]?.Data1 || "";
        const kelasName = (actualData[0]?.Data3 || "").replace(/[^0-9]/g, "");
        const semesterName = actualData[0]?.Data2 || "";

        const tpCached = await idbLoad(STORE_TP);
        if (tpCached && tpCached.length > 1) {
          const tpRows = tpCached.slice(1);

          // ─── DEBUG: tampilkan 3 baris pertama TP ───
          console.log("🔎 ALL KEYS di row pertama:", Object.keys(tpRows[0]));
          console.log("🔎 Full row pertama:", JSON.stringify(tpRows[0]));

          tpTableData = tpRows
            .filter((row: any) => {
              const rowMapel = (row.Data1 || "")
                .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
                .trim();
              const rowKelas = String(row.Data6 ?? "").trim();
              const rowSemester = String(row.Data5 ?? "").trim();
              const semesterAngka = String(
                parseInt(semesterName) || semesterName
              );
              const kelasAngka = kelasName.replace(/[^0-9]/g, "");
              return (
                rowMapel.toLowerCase() === mapelName.toLowerCase() &&
                rowKelas === kelasAngka &&
                rowSemester === semesterAngka
              );
            })
            .map((row: any) => ({
              bab: String(row.Data4 || "")
                .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
                .trim(),
              tp: String(row.Data2 || "")
                .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
                .trim(),
              rincian: String(row.Data3 || "")
                .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
                .trim(),
            }))
            .filter((item: any) => item.tp && item.rincian)
            .sort((a: any, b: any) => {
              // Urutkan berdasarkan BAB dulu, lalu nomor TP
              const babA = parseFloat(a.bab) || 0;
              const babB = parseFloat(b.bab) || 0;
              if (babA !== babB) return babA - babB;
              const [, subA] = a.tp.split(".").map(Number);
              const [, subB] = b.tp.split(".").map(Number);
              return (subA || 0) - (subB || 0);
            });
          console.log(
            `✅ DataTP dari IndexedDB: ${tpTableData.length} TP ditemukan untuk ${mapelName} Kelas ${kelasName} Sem ${semesterName}`
          );
        } else {
          console.warn(
            "⚠️ IndexedDB STORE_TP kosong, tabel TP tidak akan ditampilkan"
          );
        }
      } catch (e) {
        console.warn("Gagal load DataTP dari IndexedDB:", e);
      }

      // ─── TABEL DAFTAR TP ───
      if (tpTableData.length > 0) {
        // Cek apakah perlu halaman baru
        const tpTableHeight = 10 + tpTableData.length * 12 + 10; // estimasi tinggi
        if (currentY + tpTableHeight + 20 > 210 - margin) {
          doc.addPage();
          currentY = margin + 5;
        } else {
          currentY += 8;
        }

        // Judul tabel TP
        doc.setFont("helvetica", "bold");
        doc.setFontSize(9);
        doc.setTextColor(0, 0, 0);
        doc.text("DAFTAR TUJUAN PEMBELAJARAN (TP)", pageW / 2, currentY, {
          align: "center",
        });
        currentY += 6;

        // Header tabel TP
        const tpNoW = 8;
        const tpKodeW = 18;
        const tpRincianW = pageW - margin * 2 - tpNoW - tpKodeW;
        const tpHeaderH = 7;

        doc.setFont("helvetica", "bold");
        doc.setFontSize(7);

        // Header: No
        doc.setFillColor(41, 128, 185);
        doc.setTextColor(255, 255, 255);
        doc.rect(margin, currentY, tpNoW, tpHeaderH, "FD");
        doc.text("No", margin + tpNoW / 2, currentY + tpHeaderH / 2 + 2, {
          align: "center",
        });

        // Header: Kode TP
        doc.setFillColor(41, 128, 185);
        doc.setTextColor(255, 255, 255);
        doc.rect(margin + tpNoW, currentY, tpKodeW, tpHeaderH, "FD");
        doc.text(
          "TP",
          margin + tpNoW + tpKodeW / 2,
          currentY + tpHeaderH / 2 + 2,
          { align: "center" }
        );

        // Header: Rincian TP
        doc.setFillColor(41, 128, 185);
        doc.setTextColor(255, 255, 255);
        doc.rect(
          margin + tpNoW + tpKodeW,
          currentY,
          tpRincianW,
          tpHeaderH,
          "FD"
        );
        doc.text(
          "Rincian Tujuan Pembelajaran",
          margin + tpNoW + tpKodeW + tpRincianW / 2,
          currentY + tpHeaderH / 2 + 2,
          { align: "center" }
        );

        currentY += tpHeaderH;

        // Body tabel TP
        doc.setTextColor(0, 0, 0);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7);

        tpTableData.forEach((tpItem, idx) => {
          // Hitung tinggi baris berdasarkan panjang teks rincian
          const rincianLines = doc.splitTextToSize(
            String(tpItem.rincian),
            tpRincianW - 3
          );
          const rowH = Math.max(7, rincianLines.length * 4 + 3);

          // Cek halaman baru
          if (currentY + rowH > 210 - margin) {
            doc.addPage();
            currentY = margin;
          }

          const bgColor: [number, number, number] =
            idx % 2 === 0 ? [240, 248, 255] : [255, 255, 255];
          doc.setFillColor(...bgColor);

          // Background
          doc.rect(margin, currentY, pageW - margin * 2, rowH, "F");

          // Border
          doc.rect(margin, currentY, tpNoW, rowH, "S");
          doc.rect(margin + tpNoW, currentY, tpKodeW, rowH, "S");
          doc.rect(margin + tpNoW + tpKodeW, currentY, tpRincianW, rowH, "S");

          const textY = currentY + rowH / 2 + 1.5;

          // No
          doc.text(String(idx + 1), margin + tpNoW / 2, textY, {
            align: "center",
          });

          // Kode TP
          doc.setFont("helvetica", "bold");
          doc.text(String(tpItem.tp), margin + tpNoW + tpKodeW / 2, textY, {
            align: "center",
          });
          doc.setFont("helvetica", "normal");

          // Rincian TP (multi-line)
          const rincianX = margin + tpNoW + tpKodeW + 2;
          const rincianStartY = currentY + 4;
          rincianLines.forEach((line: string, lineIdx: number) => {
            doc.text(line, rincianX, rincianStartY + lineIdx * 4);
          });

          currentY += rowH;
        });
      }

      // ─── TANDA TANGAN ───
      doc.setTextColor(0, 0, 0);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(8);

      const ttdY = currentY + 12;
      const kepsekX = margin;
      const guruX = pageW - margin - 50;

      // Tanggal (di atas kolom guru)
      const tanggalRapor = actualData[0]?.Data_tanggal || "";
      doc.text(
        `${
          schoolData?.tanggalRapor ? `Bungeng, ${schoolData.tanggalRapor}` : ""
        }`,
        guruX,
        ttdY
      );

      // Label jabatan
      doc.text("Kepala Sekolah,", kepsekX, ttdY + 5);
      doc.text("Guru Kelas,", guruX, ttdY + 5);

      // TTD Kepsek
      if (schoolData?.ttdKepsek) {
        try {
          doc.addImage(schoolData.ttdKepsek, "PNG", kepsekX, ttdY + 7, 30, 15);
        } catch (e) {
          console.warn("Gagal load TTD Kepsek:", e);
        }
      }

      // TTD Guru
      if (schoolData?.ttdGuru) {
        try {
          doc.addImage(schoolData.ttdGuru, "PNG", guruX, ttdY + 7, 30, 15);
        } catch (e) {
          console.warn("Gagal load TTD Guru:", e);
        }
      }

      // Nama dan NIP
      const namaKepsek = schoolData?.namaKepsek || "_______________";
      const nipKepsek = schoolData?.nipKepsek || "_______________";
      const namaGuru = schoolData?.namaGuru || "_______________";
      const nipGuru = schoolData?.nipGuru || "_______________";

      doc.setFont("helvetica", "bold");
      doc.text(namaKepsek, kepsekX, ttdY + 25);
      doc.setFont("helvetica", "normal");
      doc.text(`NIP. ${nipKepsek}`, kepsekX, ttdY + 30);

      doc.setFont("helvetica", "bold");
      doc.text(namaGuru, guruX, ttdY + 25);
      doc.setFont("helvetica", "normal");
      doc.text(`NIP. ${nipGuru}`, guruX, ttdY + 30);

      const fileName =
        `Nilai_${mapel}_Kelas${kelas}_Sem${semester}.pdf`.replace(/\s+/g, "_");
      setIsSaving(false);
      setIsSaving(false);
      doc.save(fileName);
    } catch (err) {
      console.error("Error generating PDF:", err);
      alert(
        "❌ Gagal membuat PDF: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const handleImportExcel = (
    e: React.ChangeEvent<HTMLInputElement>,
    targetColumn: string
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = "";

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const arrayBuffer = evt.target?.result as ArrayBuffer;

        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, {
          defval: "",
        });

        if (jsonData.length === 0) {
          alert("⚠️ File Excel kosong.");
          return;
        }

        const keys = Object.keys(jsonData[0]);
        const namaKey = keys.find(
          (k) => k.toLowerCase().includes("nama") || k.toLowerCase() === "name"
        );
        const nilaiKey = keys.find(
          (k) =>
            k.toLowerCase().includes("nilai") ||
            k.toLowerCase().includes("score")
        );

        if (!namaKey || !nilaiKey) {
          alert(
            `⚠️ Kolom tidak ditemukan!\n\nKolom tersedia: ${keys.join(
              ", "
            )}\n\nFile harus punya kolom "Nama" dan "Nilai".`
          );
          return;
        }

        const excelMap: { [nama: string]: string } = {};
        jsonData.forEach((row) => {
          const nama = String(row[namaKey] || "").trim();
          const nilai = String(row[nilaiKey] || "").trim();
          if (nama) excelMap[nama.toLowerCase()] = nilai;
        });

        const updatedData = [...data];
        const notFound: string[] = [];
        let matchedCount = 0;
        const newChangedRows = new Set(Array.from(changedRows));

        actualData.forEach((row, rowIndex) => {
          const namaSiswa = String(row.Data4 || "").trim();
          if (!namaSiswa) return;
          const matchedNilai = excelMap[namaSiswa.toLowerCase()];
          if (matchedNilai !== undefined) {
            updatedData[rowIndex + 1] = {
              ...updatedData[rowIndex + 1],
              [targetColumn]: matchedNilai,
            };
            newChangedRows.add(rowIndex);
            matchedCount++;
          } else {
            notFound.push(namaSiswa);
          }
        });

        setData(updatedData);
        setChangedRows(newChangedRows);
        setImportResult({ matched: matchedCount, notFound, targetColumn });
      } catch (err) {
        alert(
          "❌ Gagal membaca file: " +
            (err instanceof Error ? err.message : "Unknown error")
        );
      }
    };
    reader.readAsArrayBuffer(file);
  };

  if (loading)
    return (
      <div
        style={{
          textAlign: "center",
          fontSize: "18px",
          color: "#666",
          padding: "20px",
        }}
      >
        Loading...
      </div>
    );
  if (error)
    return (
      <div
        style={{
          textAlign: "center",
          fontSize: "18px",
          color: "red",
          padding: "20px",
        }}
      >
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div
        style={{
          textAlign: "center",
          fontSize: "18px",
          color: "#666",
          padding: "20px",
        }}
      >
        No data available
      </div>
    );

  const headers = [
    "Data1",
    "Data2",
    "Data3",
    "Data4",
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
    "Data11",
    "Data12",
    "Data13",
    "Data14",
    "Data15",
    "Data16",
    "Data17",
    "Data18",
    "Data19",
    "Data20",
    "Data21",
    "Data22",
    "Data23",
    "Data24",
    "Data25",
    "Data26",
    "Data27",
    "Data28",
    "Data29",
    "Data30",
    "Data31",
    "Data32",
  ];

  const displayHeaders = headers.map((header) => data[0][header] || "");

  const readOnlyHeaders = new Set([
    "Data1",
    "Data2",
    "Data3",
    "Data4",
    "Data20",
    "Data21",
    "Data23",
    "Data24",
    "Data25",
  ]);

  const conditionalHeaders = [
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
    "Data11",
    "Data12",
    "Data13",
    "Data14",
    "Data15",
    "Data16",
    "Data17",
    "Data18",
    "Data19",
    "Data22",
  ];

  const fixedWidthHeaders = new Set([
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
    "Data11",
    "Data12",
    "Data13",
    "Data14",
    "Data15",
    "Data16",
    "Data17",
    "Data18",
    "Data19",
    "Data20",
    "Data21",
    "Data22",
    "Data23",
  ]);

  const frozenHeaders = new Set(["Data4"]);
  const hiddenHeaders = new Set([
    "Data1",
    "Data2",
    "Data3",
    "Data20",
    "Data21",
    "Data23", // ← Nilai Murni
    "Data26",
    "Data27",
    "Data28",
    "Data29",
    "Data30",
    "Data31",
    "Data32", // ← Ket.
  ]);

  const visibleHeaders = headers.filter((header, index) => {
    if (hiddenHeaders.has(header)) {
      return false;
    }

    if (conditionalHeaders.indexOf(header) !== -1) {
      return displayHeaders[index] !== "-";
    }
    return true;
  });

  const visibleDisplayHeaders = visibleHeaders.map(
    (header) => data[0][header] || ""
  );

  const actualData = data.slice(1);

  const getColumnWidth = (header: string): string => {
    if (header === "Data4") return "120px";
    if (header === "Data20") return "120px";
    if (header === "Data21") return "100px";
    if (header === "Data22") return "100px";
    if (header === "Data23") return "100px";
    if (header === "Data32") return "120px";
    if (fixedWidthHeaders.has(header)) return "50px";
    return "90px";
  };

  const getFrozenLeftPosition = (header: string): number => {
    if (header === "Data4") {
      return 80;
    }
    return 0;
  };

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "8px",
          fontSize: "16px",
          marginTop: "8px",
        }}
      >
        ✏️ Input Nilai
      </h1>

      {/* Dropdown Pilih Sheet */}
      <div
        style={{ textAlign: "center", marginBottom: "8px", padding: "0 10px" }}
      >
        <label
          style={{
            fontSize: "11px",
            color: "#666",
            display: "block",
            marginBottom: "2px",
          }}
        >
          Pilih Mapel:
        </label>
        <select
          value={selectedSheet}
          onChange={(e) => handleSheetChange(e.target.value)}
          style={{
            padding: "10px 8px",
            fontSize: "13px", // lebih kecil agar muat
            borderRadius: "4px",
            border: "1px solid #ddd",
            width: "100%",
            maxWidth: "500px",
            cursor: "pointer",
            backgroundColor: "white",
            boxSizing: "border-box",
            WebkitAppearance: "none", // hilangkan style default iOS
            appearance: "none",
            backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%23333' d='M6 8L1 3h10z'/%3E%3C/svg%3E")`,
            backgroundRepeat: "no-repeat",
            backgroundPosition: "right 10px center",
            paddingRight: "30px",
          }}
        >
          <optgroup label="── Semester 1 ──">
            {availableSheets
              .filter((sheet) => sheet.semester === "1")
              .map((sheet, index) => (
                <option key={`s1-${index}`} value={sheet.sheetName}>
                  {sheet.mapel.length > 15
                    ? sheet.mapel.substring(0, 15) + "…"
                    : sheet.mapel}{" "}
                  · {sheet.kelas}
                </option>
              ))}
          </optgroup>
          <optgroup label="── Semester 2 ──">
            {availableSheets
              .filter((sheet) => sheet.semester === "2")
              .map((sheet, index) => (
                <option key={`s2-${index}`} value={sheet.sheetName}>
                  {sheet.mapel.length > 15
                    ? sheet.mapel.substring(0, 15) + "…"
                    : sheet.mapel}{" "}
                  · {sheet.kelas}
                </option>
              ))}
          </optgroup>
        </select>
      </div>

      {/* Info Sheet + Tombol dalam satu baris */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          marginBottom: "10px",
          padding: "8px 12px",
          backgroundColor: "#f8f9fa",
          borderRadius: "8px",
          border: "1px solid #e0e0e0",
          flexWrap: "wrap",
          gap: "8px",
        }}
      >
        {/* Info kiri */}
        <div style={{ fontSize: "13px", color: "#333", fontWeight: "500" }}>
          <span style={{ color: "#1565c0", fontWeight: "bold" }}>
            {actualData[0]?.Data1 || "N/A"}
          </span>
          <span style={{ color: "#666", margin: "0 6px" }}>|</span>
          <span>
            Kelas: <strong>{actualData[0]?.Data3 || "N/A"}</strong>
          </span>
          <span style={{ color: "#666", margin: "0 6px" }}>|</span>
          <span>
            Sem: <strong>{actualData[0]?.Data2 || "N/A"}</strong>
          </span>
          <span
            style={{
              marginLeft: "8px",
              backgroundColor: "#e3f2fd",
              color: "#1565c0",
              padding: "2px 8px",
              borderRadius: "12px",
              fontSize: "12px",
              fontWeight: "bold",
            }}
          >
            {actualData.length} siswa
          </span>
        </div>

        {/* Tombol kanan */}
        <div style={{ display: "flex", gap: "6px", flexWrap: "wrap" }}>
          <button
            onClick={handleSaveAll}
            disabled={isSaving}
            style={{
              padding: "6px 12px",
              backgroundColor: isSaving ? "#ccc" : "#4CAF50",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "12px",
            }}
          >
            💾 Simpan ({changedRows.size})
          </button>
          <button
            onClick={() => {
              if (!isEditMode) {
                setOriginalData(JSON.parse(JSON.stringify(data)));
                setIsEditMode(true);
              } else {
                const confirm = window.confirm(
                  "⚠️ Batal edit?\n\nSemua perubahan yang belum disimpan akan dikembalikan ke nilai sebelumnya."
                );
                if (!confirm) return;
                setData(originalData);
                setChangedRows(new Set());
                setIsEditMode(false);
                setShowFloatingButton(false);
                setActiveInput(null);
              }
            }}
            disabled={isSaving}
            style={{
              padding: "6px 12px",
              backgroundColor: isEditMode ? "#FF9800" : "#2196F3",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "12px",
            }}
          >
            {isEditMode ? "❌ Batal" : "✏️ Edit"}
          </button>
          <button
            onClick={handleClearAllValues}
            disabled={isSaving}
            style={{
              padding: "6px 12px",
              backgroundColor: isSaving ? "#ccc" : "#f44336",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "12px",
            }}
          >
            🗑️ Hapus
          </button>
          <button
            onClick={() => setImportModalOpen(true)}
            disabled={isSaving}
            style={{
              padding: "6px 12px",
              backgroundColor: isSaving ? "#ccc" : "#9C27B0",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "12px",
            }}
          >
            📥 Import
          </button>
          <button
            onClick={handleDownloadPDF}
            disabled={isSaving}
            style={{
              padding: "6px 12px",
              backgroundColor: isSaving ? "#ccc" : "#FF5722",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "12px",
            }}
          >
            📄 PDF
          </button>
        </div>
      </div>

      {/* Table */}
      <div
        id="table-scroll-container"
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 250px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
          WebkitOverflowScrolling: "touch",
          transform: "translateZ(0)",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
            tableLayout: "fixed",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              <th
                style={{
                  padding: "8px 4px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "35px",
                  minWidth: "35px",
                  position: "sticky",
                  left: 0,
                  top: 0,
                  backgroundColor: "#f4f4f4",
                  zIndex: 3,
                  boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  fontSize: "12px",
                }}
              >
                No.
              </th>
              <th
                style={{
                  padding: "8px 4px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "45px",
                  minWidth: "45px",
                  position: "sticky",
                  left: "35px",
                  top: 0,
                  backgroundColor: "#f4f4f4",
                  zIndex: 3,
                  boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  fontSize: "12px",
                }}
              >
                Desc
              </th>
              {visibleDisplayHeaders.map((header, index) => {
                const currentHeader = visibleHeaders[index];
                const isFrozen = frozenHeaders.has(currentHeader);
                const leftPos = isFrozen
                  ? getFrozenLeftPosition(currentHeader)
                  : "auto";
                const colWidth = getColumnWidth(currentHeader);

                return (
                  <th
                    key={index}
                    onClick={(e) => {
                      if (
                        conditionalHeaders.indexOf(currentHeader) !== -1 &&
                        ["Data20", "Data21", "Data22", "Data23"].indexOf(
                          currentHeader
                        ) === -1 &&
                        displayHeaders[headers.indexOf(currentHeader)] !== "-"
                      ) {
                        const tpCode =
                          displayHeaders[headers.indexOf(currentHeader)];
                        const mapel = actualData[0]?.Data1 || "";
                        const kelasRaw = actualData[0]?.Data3 || "";
                        const kelas = kelasRaw.replace(/[^0-9]/g, ""); // "6A" → "6"
                        const semester = actualData[0]?.Data2 || "";
                        fetchTPDetails(tpCode, mapel, 0, kelas, semester);
                      }
                    }}
                    style={{
                      padding: "8px 4px",
                      textAlign: "center",
                      borderBottom: "2px solid #ddd",
                      fontWeight: "bold",
                      width: colWidth,
                      minWidth: colWidth,
                      maxWidth: colWidth,
                      position: "sticky",
                      left: isFrozen ? leftPos : "auto",
                      backgroundColor: "#f4f4f4",
                      zIndex: isFrozen ? 2 : 1,
                      boxShadow: isFrozen
                        ? "2px 0 5px rgba(0,0,0,0.1)"
                        : "none",
                      fontSize: "12px",
                      whiteSpace: "nowrap",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      cursor:
                        conditionalHeaders.indexOf(currentHeader) !== -1 &&
                        ["Data20", "Data21", "Data22", "Data23"].indexOf(
                          currentHeader
                        ) === -1 &&
                        displayHeaders[headers.indexOf(currentHeader)] !== "-"
                          ? "pointer"
                          : "default",
                      transition: "background-color 0.2s",
                    }}
                    onMouseEnter={(e) => {
                      if (
                        conditionalHeaders.indexOf(currentHeader) !== -1 &&
                        ["Data20", "Data21", "Data22", "Data23"].indexOf(
                          currentHeader
                        ) === -1 &&
                        displayHeaders[headers.indexOf(currentHeader)] !== "-"
                      ) {
                        (e.target as HTMLElement).style.backgroundColor =
                          "#e0e0e0";
                      }
                    }}
                    onMouseLeave={(e) => {
                      (e.target as HTMLElement).style.backgroundColor =
                        "#f4f4f4";
                    }}
                  >
                    {header}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "6px 4px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                    fontWeight: "bold",
                    color: "#666",
                    width: "35px",
                    minWidth: "35px",
                    position: "sticky",
                    left: 0,
                    backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                    zIndex: 2,
                    boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                    fontSize: "12px",
                  }}
                >
                  {rowIndex + 1}
                </td>
                <td
                  style={{
                    padding: "4px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                    width: "45px",
                    minWidth: "45px",
                    position: "sticky",
                    left: "35px",
                    backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                    zIndex: 2,
                    boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  }}
                >
                  <button
                    onClick={async () => {
                      setIsLoadingDesc(true);
                      const freshData = await reloadDataKehadiran();
                      setIsLoadingDesc(false);

                      if (freshData && freshData.length > rowIndex + 1) {
                        const freshRow = freshData[rowIndex + 1];
                        setSelectedStudentDesc({
                          nama: freshRow.Data4 || "",
                          descMin: freshRow.Data26 || "Tidak ada deskripsi",
                          descMax: freshRow.Data27 || "Tidak ada deskripsi",
                          tpMin: freshRow.Data28 || "-",
                          tpMax: freshRow.Data29 || "-",
                          nilaiMin: freshRow.Data30 || "-",
                          nilaiMax: freshRow.Data31 || "-",
                        });
                        setShowDescPopup(true);
                      }
                    }}
                    disabled={isLoadingDesc}
                    style={{
                      width: "100%",
                      padding: "6px",
                      backgroundColor: isLoadingDesc ? "#ccc" : "#2196F3",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: isLoadingDesc ? "not-allowed" : "pointer",
                      fontSize: "20px",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontWeight: "bold",
                    }}
                  >
                    {isLoadingDesc ? "⏳" : "±"}
                  </button>
                </td>
                {visibleHeaders.map((header, colIndex) => {
                  const isFrozen = frozenHeaders.has(header);
                  const leftPos = isFrozen
                    ? getFrozenLeftPosition(header)
                    : "auto";
                  const colWidth = getColumnWidth(header);

                  return (
                    <td
                      key={colIndex}
                      style={{
                        padding: "4px",
                        borderBottom: "1px solid #eee",
                        width: colWidth,
                        minWidth: colWidth,
                        maxWidth: colWidth,
                        position: isFrozen ? "sticky" : "static",
                        left: isFrozen ? leftPos : "auto",
                        backgroundColor: isFrozen
                          ? rowIndex % 2 === 0
                            ? "#fff"
                            : "#f9f9f9"
                          : "transparent",
                        zIndex: isFrozen ? 1 : 0,
                        boxShadow: isFrozen
                          ? "2px 0 5px rgba(0,0,0,0.1)"
                          : "none",
                      }}
                    >
                      {readOnlyHeaders.has(header) ? (
                        <div
                          style={{
                            padding: "4px 2px",
                            color: "#666",
                            fontWeight: "normal",
                            fontSize: "12px",
                            whiteSpace: "nowrap",
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            textAlign: header === "Data4" ? "left" : "center",
                          }}
                        >
                          {row[header] || ""}
                        </div>
                      ) : (
                        <input
                          id={`input-${rowIndex}-${colIndex}`}
                          type="text"
                          inputMode={header === "Data32" ? "text" : "decimal"}
                          pattern={header === "Data32" ? undefined : "[0-9]*"}
                          value={row[header] || ""}
                          disabled={!isEditMode}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          onKeyDown={(e) =>
                            handleKeyDown(
                              e,
                              rowIndex,
                              colIndex,
                              actualData.length
                            )
                          }
                          onFocus={(e) => {
                            if (!isEditMode) return;
                            e.target.select();
                            updateFloatingButtonPosition(
                              e.target,
                              rowIndex,
                              colIndex,
                              true
                            );
                          }}
                          onBlur={(e) => {
                            const relatedTarget =
                              e.relatedTarget as HTMLElement;
                            if (
                              !relatedTarget ||
                              relatedTarget.tagName !== "BUTTON"
                            ) {
                              setTimeout(() => {
                                const activeElement = document.activeElement;
                                const isInputFocused =
                                  activeElement?.tagName === "INPUT" &&
                                  activeElement?.id.startsWith("input-");
                                if (!isInputFocused) {
                                  setShowFloatingButton(false);
                                }
                              }, 150);
                            }
                          }}
                          style={{
                            width: "100%",
                            padding: "4px 2px",
                            border: isEditMode
                              ? "1px solid #ddd"
                              : "1px solid #eee",
                            borderRadius: "3px",
                            boxSizing: "border-box",
                            backgroundColor: isEditMode ? "white" : "#f5f5f5",
                            cursor: isEditMode ? "text" : "not-allowed",
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            fontSize: "12px",
                            textAlign: "center",
                            color: isEditMode ? "#000" : "#666",
                          }}
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {/* Popup TP Details */}
      {showTPPopup && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setShowTPPopup(false)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "20px",
              maxWidth: "600px",
              width: "90%",
              maxHeight: "80vh",
              overflowY: "auto",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "15px",
                borderBottom: "2px solid #4CAF50",
                paddingBottom: "10px",
              }}
            >
              <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                Rincian TP: {selectedTP}
              </h2>
              <button
                onClick={() => setShowTPPopup(false)}
                style={{
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  padding: "8px 16px",
                  cursor: "pointer",
                  fontSize: "14px",
                  fontWeight: "bold",
                }}
              >
                Tutup
              </button>
            </div>

            {loadingTP ? (
              <div
                style={{
                  textAlign: "center",
                  padding: "20px",
                  color: "#666",
                }}
              >
                Loading...
              </div>
            ) : tpDetails ? (
              <div>
                <div style={{ marginBottom: "15px" }}>
                  <strong style={{ color: "#4CAF50" }}>Mapel:</strong>{" "}
                  <span style={{ color: "#333" }}>
                    {tpDetails.mapel || "N/A"}
                  </span>
                </div>
                <div style={{ marginBottom: "15px" }}>
                  <strong style={{ color: "#4CAF50" }}>TP:</strong>{" "}
                  <span style={{ color: "#333" }}>{tpDetails.tp || "N/A"}</span>
                </div>
                <div style={{ marginBottom: "15px" }}>
                  <strong style={{ color: "#4CAF50" }}>BAB:</strong>{" "}
                  <span style={{ color: "#333" }}>
                    {tpDetails.bab || "N/A"}
                  </span>
                </div>
                <div style={{ marginBottom: "15px" }}>
                  <strong style={{ color: "#4CAF50" }}>Semester:</strong>{" "}
                  <span style={{ color: "#333" }}>
                    {tpDetails.semester || "N/A"}
                  </span>
                </div>
                <div style={{ marginBottom: "15px" }}>
                  <strong style={{ color: "#4CAF50" }}>Kelas:</strong>{" "}
                  <span style={{ color: "#333" }}>
                    {tpDetails.kelas || "N/A"}
                  </span>
                </div>
                <div>
                  <strong style={{ color: "#4CAF50" }}>Rincian TP:</strong>
                  <p
                    style={{
                      marginTop: "10px",
                      lineHeight: "1.6",
                      color: "#333",
                      backgroundColor: "#f9f9f9",
                      padding: "15px",
                      borderRadius: "4px",
                      border: "1px solid #e0e0e0",
                    }}
                  >
                    {tpDetails.rincian || "Tidak ada rincian"}
                  </p>
                </div>
              </div>
            ) : (
              <div
                style={{
                  textAlign: "center",
                  padding: "20px",
                  color: "#f44336",
                }}
              >
                Data tidak ditemukan
              </div>
            )}
          </div>
        </div>
      )}

      {/* Popup Deskripsi */}
      {showDescPopup && selectedStudentDesc && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setShowDescPopup(false)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "20px",
              maxWidth: "700px",
              width: "90%",
              maxHeight: "80vh",
              overflowY: "auto",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "15px",
                borderBottom: "2px solid #2196F3",
                paddingBottom: "10px",
              }}
            >
              <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                Deskripsi: {selectedStudentDesc.nama}
              </h2>
              <button
                onClick={() => setShowDescPopup(false)}
                style={{
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  padding: "8px 16px",
                  cursor: "pointer",
                  fontSize: "14px",
                  fontWeight: "bold",
                }}
              >
                Tutup
              </button>
            </div>

            {/* TP & Nilai */}
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: "15px",
                marginBottom: "20px",
              }}
            >
              <div
                style={{
                  backgroundColor: "#ffebee",
                  padding: "15px",
                  borderRadius: "8px",
                  border: "2px solid #f44336",
                }}
              >
                <h3
                  style={{
                    color: "#f44336",
                    fontSize: "14px",
                    marginBottom: "8px",
                    display: "flex",
                    alignItems: "center",
                    gap: "5px",
                  }}
                >
                  <span style={{ fontSize: "20px" }}>📉</span> TP Terendah
                </h3>
                <p
                  style={{
                    fontSize: "24px",
                    fontWeight: "bold",
                    color: "#c62828",
                    margin: 0,
                    textAlign: "center",
                  }}
                >
                  {selectedStudentDesc.tpMin}
                </p>
              </div>

              <div
                style={{
                  backgroundColor: "#e8f5e9",
                  padding: "15px",
                  borderRadius: "8px",
                  border: "2px solid #4CAF50",
                }}
              >
                <h3
                  style={{
                    color: "#4CAF50",
                    fontSize: "14px",
                    marginBottom: "8px",
                    display: "flex",
                    alignItems: "center",
                    gap: "5px",
                  }}
                >
                  <span style={{ fontSize: "20px" }}>📈</span> TP Tertinggi
                </h3>
                <p
                  style={{
                    fontSize: "24px",
                    fontWeight: "bold",
                    color: "#2e7d32",
                    margin: 0,
                    textAlign: "center",
                  }}
                >
                  {selectedStudentDesc.tpMax}
                </p>
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: "15px",
                marginBottom: "20px",
              }}
            >
              <div
                style={{
                  backgroundColor: "#fff3e0",
                  padding: "15px",
                  borderRadius: "8px",
                  border: "2px solid #ff9800",
                }}
              >
                <h3
                  style={{
                    color: "#ff9800",
                    fontSize: "14px",
                    marginBottom: "8px",
                    display: "flex",
                    alignItems: "center",
                    gap: "5px",
                  }}
                >
                  <span style={{ fontSize: "20px" }}>📊</span> Nilai Terendah
                </h3>
                <p
                  style={{
                    fontSize: "24px",
                    fontWeight: "bold",
                    color: "#e65100",
                    margin: 0,
                    textAlign: "center",
                  }}
                >
                  {selectedStudentDesc.nilaiMin}
                </p>
              </div>

              <div
                style={{
                  backgroundColor: "#e3f2fd",
                  padding: "15px",
                  borderRadius: "8px",
                  border: "2px solid #2196F3",
                }}
              >
                <h3
                  style={{
                    color: "#2196F3",
                    fontSize: "14px",
                    marginBottom: "8px",
                    display: "flex",
                    alignItems: "center",
                    gap: "5px",
                  }}
                >
                  <span style={{ fontSize: "20px" }}>🎯</span> Nilai Tertinggi
                </h3>
                <p
                  style={{
                    fontSize: "24px",
                    fontWeight: "bold",
                    color: "#1565c0",
                    margin: 0,
                    textAlign: "center",
                  }}
                >
                  {selectedStudentDesc.nilaiMax}
                </p>
              </div>
            </div>

            {/* Deskripsi */}
            <div style={{ marginBottom: "20px" }}>
              <h3
                style={{
                  color: "#ff9800",
                  fontSize: "16px",
                  marginBottom: "10px",
                  display: "flex",
                  alignItems: "center",
                  gap: "5px",
                }}
              >
                <span style={{ fontSize: "18px" }}>⚠️</span> Deskripsi Minimal
              </h3>
              <p
                style={{
                  lineHeight: "1.6",
                  color: "#333",
                  backgroundColor: "#fff3cd",
                  padding: "15px",
                  borderRadius: "4px",
                  border: "1px solid #ffc107",
                  margin: 0,
                }}
              >
                {selectedStudentDesc.descMin}
              </p>
            </div>

            <div>
              <h3
                style={{
                  color: "#4CAF50",
                  fontSize: "16px",
                  marginBottom: "10px",
                  display: "flex",
                  alignItems: "center",
                  gap: "5px",
                }}
              >
                <span style={{ fontSize: "18px" }}>✅</span> Deskripsi Maksimal
              </h3>
              <p
                style={{
                  lineHeight: "1.6",
                  color: "#333",
                  backgroundColor: "#d4edda",
                  padding: "15px",
                  borderRadius: "4px",
                  border: "1px solid #28a745",
                  margin: 0,
                }}
              >
                {selectedStudentDesc.descMax}
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Modal Pilih Kolom Import */}
      {importModalOpen && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0,0,0,0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setImportModalOpen(false)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "24px",
              maxWidth: "480px",
              width: "90%",
              maxHeight: "80vh",
              overflowY: "auto",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "16px",
                borderBottom: "2px solid #9C27B0",
                paddingBottom: "10px",
              }}
            >
              <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                📥 Import Nilai dari Excel
              </h2>
              <button
                onClick={() => setImportModalOpen(false)}
                style={{
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  padding: "8px 16px",
                  cursor: "pointer",
                  fontWeight: "bold",
                }}
              >
                Tutup
              </button>
            </div>
            <p
              style={{
                fontSize: "13px",
                color: "#666",
                marginBottom: "16px",
                lineHeight: "1.6",
              }}
            >
              Pilih kolom tujuan, lalu upload file Excel.
              <br />
              File harus memiliki kolom <strong>"Nama"</strong> dan{" "}
              <strong>"Nilai"</strong>.<br />
              Nama siswa akan dicocokkan otomatis (tidak case-sensitive).
            </p>
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                gap: "10px",
              }}
            >
              {IMPORTABLE_COLUMNS.filter((col) => {
                const idx = headers.indexOf(col);
                return (
                  idx !== -1 &&
                  displayHeaders[idx] &&
                  displayHeaders[idx] !== "-"
                );
              }).map((col) => {
                const idx = headers.indexOf(col);
                const label = displayHeaders[idx] || col;
                const inputId = `import-file-${col}`;
                return (
                  <label
                    key={col}
                    htmlFor={inputId}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "space-between",
                      padding: "10px 16px",
                      backgroundColor: "#f3e5f5",
                      borderRadius: "6px",
                      border: "1px solid #ce93d8",
                      cursor: "pointer",
                    }}
                    onMouseEnter={(e) =>
                      (e.currentTarget.style.backgroundColor = "#e1bee7")
                    }
                    onMouseLeave={(e) =>
                      (e.currentTarget.style.backgroundColor = "#f3e5f5")
                    }
                  >
                    <span
                      style={{
                        fontWeight: "bold",
                        color: "#6a1b9a",
                        fontSize: "14px",
                      }}
                    >
                      {label}
                    </span>
                    <span style={{ fontSize: "12px", color: "#888" }}>
                      Klik untuk upload →
                    </span>
                    <input
                      id={inputId}
                      type="file"
                      accept=".xlsx,.xls"
                      style={{ display: "none" }}
                      onChange={(e) => {
                        setImportModalOpen(false);
                        handleImportExcel(e, col);
                      }}
                    />
                  </label>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* Modal Hasil Import */}
      {importResult && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0,0,0,0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setImportResult(null)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "24px",
              maxWidth: "480px",
              width: "90%",
              maxHeight: "80vh",
              overflowY: "auto",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "16px",
                borderBottom: "2px solid #4CAF50",
                paddingBottom: "10px",
              }}
            >
              <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                ✅ Hasil Import
              </h2>
              <button
                onClick={() => setImportResult(null)}
                style={{
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  padding: "8px 16px",
                  cursor: "pointer",
                  fontWeight: "bold",
                }}
              >
                Tutup
              </button>
            </div>
            <div
              style={{
                marginBottom: "16px",
                padding: "12px",
                backgroundColor: "#e8f5e9",
                borderRadius: "6px",
                border: "1px solid #4CAF50",
              }}
            >
              <p style={{ margin: 0, fontSize: "15px", color: "#2e7d32" }}>
                ✅ <strong>{importResult.matched} siswa</strong> berhasil
                diimport ke kolom{" "}
                <strong>
                  {displayHeaders[headers.indexOf(importResult.targetColumn)] ||
                    importResult.targetColumn}
                </strong>
              </p>
              <p style={{ margin: "6px 0 0", fontSize: "12px", color: "#555" }}>
                Klik "Save All Changes" untuk menyimpan ke server.
              </p>
            </div>
            {importResult.notFound.length > 0 && (
              <div
                style={{
                  padding: "12px",
                  backgroundColor: "#fff3e0",
                  borderRadius: "6px",
                  border: "1px solid #ff9800",
                }}
              >
                <p
                  style={{
                    margin: "0 0 8px",
                    fontSize: "14px",
                    color: "#e65100",
                    fontWeight: "bold",
                  }}
                >
                  ⚠️ {importResult.notFound.length} siswa tidak ditemukan di
                  Excel:
                </p>
                <ul
                  style={{
                    margin: 0,
                    paddingLeft: "20px",
                    fontSize: "13px",
                    color: "#555",
                  }}
                >
                  {importResult.notFound.map((nama, i) => (
                    <li key={i}>{nama}</li>
                  ))}
                </ul>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Floating Arrow Button */}
      {showFloatingButton && (
        <button
          onClick={(e) => {
            e.preventDefault();
            e.stopPropagation();
            handleFloatingArrowClick();
          }}
          style={{
            position: "fixed",
            top: `${floatingButtonPosition.top}px`,
            left: `${floatingButtonPosition.left}px`,
            width: "56px",
            height: "56px",
            borderRadius: "50%",
            backgroundColor: "#4CAF50",
            color: "white",
            border: "none",
            cursor: "pointer",
            boxShadow: "0 4px 12px rgba(0,0,0,0.4)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: "28px",
            fontWeight: "bold",
            zIndex: 1001,
            transition: "all 0.2s ease",
            pointerEvents: floatingButtonPosition.visible ? "auto" : "none",
            opacity: floatingButtonPosition.visible ? 1 : 0,
            visibility: floatingButtonPosition.visible ? "visible" : "hidden",
            touchAction: "manipulation",
            WebkitTapHighlightColor: "transparent",
          }}
          onMouseEnter={(e) => {
            if (floatingButtonPosition.visible) {
              (e.target as HTMLButtonElement).style.backgroundColor = "#45a049";
              (e.target as HTMLButtonElement).style.transform = "scale(1.1)";
            }
          }}
          onMouseLeave={(e) => {
            (e.target as HTMLButtonElement).style.backgroundColor = "#4CAF50";
            (e.target as HTMLButtonElement).style.transform = "scale(1)";
          }}
        >
          ↓
        </button>
      )}
    </div>
  );
};

const DataSekolah = () => {
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [namaKepsek, setNamaKepsek] = useState("");
  const [nipKepsek, setNipKepsek] = useState("");
  const [namaGuru, setNamaGuru] = useState("");
  const [nipGuru, setNipGuru] = useState("");
  const [ttdKepsek, setTtdKepsek] = useState("");
  const [ttdGuru, setTtdGuru] = useState("");
  const [loading, setLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [isKepsekSigning, setIsKepsekSigning] = useState(false);
  const [isGuruSigning, setIsGuruSigning] = useState(false);
  const kepsekSigCanvas = useRef<SignatureCanvas>(null);
  const guruSigCanvas = useRef<SignatureCanvas>(null);
  const [namaSekolah, setNamaSekolah] = useState("");
  const [npsn, setNpsn] = useState("");
  const [alamatSekolah, setAlamatSekolah] = useState("");
  const [kodePos, setKodePos] = useState("");
  const [desaKelurahan, setDesaKelurahan] = useState("");
  const [kabKota, setKabKota] = useState("");
  const [provinsi, setProvinsi] = useState("");
  const [tahunPelajaran, setTahunPelajaran] = useState(""); // ✅ TAMBAHAN BARU
  const [tanggalRapor, setTanggalRapor] = useState(""); // ✅ TAMBAHAN BARU
  const [nilaiKKM, setNilaiKKM] = useState("");
  const [kelas, setKelas] = useState("");
  const [rombel, setRombel] = useState("");
  const [hasChanges, setHasChanges] = useState(false);
  const initialDataRef = useRef<any>(null);
  const [isEditing, setIsEditing] = useState(false);

  useEffect(() => {
    const loadSekolah = async () => {
      // ✅ Load dari IndexedDB dulu (instant)
      const cached = await idbLoad("sekolahData");
      if (cached) {
        setSchoolData(cached);
        setNamaSekolah(cached.namaSekolah || "");
        setNpsn(cached.npsn || "");
        setAlamatSekolah(cached.alamatSekolah || "");
        setKodePos(cached.kodePos || "");
        setDesaKelurahan(cached.desaKelurahan || "");
        setKabKota(cached.kabKota || "");
        setProvinsi(cached.provinsi || "");
        setTahunPelajaran(cached.tahunPelajaran || "");
        setTanggalRapor(cached.tanggalRapor || "");
        setNilaiKKM(cached.nilaiKKM || "");
        setKelas(cached.kelas || "");
        setRombel(cached.rombel || "");
        setNamaKepsek(cached.namaKepsek || "");
        setNipKepsek(cached.nipKepsek || "");
        setTtdKepsek(cached.ttdKepsek || "");
        setNamaGuru(cached.namaGuru || "");
        setNipGuru(cached.nipGuru || "");
        setTtdGuru(cached.ttdGuru || "");
        setLoading(false);
        // ✅ Simpan snapshot data awal
        initialDataRef.current = cached;
        console.log("✅ DataSekolah: dari IndexedDB");
      }

      // ✅ Background sync dari server
      try {
        const res = await fetch(`${endpoint}?action=schoolData`);
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        const data = await res.json();
        if (data.success && data.data && data.data.length > 0) {
          const record = data.data[0];
          setSchoolData(record);
          setNamaSekolah(record.namaSekolah || "");
          setNpsn(record.npsn || "");
          setAlamatSekolah(record.alamatSekolah || "");
          setKodePos(record.kodePos || "");
          setDesaKelurahan(record.desaKelurahan || "");
          setKabKota(record.kabKota || "");
          setProvinsi(record.provinsi || "");
          setTahunPelajaran(record.tahunPelajaran || "");
          setTanggalRapor(record.tanggalRapor || "");
          setNilaiKKM(record.nilaiKKM || "");
          setKelas(record.kelas || "");
          setRombel(record.rombel || "");
          setNamaKepsek(record.namaKepsek || "");
          setNipKepsek(record.nipKepsek || "");
          setTtdKepsek(record.ttdKepsek || "");
          setNamaGuru(record.namaGuru || "");
          setNipGuru(record.nipGuru || "");
          setTtdGuru(record.ttdGuru || "");
          // ✅ Simpan ke IndexedDB
          await idbSave("sekolahData", record);
          // ✅ Update snapshot dengan data terbaru dari server
          initialDataRef.current = record;
          setHasChanges(false);
          console.log("✅ DataSekolah: IndexedDB diperbarui dari server");
        }
      } catch (err) {
        console.error("Gagal fetch DataSekolah:", err);
      } finally {
        setLoading(false);
      }
    };

    loadSekolah();
  }, []);

  const handleSave = async () => {
    if (!namaSekolah || !namaKepsek || !nipKepsek || !namaGuru || !nipGuru) {
      alert(
        "⚠️ Nama Sekolah, Nama & NIP Kepsek, dan Nama & NIP Guru wajib diisi!"
      );
      return;
    }

    setIsSaving(true);

    const data: SchoolData = {
      namaSekolah,
      npsn,
      alamatSekolah,
      kodePos,
      desaKelurahan,
      kabKota,
      provinsi,
      tahunPelajaran, // ✅ TAMBAHAN BARU
      tanggalRapor, // ✅ TAMBAHAN BARU
      nilaiKKM,
      kelas,
      rombel,
      namaKepsek,
      nipKepsek,
      ttdKepsek: ttdKepsek || "",
      namaGuru,
      nipGuru,
      ttdGuru: ttdGuru || "",
    };

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify({
          action: "schoolData",
          ...data,
        }),
      });

      await response.text();

      const freshRes = await fetch(`${endpoint}?action=schoolData`);
      if (freshRes.ok) {
        const freshJson = await freshRes.json();
        if (freshJson.success && freshJson.data?.length > 0) {
          await idbSave("sekolahData", freshJson.data[0]);
          console.log("✅ IndexedDB diperbarui dengan data fresh dari server");
        }
      }
      // ✅ Update snapshot dengan data terbaru setelah save
      initialDataRef.current = data;
      alert("✅ Data sekolah berhasil diperbarui!");
      setHasChanges(false);
      setIsEditing(false);
      setIsSaving(false);
    } catch (error) {
      console.error("Error saving school data:", error);
      alert("❌ Gagal memperbarui data sekolah.");
      setIsSaving(false);
    }
  };

  const checkChanges = (field: string, value: string) => {
    const init = initialDataRef.current;
    if (!init) {
      setHasChanges(true);
      return;
    }
    const current = {
      namaSekolah,
      npsn,
      alamatSekolah,
      kodePos,
      desaKelurahan,
      kabKota,
      provinsi,
      tahunPelajaran,
      tanggalRapor,
      nilaiKKM,
      kelas,
      rombel,
      namaKepsek,
      nipKepsek,
      namaGuru,
      nipGuru,
      ttdKepsek,
      ttdGuru,
      [field]: value,
    };
    const changed = Object.keys(current).some(
      (k) => String((current as any)[k] || "") !== String(init[k] || "")
    );
    setHasChanges(changed);
  };

  // Handler functions untuk signature
  const handleClearKepsekSignature = () => kepsekSigCanvas.current?.clear();
  const handleClearGuruSignature = () => guruSigCanvas.current?.clear();

  const handleSaveKepsekSignature = () => {
    const signature = kepsekSigCanvas.current?.toDataURL("image/png");
    if (signature && !kepsekSigCanvas.current?.isEmpty()) {
      setTtdKepsek(signature);
      setIsKepsekSigning(false);
      checkChanges("ttdKepsek", signature);
    } else {
      alert("⚠️ Tanda tangan kepala sekolah kosong!");
    }
  };

  const handleSaveGuruSignature = () => {
    const signature = guruSigCanvas.current?.toDataURL("image/png");
    if (signature && !guruSigCanvas.current?.isEmpty()) {
      setTtdGuru(signature);
      setIsGuruSigning(false);
      checkChanges("ttdGuru", signature);
    } else {
      alert("⚠️ Tanda tangan guru kosong!");
    }
  };

  if (loading) {
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        Memuat data sekolah...
      </div>
    );
  }

  return (
    <div style={{ maxWidth: "800px", margin: "0 auto", padding: "20px" }}>
      <div
        style={{
          backgroundColor: "white",
          padding: "24px",
          borderRadius: "8px",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
        }}
      >
        <h2
          style={{
            fontSize: "24px",
            fontWeight: "bold",
            textAlign: "center",
            color: "#2563eb",
            marginBottom: "24px",
          }}
        >
          🏫 Data Sekolah
        </h2>

        <div style={{ marginBottom: "24px" }}>
          <h3
            style={{
              fontSize: "18px",
              fontWeight: "600",
              marginBottom: "12px",
              color: "#1e40af",
            }}
          >
            Informasi Sekolah
          </h3>

          <input
            type="text"
            placeholder="Nama Sekolah *"
            value={namaSekolah}
            onChange={(e) => {
              setNamaSekolah(e.target.value);
              checkChanges("namaSekolah", e.target.value);
            }}
            disabled={isSaving || !isEditing}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
            }}
          />

          <input
            type="text"
            placeholder="NPSN"
            value={npsn}
            onChange={(e) => {
              setNpsn(e.target.value);
              checkChanges("npsn", e.target.value);
            }}
            disabled={isSaving || !isEditing}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
            }}
          />

          <textarea
            placeholder="Alamat Sekolah"
            value={alamatSekolah}
            onChange={(e) => {
              setAlamatSekolah(e.target.value);
              checkChanges("alamatSekolah", e.target.value);
            }}
            disabled={isSaving || !isEditing}
            rows={2}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
              resize: "vertical",
            }}
          />

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "8px",
            }}
          >
            <input
              type="text"
              placeholder="Kode Pos"
              value={kodePos}
              onChange={(e) => {
                setKodePos(e.target.value);
                checkChanges("kodePos", e.target.value);
              }}
              disabled={isSaving || !isEditing}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />

            <input
              type="text"
              placeholder="Desa/Kelurahan"
              value={desaKelurahan}
              onChange={(e) => {
                setDesaKelurahan(e.target.value);
                checkChanges("desaKelurahan", e.target.value);
              }}
              disabled={isSaving || !isEditing}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />
          </div>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "8px",
              marginTop: "8px",
            }}
          >
            <input
              type="text"
              placeholder="Kabupaten/Kota"
              value={kabKota}
              onChange={(e) => {
                setKabKota(e.target.value);
                checkChanges("kabKota", e.target.value);
              }}
              disabled={isSaving || !isEditing}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />

            <input
              type="text"
              placeholder="Provinsi"
              value={provinsi}
              onChange={(e) => {
                setProvinsi(e.target.value);
                checkChanges("provinsi", e.target.value);
              }}
              disabled={isSaving || !isEditing}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />

            {/* ✅ TAMBAHAN BARU - Input Tahun Pelajaran */}
            <input
              type="text"
              placeholder="Tahun Pelajaran (contoh: 2024/2025)"
              value={tahunPelajaran}
              onChange={(e) => {
                setTahunPelajaran(e.target.value);
                checkChanges("tahunPelajaran", e.target.value);
              }}
              disabled={isSaving || !isEditing}
              style={{
                padding: "10px",
                border: "1px solid #ddd",
                borderRadius: "4px",
              }}
            />

            {/* Input Tanggal Rapor dengan Date Picker */}
            <div style={{ position: "relative", marginBottom: "8px" }}>
              <input
                type="date"
                value={
                  tanggalRapor
                    ? (() => {
                        // Convert dd/mm/yyyy to yyyy-mm-dd for input[type="date"]
                        const parts = tanggalRapor.split("/");
                        if (parts.length === 3) {
                          return `${parts[2]}-${parts[1]}-${parts[0]}`;
                        }
                        return "";
                      })()
                    : ""
                }
                onChange={(e) => {
                  // Convert yyyy-mm-dd to dd/mm/yyyy
                  const value = e.target.value;
                  if (value) {
                    const parts = value.split("-");
                    const formatted = `${parts[2]}/${parts[1]}/${parts[0]}`;
                    setTanggalRapor(formatted);
                    checkChanges("tanggalRapor", formatted);
                  }
                }}
                disabled={isSaving || !isEditing}
                style={{
                  padding: "10px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  width: "100%",
                }}
              />
              {/* Label helper untuk menunjukkan format yang disimpan */}
              {tanggalRapor && (
                <div
                  style={{
                    fontSize: "11px",
                    color: "#4CAF50",
                    marginTop: "4px",
                  }}
                >
                  ✓ Tersimpan: {tanggalRapor}
                </div>
              )}
            </div>
          </div>
        </div>

        <hr
          style={{
            margin: "24px 0",
            border: "none",
            borderTop: "1px solid #e5e7eb",
          }}
        />

        {/* Kepala Sekolah */}
        <div style={{ marginBottom: "24px" }}>
          <h3
            style={{
              fontSize: "18px",
              fontWeight: "600",
              marginBottom: "12px",
            }}
          >
            Kepala Sekolah
          </h3>
          <input
            type="text"
            placeholder="Nama Kepala Sekolah"
            value={namaKepsek}
            onChange={(e) => {
              setNamaKepsek(e.target.value);
              checkChanges("namaKepsek", e.target.value);
            }}
            disabled={isSaving || !isEditing}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
            }}
          />
          <input
            type="text"
            placeholder="NIP Kepala Sekolah"
            value={nipKepsek}
            onChange={(e) => {
              setNipKepsek(e.target.value);
              checkChanges("nipKepsek", e.target.value);
            }}
            disabled={isSaving || !isEditing}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "12px",
            }}
          />

          <p style={{ fontSize: "14px", color: "#666", marginBottom: "8px" }}>
            Tanda Tangan Kepala Sekolah
          </p>
          <div style={{ position: "relative" }}>
            <SignatureCanvas
              ref={kepsekSigCanvas}
              penColor="black"
              canvasProps={{
                style: {
                  width: "100%",
                  height: "200px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  opacity: !isKepsekSigning || isSaving ? 0.5 : 1,
                  pointerEvents: !isKepsekSigning || isSaving ? "none" : "auto",
                },
              }}
            />
            {!isKepsekSigning && (
              <div
                style={{
                  position: "absolute",
                  top: 0,
                  left: 0,
                  right: 0,
                  bottom: 0,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  backgroundColor: "rgba(200,200,200,0.3)",
                }}
              >
                <span style={{ color: "#666" }}>
                  Klik "Mulai Tanda Tangan" untuk mengaktifkan
                </span>
              </div>
            )}
          </div>
          <div style={{ display: "flex", gap: "8px", marginTop: "8px" }}>
            {!isKepsekSigning && (
              <button
                onClick={() => setIsKepsekSigning(true)}
                disabled={isSaving || !isEditing}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#10b981",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                }}
              >
                ✍️ Mulai Tanda Tangan
              </button>
            )}
            {isKepsekSigning && (
              <button
                onClick={handleSaveKepsekSignature}
                disabled={isSaving || !isEditing}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#3b82f6",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                }}
              >
                💾 Simpan Tanda Tangan
              </button>
            )}
            <button
              onClick={handleClearKepsekSignature}
              disabled={!isKepsekSigning || isSaving}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ef4444",
                color: "white",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer",
                opacity: !isKepsekSigning || isSaving ? 0.5 : 1,
              }}
            >
              🗑️ Hapus TTD
            </button>
          </div>
          {ttdKepsek && (
            <div style={{ marginTop: "12px" }}>
              <img
                src={ttdKepsek}
                alt="TTD Kepsek"
                style={{
                  maxWidth: "200px",
                  height: "80px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  display: "block",
                  marginBottom: "8px",
                }}
              />
              <button
                onClick={() => {
                  const confirm = window.confirm(
                    "⚠️ Hapus tanda tangan Kepala Sekolah?\n\nTanda tangan akan dihapus permanen setelah disimpan."
                  );
                  if (confirm) {
                    setTtdKepsek("");
                    kepsekSigCanvas.current?.clear();
                    checkChanges("ttdKepsek", "");
                  }
                }}
                disabled={isSaving || !isEditing}
                style={{
                  padding: "6px 14px",
                  backgroundColor: "#ef4444",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                  fontSize: "13px",
                }}
              >
                🗑️ Hapus TTD Tersimpan
              </button>
            </div>
          )}
        </div>

        {/* Guru - Similar structure */}
        <div style={{ marginBottom: "24px" }}>
          <h3
            style={{
              fontSize: "18px",
              fontWeight: "600",
              marginBottom: "12px",
            }}
          >
            Guru Kelas
          </h3>
          <input
            type="text"
            placeholder="Nama Guru"
            value={namaGuru}
            onChange={(e) => {
              setNamaGuru(e.target.value);
              checkChanges("namaGuru", e.target.value);
            }}
            disabled={true}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "8px",
            }}
          />
          <input
            type="text"
            placeholder="NIP Guru"
            value={nipGuru}
            onChange={(e) => {
              setNipGuru(e.target.value);
              checkChanges("nipGuru", e.target.value);
            }}
            disabled={true}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              marginBottom: "12px",
            }}
          />

          <p style={{ fontSize: "14px", color: "#666", marginBottom: "8px" }}>
            Tanda Tangan Guru
          </p>
          <div style={{ position: "relative" }}>
            <SignatureCanvas
              ref={guruSigCanvas}
              penColor="black"
              canvasProps={{
                style: {
                  width: "100%",
                  height: "200px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  opacity: !isGuruSigning || isSaving ? 0.5 : 1,
                  pointerEvents: !isGuruSigning || isSaving ? "none" : "auto",
                },
              }}
            />
            {!isGuruSigning && (
              <div
                style={{
                  position: "absolute",
                  top: 0,
                  left: 0,
                  right: 0,
                  bottom: 0,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  backgroundColor: "rgba(200,200,200,0.3)",
                }}
              >
                <span style={{ color: "#666" }}>
                  Klik "Mulai Tanda Tangan" untuk mengaktifkan
                </span>
              </div>
            )}
          </div>
          <div style={{ display: "flex", gap: "8px", marginTop: "8px" }}>
            {!isGuruSigning && (
              <button
                onClick={() => setIsGuruSigning(true)}
                disabled={isSaving || !isEditing}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#10b981",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                }}
              >
                ✍️ Mulai Tanda Tangan
              </button>
            )}
            {isGuruSigning && (
              <button
                onClick={handleSaveGuruSignature}
                disabled={isSaving || !isEditing}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#3b82f6",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                }}
              >
                💾 Simpan Tanda Tangan
              </button>
            )}
            <button
              onClick={handleClearGuruSignature}
              disabled={!isGuruSigning || isSaving}
              style={{
                padding: "8px 16px",
                backgroundColor: "#ef4444",
                color: "white",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer",
                opacity: !isGuruSigning || isSaving ? 0.5 : 1,
              }}
            >
              🗑️ Hapus TTD
            </button>
          </div>
          {ttdGuru && (
            <div style={{ marginTop: "12px" }}>
              <img
                src={ttdGuru}
                alt="TTD Guru"
                style={{
                  maxWidth: "200px",
                  height: "80px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  display: "block",
                  marginBottom: "8px",
                }}
              />
              <button
                onClick={() => {
                  const confirm = window.confirm(
                    "⚠️ Hapus tanda tangan Guru?\n\nTanda tangan akan dihapus permanen setelah disimpan."
                  );
                  if (confirm) {
                    setTtdGuru("");
                    guruSigCanvas.current?.clear();
                    checkChanges("ttdGuru", "");
                  }
                }}
                disabled={isSaving || !isEditing}
                style={{
                  padding: "6px 14px",
                  backgroundColor: "#ef4444",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                  fontSize: "13px",
                }}
              >
                🗑️ Hapus TTD Tersimpan
              </button>
            </div>
          )}
        </div>

        <div style={{ marginBottom: "24px" }}>
          <h3
            style={{
              fontSize: "18px",
              fontWeight: "600",
              marginBottom: "12px",
            }}
          >
            Kelas & Rombel
          </h3>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "8px",
            }}
          >
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontSize: "14px",
                  color: "#555",
                }}
              >
                Kelas
              </label>
              <select
                value={kelas}
                onChange={(e) => {
                  setKelas(e.target.value);
                  checkChanges("kelas", e.target.value);
                }}
                disabled={true}
                style={{
                  width: "100%",
                  padding: "10px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  fontSize: "14px",
                  backgroundColor: "white",
                }}
              >
                <option value="">-- Pilih Kelas --</option>
                {["1", "2", "3", "4", "5", "6"].map((k) => (
                  <option key={k} value={k}>
                    {k}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontSize: "14px",
                  color: "#555",
                }}
              >
                Rombel
              </label>
              <select
                value={rombel}
                onChange={(e) => {
                  setRombel(e.target.value);
                  checkChanges("rombel", e.target.value);
                }}
                disabled={true}
                style={{
                  width: "100%",
                  padding: "10px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  fontSize: "14px",
                  backgroundColor: "white",
                }}
              >
                <option value="">-- Pilih Rombel --</option>
                <option value="-">- (Tanpa Rombel)</option>
                <option value="A">A</option>
                <option value="B">B</option>
              </select>
            </div>
          </div>
        </div>

        <div style={{ marginBottom: "24px" }}>
          <h3
            style={{
              fontSize: "18px",
              fontWeight: "600",
              marginBottom: "12px",
            }}
          >
            Nilai KKM
          </h3>
          <input
            type="number"
            placeholder="Nilai KKM (contoh: 75)"
            value={nilaiKKM}
            onChange={(e) => {
              setNilaiKKM(e.target.value);
              checkChanges("nilaiKKM", e.target.value);
            }}
            disabled={isSaving || !isEditing}
            style={{
              width: "100%",
              padding: "10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
            }}
          />
        </div>

        <div
          style={{
            textAlign: "center",
            display: "flex",
            gap: "12px",
            justifyContent: "center",
          }}
        >
          <button
            onClick={() => {
              if (isEditing) {
                // Batal edit: kembalikan ke data awal
                const init = initialDataRef.current;
                if (init) {
                  setNamaSekolah(init.namaSekolah || "");
                  setNpsn(init.npsn || "");
                  setAlamatSekolah(init.alamatSekolah || "");
                  setKodePos(init.kodePos || "");
                  setDesaKelurahan(init.desaKelurahan || "");
                  setKabKota(init.kabKota || "");
                  setProvinsi(init.provinsi || "");
                  setTahunPelajaran(init.tahunPelajaran || "");
                  setTanggalRapor(init.tanggalRapor || "");
                  setNilaiKKM(init.nilaiKKM || "");
                  setKelas(init.kelas || "");
                  setRombel(init.rombel || "");
                  setNamaKepsek(init.namaKepsek || "");
                  setNipKepsek(init.nipKepsek || "");
                  setNamaGuru(init.namaGuru || "");
                  setNipGuru(init.nipGuru || "");
                  setTtdKepsek(init.ttdKepsek || "");
                  setTtdGuru(init.ttdGuru || "");
                }
                setHasChanges(false);
                setIsEditing(false);
              } else {
                setIsEditing(true);
              }
            }}
            disabled={isSaving}
            style={{
              padding: "12px 24px",
              backgroundColor: isEditing ? "#FF9800" : "#10b981",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "600",
              transition: "background-color 0.3s",
            }}
          >
            {isEditing ? "❌ Batal Edit" : "✏️ Edit Data"}
          </button>

          <button
            onClick={handleSave}
            disabled={isSaving || !hasChanges || !isEditing}
            style={{
              padding: "12px 24px",
              backgroundColor: isSaving
                ? "#93c5fd"
                : !hasChanges || !isEditing
                ? "#cbd5e1"
                : "#2563eb",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor:
                isSaving || !hasChanges || !isEditing
                  ? "not-allowed"
                  : "pointer",
              fontWeight: "600",
              transition: "background-color 0.3s",
            }}
          >
            {isSaving ? "⏳ Menyimpan..." : "💾 Simpan Data Sekolah"}
          </button>
        </div>
      </div>
    </div>
  );
};

interface RekapData {
  nama: string;
  kelas: string;
  nilaiMapel: { [mapel: string]: number | null }; // Nilai per mapel, dinamis
  jumlah: number;
  rataRata: number;
  ranking: number;
  catatan: string;
}

const DataKehadiran = () => {
  const { refreshRekapData } = useRekapData();
  const [selectedSemester, setSelectedSemester] = useState<string>("1");
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [showFloatingButton, setShowFloatingButton] = useState(false);
  const [floatingButtonPosition, setFloatingButtonPosition] = useState({
    top: 0,
    left: 0,
    visible: true,
  });
  const [activeInput, setActiveInput] = useState<{
    rowIndex: number;
    colIndex: number;
  } | null>(null);
  const [isProcessingClick, setIsProcessingClick] = useState(false);
  const [isEditMode, setIsEditMode] = useState(false);
  const [importResult, setImportResult] = useState<{
    matched: number;
    notFound: string[];
  } | null>(null);
  const [importPreview, setImportPreview] = useState<{
    matched: {
      nama: string;
      hadir: string;
      alpha: string;
      izin: string;
      sakit: string;
    }[];
    notFound: string[];
    pendingData: {
      rowIndex: number;
      hadir: string;
      alpha: string;
      izin: string;
      sakit: string;
    }[];
  } | null>(null);
  const [showStudentPopup, setShowStudentPopup] = useState(false);
  const [selectedStudent, setSelectedStudent] = useState<{
    nama: string;
    kelas: string;
    nisn: string;
  } | null>(null);

  // Selalu muat data langsung dari server saat mount
  useEffect(() => {
    setLoading(true);
    setData([]);

    const loadFromServer = async () => {
      try {
        const sheetName = `DataKehadiran${selectedSemester}`;
        const response = await fetch(`${endpoint}?sheet=${sheetName}`);
        if (!response.ok) throw new Error(`Failed to fetch ${sheetName}`);

        const jsonData = await response.json();
        setData(jsonData);
        setLoading(false);
      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
        setLoading(false);
      }
    };

    loadFromServer();
  }, [selectedSemester]);

  // Update posisi tombol floating
  useEffect(() => {
    const updateButtonPosition = () => {
      if (showFloatingButton && activeInput) {
        const { rowIndex, colIndex } = activeInput;
        const input = document.getElementById(
          `kehadiran-input-${rowIndex}-${colIndex}`
        ) as HTMLInputElement;

        if (input) {
          const rect = input.getBoundingClientRect();
          const tableContainer = document.getElementById(
            "kehadiran-table-container"
          );

          if (tableContainer) {
            const containerRect = tableContainer.getBoundingClientRect();
            const thead = tableContainer.querySelector("thead");
            const headerHeight = thead ? thead.offsetHeight : 40;

            const inputTopInContainer = rect.top - containerRect.top;
            const inputBottomInContainer = rect.bottom - containerRect.top;

            const isVisibleInContainer =
              inputTopInContainer >= headerHeight &&
              inputBottomInContainer > headerHeight &&
              rect.bottom <= containerRect.bottom &&
              rect.left >= containerRect.left - 100 &&
              rect.right <= window.innerWidth + 100;

            setFloatingButtonPosition({
              top: rect.top + rect.height / 2 - 28,
              left: rect.right + 10,
              visible: isVisibleInContainer,
            });
          }
        }
      }
    };

    const handleScroll = throttle(updateButtonPosition, 16);
    const tableContainer = document.getElementById("kehadiran-table-container");

    if (tableContainer) {
      tableContainer.addEventListener("scroll", handleScroll as any, {
        passive: true,
      });
    }

    window.addEventListener("scroll", handleScroll as any, { passive: true });

    return () => {
      if (tableContainer) {
        tableContainer.removeEventListener("scroll", handleScroll as any);
      }
      window.removeEventListener("scroll", handleScroll as any);
    };
  }, [showFloatingButton, activeInput]);

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleKeyDown = (
    e: React.KeyboardEvent<HTMLInputElement>,
    rowIndex: number,
    colIndex: number,
    actualDataLength: number
  ) => {
    if (e.key === "Enter") {
      e.preventDefault();
      const nextRow = rowIndex + 1;
      if (nextRow < actualDataLength) {
        const nextInput = document.getElementById(
          `kehadiran-input-${nextRow}-${colIndex}`
        ) as HTMLInputElement | null;
        if (nextInput) {
          nextInput.focus();
          nextInput.select();
        }
      }
    }
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);

    const headers = [
      "Data1",
      "Data2",
      "Data3",
      "Data4",
      "Data5",
      "Data6",
      "Data7",
    ];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({
        rowIndex: rowIndex + 3,
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_kehadiran_bulk",
        sheetName: `DataKehadiran${selectedSemester}`,
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());
      setIsEditMode(false);
      setShowFloatingButton(false);
      setActiveInput(null);
      setIsSaving(false);

      // Reload data langsung dari server
      const reloadResponse = await fetch(
        `${endpoint}?sheet=DataKehadiran${selectedSemester}`
      );
      if (reloadResponse.ok) {
        const reloadJson = await reloadResponse.json();
        setData(reloadJson);
      }

      refreshRekapData(true);
    } catch (err) {
      console.error("=== ERROR DETAILS ===");
      console.error(err);
      alert(
        "Error updating rows: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  const handleFloatingArrowClick = () => {
    if (isProcessingClick) return;
    setIsProcessingClick(true);

    if (activeInput) {
      const { rowIndex, colIndex } = activeInput;
      const nextRow = rowIndex + 1;
      if (nextRow < actualData.length) {
        const nextInput = document.getElementById(
          `kehadiran-input-${nextRow}-${colIndex}`
        ) as HTMLInputElement | null;
        if (nextInput) {
          nextInput.focus();
          nextInput.select();
        }
      }
    }

    setTimeout(() => {
      setIsProcessingClick(false);
    }, 300);
  };

  const handleImportKehadiran = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = "";

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, {
          defval: "",
        });

        if (jsonData.length === 0) {
          alert("⚠️ File Excel kosong.");
          return;
        }

        const keys = Object.keys(jsonData[0]);
        const namaKey = keys.find((k) => k.toLowerCase() === "nama");
        const hadirKey = keys.find((k) => k.toLowerCase() === "hadir");
        const alphaKey = keys.find(
          (k) => k.toLowerCase() === "alpha" || k.toLowerCase() === "alpa"
        );
        const izinKey = keys.find((k) => k.toLowerCase() === "izin");
        const sakitKey = keys.find((k) => k.toLowerCase() === "sakit");

        if (!namaKey) {
          alert(
            `⚠️ Kolom "Nama" tidak ditemukan!\n\nKolom tersedia: ${keys.join(
              ", "
            )}`
          );
          return;
        }

        // Buat map dari Excel
        const excelMap: {
          [nama: string]: {
            hadir: string;
            alpha: string;
            izin: string;
            sakit: string;
          };
        } = {};
        jsonData.forEach((row) => {
          const nama = String(row[namaKey] || "").trim();
          if (
            nama &&
            nama.toLowerCase() !== "total" &&
            nama.toLowerCase() !== "persen"
          ) {
            excelMap[nama] = {
              hadir: hadirKey ? String(row[hadirKey] || "") : "",
              alpha: alphaKey ? String(row[alphaKey] || "") : "",
              izin: izinKey ? String(row[izinKey] || "") : "",
              sakit: sakitKey ? String(row[sakitKey] || "") : "",
            };
          }
        });

        // Buat preview data
        const matchedPreview: {
          nama: string;
          hadir: string;
          alpha: string;
          izin: string;
          sakit: string;
        }[] = [];
        const notFound: string[] = [];
        const pendingData: {
          rowIndex: number;
          hadir: string;
          alpha: string;
          izin: string;
          sakit: string;
        }[] = [];

        actualData.forEach((row, rowIndex) => {
          const namaSiswa = String(row.Data1 || "").trim();
          if (!namaSiswa) return;

          const matchedNilai = excelMap[namaSiswa];
          if (matchedNilai !== undefined) {
            matchedPreview.push({ nama: namaSiswa, ...matchedNilai });
            pendingData.push({ rowIndex, ...matchedNilai });
          } else {
            notFound.push(namaSiswa);
          }
        });

        // Tampilkan preview, belum apply ke data
        setImportPreview({ matched: matchedPreview, notFound, pendingData });
      } catch (err) {
        alert(
          "❌ Gagal membaca file: " +
            (err instanceof Error ? err.message : "Unknown error")
        );
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleConfirmImport = async () => {
    if (!importPreview) return;

    setIsSaving(true);
    setImportPreview(null);

    const headers = [
      "Data1",
      "Data2",
      "Data3",
      "Data4",
      "Data5",
      "Data6",
      "Data7",
    ];

    // Terapkan perubahan ke data lokal dan siapkan updates untuk server
    const updatedData = [...data];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    importPreview.pendingData.forEach(
      ({ rowIndex, hadir, alpha, izin, sakit }) => {
        updatedData[rowIndex + 1] = {
          ...updatedData[rowIndex + 1],
          Data4: hadir,
          Data5: alpha,
          Data6: izin,
          Data7: sakit,
        };

        const rowData = updatedData[rowIndex + 1];
        const values = headers.map((header) => rowData[header] || "");

        updates.push({
          rowIndex: rowIndex + 3, // +3 karena: +1 header row, +1 label row, +1 untuk 1-based index
          values,
        });
      }
    );

    setData(updatedData);

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "update_kehadiran_bulk",
          sheetName: `DataKehadiran${selectedSemester}`,
          updates,
        }),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      setImportResult({
        matched: importPreview.matched.length,
        notFound: importPreview.notFound,
      });

      setChangedRows(new Set());

      // Reload data dari server
      const reloadResponse = await fetch(
        `${endpoint}?sheet=DataKehadiran${selectedSemester}`
      );
      if (reloadResponse.ok) {
        const reloadJson = await reloadResponse.json();
        setData(reloadJson);
      }

      refreshRekapData(true);
    } catch (err) {
      alert(
        "❌ Gagal menyimpan: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const updateFloatingButtonPosition = (
    element: HTMLInputElement,
    rowIndex: number,
    colIndex: number,
    forceShow: boolean = true
  ) => {
    const rect = element.getBoundingClientRect();

    setFloatingButtonPosition({
      top: rect.top + rect.height / 2 - 28,
      left: rect.right + 10,
      visible: true,
    });
    setActiveInput({ rowIndex, colIndex });

    if (forceShow) {
      setShowFloatingButton(rowIndex < actualData.length - 1);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const calcPersen = (row: any): string => {
    const hadir = parseFloat(row.Data4 || "0") || 0;
    const alpha = parseFloat(row.Data5 || "0") || 0;
    const izin = parseFloat(row.Data6 || "0") || 0;
    const sakit = parseFloat(row.Data7 || "0") || 0;
    const total = hadir + alpha + izin + sakit;
    if (total === 0) return "-";
    return ((hadir / total) * 100).toFixed(1) + "%";
  };

  const headers = [
    "Data1",
    "Data2",
    "Data3",
    "Data4",
    "Data5",
    "Data6",
    "Data7",
  ];
  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  // Data1=No, Data2=Kelas, Data3=Nama, Data4=Hadir, Data5=Alpha, Data6=Izin, Data7=Sakit
  const readOnlyHeaders = new Set(["Data1", "Data2", "Data3"]);
  const editableHeaders = ["Data4", "Data5", "Data6", "Data7"]; // Hadir, Alpha, Izin, Sakit
  const hiddenHeaders = new Set(["Data2", "Data3"]); // Data1 asumsikan NISN/No, Data2=Kelas
  const visibleHeaders = headers.filter((header) => !hiddenHeaders.has(header));
  const visibleDisplayHeaders = visibleHeaders.map(
    (header) => data[0][header] || ""
  );

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        📋 Data Kehadiran Siswa
      </h1>

      {/* Filter Semester */}
      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <label style={{ fontSize: "14px", color: "#666", marginRight: "10px" }}>
          Semester:
        </label>
        <select
          value={selectedSemester}
          onChange={(e) => {
            setSelectedSemester(e.target.value);
            setChangedRows(new Set());
            setIsEditMode(false);
          }}
          style={{
            padding: "10px 15px",
            fontSize: "16px",
            borderRadius: "4px",
            border: "1px solid #ddd",
            minWidth: "150px",
            cursor: "pointer",
            backgroundColor: "white",
          }}
        >
          <option value="1">Semester 1</option>
          <option value="2">Semester 2</option>
        </select>
      </div>

      {/* Tombol Save & Edit */}
      <div
        style={{
          textAlign: "center",
          marginBottom: "15px",
          display: "flex",
          gap: "10px",
          justifyContent: "center",
          flexWrap: "wrap",
        }}
      >
        <button
          onClick={handleSaveAll}
          disabled={isSaving || !isEditMode}
          style={{
            padding: "12px 24px",
            backgroundColor: isSaving || !isEditMode ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving || !isEditMode ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            maxWidth: "300px",
          }}
        >
          {isSaving ? "Memproses..." : `Simpan Perubahan (${changedRows.size})`}
        </button>
        <button
          onClick={() => {
            if (isEditMode) {
              const confirm = window.confirm(
                "⚠️ Batal edit?\n\nSemua perubahan yang belum disimpan akan dikembalikan."
              );
              if (!confirm) return;
              setChangedRows(new Set());
              setIsEditMode(false);
              setShowFloatingButton(false);
              setActiveInput(null);
            } else {
              setIsEditMode(true);
            }
          }}
          disabled={isSaving}
          style={{
            padding: "12px 24px",
            backgroundColor: isEditMode ? "#FF9800" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            maxWidth: "300px",
          }}
        >
          {isEditMode ? "❌ Batal Edit" : "✏️ Edit Nilai"}
        </button>
        <label
          htmlFor="import-kehadiran-file"
          style={{
            padding: "12px 24px",
            backgroundColor: "#9C27B0",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            maxWidth: "300px",
            display: "inline-block",
          }}
          onMouseEnter={(e) =>
            (e.currentTarget.style.backgroundColor = "#7B1FA2")
          }
          onMouseLeave={(e) =>
            (e.currentTarget.style.backgroundColor = "#9C27B0")
          }
        >
          📥 Import Excel
        </label>
        <input
          id="import-kehadiran-file"
          type="file"
          accept=".xlsx,.xls"
          style={{ display: "none" }}
          onChange={handleImportKehadiran}
        />
      </div>

      {/* Table */}
      <div
        id="kehadiran-table-container"
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 200px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
            tableLayout: "fixed",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              <th
                style={{
                  padding: "8px 4px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "40px",
                  minWidth: "40px",
                  position: "sticky",
                  left: 0,
                  backgroundColor: "#f4f4f4",
                  zIndex: 2,
                  boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  fontSize: "12px",
                }}
              >
                No.
              </th>
              {visibleDisplayHeaders.map((header, index) => {
                const currentHeader = visibleHeaders[index];
                return (
                  <th
                    key={index}
                    style={{
                      padding: "8px 4px",
                      textAlign: "center",
                      borderBottom: "2px solid #ddd",
                      fontWeight: "bold",
                      width: currentHeader === "Data1" ? "150px" : "60px",
                      minWidth: currentHeader === "Data1" ? "150px" : "60px",
                      position: currentHeader === "Data1" ? "sticky" : "static",
                      left: currentHeader === "Data1" ? "40px" : "auto",
                      backgroundColor: "#f4f4f4",
                      zIndex: currentHeader === "Data1" ? 2 : 1,
                      boxShadow:
                        currentHeader === "Data1"
                          ? "2px 0 5px rgba(0,0,0,0.1)"
                          : "none",
                      fontSize: "12px",
                    }}
                  >
                    {header}
                  </th>
                );
              })}
              <th
                style={{
                  padding: "8px 4px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "60px",
                  minWidth: "60px",
                  fontSize: "12px",
                  backgroundColor: "#e3f2fd",
                }}
              >
                % Hadir
              </th>
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "6px 4px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                    fontWeight: "bold",
                    color: "#666",
                    width: "40px",
                    minWidth: "40px",
                    position: "sticky",
                    left: 0,
                    backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                    zIndex: 1,
                    boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                    fontSize: "12px",
                  }}
                >
                  {rowIndex + 1}
                </td>
                {visibleHeaders.map((header, colIndex) => {
                  const isNama = header === "Data1";
                  const isEditable = editableHeaders.indexOf(header) !== -1;
                  return (
                    <td
                      key={colIndex}
                      style={{
                        padding: "4px",
                        borderBottom: "1px solid #eee",
                        position: header === "Data1" ? "sticky" : "static",
                        left: header === "Data1" ? "40px" : "auto",
                        width: header === "Data1" ? "150px" : "60px",
                        minWidth: header === "Data1" ? "150px" : "60px",
                        backgroundColor:
                          header === "Data1"
                            ? rowIndex % 2 === 0
                              ? "#fff"
                              : "#f9f9f9"
                            : "transparent",
                        zIndex: header === "Data1" ? 1 : 0,
                        boxShadow:
                          header === "Data1"
                            ? "2px 0 5px rgba(0,0,0,0.1)"
                            : "none",
                      }}
                    >
                      {readOnlyHeaders.has(header) || !isEditable ? (
                        <div
                          style={{
                            padding: "4px 2px",
                            color: "#666",
                            fontSize: "11px",
                            textAlign: isNama ? "left" : "center",
                            cursor: header === "Data1" ? "pointer" : "default",
                          }}
                          onClick={() => {
                            if (header === "Data1") {
                              setSelectedStudent({
                                nama: row.Data3 || "",
                                kelas: row.Data2 || "N/A",
                                nisn: row.Data1 || "N/A",
                              });
                              setShowStudentPopup(true);
                            }
                          }}
                        >
                          {row[header] || ""}
                        </div>
                      ) : (
                        <input
                          id={`kehadiran-input-${rowIndex}-${colIndex}`}
                          type="text"
                          inputMode="decimal"
                          pattern="[0-9]*"
                          value={row[header] || ""}
                          disabled={!isEditMode}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          onKeyDown={(e) =>
                            handleKeyDown(
                              e,
                              rowIndex,
                              colIndex,
                              actualData.length
                            )
                          }
                          onFocus={(e) => {
                            e.target.select();
                            updateFloatingButtonPosition(
                              e.target,
                              rowIndex,
                              colIndex,
                              true
                            );
                          }}
                          onBlur={(e) => {
                            const relatedTarget =
                              e.relatedTarget as HTMLElement;
                            if (
                              !relatedTarget ||
                              relatedTarget.tagName !== "BUTTON"
                            ) {
                              setTimeout(() => {
                                const activeElement = document.activeElement;
                                const isInputFocused =
                                  activeElement?.tagName === "INPUT" &&
                                  activeElement?.id.startsWith(
                                    "kehadiran-input-"
                                  );
                                if (!isInputFocused) {
                                  setShowFloatingButton(false);
                                }
                              }, 150);
                            }
                          }}
                          style={{
                            width: "100%",
                            padding: "2px",
                            border: isEditMode
                              ? "1px solid #ddd"
                              : "1px solid #eee",
                            borderRadius: "3px",
                            boxSizing: "border-box",
                            backgroundColor: isEditMode ? "white" : "#f5f5f5",
                            cursor: isEditMode ? "text" : "not-allowed",
                            fontSize: "11px",
                            textAlign: "center",
                            color: isEditMode ? "#000" : "#666",
                          }}
                        />
                      )}
                    </td>
                  );
                })}
                <td
                  style={{
                    padding: "4px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                    fontSize: "11px",
                    fontWeight: "bold",
                    color:
                      calcPersen(row) === "-"
                        ? "#999"
                        : parseFloat(calcPersen(row)) >= 75
                        ? "#2e7d32"
                        : "#e65100",
                    backgroundColor:
                      calcPersen(row) === "-"
                        ? "transparent"
                        : parseFloat(calcPersen(row)) >= 75
                        ? "#e8f5e9"
                        : "#fff3e0",
                    borderRadius: "3px",
                  }}
                >
                  {calcPersen(row)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Floating Arrow Button */}
      {showFloatingButton && (
        <button
          onClick={handleFloatingArrowClick}
          style={{
            position: "fixed",
            top: `${floatingButtonPosition.top}px`,
            left: `${floatingButtonPosition.left}px`,
            width: "56px",
            height: "56px",
            borderRadius: "50%",
            backgroundColor: "#4CAF50",
            color: "white",
            border: "none",
            cursor: "pointer",
            boxShadow: "0 4px 12px rgba(0,0,0,0.4)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: "28px",
            fontWeight: "bold",
            zIndex: 1001,
            transition: "all 0.2s ease",
            pointerEvents: floatingButtonPosition.visible ? "auto" : "none",
            opacity: floatingButtonPosition.visible ? 1 : 0,
            visibility: floatingButtonPosition.visible ? "visible" : "hidden",
          }}
        >
          ↓
        </button>
      )}

      {/* Modal Preview Import */}
      {importPreview && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0,0,0,0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setImportPreview(null)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "24px",
              maxWidth: "600px",
              width: "90%",
              maxHeight: "85vh",
              overflowY: "auto",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            {/* Header */}
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "16px",
                borderBottom: "2px solid #9C27B0",
                paddingBottom: "10px",
              }}
            >
              <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                👁️ Preview Import Kehadiran
              </h2>
              <button
                onClick={() => setImportPreview(null)}
                style={{
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  padding: "8px 16px",
                  cursor: "pointer",
                  fontWeight: "bold",
                }}
              >
                ✕
              </button>
            </div>

            {/* Ringkasan */}
            <div style={{ display: "flex", gap: "10px", marginBottom: "16px" }}>
              <div
                style={{
                  flex: 1,
                  padding: "10px",
                  backgroundColor: "#e8f5e9",
                  borderRadius: "6px",
                  border: "1px solid #4CAF50",
                  textAlign: "center",
                }}
              >
                <div
                  style={{
                    fontSize: "22px",
                    fontWeight: "bold",
                    color: "#2e7d32",
                  }}
                >
                  {importPreview.matched.length}
                </div>
                <div style={{ fontSize: "12px", color: "#555" }}>
                  Siswa cocok
                </div>
              </div>
              <div
                style={{
                  flex: 1,
                  padding: "10px",
                  backgroundColor: "#fff3e0",
                  borderRadius: "6px",
                  border: "1px solid #ff9800",
                  textAlign: "center",
                }}
              >
                <div
                  style={{
                    fontSize: "22px",
                    fontWeight: "bold",
                    color: "#e65100",
                  }}
                >
                  {importPreview.notFound.length}
                </div>
                <div style={{ fontSize: "12px", color: "#555" }}>
                  Tidak cocok
                </div>
              </div>
            </div>

            {/* Tabel preview data yang cocok */}
            {importPreview.matched.length > 0 && (
              <div style={{ marginBottom: "16px" }}>
                <p
                  style={{
                    margin: "0 0 8px",
                    fontWeight: "bold",
                    color: "#2e7d32",
                    fontSize: "14px",
                  }}
                >
                  ✅ Data yang akan diimport:
                </p>
                <div style={{ overflowX: "auto" }}>
                  <table
                    style={{
                      width: "100%",
                      borderCollapse: "collapse",
                      fontSize: "12px",
                    }}
                  >
                    <thead>
                      <tr style={{ backgroundColor: "#f3e5f5" }}>
                        <th
                          style={{
                            padding: "6px 8px",
                            border: "1px solid #ddd",
                            textAlign: "left",
                          }}
                        >
                          Nama
                        </th>
                        <th
                          style={{
                            padding: "6px 8px",
                            border: "1px solid #ddd",
                            textAlign: "center",
                          }}
                        >
                          Hadir
                        </th>
                        <th
                          style={{
                            padding: "6px 8px",
                            border: "1px solid #ddd",
                            textAlign: "center",
                          }}
                        >
                          Alpha
                        </th>
                        <th
                          style={{
                            padding: "6px 8px",
                            border: "1px solid #ddd",
                            textAlign: "center",
                          }}
                        >
                          Izin
                        </th>
                        <th
                          style={{
                            padding: "6px 8px",
                            border: "1px solid #ddd",
                            textAlign: "center",
                          }}
                        >
                          Sakit
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {importPreview.matched.map((item, i) => (
                        <tr
                          key={i}
                          style={{
                            backgroundColor: i % 2 === 0 ? "#fff" : "#f9f9f9",
                          }}
                        >
                          <td
                            style={{
                              padding: "5px 8px",
                              border: "1px solid #eee",
                            }}
                          >
                            {item.nama}
                          </td>
                          <td
                            style={{
                              padding: "5px 8px",
                              border: "1px solid #eee",
                              textAlign: "center",
                            }}
                          >
                            {item.hadir}
                          </td>
                          <td
                            style={{
                              padding: "5px 8px",
                              border: "1px solid #eee",
                              textAlign: "center",
                            }}
                          >
                            {item.alpha}
                          </td>
                          <td
                            style={{
                              padding: "5px 8px",
                              border: "1px solid #eee",
                              textAlign: "center",
                            }}
                          >
                            {item.izin}
                          </td>
                          <td
                            style={{
                              padding: "5px 8px",
                              border: "1px solid #eee",
                              textAlign: "center",
                            }}
                          >
                            {item.sakit}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Daftar nama tidak cocok */}
            {importPreview.notFound.length > 0 && (
              <div
                style={{
                  marginBottom: "16px",
                  padding: "12px",
                  backgroundColor: "#fff3e0",
                  borderRadius: "6px",
                  border: "1px solid #ff9800",
                }}
              >
                <p
                  style={{
                    margin: "0 0 6px",
                    fontSize: "13px",
                    color: "#e65100",
                    fontWeight: "bold",
                  }}
                >
                  ⚠️ Nama berikut tidak ditemukan di Excel (tidak akan
                  diimport):
                </p>
                <p
                  style={{ margin: "0 0 6px", fontSize: "11px", color: "#777" }}
                >
                  Nama harus sama persis termasuk huruf kapital dan spasi.
                </p>
                <ul
                  style={{
                    margin: 0,
                    paddingLeft: "18px",
                    fontSize: "12px",
                    color: "#555",
                  }}
                >
                  {importPreview.notFound.map((nama, i) => (
                    <li key={i}>{nama}</li>
                  ))}
                </ul>
              </div>
            )}

            {/* Tombol Batal & Kirim */}
            <div
              style={{
                display: "flex",
                gap: "10px",
                justifyContent: "flex-end",
              }}
            >
              <button
                onClick={() => setImportPreview(null)}
                style={{
                  padding: "10px 24px",
                  backgroundColor: "#757575",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                  fontWeight: "bold",
                  fontSize: "14px",
                }}
              >
                ❌ Batal
              </button>
              <button
                onClick={handleConfirmImport}
                disabled={importPreview.matched.length === 0 || isSaving}
                style={{
                  padding: "10px 24px",
                  backgroundColor:
                    importPreview.matched.length === 0 || isSaving
                      ? "#ccc"
                      : "#4CAF50",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor:
                    importPreview.matched.length === 0 || isSaving
                      ? "not-allowed"
                      : "pointer",
                  fontWeight: "bold",
                  fontSize: "14px",
                }}
              >
                {isSaving
                  ? "⏳ Menyimpan..."
                  : `✅ Kirim Data (${importPreview.matched.length} siswa)`}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Hasil Import */}
      {importResult && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0,0,0,0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setImportResult(null)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "24px",
              maxWidth: "480px",
              width: "90%",
              maxHeight: "80vh",
              overflowY: "auto",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "16px",
                borderBottom: "2px solid #4CAF50",
                paddingBottom: "10px",
              }}
            >
              <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                ✅ Hasil Import Kehadiran
              </h2>
              <button
                onClick={() => setImportResult(null)}
                style={{
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  padding: "8px 16px",
                  cursor: "pointer",
                  fontWeight: "bold",
                }}
              >
                Tutup
              </button>
            </div>
            <div
              style={{
                marginBottom: "16px",
                padding: "12px",
                backgroundColor: "#e8f5e9",
                borderRadius: "6px",
                border: "1px solid #4CAF50",
              }}
            >
              <p style={{ margin: 0, fontSize: "15px", color: "#2e7d32" }}>
                ✅ <strong>{importResult.matched} siswa</strong> berhasil
                diimport
              </p>
              <p style={{ margin: "6px 0 0", fontSize: "12px", color: "#555" }}>
                Klik "Simpan Perubahan" untuk menyimpan ke server.
              </p>
            </div>
            {importResult.notFound.length > 0 && (
              <div
                style={{
                  padding: "12px",
                  backgroundColor: "#fff3e0",
                  borderRadius: "6px",
                  border: "1px solid #ff9800",
                }}
              >
                <p
                  style={{
                    margin: "0 0 8px",
                    fontSize: "14px",
                    color: "#e65100",
                    fontWeight: "bold",
                  }}
                >
                  ⚠️ {importResult.notFound.length} nama siswa tidak cocok
                  dengan data Excel:
                </p>
                <p
                  style={{ margin: "0 0 8px", fontSize: "12px", color: "#777" }}
                >
                  Nama harus sama persis termasuk huruf kapital dan spasi.
                </p>
                <ul
                  style={{
                    margin: 0,
                    paddingLeft: "20px",
                    fontSize: "13px",
                    color: "#555",
                  }}
                >
                  {importResult.notFound.map((nama, i) => (
                    <li key={i}>{nama}</li>
                  ))}
                </ul>
              </div>
            )}
          </div>
        </div>
      )}

      {showStudentPopup && selectedStudent && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setShowStudentPopup(false)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "20px",
              maxWidth: "400px",
              width: "90%",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "15px",
                borderBottom: "2px solid #2196F3",
                paddingBottom: "10px",
              }}
            >
              <h2 style={{ margin: 0, color: "#333", fontSize: "18px" }}>
                Detail Siswa: {selectedStudent.nama}
              </h2>
              <button
                onClick={() => setShowStudentPopup(false)}
                style={{
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  padding: "8px 16px",
                  cursor: "pointer",
                  fontSize: "14px",
                  fontWeight: "bold",
                }}
              >
                Tutup
              </button>
            </div>
            <div style={{ marginBottom: "15px" }}>
              <strong style={{ color: "#4CAF50" }}>Nama:</strong>{" "}
              <span style={{ color: "#333" }}>{selectedStudent.nama}</span>
            </div>
            <div style={{ marginBottom: "15px" }}>
              <strong style={{ color: "#4CAF50" }}>Kelas:</strong>{" "}
              <span style={{ color: "#333" }}>{selectedStudent.kelas}</span>
            </div>
            <div style={{ marginBottom: "15px" }}>
              <strong style={{ color: "#4CAF50" }}>NISN:</strong>{" "}
              <span style={{ color: "#333" }}>{selectedStudent.nisn}</span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

const InputTP = () => {
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editTP, setEditTP] = useState({
    mapel: "",
    tp: "",
    rincian: "",
    bab: "",
    semester: "",
    kelas: "",
  });
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isAddingNew, setIsAddingNew] = useState(false);
  const [newTP, setNewTP] = useState({
    mapel: "",
    tp: "",
    rincian: "",
    bab: "",
    semester: "",
    kelas: "",
  });
  const [availableMapel, setAvailableMapel] = useState<string[]>([]);
  const [availableSemester, setAvailableSemester] = useState<string[]>([]);
  const [availableKelas, setAvailableKelas] = useState<string[]>([]);
  const [filterMapel, setFilterMapel] = useState<string>("");
  const [filterSemester, setFilterSemester] = useState<string>("");
  const [filterKelas, setFilterKelas] = useState<string>("");

  // ✅ TAMBAH: Row mapping untuk track index asli
  const [rowMapping, setRowMapping] = useState<number[]>([]);
  const targetKelasRef = useRef<string>("");

  const processTPData = (
    tpRaw: any[],
    mapelRaw: any[],
    targetKelas: string
  ) => {
    if (!tpRaw || tpRaw.length < 2) return;

    const headers = tpRaw[0];
    const allRows = tpRaw.slice(1);
    const filteredWithMapping: { row: any; originalIndex: number }[] = [];

    allRows.forEach((row: any, index: number) => {
      const hasData =
        (row.Data1 && String(row.Data1).trim() !== "") ||
        (row.Data3 && String(row.Data3).trim() !== "");
      const rowKelas = String(row.Data6 || "").trim();
      const kelasMatch = !targetKelas || rowKelas === targetKelas;
      if (hasData && kelasMatch) {
        filteredWithMapping.push({ row, originalIndex: index + 2 });
      }
    });

    const filteredData = filteredWithMapping.map((i) => i.row);
    const mapping = filteredWithMapping.map((i) => i.originalIndex);

    setData([headers, ...filteredData]);
    setRowMapping(mapping);

    const semSet = new Set<string>();
    filteredData.forEach((row: any) => {
      if (row.Data5) semSet.add(String(row.Data5));
    });
    setAvailableSemester(Array.from(semSet).sort());

    if (mapelRaw && mapelRaw.length > 1) {
      const mapelList = mapelRaw
        .slice(1)
        .filter((r: any) => r.Data1?.trim())
        .map((r: any) =>
          r.Data1.replace(/[\u200B\u200C\u200D\uFEFF]/g, "").trim()
        )
        .filter((m: string) => m !== "");
      setAvailableMapel([...new Set<string>(mapelList)].sort());
    }
  };

  const [idbReady, setIdbReady] = useState(false);
  const syncQueueRef = useRef<(() => Promise<void>)[]>([]);

  const {
    tpData,
    mapelListData,
    loading: contextLoading,
    refreshRekapData,
    updateLocalData,
    refreshMapelSheet,
  } = useRekapData();

  // ✅ TAMBAH BARU: Baca localStorage cache langsung saat mount (sebelum context siap)
  useEffect(() => {
    try {
      const cachedTP = localStorage.getItem("cache_tpData");
      const cachedMapel = localStorage.getItem("cache_mapelData");

      if (!cachedTP || !cachedMapel) return;

      const tpJson = JSON.parse(cachedTP);
      const mapelJson = JSON.parse(cachedMapel);

      if (!tpJson || tpJson.length < 2 || !mapelJson || mapelJson.length < 1)
        return;

      const headers = tpJson[0];
      const allRows = tpJson.slice(1);

      const filteredDataWithMapping: { row: any; originalIndex: number }[] = [];
      allRows.forEach((row: any, index: number) => {
        const hasData =
          (row.Data1 && String(row.Data1).trim() !== "") ||
          (row.Data3 && String(row.Data3).trim() !== "");
        if (hasData) {
          filteredDataWithMapping.push({
            row,
            originalIndex: index + 3,
          });
        }
      });

      const filteredData = filteredDataWithMapping.map((item) => item.row);
      const mapping = filteredDataWithMapping.map((item) => item.originalIndex);

      setData([headers, ...filteredData]);
      setRowMapping(mapping);

      const semesterSet = new Set<string>();
      filteredData.forEach((row: any) => {
        if (row.Data5) semesterSet.add(String(row.Data5));
      });
      setAvailableSemester(Array.from(semesterSet).sort());

      const mapelRows = mapelJson
        .slice(1)
        .filter(
          (row: any) =>
            row.Data1 &&
            typeof row.Data1 === "string" &&
            row.Data1.trim() !== ""
        );
      const mapelList = mapelRows
        .map((row: any) =>
          row.Data1.replace(/[\u200B\u200C\u200D\uFEFF]/g, "").trim()
        )
        .filter((m: string) => m !== "");
      setAvailableMapel(Array.from(new Set<string>(mapelList)).sort());

      setLoading(false); // ← tampil langsung dari cache, tanpa tunggu context/fetch
      console.log("✅ InputTP: data dimuat dari localStorage cache");
    } catch (e) {
      console.warn("Gagal baca cache InputTP:", e);
    }
  }, []); // ← [] = hanya run sekali saat mount

  // ✅ BARU: Load dari IndexedDB saat mount (instant, tanpa server)
  useEffect(() => {
    const loadFromIDB = async () => {
      try {
        const [tpCached, mapelCached] = await Promise.all([
          idbLoad(STORE_TP),
          idbLoad(STORE_MAPEL),
        ]);

        if (!tpCached || tpCached.length < 2) {
          console.log("IndexedDB kosong, tunggu context/server");
          return;
        }

        const headers = tpCached[0];
        const allRows = tpCached.slice(1);
        const filteredWithMapping: { row: any; originalIndex: number }[] = [];

        allRows.forEach((row: any, index: number) => {
          const hasData =
            (row.Data1 && String(row.Data1).trim() !== "") ||
            (row.Data3 && String(row.Data3).trim() !== "");
          if (hasData) {
            filteredWithMapping.push({ row, originalIndex: index + 3 });
          }
        });

        const filteredData = filteredWithMapping.map((i) => i.row);
        const mapping = filteredWithMapping.map((i) => i.originalIndex);

        setData([headers, ...filteredData]);
        setRowMapping(mapping);

        const semSet = new Set<string>();
        filteredData.forEach((row: any) => {
          if (row.Data5) semSet.add(String(row.Data5));
        });
        setAvailableSemester(Array.from(semSet).sort());

        if (mapelCached && mapelCached.length > 1) {
          const mapelList = mapelCached
            .slice(1)
            .filter((r: any) => r.Data1?.trim())
            .map((r: any) =>
              r.Data1.replace(/[\u200B\u200C\u200D\uFEFF]/g, "").trim()
            )
            .filter((m: string) => m !== "");
          setAvailableMapel([...new Set<string>(mapelList)].sort());
        }

        setIdbReady(true);
        setLoading(false);
        console.log("✅ InputTP: data dimuat dari IndexedDB (instant)");
      } catch (e) {
        console.warn("Gagal load IndexedDB:", e);
      }
    };

    loadFromIDB();
  }, []);

  useEffect(() => {
    const init = async () => {
      // Coba load dari IndexedDB dulu
      const [tpCached, mapelCached, kelasCached] = await Promise.all([
        idbLoad(STORE_TP),
        idbLoad(STORE_MAPEL),
        idbLoad("kelasData"),
      ]);

      if (tpCached && tpCached.length > 1) {
        // ✅ Prioritas baca kelas dari sekolahData, bukan kelasData
        const sekolahCached = await idbLoad("sekolahData");
        const savedKelas = sekolahCached?.kelas
          ? String(sekolahCached.kelas).trim()
          : kelasCached?.[0]?.kelas || "";

        targetKelasRef.current = savedKelas;

        unstable_batchedUpdates(() => {
          setAvailableKelas(savedKelas ? [savedKelas] : []);
          setFilterKelas(savedKelas);
          processTPData(tpCached, mapelCached || [], savedKelas);
          setLoading(false);
        });
        console.log("✅ InputTP: dari IndexedDB");

        // Background sync dari Google Sheet → update IndexedDB jika ada perubahan
        fetchAndSyncFromServer(false);
      } else {
        // IndexedDB kosong → fetch server, simpan ke IndexedDB
        fetchAndSyncFromServer(true);
      }
    };

    const fetchAndSyncFromServer = async (showLoading: boolean) => {
      if (showLoading) setLoading(true);
      try {
        let targetKelas = targetKelasRef.current;

        if (!targetKelas) {
          // ✅ Cek IndexedDB sekolahData dulu
          const sekolahCached = await idbLoad("sekolahData");
          if (sekolahCached?.kelas) {
            targetKelas = String(sekolahCached.kelas).trim();
          } else {
            // Fallback fetch server
            const schoolRes = await fetch(`${endpoint}?action=schoolData`);
            if (schoolRes.ok) {
              const schoolJson = await schoolRes.json();
              if (schoolJson.success && schoolJson.data?.length > 0) {
                targetKelas = String(schoolJson.data[0].kelas || "").trim();
                await idbSave("sekolahData", schoolJson.data[0]);
              }
            }
          }
          targetKelasRef.current = targetKelas;
          setAvailableKelas(targetKelas ? [targetKelas] : []);
          setFilterKelas(targetKelas);
          await idbSave("kelasData", [{ kelas: targetKelas }]);
        }

        const [tpRes, mapelRes] = await Promise.all([
          fetch(`${endpoint}?sheet=DataTP`),
          fetch(`${endpoint}?sheet=DataMapel`),
        ]);

        if (!tpRes.ok || !mapelRes.ok) throw new Error("Gagal fetch server");

        const [tpJson, mapelJson] = await Promise.all([
          tpRes.json(),
          mapelRes.json(),
        ]);

        // Simpan ke IndexedDB
        await idbSave(STORE_TP, tpJson);
        await idbSave(STORE_MAPEL, mapelJson);
        await idbSave("kelasData", [{ kelas: targetKelas }]);
        console.log("✅ IndexedDB diperbarui dari Google Sheet");

        // Update tampilan
        processTPData(tpJson, mapelJson, targetKelas);
      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
      } finally {
        if (showLoading) setLoading(false);
      }
    };

    init();
  }, []); // ← tidak bergantung context sama sekali

  // Function untuk auto-generate TP berdasarkan MAPEL dan BAB
  const getNextTP = (mapel: string, bab: string): string => {
    if (!mapel || !bab) return "";

    const actualData = data.slice(1);

    const filteredData = actualData.filter((row: any) => row.Data1 === mapel);

    if (filteredData.length === 0) {
      return `${bab}.1`;
    }

    // ✅ UBAH: Bersihkan \u200B saat membaca data lama
    const tpList = filteredData
      .map((row: any) => {
        const tp = row.Data2 || "";
        return tp.replace(/[\u200B\u200C\u200D\uFEFF]/g, "").trim();
      })
      .filter((tp: string) => {
        if (!tp || typeof tp !== "string") return false;
        return tp.split(".")[0] === bab;
      })
      .sort((a: string, b: string) => {
        const [, aSub] = a.split(".").map(Number);
        const [, bSub] = b.split(".").map(Number);
        return aSub - bSub;
      });

    if (tpList.length === 0) {
      return `${bab}.1`;
    }

    const lastTP = tpList[tpList.length - 1];
    const [mainNum, subNum] = lastTP.split(".").map(Number);

    const nextTP = `${mainNum}.${subNum + 1}`;
    return nextTP;
  };

  const getTPCount = (
    mapel: string,
    semester: string,
    kelas: string
  ): number => {
    if (!mapel || !semester || !kelas) return 0;
    return data
      .slice(1)
      .filter(
        (row: any) =>
          row.Data1 === mapel &&
          String(row.Data5) === String(semester) &&
          String(row.Data6) === String(kelas)
      ).length;
  };

  const handleMapelChange = async (selectedMapel: string) => {
    // Coba ambil semester & kelas dari DataTP dulu
    const actualDataTP = data.slice(1);
    const mapelDataTP = actualDataTP.find(
      (row: any) => row.Data1 === selectedMapel
    );

    if (mapelDataTP) {
      // Sudah ada di DataTP, langsung pakai
      setNewTP((prev) => ({
        ...prev,
        mapel: selectedMapel,
        tp: "",
        bab: "",
        semester: String(mapelDataTP.Data5 || ""),
        kelas: String(mapelDataTP.Data6 || ""),
      }));
    } else {
      // Belum ada di DataTP, fetch dari DataMapel
      try {
        const mapelResponse = await fetch(`${endpoint}?sheet=DataMapel`);
        if (mapelResponse.ok) {
          const mapelJson = await mapelResponse.json();
          const mapelRow = mapelJson
            .slice(1)
            .find(
              (row: any) =>
                (row.Data1 || "")
                  .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
                  .trim() === selectedMapel
            );

          setNewTP((prev) => ({
            ...prev,
            mapel: selectedMapel,
            tp: "",
            bab: "",
            semester: mapelRow ? String(mapelRow.Data2 || "") : "", // Data2 = semester di DataMapel
            kelas: mapelRow ? String(mapelRow.Data3 || "") : "", // Data3 = kelas di DataMapel
          }));
        }
      } catch (err) {
        console.error("Error fetching DataMapel untuk semester/kelas:", err);
        setNewTP((prev) => ({
          ...prev,
          mapel: selectedMapel,
          tp: "",
          bab: "",
          semester: "",
          kelas: "",
        }));
      }
    }
  };

  const handleBabChange = (selectedBab: string) => {
    setNewTP((prev) => ({
      ...prev,
      bab: selectedBab,
      tp: prev.mapel && selectedBab ? getNextTP(prev.mapel, selectedBab) : "",
    }));
  };

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);

    const headers = ["Data1", "Data2", "Data3", "Data4", "Data5", "Data6"];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");

      // ✅ GUNAKAN ROW MAPPING
      const originalRowIndex = rowMapping[rowIndex];

      console.log(
        `📝 Saving row ${rowIndex} (UI) -> Row ${originalRowIndex} (Sheet):`,
        values
      );

      updates.push({
        rowIndex: originalRowIndex + 1, // +1 karena sheet row dimulai dari 1
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: "DataTP",
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());

      setIsSaving(false);
      refreshRekapData(true);
    } catch (err) {
      console.error("=== ERROR DETAILS ===");
      console.error(err);
      alert(
        "Error updating rows: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  const handleAddNew = async () => {
    if (
      !newTP.mapel ||
      !newTP.tp ||
      !newTP.rincian ||
      !newTP.bab ||
      !newTP.semester ||
      !newTP.kelas
    ) {
      alert("⚠️ Semua field wajib diisi!");
      return;
    }

    // Validasi batas 10 TP
    const currentCount = getTPCount(newTP.mapel, newTP.semester, newTP.kelas);
    if (currentCount >= 10) {
      alert(
        `⛔ Batas maksimal tercapai!\n\n` +
          `${newTP.mapel} — Semester ${newTP.semester} — Kelas ${newTP.kelas}\n` +
          `sudah memiliki ${currentCount} TP.\n\n` +
          `Maksimal 10 TP per mata pelajaran per semester.`
      );
      return;
    }

    setIsSaving(true);

    const cleanTP = {
      mapel: newTP.mapel.replace(/[\u200B\u200C\u200D\uFEFF]/g, "").trim(),
      tp: newTP.tp.replace(/[\u200B\u200C\u200D\uFEFF]/g, "").trim(),
      rincian: newTP.rincian.replace(/[\u200B\u200C\u200D\uFEFF]/g, "").trim(),
      bab: newTP.bab.replace(/[\u200B\u200C\u200D\uFEFF]/g, "").trim(),
      semester: newTP.semester
        .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
        .trim(),
      kelas: newTP.kelas.replace(/[\u200B\u200C\u200D\uFEFF]/g, "").trim(),
    };

    try {
      const requestBody = {
        action: "add_tp",
        data: cleanTP,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("✅ Data TP baru berhasil ditambahkan!");

      const newRow = {
        Data1: cleanTP.mapel, // ⬅️ ganti newTP menjadi cleanTP
        Data2: cleanTP.tp,
        Data3: cleanTP.rincian,
        Data4: cleanTP.bab,
        Data5: cleanTP.semester,
        Data6: cleanTP.kelas,
      };
      const updatedData = [...data, newRow];
      setData(updatedData);
      updateLocalData("tp", updatedData);

      setData(updatedData);
      updateLocalData("tp", updatedData);
      // ✅ BARU: Sync ke IndexedDB
      await idbSave(STORE_TP, updatedData);
      console.log("✅ IndexedDB diperbarui setelah tambah TP");

      setNewTP({
        mapel: "",
        tp: "",
        rincian: "",
        bab: "",
        semester: "",
        kelas: "",
      });
      setIsAddingNew(false);
      setIsSaving(false);

      await refreshMapelSheet(cleanTP.mapel);

      setTimeout(() => {
        refreshRekapData(true);
      }, 2000);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menambah data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  const handleEdit = (rowIndex: number) => {
    const rowData = actualData[rowIndex];
    setEditingIndex(rowIndex);
    setEditTP({
      mapel: rowData.Data1 || "",
      tp: rowData.Data2 || "",
      rincian: rowData.Data3 || "",
      bab: rowData.Data4 || "",
      semester: rowData.Data5 || "",
      kelas: rowData.Data6 || "",
    });
    setIsAddingNew(false); // tutup form tambah jika terbuka
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const handleSaveEdit = async () => {
    if (
      !editTP.mapel ||
      !editTP.tp ||
      !editTP.rincian ||
      !editTP.bab ||
      !editTP.semester ||
      !editTP.kelas
    ) {
      alert("⚠️ Semua field wajib diisi!");
      return;
    }

    if (editingIndex === null) return;

    setIsSaving(true);

    const headers = ["Data1", "Data2", "Data3", "Data4", "Data5", "Data6"];
    const values = [
      String(editTP.mapel)
        .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
        .trim(),
      String(editTP.tp)
        .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
        .trim(),
      String(editTP.rincian)
        .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
        .trim(),
      String(editTP.bab)
        .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
        .trim(),
      String(editTP.semester)
        .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
        .trim(),
      String(editTP.kelas)
        .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
        .trim(),
    ];

    const originalRowIndex = rowMapping[editingIndex];

    try {
      const requestBody = {
        action: "update_tp_bulk",
        updates: [
          {
            rowIndex: originalRowIndex + 1,
            values: values,
          },
        ],
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      alert("✅ Data TP berhasil diperbarui!");

      // Update local state
      const updatedData = [...data];
      updatedData[editingIndex + 1] = {
        Data1: values[0], // ⬅️ gunakan values[] yang sudah clean
        Data2: values[1],
        Data3: values[2],
        Data4: values[3],
        Data5: values[4],
        Data6: values[5],
      };
      setData(updatedData);
      updateLocalData("tp", updatedData);
      await idbSave(STORE_TP, updatedData);
      console.log("✅ IndexedDB diperbarui setelah edit TP");

      setEditingIndex(null);
      setEditTP({
        mapel: "",
        tp: "",
        rincian: "",
        bab: "",
        semester: "",
        kelas: "",
      });
      setIsSaving(false);

      refreshRekapData(true);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error memperbarui data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  // ✅ UBAH FUNGSI DELETE - Gunakan row mapping
  const handleDelete = async (rowIndex: number) => {
    const rowData = actualData[rowIndex];
    const mapelName = rowData.Data1;

    const confirmation = window.confirm(
      `Apakah Anda yakin ingin menghapus data TP ini?\n\nMapel: ${rowData.Data1}\nTP: ${rowData.Data2}\nRincian: ${rowData.Data3}\n\n⚠️ Data akan dikosongkan (baris tetap ada)`
    );

    if (!confirmation) return;

    setIsSaving(true);

    try {
      // ✅ GUNAKAN ROW MAPPING untuk mendapatkan row index asli
      const originalRowIndex = rowMapping[rowIndex];

      console.log(
        `🗑️ Deleting row ${rowIndex} (UI) -> Row ${originalRowIndex} (Sheet)`
      );

      const requestBody = {
        action: "delete_row",
        sheetName: "DataTP",
        rowIndex: originalRowIndex, // ← HAPUS +1
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result = await response.json();

      if (!result.success) {
        throw new Error(result.error || "Delete failed");
      }

      alert("✅ Data TP berhasil dihapus!");

      // Update local state
      const updatedData = [...data];
      const emptyRow: RowData = {};
      headers.forEach((header) => {
        emptyRow[header] = "";
      });
      updatedData[rowIndex + 1] = emptyRow;
      setData(updatedData);
      updateLocalData("tp", updatedData);
      await idbSave(STORE_TP, updatedData);
      console.log("✅ IndexedDB diperbarui setelah hapus TP");

      setIsSaving(false);

      if (mapelName) {
        console.log(`🔄 Refreshing MAPEL sheets for: ${mapelName}`);
        await refreshMapelSheet(mapelName);
      }

      setTimeout(() => {
        refreshRekapData(true);
      }, 2000);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menghapus data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  // ← HAPUS kondisi "No data available"

  const headers = ["Data1", "Data2", "Data3", "Data4", "Data5", "Data6"];
  const displayHeaders = headers.map((header) => data[0][header] || "");

  // ✅ Data sudah terfilter saat load, jadi langsung pakai
  const actualData = data.slice(1);

  const filteredData = actualData.filter((row: any) => {
    const matchMapel = filterMapel === "" || row.Data1 === filterMapel;
    const matchSemester =
      filterSemester === "" || String(row.Data5) === filterSemester;
    const matchKelas = filterKelas === "" || String(row.Data6) === filterKelas; // ← TAMBAH
    return matchMapel && matchSemester && matchKelas;
  });

  // Cek limit TP untuk form TAMBAH
  const newTPCount = getTPCount(newTP.mapel, newTP.semester, newTP.kelas);
  const isNewTPLimitReached = !!(
    newTP.mapel &&
    newTP.semester &&
    newTP.kelas &&
    newTPCount >= 10
  );

  // Cek limit TP untuk form EDIT (exclude row yang sedang diedit)
  const editTPCountExcludeSelf =
    editingIndex !== null
      ? data
          .slice(1)
          .filter(
            (row: any, idx: number) =>
              idx !== editingIndex &&
              row.Data1 === editTP.mapel &&
              String(row.Data5) === String(editTP.semester) &&
              String(row.Data6) === String(editTP.kelas)
          ).length
      : 0;
  const isEditTPLimitReached = !!(
    editingIndex !== null &&
    editTP.mapel &&
    editTP.semester &&
    editTP.kelas &&
    editTPCountExcludeSelf >= 10
  );

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        📚 Data Tujuan Pembelajaran (TP)
      </h1>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={() => setIsAddingNew(!isAddingNew)}
          style={{
            padding: "12px 24px",
            backgroundColor: isAddingNew ? "#f44336" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            marginRight: "10px",
          }}
        >
          {isAddingNew ? "❌ Batal Tambah" : "➕ Tambah Data Baru"}
        </button>

        <button
          onClick={handleSaveAll}
          disabled={isSaving || changedRows.size === 0}
          style={{
            padding: "12px 24px",
            backgroundColor:
              isSaving || changedRows.size === 0 ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor:
              isSaving || changedRows.size === 0 ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
          }}
        >
          {isSaving
            ? "Memproses..."
            : `💾 Simpan Perubahan (${changedRows.size})`}
        </button>
      </div>

      {/* Form Tambah TP - tetap sama */}
      {isAddingNew && (
        <div
          style={{
            backgroundColor: "#f0f8ff",
            padding: "20px",
            borderRadius: "8px",
            marginBottom: "20px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          }}
        >
          <h3 style={{ marginBottom: "15px", color: "#2196F3" }}>
            Form Tambah TP Baru
          </h3>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "15px",
            }}
          >
            {/* LANGKAH 1: MAPEL */}
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                1️⃣ Mapel:
              </label>
              <select
                value={newTP.mapel}
                onChange={(e) => {
                  const value = e.target.value;
                  if (value === "__custom__") {
                    const customMapel = prompt("Masukkan nama Mapel baru:");
                    if (customMapel) {
                      setNewTP({
                        mapel: customMapel.toUpperCase(),
                        tp: "",
                        rincian: "",
                        bab: "",
                        semester: "",
                        kelas: "",
                      });
                    }
                  } else {
                    handleMapelChange(value);
                  }
                }}
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                }}
              >
                <option value="">-- Pilih Mapel --</option>
                {availableMapel.map((mapel, index) => (
                  <option key={index} value={mapel}>
                    {mapel}
                  </option>
                ))}
                <option value="__custom__">➕ Tambah Mapel Baru...</option>
              </select>
            </div>

            {/* LANGKAH 2: BAB */}
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                2️⃣ BAB:
              </label>
              <input
                type="text"
                value={newTP.bab}
                onChange={(e) => handleBabChange(e.target.value)}
                placeholder="Contoh: 1, 2, 3, dst"
                disabled={!newTP.mapel}
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  backgroundColor: !newTP.mapel ? "#f5f5f5" : "white",
                  cursor: !newTP.mapel ? "not-allowed" : "text",
                }}
              />
              {!newTP.mapel && (
                <small style={{ color: "#999", fontSize: "11px" }}>
                  Pilih Mapel terlebih dahulu
                </small>
              )}
            </div>

            {/* LANGKAH 3: TP (AUTO) */}
            <div style={{ gridColumn: "1 / -1" }}>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                3️⃣ TP:
                {newTP.mapel && newTP.bab && (
                  <span
                    style={{
                      color: "#4CAF50",
                      fontSize: "12px",
                      marginLeft: "8px",
                    }}
                  >
                    ✓ Otomatis terisi berdasarkan BAB
                  </span>
                )}
              </label>
              <input
                type="text"
                value={newTP.tp}
                readOnly
                placeholder="TP akan terisi otomatis setelah Mapel & BAB diisi"
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "2px solid " + (newTP.tp ? "#4CAF50" : "#ddd"),
                  borderRadius: "4px",
                  backgroundColor: newTP.tp ? "#e8f5e9" : "#f5f5f5",
                  fontSize: "16px",
                  fontWeight: "bold",
                  color: newTP.tp ? "#2e7d32" : "#999",
                  textAlign: "center",
                  cursor: "not-allowed",
                }}
              />
            </div>

            {/* RINCIAN TP */}
            <div style={{ gridColumn: "1 / -1" }}>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                Rincian TP:
              </label>
              <textarea
                value={newTP.rincian}
                onChange={(e) =>
                  setNewTP({ ...newTP, rincian: e.target.value })
                }
                placeholder="Jelaskan rincian tujuan pembelajaran..."
                rows={3}
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  resize: "vertical",
                }}
              />
            </div>

            {/* SEMESTER & KELAS */}
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                Semester:
                {newTP.mapel && (
                  <span
                    style={{
                      color: "#2196F3",
                      fontSize: "12px",
                      marginLeft: "5px",
                    }}
                  >
                    (Auto-fill)
                  </span>
                )}
              </label>
              <input
                type="text"
                value={newTP.semester}
                onChange={(e) =>
                  setNewTP({ ...newTP, semester: e.target.value })
                }
                placeholder="Contoh: 1"
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  backgroundColor: newTP.mapel ? "#f0f8ff" : "white",
                }}
              />
            </div>
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                Kelas:
                {newTP.mapel && (
                  <span
                    style={{
                      color: "#2196F3",
                      fontSize: "12px",
                      marginLeft: "5px",
                    }}
                  >
                    (Auto-fill)
                  </span>
                )}
              </label>
              <input
                type="text"
                value={newTP.kelas}
                onChange={(e) => setNewTP({ ...newTP, kelas: e.target.value })}
                placeholder="Contoh: 6"
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  backgroundColor: newTP.mapel ? "#f0f8ff" : "white",
                }}
              />
            </div>
          </div>

          {newTP.mapel && newTP.bab && newTP.tp && (
            <div
              style={{
                marginTop: "15px",
                padding: "12px",
                backgroundColor: "#e8f5e9",
                border: "1px solid #4CAF50",
                borderRadius: "4px",
              }}
            >
              <strong style={{ color: "#2e7d32" }}>✓ Siap disimpan:</strong>
              <div
                style={{ marginTop: "5px", fontSize: "14px", color: "#333" }}
              >
                {newTP.mapel} - BAB {newTP.bab} - TP {newTP.tp}
              </div>
            </div>
          )}

          {isNewTPLimitReached && (
            <div
              style={{
                marginTop: "12px",
                padding: "12px 16px",
                backgroundColor: "#fdecea",
                border: "1px solid #f44336",
                borderRadius: "6px",
                color: "#b71c1c",
                fontSize: "14px",
              }}
            >
              ⛔ <strong>Batas maksimal tercapai!</strong> {newTP.mapel} —
              Semester {newTP.semester} — Kelas {newTP.kelas} sudah memiliki{" "}
              <strong>{newTPCount} TP</strong>. Maksimal 10 TP per mata
              pelajaran per semester.
            </div>
          )}

          <button
            onClick={handleAddNew}
            disabled={isSaving || isNewTPLimitReached}
            style={{
              marginTop: "15px",
              padding: "12px 24px",
              backgroundColor:
                isSaving || isNewTPLimitReached ? "#ccc" : "#4CAF50",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor:
                isSaving || isNewTPLimitReached ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "16px",
            }}
          >
            {isSaving ? "Menyimpan..." : "💾 Simpan Data Baru"}
          </button>
        </div>
      )}

      {/* Form Edit TP */}
      {editingIndex !== null && (
        <div
          style={{
            backgroundColor: "#fff8e1",
            padding: "20px",
            borderRadius: "8px",
            marginBottom: "20px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
            border: "2px solid #FFC107",
          }}
        >
          <h3 style={{ marginBottom: "15px", color: "#F57F17" }}>
            ✏️ Edit Data TP — Row {editingIndex + 1}
          </h3>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "15px",
            }}
          >
            {/* MAPEL - dropdown */}
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                1️⃣ Mapel:
              </label>
              <select
                value={editTP.mapel}
                onChange={(e) => {
                  const value = e.target.value;
                  if (value === "__custom__") {
                    const customMapel = prompt("Masukkan nama Mapel baru:");
                    if (customMapel) {
                      setEditTP((prev) => ({
                        ...prev,
                        mapel: customMapel.toUpperCase(),
                      }));
                    }
                  } else {
                    // Auto-fill semester & kelas dari data yang ada
                    const actualData = data.slice(1);
                    const mapelData = actualData.find(
                      (row: any) => row.Data1 === value
                    );
                    setEditTP((prev) => ({
                      ...prev,
                      mapel: value,
                      semester: mapelData?.Data5 || prev.semester,
                      kelas: mapelData?.Data6 || prev.kelas,
                    }));
                  }
                }}
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                }}
              >
                <option value="">-- Pilih Mapel --</option>
                {availableMapel.map((mapel, index) => (
                  <option key={index} value={mapel}>
                    {mapel}
                  </option>
                ))}
                <option value="__custom__">➕ Tambah Mapel Baru...</option>
              </select>
            </div>

            {/* BAB - dengan auto-generate TP */}
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                2️⃣ BAB:
              </label>
              <input
                type="text"
                value={editTP.bab}
                onChange={(e) => {
                  const selectedBab = e.target.value;
                  // Hitung next TP berdasarkan mapel & bab,
                  // tapi EXCLUDE row yang sedang diedit agar tidak menghitung dirinya sendiri
                  const getNextTPForEdit = (
                    mapel: string,
                    bab: string
                  ): string => {
                    if (!mapel || !bab) return "";
                    const actualRows = data.slice(1);
                    const filteredData = actualRows.filter(
                      (row: any, idx: number) =>
                        row.Data1 === mapel && idx !== editingIndex
                    );
                    const tpList = filteredData
                      .map((row: any) => row.Data2)
                      .filter((tp: string) => {
                        if (!tp || typeof tp !== "string") return false;
                        return tp.split(".")[0] === bab;
                      })
                      .sort((a: string, b: string) => {
                        const [, aSub] = a.split(".").map(Number);
                        const [, bSub] = b.split(".").map(Number);
                        return aSub - bSub;
                      });
                    if (tpList.length === 0) return `${bab}.1`;
                    const lastTP = tpList[tpList.length - 1];
                    const [mainNum, subNum] = lastTP.split(".").map(Number);
                    return `${mainNum}.${subNum + 1}`;
                  };

                  setEditTP((prev) => ({
                    ...prev,
                    bab: selectedBab,
                    tp:
                      editTP.mapel && selectedBab
                        ? getNextTPForEdit(editTP.mapel, selectedBab)
                        : prev.tp,
                  }));
                }}
                placeholder="Contoh: 1, 2, 3, dst"
                disabled={!editTP.mapel}
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  backgroundColor: !editTP.mapel ? "#f5f5f5" : "white",
                  cursor: !editTP.mapel ? "not-allowed" : "text",
                }}
              />
              {!editTP.mapel && (
                <small style={{ color: "#999", fontSize: "11px" }}>
                  Pilih Mapel terlebih dahulu
                </small>
              )}
            </div>

            {/* TP - auto tapi bisa diedit manual */}
            <div style={{ gridColumn: "1 / -1" }}>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                3️⃣ TP:
                {editTP.mapel && editTP.bab && (
                  <span
                    style={{
                      color: "#FF9800",
                      fontSize: "12px",
                      marginLeft: "8px",
                    }}
                  >
                    ✓ Otomatis terisi (bisa diubah manual)
                  </span>
                )}
              </label>
              <input
                type="text"
                value={editTP.tp}
                onChange={(e) => setEditTP({ ...editTP, tp: e.target.value })}
                placeholder="TP akan terisi otomatis atau isi manual"
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "2px solid " + (editTP.tp ? "#FF9800" : "#ddd"),
                  borderRadius: "4px",
                  backgroundColor: editTP.tp ? "#fff8e1" : "white",
                  fontSize: "16px",
                  fontWeight: "bold",
                  color: editTP.tp ? "#E65100" : "#999",
                  textAlign: "center",
                }}
              />
            </div>

            {/* Rincian TP */}
            <div style={{ gridColumn: "1 / -1" }}>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                Rincian TP:
              </label>
              <textarea
                value={editTP.rincian}
                onChange={(e) =>
                  setEditTP({ ...editTP, rincian: e.target.value })
                }
                rows={3}
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  resize: "vertical",
                }}
              />
            </div>

            {/* Semester & Kelas */}
            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                Semester:
                {editTP.mapel && (
                  <span
                    style={{
                      color: "#FF9800",
                      fontSize: "12px",
                      marginLeft: "5px",
                    }}
                  >
                    (Auto-fill)
                  </span>
                )}
              </label>
              <input
                type="text"
                value={editTP.semester}
                onChange={(e) =>
                  setEditTP({ ...editTP, semester: e.target.value })
                }
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  backgroundColor: editTP.mapel ? "#fff8e1" : "white",
                }}
              />
            </div>

            <div>
              <label
                style={{
                  display: "block",
                  marginBottom: "5px",
                  fontWeight: "bold",
                }}
              >
                Kelas:
                {editTP.mapel && (
                  <span
                    style={{
                      color: "#FF9800",
                      fontSize: "12px",
                      marginLeft: "5px",
                    }}
                  >
                    (Auto-fill)
                  </span>
                )}
              </label>
              <input
                type="text"
                value={editTP.kelas}
                onChange={(e) =>
                  setEditTP({ ...editTP, kelas: e.target.value })
                }
                style={{
                  width: "100%",
                  padding: "8px",
                  border: "1px solid #ddd",
                  borderRadius: "4px",
                  backgroundColor: editTP.mapel ? "#fff8e1" : "white",
                }}
              />
            </div>
          </div>

          {/* Preview */}
          {editTP.mapel && editTP.bab && editTP.tp && (
            <div
              style={{
                marginTop: "15px",
                padding: "12px",
                backgroundColor: "#fff3e0",
                border: "1px solid #FF9800",
                borderRadius: "4px",
              }}
            >
              <strong style={{ color: "#E65100" }}>✓ Siap disimpan:</strong>
              <div
                style={{ marginTop: "5px", fontSize: "14px", color: "#333" }}
              >
                {editTP.mapel} - BAB {editTP.bab} - TP {editTP.tp}
              </div>
            </div>
          )}

          <div style={{ marginTop: "15px" }}>
            {isEditTPLimitReached && (
              <div
                style={{
                  marginBottom: "12px",
                  padding: "12px 16px",
                  backgroundColor: "#fdecea",
                  border: "1px solid #f44336",
                  borderRadius: "6px",
                  color: "#b71c1c",
                  fontSize: "14px",
                }}
              >
                ⛔ <strong>Batas maksimal tercapai!</strong> {editTP.mapel} —
                Semester {editTP.semester} — Kelas {editTP.kelas} sudah memiliki{" "}
                <strong>{editTPCountExcludeSelf} TP</strong> lainnya. Maksimal
                10 TP per mata pelajaran per semester.
              </div>
            )}

            <div style={{ display: "flex", gap: "10px" }}>
              <button
                onClick={handleSaveEdit}
                disabled={isSaving || isEditTPLimitReached}
                style={{
                  padding: "12px 24px",
                  backgroundColor:
                    isSaving || isEditTPLimitReached ? "#ccc" : "#FF9800",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor:
                    isSaving || isEditTPLimitReached
                      ? "not-allowed"
                      : "pointer",
                  fontWeight: "bold",
                  fontSize: "16px",
                }}
              >
                {isSaving ? "Menyimpan..." : "💾 Simpan Perubahan"}
              </button>
              <button
                onClick={() => {
                  setEditingIndex(null);
                  setEditTP({
                    mapel: "",
                    tp: "",
                    rincian: "",
                    bab: "",
                    semester: "",
                    kelas: "",
                  });
                }}
                style={{
                  padding: "12px 24px",
                  backgroundColor: "#9E9E9E",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                  fontWeight: "bold",
                  fontSize: "16px",
                }}
              >
                ❌ Batal
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Filter Bar */}
      <div
        style={{
          display: "flex",
          gap: "12px",
          alignItems: "center",
          marginBottom: "15px",
          padding: "12px 16px",
          backgroundColor: "#f8f9fa",
          borderRadius: "8px",
          border: "1px solid #e0e0e0",
          flexWrap: "wrap",
        }}
      >
        <span style={{ fontWeight: "bold", fontSize: "14px", color: "#555" }}>
          🔍 Filter:
        </span>

        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
          <label
            style={{ fontSize: "13px", color: "#666", whiteSpace: "nowrap" }}
          >
            Mapel:
          </label>
          <select
            value={filterMapel}
            onChange={(e) => setFilterMapel(e.target.value)}
            style={{
              padding: "6px 10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              fontSize: "13px",
              backgroundColor: filterMapel ? "#e3f2fd" : "white",
              color: filterMapel ? "#1565C0" : "#333",
              fontWeight: filterMapel ? "bold" : "normal",
              minWidth: "160px",
            }}
          >
            <option value="">-- Semua Mapel --</option>
            {availableMapel.map((mapel, index) => (
              <option key={index} value={mapel}>
                {mapel}
              </option>
            ))}
          </select>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
          <label
            style={{ fontSize: "13px", color: "#666", whiteSpace: "nowrap" }}
          >
            Semester:
          </label>
          <select
            value={filterSemester}
            onChange={(e) => setFilterSemester(e.target.value)}
            style={{
              padding: "6px 10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              fontSize: "13px",
              backgroundColor: filterSemester ? "#e3f2fd" : "white",
              color: filterSemester ? "#1565C0" : "#333",
              fontWeight: filterSemester ? "bold" : "normal",
              minWidth: "120px",
            }}
          >
            <option value="">-- Semua Semester --</option>
            {availableSemester.map((sem, index) => (
              <option key={index} value={sem}>
                Semester {sem}
              </option>
            ))}
          </select>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
          <label
            style={{ fontSize: "13px", color: "#666", whiteSpace: "nowrap" }}
          >
            Kelas:
          </label>
          <select
            value={filterKelas}
            onChange={(e) => setFilterKelas(e.target.value)}
            style={{
              padding: "6px 10px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              fontSize: "13px",
              backgroundColor: filterKelas ? "#e3f2fd" : "white",
              color: filterKelas ? "#1565C0" : "#333",
              fontWeight: filterKelas ? "bold" : "normal",
              minWidth: "100px",
            }}
          >
            <option value="">-- Semua Kelas --</option>
            {availableKelas.map((kelas, index) => (
              <option key={index} value={kelas}>
                Kelas {kelas}
              </option>
            ))}
          </select>
        </div>

        {(filterMapel || filterSemester || filterKelas) && (
          <button
            onClick={() => {
              setFilterMapel("");
              setFilterSemester("");
              setFilterKelas("");
            }}
            style={{
              padding: "6px 12px",
              backgroundColor: "#ff5722",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
              fontSize: "12px",
              fontWeight: "bold",
            }}
          >
            ✖ Reset Filter
          </button>
        )}

        <span style={{ fontSize: "12px", color: "#888", marginLeft: "auto" }}>
          Menampilkan{" "}
          <strong style={{ color: "#333" }}>{filteredData.length}</strong> dari{" "}
          <strong style={{ color: "#333" }}>{actualData.length}</strong> data
        </span>
      </div>

      {/* Tabel Data TP */}
      <div
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 300px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            width: "100%",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "40px",
                  minWidth: "40px",
                  backgroundColor: "#f4f4f4",
                  fontSize: "12px",
                }}
              >
                No.
              </th>
              {/* Bab (index 3) */}
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "60px",
                  minWidth: "60px",
                  backgroundColor: "#f4f4f4",
                  fontSize: "12px",
                }}
              >
                {displayHeaders[3]}
              </th>
              {/* TP (index 1) */}
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "80px",
                  minWidth: "80px",
                  backgroundColor: "#f4f4f4",
                  fontSize: "12px",
                }}
              >
                {displayHeaders[1]}
              </th>
              {/* Rincian TP (index 2) */}
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "400px",
                  minWidth: "180px",
                  backgroundColor: "#f4f4f4",
                  fontSize: "12px",
                }}
              >
                {displayHeaders[2]}
              </th>

              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "100px",
                  backgroundColor: "#f4f4f4",
                  fontSize: "12px",
                }}
              >
                Aksi
              </th>
            </tr>
          </thead>
          <tbody>
            {filteredData.map((row, filteredIndex) => {
              const rowIndex = actualData.findIndex(
                (r: any) =>
                  r.Data1 === row.Data1 &&
                  r.Data2 === row.Data2 &&
                  r.Data3 === row.Data3
              );

              return (
                <tr
                  key={filteredIndex}
                  style={{
                    backgroundColor:
                      filteredIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                  }}
                >
                  <td
                    style={{
                      padding: "8px",
                      borderBottom: "1px solid #eee",
                      textAlign: "center",
                      fontSize: "12px",
                      color: "#666",
                      verticalAlign: "middle",
                      width: "40px",
                      fontWeight: "bold",
                    }}
                  >
                    {filteredIndex + 1}
                  </td>
                  {/* Bab (Data4, index 3) */}
                  <td
                    style={{
                      padding: "8px",
                      borderBottom: "1px solid #eee",
                      verticalAlign: "middle",
                      textAlign: "center",
                    }}
                  >
                    <input
                      type="text"
                      value={row["Data4"] || ""}
                      readOnly
                      style={{
                        width: "100%",
                        padding: "6px",
                        border: "1px solid #eee",
                        borderRadius: "3px",
                        fontSize: "12px",
                        backgroundColor: "#f9f9f9",
                        cursor: "default",
                        color: "#333",
                        textAlign: "center",
                      }}
                    />
                  </td>
                  {/* TP (Data2, index 1) */}
                  <td
                    style={{
                      padding: "8px",
                      borderBottom: "1px solid #eee",
                      verticalAlign: "middle",
                      textAlign: "center",
                    }}
                  >
                    <input
                      type="text"
                      value={row["Data2"] || ""}
                      readOnly
                      style={{
                        width: "100%",
                        padding: "6px",
                        border: "1px solid #eee",
                        borderRadius: "3px",
                        fontSize: "12px",
                        backgroundColor: "#f9f9f9",
                        cursor: "default",
                        color: "#333",
                        textAlign: "center",
                      }}
                    />
                  </td>
                  {/* Rincian TP (Data3, index 2) */}
                  <td
                    style={{
                      padding: "8px",
                      borderBottom: "1px solid #eee",
                      verticalAlign: "middle",
                    }}
                  >
                    <div
                      style={{
                        width: "100%",
                        padding: "6px",
                        border: "1px solid #eee",
                        borderRadius: "3px",
                        fontSize: "12px",
                        backgroundColor: "#f9f9f9",
                        color: "#333",
                        whiteSpace: "pre-wrap",
                        wordBreak: "break-word",
                        lineHeight: "1.5",
                        minWidth: "180px",
                      }}
                    >
                      {row["Data3"] || ""}
                    </div>
                  </td>
                  <td
                    style={{
                      padding: "8px",
                      borderBottom: "1px solid #eee",
                      textAlign: "center",
                      verticalAlign: "middle",
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                        gap: "6px",
                        justifyContent: "center",
                      }}
                    >
                      <button
                        onClick={() => handleEdit(rowIndex)}
                        disabled={isSaving}
                        style={{
                          padding: "6px 12px",
                          backgroundColor: isSaving ? "#ccc" : "#FF9800",
                          color: "white",
                          border: "none",
                          borderRadius: "4px",
                          cursor: isSaving ? "not-allowed" : "pointer",
                          fontSize: "12px",
                          fontWeight: "bold",
                        }}
                      >
                        ✏️ Edit
                      </button>
                      <button
                        onClick={() => handleDelete(rowIndex)}
                        disabled={isSaving}
                        style={{
                          padding: "6px 12px",
                          backgroundColor: isSaving ? "#ccc" : "#f44336",
                          color: "white",
                          border: "none",
                          borderRadius: "4px",
                          cursor: isSaving ? "not-allowed" : "pointer",
                          fontSize: "12px",
                          fontWeight: "bold",
                        }}
                      >
                        🗑️ Hapus
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DataMapel = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isAddingNew, setIsAddingNew] = useState(false);
  const [newMapel, setNewMapel] = useState("");

  // ✅ Gunakan data dari context
  const { mapelListData, refreshRekapData, updateLocalData } = useRekapData();

  // ✅ TAMBAH INI - Force refresh saat halaman dibuka
  useEffect(() => {
    console.log("DataMapel mounted - forcing immediate refresh");
    setLoading(true);

    const refreshData = async () => {
      try {
        const response = await fetch(`${endpoint}?sheet=DataMapel`);
        if (!response.ok) {
          throw new Error("Failed to fetch DataMapel");
        }

        const jsonData = await response.json();
        setData(jsonData);

        setLoading(false);
        console.log("✅ DataMapel data refreshed successfully");
      } catch (err) {
        console.error("Error refreshing DataMapel:", err);
        setError(err instanceof Error ? err.message : "Unknown error");
        setLoading(false);
      }
    };

    refreshData();

    return () => {
      console.log("DataMapel unmounted");
    };
  }, []); // Empty dependency = hanya run saat mount

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);
    const headers = ["Data1"];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({ rowIndex: rowIndex + 3, values: values });
    });

    try {
      const requestBody = {
        action: "update_bulk",
        sheetName: "DataMapel",
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());
      setIsSaving(false);
      await syncSheetsToContext();
      refreshRekapData(true);
    } catch (err) {
      console.error("=== ERROR DETAILS ===");
      console.error(err);
      alert(
        "Error updating rows: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  const handleAddNew = async () => {
    if (!newMapel.trim()) {
      alert("⚠️ Nama Mata Pelajaran wajib diisi!");
      return;
    }

    setIsSaving(true);

    try {
      const requestBody = {
        action: "add_mapel",
        mapel: newMapel.toUpperCase().trim(),
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      alert("✅ Mata Pelajaran baru berhasil ditambahkan!");

      // ✅ UBAH 3 BARIS INI MENJADI 4 BARIS:
      const newRow = {
        Data1: newMapel.toUpperCase().trim(),
      };
      const updatedData = [...data, newRow]; // ✅ BARIS BARU
      setData(updatedData); // ✅ UBAH: pakai updatedData
      updateLocalData("mapel", updatedData); // ✅ BARIS BARU

      setNewMapel("");
      setIsAddingNew(false);

      // Kembalikan tombol dulu, refresh di background
      setIsSaving(false);

      // Sync dropdown InputNilai langsung
      await syncSheetsToContext();

      // Background refresh lainnya
      refreshRekapData(false);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menambah data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  const handleDelete = async (rowIndex: number) => {
    const rowData = actualData[rowIndex];
    const confirmation = window.confirm(
      `Apakah Anda yakin ingin menghapus mata pelajaran ini?\n\nMapel: ${rowData.Data1}\n\n⚠️ Data akan dikosongkan (baris tetap ada)`
    );

    if (!confirmation) return;

    setIsSaving(true);

    try {
      const requestBody = {
        action: "delete_row",
        sheetName: "DataMapel",
        rowIndex: rowIndex + 3,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("✅ Mata pelajaran berhasil dihapus!");

      // ✅ UBAH 3 BARIS INI MENJADI 4 BARIS:
      const updatedData = [...data];
      updatedData[rowIndex + 1] = { Data1: "" };
      setData(updatedData);
      updateLocalData("mapel", updatedData);

      await syncSheetsToContext();

      setIsSaving(false);
      refreshRekapData(false);
    } catch (err) {
      console.error("Error:", err);
      alert(
        "Error menghapus data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  const syncSheetsToContext = async () => {
    try {
      const sheetsRes = await fetch(`${endpoint}?action=listSheets`);
      if (!sheetsRes.ok) return;

      const sheetsRaw: SheetInfo[] = await sheetsRes.json();
      const filteredSheets = sheetsRaw.filter((sheet) => {
        const mapel = sheet.mapel || "";
        return (
          mapel.trim() !== "" &&
          !mapel.includes("#REF!") &&
          !mapel.includes("#N/A") &&
          !mapel.includes("N/A") &&
          !mapel.toUpperCase().includes("ERROR")
        );
      });

      updateLocalData("sheets", filteredSheets);
      console.log("✅ availableSheets synced:", filteredSheets.length);
    } catch (err) {
      console.error("Gagal sync sheets:", err);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const actualData = data.slice(1);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "800px" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        📚 Data Mata Pelajaran
      </h1>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={() => setIsAddingNew(!isAddingNew)}
          style={{
            padding: "12px 24px",
            backgroundColor: isAddingNew ? "#f44336" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            marginRight: "10px",
          }}
        >
          {isAddingNew ? "❌ Batal Tambah" : "➕ Tambah Mata Pelajaran Baru"}
        </button>

        <button
          onClick={handleSaveAll}
          disabled={isSaving || changedRows.size === 0}
          style={{
            padding: "12px 24px",
            backgroundColor:
              isSaving || changedRows.size === 0 ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor:
              isSaving || changedRows.size === 0 ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
          }}
        >
          {isSaving
            ? "Memproses..."
            : `💾 Simpan Perubahan (${changedRows.size})`}
        </button>
      </div>

      {isAddingNew && (
        <div
          style={{
            backgroundColor: "#f0f8ff",
            padding: "20px",
            borderRadius: "8px",
            marginBottom: "20px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          }}
        >
          <h3 style={{ marginBottom: "15px", color: "#2196F3" }}>
            Form Tambah Mata Pelajaran Baru
          </h3>
          <input
            type="text"
            value={newMapel}
            onChange={(e) => setNewMapel(e.target.value)}
            placeholder="Contoh: BAHASA INGGRIS"
            style={{
              width: "100%",
              padding: "12px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              fontSize: "16px",
              marginBottom: "15px",
            }}
          />
          <button
            onClick={handleAddNew}
            disabled={isSaving}
            style={{
              padding: "12px 24px",
              backgroundColor: isSaving ? "#ccc" : "#4CAF50",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "16px",
            }}
          >
            {isSaving ? "Menyimpan..." : "💾 Simpan Mata Pelajaran Baru"}
          </button>
        </div>
      )}

      <div
        style={{
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          overflow: "hidden",
        }}
      >
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead style={{ backgroundColor: "#f4f4f4" }}>
            <tr>
              <th
                style={{
                  padding: "12px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "80px",
                }}
              >
                No
              </th>
              <th
                style={{
                  padding: "12px",
                  textAlign: "left",
                  borderBottom: "2px solid #ddd",
                }}
              >
                Nama Mata Pelajaran
              </th>
              <th
                style={{
                  padding: "12px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "120px",
                }}
              >
                Aksi
              </th>
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "12px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                  }}
                >
                  {rowIndex + 1}
                </td>
                <td style={{ padding: "8px", borderBottom: "1px solid #eee" }}>
                  <input
                    type="text"
                    value={row.Data1 || ""}
                    onChange={(e) =>
                      handleInputChange(rowIndex, "Data1", e.target.value)
                    }
                    style={{
                      width: "100%",
                      padding: "8px",
                      border: "1px solid #ddd",
                      borderRadius: "4px",
                      fontSize: "14px",
                    }}
                  />
                </td>
                {/* ✅ KOLOM AKSI BARU */}
                <td
                  style={{
                    padding: "8px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                  }}
                >
                  <button
                    onClick={() => handleDelete(rowIndex)}
                    disabled={isSaving}
                    style={{
                      padding: "8px 16px",
                      backgroundColor: isSaving ? "#ccc" : "#f44336",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: isSaving ? "not-allowed" : "pointer",
                      fontSize: "14px",
                      fontWeight: "bold",
                    }}
                  >
                    🗑️ Hapus
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DataKokurikuler = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [dataSnapshot, setDataSnapshot] = useState<RowData[]>([]);
  const [selectedSemester, setSelectedSemester] = useState<string>("1");

  // ✅ Gunakan data dari context
  const { kokurikulerData, refreshRekapData, updateLocalData } = useRekapData(); // ✅ Tambah updateLocalData

  // ✅ BARU: Force refresh hanya DataKokurikuler saat mount
  useEffect(() => {
    console.log("DataKokurikuler mounted - forcing immediate refresh");
    setLoading(true);
    setChangedRows(new Set()); // reset perubahan saat ganti semester
    setIsEditing(false);

    const refreshDataKokurikuler = async () => {
      try {
        const response = await fetch(
          `${endpoint}?sheet=DataKokurikuler${selectedSemester}`
        );
        if (!response.ok) {
          throw new Error("Failed to fetch DataKokurikuler");
        }

        const jsonData = await response.json();
        setData(jsonData);
        updateLocalData("kokurikuler", jsonData); // Update context
        setLoading(false);
        console.log("✅ DataKokurikuler refreshed successfully");
      } catch (err) {
        console.error("Error refreshing DataKokurikuler:", err);
        setError(err instanceof Error ? err.message : "Unknown error");
        setLoading(false);
      }
    };

    refreshDataKokurikuler();

    return () => {
      console.log("DataKokurikuler unmounted");
    };
  }, [selectedSemester]);

  const editableKokurikulerHeaders = [
    "Data2",
    "Data3",
    "Data4",
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
  ];

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = JSON.parse(JSON.stringify(data));

    // Jika nilai yang dipilih adalah "1" atau "4", kosongkan kolom lain
    // yang memiliki nilai sama dalam baris yang sama
    if (value === "1" || value === "4") {
      editableKokurikulerHeaders.forEach((h) => {
        if (
          h !== header &&
          String(updatedData[rowIndex + 1][h]).trim() === value
        ) {
          updatedData[rowIndex + 1][h] = "";
        }
      });
    }

    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => {
      const newSet = new Set(prev);
      newSet.add(rowIndex);
      return newSet;
    });
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);

    // Data10 (Deskripsi) tidak disertakan karena berisi formula
    const headers = [
      "Data1",
      "Data2",
      "Data3",
      "Data4",
      "Data5",
      "Data6",
      "Data7",
      "Data8",
      "Data9",
    ];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({
        rowIndex: rowIndex + 3,
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_kokurikuler_bulk",
        sheetName: `DataKokurikuler${selectedSemester}`,
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());
      setIsSaving(false);
      setIsEditing(false); // ← TAMBAHKAN INI
      setLoading(true);

      try {
        const response = await fetch(
          `${endpoint}?sheet=DataKokurikuler${selectedSemester}`
        );
        if (response.ok) {
          const jsonData = await response.json();
          setData(jsonData);
          updateLocalData("kokurikuler", jsonData);
        }
      } catch (err) {
        console.error("Error refreshing DataKokurikuler:", err);
      } finally {
        setLoading(false);
      }
    } catch (err) {
      console.error("=== ERROR DETAILS ===");
      console.error(err);
      alert(
        "Error updating rows: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const headers = [
    "Data1",
    "Data2",
    "Data3",
    "Data4",
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
  ];
  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  // Data1=nama siswa (frozen, readonly)
  // Data2-Data9 = kolom nilai kokurikuler (editable)
  // Data10 = Deskripsi (readonly)
  const readOnlyHeaders = new Set(["Data1", "Data10"]);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        🌟 Data Kokurikuler
      </h1>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <label style={{ fontSize: "14px", color: "#666", marginRight: "10px" }}>
          Semester:
        </label>
        <select
          value={selectedSemester}
          onChange={(e) => setSelectedSemester(e.target.value)}
          style={{
            padding: "10px 15px",
            fontSize: "16px",
            borderRadius: "4px",
            border: "1px solid #ddd",
            minWidth: "150px",
            cursor: "pointer",
            backgroundColor: "white",
          }}
        >
          <option value="1">Semester 1</option>
          <option value="2">Semester 2</option>
        </select>
      </div>

      <div
        style={{
          textAlign: "center",
          marginBottom: "15px",
          display: "flex",
          gap: "10px",
          justifyContent: "center",
        }}
      >
        <button
          onClick={handleSaveAll}
          disabled={isSaving || !isEditing}
          style={{
            padding: "12px 24px",
            backgroundColor: isSaving || !isEditing ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving || !isEditing ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            maxWidth: "300px",
          }}
        >
          {isSaving ? "Memproses..." : `Simpan Perubahan (${changedRows.size})`}
        </button>
        <button
          onClick={() => {
            if (!isEditing) {
              // Mulai edit: simpan deep copy snapshot data saat ini
              setDataSnapshot(JSON.parse(JSON.stringify(data)));
              setIsEditing(true);
            } else {
              // Batal edit: kembalikan data ke deep copy snapshot
              setData(JSON.parse(JSON.stringify(dataSnapshot)));
              setChangedRows(new Set());
              setIsEditing(false);
            }
          }}
          disabled={isSaving}
          style={{
            padding: "12px 24px",
            backgroundColor: isEditing ? "#f44336" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            maxWidth: "300px",
          }}
        >
          {isEditing ? "❌ Batal Edit" : "✏️ Edit"}
        </button>
      </div>

      <div
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 200px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
            tableLayout: "fixed",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              <th
                style={{
                  padding: "8px 4px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "40px",
                  minWidth: "40px",
                  position: "sticky",
                  left: 0,
                  backgroundColor: "#f4f4f4",
                  zIndex: 2,
                  boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  fontSize: "12px",
                }}
              >
                No.
              </th>
              {displayHeaders.map((header, index) => {
                const currentHeader = headers[index];
                const isNameColumn = currentHeader === "Data1";
                return (
                  <th
                    key={index}
                    style={{
                      padding: "8px 4px",
                      textAlign: "center",
                      borderBottom: "2px solid #ddd",
                      fontWeight: "bold",
                      width: isNameColumn ? "100px" : "120px",
                      minWidth: isNameColumn ? "100px" : "120px",
                      position: isNameColumn ? "sticky" : "static",
                      left: isNameColumn ? 0 : "auto",
                      backgroundColor: "#f4f4f4",
                      zIndex: isNameColumn ? 2 : 1,
                      boxShadow: isNameColumn
                        ? "2px 0 5px rgba(0,0,0,0.1)"
                        : "none",
                      fontSize: "12px",
                    }}
                  >
                    {header}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "6px 4px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                    fontWeight: "bold",
                    color: "#666",
                    width: "40px",
                    minWidth: "40px",
                    position: "sticky",
                    left: 0,
                    backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                    zIndex: 1,
                    boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                    fontSize: "12px",
                  }}
                >
                  {rowIndex + 1}
                </td>
                {headers.map((header, colIndex) => {
                  const isNameColumn = header === "Data1";
                  const isReadOnly = readOnlyHeaders.has(header);
                  return (
                    <td
                      key={colIndex}
                      style={{
                        padding: "4px",
                        borderBottom: "1px solid #eee",
                        position: isNameColumn ? "sticky" : "static",
                        left: isNameColumn ? 0 : "auto",
                        backgroundColor: isNameColumn
                          ? rowIndex % 2 === 0
                            ? "#fff"
                            : "#f9f9f9"
                          : "transparent",
                        zIndex: isNameColumn ? 1 : 0,
                        boxShadow: isNameColumn
                          ? "2px 0 5px rgba(0,0,0,0.1)"
                          : "none",
                      }}
                    >
                      {isReadOnly ? (
                        <div
                          style={{
                            padding: "4px 2px",
                            color: "#666",
                            fontSize: "12px",
                            textAlign: isNameColumn ? "left" : "center",
                          }}
                        >
                          {row[header] || ""}
                        </div>
                      ) : (
                        <select
                          value={row[header] || ""}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          disabled={!isEditing}
                          style={{
                            width: "100%",
                            padding: "4px 2px",
                            border: "1px solid #ddd",
                            borderRadius: "3px",
                            boxSizing: "border-box",
                            backgroundColor: isEditing ? "white" : "#f5f5f5",
                            fontSize: "12px",
                            textAlign: "center",
                            cursor: isEditing ? "pointer" : "not-allowed",
                          }}
                        >
                          <option value="">-</option>
                          <option value="1">1</option>
                          <option value="4">4</option>
                        </select>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DataEkstrakurikuler = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [dataSnapshot, setDataSnapshot] = useState<RowData[]>([]);
  const [selectedSemester, setSelectedSemester] = useState<string>("1");

  // ✅ Gunakan data dari context
  const { ekstrakurikulerData, refreshRekapData } = useRekapData();

  useEffect(() => {
    console.log("DataEkstrakurikuler mounted - forcing immediate refresh");
    setLoading(true);
    setChangedRows(new Set());
    setIsEditing(false);

    const refreshData = async () => {
      try {
        const response = await fetch(
          `${endpoint}?sheet=DataEkstrakurikuler${selectedSemester}`
        );
        if (!response.ok) {
          throw new Error("Failed to fetch DataEkstrakurikuler");
        }

        const jsonData = await response.json();
        setData(jsonData);
        setLoading(false);
        console.log("✅ DataEkstrakurikuler refreshed successfully");
      } catch (err) {
        console.error("Error refreshing DataEkstrakurikuler:", err);
        setError(err instanceof Error ? err.message : "Unknown error");
        setLoading(false);
      }
    };

    refreshData();
  }, [selectedSemester]);

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);

    const headers = ["Data1", "Data2", "Data3", "Data4", "Data5", "Data6"];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      updates.push({
        rowIndex: rowIndex + 3,
        values: values,
      });
    });

    try {
      const requestBody = {
        action: "update_ekstrakurikuler_bulk",
        sheetName: `DataEkstrakurikuler${selectedSemester}`,
        updates: updates,
      };
      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      alert("Semua perubahan berhasil disimpan!");
      setChangedRows(new Set());
      setIsSaving(false);
      setIsEditing(false);
      setLoading(true);

      try {
        const response = await fetch(
          `${endpoint}?sheet=DataEkstrakurikuler${selectedSemester}`
        );
        if (response.ok) {
          const jsonData = await response.json();
          setData(jsonData);
        }
      } catch (err) {
        console.error("Error refreshing DataEkstrakurikuler:", err);
      } finally {
        setLoading(false);
      }
    } catch (err) {
      console.error("=== ERROR DETAILS ===");
      console.error(err);
      alert(
        "Error updating rows: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
      setIsSaving(false);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const headers = ["Data1", "Data2", "Data3", "Data4", "Data5", "Data6"];
  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  const readOnlyHeaders = new Set([
    "Data1",
    "Data3",
    "Data5",
    "Data6",
    "Data7",
  ]);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        🎯 Data Ekstrakurikuler
      </h1>

      <div
        style={{
          textAlign: "center",
          marginBottom: "15px",
          display: "flex",
          gap: "10px",
          justifyContent: "center",
        }}
      >
        <label style={{ fontSize: "14px", color: "#666", marginRight: "10px" }}>
          Semester:
        </label>
        <select
          value={selectedSemester}
          onChange={(e) => setSelectedSemester(e.target.value)}
          style={{
            padding: "10px 15px",
            fontSize: "16px",
            borderRadius: "4px",
            border: "1px solid #ddd",
            minWidth: "150px",
            cursor: "pointer",
            backgroundColor: "white",
          }}
        >
          <option value="1">Semester 1</option>
          <option value="2">Semester 2</option>
        </select>
      </div>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={handleSaveAll}
          disabled={isSaving || !isEditing}
          style={{
            padding: "12px 24px",
            backgroundColor: isSaving || !isEditing ? "#ccc" : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving || !isEditing ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            width: "100%",
            maxWidth: "300px",
          }}
        >
          {isSaving ? "Memproses..." : `Simpan Perubahan (${changedRows.size})`}
        </button>
        <button
          onClick={() => {
            if (!isEditing) {
              setDataSnapshot(JSON.parse(JSON.stringify(data)));
              setIsEditing(true);
            } else {
              setData(JSON.parse(JSON.stringify(dataSnapshot)));
              setChangedRows(new Set());
              setIsEditing(false);
            }
          }}
          disabled={isSaving}
          style={{
            padding: "12px 24px",
            backgroundColor: isEditing ? "#f44336" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            maxWidth: "300px",
          }}
        >
          {isEditing ? "❌ Batal Edit" : "✏️ Edit"}
        </button>
      </div>

      <div
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 200px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
            tableLayout: "fixed",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              {displayHeaders.map((header, index) => {
                const currentHeader = headers[index];
                const isNameColumn = currentHeader === "Data1";
                return (
                  <th
                    key={index}
                    style={{
                      padding: "8px 4px",
                      textAlign: "center",
                      borderBottom: "2px solid #ddd",
                      fontWeight: "bold",
                      width: isNameColumn ? "200px" : "120px",
                      minWidth: isNameColumn ? "200px" : "120px",
                      position: isNameColumn ? "sticky" : "static",
                      left: isNameColumn ? 0 : "auto",
                      backgroundColor: "#f4f4f4",
                      zIndex: isNameColumn ? 2 : 1,
                      boxShadow: isNameColumn
                        ? "2px 0 5px rgba(0,0,0,0.1)"
                        : "none",
                      fontSize: "12px",
                    }}
                  >
                    {header}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                {headers.map((header, colIndex) => {
                  const isNameColumn = header === "Data1";
                  const isReadOnly = readOnlyHeaders.has(header);
                  return (
                    <td
                      key={colIndex}
                      style={{
                        padding: "4px",
                        borderBottom: "1px solid #eee",
                        position: isNameColumn ? "sticky" : "static",
                        left: isNameColumn ? 0 : "auto",
                        backgroundColor: isNameColumn
                          ? rowIndex % 2 === 0
                            ? "#fff"
                            : "#f9f9f9"
                          : "transparent",
                        zIndex: isNameColumn ? 1 : 0,
                        boxShadow: isNameColumn
                          ? "2px 0 5px rgba(0,0,0,0.1)"
                          : "none",
                      }}
                    >
                      {isReadOnly ? (
                        <div
                          style={{
                            padding: "4px 2px",
                            color: "#666",
                            fontSize: "12px",
                            textAlign: isNameColumn ? "left" : "center",
                          }}
                        >
                          {row[header] || ""}
                        </div>
                      ) : header === "Data2" ? (
                        <select
                          value={row[header] || ""}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          disabled={!isEditing}
                          style={{
                            width: "100%",
                            padding: "4px 2px",
                            border: "1px solid #ddd",
                            borderRadius: "3px",
                            boxSizing: "border-box" as const,
                            backgroundColor: isEditing ? "white" : "#f5f5f5",
                            cursor: isEditing ? "pointer" : "not-allowed",
                            fontSize: "12px",
                            textAlign: "center",
                          }}
                        >
                          <option value="">-- Pilih --</option>
                          <option value="PRAMUKA">PRAMUKA</option>
                        </select>
                      ) : (
                        <input
                          type="text"
                          inputMode="decimal"
                          value={row[header] || ""}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          disabled={!isEditing}
                          style={{
                            width: "100%",
                            padding: "4px 2px",
                            border: "1px solid #ddd",
                            borderRadius: "3px",
                            boxSizing: "border-box" as const,
                            backgroundColor: isEditing ? "white" : "#f5f5f5",
                            cursor: isEditing ? "text" : "not-allowed",
                            fontSize: "12px",
                            textAlign: "center",
                          }}
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const DataSiswa = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [changedRows, setChangedRows] = useState<Set<number>>(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [dataSnapshot, setDataSnapshot] = useState<RowData[]>([]);
  const [kelasOptions, setKelasOptions] = useState<string[]>([]);

  const [isAddingNew, setIsAddingNew] = useState(false);
  const [deleteNilai, setDeleteNilai] = useState<{
    [rowIndex: number]: boolean;
  }>({});
  const [isImporting, setIsImporting] = useState(false);
  const [importSiswaResult, setImportSiswaResult] = useState<{
    added: number;
    skipped: string[];
  } | null>(null);
  const [importSiswaPreview, setImportSiswaPreview] = useState<{
    toImport: any[];
    skipped: string[];
  } | null>(null);

  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
  const [isDeletingBulk, setIsDeletingBulk] = useState(false);
  const [deleteNilaiBulk, setDeleteNilaiBulk] = useState(false);
  const [newSiswa, setNewSiswa] = useState({
    nama: "",
    kelas: "",
    nis: "",
    nisn: "",
    namaOrtu: "",
  });

  const { refreshRekapData } = useRekapData();

  // ✅ Load data dari server saat mount
  const processSiswaData = (jsonData: any[]) => {
    if (jsonData.length > 1) {
      const headers = jsonData[0];
      const filtered = jsonData
        .slice(1)
        .filter((row: any) => row.Data1 && String(row.Data1).trim() !== "");
      setData([headers, ...filtered]);
    } else {
      setData(jsonData);
    }
  };

  const loadDataFromServer = async (showLoading: boolean = true) => {
    if (showLoading) setLoading(true);
    setError(null);

    try {
      // ✅ Cek IndexedDB dulu
      const cached = await idbLoad("siswaData");
      if (cached && cached.length > 1) {
        processSiswaData(cached);
        setLoading(false);
        console.log("✅ DataSiswa: dari IndexedDB");
      }

      // ✅ Background sync dari server
      const response = await fetch(`${endpoint}?sheet=DataSiswa`);
      if (!response.ok) throw new Error("Failed to fetch DataSiswa");

      const jsonData = await response.json();

      // Simpan ke IndexedDB
      await idbSave("siswaData", jsonData);
      console.log("✅ DataSiswa: IndexedDB diperbarui dari server");

      processSiswaData(jsonData);
    } catch (err) {
      console.error("Error loading DataSiswa:", err);
      if (!data.length) {
        setError(err instanceof Error ? err.message : "Unknown error");
      }
    } finally {
      setLoading(false);
    }
  };

  // Load saat pertama kali
  useEffect(() => {
    loadDataFromServer();
    loadKelasOptions();
  }, []);

  const loadKelasOptions = async () => {
    try {
      const res = await fetch(`${endpoint}?action=schoolData`);
      if (!res.ok) return;
      const json = await res.json();
      if (json.success && json.data?.length > 0) {
        const record = json.data[0];
        const kelas = String(record.kelas || "").trim();
        const rombel = String(record.rombel || "").trim();

        if (kelas) {
          // ✅ Jika rombel kosong atau "-", gunakan kelas saja
          const kelasRombel =
            rombel && rombel !== "-" ? `${kelas}${rombel}` : kelas;
          setKelasOptions([kelasRombel]);
        }
      }
    } catch (err) {
      console.error("Gagal fetch kelas options:", err);
    }
  };

  const handleInputChange = (
    rowIndex: number,
    header: string,
    value: string
  ) => {
    const updatedData = [...data];
    updatedData[rowIndex + 1][header] = value;
    setData(updatedData);
    setChangedRows((prev) => new Set([...Array.from(prev), rowIndex]));
  };

  const handleSaveAll = async () => {
    if (changedRows.size === 0) {
      alert("Tidak ada perubahan untuk disimpan!");
      return;
    }

    setIsSaving(true);
    const headers = ["Data1", "Data2", "Data3", "Data4", "Data5"];
    const updates: Array<{ rowIndex: number; values: string[] }> = [];

    changedRows.forEach((rowIndex) => {
      const rowData = data[rowIndex + 1];
      const values = headers.map((header) => rowData[header] || "");
      const sheetRowIndex = Number(rowData._rowIndex ?? rowIndex + 3);
      updates.push({ rowIndex: sheetRowIndex, values: values });
    });

    try {
      const requestBody = {
        action: "update_siswa_bulk",
        sheetName: "DataSiswa",
        updates: updates,
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      const result = await response.json();
      if (!result.success) throw new Error(result.error || "Gagal menyimpan");

      // ✅ Update IndexedDB dengan data yang sudah diedit (tanpa fetch ulang)
      await idbSave("siswaData", data);

      setChangedRows(new Set());
      setIsEditing(false);
      setDataSnapshot([]);

      alert("✅ Semua perubahan berhasil disimpan!");

      // Background sync untuk pastikan data sinkron dengan server
      loadDataFromServer(false);
      setTimeout(() => refreshRekapData(true), 2000);
    } catch (err) {
      alert(
        "❌ Error updating rows: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const handleBulkDelete = async () => {
    if (selectedRows.size === 0) return;

    const sortedIndexes = Array.from(selectedRows).sort((a, b) => b - a);
    const selectedNames = sortedIndexes
      .map((i) => actualData[i]?.Data1 || "")
      .filter(Boolean);

    const confirmation = window.confirm(
      `⚠️ Hapus ${selectedRows.size} siswa sekaligus?\n\n${selectedNames.join(
        "\n"
      )}\n\n` +
        (deleteNilaiBulk
          ? "⚠️ Nilai di Input Nilai JUGA akan dikosongkan!"
          : "ℹ️ Nilai di Input Nilai TIDAK akan dihapus.")
    );
    if (!confirmation) return;

    setIsDeletingBulk(true);

    try {
      // ✅ Kumpulkan semua data yang diperlukan
      const rowIndexes = sortedIndexes.map(
        (i) => Number(actualData[i]._rowIndex) || i + 3
      );

      const siswaList = sortedIndexes.map((i) => ({
        nama: actualData[i].Data1,
        kelas: actualData[i].Data2,
      }));

      // ✅ Kirim SATU request untuk semua
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "delete_siswa_bulk",
          rowIndexes,
          deleteNilai: deleteNilaiBulk,
          siswaList,
        }),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      const result = await response.json();
      if (!result.success) throw new Error(result.error || "Gagal menghapus");

      // ✅ Update UI
      const selectedSet = new Set(sortedIndexes);
      const newData = [
        data[0],
        ...actualData.filter((_, i) => !selectedSet.has(i)),
      ];
      setData(newData);
      await idbSave("siswaData", newData);

      setSelectedRows(new Set());
      setDeleteNilaiBulk(false);

      if (deleteNilaiBulk) {
        alert(
          `✅ ${result.deletedCount} siswa berhasil dihapus!\n\n` +
            `🗑️ Nilai dikosongkan dari ${result.affectedMapelSheets} sheet mapel.`
        );
      } else {
        alert(`✅ ${result.deletedCount} siswa berhasil dihapus!`);
      }

      loadDataFromServer(false);
      setTimeout(() => refreshRekapData(true), 2000);
    } catch (err) {
      alert(
        "❌ Error: " + (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsDeletingBulk(false);
    }
  };

  const handleImportSiswa = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = "";

    // ✅ Ambil kelas dari kelasOptions (dari data sekolah), bukan dari Excel
    const kelasFromSekolah = kelasOptions[0] || "";

    if (!kelasFromSekolah) {
      alert(
        "⚠️ Data kelas belum tersedia. Pastikan Data Sekolah sudah diisi (Kelas & Rombel)."
      );
      return;
    }

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        setIsImporting(true);
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, {
          defval: "",
        });

        if (jsonData.length === 0) {
          alert("⚠️ File Excel kosong.");
          setIsImporting(false);
          return;
        }

        const existingNames = actualData.map((r) =>
          String(r.Data1 || "")
            .trim()
            .toLowerCase()
        );

        const toImport: any[] = [];
        const skipped: string[] = [];

        jsonData.forEach((row) => {
          const nama = String(row["Nama Lengkap"] || "").trim();
          if (!nama) return;

          if (existingNames.includes(nama.toLowerCase())) {
            skipped.push(nama);
            return;
          }

          const nis = String(row["NIS"] || "").trim();
          const nisn = String(row["NISN"] || "").trim();
          const namaOrtu = String(row["Nama Ayah"] || "").trim();

          // ✅ Kelas selalu dari data sekolah, abaikan kelas dari Excel
          toImport.push({ nama, kelas: kelasFromSekolah, nis, nisn, namaOrtu });
        });

        if (toImport.length === 0) {
          alert("⚠️ Tidak ada data baru untuk diimport.");
          setIsImporting(false);
          return;
        }

        setImportSiswaPreview({ toImport, skipped });
      } catch (err) {
        alert(
          "❌ Gagal membaca file: " +
            (err instanceof Error ? err.message : "Unknown error")
        );
      } finally {
        setIsImporting(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleConfirmImport = async () => {
    if (!importSiswaPreview) {
      console.warn("handleConfirmImport dipanggil tanpa preview aktif!");
      return;
    }
    if (isImporting) return; // prevent double call

    setIsImporting(true);

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "import_siswa_bulk",
          data: importSiswaPreview.toImport,
        }),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);
      const result = await response.json();
      if (!result.success) throw new Error(result.error || "Import gagal");

      setImportSiswaPreview(null);
      setImportSiswaResult({
        added: importSiswaPreview.toImport.length,
        skipped: importSiswaPreview.skipped,
      });

      await loadDataFromServer(false);
      setTimeout(() => {
        refreshRekapData(true);
      }, 2000);
    } catch (err) {
      alert(
        "❌ Gagal import: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsImporting(false);
    }
  };

  const handleAddNew = async () => {
    if (!newSiswa.nama.trim() || !newSiswa.kelas.trim()) {
      alert("⚠️ Nama dan Kelas wajib diisi!");
      return;
    }

    setIsSaving(true);

    const namaSiswa = newSiswa.nama.trim();
    const kelas = newSiswa.kelas.trim();
    const nis = newSiswa.nis.trim();
    const nisn = newSiswa.nisn.trim();
    const namaOrtu = newSiswa.namaOrtu.trim();

    // ✅ OPTIMISTIC: Tambah ke UI langsung
    const newRow: RowData = {
      Data1: namaSiswa,
      Data2: kelas,
      Data3: nis,
      Data4: nisn,
      Data5: namaOrtu,
      _rowIndex: "-1",
    };
    const optimisticData = [...data, newRow];
    setData(optimisticData);
    await idbSave("siswaData", optimisticData);

    setNewSiswa({ nama: "", kelas: "", nis: "", nisn: "", namaOrtu: "" });
    setIsAddingNew(false);

    try {
      const requestBody = {
        action: "add_siswa",
        data: { nama: namaSiswa, kelas, nis, nisn, namaOrtu },
      };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      alert(`✅ Data siswa "${namaSiswa}" berhasil ditambahkan!`);

      // Background sync untuk dapatkan _rowIndex yang benar
      loadDataFromServer(false);
      setTimeout(() => refreshRekapData(true), 2000);
    } catch (err) {
      // ✅ ROLLBACK jika gagal
      setData(data);
      await idbSave("siswaData", data);
      setNewSiswa({ nama: namaSiswa, kelas, nis, nisn, namaOrtu });
      setIsAddingNew(true);
      alert(
        "❌ Error menambah data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (rowIndex: number) => {
    const rowData = actualData[rowIndex];
    const hapusNilai = deleteNilai[rowIndex] || false;

    const confirmation = window.confirm(
      `Apakah Anda yakin ingin menghapus data siswa ini?\n\nNama: ${rowData.Data1}\nKelas: ${rowData.Data2}\n\n` +
        (hapusNilai
          ? "⚠️ Nilai di halaman Input Nilai JUGA akan dikosongkan!"
          : "ℹ️ Nilai di halaman Input Nilai TIDAK akan dihapus.")
    );
    if (!confirmation) return;

    setIsSaving(true);

    try {
      const requestBody = {
        action: "delete_row",
        sheetName: "DataSiswa",
        rowIndex: Number(rowData._rowIndex) || rowIndex + 3,
        deleteNilai: hapusNilai,
        namaSiswa: rowData.Data1,
        kelas: rowData.Data2,
      };

      // ✅ Tunggu server selesai dulu (termasuk hapus nilai jika dicentang)
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);

      const result = await response.json();
      if (!result.success) throw new Error(result.error || "Gagal menghapus");

      // ✅ Update UI setelah server confirm berhasil
      const newData = [data[0], ...actualData.filter((_, i) => i !== rowIndex)];
      setData(newData);
      await idbSave("siswaData", newData);

      setDeleteNilai((prev) => {
        const updated = { ...prev };
        delete updated[rowIndex];
        return updated;
      });

      // ✅ Alert muncul setelah semua proses selesai
      if (hapusNilai) {
        alert(
          `✅ Data siswa "${rowData.Data1}" berhasil dihapus!\n\n` +
            `🗑️ Nilai juga berhasil dikosongkan dari ${result.affectedMapelSheets} sheet mapel.`
        );
      } else {
        alert(`✅ Data siswa "${rowData.Data1}" berhasil dihapus!`);
      }

      loadDataFromServer(false);
      setTimeout(() => refreshRekapData(true), 2000);
    } catch (err) {
      alert(
        "❌ Error menghapus data: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsSaving(false);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const headers = ["Data1", "Data2", "Data3", "Data4", "Data5"];
  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        👨‍🎓 Data Siswa
      </h1>

      <div
        style={{
          textAlign: "center",
          marginBottom: "15px",
          display: "flex",
          gap: "10px",
          justifyContent: "center",
          flexWrap: "wrap",
        }}
      >
        <button
          onClick={() => setIsAddingNew(!isAddingNew)}
          style={{
            padding: "12px 24px",
            backgroundColor: isAddingNew ? "#f44336" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "16px",
          }}
        >
          {isAddingNew ? "❌ Batal Tambah" : "➕ Tambah Siswa Baru"}
        </button>

        <label
          htmlFor="import-siswa-excel"
          style={{
            padding: "12px 24px",
            backgroundColor: isImporting ? "#ccc" : "#FF9800",
            color: "white",
            borderRadius: "4px",
            cursor: isImporting ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            display: "inline-block",
          }}
        >
          {isImporting ? "⏳ Mengimport..." : "📥 Import Excel"}
          <input
            id="import-siswa-excel"
            type="file"
            accept=".xlsx,.xls"
            style={{ display: "none" }}
            onChange={handleImportSiswa}
            disabled={isImporting}
          />
        </label>

        <button
          onClick={() => {
            if (!isEditing) {
              setDataSnapshot(JSON.parse(JSON.stringify(data)));
              setIsEditing(true);
            } else {
              const confirm = window.confirm(
                "⚠️ Batal edit?\n\nSemua perubahan yang belum disimpan akan dikembalikan."
              );
              if (!confirm) return;
              setData(JSON.parse(JSON.stringify(dataSnapshot)));
              setChangedRows(new Set());
              setIsEditing(false);
            }
          }}
          disabled={isSaving}
          style={{
            padding: "12px 24px",
            backgroundColor: isEditing ? "#FF9800" : "#9C27B0",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isSaving ? "not-allowed" : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
          }}
        >
          {isEditing ? "❌ Batal Edit" : "✏️ Edit Data Siswa"}
        </button>

        <button
          onClick={handleSaveAll}
          disabled={isSaving || changedRows.size === 0 || !isEditing}
          style={{
            padding: "12px 24px",
            backgroundColor:
              isSaving || changedRows.size === 0 || !isEditing
                ? "#ccc"
                : "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor:
              isSaving || changedRows.size === 0 || !isEditing
                ? "not-allowed"
                : "pointer",
            fontWeight: "bold",
            fontSize: "16px",
          }}
        >
          {isSaving
            ? "Memproses..."
            : `💾 Simpan Perubahan (${changedRows.size})`}
        </button>
      </div>

      {selectedRows.size > 0 && (
        <div
          style={{
            display: "flex",
            gap: "10px",
            alignItems: "center",
            padding: "10px 16px",
            backgroundColor: "#ffebee",
            borderRadius: "8px",
            border: "1px solid #ef9a9a",
            flexWrap: "wrap",
            justifyContent: "center",
          }}
        >
          <span
            style={{ fontWeight: "bold", color: "#c62828", fontSize: "14px" }}
          >
            🗑️ {selectedRows.size} siswa dipilih
          </span>
          <label
            style={{
              display: "flex",
              alignItems: "center",
              gap: "5px",
              fontSize: "12px",
              color: "#e65100",
              cursor: "pointer",
            }}
          >
            <input
              type="checkbox"
              checked={deleteNilaiBulk}
              onChange={(e) => setDeleteNilaiBulk(e.target.checked)}
              style={{ cursor: "pointer", width: "14px", height: "14px" }}
            />
            Hapus nilai juga
          </label>
          <button
            onClick={handleBulkDelete}
            disabled={isDeletingBulk}
            style={{
              padding: "8px 16px",
              backgroundColor: isDeletingBulk ? "#ccc" : "#f44336",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isDeletingBulk ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "13px",
            }}
          >
            {isDeletingBulk
              ? "⏳ Menghapus..."
              : `🗑️ Hapus ${selectedRows.size} Siswa`}
          </button>
          <button
            onClick={() => setSelectedRows(new Set())}
            disabled={isDeletingBulk}
            style={{
              padding: "8px 16px",
              backgroundColor: "#9E9E9E",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
              fontWeight: "bold",
              fontSize: "13px",
            }}
          >
            ✖ Batal Pilih
          </button>
        </div>
      )}

      {isAddingNew && (
        <div
          style={{
            backgroundColor: "#f0f8ff",
            padding: "20px",
            borderRadius: "8px",
            marginBottom: "20px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          }}
        >
          <h3 style={{ marginBottom: "15px", color: "#2196F3" }}>
            Form Tambah Siswa Baru
          </h3>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "15px",
            }}
          >
            <input
              type="text"
              value={newSiswa.nama}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, nama: e.target.value })
              }
              placeholder="Nama Siswa *"
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
              }}
            />
            <select
              value={newSiswa.kelas}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, kelas: e.target.value })
              }
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
                backgroundColor: "white",
                cursor: "pointer",
              }}
            >
              <option value="">-- Pilih Kelas *--</option>
              {kelasOptions.map((k) => (
                <option key={k} value={k}>
                  {k}
                </option>
              ))}
            </select>
            <input
              type="text"
              value={newSiswa.nis}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, nis: e.target.value })
              }
              placeholder="NIS"
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
              }}
            />
            <input
              type="text"
              value={newSiswa.nisn}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, nisn: e.target.value })
              }
              placeholder="NISN"
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
              }}
            />
            <input
              type="text"
              value={newSiswa.namaOrtu}
              onChange={(e) =>
                setNewSiswa({ ...newSiswa, namaOrtu: e.target.value })
              }
              placeholder="Nama Orang Tua"
              style={{
                padding: "12px",
                border: "1px solid #ddd",
                borderRadius: "4px",
                fontSize: "16px",
                gridColumn: "1 / -1",
              }}
            />
          </div>
          <button
            onClick={handleAddNew}
            disabled={isSaving}
            style={{
              marginTop: "15px",
              padding: "12px 24px",
              backgroundColor: isSaving ? "#ccc" : "#4CAF50",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: isSaving ? "not-allowed" : "pointer",
              fontWeight: "bold",
              fontSize: "16px",
            }}
          >
            {isSaving ? "Menyimpan..." : "💾 Simpan Siswa Baru"}
          </button>
        </div>
      )}

      <div
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 300px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  width: "40px",
                  backgroundColor: "#f4f4f4",
                }}
              >
                <input
                  type="checkbox"
                  checked={
                    selectedRows.size === actualData.length &&
                    actualData.length > 0
                  }
                  onChange={(e) => {
                    if (e.target.checked) {
                      setSelectedRows(new Set(actualData.map((_, i) => i)));
                    } else {
                      setSelectedRows(new Set());
                    }
                  }}
                  style={{ cursor: "pointer", width: "16px", height: "16px" }}
                />
              </th>
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "50px",
                }}
              >
                No
              </th>
              {displayHeaders.map((header, index) => (
                <th
                  key={index}
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "2px solid #ddd",
                    fontWeight: "bold",
                    minWidth: "150px",
                    backgroundColor: "#f4f4f4",
                    fontSize: "12px",
                  }}
                >
                  {header}
                </th>
              ))}
              <th
                style={{
                  padding: "8px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "100px",
                  backgroundColor: "#f4f4f4",
                  fontSize: "12px",
                }}
              >
                Aksi
              </th>
            </tr>
          </thead>
          <tbody>
            {actualData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                  }}
                >
                  <input
                    type="checkbox"
                    checked={selectedRows.has(rowIndex)}
                    onChange={(e) => {
                      const updated = new Set(selectedRows);
                      if (e.target.checked) {
                        updated.add(rowIndex);
                      } else {
                        updated.delete(rowIndex);
                      }
                      setSelectedRows(updated);
                    }}
                    style={{ cursor: "pointer", width: "16px", height: "16px" }}
                  />
                </td>
                <td
                  style={{
                    padding: "8px",
                    textAlign: "center",
                    borderBottom: "1px solid #eee",
                    fontWeight: "bold",
                    color: "#666",
                  }}
                >
                  {rowIndex + 1}
                </td>
                {headers.map((header, colIndex) => (
                  <td
                    key={colIndex}
                    style={{
                      padding: "8px",
                      borderBottom: "1px solid #eee",
                    }}
                  >
                    {header === "Data2" ? (
                      isEditing ? (
                        <select
                          value={row[header] || ""}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          style={{
                            width: "100%",
                            padding: "6px",
                            border: "1px solid #ddd",
                            borderRadius: "3px",
                            fontSize: "12px",
                            backgroundColor: "white",
                            cursor: "pointer",
                          }}
                        >
                          <option value="">-- Pilih Kelas --</option>
                          {kelasOptions.map((k) => (
                            <option key={k} value={k}>
                              {k}
                            </option>
                          ))}
                        </select>
                      ) : (
                        <input
                          type="text"
                          value={row[header] || ""}
                          onChange={(e) =>
                            handleInputChange(rowIndex, header, e.target.value)
                          }
                          disabled={true}
                          style={{
                            width: "100%",
                            padding: "6px",
                            border: "1px solid #ddd",
                            borderRadius: "3px",
                            fontSize: "12px",
                            backgroundColor: "#f5f5f5",
                            cursor: "not-allowed",
                            color: "#666",
                          }}
                        />
                      )
                    ) : (
                      <input
                        type="text"
                        value={row[header] || ""}
                        onChange={(e) =>
                          handleInputChange(rowIndex, header, e.target.value)
                        }
                        disabled={!isEditing}
                        style={{
                          width: "100%",
                          padding: "6px",
                          border: "1px solid #ddd",
                          borderRadius: "3px",
                          fontSize: "12px",
                          backgroundColor: isEditing ? "white" : "#f5f5f5",
                          cursor: isEditing ? "text" : "not-allowed",
                          color: isEditing ? "#000" : "#666",
                        }}
                      />
                    )}
                  </td>
                ))}
                <td
                  style={{
                    padding: "8px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                  }}
                >
                  <label
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: "5px",
                      fontSize: "11px",
                      color: "#e65100",
                      marginBottom: "6px",
                      cursor: "pointer",
                      whiteSpace: "nowrap",
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={deleteNilai[rowIndex] || false}
                      onChange={(e) =>
                        setDeleteNilai((prev) => ({
                          ...prev,
                          [rowIndex]: e.target.checked,
                        }))
                      }
                      style={{
                        cursor: "pointer",
                        width: "14px",
                        height: "14px",
                      }}
                    />
                    Hapus nilai
                  </label>
                  <button
                    onClick={() => handleDelete(rowIndex)}
                    disabled={isSaving}
                    style={{
                      padding: "6px 12px",
                      backgroundColor: isSaving ? "#ccc" : "#f44336",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: isSaving ? "not-allowed" : "pointer",
                      fontSize: "12px",
                      fontWeight: "bold",
                    }}
                  >
                    🗑️ Hapus
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {importSiswaPreview && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0,0,0,0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "24px",
              maxWidth: "600px",
              width: "90%",
              maxHeight: "80vh",
              overflowY: "auto",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
          >
            <h2 style={{ margin: "0 0 16px", color: "#333", fontSize: "18px" }}>
              📋 Preview Data Import Siswa
            </h2>

            <div
              style={{
                padding: "10px 12px",
                backgroundColor: "#e3f2fd",
                borderRadius: "6px",
                border: "1px solid #2196F3",
                marginBottom: "12px",
              }}
            >
              <p style={{ margin: 0, color: "#1565c0" }}>
                <strong>{importSiswaPreview.toImport.length} siswa</strong> akan
                diimport
                {importSiswaPreview.skipped.length > 0 && (
                  <span style={{ color: "#e65100" }}>
                    {" "}
                    · <strong>{importSiswaPreview.skipped.length}</strong>{" "}
                    dilewati (sudah ada)
                  </span>
                )}
              </p>
            </div>

            {/* Tabel preview */}
            <div style={{ overflowX: "auto", marginBottom: "16px" }}>
              <table
                style={{
                  width: "100%",
                  borderCollapse: "collapse",
                  fontSize: "13px",
                }}
              >
                <thead>
                  <tr style={{ backgroundColor: "#f4f4f4" }}>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #ddd",
                        textAlign: "center",
                      }}
                    >
                      No
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #ddd",
                        textAlign: "left",
                      }}
                    >
                      Nama Siswa
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #ddd",
                        textAlign: "center",
                      }}
                    >
                      Kelas
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #ddd",
                        textAlign: "center",
                      }}
                    >
                      NIS
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #ddd",
                        textAlign: "center",
                      }}
                    >
                      NISN
                    </th>
                    <th
                      style={{
                        padding: "8px",
                        border: "1px solid #ddd",
                        textAlign: "left",
                      }}
                    >
                      Nama Ortu
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {importSiswaPreview.toImport.map((s, i) => (
                    <tr
                      key={i}
                      style={{
                        backgroundColor: i % 2 === 0 ? "#fff" : "#f9f9f9",
                      }}
                    >
                      <td
                        style={{
                          padding: "6px 8px",
                          border: "1px solid #eee",
                          textAlign: "center",
                          color: "#666",
                        }}
                      >
                        {i + 1}
                      </td>
                      <td
                        style={{ padding: "6px 8px", border: "1px solid #eee" }}
                      >
                        {s.nama}
                      </td>
                      <td
                        style={{
                          padding: "6px 8px",
                          border: "1px solid #eee",
                          textAlign: "center",
                        }}
                      >
                        {s.kelas}
                      </td>
                      <td
                        style={{
                          padding: "6px 8px",
                          border: "1px solid #eee",
                          textAlign: "center",
                        }}
                      >
                        {s.nis || "-"}
                      </td>
                      <td
                        style={{
                          padding: "6px 8px",
                          border: "1px solid #eee",
                          textAlign: "center",
                        }}
                      >
                        {s.nisn || "-"}
                      </td>
                      <td
                        style={{ padding: "6px 8px", border: "1px solid #eee" }}
                      >
                        {s.namaOrtu || "-"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Daftar yang dilewati */}
            {importSiswaPreview.skipped.length > 0 && (
              <div
                style={{
                  padding: "10px 12px",
                  backgroundColor: "#fff3e0",
                  borderRadius: "6px",
                  border: "1px solid #ff9800",
                  marginBottom: "16px",
                }}
              >
                <p
                  style={{
                    margin: "0 0 6px",
                    color: "#e65100",
                    fontWeight: "bold",
                    fontSize: "13px",
                  }}
                >
                  ⚠️ Dilewati karena sudah ada:
                </p>
                <p style={{ margin: 0, fontSize: "12px", color: "#555" }}>
                  {importSiswaPreview.skipped.join(", ")}
                </p>
              </div>
            )}

            {/* Tombol aksi */}
            <div style={{ display: "flex", gap: "10px" }}>
              <button
                onClick={handleConfirmImport}
                disabled={isImporting}
                style={{
                  flex: 1,
                  padding: "12px",
                  backgroundColor: isImporting ? "#ccc" : "#4CAF50",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: isImporting ? "not-allowed" : "pointer",
                  fontWeight: "bold",
                  fontSize: "14px",
                }}
              >
                {isImporting
                  ? "⏳ Mengirim..."
                  : `✅ Kirim ${importSiswaPreview.toImport.length} Data`}
              </button>
              <button
                onClick={() => setImportSiswaPreview(null)}
                disabled={isImporting}
                style={{
                  flex: 1,
                  padding: "12px",
                  backgroundColor: "#f44336",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: isImporting ? "not-allowed" : "pointer",
                  fontWeight: "bold",
                  fontSize: "14px",
                }}
              >
                ❌ Batalkan
              </button>
            </div>
          </div>
        </div>
      )}

      {importSiswaResult && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0,0,0,0.5)",
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            zIndex: 1000,
          }}
          onClick={() => setImportSiswaResult(null)}
        >
          <div
            style={{
              backgroundColor: "white",
              borderRadius: "8px",
              padding: "24px",
              maxWidth: "480px",
              width: "90%",
              maxHeight: "80vh",
              overflowY: "auto",
              boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <h2 style={{ margin: "0 0 16px", color: "#333", fontSize: "18px" }}>
              ✅ Hasil Import Siswa
            </h2>
            <div
              style={{
                padding: "12px",
                backgroundColor: "#e8f5e9",
                borderRadius: "6px",
                border: "1px solid #4CAF50",
                marginBottom: "12px",
              }}
            >
              <p style={{ margin: 0, color: "#2e7d32", fontWeight: "bold" }}>
                ✅ {importSiswaResult.added} siswa berhasil diimport
              </p>
            </div>
            {importSiswaResult.skipped.length > 0 && (
              <div
                style={{
                  padding: "12px",
                  backgroundColor: "#fff3e0",
                  borderRadius: "6px",
                  border: "1px solid #ff9800",
                }}
              >
                <p
                  style={{
                    margin: "0 0 8px",
                    color: "#e65100",
                    fontWeight: "bold",
                  }}
                >
                  ⚠️ {importSiswaResult.skipped.length} siswa dilewati (sudah
                  ada):
                </p>
                <ul
                  style={{
                    margin: 0,
                    paddingLeft: "20px",
                    fontSize: "13px",
                    color: "#555",
                  }}
                >
                  {importSiswaResult.skipped.map((nama, i) => (
                    <li key={i}>{nama}</li>
                  ))}
                </ul>
              </div>
            )}
            <button
              onClick={() => setImportSiswaResult(null)}
              style={{
                marginTop: "16px",
                padding: "10px 20px",
                backgroundColor: "#4CAF50",
                color: "white",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer",
                fontWeight: "bold",
                width: "100%",
              }}
            >
              Tutup
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

const getFase = (kelas: string): string => {
  const kelasStr = String(kelas || "").replace(/[^0-9]/g, "");
  const kelasNum = parseInt(kelasStr);
  if (kelasNum === 1 || kelasNum === 2) return "A";
  if (kelasNum === 3 || kelasNum === 4) return "B";
  if (kelasNum === 5 || kelasNum === 6) return "C";
  return "-";
};

const PreviewRaporModal: React.FC<{
  rowData: any;
  selectedSemester: string;
  availableSheets: any[];
  schoolData: any;
  onClose: () => void;
}> = ({ rowData, selectedSemester, availableSheets, schoolData, onClose }) => {
  const [loadingPreview, setLoadingPreview] = React.useState(true);
  const [previewData, setPreviewData] = React.useState<any>(null);

  const namaSiswa = rowData?.Data1 || "-";
  const [kelasFromSekolah, setKelasFromSekolah] = React.useState<string>("-");

  React.useEffect(() => {
    const fetchKelasFromSekolah = async () => {
      try {
        const cached = await idbLoad("sekolahData");
        if (cached?.kelas) {
          const kelasVal = String(cached.kelas).trim();
          const rombelVal = String(cached.rombel || "").trim();
          setKelasFromSekolah(
            rombelVal && rombelVal !== "-"
              ? `${kelasVal}${rombelVal}`
              : kelasVal
          );
        }
      } catch (e) {
        console.warn("Gagal fetch kelas dari IndexedDB sekolah:", e);
      }
    };
    fetchKelasFromSekolah();
  }, []);

  const kelas =
    kelasFromSekolah !== "-" ? kelasFromSekolah : rowData?.Data2 || "-";
  const namaOrtu = rowData?.Data5 || "-";
  const [nisn, setNisn] = React.useState("-");

  React.useEffect(() => {
    const fetchNisn = async () => {
      try {
        const siswaRes = await fetch(`${endpoint}?sheet=DataSiswa`);
        if (siswaRes.ok) {
          const siswaJson = await siswaRes.json();
          const siswaRow = siswaJson
            .slice(1)
            .find((r: any) => r.Data1 === namaSiswa);
          if (siswaRow?.Data4) {
            setNisn(
              String(siswaRow.Data4)
                .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
                .trim()
            );
          }
        }
      } catch (e) {
        console.warn("Gagal fetch NISN:", e);
      }
    };
    fetchNisn();
  }, [namaSiswa]);

  const cleanText = (t: string) =>
    t
      ? t
          .replace(/\s+/g, " ")
          .replace(/\s([.,;:!?])/g, "$1")
          .trim()
      : "";

  const generateCatatan = (ranking: number | null) => {
    if (!ranking || isNaN(ranking)) return "Terus semangat belajar ananda!";
    if (ranking === 1) return "Pertahankan prestasi ananda!";
    if (ranking <= 5)
      return "Sudah baik, namun tingkatkan lagi prestasi ananda!";
    if (ranking <= 10) return "Fokus, rajin dan lebih semangat lagi!";
    if (ranking <= 15) return "Tingkatkan semangat ananda sewaktu belajar!";
    return "Lebih rajin lagi mengulang pelajaran di rumah ya!";
  };

  const getNilaiColor = (nilai: string) => {
    const n = parseFloat(nilai);
    if (isNaN(n)) return { bg: "#f5f5f5", color: "#999", border: "#ddd" };
    if (n >= 85) return { bg: "#e8f5e9", color: "#2e7d32", border: "#a5d6a7" };
    if (n >= 70) return { bg: "#fff3e0", color: "#e65100", border: "#ffcc80" };
    return { bg: "#ffebee", color: "#c62828", border: "#ef9a9a" };
  };

  React.useEffect(() => {
    const load = async () => {
      setLoadingPreview(true);
      try {
        const semSheets = availableSheets.filter(
          (s: any) => s.semester === selectedSemester
        );

        const deskripsiResults = await Promise.all(
          semSheets.map(async (sheet: any, index: number) => {
            try {
              const res = await fetch(`${endpoint}?sheet=${sheet.sheetName}`);
              if (!res.ok)
                return {
                  mapel: sheet.mapel,
                  nilai: "-",
                  descMin: "",
                  descMax: "",
                };
              const json = await res.json();
              const siswa = json
                .slice(1)
                .find((r: any) => r.Data4 === namaSiswa);
              return {
                mapel: sheet.mapel,
                nilai: rowData[`Data${6 + index}`] || "-",
                descMin: siswa?.Data26 ? cleanText(siswa.Data26) : "",
                descMax: siswa?.Data27 ? cleanText(siswa.Data27) : "",
              };
            } catch {
              return {
                mapel: sheet.mapel,
                nilai: "-",
                descMin: "",
                descMax: "",
              };
            }
          })
        );

        const [kehadiranRaw, kokurikulerRaw, ekstraRaw] = await Promise.all([
          fetch(`${endpoint}?sheet=DataKehadiran${selectedSemester}`).then(
            (r) => (r.ok ? r.json() : null)
          ),
          fetch(`${endpoint}?sheet=DataKokurikuler${selectedSemester}`).then(
            (r) => (r.ok ? r.json() : null)
          ),
          fetch(
            `${endpoint}?sheet=DataEkstrakurikuler${selectedSemester}`
          ).then((r) => (r.ok ? r.json() : null)),
        ]);

        const kehadiran =
          kehadiranRaw?.slice(1).find((k: any) => k.Data1 === namaSiswa) ||
          null;

        let kokurikuler = "-";
        if (kokurikulerRaw) {
          const kd = kokurikulerRaw
            .slice(1)
            .find((k: any) => k.Data1 === namaSiswa);
          if (kd?.Data10) kokurikuler = kd.Data10;
        }

        const ekstrakurikuler: { nama: string; ket: string }[] = [];
        if (ekstraRaw) {
          const ed = ekstraRaw.slice(1).find((k: any) => k.Data1 === namaSiswa);
          if (ed) {
            if (ed.Data2)
              ekstrakurikuler.push({ nama: ed.Data2, ket: ed.Data3 || "-" });
            if (ed.Data4)
              ekstrakurikuler.push({ nama: ed.Data4, ket: ed.Data5 || "-" });
            if (ed.Data6) ekstrakurikuler.push({ nama: ed.Data6, ket: "-" });
          }
        }
        if (ekstrakurikuler.length === 0)
          ekstrakurikuler.push({ nama: "-", ket: "-" });

        const ranking = rowData?.Data17 ? parseInt(rowData.Data17) : null;

        setPreviewData({
          nilaiMapel: deskripsiResults,
          kehadiran,
          kokurikuler,
          ekstrakurikuler,
          catatan: generateCatatan(ranking),
        });
      } catch (err) {
        console.error(err);
      } finally {
        setLoadingPreview(false);
      }
    };
    load();
  }, []);

  const SectionTitle = ({ children }: { children: string }) => (
    <div
      style={{
        fontWeight: "bold",
        fontSize: "10pt",
        background: "#1565c0",
        color: "#fff",
        padding: "5px 10px",
        borderRadius: "2px",
        margin: "14px 0 6px",
        letterSpacing: "0.3px",
      }}
    >
      {children}
    </div>
  );

  const TH = ({
    children,
    style,
  }: {
    children: React.ReactNode;
    style?: React.CSSProperties;
  }) => (
    <th
      style={{
        background: "#e3eaf5",
        border: "1px solid #b0bec5",
        padding: "5px 8px",
        fontSize: "9pt",
        fontWeight: "bold",
        textAlign: "center",
        ...style,
      }}
    >
      {children}
    </th>
  );

  const TD = ({
    children,
    style,
  }: {
    children: React.ReactNode;
    style?: React.CSSProperties;
  }) => (
    <td
      style={{
        border: "1px solid #cfd8dc",
        padding: "5px 8px",
        fontSize: "9pt",
        verticalAlign: "top",
        ...style,
      }}
    >
      {children}
    </td>
  );

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        backgroundColor: "rgba(0,0,0,0.72)",
        zIndex: 2000,
        overflowY: "auto",
        padding: "8px 4px 60px",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
      }}
    >
      {/* Toolbar */}
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          width: "100%",
          maxWidth: "820px",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: "8px",
          background: "#1e1e2e",
          borderRadius: "8px",
          padding: "8px 12px",
          boxSizing: "border-box",
        }}
      >
        <span style={{ color: "#fff", fontWeight: "bold", fontSize: "14px" }}>
          👁️ Preview Rapor — {namaSiswa}
        </span>
        <button
          onClick={onClose}
          style={{
            background: "#e53935",
            color: "#fff",
            border: "none",
            borderRadius: "6px",
            padding: "8px 20px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "13px",
          }}
        >
          ✕ Tutup
        </button>
      </div>

      {/* Kertas */}
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: "#fff",
          width: "100%",
          maxWidth: "820px",
          boxShadow: "0 8px 40px rgba(0,0,0,0.45)",
          borderRadius: "4px",
          padding: "clamp(12px, 4vw, 40px) clamp(12px, 4vw, 40px) 44px",
          fontFamily: "'Times New Roman', Times, serif",
          fontSize: "clamp(8pt, 2.5vw, 10pt)",
          color: "#111",
          boxSizing: "border-box",
        }}
      >
        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: "8px" }}>
          <div
            style={{
              fontWeight: "bold",
              fontSize: "13pt",
              textTransform: "uppercase",
              letterSpacing: "1px",
            }}
          >
            LAPORAN HASIL BELAJAR (RAPOR)
          </div>
          {schoolData?.namaSekolah && (
            <div style={{ fontSize: "9pt", color: "#555", marginTop: "3px" }}>
              {schoolData.namaSekolah}
            </div>
          )}
          {(schoolData?.alamatSekolah || schoolData?.kabKota) && (
            <div style={{ fontSize: "8.5pt", color: "#777" }}>
              {[schoolData?.alamatSekolah, schoolData?.kabKota]
                .filter(Boolean)
                .join(", ")}
            </div>
          )}
        </div>
        <hr
          style={{
            border: "none",
            borderTop: "2.5px solid #111",
            margin: "8px 0 14px",
          }}
        />

        {/* Info Siswa */}
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: "2px 12px",
            marginBottom: "16px",
            fontSize: "clamp(7pt, 2vw, 8.5pt)",
          }}
        >
          {[
            ["Nama Peserta Didik", namaSiswa.toUpperCase()],
            ["Kelas", kelas],
            ["NISN", nisn],
            ["Fase", getFase(kelas)],
            ["Nama Sekolah", schoolData?.namaSekolah || "-"],
            ["Semester", selectedSemester],
            ["Nama Orang Tua", namaOrtu],
            ["Tahun Pelajaran", schoolData?.tahunPelajaran || "-"],
          ].map(([label, value], i) => (
            <div
              key={i}
              style={{ display: "flex", gap: "3px", padding: "1px 0" }}
            >
              <span style={{ minWidth: "90px", flexShrink: 0, color: "#444" }}>
                {label}
              </span>
              <span style={{ color: "#444" }}>:</span>
              <span style={{ fontWeight: "bold", wordBreak: "break-word" }}>
                {value}
              </span>
            </div>
          ))}
        </div>

        {/* Loading */}
        {loadingPreview ? (
          <div
            style={{
              textAlign: "center",
              padding: "60px 0",
              color: "#888",
              fontSize: "13pt",
            }}
          >
            ⏳ Memuat data rapor...
          </div>
        ) : previewData ? (
          <>
            {/* A. Nilai Mata Pelajaran */}
            <SectionTitle>A. Nilai Mata Pelajaran</SectionTitle>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: "9pt",
                marginBottom: "4px",
              }}
            >
              <thead>
                <tr>
                  <TH style={{ width: "30px" }}>NO.</TH>
                  <TH style={{ textAlign: "left", width: "140px" }}>
                    Mata Pelajaran
                  </TH>
                  <TH style={{ width: "72px" }}>Nilai Akhir</TH>
                  <TH style={{ textAlign: "left" }}>Capaian Kompetensi</TH>
                </tr>
              </thead>
              <tbody>
                {previewData.nilaiMapel.map((item: any, i: number) => {
                  const nc = getNilaiColor(item.nilai);
                  const capaian =
                    [item.descMax, item.descMin].filter(Boolean).join("\n\n") ||
                    "-";
                  return (
                    <tr
                      key={i}
                      style={{ background: i % 2 === 0 ? "#fff" : "#f9f9f9" }}
                    >
                      <TD style={{ textAlign: "center", color: "#666" }}>
                        {i + 1}
                      </TD>
                      <TD style={{ fontWeight: "500" }}>{item.mapel}</TD>
                      <TD style={{ textAlign: "center" }}>
                        <span
                          style={{
                            display: "inline-block",
                            background: nc.bg,
                            color: nc.color,
                            border: `1px solid ${nc.border}`,
                            borderRadius: "4px",
                            padding: "2px 8px",
                            fontWeight: "bold",
                            fontSize: "10.5pt",
                            minWidth: "38px",
                            textAlign: "center",
                          }}
                        >
                          {item.nilai}
                        </span>
                      </TD>
                      <TD
                        style={{
                          lineHeight: "1.55",
                          whiteSpace: "pre-line",
                          fontSize: "8.5pt",
                          color: "#333",
                        }}
                      >
                        {capaian}
                      </TD>
                    </tr>
                  );
                })}
              </tbody>
            </table>

            {/* B. Kokurikuler */}
            <SectionTitle>B. Kokurikuler</SectionTitle>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: "9pt",
              }}
            >
              <thead>
                <tr>
                  <TH style={{ textAlign: "left" }}>Deskripsi Kokurikuler</TH>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <TD style={{ lineHeight: "1.6", whiteSpace: "pre-line" }}>
                    {previewData.kokurikuler}
                  </TD>
                </tr>
              </tbody>
            </table>

            {/* C. Ekstrakurikuler */}
            <SectionTitle>C. Ekstrakurikuler</SectionTitle>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: "9pt",
              }}
            >
              <thead>
                <tr>
                  <TH style={{ textAlign: "left", width: "50%" }}>
                    Ekstrakurikuler
                  </TH>
                  <TH style={{ textAlign: "left" }}>Keterangan</TH>
                </tr>
              </thead>
              <tbody>
                {previewData.ekstrakurikuler.map((e: any, i: number) => (
                  <tr
                    key={i}
                    style={{ background: i % 2 === 0 ? "#fff" : "#f9f9f9" }}
                  >
                    <TD>{e.nama}</TD>
                    <TD>{e.ket}</TD>
                  </tr>
                ))}
              </tbody>
            </table>

            {/* D. Ketidakhadiran */}
            {previewData.kehadiran && (
              <>
                <SectionTitle>D. Ketidakhadiran</SectionTitle>
                <table
                  style={{
                    width: "50%",
                    borderCollapse: "collapse",
                    fontSize: "9pt",
                  }}
                >
                  <thead>
                    <tr>
                      <TH style={{ textAlign: "left" }}>Keterangan</TH>
                      <TH style={{ width: "100px" }}>Jumlah Hari</TH>
                    </tr>
                  </thead>
                  <tbody>
                    {[
                      ["Sakit", previewData.kehadiran.Data7],
                      ["Izin", previewData.kehadiran.Data6],
                      ["Tanpa Keterangan", previewData.kehadiran.Data5],
                    ].map(([label, val], i) => (
                      <tr
                        key={i}
                        style={{ background: i % 2 === 0 ? "#fff" : "#f9f9f9" }}
                      >
                        <TD>{label}</TD>
                        <TD style={{ textAlign: "center", fontWeight: "bold" }}>
                          {val || "0"} hari
                        </TD>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}

            {/* E. Catatan Guru */}
            <SectionTitle>E. Catatan Guru</SectionTitle>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: "9pt",
              }}
            >
              <thead>
                <tr>
                  <TH style={{ textAlign: "left" }}>Catatan</TH>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <TD style={{ lineHeight: "1.6", fontStyle: "italic" }}>
                    {previewData.catatan}
                  </TD>
                </tr>
              </tbody>
            </table>

            {/* Tanda Tangan */}
            <div style={{ marginTop: "28px", fontSize: "9pt" }}>
              {/* Baris 1: Ortu (kiri) dan Guru (kanan) */}
              <div style={{ display: "flex", justifyContent: "space-between" }}>
                {/* Ortu */}
                <div style={{ textAlign: "center", minWidth: "150px" }}>
                  <div>Mengetahui :</div>
                  <div>Orang Tua / Wali,</div>
                  <div style={{ height: "54px" }} />
                  <div
                    style={{
                      borderBottom: "1px solid #333",
                      marginBottom: "2px",
                    }}
                  >
                    {namaOrtu}
                  </div>
                </div>

                {/* Guru */}
                <div style={{ textAlign: "center", minWidth: "150px" }}>
                  <div>
                    {schoolData?.tanggalRapor
                      ? `Bungeng, ${schoolData.tanggalRapor}`
                      : ""}
                  </div>
                  <div>Wali Kelas,</div>
                  {schoolData?.ttdGuru ? (
                    <img
                      src={schoolData.ttdGuru}
                      alt="TTD Guru"
                      style={{
                        height: "54px",
                        maxWidth: "120px",
                        objectFit: "contain",
                      }}
                    />
                  ) : (
                    <div style={{ height: "54px" }} />
                  )}
                  <div
                    style={{
                      fontWeight: "bold",
                      borderBottom: "1px solid #333",
                      marginBottom: "2px",
                    }}
                  >
                    {schoolData?.namaGuru || "_______________"}
                  </div>
                  <div style={{ color: "#555" }}>
                    NIP. {schoolData?.nipGuru || "_______________"}
                  </div>
                </div>
              </div>

              {/* Baris 2: Kepsek di tengah, lebih ke bawah */}
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  marginTop: "32px",
                }}
              >
                <div style={{ textAlign: "center", minWidth: "150px" }}>
                  <div>Mengetahui,</div>
                  <div>Kepala Sekolah</div>
                  {schoolData?.ttdKepsek ? (
                    <img
                      src={schoolData.ttdKepsek}
                      alt="TTD Kepsek"
                      style={{
                        height: "54px",
                        maxWidth: "120px",
                        objectFit: "contain",
                      }}
                    />
                  ) : (
                    <div style={{ height: "54px" }} />
                  )}
                  <div
                    style={{
                      fontWeight: "bold",
                      borderBottom: "1px solid #333",
                      marginBottom: "2px",
                    }}
                  >
                    {schoolData?.namaKepsek || "_______________"}
                  </div>
                  <div style={{ color: "#555" }}>
                    NIP. {schoolData?.nipKepsek || "_______________"}
                  </div>
                </div>
              </div>
            </div>
          </>
        ) : (
          <div
            style={{ textAlign: "center", padding: "40px 0", color: "#e53935" }}
          >
            ❌ Gagal memuat data. Coba tutup dan buka lagi.
          </div>
        )}
      </div>
    </div>
  );
};

const RekapNilai = () => {
  const [data, setData] = useState<RowData[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedKelas, setSelectedKelas] = useState<string>("ALL");
  const [selectedSemester, setSelectedSemester] = useState<string>("1"); // ✅ TAMBAH STATE INI
  const [availableKelas, setAvailableKelas] = useState<string[]>([]);
  const [downloadingId, setDownloadingId] = useState<string | null>(null);
  const [downloadingSampulId, setDownloadingSampulId] = useState<string | null>(
    null
  );
  const [previewSiswa, setPreviewSiswa] = useState<any>(null);
  const [isDownloadingAll, setIsDownloadingAll] = useState(false);
  const [isDownloadingExcel, setIsDownloadingExcel] = useState(false);
  const [localSchoolData, setLocalSchoolData] = useState<SchoolData | null>(
    null
  );
  const [cachedSiswaData, setCachedSiswaData] = useState<any[]>([]);

  const {
    rekapNilaiData,
    rekapNilai2Data, // ✅ TAMBAH INI
    availableSheets,
    schoolData,
    kehadiranData,
    refreshRekapData,
  } = useRekapData();

  // ✅ UBAH useEffect - tambah selectedSemester sebagai dependency
  useEffect(() => {
    console.log(
      `RekapNilai mounted - forcing refresh for semester ${selectedSemester}`
    );
    setLoading(true);

    const refreshData = async () => {
      try {
        if (!localSchoolData) {
          const cachedSekolah = await idbLoad("sekolahData");
          if (cachedSekolah) {
            setLocalSchoolData(cachedSekolah);
          }
        }

        if (cachedSiswaData.length === 0) {
          const cachedSiswa = await idbLoad("siswaData");
          if (cachedSiswa && cachedSiswa.length > 1) {
            setCachedSiswaData(cachedSiswa.slice(1));
          }
        }
        const sheetName =
          selectedSemester === "1" ? "RekapNilai1" : "RekapNilai2"; // ✅ UBAH INI
        const response = await fetch(`${endpoint}?sheet=${sheetName}`);

        if (!response.ok) {
          throw new Error(`Failed to fetch ${sheetName} data`);
        }

        const jsonData = await response.json();

        const actualData = jsonData.slice(1).filter((row: any) => {
          return (
            row.Data1 &&
            typeof row.Data1 === "string" &&
            row.Data1.trim() !== ""
          );
        });

        const filteredJson = [jsonData[0], ...actualData];
        setData(filteredJson);

        const kelasSet = new Set<string>();
        actualData.forEach((row: any) => {
          if (row.Data2) {
            kelasSet.add(row.Data2);
          }
        });
        setAvailableKelas(Array.from(kelasSet).sort());

        setLoading(false);
        console.log(`✅ ${sheetName} data refreshed successfully`);
      } catch (err) {
        console.error("Error refreshing RekapNilai:", err);
        setError(err instanceof Error ? err.message : "Unknown error");
        setLoading(false);
      }
    };

    refreshData();
  }, [selectedSemester]); // ✅ TAMBAH selectedSemester

  const cleanText = (text: string): string => {
    if (!text) return "";
    return text
      .replace(/\s+/g, " ")
      .replace(/\s([.,;:!?])/g, "$1")
      .trim();
  };

  const generateCatatanGuru = (ranking: number | null): string => {
    if (!ranking || isNaN(ranking)) {
      return "Terus semangat belajar ananda!";
    }

    if (ranking === 1) {
      return "Pertahankan prestasi ananda!";
    } else if (ranking >= 2 && ranking <= 5) {
      return "Sudah baik, namun tingkatkan lagi prestasi ananda!";
    } else if (ranking >= 6 && ranking <= 10) {
      return "Fokus, rajin dan lebih semangat lagi!";
    } else if (ranking >= 11 && ranking <= 15) {
      return "Tingkatkan semangat ananda sewaktu belajar!";
    } else if (ranking >= 16) {
      return "Lebih rajin lagi mengulang pelajaran di rumah ya!";
    }

    return "Terus semangat belajar ananda!";
  };

  const downloadRekapExcel = () => {
    if (filteredData.length === 0) return;
    setIsDownloadingExcel(true);

    try {
      const workbook = XLSX.utils.book_new();

      // Kelompokkan per kelas
      const groupedByKelas: { [kelas: string]: any[] } = {};
      filteredData.forEach((row) => {
        const kelas = row.Data2 || "Tanpa Kelas";
        if (!groupedByKelas[kelas]) groupedByKelas[kelas] = [];
        groupedByKelas[kelas].push(row);
      });

      const allRekapHeaders = [
        "Data1",
        "Data6",
        "Data7",
        "Data8",
        "Data9",
        "Data10",
        "Data11",
        "Data12",
        "Data13",
        "Data14",
      ];

      const rekapHeaders = allRekapHeaders.filter((h) => {
        if (h === "Data1") return true;
        const dispH = (data[0]?.[h] || "").trim();
        return (
          dispH !== "" &&
          dispH !== h &&
          !dispH.includes("#REF!") &&
          !dispH.includes("#N/A") &&
          !dispH.toUpperCase().includes("N/A")
        );
      });
      const excludeKeys = ["Data17"];
      const nilaiHeaders = rekapHeaders
        .slice(1)
        .filter((h) => !excludeKeys.includes(h));
      const mapelOnlyHeaders = nilaiHeaders.filter(
        (h) => !["Data15", "Data16"].includes(h)
      );

      const getVals = (rows: any[], h: string) =>
        rows.map((row) => parseFloat(row[h])).filter((v) => !isNaN(v));

      Object.entries(groupedByKelas).forEach(([kelas, rows]) => {
        // ─── Header baris 1: label kolom ───
        const headerRow = [
          "No",
          data[0]?.["Data1"] || "NAMA SISWA",
          ...nilaiHeaders.map((h) => data[0]?.[h] || h),
        ];

        // ─── Data siswa ───
        const dataRows = rows.map((row, idx) => [
          idx + 1,
          ...rekapHeaders.map((h) =>
            row[h] !== undefined && row[h] !== null ? row[h] : ""
          ),
        ]);

        // ─── Baris summary ───
        const buildSummaryRow = (
          label: string,
          mapelFn: (h: string) => string | number,
          data15Val: string | number,
          data16Val: string | number
        ) => {
          return [
            "",
            label,
            ...nilaiHeaders.map((h) => {
              if (h === "Data15") return data15Val;
              if (h === "Data16") return data16Val;
              return mapelFn(h);
            }),
          ];
        };

        // Jumlah
        const jumlahRow = buildSummaryRow(
          "Jumlah",
          (h) => {
            const v = getVals(rows, h);
            return v.length > 0 ? v.reduce((a, b) => a + b, 0) : "";
          },
          (() => {
            const v = rows
              .map((row) => {
                const vals = mapelOnlyHeaders
                  .map((mh) => parseFloat(row[mh]))
                  .filter((v) => !isNaN(v));
                return vals.length > 0 ? vals.reduce((a, b) => a + b, 0) : null;
              })
              .filter((v) => v !== null) as number[];
            return v.length > 0 ? v.reduce((a, b) => a + b, 0) : "";
          })(),
          (() => {
            const totalJumlah = rows
              .map((row) => {
                const v = mapelOnlyHeaders
                  .map((mh) => parseFloat(row[mh]))
                  .filter((v) => !isNaN(v));
                return v.length > 0 ? v.reduce((a, b) => a + b, 0) : null;
              })
              .filter((v) => v !== null) as number[];
            const grand = totalJumlah.reduce((a, b) => a + b, 0);
            return mapelOnlyHeaders.length > 0
              ? parseFloat((grand / mapelOnlyHeaders.length).toFixed(2))
              : "";
          })()
        );

        // Rata-rata Kelas
        const rataRataRow = buildSummaryRow(
          "Rata-rata Kelas",
          (h) => {
            const v = getVals(rows, h);
            return v.length > 0
              ? parseFloat((v.reduce((a, b) => a + b, 0) / v.length).toFixed(2))
              : "";
          },
          (() => {
            const total = mapelOnlyHeaders.reduce((sum, mh) => {
              const v = getVals(rows, mh);
              return (
                sum +
                (v.length > 0 ? v.reduce((a, b) => a + b, 0) / v.length : 0)
              );
            }, 0);
            return mapelOnlyHeaders.length > 0
              ? parseFloat(total.toFixed(2))
              : "";
          })(),
          (() => {
            const total = mapelOnlyHeaders.reduce((sum, mh) => {
              const v = getVals(rows, mh);
              return (
                sum +
                (v.length > 0 ? v.reduce((a, b) => a + b, 0) / v.length : 0)
              );
            }, 0);
            return mapelOnlyHeaders.length > 0
              ? parseFloat((total / mapelOnlyHeaders.length).toFixed(2))
              : "";
          })()
        );

        // Daya Serap
        const dayaSerapRow = buildSummaryRow(
          "Daya Serap",
          (h) => {
            const v = getVals(rows, h);
            return v.length > 0
              ? `${Math.round(v.reduce((a, b) => a + b, 0) / v.length)}%`
              : "";
          },
          (() => {
            const total = mapelOnlyHeaders.reduce((sum, mh) => {
              const v = getVals(rows, mh);
              return (
                sum +
                (v.length > 0 ? v.reduce((a, b) => a + b, 0) / v.length : 0)
              );
            }, 0);
            return mapelOnlyHeaders.length > 0 ? `${Math.round(total)}%` : "";
          })(),
          (() => {
            const total = mapelOnlyHeaders.reduce((sum, mh) => {
              const v = getVals(rows, mh);
              return (
                sum +
                (v.length > 0 ? v.reduce((a, b) => a + b, 0) / v.length : 0)
              );
            }, 0);
            return mapelOnlyHeaders.length > 0
              ? `${Math.round(total / mapelOnlyHeaders.length)}%`
              : "";
          })()
        );

        // Nilai Terbesar
        const nilaiTerbesarRow = buildSummaryRow(
          "Nilai Terbesar",
          (h) => {
            const v = getVals(rows, h);
            return v.length > 0 ? Math.max(...v) : "";
          },
          (() => {
            const perSiswa = rows
              .map((row) => {
                const v = mapelOnlyHeaders
                  .map((mh) => parseFloat(row[mh]))
                  .filter((v) => !isNaN(v));
                return v.length > 0 ? v.reduce((a, b) => a + b, 0) : null;
              })
              .filter((v) => v !== null) as number[];
            return perSiswa.length > 0 ? Math.max(...perSiswa) : "";
          })(),
          (() => {
            const perSiswa = rows
              .map((row) => {
                const v = mapelOnlyHeaders
                  .map((mh) => parseFloat(row[mh]))
                  .filter((v) => !isNaN(v));
                return v.length > 0
                  ? v.reduce((a, b) => a + b, 0) / v.length
                  : null;
              })
              .filter((v) => v !== null) as number[];
            return perSiswa.length > 0
              ? parseFloat(Math.max(...perSiswa).toFixed(2))
              : "";
          })()
        );

        // Nilai Terkecil
        const nilaiTerkecilRow = buildSummaryRow(
          "Nilai Terkecil",
          (h) => {
            const v = getVals(rows, h);
            return v.length > 0 ? Math.min(...v) : "";
          },
          (() => {
            const total = mapelOnlyHeaders.reduce((sum, mh) => {
              const v = getVals(rows, mh);
              return sum + (v.length > 0 ? Math.min(...v) : 0);
            }, 0);
            return mapelOnlyHeaders.length > 0 ? total : "";
          })(),
          (() => {
            const perSiswa = rows
              .map((row) => {
                const v = mapelOnlyHeaders
                  .map((mh) => parseFloat(row[mh]))
                  .filter((v) => !isNaN(v));
                return v.length > 0
                  ? v.reduce((a, b) => a + b, 0) / v.length
                  : null;
              })
              .filter((v) => v !== null) as number[];
            return perSiswa.length > 0
              ? parseFloat(Math.min(...perSiswa).toFixed(2))
              : "";
          })()
        );

        // Gabungkan semua baris (tanpa summary)
        const allRows = [headerRow, ...dataRows];

        // Buat worksheet
        const ws = XLSX.utils.aoa_to_sheet(allRows);

        // Style lebar kolom
        ws["!cols"] = [
          { wch: 5 }, // No
          { wch: 25 }, // Nama
          ...nilaiHeaders.map(() => ({ wch: 12 })),
        ];

        // Nama sheet (maks 31 karakter, karakter invalid diganti)
        const sheetName = `Kelas ${kelas}`
          .substring(0, 31)
          .replace(/[\\/*?:\[\]]/g, "_");
        XLSX.utils.book_append_sheet(workbook, ws, sheetName);
      });

      // Simpan file
      const fileName =
        selectedKelas === "ALL"
          ? `Rekap_Nilai_Semua_Kelas_Sem${selectedSemester}.xlsx`
          : `Rekap_Nilai_Kelas${selectedKelas}_Sem${selectedSemester}.xlsx`;

      const wbout = XLSX.write(workbook, {
        bookType: "xlsx",
        type: "array",
      });

      const blob = new Blob([wbout], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName;
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();

      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 1000);
    } catch (err) {
      alert(
        "❌ Gagal membuat Excel: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsDownloadingExcel(false);
    }
  };

  const compressImageToBase64 = (
    base64: string,
    maxWidth = 200,
    quality = 0.5
  ): Promise<string> => {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        const scale = Math.min(1, maxWidth / img.width);
        canvas.width = img.width * scale;
        canvas.height = img.height * scale;
        const ctx = canvas.getContext("2d")!;
        ctx.fillStyle = "white";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = base64.startsWith("data:")
        ? base64
        : "data:image/png;base64," + base64;
    });
  };

  const downloadAllRekapPDF = async () => {
    if (filteredData.length === 0) return;
    setIsDownloadingAll(true);

    let currentSchoolData = localSchoolData;
    if (!currentSchoolData) {
      try {
        const schoolRes = await fetch(`${endpoint}?action=schoolData`);
        if (schoolRes.ok) {
          const schoolJson = await schoolRes.json();
          if (schoolJson.success && schoolJson.data?.length > 0) {
            currentSchoolData = schoolJson.data[0];
            setLocalSchoolData(schoolJson.data[0]);
          }
        }
      } catch (e) {
        console.warn("Gagal fetch schoolData:", e);
      }
    }

    try {
      const doc = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "a4",
      });
      const pageW = 297;
      const pageH = 210;
      const margin = 10;

      const allHeaders = [
        "Data1",
        "Data6",
        "Data7",
        "Data8",
        "Data9",
        "Data10",
        "Data11",
        "Data12",
        "Data13",
        "Data14",
        "Data15",
        "Data16",
        "Data17",
      ];

      // Filter kolom yang tidak punya nama mapel valid
      const fixedKeepHeaders = new Set(["Data1", "Data15", "Data16", "Data17"]);
      const headers = allHeaders.filter((h) => {
        if (fixedKeepHeaders.has(h)) return true;
        const dispH = (data[0]?.[h] || "").trim();
        return (
          dispH !== "" &&
          dispH !== h && // bukan fallback ke nama key-nya sendiri
          !dispH.includes("#REF!") &&
          !dispH.includes("#N/A") &&
          !dispH.toUpperCase().includes("N/A")
        );
      });

      const displayHdrs = headers.map((h) => data[0]?.[h] || h);

      // Override tampilan header tertentu di PDF
      const headerOverrides: { [key: string]: string } = {
        Data16: "RATA\nRATA", // Pecah manual agar tidak terpotong
      };
      const displayHdrsPDF = displayHdrs.map((h, idx) => {
        const key = headers[idx];
        return headerOverrides[key] ?? h;
      });

      // Hitung lebar kolom
      const noColW = 8;
      const namaColW = 60;

      // Kolom fixed (non-mapel): Data15=Jumlah, Data16=Rata-rata, Data17=Ranking
      const fixedSpecialCols = new Set(["Data15", "Data16", "Data17"]);
      // Hitung lebar fixedSpecial berdasarkan kata terpanjang di headernya
      doc.setFontSize(6);
      const fixedSpecialHeaders = ["Data15", "Data16", "Data17"];
      let fixedSpecialW = 14; // default minimum
      fixedSpecialHeaders.forEach((h) => {
        const dispH = (data[0][h] || "").toUpperCase();
        const words = dispH.trim().split(/\s+/);
        const longestWord = words.reduce(
          (a, b) => (doc.getTextWidth(a) >= doc.getTextWidth(b) ? a : b),
          ""
        );
        const minW = doc.getTextWidth(longestWord) + 4;
        if (minW > fixedSpecialW) fixedSpecialW = minW;
      });
      const fixedSpecialCount = headers
        .slice(1)
        .filter((h) => fixedSpecialCols.has(h)).length;

      // Kolom mapel murni (Data6 - Data14)
      const mapelCols = headers
        .slice(1)
        .filter((h) => !fixedSpecialCols.has(h) && h !== "Data1");

      // Hitung lebar minimum tiap kolom mapel berdasarkan kata terpanjang di headernya
      doc.setFontSize(6);
      const mapelMinWidths: { [h: string]: number } = {};
      mapelCols.forEach((h) => {
        const dispH = (data[0][h] || "").toUpperCase();
        const words = dispH.trim().split(/\s+/);
        const longestWord = words.reduce(
          (a, b) => (doc.getTextWidth(a) >= doc.getTextWidth(b) ? a : b),
          ""
        );
        // Lebar minimum = lebar kata terpanjang + padding
        mapelMinWidths[h] = doc.getTextWidth(longestWord) + 4;
      });

      // Total lebar minimum semua kolom mapel
      const totalMinW = Object.values(mapelMinWidths).reduce(
        (a, b) => a + b,
        0
      );

      // Sisa lebar setelah kolom fixed
      const usedW =
        margin * 2 + noColW + namaColW + fixedSpecialCount * fixedSpecialW;
      const remainingW = pageW - usedW;

      // Distribusikan sisa lebar secara proporsional berdasarkan lebar minimum
      const scale = remainingW / totalMinW;
      const mapelColWidths: { [h: string]: number } = {};
      mapelCols.forEach((h) => {
        mapelColWidths[h] = mapelMinWidths[h] * scale;
      });

      // Fallback otherColW untuk kompatibilitas
      const otherColW = remainingW / (mapelCols.length || 1);

      const columnStyles: { [key: number]: any } = {
        0: { cellWidth: noColW, halign: "center" },
      };
      headers.forEach((h, idx) => {
        if (h === "Data1")
          columnStyles[idx + 1] = { cellWidth: namaColW, halign: "left" };
        else if (fixedSpecialCols.has(h))
          columnStyles[idx + 1] = {
            cellWidth: fixedSpecialW,
            halign: "center",
          };
        else
          columnStyles[idx + 1] = {
            cellWidth: mapelColWidths[h] ?? otherColW,
            halign: "center",
          };
      });

      // Kelompokkan data per kelas jika filter ALL, atau langsung pakai filteredData
      const groupedByKelas: { [kelas: string]: any[] } = {};
      filteredData.forEach((row) => {
        const kelas = row.Data2 || "Tanpa Kelas";
        if (!groupedByKelas[kelas]) groupedByKelas[kelas] = [];
        groupedByKelas[kelas].push(row);
      });

      let isFirstPage = true;

      Object.entries(groupedByKelas).forEach(async ([kelas, rows]) => {
        if (!isFirstPage) {
          doc.addPage();
        }
        isFirstPage = false;

        // Judul
        doc.setFont("helvetica", "bold");
        doc.setFontSize(13);
        doc.text(`REKAP NILAI SISWA - KELAS ${kelas}`, pageW / 2, margin + 5, {
          align: "center",
        });

        doc.setFont("helvetica", "normal");
        doc.setFontSize(9);
        doc.text(
          `Semester: ${selectedSemester}   |   Jumlah Siswa: ${rows.length}`,
          pageW / 2,
          margin + 11,
          { align: "center" }
        );

        // Pecah display header jadi multiline per kata
        const multilineHdrs = displayHdrsPDF.map((h) => {
          if (!h) return h;

          const hKey = headers[displayHdrs.indexOf(h)];
          const colW = fixedSpecialCols.has(hKey)
            ? fixedSpecialW
            : hKey === "Data1"
            ? namaColW
            : mapelColWidths[hKey] ?? otherColW;

          doc.setFontSize(6);

          // Pecah per tanda hubung juga, tidak hanya spasi
          const rawWords = h.trim().toUpperCase().split(/\s+/);
          const words: string[] = [];
          rawWords.forEach((word) => {
            // Jika kata terlalu lebar untuk kolom, pecah per karakter tanda hubung
            if (doc.getTextWidth(word) > colW - 2 && word.includes("-")) {
              const parts = word.split("-");
              parts.forEach((part, i) => {
                words.push(i < parts.length - 1 ? part + "-" : part);
              });
            } else {
              words.push(word);
            }
          });

          const lines: string[] = [];
          let currentLine = "";

          words.forEach((word) => {
            const testLine = currentLine ? `${currentLine} ${word}` : word;
            if (doc.getTextWidth(testLine) <= colW - 2) {
              currentLine = testLine;
            } else {
              if (currentLine) lines.push(currentLine);
              currentLine = word;
            }
          });
          if (currentLine) lines.push(currentLine);

          return lines.join("\n");
        });

        const excludeKeys = ["Data17"];
        const nilaiHeaders = headers
          .slice(1)
          .filter((h) => !excludeKeys.includes(h));
        const mapelOnlyHeaders = nilaiHeaders.filter(
          (h) => !["Data15", "Data16"].includes(h)
        );

        const getVals = (h: string) =>
          rows.map((row) => parseFloat(row[h])).filter((v) => !isNaN(v));

        const calcHoriz = (h: string, type: "sum" | "avg") => {
          const perSiswa = rows
            .map((row) => {
              const v = mapelOnlyHeaders
                .map((mh) => parseFloat(row[mh]))
                .filter((v) => !isNaN(v));
              return v.length > 0
                ? type === "sum"
                  ? v.reduce((a, b) => a + b, 0)
                  : v.reduce((a, b) => a + b, 0) / v.length
                : null;
            })
            .filter((v) => v !== null) as number[];
          return perSiswa;
        };

        const buildSummaryValues = (
          mapelFn: (h: string) => string,
          data15Fn: (vals: number[]) => string,
          data16Fn: (vals: number[]) => string
        ) => {
          return nilaiHeaders.map((h) => {
            if (h === "Data15") return data15Fn(calcHoriz("Data15", "sum"));
            if (h === "Data16") return data16Fn(calcHoriz("Data16", "avg"));
            return mapelFn(h);
          });
        };

        const summaryRows = [
          {
            label: "Jumlah",
            values: buildSummaryValues(
              (h) => {
                const v = getVals(h);
                return v.length > 0 ? String(v.reduce((a, b) => a + b, 0)) : "";
              },
              (v) => (v.length > 0 ? String(v.reduce((a, b) => a + b, 0)) : ""),
              (_v) => {
                const totalJumlah = rows
                  .map((row) => {
                    const v = mapelOnlyHeaders
                      .map((mh) => parseFloat(row[mh]))
                      .filter((v) => !isNaN(v));
                    return v.length > 0 ? v.reduce((a, b) => a + b, 0) : null;
                  })
                  .filter((v) => v !== null) as number[];
                const grandTotal = totalJumlah.reduce((a, b) => a + b, 0);
                return mapelOnlyHeaders.length > 0
                  ? (grandTotal / mapelOnlyHeaders.length).toFixed(2)
                  : "";
              }
            ),
            fill: [227, 242, 253] as [number, number, number],
          },
          {
            label: "Rata-rata Kelas",
            values: buildSummaryValues(
              (h) => {
                const v = getVals(h);
                return v.length > 0
                  ? (v.reduce((a, b) => a + b, 0) / v.length).toFixed(2)
                  : "";
              },
              // Data15: jumlah semua rata-rata per mapel
              (_v) => {
                const totalRataMapel = mapelOnlyHeaders.reduce((sum, mh) => {
                  const vals = rows
                    .map((row) => parseFloat(row[mh]))
                    .filter((v) => !isNaN(v));
                  const avg =
                    vals.length > 0
                      ? vals.reduce((a, b) => a + b, 0) / vals.length
                      : 0;
                  return sum + avg;
                }, 0);
                return mapelOnlyHeaders.length > 0
                  ? totalRataMapel.toFixed(2)
                  : "";
              },
              // Data16: jumlah rata-rata semua mapel dibagi jumlah mapel
              (_v) => {
                const totalRataMapel = mapelOnlyHeaders.reduce((sum, mh) => {
                  const vals = rows
                    .map((row) => parseFloat(row[mh]))
                    .filter((v) => !isNaN(v));
                  const avg =
                    vals.length > 0
                      ? vals.reduce((a, b) => a + b, 0) / vals.length
                      : 0;
                  return sum + avg;
                }, 0);
                return mapelOnlyHeaders.length > 0
                  ? (totalRataMapel / mapelOnlyHeaders.length).toFixed(2)
                  : "";
              }
            ),
            fill: [232, 245, 233] as [number, number, number],
          },
          {
            label: "Daya Serap",
            values: buildSummaryValues(
              (h) => {
                const v = getVals(h);
                if (v.length === 0) return "";
                return `${Math.round(
                  v.reduce((a, b) => a + b, 0) / v.length
                )}%`;
              },
              // Data15: jumlah daya serap per mapel
              (_v) => {
                const totalDS = mapelOnlyHeaders.reduce((sum, mh) => {
                  const vals = rows
                    .map((row) => parseFloat(row[mh]))
                    .filter((v) => !isNaN(v));
                  const avg =
                    vals.length > 0
                      ? vals.reduce((a, b) => a + b, 0) / vals.length
                      : 0;
                  return sum + avg;
                }, 0);
                return mapelOnlyHeaders.length > 0
                  ? `${Math.round(totalDS)}%`
                  : "";
              },
              // Data16: rata-rata daya serap dibagi jumlah mapel
              (_v) => {
                const totalDS = mapelOnlyHeaders.reduce((sum, mh) => {
                  const vals = rows
                    .map((row) => parseFloat(row[mh]))
                    .filter((v) => !isNaN(v));
                  const avg =
                    vals.length > 0
                      ? vals.reduce((a, b) => a + b, 0) / vals.length
                      : 0;
                  return sum + avg;
                }, 0);
                return mapelOnlyHeaders.length > 0
                  ? `${Math.round(totalDS / mapelOnlyHeaders.length)}%`
                  : "";
              }
            ),
            fill: [255, 243, 224] as [number, number, number],
          },
          {
            label: "Nilai Terbesar",
            values: buildSummaryValues(
              (h) => {
                const v = getVals(h);
                return v.length > 0 ? String(Math.max(...v)) : "";
              },
              // Data15: nilai terbesar dari jumlah per siswa
              (_v) => {
                const perSiswaJumlah = rows
                  .map((row) => {
                    const v = mapelOnlyHeaders
                      .map((mh) => parseFloat(row[mh]))
                      .filter((v) => !isNaN(v));
                    return v.length > 0 ? v.reduce((a, b) => a + b, 0) : null;
                  })
                  .filter((v) => v !== null) as number[];
                return perSiswaJumlah.length > 0
                  ? String(Math.max(...perSiswaJumlah))
                  : "";
              },
              // Data16: nilai terbesar dari rata-rata per siswa
              (_v) => {
                const perSiswaRata = rows
                  .map((row) => {
                    const v = mapelOnlyHeaders
                      .map((mh) => parseFloat(row[mh]))
                      .filter((v) => !isNaN(v));
                    return v.length > 0
                      ? v.reduce((a, b) => a + b, 0) / v.length
                      : null;
                  })
                  .filter((v) => v !== null) as number[];
                return perSiswaRata.length > 0
                  ? Math.max(...perSiswaRata).toFixed(2)
                  : "";
              }
            ),
            fill: [252, 228, 236] as [number, number, number],
          },
          {
            label: "Nilai Terkecil",
            values: buildSummaryValues(
              (h) => {
                const v = getVals(h);
                return v.length > 0 ? String(Math.min(...v)) : "";
              },
              // Data15: jumlah semua nilai terkecil per mapel
              (_v) => {
                const totalTerkecil = mapelOnlyHeaders.reduce((sum, mh) => {
                  const vals = rows
                    .map((row) => parseFloat(row[mh]))
                    .filter((v) => !isNaN(v));
                  const min = vals.length > 0 ? Math.min(...vals) : 0;
                  return sum + min;
                }, 0);
                return mapelOnlyHeaders.length > 0 ? String(totalTerkecil) : "";
              },
              // Data16: nilai terkecil dari rata-rata per siswa
              (_v) => {
                const perSiswaRata = rows
                  .map((row) => {
                    const v = mapelOnlyHeaders
                      .map((mh) => parseFloat(row[mh]))
                      .filter((v) => !isNaN(v));
                    return v.length > 0
                      ? v.reduce((a, b) => a + b, 0) / v.length
                      : null;
                  })
                  .filter((v) => v !== null) as number[];
                return perSiswaRata.length > 0
                  ? Math.min(...perSiswaRata).toFixed(2)
                  : "";
              }
            ),
            fill: [243, 229, 245] as [number, number, number],
          },
        ];

        const tableBody = [
          ...rows.map((row, idx) => [
            idx + 1,
            ...headers.map((h) =>
              row[h] !== undefined && row[h] !== null ? String(row[h]) : ""
            ),
          ]),
          // Baris summary
          ...summaryRows.map((s) => ["", s.label, ...s.values]),
        ];

        autoTable(doc, {
          startY: margin + 15,
          head: [["NO", ...multilineHdrs]],
          body: tableBody,
          theme: "grid",
          headStyles: {
            fillColor: [41, 128, 185],
            textColor: 255,
            fontStyle: "bold",
            halign: "center",
            fontSize: 6,
            cellPadding: 2,
            minCellHeight: 12,
            valign: "middle",
            overflow: "linebreak",
          },
          bodyStyles: {
            fontSize: 7,
            cellPadding: 2,
          },
          columnStyles: columnStyles,
          styles: {
            overflow: "ellipsize",
          },
          margin: { left: margin, right: margin },
          didParseCell: (hookData) => {
            if (hookData.section === "body") {
              const totalRows = rows.length;
              const rowIdx = hookData.row.index;

              if (rowIdx >= totalRows) {
                // Baris summary
                const summaryIdx = rowIdx - totalRows;
                const summaryFills: [number, number, number][] = [
                  [227, 242, 253],
                  [232, 245, 233],
                  [255, 243, 224],
                  [252, 228, 236],
                  [243, 229, 245],
                ];
                hookData.cell.styles.fillColor = summaryFills[summaryIdx] || [
                  255, 255, 255,
                ];
                hookData.cell.styles.fontStyle = "bold";
                hookData.cell.styles.textColor = [0, 0, 0];
              } else if (rowIdx % 2 === 0) {
                // Warna baris selang-seling untuk data siswa
                hookData.cell.styles.fillColor = [224, 255, 255];
              }
            }
          },
        });

        // Footer per kelas
        const finalY = doc.lastAutoTable.finalY + 5;
        doc.setFont("helvetica", "italic");
        doc.setFontSize(8);
        doc.setTextColor(100, 100, 100);
        doc.text(`Total: ${rows.length} siswa`, margin, finalY);
        doc.setTextColor(0, 0, 0);

        // ⬇️ TAMBAHKAN TANDA TANGAN DI SINI
        const ttdStartY = finalY + 10;
        const kepsekX = margin;
        const guruX = pageW - margin - 50;

        doc.setFont("helvetica", "normal");
        doc.setFontSize(9);

        // Tanggal
        const tanggalRapor = localSchoolData?.tanggalRapor || "";
        if (tanggalRapor) {
          doc.text(`Bungeng, ${tanggalRapor}`, guruX, ttdStartY);
        }

        // Label jabatan
        doc.text("Kepala Sekolah,", kepsekX, ttdStartY + 5);
        doc.text("Wali Kelas,", guruX, ttdStartY + 5);

        // TTD Kepsek
        if (currentSchoolData?.ttdKepsek) {
          try {
            const compressed = await compressImageToBase64(
              currentSchoolData.ttdKepsek,
              800,
              1.0
            );
            doc.addImage(compressed, "JPEG", kepsekX, ttdStartY + 7, 30, 15);
          } catch (e) {
            console.warn("Gagal load TTD Kepsek:", e);
          }
        }

        // TTD Guru
        if (currentSchoolData?.ttdGuru) {
          try {
            const compressed = await compressImageToBase64(
              currentSchoolData.ttdGuru,
              800,
              1.0
            );
            doc.addImage(compressed, "JPEG", guruX, ttdStartY + 7, 30, 15);
          } catch (e) {
            console.warn("Gagal load TTD Guru:", e);
          }
        }

        // Nama dan NIP Kepsek
        const namaKepsek = localSchoolData?.namaKepsek || "_______________";
        const nipKepsek = localSchoolData?.nipKepsek || "_______________";
        doc.setFont("helvetica", "bold");
        doc.text(namaKepsek, kepsekX, ttdStartY + 25);
        doc.setLineWidth(0.3);
        const kepsekTextW = doc.getTextWidth(namaKepsek);
        doc.line(
          kepsekX,
          ttdStartY + 26,
          kepsekX + kepsekTextW,
          ttdStartY + 26
        );
        doc.setFont("helvetica", "normal");
        doc.text(`NIP. ${nipKepsek}`, kepsekX, ttdStartY + 30);

        // Nama dan NIP Guru
        const namaGuru = localSchoolData?.namaGuru || "_______________";
        const nipGuru = localSchoolData?.nipGuru || "_______________";
        doc.setFont("helvetica", "bold");
        doc.text(namaGuru, guruX, ttdStartY + 25);
        doc.setLineWidth(0.3);
        const guruTextW = doc.getTextWidth(namaGuru);
        doc.line(guruX, ttdStartY + 26, guruX + guruTextW, ttdStartY + 26);
        doc.setFont("helvetica", "normal");
        doc.text(`NIP. ${nipGuru}`, guruX, ttdStartY + 30);
      });

      // Nomor halaman
      const totalPages = doc.getNumberOfPages();
      for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(8);
        doc.setTextColor(100, 100, 100);
        doc.text(`Halaman ${i} dari ${totalPages}`, pageW - margin, pageH - 5, {
          align: "right",
        });
      }

      const fileName =
        selectedKelas === "ALL"
          ? `Rekap_Nilai_Semua_Kelas_Sem${selectedSemester}.pdf`
          : `Rekap_Nilai_Kelas${selectedKelas}_Sem${selectedSemester}.pdf`;

      doc.save(fileName);
    } catch (err) {
      alert(
        "❌ Gagal membuat PDF: " +
          (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setIsDownloadingAll(false);
    }
  };

  const getSchoolData = () => localSchoolData || schoolData;

  const downloadSampulPDF = async (rowData: any) => {
    const namaSiswa = rowData.Data1 || "-";
    const nisn = rowData.Data4 || "-";
    const nis = rowData.Data3 || "-";

    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
    });

    const pageW = 210;
    const pageH = 297;
    const margin = 20;

    // ─── BORDER LUAR (garis tebal) ───
    doc.setDrawColor(0);
    doc.setLineWidth(2);
    doc.rect(margin, margin, pageW - margin * 2, pageH - margin * 2);

    // ─── BORDER DALAM (garis tipis) ───
    doc.setLineWidth(0.5);
    doc.rect(
      margin + 3,
      margin + 3,
      pageW - (margin + 3) * 2,
      pageH - (margin + 3) * 2
    );

    const centerX = pageW / 2;

    // ─── MUAT LOGO TUT WURI ───
    try {
      const response = await fetch("/logo-tutwurih.png");
      const blob = await response.blob();
      const logoBase64: string = await new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as string);
        reader.readAsDataURL(blob);
      });

      const logoSize = 40;
      const logoX = centerX - logoSize / 2;
      const logoY = margin + 12;
      doc.addImage(logoBase64, "PNG", logoX, logoY, logoSize, logoSize);
    } catch (e) {
      console.warn("Gagal load logo Tut Wuri:", e);
    }

    // ─── JUDUL RAPOR (di bawah logo) ───
    let y = margin + 65;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("RAPOR", centerX, y, { align: "center" });

    y += 10;
    doc.setFontSize(18);
    doc.text("P E S E R T A   D I D I K", centerX, y, { align: "center" });

    y += 9;
    doc.setFontSize(16);
    doc.text("SEKOLAH DASAR", centerX, y, { align: "center" });

    y += 9;
    doc.text("(SD)", centerX, y, { align: "center" });

    // ─── NAMA SEKOLAH ───
    y += 8;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.text(getSchoolData()?.namaSekolah || "", centerX, y, {
      align: "center",
    });

    // ─── GARIS PEMISAH ───
    y += 10;
    doc.setLineWidth(0.5);
    doc.line(margin + 10, y, pageW - margin - 10, y);

    // ─── SPACER ───
    y += 25;

    // ─── LABEL "Nama Peserta Didik :" ───
    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.text("Nama Peserta Didik :", centerX, y, { align: "center" });

    // ─── KOTAK NAMA SISWA ───
    y += 7;
    const boxW = 120;
    const boxX = centerX - boxW / 2;
    doc.setLineWidth(0.5);
    doc.rect(boxX, y, boxW, 12);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.text(namaSiswa.toUpperCase(), centerX, y + 8.5, { align: "center" });

    // ─── LABEL "NISN/NIS" ───
    y += 20;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.text("NISN/NIS", centerX, y, { align: "center" });

    // ─── KOTAK NISN/NIS ───
    y += 7;
    doc.setLineWidth(0.5);
    doc.rect(boxX, y, boxW, 12);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(13);
    const nisnClean = String(nisn)
      .replace(/[\u200B\u200C\u200D\uFEFF\s]/g, "")
      .trim();
    const nisClean = String(nis)
      .replace(/[\u200B\u200C\u200D\uFEFF\s]/g, "")
      .trim();
    const nisnNisText = `${nisnClean} / ${nisClean}`;
    doc.text(nisnNisText, centerX, y + 8.5, { align: "center" });

    // ─── GARIS PEMISAH FOOTER ───
    const footerY = pageH - margin - 30;
    doc.setLineWidth(0.5);
    doc.line(margin + 10, footerY, pageW - margin - 10, footerY);

    // ─── FOOTER ───
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.text("K E M E N D I K D A S M E N", centerX, footerY + 10, {
      align: "center",
    });

    doc.setFontSize(12);
    doc.text("REPUBLIK INDONESIA", centerX, footerY + 19, { align: "center" });

    const fileName = `Sampul_${namaSiswa.replace(/\s+/g, "_")}.pdf`;
    doc.save(fileName);
  };

  const downloadRaporPDF = async (rowData: any) => {
    setDownloadingId(rowData.Data1);

    let latestSchoolData = localSchoolData || schoolData;
    if (!latestSchoolData || !latestSchoolData.namaGuru) {
      try {
        const schoolRes = await fetch(`${endpoint}?action=schoolData`);
        if (schoolRes.ok) {
          const schoolJson = await schoolRes.json();
          if (schoolJson.success && schoolJson.data?.length > 0) {
            latestSchoolData = schoolJson.data[0];
            setLocalSchoolData(schoolJson.data[0]); // simpan untuk berikutnya
          }
        }
      } catch (e) {
        console.warn("Gagal fetch schoolData:", e);
      }
    }

    console.log("=== START PDF GENERATION ===");
    console.log("Siswa:", rowData.Data1);
    console.log("Semester:", selectedSemester); // ✅ TAMBAH LOG INI

    console.log("=== ALL ROW DATA ===");
    Object.keys(rowData).forEach((key) => {
      console.log(`${key}: ${rowData[key]}`);
    });

    try {
      const doc = new jsPDF();

      const namaSiswa = rowData.Data1 || "-";
      let kelas = rowData.Data2 || "-";
      try {
        const sekolahCached = await idbLoad("sekolahData");
        if (sekolahCached?.kelas) {
          const kelasVal = String(sekolahCached.kelas).trim();
          const rombelVal = String(sekolahCached.rombel || "").trim();
          kelas =
            rombelVal && rombelVal !== "-"
              ? `${kelasVal}${rombelVal}`
              : kelasVal;
        }
      } catch (e) {
        console.warn("Gagal fetch kelas dari sekolahData:", e);
      }
      const namaOrtu = rowData.Data5 || "-";

      // Ambil NISN dari cache
      let nisn = "-";
      const siswaRow = cachedSiswaData.find((r: any) => r.Data1 === namaSiswa);
      if (siswaRow?.Data4) {
        nisn = String(siswaRow.Data4)
          .replace(/[\u200B\u200C\u200D\uFEFF]/g, "")
          .trim();
      }

      // Header
      doc.setFontSize(16);
      doc.setFont("helvetica", "bold");
      doc.text("LAPORAN HASIL BELAJAR (RAPOR)", 105, 20, { align: "center" });

      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");

      const leftCol = 20;
      const rightCol = 130;
      const leftColTTD = 25;
      const centerColTTD = 100;
      const rightColTTD = 150;
      let y = 35;

      doc.text("Nama Peserta Didik", leftCol, y);
      doc.text(": " + namaSiswa.toUpperCase(), leftCol + 50, y);
      doc.text("Kelas", rightCol, y);
      doc.text(": " + kelas, rightCol + 30, y);

      y += 7;
      doc.text("NISN", leftCol, y);
      doc.text(`: ${nisn}`, leftCol + 50, y);
      doc.text("Fase", rightCol, y);
      doc.text(`: ${getFase(kelas)}`, rightCol + 30, y);

      y += 7;
      doc.text("Nama Sekolah", leftCol, y);
      doc.text(
        ": " + (latestSchoolData?.namaSekolah || "UPT SD NEGERI 2 BATANG"),
        leftCol + 50,
        y
      );
      doc.text("Semester", rightCol, y);
      doc.text(": " + selectedSemester, rightCol + 30, y); // ✅ UBAH INI

      y += 7;
      doc.text("Alamat Sekolah", leftCol, y);
      const alamatLengkap = `${
        latestSchoolData?.alamatSekolah || "Desa Bungeng, Kecamatan Batang"
      }, ${latestSchoolData?.kabKota || ""}`;
      doc.text(": " + alamatLengkap, leftCol + 50, y);
      doc.text("Tahun Pelajaran", rightCol, y);
      doc.text(
        ": " + (latestSchoolData?.tahunPelajaran || "2023/2024"),
        rightCol + 30,
        y
      );

      // ✅ FILTER sheet berdasarkan semester
      const semesterSheets = availableSheets.filter(
        (sheet) => sheet.semester === selectedSemester
      );

      // ✅✅✅ TAMBAHKAN DEBUG INI
      console.log("=== DEBUG PDF GENERATION ===");
      console.log("Selected Semester:", selectedSemester);
      console.log("All Available Sheets:", availableSheets.length);
      console.log("Available Sheets Detail:", availableSheets);
      console.log("Filtered Semester Sheets:", semesterSheets.length);
      console.log("Semester Sheets Detail:", semesterSheets);

      if (semesterSheets.length === 0) {
        console.error(
          "❌ ERROR: No sheets found for semester",
          selectedSemester
        );
        console.error(
          "Available sheets data:",
          JSON.stringify(availableSheets, null, 2)
        );
        alert(
          `⚠️ Tidak ada data mata pelajaran untuk Semester ${selectedSemester}!\n\nSilakan cek sheet MAPEL${selectedSemester}01, MAPEL${selectedSemester}02, dst di Google Sheets.`
        );
        setDownloadingId(null);
        return;
      }

      console.log(
        `📚 Found ${semesterSheets.length} sheets for semester ${selectedSemester}`
      );

      // ✅ Fetch deskripsi hanya untuk semester yang dipilih
      const [
        kokurikulerResult,
        ekstrakurikulerResult,
        kehadiranResult,
        ...deskripsiRaw
      ] = await Promise.all([
        fetch(`${endpoint}?sheet=DataKokurikuler${selectedSemester}`)
          .then((res) => (res.ok ? res.json() : null))
          .catch(() => null),
        fetch(`${endpoint}?sheet=DataEkstrakurikuler${selectedSemester}`)
          .then((res) => (res.ok ? res.json() : null))
          .catch(() => null),
        fetch(`${endpoint}?sheet=DataKehadiran${selectedSemester}`)
          .then((res) => (res.ok ? res.json() : null))
          .catch(() => null),
        ...semesterSheets.map((sheet) =>
          fetch(`${endpoint}?sheet=${sheet.sheetName}`)
            .then((res) => (res.ok ? res.json() : null))
            .catch(() => null)
        ),
      ]);

      const deskripsiData = semesterSheets.map((sheet, index) => {
        const jsonData = deskripsiRaw[index];
        if (!jsonData) return { mapel: sheet.mapel, descMin: "", descMax: "" };
        const siswaData = jsonData
          .slice(1)
          .find((row: any) => row.Data4 === namaSiswa);
        return {
          mapel: sheet.mapel,
          descMin: siswaData?.Data26 || "",
          descMax: siswaData?.Data27 || "",
        };
      });

      // ✅ Ambil nilai dari kolom sesuai semester
      const nilaiMapel: { [key: string]: number | null } = {};
      semesterSheets.forEach((sheet, index) => {
        const dataKey = `Data${6 + index}`;
        const nilaiStr = rowData[dataKey];
        nilaiMapel[sheet.mapel] = nilaiStr ? parseFloat(nilaiStr) : null;
      });

      y += 10;
      const mapelColumns = semesterSheets.map((sheet) => sheet.mapel);
      const tableData = mapelColumns.map((mapel, index) => {
        const nilai = nilaiMapel[mapel];
        const desc = deskripsiData.find((d) => d.mapel === mapel);

        let capaianText = "";
        if (desc?.descMax) {
          capaianText += cleanText(desc.descMax);
        }
        if (desc?.descMin) {
          if (capaianText) capaianText += "\n\n";
          capaianText += cleanText(desc.descMin);
        }
        if (!capaianText) {
          capaianText = "-";
        }

        let nilaiText = "-";
        if (nilai !== null && nilai !== undefined) {
          nilaiText = String(nilai);
        }

        return [index + 1, mapel, nilaiText, capaianText];
      });

      autoTable(doc, {
        startY: y,
        head: [["No.", "Mata Pelajaran", "Nilai Akhir", "Capaian Kompetensi"]],
        body: tableData,
        theme: "grid",
        headStyles: {
          fillColor: [200, 200, 200],
          textColor: 0,
          fontStyle: "bold",
          halign: "center",
        },
        columnStyles: {
          0: { cellWidth: 15, halign: "center" },
          1: { cellWidth: 50 },
          2: { cellWidth: 25, halign: "center" },
          3: {
            cellWidth: 90,
            cellPadding: 3,
            overflow: "linebreak",
            valign: "top",
          },
        },
        styles: {
          fontSize: 9,
          cellPadding: 3,
          overflow: "linebreak",
          cellWidth: "wrap",
        },
        rowPageBreak: "avoid",
        pageBreak: "auto",
      });

      let additionalY = doc.lastAutoTable.finalY + 10;

      let kokurikulerText = "-";
      if (kokurikulerResult) {
        const studentKokurikuler = kokurikulerResult
          .slice(1)
          .find((k: any) => k.Data1 === namaSiswa);

        if (studentKokurikuler && studentKokurikuler.Data10) {
          kokurikulerText = studentKokurikuler.Data10;
        }
      }

      const remainingSpace = 297 - additionalY;
      const estimatedTableHeight = 30;

      if (remainingSpace < estimatedTableHeight + 60) {
        doc.addPage();
        additionalY = 20;
      }

      doc.setFont("helvetica", "bold");
      doc.text("Kokurikuler", leftCol, additionalY);
      doc.setFont("helvetica", "normal");
      additionalY += 7;

      autoTable(doc, {
        startY: additionalY,
        head: [["Deskripsi Kokurikuler"]],
        body: [[kokurikulerText]],
        theme: "grid",
        headStyles: {
          fillColor: [200, 200, 200],
          textColor: 0,
          fontStyle: "bold",
          halign: "center",
        },
        columnStyles: {
          0: { cellWidth: 140, halign: "left" },
        },
        styles: {
          fontSize: 9,
          cellPadding: 5,
        },
        margin: { left: leftCol },
      });

      additionalY = doc.lastAutoTable.finalY + 10;

      let ekstrakurikulerData: any[] = [];
      if (ekstrakurikulerResult) {
        const studentEkstrakurikuler = ekstrakurikulerResult
          .slice(1)
          .find((k: any) => k.Data1 === namaSiswa);

        if (studentEkstrakurikuler) {
          if (studentEkstrakurikuler.Data2 || studentEkstrakurikuler.Data3) {
            ekstrakurikulerData.push([
              studentEkstrakurikuler.Data2 || "-",
              studentEkstrakurikuler.Data3 || "-",
            ]);
          }

          if (studentEkstrakurikuler.Data4 || studentEkstrakurikuler.Data5) {
            ekstrakurikulerData.push([
              studentEkstrakurikuler.Data4 || "-",
              studentEkstrakurikuler.Data5 || "-",
            ]);
          }

          if (studentEkstrakurikuler.Data6) {
            ekstrakurikulerData.push([
              studentEkstrakurikuler.Data6 || "-",
              "-",
            ]);
          }
        }
      }

      if (ekstrakurikulerData.length === 0) {
        ekstrakurikulerData = [["-", "-"]];
      }

      const remainingSpace2 = 297 - additionalY;
      const estimatedTableHeight2 = 20 + ekstrakurikulerData.length * 8;

      if (remainingSpace2 < estimatedTableHeight2 + 60) {
        doc.addPage();
        additionalY = 20;
      }

      doc.setFont("helvetica", "bold");
      doc.text("Ekstrakurikuler", leftCol, additionalY);
      doc.setFont("helvetica", "normal");
      additionalY += 7;

      autoTable(doc, {
        startY: additionalY,
        head: [["Ekstrakurikuler", "Keterangan"]],
        body: ekstrakurikulerData,
        theme: "grid",
        headStyles: {
          fillColor: [200, 200, 200],
          textColor: 0,
          fontStyle: "bold",
          halign: "center",
        },
        columnStyles: {
          0: { cellWidth: 70, halign: "left" },
          1: { cellWidth: 70, halign: "left" },
        },
        styles: {
          fontSize: 9,
          cellPadding: 5,
        },
        margin: { left: leftCol },
      });

      additionalY = doc.lastAutoTable.finalY + 10;

      // Identifikasi kolom ranking (SESUAIKAN dengan kolom ranking Anda!)
      const rankingKey = "Data17"; // ⚠️ UBAH ini sesuai kolom ranking di sheet Anda
      const ranking = rowData[rankingKey]
        ? parseInt(rowData[rankingKey])
        : null;

      // Generate catatan otomatis berdasarkan ranking
      const catatan = generateCatatanGuru(ranking);

      console.log("=== CATATAN GURU DEBUG ===");
      console.log("Ranking Key:", rankingKey);
      console.log("Ranking Value:", ranking);
      console.log("Generated Catatan:", catatan);

      const remainingSpaceBeforeCatatan = 297 - additionalY;
      const estimatedCatatanHeight = 25;

      if (remainingSpaceBeforeCatatan < estimatedCatatanHeight + 60) {
        doc.addPage();
        additionalY = 20;
      }

      doc.setFont("helvetica", "bold");
      doc.text("Catatan Guru", leftCol, additionalY);
      doc.setFont("helvetica", "normal");
      additionalY += 7;

      autoTable(doc, {
        startY: additionalY,
        head: [["Catatan"]],
        body: [[catatan]],
        theme: "grid",
        headStyles: {
          fillColor: [200, 200, 200],
          textColor: 0,
          fontStyle: "bold",
          halign: "center",
        },
        columnStyles: {
          0: { cellWidth: 140, halign: "left" },
        },
        styles: {
          fontSize: 9,
          cellPadding: 5,
        },
        margin: { left: leftCol },
      });

      additionalY = doc.lastAutoTable.finalY + 10;

      const kehadiranList: RowData[] = kehadiranResult
        ? kehadiranResult.slice(1)
        : [];
      const studentKehadiran = kehadiranList.find(
        (k: RowData) => k.Data1 === namaSiswa
      );

      const requiredSpace = 170;
      const remainingSpaceBeforeKehadiran = 297 - additionalY;

      if (remainingSpaceBeforeKehadiran < requiredSpace) {
        doc.addPage();
        additionalY = 20;
      }

      if (studentKehadiran) {
        doc.setFont("helvetica", "bold");
        doc.text("Ketidakhadiran", leftCol, additionalY);
        doc.setFont("helvetica", "normal");
        additionalY += 7;

        const kehadiranStartY = additionalY;

        autoTable(doc, {
          startY: kehadiranStartY,
          head: [["Keterangan", "Jumlah Hari"]],
          body: [
            ["Sakit", `${studentKehadiran.Data7 || "0"} hari`],
            ["Izin", `${studentKehadiran.Data6 || "0"} hari`],
            ["Tanpa Keterangan", `${studentKehadiran.Data5 || "0"} hari`],
          ],
          theme: "grid",
          headStyles: {
            fillColor: [200, 200, 200],
            textColor: 0,
            fontStyle: "bold",
            halign: "center",
          },
          columnStyles: {
            0: { cellWidth: 45, halign: "left" },
            1: { cellWidth: 25, halign: "center" },
          },
          styles: {
            fontSize: 9,
            cellPadding: 3,
          },
          margin: { left: leftCol },
        });

        const kehadiranEndY = doc.lastAutoTable.finalY;

        // ✅ Tabel Keputusan hanya muncul di Semester 2
        if (selectedSemester === "2") {
          const keputusanX = leftCol + 75;

          // Deteksi nomor kelas dari string kelas
          const kelasRaw = (kelas || "").toString().trim();

          const romawiKeAngka = (str: string): number => {
            const map: { [key: string]: number } = {
              VI: 6,
              V: 5,
              IV: 4,
              III: 3,
              II: 2,
              I: 1,
            };
            const upper = str.toUpperCase();
            for (const romawi of ["VI", "V", "IV", "III", "II", "I"]) {
              if (upper.startsWith(romawi)) {
                return map[romawi];
              }
            }
            return 0;
          };

          let kelasAngka = 0;
          const matchAngka = kelasRaw.match(/\d+/);
          if (matchAngka) {
            kelasAngka = parseInt(matchAngka[0]);
          } else {
            kelasAngka = romawiKeAngka(kelasRaw);
          }

          const romawi = ["", "I", "II", "III", "IV", "V", "VI"];
          const terbilang = [
            "",
            "Satu",
            "Dua",
            "Tiga",
            "Empat",
            "Lima",
            "Enam",
          ];

          if (kelasAngka >= 1 && kelasAngka <= 5) {
            // Kelas 1–5: tampilkan format naik kelas
            const kelasBerikutAngka = kelasAngka + 1;
            const kelasBerikutRomawi = romawi[kelasBerikutAngka];
            const kelasBerikutTerbilang = terbilang[kelasBerikutAngka];

            const tX = keputusanX;
            const tW = 105;
            const tStartY = kehadiranStartY;

            // Gambar border dan header tabel
            doc.setDrawColor(0);
            doc.setLineWidth(0.3);
            doc.setFillColor(200, 200, 200);
            doc.rect(tX, tStartY, tW, 7, "FD");
            doc.setFont("helvetica", "bold");
            doc.setFontSize(9);
            doc.text("Keputusan", tX + tW / 2, tStartY + 5, {
              align: "center",
            });

            // Gambar body tabel
            const bodyH = 30;
            doc.rect(tX, tStartY + 7, tW, bodyH, "D");

            doc.setFont("helvetica", "normal");
            doc.setFontSize(9);

            const lineX = tX + 3;
            let lineY = tStartY + 13;

            doc.text("Keputusan :", lineX, lineY);
            lineY += 5;
            doc.text("Berdasarkan pencapaian kompetensi pada", lineX, lineY);
            lineY += 5;
            doc.text("semester ke-1 dan ke-2, peserta didik", lineX, lineY);
            lineY += 5;

            // "Naik ke kelas : V (Lima)"
            doc.text("Naik ke kelas", lineX, lineY);
            doc.text(
              `: ${kelasBerikutRomawi} (${kelasBerikutTerbilang})`,
              lineX + 35,
              lineY
            );
            lineY += 5;

            // "Tinggal di kelas" dengan strikethrough
            const strikeText = "Tinggal di kelas";
            doc.text(strikeText, lineX, lineY);
            const strikeWidth = doc.getTextWidth(strikeText);
            doc.setLineWidth(0.4);
            doc.line(lineX, lineY - 1, lineX + strikeWidth, lineY - 1);
            doc.text(": -", lineX + 35, lineY);
          } else {
            // Kelas 6 atau tidak diketahui: pakai template lama
            doc.setFont("helvetica", "bold");
            doc.text("Keputusan", keputusanX, additionalY - 7);
            doc.setFont("helvetica", "normal");

            autoTable(doc, {
              startY: kehadiranStartY,
              head: [["Keputusan"]],
              body: [
                [
                  {
                    content:
                      "Berdasarkan hasil capaian pembelajaran seluruh kompetensi\n" +
                      `Ananda         :    ${namaSiswa.toUpperCase()}\n` +
                      "Dinyatakan  :    LULUS  /  TIDAK LULUS\n" +
                      "*Coret yang tidak perlu",
                    styles: { fontStyle: "normal" },
                  },
                ],
              ],
              theme: "grid",
              headStyles: {
                fillColor: [200, 200, 200],
                textColor: 0,
                fontStyle: "bold",
                halign: "center",
              },
              columnStyles: {
                0: { cellWidth: 105, halign: "left" },
              },
              styles: {
                fontSize: 9,
                cellPadding: 4,
                overflow: "linebreak",
              },
              margin: { left: keputusanX },
            });
          }
        }

        additionalY = Math.max(kehadiranEndY, doc.lastAutoTable.finalY) + 15;
      }

      const ttdY = additionalY;
      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");

      const tanggalRapor = latestSchoolData?.tanggalRapor || "23 Desember 2023";
      doc.text(`Bungeng, ${tanggalRapor}`, rightColTTD, ttdY + 10, {
        align: "left",
      });

      doc.text("Mengetahui :", leftColTTD, ttdY + 10);
      doc.text("Orang Tua / Wali,", leftColTTD, ttdY + 15);
      doc.text(namaOrtu || "_______________", leftColTTD, ttdY + 40);

      doc.setFontSize(10);
      doc.text("Wali Kelas,", rightColTTD, ttdY + 15);

      if (latestSchoolData?.ttdGuru) {
        try {
          const compressed = await compressImageToBase64(
            latestSchoolData.ttdGuru,
            800,
            1.0
          );
          doc.addImage(compressed, "JPEG", rightColTTD - 4, ttdY + 17, 40, 20);
        } catch (error) {
          console.log("Error adding guru signature:", error);
        }
      }

      doc.setFont("helvetica", "bold");
      const namaGuru = latestSchoolData?.namaGuru || "_______________";
      doc.text(namaGuru, rightColTTD, ttdY + 40);

      doc.setLineWidth(0.3);
      const guruTextWidth = doc.getTextWidth(namaGuru);
      doc.line(rightColTTD, ttdY + 41, rightColTTD + guruTextWidth, ttdY + 41);

      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(
        `NIP. ${latestSchoolData?.nipGuru || "_______________"}`,
        rightColTTD,
        ttdY + 45
      );

      const kepsekY = ttdY + 55;

      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");
      doc.text("Mengetahui,", centerColTTD, kepsekY, { align: "center" });
      doc.text("Kepala Sekolah", centerColTTD, kepsekY + 5, {
        align: "center",
      });

      if (latestSchoolData?.ttdKepsek) {
        try {
          const compressed = await compressImageToBase64(
            latestSchoolData.ttdKepsek,
            800,
            1.0
          );
          doc.addImage(
            compressed,
            "JPEG",
            centerColTTD - 20,
            kepsekY + 7,
            40,
            20
          );
        } catch (error) {
          console.log("Error adding kepsek signature:", error);
        }
      }

      doc.setFont("helvetica", "bold");
      const namaKepsek = latestSchoolData?.namaKepsek || "_______________";
      doc.text(namaKepsek, centerColTTD, kepsekY + 30, { align: "center" });

      doc.setLineWidth(0.3);
      const kepsekTextWidth = doc.getTextWidth(namaKepsek);
      doc.line(
        centerColTTD - kepsekTextWidth / 2,
        kepsekY + 31,
        centerColTTD + kepsekTextWidth / 2,
        kepsekY + 31
      );

      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(
        `NIP. ${latestSchoolData?.nipKepsek || "_______________"}`,
        centerColTTD,
        kepsekY + 35,
        { align: "center" }
      );

      // ✅ UBAH nama file
      const fileName = `Rapor_Sem${selectedSemester}_${namaSiswa.replace(
        /\s+/g,
        "_"
      )}.pdf`;
      doc.save(fileName);

      console.log("=== PDF GENERATION COMPLETE ===");
      console.log(`✅ PDF berhasil didownload: ${fileName}`);
    } catch (error) {
      console.error("=== PDF ERROR DETAILS ===");
      console.error("Error:", error);

      const errorMessage =
        error instanceof Error ? error.message : String(error);
      const errorStack = error instanceof Error ? error.stack : "";

      console.error("Error message:", errorMessage);
      console.error("Error stack:", errorStack);

      alert(
        `Gagal membuat PDF untuk ${rowData.Data1}\n\nError: ${errorMessage}\n\nCek console untuk detail lengkap.`
      );
    } finally {
      setDownloadingId(null);
    }
  };

  if (loading)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>Loading...</div>
    );
  if (error)
    return (
      <div style={{ textAlign: "center", color: "red", padding: "20px" }}>
        Error: {error}
      </div>
    );
  if (data.length === 0)
    return (
      <div style={{ textAlign: "center", padding: "20px" }}>
        No data available
      </div>
    );

  const allHeaders = [
    "Data1",
    "Data2",
    "Data3",
    "Data4",
    "Data5",
    "Data6",
    "Data7",
    "Data8",
    "Data9",
    "Data10",
    "Data11",
    "Data12",
    "Data13",
    "Data14",
    "Data15",
    "Data16",
    "Data17",
  ];

  const hiddenHeaders = new Set(["Data2", "Data3", "Data4", "Data5"]);

  const fixedKeepHeaders = new Set(["Data1", "Data15", "Data16", "Data17"]);

  const headers = allHeaders.filter((h) => {
    if (hiddenHeaders.has(h)) return false; // ← sembunyikan kolom ini
    if (fixedKeepHeaders.has(h)) return true;
    const dispH = (data[0]?.[h] || "").trim();
    return (
      dispH !== "" &&
      dispH !== h &&
      !dispH.includes("#REF!") &&
      !dispH.includes("#N/A") &&
      !dispH.toUpperCase().includes("N/A")
    );
  });

  const displayHeaders = headers.map((header) => data[0][header] || "");
  const actualData = data.slice(1);

  const filteredData =
    selectedKelas === "ALL"
      ? actualData
      : actualData.filter((row) => row.Data2 === selectedKelas);

  return (
    <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
      <h1
        style={{
          textAlign: "center",
          color: "#333",
          marginBottom: "15px",
          fontSize: "20px",
        }}
      >
        📊 Rekap Nilai Siswa
      </h1>

      {/* ✅ TAMBAH: Filter Semester DAN Kelas */}
      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <label style={{ fontSize: "14px", color: "#666", marginRight: "10px" }}>
          Semester:
        </label>
        <select
          value={selectedSemester}
          onChange={(e) => setSelectedSemester(e.target.value)}
          style={{
            padding: "10px 15px",
            fontSize: "16px",
            borderRadius: "4px",
            border: "1px solid #ddd",
            minWidth: "150px",
            cursor: "pointer",
            backgroundColor: "white",
            marginRight: "20px",
          }}
        >
          <option value="1">Semester 1</option>
          <option value="2">Semester 2</option>
        </select>

        <label style={{ fontSize: "14px", color: "#666", marginRight: "10px" }}>
          Filter Kelas:
        </label>
        <select
          value={selectedKelas}
          onChange={(e) => setSelectedKelas(e.target.value)}
          style={{
            padding: "10px 15px",
            fontSize: "16px",
            borderRadius: "4px",
            border: "1px solid #ddd",
            minWidth: "200px",
            cursor: "pointer",
            backgroundColor: "white",
          }}
        >
          <option value="ALL">Semua Kelas</option>
          {availableKelas.map((kelas, index) => (
            <option key={index} value={kelas}>
              {kelas}
            </option>
          ))}
        </select>
      </div>

      <div
        style={{
          textAlign: "center",
          marginBottom: "15px",
          fontSize: "14px",
          color: "#666",
        }}
      >
        Menampilkan {filteredData.length} siswa - Semester {selectedSemester}
      </div>

      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={downloadAllRekapPDF}
          disabled={isDownloadingAll || filteredData.length === 0}
          style={{
            padding: "12px 24px",
            backgroundColor:
              isDownloadingAll || filteredData.length === 0
                ? "#ccc"
                : "#E91E63",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor:
              isDownloadingAll || filteredData.length === 0
                ? "not-allowed"
                : "pointer",
            fontWeight: "bold",
            fontSize: "14px",
          }}
        >
          {isDownloadingAll
            ? "⏳ Membuat PDF..."
            : `📊 Download Rekap PDF (${filteredData.length} siswa)`}
        </button>
      </div>
      <div style={{ textAlign: "center", marginBottom: "15px" }}>
        <button
          onClick={downloadRekapExcel}
          disabled={isDownloadingExcel || filteredData.length === 0}
          style={{
            padding: "12px 24px",
            backgroundColor:
              isDownloadingExcel || filteredData.length === 0
                ? "#ccc"
                : "#217346",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor:
              isDownloadingExcel || filteredData.length === 0
                ? "not-allowed"
                : "pointer",
            fontWeight: "bold",
            fontSize: "14px",
          }}
        >
          {isDownloadingExcel
            ? "⏳ Membuat Excel..."
            : `📗 Download Rekap Excel (${filteredData.length} siswa)`}
        </button>
      </div>
      <div
        style={{
          overflowX: "auto",
          overflowY: "auto",
          maxHeight: "calc(100vh - 200px)",
          boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
          borderRadius: "8px",
          position: "relative",
        }}
      >
        <table
          style={{
            borderCollapse: "separate",
            borderSpacing: 0,
            minWidth: "100%",
            width: "max-content",
            tableLayout: "fixed",
          }}
        >
          <thead style={{ position: "sticky", top: 0, zIndex: 100 }}>
            <tr style={{ backgroundColor: "#f4f4f4" }}>
              <th
                style={{
                  padding: "8px 4px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "30px",
                  minWidth: "30px",
                  position: "sticky",
                  left: 0,
                  backgroundColor: "#f4f4f4",
                  zIndex: 3,
                  boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                  fontSize: "12px",
                }}
              >
                No
              </th>
              {displayHeaders.map((header, index) => {
                const currentHeader = headers[index];
                const isNamaColumn = currentHeader === "Data1";

                return (
                  <th
                    key={index}
                    style={{
                      padding: "8px 4px",
                      textAlign: "center",
                      borderBottom: "2px solid #ddd",
                      fontWeight: "bold",
                      width: isNamaColumn ? "150px" : "100px",
                      minWidth: isNamaColumn ? "150px" : "100px",
                      position: isNamaColumn ? "sticky" : "static",
                      left: isNamaColumn ? "30px" : "auto",
                      backgroundColor: "#f4f4f4",
                      zIndex: isNamaColumn ? 2 : 1,
                      boxShadow: isNamaColumn
                        ? "2px 0 5px rgba(0,0,0,0.1)"
                        : "none",
                      fontSize: "12px",
                      whiteSpace: isNamaColumn ? "nowrap" : "normal",
                      wordBreak: "normal",
                      overflowWrap: "normal",
                      overflow: "visible",
                      lineHeight: "1.3",
                    }}
                  >
                    {header}
                  </th>
                );
              })}
              <th
                style={{
                  padding: "8px 4px",
                  textAlign: "center",
                  borderBottom: "2px solid #ddd",
                  fontWeight: "bold",
                  width: "100px",
                  minWidth: "100px",
                  backgroundColor: "#f4f4f4",
                  zIndex: 1,
                  fontSize: "12px",
                }}
              >
                Aksi
              </th>
            </tr>
          </thead>
          <tbody>
            {filteredData.map((row, rowIndex) => (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                }}
              >
                <td
                  style={{
                    padding: "6px 4px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                    fontWeight: "bold",
                    color: "#666",
                    width: "30px",
                    minWidth: "30px",
                    position: "sticky",
                    left: 0,
                    backgroundColor: rowIndex % 2 === 0 ? "#fff" : "#f9f9f9",
                    zIndex: 2,
                    boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                    fontSize: "12px",
                  }}
                >
                  {rowIndex + 1}
                </td>
                {headers.map((header, colIndex) => {
                  const isNamaColumn = header === "Data1";

                  return (
                    <td
                      key={colIndex}
                      style={{
                        padding: "4px",
                        borderBottom: "1px solid #eee",
                        position: isNamaColumn ? "sticky" : "static",
                        left: isNamaColumn ? "30px" : "auto",
                        backgroundColor: isNamaColumn
                          ? rowIndex % 2 === 0
                            ? "#fff"
                            : "#f9f9f9"
                          : "transparent",
                        zIndex: isNamaColumn ? 1 : 0,
                        boxShadow: isNamaColumn
                          ? "2px 0 5px rgba(0,0,0,0.1)"
                          : "none",
                      }}
                    >
                      <div
                        style={{
                          padding: "4px 2px",
                          color: "#666",
                          fontWeight: "normal",
                          fontSize: "12px",
                          whiteSpace: "nowrap",
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          textAlign: isNamaColumn ? "left" : "center",
                        }}
                      >
                        {row[header] || ""}
                      </div>
                    </td>
                  );
                })}
                <td
                  style={{
                    padding: "4px",
                    borderBottom: "1px solid #eee",
                    textAlign: "center",
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      gap: "4px",
                      alignItems: "center",
                    }}
                  >
                    <button
                      onClick={() => downloadRaporPDF(row)}
                      disabled={downloadingId === row.Data1}
                      style={{
                        padding: "6px 12px",
                        backgroundColor:
                          downloadingId === row.Data1 ? "#ccc" : "#2196F3",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor:
                          downloadingId === row.Data1
                            ? "not-allowed"
                            : "pointer",
                        fontSize: "11px",
                        fontWeight: "bold",
                        whiteSpace: "nowrap",
                        width: "80px",
                      }}
                    >
                      {downloadingId === row.Data1 ? "⏳" : "📄 PDF"}
                    </button>
                    <button
                      onClick={async () => {
                        setDownloadingSampulId(row.Data1);
                        await downloadSampulPDF(row);
                        setDownloadingSampulId(null);
                      }}
                      disabled={downloadingSampulId === row.Data1}
                      style={{
                        padding: "6px 12px",
                        backgroundColor:
                          downloadingSampulId === row.Data1
                            ? "#ccc"
                            : "#FF9800",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor:
                          downloadingSampulId === row.Data1
                            ? "not-allowed"
                            : "pointer",
                        fontSize: "11px",
                        fontWeight: "bold",
                        whiteSpace: "nowrap",
                        width: "80px",
                      }}
                    >
                      {downloadingSampulId === row.Data1 ? "⏳" : "🖨️ Sampul"}
                    </button>
                    <button
                      onClick={() => setPreviewSiswa(row)}
                      style={{
                        padding: "6px 12px",
                        backgroundColor: "#673AB7",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "11px",
                        fontWeight: "bold",
                        whiteSpace: "nowrap",
                        width: "80px",
                      }}
                    >
                      👁️ Preview
                    </button>
                  </div>
                </td>
              </tr>
            ))}
            {(() => {
              // Kolom nilai dimulai dari Data6 (index 5 di headers)
              const excludeKeys = ["Data17"];
              const nilaiHeaders = headers
                .slice(1)
                .filter((h) => !excludeKeys.includes(h));

              // Hitung per kolom
              const summaryRows: {
                label: string;
                values: (string | number)[];
              }[] = [];

              const getValues = (h: string) =>
                filteredData
                  .map((row) => parseFloat(row[h]))
                  .filter((v) => !isNaN(v));

              const jumlah = nilaiHeaders.map((h) => {
                const vals = getValues(h);
                return vals.length > 0 ? vals.reduce((a, b) => a + b, 0) : "";
              });

              const rataRata = nilaiHeaders.map((h) => {
                const vals = getValues(h);
                return vals.length > 0
                  ? (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2)
                  : "";
              });

              const dayaSerap = nilaiHeaders.map((h) => {
                const vals = getValues(h);
                if (vals.length === 0) return "";
                const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
                return `${Math.round(avg)}%`;
              });

              const nilaiTerbesar = nilaiHeaders.map((h) => {
                const vals = getValues(h);
                return vals.length > 0 ? Math.max(...vals) : "";
              });

              const nilaiTerkecil = nilaiHeaders.map((h) => {
                const vals = getValues(h);
                return vals.length > 0 ? Math.min(...vals) : "";
              });

              // Kolom mapel murni (tidak termasuk Data15, Data16, Data17)
              const mapelOnlyHeaders = nilaiHeaders.filter(
                (h) => !["Data15", "Data16"].includes(h)
              );

              // Hitung horizontal untuk kolom Data15 dan Data16
              const calcHorizontal = (
                rows: any[],
                type: "sum" | "avg"
              ): { [key: string]: string } => {
                const result: { [key: string]: string } = {};

                // Baris Jumlah: total mapel semua siswa per kolom horizontal
                if (type === "sum") {
                  // Data15: jumlah semua nilai mapel dari semua siswa
                  const allMapelVals: number[] = [];
                  filteredData.forEach((row: RowData) => {
                    mapelOnlyHeaders.forEach((h) => {
                      const v = parseFloat(row[h]);
                      if (!isNaN(v)) allMapelVals.push(v);
                    });
                  });
                  result["Data15"] =
                    allMapelVals.length > 0
                      ? String(allMapelVals.reduce((a, b) => a + b, 0))
                      : "";

                  // Data16: rata-rata dari semua jumlah per siswa
                  const perSiswaJumlah = filteredData
                    .map((row) => {
                      const vals = mapelOnlyHeaders
                        .map((h) => parseFloat(row[h]))
                        .filter((v) => !isNaN(v));
                      return vals.length > 0
                        ? vals.reduce((a, b) => a + b, 0)
                        : null;
                    })
                    .filter((v) => v !== null) as number[];
                  result["Data16"] =
                    perSiswaJumlah.length > 0
                      ? (
                          perSiswaJumlah.reduce((a, b) => a + b, 0) /
                          perSiswaJumlah.length
                        ).toFixed(2)
                      : "";
                }

                if (type === "avg") {
                  // Data15: rata-rata jumlah per siswa
                  const perSiswaJumlah = filteredData
                    .map((row) => {
                      const vals = mapelOnlyHeaders
                        .map((h) => parseFloat(row[h]))
                        .filter((v) => !isNaN(v));
                      return vals.length > 0
                        ? vals.reduce((a, b) => a + b, 0)
                        : null;
                    })
                    .filter((v) => v !== null) as number[];
                  result["Data15"] =
                    perSiswaJumlah.length > 0
                      ? (
                          perSiswaJumlah.reduce((a, b) => a + b, 0) /
                          perSiswaJumlah.length
                        ).toFixed(2)
                      : "";

                  // Data16: rata-rata dari rata-rata per siswa
                  const perSiswaRata = filteredData
                    .map((row) => {
                      const vals = mapelOnlyHeaders
                        .map((h) => parseFloat(row[h]))
                        .filter((v) => !isNaN(v));
                      return vals.length > 0
                        ? vals.reduce((a, b) => a + b, 0) / vals.length
                        : null;
                    })
                    .filter((v) => v !== null) as number[];
                  result["Data16"] =
                    perSiswaRata.length > 0
                      ? (
                          perSiswaRata.reduce((a, b) => a + b, 0) /
                          perSiswaRata.length
                        ).toFixed(2)
                      : "";
                }

                return result;
              };

              // Bangun values array per summary row
              const buildValues = (
                baseValues: (string | number)[],
                horizontal: { [key: string]: string }
              ): (string | number)[] => {
                return nilaiHeaders.map((h, idx) => {
                  if (horizontal[h] !== undefined) return horizontal[h];
                  return (
                    baseValues[
                      nilaiHeaders
                        .filter((hh) => !["Data15", "Data16"].includes(hh))
                        .indexOf(h)
                    ] ?? ""
                  );
                });
              };

              // Hitung base values hanya dari mapelOnlyHeaders
              const getMapelVals = (h: string) =>
                filteredData
                  .map((row) => parseFloat(row[h]))
                  .filter((v) => !isNaN(v));

              const jumlahBase = mapelOnlyHeaders.map((h) => {
                const vals = getMapelVals(h);
                return vals.length > 0 ? vals.reduce((a, b) => a + b, 0) : "";
              });
              const rataRataBase = mapelOnlyHeaders.map((h) => {
                const vals = getMapelVals(h);
                return vals.length > 0
                  ? (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2)
                  : "";
              });
              const dayaSerapBase = mapelOnlyHeaders.map((h) => {
                const vals = getMapelVals(h);
                if (vals.length === 0) return "";
                const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
                return `${Math.round(avg)}%`;
              });
              const nilaiTerbesarBase = mapelOnlyHeaders.map((h) => {
                const vals = getMapelVals(h);
                return vals.length > 0 ? Math.max(...vals) : "";
              });
              const nilaiTerkecilBase = mapelOnlyHeaders.map((h) => {
                const vals = getMapelVals(h);
                return vals.length > 0 ? Math.min(...vals) : "";
              });

              // Gabungkan base values dengan horizontal untuk Data15/Data16
              const jumlahFinal = nilaiHeaders.map((h) => {
                if (h === "Data15") {
                  const vals = filteredData
                    .map((row) => {
                      const v = mapelOnlyHeaders
                        .map((mh) => parseFloat(row[mh]))
                        .filter((v) => !isNaN(v));
                      return v.length > 0 ? v.reduce((a, b) => a + b, 0) : null;
                    })
                    .filter((v) => v !== null) as number[];
                  return vals.length > 0 ? vals.reduce((a, b) => a + b, 0) : "";
                }
                if (h === "Data16") {
                  // Total jumlah semua siswa dibagi jumlah mapel
                  const totalJumlah = filteredData
                    .map((row) => {
                      const v = mapelOnlyHeaders
                        .map((mh) => parseFloat(row[mh]))
                        .filter((v) => !isNaN(v));
                      return v.length > 0 ? v.reduce((a, b) => a + b, 0) : null;
                    })
                    .filter((v) => v !== null) as number[];
                  const grandTotal = totalJumlah.reduce((a, b) => a + b, 0);
                  return mapelOnlyHeaders.length > 0
                    ? (grandTotal / mapelOnlyHeaders.length).toFixed(2)
                    : "";
                }
                const idx = mapelOnlyHeaders.indexOf(h);
                return idx >= 0 ? jumlahBase[idx] : "";
              });

              const rataRataFinal = nilaiHeaders.map((h) => {
                if (h === "Data15") {
                  // Jumlah semua rata-rata per mapel
                  const totalRataMapel = mapelOnlyHeaders.reduce((sum, mh) => {
                    const vals = filteredData
                      .map((row: RowData) => parseFloat(row[mh]))
                      .filter((v) => !isNaN(v));
                    const avg =
                      vals.length > 0
                        ? vals.reduce((a, b) => a + b, 0) / vals.length
                        : 0;
                    return sum + avg;
                  }, 0);
                  return mapelOnlyHeaders.length > 0
                    ? totalRataMapel.toFixed(2)
                    : "";
                }
                if (h === "Data16") {
                  // Jumlah rata-rata semua mapel dibagi jumlah mapel
                  const totalRataMapel = mapelOnlyHeaders.reduce((sum, mh) => {
                    const vals = filteredData
                      .map((row: RowData) => parseFloat(row[mh]))
                      .filter((v) => !isNaN(v));
                    const avg =
                      vals.length > 0
                        ? vals.reduce((a, b) => a + b, 0) / vals.length
                        : 0;
                    return sum + avg;
                  }, 0);
                  return mapelOnlyHeaders.length > 0
                    ? (totalRataMapel / mapelOnlyHeaders.length).toFixed(2)
                    : "";
                }
                const idx = mapelOnlyHeaders.indexOf(h);
                return idx >= 0 ? rataRataBase[idx] : "";
              });

              const dayaSerapFinal = nilaiHeaders.map((h) => {
                if (h === "Data15") {
                  // Jumlah daya serap per mapel (rata-rata per mapel dalam %)
                  const totalDayaSerap = mapelOnlyHeaders.reduce((sum, mh) => {
                    const vals = filteredData
                      .map((row: RowData) => parseFloat(row[mh]))
                      .filter((v) => !isNaN(v));
                    const avg =
                      vals.length > 0
                        ? vals.reduce((a, b) => a + b, 0) / vals.length
                        : 0;
                    return sum + avg;
                  }, 0);
                  return mapelOnlyHeaders.length > 0
                    ? `${Math.round(totalDayaSerap)}%`
                    : "";
                }
                if (h === "Data16") {
                  // Rata-rata daya serap semua mapel dibagi jumlah mapel
                  const totalDayaSerap = mapelOnlyHeaders.reduce((sum, mh) => {
                    const vals = filteredData
                      .map((row: RowData) => parseFloat(row[mh]))
                      .filter((v) => !isNaN(v));
                    const avg =
                      vals.length > 0
                        ? vals.reduce((a, b) => a + b, 0) / vals.length
                        : 0;
                    return sum + avg;
                  }, 0);
                  return mapelOnlyHeaders.length > 0
                    ? `${Math.round(totalDayaSerap / mapelOnlyHeaders.length)}%`
                    : "";
                }
                const idx = mapelOnlyHeaders.indexOf(h);
                return idx >= 0 ? dayaSerapBase[idx] : "";
              });

              const nilaiTerbesarFinal = nilaiHeaders.map((h) => {
                if (h === "Data15") {
                  // Nilai terbesar dari jumlah nilai per siswa
                  const perSiswaJumlah = filteredData
                    .map((row: RowData) => {
                      const v = mapelOnlyHeaders
                        .map((mh) => parseFloat(row[mh]))
                        .filter((v) => !isNaN(v));
                      return v.length > 0 ? v.reduce((a, b) => a + b, 0) : null;
                    })
                    .filter((v) => v !== null) as number[];
                  return perSiswaJumlah.length > 0
                    ? Math.max(...perSiswaJumlah)
                    : "";
                }
                if (h === "Data16") {
                  // Nilai terbesar dari rata-rata per siswa
                  const perSiswaRata = filteredData
                    .map((row: RowData) => {
                      const v = mapelOnlyHeaders
                        .map((mh) => parseFloat(row[mh]))
                        .filter((v) => !isNaN(v));
                      return v.length > 0
                        ? v.reduce((a, b) => a + b, 0) / v.length
                        : null;
                    })
                    .filter((v) => v !== null) as number[];
                  return perSiswaRata.length > 0
                    ? Math.max(...perSiswaRata).toFixed(2)
                    : "";
                }
                const idx = mapelOnlyHeaders.indexOf(h);
                return idx >= 0 ? nilaiTerbesarBase[idx] : "";
              });

              const nilaiTerkecilFinal = nilaiHeaders.map((h) => {
                if (h === "Data15") {
                  // Jumlah semua nilai terkecil per mapel
                  const totalTerkecil = mapelOnlyHeaders.reduce((sum, mh) => {
                    const vals = filteredData
                      .map((row: RowData) => parseFloat(row[mh]))
                      .filter((v) => !isNaN(v));
                    const min = vals.length > 0 ? Math.min(...vals) : 0;
                    return sum + min;
                  }, 0);
                  return mapelOnlyHeaders.length > 0
                    ? String(totalTerkecil)
                    : "";
                }
                if (h === "Data16") {
                  // Nilai terkecil dari rata-rata per siswa
                  const perSiswaRata = filteredData
                    .map((row: RowData) => {
                      const v = mapelOnlyHeaders
                        .map((mh) => parseFloat(row[mh]))
                        .filter((v) => !isNaN(v));
                      return v.length > 0
                        ? v.reduce((a, b) => a + b, 0) / v.length
                        : null;
                    })
                    .filter((v) => v !== null) as number[];
                  return perSiswaRata.length > 0
                    ? Math.min(...perSiswaRata).toFixed(2)
                    : "";
                }
                const idx = mapelOnlyHeaders.indexOf(h);
                return idx >= 0 ? nilaiTerkecilBase[idx] : "";
              });

              const summaryData = [
                {
                  label: "Jumlah",
                  values: jumlahFinal,
                  bold: true,
                  bg: "#e3f2fd",
                },
                {
                  label: "Rata-rata Kelas",
                  values: rataRataFinal,
                  bold: true,
                  bg: "#e8f5e9",
                },
                {
                  label: "Daya Serap",
                  values: dayaSerapFinal,
                  bold: true,
                  bg: "#fff3e0",
                },
                {
                  label: "Nilai Terbesar",
                  values: nilaiTerbesarFinal,
                  bold: true,
                  bg: "#fce4ec",
                },
                {
                  label: "Nilai Terkecil",
                  values: nilaiTerkecilFinal,
                  bold: true,
                  bg: "#f3e5f5",
                },
              ];

              return summaryData.map((summary, sIdx) => (
                <tr
                  key={`summary-${sIdx}`}
                  style={{ backgroundColor: summary.bg }}
                >
                  <td
                    style={{
                      padding: "6px 4px",
                      borderBottom: "1px solid #ddd",
                      borderTop: sIdx === 0 ? "2px solid #333" : "none",
                      textAlign: "center",
                      fontWeight: "bold",
                      color: "#666",
                      width: "30px",
                      minWidth: "30px",
                      position: "sticky",
                      left: 0,
                      backgroundColor: summary.bg,
                      zIndex: 2,
                      boxShadow: "2px 0 5px rgba(0,0,0,0.1)",
                      fontSize: "12px",
                    }}
                  >
                    -
                  </td>
                  {headers.map((header, colIndex) => {
                    const isNamaColumn = header === "Data1";
                    const nilaiIdx = nilaiHeaders.indexOf(header);
                    const value = isNamaColumn
                      ? summary.label
                      : nilaiIdx >= 0
                      ? summary.values[nilaiIdx]
                      : "";

                    return (
                      <td
                        key={colIndex}
                        style={{
                          padding: "4px",
                          borderBottom: "1px solid #ddd",
                          borderTop: sIdx === 0 ? "2px solid #333" : "none",
                          position: isNamaColumn ? "sticky" : "static",
                          left: isNamaColumn ? "30px" : "auto",
                          backgroundColor: summary.bg,
                          zIndex: isNamaColumn ? 1 : 0,
                          boxShadow: isNamaColumn
                            ? "2px 0 5px rgba(0,0,0,0.1)"
                            : "none",
                          textAlign: isNamaColumn ? "left" : "center",
                          fontWeight: summary.bold ? "bold" : "normal",
                          fontSize: "12px",
                          color: "#333",
                        }}
                      >
                        {String(value ?? "")}
                      </td>
                    );
                  })}
                  <td
                    style={{
                      padding: "4px",
                      borderBottom: "1px solid #ddd",
                      borderTop: sIdx === 0 ? "2px solid #333" : "none",
                    }}
                  />
                </tr>
              ));
            })()}
          </tbody>
        </table>
      </div>
      {/* ── MODAL PREVIEW RAPOR ── */}
      {previewSiswa && (
        <PreviewRaporModal
          rowData={previewSiswa}
          selectedSemester={selectedSemester}
          availableSheets={availableSheets}
          schoolData={localSchoolData || schoolData}
          onClose={() => setPreviewSiswa(null)}
        />
      )}
    </div>
  );
};

// ✅ KOMPONEN BARU - AppContent (yang menggunakan context)
const AppContent = () => {
  const [menuOpen, setMenuOpen] = useState(false);
  const { loading } = useRekapData();
  const navigate = useNavigate();

  useEffect(() => {
    navigate("/");
  }, []);

  return (
    <div>
      <style>{`
        @keyframes pulse {
          0%, 100% { opacity: 1; }
          50% { opacity: 0.6; }
        }
      `}</style>
      <div style={{ padding: "10px", margin: "0 auto", maxWidth: "100vw" }}>
        {/* ✅ TAMBAH: Indikator syncing */}
        {loading && (
          <div
            style={{
              position: "fixed",
              top: "15px",
              right: "15px",
              backgroundColor: "#2196F3",
              color: "white",
              padding: "8px 16px",
              borderRadius: "20px",
              fontSize: "12px",
              zIndex: 1003,
              boxShadow: "0 2px 10px rgba(0,0,0,0.2)",
              animation: "pulse 1.5s infinite",
            }}
          >
            🔄 Menyinkronkan data...
          </div>
        )}
        {/* Hamburger Menu Button */}
        <button
          onClick={() => setMenuOpen(!menuOpen)}
          style={{
            position: "fixed",
            top: "15px",
            left: "15px",
            width: "50px",
            height: "50px",
            backgroundColor: "#4CAF50",
            color: "white",
            border: "none",
            borderRadius: "8px",
            cursor: "pointer",
            zIndex: 1002,
            display: "flex",
            flexDirection: "column",
            justifyContent: "center",
            alignItems: "center",
            gap: "5px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.2)",
          }}
        >
          <div
            style={{
              width: "25px",
              height: "3px",
              backgroundColor: "white",
              borderRadius: "2px",
            }}
          ></div>
          <div
            style={{
              width: "25px",
              height: "3px",
              backgroundColor: "white",
              borderRadius: "2px",
            }}
          ></div>
          <div
            style={{
              width: "25px",
              height: "3px",
              backgroundColor: "white",
              borderRadius: "2px",
            }}
          ></div>
        </button>

        {/* Menu Overlay */}
        {menuOpen && (
          <div
            style={{
              position: "fixed",
              top: 0,
              left: 0,
              right: 0,
              bottom: 0,
              backgroundColor: "rgba(0, 0, 0, 0.5)",
              zIndex: 1001,
            }}
            onClick={() => setMenuOpen(false)}
          >
            {/* Menu Panel */}
            <div
              style={{
                position: "fixed",
                top: 0,
                left: 0,
                width: "280px",
                height: "100vh",
                backgroundColor: "white",
                boxShadow: "2px 0 10px rgba(0,0,0,0.3)",
                padding: "80px 20px 20px 20px",
                display: "flex",
                flexDirection: "column",
                gap: "10px",
                overflowY: "auto",
                overflowX: "hidden",
              }}
              onClick={(e) => e.stopPropagation()}
            >
              <h2
                style={{
                  margin: "0 0 20px 0",
                  color: "#333",
                  fontSize: "20px",
                }}
              >
                📚 Menu
              </h2>

              <Link
                to="/"
                onClick={() => setMenuOpen(false)}
                style={{
                  padding: "15px 20px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "8px",
                  textDecoration: "none",
                  color: "#333",
                  fontWeight: "500",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = "#4CAF50")
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = "#f0f0f0")
                }
              >
                👨‍🎓 Data Siswa
              </Link>

              <Link
                to="/input-nilai"
                onClick={() => setMenuOpen(false)}
                style={{
                  padding: "15px 20px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "8px",
                  textDecoration: "none",
                  color: "#333",
                  fontWeight: "500",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = "#4CAF50")
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = "#f0f0f0")
                }
              >
                📝 Input Nilai
              </Link>

              <Link
                to="/input-tp"
                onClick={() => setMenuOpen(false)}
                style={{
                  padding: "15px 20px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "8px",
                  textDecoration: "none",
                  color: "#333",
                  fontWeight: "500",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = "#4CAF50")
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = "#f0f0f0")
                }
              >
                📚 Input TP
              </Link>

              <Link
                to="/data-mapel"
                onClick={() => setMenuOpen(false)}
                style={{
                  padding: "15px 20px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "8px",
                  textDecoration: "none",
                  color: "#333",
                  fontWeight: "500",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = "#4CAF50")
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = "#f0f0f0")
                }
              >
                📖 Data Mata Pelajaran
              </Link>

              <Link
                to="/kehadiran"
                onClick={() => setMenuOpen(false)}
                style={{
                  padding: "15px 20px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "8px",
                  textDecoration: "none",
                  color: "#333",
                  fontWeight: "500",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = "#4CAF50")
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = "#f0f0f0")
                }
              >
                📋 Data Kehadiran
              </Link>

              <Link
                to="/kokurikuler"
                onClick={() => setMenuOpen(false)}
                style={{
                  padding: "15px 20px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "8px",
                  textDecoration: "none",
                  color: "#333",
                  fontWeight: "500",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = "#4CAF50")
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = "#f0f0f0")
                }
              >
                🌟 Data Kokurikuler
              </Link>

              <Link
                to="/ekstrakurikuler"
                onClick={() => setMenuOpen(false)}
                style={{
                  padding: "15px 20px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "8px",
                  textDecoration: "none",
                  color: "#333",
                  fontWeight: "500",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = "#4CAF50")
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = "#f0f0f0")
                }
              >
                🎯 Data Ekstrakurikuler
              </Link>

              <Link
                to="/data-sekolah"
                onClick={() => setMenuOpen(false)}
                style={{
                  padding: "15px 20px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "8px",
                  textDecoration: "none",
                  color: "#333",
                  fontWeight: "500",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = "#4CAF50")
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = "#f0f0f0")
                }
              >
                🏫 Data Sekolah
              </Link>
              <Link
                to="/rekap-nilai"
                onClick={() => setMenuOpen(false)}
                style={{
                  padding: "15px 20px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "8px",
                  textDecoration: "none",
                  color: "#333",
                  fontWeight: "500",
                  transition: "background-color 0.2s",
                }}
                onMouseEnter={(e) =>
                  (e.currentTarget.style.backgroundColor = "#4CAF50")
                }
                onMouseLeave={(e) =>
                  (e.currentTarget.style.backgroundColor = "#f0f0f0")
                }
              >
                📊 Rekap Nilai
              </Link>
            </div>
          </div>
        )}

        <Routes>
          <Route path="/" element={<DataSiswa />} />
          <Route path="/input-nilai" element={<InputNilai />} />
          <Route path="/kehadiran" element={<DataKehadiran />} />
          <Route path="/kokurikuler" element={<DataKokurikuler />} />
          <Route path="/ekstrakurikuler" element={<DataEkstrakurikuler />} />
          <Route path="/input-tp" element={<InputTP />} />
          <Route path="/data-mapel" element={<DataMapel />} />
          <Route path="/data-sekolah" element={<DataSekolah />} />
          <Route path="/rekap-nilai" element={<RekapNilai />} />
        </Routes>
      </div>
    </div>
  );
};

// ✅ KOMPONEN APP YANG BARU - Hanya sebagai wrapper Provider
const App = () => {
  return (
    <Router>
      <RekapProvider>
        <AppContent />
      </RekapProvider>
    </Router>
  );
};

export default App;
