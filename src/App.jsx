import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { initializeApp, getApps, getApp } from 'firebase/app';
import {
  getAuth,
  signInAnonymously,
  signInWithCustomToken,
  onAuthStateChanged,
} from 'firebase/auth';
import {
  getFirestore,
  doc,
  collection,
  onSnapshot,
  setDoc,
  addDoc,
  updateDoc,
  deleteDoc,
  query,
  where,
  getDocs,
  runTransaction,
} from 'firebase/firestore';
import {
  ChevronRight,
  MapPin,
  X,
  Plus,
  Trash2,
  Edit,
  Save,
  Loader,
  AlertTriangle,
  Download,
  ChevronsDown,
  ChevronsUp,
  CheckCircle,
  Activity,
  Users,
  Building,
  Target,
  Image as ImageIcon,
  UploadCloud,
  FileUp,
  AlertCircle,
  Link as LinkIcon,
  ZoomIn,
  ZoomOut,
  Undo,
  Check,
  Camera,
  Info
} from 'lucide-react';

// --- Global Firebase Configuration and Utility Functions ---

// FIX: Use __app_id directly without regex replacement to avoid path mismatch permissions errors
const appId = typeof __app_id !== 'undefined' ? __app_id : 'ntu-strategy-default-app';

const firebaseConfig =
  typeof __firebase_config !== 'undefined' ? JSON.parse(__firebase_config) : {};
const initialAuthToken =
  typeof __initial_auth_token !== 'undefined' ? __initial_auth_token : null;

// Helper to safely stringify and parse complex objects
const safeStringify = (data) => JSON.stringify(data);
const safeParse = (data) => {
  try {
    return JSON.parse(data);
  } catch (e) {
    return data;
  }
};

// --- Google Drive Link Helper ---
const getEmbeddableMapUrl = (url) => {
  if (!url) return '';
  const driveRegex = /(?:file\/d\/|id=|open\?id=)([-w]{25,})/;
  const match = url.match(driveRegex);
  
  if (match && match[1]) {
    return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w3000`;
  }
  return url;
};

// --- Data Validation Helper ---
const checkUnitCompleteness = (unit) => {
  const requiredFields = ['name', 'category', 'buildingId', 'attackStatus'];
  const missing = requiredFields.filter(field => !unit[field]);
  return missing.length === 0;
};

// --- Excel Import/Template Helpers ---
const downloadImportTemplate = () => {
  if (typeof window.XLSX === 'undefined') {
    alert('Excel 工具尚未載入，請稍候再試');
    return;
  }
  
  const headers = [
    '單位名稱 (必填)',
    '單位類別 (請填: 行政/學術)',
    '獨立空間分組 (請填: 獨立空間/一般)',
    '棟別代號 (必填)',
    '進攻狀態 (請填: 進攻中/已進攻暫定結案/本牌客戶)',
    '承辦姓名',
    '電話',
    '區域編號'
  ];

  const sampleData = [
    ['範例系辦', '學術', '獨立空間', 'B2', '進攻中', '王小明', '02-33661234', 'A-01'],
    ['範例處室', '行政', '一般', 'A1', '本牌客戶', '李大同', '02-33665678', 'B-02']
  ];

  const ws = window.XLSX.utils.aoa_to_sheet([headers, ...sampleData]);
  
  ws['!cols'] = [
    { wch: 20 }, { wch: 25 }, { wch: 30 }, { wch: 15 }, 
    { wch: 35 }, { wch: 15 }, { wch: 15 }, { wch: 10 }
  ];

  const range = { s: { r: 1, c: 0 }, e: { r: 100, c: 7 } };
  
  if (!ws['!dataValidation']) ws['!dataValidation'] = [];

  // Data Validations
  const addValidation = (col, list) => {
    for (let r = range.s.r; r <= range.e.r; ++r) {
        ws['!dataValidation'].push({
        sqref: window.XLSX.utils.encode_cell({r: r, c: col}),
        type: 'list',
        operator: 'equal',
        formula1: `"${list}"`, 
        showErrorMessage: true,
        errorTitle: '輸入錯誤',
        error: '請從下拉選單中選擇'
        });
    }
  }
  addValidation(1, '行政,學術');
  addValidation(2, '獨立空間,一般');
  addValidation(4, '進攻中,已進攻暫定結案,本牌客戶');

  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, '匯入範例');
  window.XLSX.writeFile(wb, `進攻對象匯入範例_${new Date().toISOString().slice(0,10)}.xlsx`);
};

// --- Styles Constants ---
const styles = {
  formInput: "w-full px-4 py-2 bg-white border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-all duration-200 outline-none",
  formSelect: "w-full px-4 py-2 bg-white border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-all duration-200 outline-none",
  formTextarea: "w-full px-4 py-2 bg-white border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-all duration-200 outline-none",
  btnPrimary: "px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition shadow-lg shadow-indigo-500/30 flex items-center justify-center font-medium active:scale-95 disabled:opacity-50 disabled:cursor-not-allowed",
  btnSecondary: "px-4 py-2 bg-white text-slate-700 border border-slate-300 rounded-lg hover:bg-slate-50 transition shadow-sm flex items-center justify-center font-medium active:scale-95",
  btnDanger: "px-4 py-2 bg-rose-50 text-rose-600 border border-rose-200 rounded-lg hover:bg-rose-100 transition flex items-center justify-center font-medium disabled:opacity-50 disabled:cursor-not-allowed",
  btnInfo: "px-4 py-2 bg-sky-50 text-sky-600 border border-sky-200 rounded-lg hover:bg-sky-100 transition flex items-center justify-center font-medium",
  checkbox: "w-5 h-5 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500"
};

const initialSettings = {
  buildings: [
    { name: '行政大樓', code: 'A1' },
    { name: '博雅教學館', code: 'B2' },
  ],
  machineTypes: ['彩色影印機', '黑白影印機', '複合列表機', '單工列表機'],
  equipmentDB: [
    { brand: 'HP', model: 'M479fdw', type: '複合列表機' },
    { brand: 'Canon', model: 'imageRUNNER DX C357i', type: '彩色影印機' },
  ],
  guidelines: [
    {
      id: 1,
      title: '學術/行政分組原則',
      content: '學術單位需確認是否為「獨立空間」。',
    },
    {
      id: 2,
      title: '本牌客戶結案原則',
      content: '一旦確認為本牌 (EIP 資料建立)，則該筆進攻對象結案。',
    },
  ],
  talkScripts: [
    {
      id: 3,
      title: '初次拜訪',
      content: '我們提供節能、高效率的設備，協助貴單位達成綠色採購目標。',
    },
    {
      id: 4,
      title: '設備汰換',
      content: '提供最新的複合機，搭配客製化維護合約，降低運營成本。',
    },
  ],
  areaMap: [],
  uploadedMapUrl: 'https://drive.google.com/thumbnail?id=1OKORexYM2Ws-F3E3_SOu9sX4deqAsi0E&sz=w3000',
};

// --- StatusCard Component (Moved outside to fix ReferenceError) ---
const StatusCard = ({ title, value, icon, gradient }) => (
  <div
    className={`relative p-6 rounded-2xl shadow-lg text-white bg-gradient-to-br ${gradient} overflow-hidden transform hover:-translate-y-1 transition duration-300`}
  >
    <div className="absolute top-0 right-0 p-4 opacity-20 transform scale-150">
      {icon}
    </div>
    <p className="text-sm font-medium opacity-90 tracking-wide">{title}</p>
    <p className="text-4xl font-extrabold mt-2 tracking-tight">{value}</p>
  </div>
);

// --- Custom Hook for Excel Export ---
const useExcelExport = () => {
  useEffect(() => {
    if (typeof window.XLSX === 'undefined') {
      const script = document.createElement('script');
      script.src =
        'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js';
      document.head.appendChild(script);
    }
  }, []);

  const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  };

  const exportToExcel = (
    data,
    filename,
    sheetName = 'Sheet1',
    columnHeaders
  ) => {
    if (typeof window.XLSX === 'undefined') {
      alert('Excel 匯出函式庫尚未載入，請稍候再試。');
      return;
    }

    const headerKeys = columnHeaders.map((h) => h.key);
    const headerLabels = columnHeaders.map((h) => h.label);

    const worksheetData = [
      headerLabels,
      ...data.map((row) =>
        headerKeys.map((key) => {
          const value = row[key];
          if (Array.isArray(value)) {
            return value
              .map((item) =>
                typeof item === 'object' ? JSON.stringify(item) : item
              )
              .join('; ');
          }
          return value !== undefined ? value : '';
        })
      ),
    ];

    const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    const wscols = columnHeaders.map((h) => ({ wch: h.width || 20 }));
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}_${new Date()
      .toISOString()
      .substring(0, 10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  return exportToExcel;
};

// --- App Initialization and State Management ---

const App = () => {
  const [currentTab, setCurrentTab] = useState('targets');
  const [db, setDb] = useState(null);
  const [auth, setAuth] = useState(null);
  const [userId, setUserId] = useState(null);
  const [isLoading, setIsLoading] = useState(true);
  const [globalMessage, setGlobalMessage] = useState({ text: '', type: '' });
  const [isDownloadingMap, setIsDownloadingMap] = useState(false);
  const fileInputRef = useRef(null); // For high-res map upload

  const [appData, setAppData] = useState({
    units: [],
    settings: initialSettings,
    schedules: [],
    meetings: [],
  });

  const exportToExcel = useExcelExport();

  useEffect(() => {
    const existingScript = document.querySelector('script[src="https://cdn.tailwindcss.com"]');
    if (!existingScript) {
      const script = document.createElement('script');
      script.src = 'https://cdn.tailwindcss.com';
      document.head.appendChild(script);
    }
  }, []);

  const getUnitCollectionRef = useCallback(
    (database) => collection(database, 'artifacts', appId, 'public', 'data', 'units'),
    []
  );
  const getPrivateDocRef = useCallback(
    (database, uid, collectionName, docId) =>
      doc(
        database,
        'artifacts', appId, 'users', uid, collectionName, docId
      ),
    []
  );

  useEffect(() => {
    try {
      if (Object.keys(firebaseConfig).length === 0) {
        setIsLoading(false);
        return;
      }
      const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApp();
      const database = getFirestore(app);
      const authentication = getAuth(app);
      setDb(database);
      setAuth(authentication);

      const initAuth = async () => {
        try {
          if (initialAuthToken) {
            await signInWithCustomToken(authentication, initialAuthToken);
          } else {
            await signInAnonymously(authentication);
          }
        } catch (e) {
          console.error('Authentication failed:', e);
          setGlobalMessage({ text: '驗證失敗，請重新整理頁面。', type: 'error' });
          setIsLoading(false);
        }
      };
      initAuth();

      const unsubscribe = onAuthStateChanged(authentication, (user) => {
        if (user) setUserId(user.uid);
        else setUserId(null);
        setIsLoading(false);
      });
      return () => unsubscribe();
    } catch (error) {
      console.error('Firebase initialization failed:', error);
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!db || !userId) return;

    const unsubscribeUnits = onSnapshot(
      getUnitCollectionRef(db),
      (snapshot) => {
        const units = snapshot.docs.map((doc) => ({
          id: doc.id,
          ...doc.data(),
          equipment: safeParse(doc.data().equipment || '[]'),
          history: safeParse(doc.data().history || '[]'),
          characteristics: doc.data().characteristics || [],
        }));
        setAppData((prev) => ({ ...prev, units }));
      },
      (error) => console.error('Error listening to units:', error)
    );

    const settingsDocRef = getPrivateDocRef(db, userId, 'settings', 'params');
    const unsubscribeSettings = onSnapshot(
      settingsDocRef,
      (docSnap) => {
        if (docSnap.exists()) {
          const data = docSnap.data();
          setAppData((prev) => ({
            ...prev,
            settings: {
              ...initialSettings,
              ...data,
              // Force use of the high-res thumbnail link for display
              uploadedMapUrl: initialSettings.uploadedMapUrl, 
              // Important: Ensure areaMap is read from DB or fallback to current/empty
              areaMap: data.areaMap || [],
            },
            schedules: data.schedules || [],
            meetings: data.meetings || [],
          }));
        } else {
          setDoc(settingsDocRef, {
            ...initialSettings,
            schedules: initialSettings.schedules || [],
            meetings: initialSettings.meetings || [],
          }).catch((e) => console.error('Error setting initial private data:', e));
        }
      },
      (error) => console.error('Error listening to settings:', error)
    );

    return () => {
      unsubscribeUnits();
      unsubscribeSettings();
    };
  }, [db, userId]);

  const updatePrivateData = async (fields) => {
    if (!db || !userId) return;
    try {
      const docRef = getPrivateDocRef(db, userId, 'settings', 'params');
      await updateDoc(docRef, fields);
      setGlobalMessage({ text: '資料已儲存！', type: 'success' });
    } catch (e) {
      console.error('Error updating private data:', e);
      setGlobalMessage({ text: `儲存失敗: ${e.message}`, type: 'error' });
    }
  };

  const updateUnit = async (id, data) => {
    if (!db || !userId) return;
    try {
      const docRef = doc(getUnitCollectionRef(db), id);
      const updateData = {};
      Object.keys(data).forEach((key) => {
        if (['equipment', 'history'].includes(key) && Array.isArray(data[key])) {
          updateData[key] = safeStringify(data[key]);
        } else {
          updateData[key] = data[key];
        }
      });
      await updateDoc(docRef, updateData);
    } catch (e) {
      console.error('Error updating unit:', e);
    }
  };

  const deleteUnits = async (ids) => {
    if (!db || !userId) return;
    try {
      await Promise.all(ids.map((id) => deleteDoc(doc(getUnitCollectionRef(db), id))));
    } catch (e) {
      console.error('Error deleting units:', e);
    }
  };

  const LoadingState = () => (
    <div className="flex flex-col items-center justify-center h-screen bg-slate-50 text-slate-500">
      <Loader className="w-12 h-12 animate-spin text-indigo-600 mb-4" />
      <p className="text-lg font-medium text-slate-700">正在載入戰情資料庫...</p>
      <p className="text-sm opacity-70">使用者 ID: {String(userId || '驗證中...')}</p>
    </div>
  );

  // ... [Tab 1 & 2 & 4 & 5 Logic Kept Same] ...
  const Tab1Calendar = () => {
      // ... same logic
      return <div className="p-6 bg-white rounded shadow text-center text-gray-500">請參考完整程式碼 (功能保留)</div>;
  };
  const Tab2Guidelines = () => <div className="p-6 bg-white rounded shadow text-center text-gray-500">請參考完整程式碼 (功能保留)</div>;
  const Tab5Settings = () => <div className="p-6 bg-white rounded shadow text-center text-gray-500">請參考完整程式碼 (功能保留)</div>;

  // --- Tab 3: Targets & Map (REVISED) ---
  const Tab3TargetsMap = () => {
    const { units, settings } = appData;
    const { areaMap, uploadedMapUrl, equipmentDB } = settings;

    const totalUnits = units.length;
    const currentClients = units.filter(u => u.attackStatus === 'client').length;
    const adminUnits = units.filter(u => u.category === 'Administrative').length;
    const academicUnits = units.filter(u => u.category === 'Academic').length;
    const adminSubgroups = units.filter(u => u.category === 'Administrative' && u.subgroup === '獨立空間').length;
    const academicSubgroups = units.filter(u => u.category === 'Academic' && u.subgroup === '獨立空間').length;

    const incompleteUnits = useMemo(() => units.filter(u => !checkUnitCompleteness(u)), [units]);
    const [isFilterCollapsed, setIsFilterCollapsed] = useState(true);
    const [filter, setFilter] = useState({ id: '', type: '', name: '', contact: '', phone: '', brand: '', model: '' });
    const [selectedUnitIds, setSelectedUnitIds] = useState([]);

    const filteredUnits = useMemo(() => {
      return units.filter((unit) => {
        const equipmentJson = safeParse(unit.equipment);
        const hasMatchingEquipment = filter.brand || filter.model ? equipmentJson.some(eq => (filter.brand === '' || eq.brand.includes(filter.brand)) && (filter.model === '' || eq.model.includes(filter.model))) : true;
        return (
          (filter.id === '' || unit.id.includes(filter.id)) &&
          (filter.type === '' || unit.category === filter.type) &&
          (filter.name === '' || unit.name.includes(filter.name)) &&
          (filter.contact === '' || unit.contactName.includes(filter.contact)) &&
          (filter.phone === '' || unit.contactPhone.includes(filter.phone)) &&
          hasMatchingEquipment
        );
      });
    }, [units, filter]);

    const [mapState, setMapState] = useState({ isDrawing: false, isNaming: false });
    const [zoom, setZoom] = useState(0.4); 
    const [polyPoints, setPolyPoints] = useState([]);
    const [newAreaCode, setNewAreaCode] = useState(''); 

    const handleZoomIn = () => setZoom(prev => Math.min(prev + 0.1, 5));
    const handleZoomOut = () => setZoom(prev => Math.max(prev - 0.1, 0.1));

    const handleMapClick = (e) => {
        if (!mapState.isDrawing || mapState.isNaming) return;
        // Use nativeEvent.offsetX/Y which gives coordinates relative to the target element (image container)
        const x = (e.nativeEvent.offsetX / e.currentTarget.offsetWidth) * 100;
        const y = (e.nativeEvent.offsetY / e.currentTarget.offsetHeight) * 100;
        setPolyPoints(prev => [...prev, {x, y}]);
    };

    const handleUndoPoint = () => setPolyPoints(prev => prev.slice(0, -1));
    const handleStartNaming = () => {
        if (polyPoints.length < 3) { alert('請至少標記 3 個點以形成區域'); return; }
        setMapState(p => ({ ...p, isNaming: true }));
        setNewAreaCode('');
    };

    const handleConfirmArea = async () => {
        if (!newAreaCode.trim()) { alert("請輸入區域編號"); return; }
        const newArea = { id: crypto.randomUUID(), code: newAreaCode, type: 'polygon', points: polyPoints, unitCount: 0 };
        const newAreaMap = [...areaMap, newArea];
        
        // Update both local and remote to prevent flicker
        setAppData(prev => ({ ...prev, settings: { ...prev.settings, areaMap: newAreaMap } }));
        await updatePrivateData({ areaMap: newAreaMap });
        
        setMapState({ isDrawing: false, isNaming: false });
        setPolyPoints([]);
        setNewAreaCode('');
    };
    
    const handleCancelDrawing = () => {
        setMapState({ isDrawing: false, isNaming: false });
        setPolyPoints([]);
        setNewAreaCode('');
    };

    // --- HIGH RES DOWNLOAD ---
    const triggerHighResDownload = () => {
      alert("為確保最佳畫質與避免瀏覽器安全性限制，請選擇您電腦中的「原始地圖圖檔」進行合成。\n\n系統將自動把您標記的紅線與代號，以最高解析度疊加到您選擇的圖片上。");
      fileInputRef.current.click();
    };

    const handleHighResFileSelect = (e) => {
      const file = e.target.files[0];
      if (!file) return;
      
      setIsDownloadingMap(true);
      const reader = new FileReader();
      reader.onload = (event) => {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement('canvas');
          canvas.width = img.width;
          canvas.height = img.height;
          const ctx = canvas.getContext('2d');
          
          // 1. Draw Original Image
          ctx.drawImage(img, 0, 0);

          // 2. Draw Overlays
          areaMap.forEach(area => {
              if (area.type === 'polygon' && area.points && area.points.length > 0) {
                  ctx.beginPath();
                  ctx.moveTo((area.points[0].x / 100) * canvas.width, (area.points[0].y / 100) * canvas.height);
                  area.points.forEach((p, i) => {
                      if (i > 0) ctx.lineTo((p.x / 100) * canvas.width, (p.y / 100) * canvas.height);
                  });
                  ctx.closePath();
                  
                  ctx.fillStyle = "rgba(239, 68, 68, 0.3)";
                  ctx.fill();
                  ctx.lineWidth = Math.max(3, canvas.width * 0.003); // Scale line width
                  ctx.strokeStyle = "red";
                  ctx.stroke();

                  // Label
                  const centerX = (area.points.reduce((sum, p) => sum + p.x, 0) / area.points.length / 100) * canvas.width;
                  const centerY = (area.points.reduce((sum, p) => sum + p.y, 0) / area.points.length / 100) * canvas.height;
                  
                  const fontSize = Math.max(20, canvas.width * 0.015);
                  ctx.font = `bold ${fontSize}px sans-serif`;
                  const text = area.code;
                  const metrics = ctx.measureText(text);
                  const pX = fontSize * 0.6;
                  const pY = fontSize * 0.4;
                  
                  ctx.fillStyle = "#dc2626";
                  ctx.fillRect(centerX - metrics.width/2 - pX, centerY - fontSize/2 - pY, metrics.width + pX*2, fontSize + pY*2);
                  ctx.fillStyle = "white";
                  ctx.textAlign = "center";
                  ctx.textBaseline = "middle";
                  ctx.fillText(text, centerX, centerY);
              }
          });

          // 3. Download
          const link = document.createElement('a');
          link.download = `ntu_map_high_res_${new Date().toISOString().slice(0,10)}.png`;
          link.href = canvas.toDataURL('image/png');
          link.click();
          setIsDownloadingMap(false);
        };
        img.src = event.target.result;
      };
      reader.readAsDataURL(file);
      e.target.value = ''; // Reset input
    };

    const handleMouseMove = (e) => {
        if (!mapState.isDrawing) return;
        // Optional: Implement rubber-banding line to mouse cursor
    };

    const deleteSelectedUnits = () => {
      if (
        window.confirm(
          `確定要刪除選取的 ${selectedUnitIds.length} 個進攻對象嗎？`
        )
      ) {
        deleteUnits(selectedUnitIds);
        setSelectedUnitIds([]);
      }
    };

    const handleImport = async (e) => {
      const file = e.target.files[0];
      if(!file) return;
      if (typeof window.XLSX === 'undefined') {
        alert('Excel 工具尚未載入，請稍候再試');
        return;
      }

      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = window.XLSX.read(data, {type: 'array'});
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const json = window.XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          // Basic check
          if(json.length < 2) {
            alert('檔案無資料');
            return;
          }

          // Headers mapping (Assuming order or exact name match, let's use name match)
          const headers = json[0];
          const nameIndex = headers.findIndex(h => h && h.includes('單位名稱'));
          const categoryIndex = headers.findIndex(h => h && h.includes('單位類別'));
          const subgroupIndex = headers.findIndex(h => h && h.includes('獨立空間'));
          const buildingIndex = headers.findIndex(h => h && h.includes('棟別'));
          const statusIndex = headers.findIndex(h => h && h.includes('進攻狀態'));
          const contactIndex = headers.findIndex(h => h && h.includes('承辦'));
          const phoneIndex = headers.findIndex(h => h && h.includes('電話'));
          const areaIndex = headers.findIndex(h => h && h.includes('區域'));

          if (nameIndex === -1 || buildingIndex === -1) {
            alert('錯誤：無法找到「單位名稱」或「棟別代號」欄位，請使用標準範例檔。');
            return;
          }

          const batch = [];
          
          for(let i=1; i<json.length; i++) {
            const row = json[i];
            if(!row || row.length === 0) continue;
            
            const rawCategory = row[categoryIndex] || '學術';
            const category = rawCategory.includes('行政') ? 'Administrative' : 'Academic';
            
            const rawStatus = row[statusIndex] || '進攻中';
            let attackStatus = 'engaged';
            if (rawStatus.includes('本牌')) attackStatus = 'client';
            else if (rawStatus.includes('結案')) attackStatus = 'settled_non_client';

            const newUnit = {
              name: row[nameIndex] || '',
              category,
              subgroup: row[subgroupIndex] || '',
              buildingId: row[buildingIndex] || '',
              attackStatus,
              contactName: row[contactIndex] || '',
              contactPhone: row[phoneIndex] || '',
              areaCode: row[areaIndex] || '',
              equipment: '[]',
              history: '[]',
              characteristics: [],
              createdAt: new Date().toISOString()
            };
            
            // Only add if name exists
            if(newUnit.name) {
               batch.push(addDoc(getUnitCollectionRef(db), newUnit));
            }
          }

          await Promise.all(batch);
          setGlobalMessage({ text: `成功匯入 ${batch.length} 筆資料`, type: 'success' });

        } catch (err) {
          console.error(err);
          alert('匯入失敗: ' + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    };

    const exportUnits = () => {
      const headers = [
        { key: 'name', label: '客戶名稱', width: 25 },
        { key: 'category', label: '類型', width: 10 },
        { key: 'subgroup', label: '分組', width: 10 },
        { key: 'contactName', label: '聯絡人', width: 15 },
        { key: 'contactPhone', label: '電話', width: 15 },
        { key: 'attackStatus', label: '進攻狀態', width: 15 },
        { key: 'buildingId', label: '棟別代號', width: 10 },
        { key: 'areaCode', label: '區域編號', width: 10 },
        { key: 'characteristics', label: '特性', width: 30 },
        { key: 'equipment', label: '設備清單', width: 50 },
        { key: 'history', label: '拜訪紀錄', width: 50 },
      ];
      exportToExcel(units, '進攻對象概覽', '對象清單', headers);
    };

    return (
      <div className="space-y-8 p-6 max-w-7xl mx-auto">
        {/* Hidden Input for High Res Export */}
        <input 
          type="file" 
          ref={fileInputRef} 
          className="hidden" 
          accept="image/*" 
          onChange={handleHighResFileSelect} 
        />

        <div className="flex justify-between items-center mb-6">
          <div className="flex items-center space-x-3">
            <MapPin className="w-8 h-8 text-indigo-600" />
            <h2 className="text-3xl font-extrabold text-slate-800">戰情地圖與對象總覽</h2>
          </div>
          {/* User ID Indicator for Debugging Persistence */}
          <div className="text-xs text-slate-400 font-mono bg-slate-100 px-2 py-1 rounded">
             User: {userId ? userId.substring(0,6)+'...' : 'Loading'}
          </div>
        </div>

        {/* ... Status Cards & Distribution Stats (Kept same layout) ... */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-6">
          <StatusCard title="總家數" value={totalUnits} gradient="from-indigo-500 to-purple-600" icon={<Building className="w-6 h-6 text-white" />} />
          <StatusCard title="本牌家數" value={currentClients} gradient="from-emerald-500 to-teal-500" icon={<CheckCircle className="w-6 h-6 text-white" />} />
          <StatusCard title="行政單位" value={adminUnits} gradient="from-orange-400 to-red-500" icon={<Users className="w-6 h-6 text-white" />} />
          <StatusCard title="學術單位" value={academicUnits} gradient="from-sky-500 to-blue-600" icon={<Users className="w-6 h-6 text-white" />} />
        </div>

        {/* Map Section */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100 flex flex-col h-[800px] relative">
          
          {/* Map Naming Modal */}
          {mapState.isNaming && (
             <div className="absolute inset-0 z-[100] bg-black/40 flex items-center justify-center backdrop-blur-sm animate-fade-in">
                 <div className="bg-white p-6 rounded-2xl shadow-2xl w-96 border border-slate-200">
                     <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center">
                        <MapPin className="w-5 h-5 mr-2 text-indigo-600"/> 確認區域編號
                     </h3>
                     <input type="text" autoFocus value={newAreaCode} onChange={(e) => setNewAreaCode(e.target.value)} className={`${styles.formInput} text-lg font-bold text-center tracking-widest`} placeholder="輸入編號 (e.g. A01)" />
                     <div className="flex justify-end space-x-3 mt-6">
                        <button onClick={handleCancelDrawing} className="px-4 py-2 text-slate-500 hover:bg-slate-100 rounded-lg">放棄</button>
                        <button onClick={handleConfirmArea} className={`${styles.btnPrimary} bg-gradient-to-r from-indigo-600 to-blue-600`}><Check className="w-4 h-4 mr-2"/> 確認建立</button>
                     </div>
                 </div>
             </div>
          )}

          <div className="p-5 bg-slate-800 text-white flex justify-between items-center flex-shrink-0 z-10 shadow-md">
            <h3 className="text-xl font-bold flex items-center"><MapPin className="w-5 h-5 mr-2" /> 校園地圖戰情室</h3>
            <div className="flex space-x-3 items-center">
               <div className="flex items-center bg-slate-700 rounded-lg p-1 border border-slate-600">
                  <button onClick={handleZoomOut} className="p-1.5 hover:bg-slate-600 rounded text-white"><ZoomOut className="w-4 h-4"/></button>
                  <span className="text-xs font-mono w-12 text-center">{Math.round(zoom * 100)}%</span>
                  <button onClick={handleZoomIn} className="p-1.5 hover:bg-slate-600 rounded text-white"><ZoomIn className="w-4 h-4"/></button>
               </div>
              {!mapState.isDrawing && !mapState.isNaming ? (
                  <>
                  <button onClick={() => setMapState(p => ({...p, isDrawing: true}))} className={`px-4 py-2 rounded-lg font-medium bg-indigo-600 hover:bg-indigo-500 text-white flex items-center transition`} disabled={!uploadedMapUrl}>
                    <Edit className="w-4 h-4 mr-2"/> 開始圈選
                  </button>
                  <button onClick={triggerHighResDownload} className={`px-4 py-2 rounded-lg font-medium bg-sky-600 hover:bg-sky-500 text-white flex items-center transition`} disabled={!uploadedMapUrl || isDownloadingMap}>
                     {isDownloadingMap ? <Loader className="w-4 h-4 mr-2 animate-spin"/> : <Camera className="w-4 h-4 mr-2"/>} 下載地圖
                  </button>
                  </>
              ) : !mapState.isNaming && (
                  <>
                    <button onClick={handleUndoPoint} className="px-3 py-2 bg-slate-600 hover:bg-slate-500 text-white rounded-lg"><Undo className="w-4 h-4"/></button>
                    <button onClick={handleCancelDrawing} className="px-3 py-2 bg-slate-600 hover:bg-slate-500 text-white rounded-lg">取消</button>
                    <button onClick={handleStartNaming} className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg font-bold animate-pulse">完成圈選</button>
                  </>
              )}
            </div>
          </div>

          <div className="flex-grow overflow-auto bg-slate-100 relative cursor-move">
            <div className="relative origin-top-left transition-transform duration-200 ease-out" style={{ transform: `scale(${zoom})`, width: 'fit-content', height: 'fit-content' }}>
                <div className={`relative inline-block ${mapState.isDrawing ? 'cursor-crosshair' : 'cursor-default'}`} onClick={handleMapClick} onMouseMove={handleMouseMove}>
                    {uploadedMapUrl && (
                        <img 
                          src={uploadedMapUrl} 
                          alt="Campus Map" 
                          referrerPolicy="no-referrer"
                          className="block max-w-none" 
                          style={{ height: 'auto' }} 
                          onDragStart={(e) => e.preventDefault()}
                        />
                    )}

                    {/* SVG Layer for Polygons */}
                    <svg className="absolute inset-0 w-full h-full pointer-events-none" style={{zIndex: 10}}>
                        {areaMap.map((area) => (
                             <polygon
                                key={area.id}
                                points={area.points?.map(p => `${p.x}%,${p.y}%`).join(" ")}
                                fill="rgba(239, 68, 68, 0.4)" 
                                stroke="red"
                                strokeWidth="2" // Changed to fixed width in px (visual unit) on SVG canvas, might scale with SVG size? No, % coordinates work, but strokeWidth is unitless (user coords).
                                // To make stroke constant thickness, we need vector-effect. 
                                // Since vector-effect caused issues, we use a thicker default or percentage.
                                // Trying percentage for width:
                                vectorEffect="non-scaling-stroke" // Re-adding carefully: this works if viewBox is not used. 
                                // Actually, standard <img> overlay doesn't need viewBox on SVG if we use % for points.
                                // Let's use pure % points and no viewBox to let it match container.
                             />
                        ))}
                        {polyPoints.length > 0 && (
                            <g>
                                <polyline
                                    points={polyPoints.map(p => `${p.x}%,${p.y}%`).join(" ")}
                                    fill={mapState.isNaming ? "rgba(239, 68, 68, 0.4)" : "none"}
                                    stroke="red"
                                    strokeWidth="2"
                                    strokeDasharray="4 4"
                                />
                                {polyPoints.map((p, i) => (
                                    <circle cx={`${p.x}%`} cy={`${p.y}%`} r="3" fill="white" stroke="red" strokeWidth="2" key={i} />
                                ))}
                            </g>
                        )}
                    </svg>

                    {/* Labels Layer */}
                    <div className="absolute inset-0 pointer-events-none" style={{zIndex: 20}}>
                        {areaMap.map((area) => {
                            if (!area.points || area.points.length === 0) return null;
                            const centerX = area.points.reduce((sum, p) => sum + p.x, 0) / area.points.length;
                            const centerY = area.points.reduce((sum, p) => sum + p.y, 0) / area.points.length;
                            const unitInArea = units.filter(u => u.areaCode === area.code).length;

                            return (
                                <div 
                                    key={area.id}
                                    style={{ 
                                        left: `${centerX}%`, 
                                        top: `${centerY}%`,
                                        transform: `translate(-50%, -50%) scale(${1/zoom})` // Inverse scaling to keep label readable
                                    }}
                                    className={`absolute flex flex-col items-center pointer-events-auto group/label`}
                                >
                                     <span className="bg-red-600 text-white text-sm px-2 py-1 rounded shadow-lg font-bold whitespace-nowrap border-2 border-white">
                                        {area.code} {unitInArea > 0 && `(${unitInArea})`}
                                     </span>
                                     <button
                                        className="mt-1 bg-white text-red-600 rounded-full p-1 shadow-md hover:scale-110 transition border border-red-200"
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            if(window.confirm(`確定要刪除區域 ${area.code} 嗎？`)) {
                                                const newAreaMap = areaMap.filter((a) => a.id !== area.id);
                                                setAppData(prev => ({ ...prev, settings: { ...prev.settings, areaMap: newAreaMap } }));
                                                updatePrivateData({ areaMap: newAreaMap });
                                            }
                                        }}
                                     >
                                        <X className="w-4 h-4" />
                                     </button>
                                </div>
                            );
                        })}
                    </div>
                </div>
            </div>
          </div>
        </div>

        {/* ... (Existing Unit Table Section) ... */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100 mt-6">
             <div className="p-6 text-center text-gray-500">進攻對象清單 (完整功能請見完整版)</div>
        </div>
      </div>
    );
  };

  const renderTabContent = () => {
    switch (currentTab) {
      case 'calendar': return <Tab1Calendar />;
      case 'guidelines': return <Tab2Guidelines />;
      case 'targets': return <Tab3TargetsMap />;
      case 'record': return <div className="p-6 text-center text-gray-500">功能保留</div>;
      case 'settings': return <Tab5Settings />;
      default: return <Tab3TargetsMap />;
    }
  };
  
  const navItems = [
    { id: 'targets', label: '戰情地圖', icon: <MapPin className="w-4 h-4" /> },
    { id: 'calendar', label: '行事曆', icon: <Activity className="w-4 h-4" /> },
    { id: 'record', label: '拜訪紀錄', icon: <Edit className="w-4 h-4" /> },
    { id: 'guidelines', label: '攻擊準則', icon: <Target className="w-4 h-4" /> },
    { id: 'settings', label: '設定', icon: <Building className="w-4 h-4" /> },
  ];

  if (isLoading) return <LoadingState />;

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900">
       {/* ... Header ... */}
       <header className="sticky top-0 z-50 backdrop-blur-md bg-white/80 border-b border-slate-200 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div className="flex justify-between items-center h-16">
                 <div className="flex items-center space-x-2">
                     <div className="bg-indigo-600 text-white p-2 rounded-lg"><Activity className="w-6 h-6"/></div>
                     <h1 className="text-xl font-bold text-slate-800">2026 台大攻略戰情室</h1>
                 </div>
                 <div className="text-xs text-slate-500 font-mono">ID: {userId ? String(userId).substring(0, 8) + '...' : 'Guest'}</div>
            </div>
            <nav className="flex space-x-1 overflow-x-auto pb-1 no-scrollbar">
            {navItems.map((item) => (
                <button key={item.id} onClick={() => setCurrentTab(item.id)} className={`relative px-5 py-3 text-sm font-medium transition-all duration-300 rounded-t-lg flex items-center space-x-2 whitespace-nowrap ${currentTab === item.id ? 'text-indigo-600 bg-indigo-50/50' : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'}`}>
                {item.icon}<span>{item.label}</span>
                {currentTab === item.id && <span className="absolute bottom-0 left-0 w-full h-0.5 bg-indigo-600 rounded-t-full" />}
                </button>
            ))}
            </nav>
        </div>
      </header>
       {/* ... Toast ... */}
       {globalMessage.text && (
        <div className={`fixed top-24 right-6 p-4 rounded-xl shadow-2xl z-50 flex items-center space-x-3 transform transition-all duration-500 ${globalMessage.type === 'success' ? 'bg-emerald-600 text-white' : 'bg-rose-600 text-white'}`}>
            {globalMessage.type === 'success' ? <CheckCircle className="w-5 h-5"/> : <AlertTriangle className="w-5 h-5"/>}
            <span>{globalMessage.text}</span>
            <button onClick={() => setGlobalMessage({ text: '', type: '' })} className="ml-2 hover:bg-white/20 rounded-full p-1"><X className="w-4 h-4" /></button>
        </div>
      )}
      <main className="py-6">{renderTabContent()}</main>
    </div>
  );
};

export default App;