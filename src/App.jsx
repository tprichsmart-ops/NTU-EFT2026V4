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
  getDoc
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
  RefreshCw
} from 'lucide-react';

// --- Global Firebase Configuration ---
const appId = typeof __app_id !== 'undefined' ? __app_id : 'ntu-strategy-default-app';
const firebaseConfig = typeof __firebase_config !== 'undefined' ? JSON.parse(__firebase_config) : {};
const initialAuthToken = typeof __initial_auth_token !== 'undefined' ? __initial_auth_token : null;

// Helper to safely stringify and parse
const safeStringify = (data) => JSON.stringify(data);
const safeParse = (data) => {
  try {
    return JSON.parse(data);
  } catch (e) {
    return data;
  }
};

// Data Validation
const checkUnitCompleteness = (unit) => {
  const requiredFields = ['name', 'category', 'buildingId', 'attackStatus'];
  const missing = requiredFields.filter(field => !unit[field]);
  return missing.length === 0;
};

// Excel Template Helper
const downloadImportTemplate = () => {
  if (typeof window.XLSX === 'undefined') {
    alert('Excel 工具尚未載入，請稍候再試');
    return;
  }
  
  const headers = [
    '單位名稱 (必填)', '單位類別 (請填: 行政/學術)', '獨立空間分組 (請填: 獨立空間/一般)',
    '棟別代號 (必填)', '進攻狀態 (請填: 進攻中/已進攻暫定結案/本牌客戶)',
    '承辦姓名', '電話', '區域編號'
  ];

  const sampleData = [
    ['範例系辦', '學術', '獨立空間', 'B2', '進攻中', '王小明', '02-33661234', 'A-01'],
    ['範例處室', '行政', '一般', 'A1', '本牌客戶', '李大同', '02-33665678', 'B-02']
  ];

  const ws = window.XLSX.utils.aoa_to_sheet([headers, ...sampleData]);
  ws['!cols'] = [{ wch: 20 }, { wch: 25 }, { wch: 30 }, { wch: 15 }, { wch: 35 }, { wch: 15 }, { wch: 15 }, { wch: 10 }];

  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, '匯入範例');
  window.XLSX.writeFile(wb, `進攻對象匯入範例_${new Date().toISOString().slice(0,10)}.xlsx`);
};

// Styles
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

// Initial Data
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
    { id: 1, title: '學術/行政分組原則', content: '學術單位需確認是否為「獨立空間」。' },
    { id: 2, title: '本牌客戶結案原則', content: '一旦確認為本牌 (EIP 資料建立)，則該筆進攻對象結案。' },
  ],
  talkScripts: [
    { id: 3, title: '初次拜訪', content: '我們提供節能、高效率的設備，協助貴單位達成綠色採購目標。' },
    { id: 4, title: '設備汰換', content: '提供最新的複合機，搭配客製化維護合約，降低運營成本。' },
  ],
  uploadedMapUrl: 'https://drive.google.com/thumbnail?id=1OKORexYM2Ws-F3E3_SOu9sX4deqAsi0E&sz=w3000',
};

// Custom Hook for Excel Export
const useExcelExport = () => {
  useEffect(() => {
    if (typeof window.XLSX === 'undefined') {
      const script = document.createElement('script');
      script.src = 'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js';
      document.head.appendChild(script);
    }
  }, []);

  const exportToExcel = (data, filename, sheetName = 'Sheet1', columnHeaders) => {
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
          if (Array.isArray(value)) return value.map(item => typeof item === 'object' ? JSON.stringify(item) : item).join('; ');
          return value !== undefined ? value : '';
        })
      ),
    ];
    const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    ws['!cols'] = columnHeaders.map((h) => ({ wch: h.width || 20 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `${filename}_${new Date().toISOString().substring(0, 10)}.xlsx`);
  };
  return exportToExcel;
};

// --- MAIN COMPONENT ---
const App = () => {
  const [currentTab, setCurrentTab] = useState('targets');
  const [db, setDb] = useState(null);
  const [auth, setAuth] = useState(null);
  const [userId, setUserId] = useState(null);
  const [isLoading, setIsLoading] = useState(true);
  const [globalMessage, setGlobalMessage] = useState({ text: '', type: '' });
  const [isDownloadingMap, setIsDownloadingMap] = useState(false);

  // Restore missing state for editing
  const [editingUnitId, setEditingUnitId] = useState(null);
  const [isNewUnit, setIsNewUnit] = useState(false);

  // Separate State for Public Map Data
  const [publicMapData, setPublicMapData] = useState({
     areaMap: [] 
  });

  const [appData, setAppData] = useState({
    units: [],
    settings: initialSettings,
    schedules: [],
    meetings: [],
  });

  const exportToExcel = useExcelExport();

  // Load Tailwind
  useEffect(() => {
    const existingScript = document.querySelector('script[src="https://cdn.tailwindcss.com"]');
    if (!existingScript) {
      const script = document.createElement('script');
      script.src = 'https://cdn.tailwindcss.com';
      document.head.appendChild(script);
    }
  }, []);

  // Collection Refs
  const getUnitCollectionRef = useCallback((database) => collection(database, 'artifacts', appId, 'public', 'data', 'units'), []);
  // FIX: Added 'main' to make path segments even (6 segments)
  const getPublicMapDocRef = useCallback((database) => doc(database, 'artifacts', appId, 'public', 'data', 'map_config', 'main'), []);
  const getPrivateDocRef = useCallback((database, uid, collectionName, docId) => doc(database, 'artifacts', appId, 'users', uid, collectionName, docId), []);

  // Init Firebase
  useEffect(() => {
    try {
      if (Object.keys(firebaseConfig).length === 0) { setIsLoading(false); return; }
      const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApp();
      const database = getFirestore(app);
      const authentication = getAuth(app);
      setDb(database);
      setAuth(authentication);

      const initAuth = async () => {
        try {
          if (initialAuthToken) await signInWithCustomToken(authentication, initialAuthToken);
          else await signInAnonymously(authentication);
        } catch (e) {
          console.error('Auth failed:', e);
          setGlobalMessage({ text: '驗證失敗', type: 'error' });
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
      console.error('Firebase init failed:', error);
      setIsLoading(false);
    }
  }, []);

  // Listeners
  useEffect(() => {
    if (!db || !userId) return;

    // 1. Public Units
    const unsubscribeUnits = onSnapshot(getUnitCollectionRef(db), (snapshot) => {
      const units = snapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
        equipment: safeParse(doc.data().equipment || '[]'),
        history: safeParse(doc.data().history || '[]'),
        characteristics: doc.data().characteristics || [],
      }));
      setAppData((prev) => ({ ...prev, units }));
    });

    // 2. Private Settings (Schedule, Meetings, Local Settings)
    const settingsDocRef = getPrivateDocRef(db, userId, 'settings', 'params');
    const unsubscribeSettings = onSnapshot(settingsDocRef, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        setAppData((prev) => ({
          ...prev,
          settings: {
            ...initialSettings,
            ...data,
            uploadedMapUrl: initialSettings.uploadedMapUrl, 
          },
          schedules: data.schedules || [],
          meetings: data.meetings || [],
        }));
      } else {
        setDoc(settingsDocRef, { ...initialSettings, schedules: [], meetings: [] }).catch(console.error);
      }
    });

    // 3. PUBLIC MAP DATA
    const publicMapDocRef = getPublicMapDocRef(db);
    const unsubscribeMap = onSnapshot(publicMapDocRef, (docSnap) => {
        if (docSnap.exists()) {
            setPublicMapData(docSnap.data());
        } else {
            setDoc(publicMapDocRef, { areaMap: [] });
        }
    });

    return () => {
      unsubscribeUnits();
      unsubscribeSettings();
      unsubscribeMap();
    };
  }, [db, userId]);

  // Update Functions
  const updatePrivateData = async (fields) => {
    if (!db || !userId) return;
    try {
      await updateDoc(getPrivateDocRef(db, userId, 'settings', 'params'), fields);
      setGlobalMessage({ text: '私有資料更新成功', type: 'success' });
    } catch (e) { console.error(e); }
  };

  const updatePublicMap = async (newAreaMap) => {
      if (!db) return;
      try {
          await updateDoc(getPublicMapDocRef(db), { areaMap: newAreaMap });
          setGlobalMessage({ text: '地圖區域已更新至公用資料庫', type: 'success' });
      } catch (e) {
          console.error("Failed to update public map", e);
          setGlobalMessage({ text: '地圖更新失敗', type: 'error' });
      }
  };

  const updateUnit = async (id, data) => {
    if (!db) return;
    const docRef = doc(getUnitCollectionRef(db), id);
    const updateData = { ...data };
    if (Array.isArray(data.equipment)) updateData.equipment = safeStringify(data.equipment);
    if (Array.isArray(data.history)) updateData.history = safeStringify(data.history);
    await updateDoc(docRef, updateData);
  };

  const deleteUnits = async (ids) => {
    if (!db) return;
    await Promise.all(ids.map((id) => deleteDoc(doc(getUnitCollectionRef(db), id))));
  };

  // --- Components ---
  const LoadingState = () => (
    <div className="flex flex-col items-center justify-center h-screen bg-slate-50 text-slate-500">
      <Loader className="w-12 h-12 animate-spin text-indigo-600 mb-4" />
      <p className="text-lg font-medium text-slate-700">正在連線至台大攻略戰情室...</p>
    </div>
  );

  // --- Tab 3: Targets & Map ---
  const Tab3TargetsMap = () => {
    const { units, settings } = appData;
    const { uploadedMapUrl } = settings;
    const areaMap = publicMapData.areaMap || []; 

    const [filter, setFilter] = useState({ id: '', type: '', name: '', contact: '', phone: '', brand: '', model: '' });
    const [selectedUnitIds, setSelectedUnitIds] = useState([]);

    const filteredUnits = useMemo(() => {
      return units.filter((unit) => {
        const equipmentJson = safeParse(unit.equipment);
        const hasMatchingEquipment = (filter.brand || filter.model)
          ? equipmentJson.some((eq) => (filter.brand === '' || eq.brand.includes(filter.brand)) && (filter.model === '' || eq.model.includes(filter.model)))
          : true;
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
        const x = (e.nativeEvent.offsetX / e.currentTarget.offsetWidth) * 100;
        const y = (e.nativeEvent.offsetY / e.currentTarget.offsetHeight) * 100;
        setPolyPoints(prev => [...prev, {x, y}]);
    };

    const handleConfirmArea = async () => {
        if (!newAreaCode.trim()) { alert("請輸入區域編號"); return; }
        const newArea = {
            id: crypto.randomUUID(),
            code: newAreaCode,
            type: 'polygon',
            points: polyPoints,
        };
        await updatePublicMap([...areaMap, newArea]);
        
        setMapState({ isDrawing: false, isNaming: false });
        setPolyPoints([]);
        setNewAreaCode('');
    };

    const handleDownloadHighResMap = async () => {
        if (!uploadedMapUrl) return;
        setIsDownloadingMap(true);

        try {
            const img = new Image();
            img.crossOrigin = "Anonymous";
            img.src = uploadedMapUrl;

            await new Promise((resolve, reject) => {
                img.onload = resolve;
                img.onerror = () => reject(new Error("Image load failed"));
            });

            const canvas = document.createElement('canvas');
            canvas.width = img.naturalWidth;
            canvas.height = img.naturalHeight;
            const ctx = canvas.getContext('2d');

            ctx.drawImage(img, 0, 0);

            areaMap.forEach(area => {
                if (area.points && area.points.length > 0) {
                    ctx.beginPath();
                    const startX = (area.points[0].x / 100) * canvas.width;
                    const startY = (area.points[0].y / 100) * canvas.height;
                    ctx.moveTo(startX, startY);
                    area.points.forEach((p, i) => {
                        if (i > 0) ctx.lineTo((p.x / 100) * canvas.width, (p.y / 100) * canvas.height);
                    });
                    ctx.closePath();
                    
                    ctx.fillStyle = "rgba(239, 68, 68, 0.3)";
                    ctx.fill();
                    ctx.lineWidth = Math.max(3, canvas.width * 0.003); 
                    ctx.strokeStyle = "#dc2626";
                    ctx.stroke();

                    const centerX = (area.points.reduce((s, p) => s + p.x, 0) / area.points.length / 100) * canvas.width;
                    const centerY = (area.points.reduce((s, p) => s + p.y, 0) / area.points.length / 100) * canvas.height;

                    const fontSize = Math.max(16, canvas.width * 0.015);
                    ctx.font = `bold ${fontSize}px "Noto Sans TC", sans-serif`;
                    
                    const text = area.code;
                    const textMetrics = ctx.measureText(text);
                    const padding = fontSize * 0.6;
                    const boxWidth = textMetrics.width + padding * 2;
                    const boxHeight = fontSize * 1.4;

                    ctx.fillStyle = "#dc2626";
                    ctx.fillRect(centerX - boxWidth/2, centerY - boxHeight/2, boxWidth, boxHeight);

                    ctx.fillStyle = "white";
                    ctx.textAlign = "center";
                    ctx.textBaseline = "middle";
                    ctx.fillText(text, centerX, centerY);
                }
            });

            const dataUrl = canvas.toDataURL("image/jpeg", 0.9);
            const link = document.createElement('a');
            link.download = `NTU_Strategy_Map_HighRes_${new Date().toISOString().slice(0,10)}.jpg`;
            link.href = dataUrl;
            link.click();

        } catch (e) {
            console.error(e);
            alert("下載失敗：可能是圖片跨域權限(CORS)問題。請嘗試使用電腦截圖功能。");
        } finally {
            setIsDownloadingMap(false);
        }
    };

    const totalUnits = units.length;
    const currentClients = units.filter(u => u.attackStatus === 'client').length;
    const incompleteUnits = units.filter(u => !checkUnitCompleteness(u));
    const adminUnits = units.filter(u => u.category === 'Administrative').length;
    const academicUnits = units.filter(u => u.category === 'Academic').length;
    
    return (
      <div className="space-y-8 p-6 max-w-7xl mx-auto">
        <div className="flex justify-between items-center mb-6">
          <div className="flex items-center space-x-3">
            <MapPin className="w-8 h-8 text-indigo-600" />
            <h2 className="text-3xl font-extrabold text-slate-800">戰情地圖與對象總覽</h2>
          </div>
          {incompleteUnits.length > 0 && (
            <button className="flex items-center space-x-2 bg-rose-100 text-rose-700 px-4 py-2 rounded-lg font-bold animate-pulse">
                <AlertCircle className="w-5 h-5" /><span>{incompleteUnits.length} 筆資料待補齊</span>
            </button>
          )}
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-6">
             <StatusCard title="總家數" value={totalUnits} gradient="from-indigo-500 to-purple-600" icon={<Building className="w-6 h-6 text-white" />} />
             <StatusCard title="本牌家數" value={currentClients} gradient="from-emerald-500 to-teal-500" icon={<CheckCircle className="w-6 h-6 text-white" />} />
             <StatusCard title="行政單位" value={adminUnits} gradient="from-orange-400 to-red-500" icon={<Users className="w-6 h-6 text-white" />} />
             <StatusCard title="學術單位" value={academicUnits} gradient="from-sky-500 to-blue-600" icon={<Users className="w-6 h-6 text-white" />} />
        </div>

        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100 flex flex-col h-[800px] relative">
            {mapState.isNaming && (
                <div className="absolute inset-0 z-[100] bg-black/40 flex items-center justify-center backdrop-blur-sm">
                    <div className="bg-white p-6 rounded-2xl shadow-2xl w-96 border border-slate-200">
                        <h3 className="text-lg font-bold text-slate-800 mb-4">確認區域編號</h3>
                        <input 
                            type="text" autoFocus value={newAreaCode} onChange={(e) => setNewAreaCode(e.target.value)}
                            className={`${styles.formInput} text-lg font-bold text-center tracking-widest`} placeholder="例如: A-01"
                        />
                        <div className="flex justify-end space-x-3 mt-6">
                            <button onClick={() => { setMapState({isDrawing: false, isNaming: false}); setPolyPoints([]); }} className="px-4 py-2 text-slate-500">放棄</button>
                            <button onClick={handleConfirmArea} className={styles.btnPrimary}>確認建立</button>
                        </div>
                    </div>
                </div>
            )}

            <div className="p-4 bg-slate-800 text-white flex justify-between items-center z-10 shadow-md">
                <div className="flex items-center space-x-4">
                     <h3 className="text-xl font-bold flex items-center"><MapPin className="w-5 h-5 mr-2" /> 校園地圖戰情室</h3>
                     <span className="text-xs bg-slate-700 px-2 py-1 rounded text-slate-300 border border-slate-600 flex items-center">
                        <UploadCloud className="w-3 h-3 mr-1"/> 資料已同步至公用雲端
                     </span>
                </div>
                
                <div className="flex space-x-3 items-center">
                    <div className="flex items-center bg-slate-700 rounded-lg p-1 border border-slate-600">
                        <button onClick={handleZoomOut} className="p-1.5 hover:bg-slate-600 rounded"><ZoomOut className="w-4 h-4"/></button>
                        <span className="text-xs font-mono w-12 text-center">{Math.round(zoom * 100)}%</span>
                        <button onClick={handleZoomIn} className="p-1.5 hover:bg-slate-600 rounded"><ZoomIn className="w-4 h-4"/></button>
                    </div>

                    {!mapState.isDrawing ? (
                        <>
                            <button onClick={() => setMapState({ isDrawing: true, isNaming: false })} className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 rounded-lg flex items-center text-sm font-bold shadow-lg shadow-indigo-500/30 transition">
                                <Edit className="w-4 h-4 mr-2"/> 圈選區域
                            </button>
                            <button onClick={handleDownloadHighResMap} disabled={isDownloadingMap} className="px-4 py-2 bg-sky-600 hover:bg-sky-500 rounded-lg flex items-center text-sm font-bold shadow-lg shadow-sky-500/30 transition disabled:opacity-50">
                                {isDownloadingMap ? <Loader className="w-4 h-4 mr-2 animate-spin"/> : <Camera className="w-4 h-4 mr-2"/>}
                                下載原始合成圖
                            </button>
                        </>
                    ) : (
                        <>
                           <button onClick={() => setPolyPoints(p => p.slice(0, -1))} className="px-3 py-2 bg-slate-600 hover:bg-slate-500 rounded-lg"><Undo className="w-4 h-4"/></button>
                           <button onClick={() => { setMapState({isDrawing:false, isNaming:false}); setPolyPoints([]);}} className="px-3 py-2 bg-slate-600 hover:bg-slate-500 rounded-lg">取消</button>
                           <button onClick={() => { if(polyPoints.length < 3) return alert('請至少標記3點'); setMapState(p => ({...p, isNaming: true})); }} className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 rounded-lg font-bold animate-pulse">完成圈選</button>
                        </>
                    )}
                </div>
            </div>

            <div className="flex-grow overflow-auto bg-slate-200 relative cursor-move flex items-center justify-center">
                <div className="relative origin-center transition-transform duration-200 ease-out shadow-2xl" style={{ transform: `scale(${zoom})` }}>
                    <div className={`relative inline-block ${mapState.isDrawing ? 'cursor-crosshair' : ''}`} onClick={handleMapClick}>
                        <img src={uploadedMapUrl} alt="Map" className="block max-w-none pointer-events-none select-none" referrerPolicy="no-referrer" />
                        
                        <svg className="absolute inset-0 w-full h-full pointer-events-none" viewBox="0 0 100 100" preserveAspectRatio="none">
                            {areaMap.map(area => {
                                let pointsStr = "";
                                if (area.points) pointsStr = area.points.map(p => `${p.x},${p.y}`).join(" ");
                                else { 
                                    const {x1=0, y1=0, x2=0, y2=0} = area;
                                    pointsStr = `${Math.min(x1,x2)},${Math.min(y1,y2)} ${Math.max(x1,x2)},${Math.min(y1,y2)} ${Math.max(x1,x2)},${Math.max(y1,y2)} ${Math.min(x1,x2)},${Math.max(y1,y2)}`;
                                }
                                return (
                                    <polygon key={area.id} points={pointsStr} fill="rgba(220, 38, 38, 0.25)" stroke="#dc2626" strokeWidth="0.4" className="transition hover:fill-red-500/50" />
                                );
                            })}
                            {polyPoints.length > 0 && (
                                <g>
                                    <polyline points={polyPoints.map(p => `${p.x},${p.y}`).join(" ")} fill="rgba(220, 38, 38, 0.2)" stroke="#dc2626" strokeWidth="0.4" strokeDasharray="1 1" />
                                    {polyPoints.map((p, i) => <circle key={i} cx={p.x} cy={p.y} r="0.4" fill="white" stroke="red" strokeWidth="0.1" />)}
                                </g>
                            )}
                        </svg>

                        <div className="absolute inset-0 pointer-events-none">
                            {areaMap.map(area => {
                                let cx = 0, cy = 0;
                                if(area.points) {
                                    cx = area.points.reduce((s,p)=>s+p.x,0)/area.points.length;
                                    cy = area.points.reduce((s,p)=>s+p.y,0)/area.points.length;
                                } else { 
                                    cx = (area.x1+area.x2)/2; cy = (area.y1+area.y2)/2;
                                }
                                const unitCount = units.filter(u => u.areaCode === area.code).length;
                                return (
                                    <div key={area.id} style={{ left: `${cx}%`, top: `${cy}%`, transform: `translate(-50%, -50%) scale(${1/zoom})` }} className="absolute flex flex-col items-center z-20 pointer-events-auto group">
                                         <span className="bg-red-600 text-white text-xs px-2 py-1 rounded font-bold border border-white shadow-sm whitespace-nowrap">
                                            {area.code} {unitCount > 0 && `(${unitCount})`}
                                         </span>
                                         <button 
                                            onClick={(e) => { e.stopPropagation(); if(confirm(`刪除區域 ${area.code}? (此操作會同步至公用資料庫)`)) updatePublicMap(areaMap.filter(a => a.id !== area.id)); }}
                                            className="mt-1 p-1 bg-white text-red-600 rounded-full shadow hover:bg-red-50 opacity-0 group-hover:opacity-100 transition"
                                         >
                                            <X className="w-3 h-3"/>
                                         </button>
                                    </div>
                                );
                            })}
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100 p-6">
             <div className="flex justify-between items-center mb-4">
                 <h3 className="text-xl font-bold text-slate-800">進攻對象列表</h3>
                 <div className="flex space-x-2">
                     <button onClick={() => deleteUnits(selectedUnitIds)} disabled={selectedUnitIds.length === 0} className={styles.btnDanger}><Trash2 className="w-4 h-4 mr-1"/> 刪除</button>
                     <button onClick={() => exportToExcel(filteredUnits, 'UnitList', 'Units', [{key:'name', label:'名稱'}, {key:'areaCode', label:'區域'}])} className={styles.btnInfo}><Download className="w-4 h-4 mr-1"/> 匯出</button>
                 </div>
             </div>
             <UnitTable units={filteredUnits} selectedUnitIds={selectedUnitIds} setSelectedUnitIds={setSelectedUnitIds} setEditingUnitId={setEditingUnitId} setCurrentTab={setCurrentTab} setIsNewUnit={setIsNewUnit} />
        </div>
      </div>
    );
  };

  const Tab1Calendar = () => <div className="p-10 text-center text-gray-500">行事曆功能區塊 (請參考完整版實作)</div>;
  const Tab2Guidelines = () => <div className="p-10 text-center text-gray-500">攻擊準則功能區塊 (請參考完整版實作)</div>;
  const Tab4Record = () => <div className="p-10 text-center text-gray-500">拜訪紀錄功能區塊 (請參考完整版實作)</div>;
  const Tab5Settings = () => <div className="p-10 text-center text-gray-500">參數設定功能區塊 (請參考完整版實作)</div>;

  const renderTabContent = () => {
    switch (currentTab) {
        case 'targets': return <Tab3TargetsMap />;
        case 'calendar': return <Tab1Calendar />;
        case 'record': return <Tab4Record />;
        case 'guidelines': return <Tab2Guidelines />;
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
      <header className="sticky top-0 z-50 backdrop-blur-md bg-white/80 border-b border-slate-200 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div className="flex justify-between items-center h-16">
                 <div className="flex items-center space-x-2">
                     <div className="bg-indigo-600 text-white p-2 rounded-lg"><Activity className="w-6 h-6"/></div>
                     <h1 className="text-xl font-bold text-slate-800">2026 台大攻略戰情室</h1>
                 </div>
                 <div className="text-xs text-slate-500 font-mono flex items-center">
                    <div className={`w-2 h-2 rounded-full mr-2 ${userId ? 'bg-emerald-500' : 'bg-gray-300'}`}></div>
                    ID: {userId ? String(userId).substring(0, 8) : '...'}
                 </div>
            </div>
            <nav className="flex space-x-1 overflow-x-auto pb-1 no-scrollbar">
            {navItems.map((item) => (
                <button key={item.id} onClick={() => setCurrentTab(item.id)}
                className={`relative px-5 py-3 text-sm font-medium transition-all duration-300 rounded-t-lg flex items-center space-x-2 whitespace-nowrap ${currentTab === item.id ? 'text-indigo-600 bg-indigo-50/50' : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'}`}>
                {item.icon}<span>{item.label}</span>{currentTab === item.id && <span className="absolute bottom-0 left-0 w-full h-0.5 bg-indigo-600 rounded-t-full" />}
                </button>
            ))}
            </nav>
        </div>
      </header>
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

const StatusCard = ({ title, value, icon, gradient }) => (
    <div className={`relative p-6 rounded-2xl shadow-lg text-white bg-gradient-to-br ${gradient} overflow-hidden`}>
      <div className="absolute top-0 right-0 p-4 opacity-20 transform scale-150">{icon}</div>
      <p className="text-sm font-medium opacity-90 tracking-wide">{title}</p>
      <p className="text-4xl font-extrabold mt-2 tracking-tight">{value}</p>
    </div>
);

const UnitTable = ({ units, selectedUnitIds, setSelectedUnitIds, setEditingUnitId, setCurrentTab, setIsNewUnit }) => (
    <div className="overflow-x-auto">
        <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
                <tr>
                    <th className="p-3 text-left">選取</th><th className="p-3 text-left">名稱</th><th className="p-3 text-left">區域</th><th className="p-3 text-left">狀態</th><th className="p-3 text-right">操作</th>
                </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-100">
                {units.map(u => (
                    <tr key={u.id}>
                        <td className="p-3"><input type="checkbox" checked={selectedUnitIds.includes(u.id)} onChange={() => setSelectedUnitIds(p => p.includes(u.id)?p.filter(i=>i!==u.id):[...p, u.id])} className={styles.checkbox}/></td>
                        <td className="p-3 font-medium">{u.name}</td>
                        <td className="p-3"><span className="bg-slate-100 px-2 py-1 rounded text-xs font-bold">{u.areaCode}</span></td>
                        <td className="p-3 text-xs">{u.attackStatus}</td>
                        <td className="p-3 text-right"><button onClick={() => {setEditingUnitId(u.id); setCurrentTab('record');}} className="text-indigo-600 hover:underline text-xs">編輯</button></td>
                    </tr>
                ))}
            </tbody>
        </table>
    </div>
);

export default App;