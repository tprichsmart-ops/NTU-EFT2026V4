import React, { useState, useEffect, useCallback, useMemo } from 'react';
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
  Link as LinkIcon
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
// Automatically converts Google Drive sharing links to embeddable image links
const getEmbeddableMapUrl = (url) => {
  if (!url) return '';
  // Regex to extract File ID from common Google Drive link formats
  const driveRegex = /(?:file\/d\/|id=|open\?id=)([-w]{25,})/;
  const match = url.match(driveRegex);
  
  if (match && match[1]) {
    // Convert to the export=view format which works for <img> tags (if public)
    return `https://drive.google.com/uc?export=view&id=${match[1]}`;
  }
  return url;
};

// --- Data Validation Helper ---
const checkUnitCompleteness = (unit) => {
  // Define required fields
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
  
  // Headers with explicit instructions in styling/text
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
  
  // Set column widths
  ws['!cols'] = [
    { wch: 20 }, { wch: 25 }, { wch: 30 }, { wch: 15 }, 
    { wch: 35 }, { wch: 15 }, { wch: 15 }, { wch: 10 }
  ];

  // --- Add Data Validation (Dropdowns) ---
  // Range: Start from row 2 (index 1) to row 100 (index 99)
  const range = { s: { r: 1, c: 0 }, e: { r: 100, c: 7 } };
  
  if (!ws['!dataValidation']) ws['!dataValidation'] = [];

  // Column B (Index 1): 單位類別
  for (let r = range.s.r; r <= range.e.r; ++r) {
    ws['!dataValidation'].push({
      sqref: window.XLSX.utils.encode_cell({r: r, c: 1}),
      type: 'list',
      operator: 'equal',
      formula1: '"行政,學術"', 
      showErrorMessage: true,
      errorTitle: '輸入錯誤',
      error: '請從下拉選單中選擇：行政 或 學術'
    });
  }

  // Column C (Index 2): 獨立空間分組
  for (let r = range.s.r; r <= range.e.r; ++r) {
    ws['!dataValidation'].push({
      sqref: window.XLSX.utils.encode_cell({r: r, c: 2}),
      type: 'list',
      operator: 'equal',
      formula1: '"獨立空間,一般"',
      showErrorMessage: true,
      errorTitle: '輸入錯誤',
      error: '請從下拉選單中選擇：獨立空間 或 一般'
    });
  }

  // Column E (Index 4): 進攻狀態
  for (let r = range.s.r; r <= range.e.r; ++r) {
    ws['!dataValidation'].push({
      sqref: window.XLSX.utils.encode_cell({r: r, c: 4}),
      type: 'list',
      operator: 'equal',
      formula1: '"進攻中,已進攻暫定結案,本牌客戶"',
      showErrorMessage: true,
      errorTitle: '輸入錯誤',
      error: '請從下拉選單中選擇有效狀態'
    });
  }

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

// --- Initial Data Structures for Defaults ---

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
  // Hardcoded to the specific Google Drive image provided
  uploadedMapUrl: 'https://drive.google.com/uc?export=view&id=1OKORexYM2Ws-F3E3_SOu9sX4deqAsi0E',
};

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

      // FIX: Check if app is already initialized to prevent "duplicate app" errors
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
        if (user) {
          setUserId(user.uid);
        } else {
          setUserId(null);
        }
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
      (error) => {
        console.error('Error listening to units:', error);
      }
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
              // Only override with saved data if it's NOT the uploadedMapUrl,
              // or optionally, use the saved one. 
              // Since the user asked to "directly embed", we prioritize the hardcoded one 
              // for this session or ensure initialSettings.uploadedMapUrl is used if no custom one exists.
              // However, to strictly follow "directly embed", we can force it here or just let initialSettings take precedence if DB is empty.
              // But if the DB has an old URL, it might overwrite. 
              // Let's force the hardcoded URL for the map to ensure the request is met even if old data exists.
              ...data,
              uploadedMapUrl: initialSettings.uploadedMapUrl, // Force hardcoded URL
              guidelines: data.guidelines || initialSettings.guidelines,
              talkScripts: data.talkScripts || initialSettings.talkScripts,
              areaMap: data.areaMap || initialSettings.areaMap,
            },
            schedules: data.schedules || [],
            meetings: data.meetings || [],
          }));
        } else {
          setDoc(settingsDocRef, {
            ...initialSettings,
            schedules: initialSettings.schedules || [],
            meetings: initialSettings.meetings || [],
          }).catch((e) =>
            console.error('Error setting initial private data:', e)
          );
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
      setGlobalMessage({ text: '資料更新成功！', type: 'success' });
    } catch (e) {
      console.error('Error updating private data:', e);
      if (e.code === 'permission-denied') {
        setGlobalMessage({ text: '權限不足或檔案過大 (限制 1MB)', type: 'error' });
      } else {
        setGlobalMessage({ text: `資料更新失敗: ${e.message}`, type: 'error' });
      }
    }
  };

  const updateUnit = async (id, data) => {
    if (!db || !userId) return;
    try {
      const docRef = doc(getUnitCollectionRef(db), id);
      const updateData = {};
      Object.keys(data).forEach((key) => {
        if (
          ['equipment', 'history'].includes(key) &&
          Array.isArray(data[key])
        ) {
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
      await Promise.all(
        ids.map((id) => deleteDoc(doc(getUnitCollectionRef(db), id)))
      );
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

  // --- Tab 1: 進攻行事曆 ---
  const Tab1Calendar = () => {
    const totalUnits = appData.units.length;
    const currentClients = appData.units.filter(
      (u) => u.attackStatus === 'client'
    ).length;
    const settledNonClients = appData.units.filter(
      (u) => u.attackStatus === 'settled_non_client'
    ).length;

    const [isAddingSchedule, setIsAddingSchedule] = useState(false);
    const [isAddingMeeting, setIsAddingMeeting] = useState(false);
    const [scheduleCollapsed, setScheduleCollapsed] = useState(true);
    const [selectedScheduleIds, setSelectedScheduleIds] = useState([]);
    const [selectedMeetingIds, setSelectedMeetingIds] = useState([]);

    const scheduleTableData = useMemo(() => {
      const now = new Date();
      return appData.schedules
        .map((s) => ({
          ...s,
          isExpired: s.scheduleDate && new Date(s.scheduleDate) < now,
        }))
        .sort((a, b) => new Date(a.scheduleDate) - new Date(b.scheduleDate));
    }, [appData.schedules]);

    const filteredSchedules = scheduleTableData.filter((s) =>
      scheduleCollapsed ? !s.isExpired : true
    );

    const exportSchedules = () => {
      const headers = [
        { key: 'scheduleDate', label: '時程', width: 15 },
        { key: 'area', label: '區域', width: 10 },
        { key: 'personnel', label: '人數/家數/區域', width: 30 },
        { key: 'memo', label: '待辦&備忘', width: 40 },
        { key: 'resourceContent', label: '資源(配給內容)', width: 20 },
        { key: 'resourceMonth', label: '配給月份', width: 15 },
        { key: 'resourceAmount', label: '配給金額', width: 15 },
        { key: 'isExpired', label: '是否過期', width: 10 },
      ];
      exportToExcel(scheduleTableData, '進攻行事曆-排程', '排程', headers);
    };

    const exportMeetings = () => {
      const headers = [
        { key: 'date', label: '日期', width: 15 },
        { key: 'attendees', label: '與會人員', width: 20 },
        { key: 'summary', label: '總結', width: 40 },
        { key: 'todo', label: '待辦', width: 40 },
        { key: 'nextMeetingDate', label: '下次開會時間', width: 15 },
        { key: 'nextAttendees', label: '下次與會人員', width: 20 },
        { key: 'nextTopics', label: '下次議題', width: 30 },
      ];
      exportToExcel(
        appData.meetings,
        '進攻行事曆-會議紀錄',
        '戰勤會議紀錄',
        headers
      );
    };

    const deleteSelectedSchedules = () => {
      if (selectedScheduleIds.length === 0) return;
      const updatedSchedules = appData.schedules.filter(
        (s) => !selectedScheduleIds.includes(s.id)
      );
      updatePrivateData({ schedules: updatedSchedules });
      setSelectedScheduleIds([]);
    };

    const deleteSelectedMeetings = () => {
      if (selectedMeetingIds.length === 0) return;
      const updatedMeetings = appData.meetings.filter(
        (m) => !selectedMeetingIds.includes(m.id)
      );
      updatePrivateData({ meetings: updatedMeetings });
      setSelectedMeetingIds([]);
    };

    const ScheduleRow = ({ schedule, index, isEditing, setIsEditing }) => {
      const [editData, setEditData] = useState(schedule);

      const handleSave = () => {
        const updatedSchedules = appData.schedules.map((s) =>
          s.id === schedule.id ? editData : s
        );
        updatePrivateData({ schedules: updatedSchedules });
        setIsEditing(null);
      };

      return (
        <tr
          className={`border-b transition hover:bg-indigo-50/50 ${
            schedule.isExpired ? 'bg-gray-50 opacity-60' : 'bg-white'
          }`}
        >
          <td className="p-3">
            <input
              type="checkbox"
              checked={selectedScheduleIds.includes(schedule.id)}
              onChange={() => {
                setSelectedScheduleIds((prev) =>
                  prev.includes(schedule.id)
                    ? prev.filter((id) => id !== schedule.id)
                    : [...prev, schedule.id]
                );
              }}
              className={styles.checkbox}
            />
          </td>
          <td className="p-3 text-sm font-medium text-gray-900">
            {isEditing === schedule.id ? (
              <input
                type="date"
                value={editData.scheduleDate || ''}
                onChange={(e) =>
                  setEditData((p) => ({ ...p, scheduleDate: e.target.value }))
                }
                className={styles.formInput}
              />
            ) : (
              schedule.scheduleDate
            )}
          </td>
          <td className="p-3 text-sm text-gray-700 whitespace-nowrap">
            {isEditing === schedule.id ? (
              <input
                type="text"
                value={editData.personnel || ''}
                onChange={(e) =>
                  setEditData((p) => ({ ...p, personnel: e.target.value }))
                }
                className={styles.formInput}
                placeholder="人員/家數/區域"
              />
            ) : (
              `${schedule.personnel} / ${schedule.area || 'N/A'}`
            )}
          </td>
          <td className="p-3 text-sm text-gray-700">
            {isEditing === schedule.id ? (
              <textarea
                value={editData.memo || ''}
                onChange={(e) =>
                  setEditData((p) => ({ ...p, memo: e.target.value }))
                }
                className={styles.formTextarea}
                placeholder="待辦&備忘"
              />
            ) : (
              schedule.memo
            )}
          </td>
          <td className="p-3 text-sm text-gray-700">
            {isEditing === schedule.id ? (
              <input
                type="text"
                value={editData.resourceContent || ''}
                onChange={(e) =>
                  setEditData((p) => ({
                    ...p,
                    resourceContent: e.target.value,
                  }))
                }
                className={styles.formInput}
                placeholder="配給內容"
              />
            ) : (
              editData.resourceContent
            )}
          </td>
          <td className="p-3 text-sm text-gray-700 whitespace-nowrap">
            {isEditing === schedule.id ? (
              <input
                type="month"
                value={editData.resourceMonth || ''}
                onChange={(e) =>
                  setEditData((p) => ({ ...p, resourceMonth: e.target.value }))
                }
                className={styles.formInput}
              />
            ) : (
              editData.resourceMonth
            )}
          </td>
          <td className="p-3 text-sm text-gray-700">
            {isEditing === schedule.id ? (
              <input
                type="number"
                value={editData.resourceAmount || 0}
                onChange={(e) =>
                  setEditData((p) => ({
                    ...p,
                    resourceAmount: parseInt(e.target.value) || 0,
                  }))
                }
                className={styles.formInput}
              />
            ) : (
              <span className="font-mono text-emerald-600 font-bold">
                ${(editData.resourceAmount || 0).toLocaleString()}
              </span>
            )}
          </td>
          <td className="p-3 text-right whitespace-nowrap">
            {isEditing === schedule.id ? (
              <button
                onClick={handleSave}
                className="p-2 text-emerald-600 hover:text-emerald-800 bg-emerald-50 rounded-full hover:bg-emerald-100 transition"
              >
                <Save className="w-4 h-4" />
              </button>
            ) : (
              <button
                onClick={() => setIsEditing(schedule.id)}
                className="p-2 text-indigo-600 hover:text-indigo-800 bg-indigo-50 rounded-full hover:bg-indigo-100 transition"
              >
                <Edit className="w-4 h-4" />
              </button>
            )}
          </td>
        </tr>
      );
    };

    const MeetingRow = ({ meeting, index, isEditing, setIsEditing }) => {
      const [editData, setEditData] = useState(meeting);

      const handleSave = () => {
        const updatedMeetings = appData.meetings.map((m) =>
          m.id === meeting.id ? editData : m
        );
        updatePrivateData({ meetings: updatedMeetings });
        setIsEditing(null);
      };

      return (
        <tr className="bg-white border-b hover:bg-indigo-50/50 transition">
          <td className="p-3">
            <input
              type="checkbox"
              checked={selectedMeetingIds.includes(meeting.id)}
              onChange={() => {
                setSelectedMeetingIds((prev) =>
                  prev.includes(meeting.id)
                    ? prev.filter((id) => id !== meeting.id)
                    : [...prev, meeting.id]
                );
              }}
              className={styles.checkbox}
            />
          </td>
          <td className="p-3 text-sm font-medium text-gray-900">
            {isEditing === meeting.id ? (
              <input
                type="date"
                value={editData.date || ''}
                onChange={(e) =>
                  setEditData((p) => ({ ...p, date: e.target.value }))
                }
                className={styles.formInput}
              />
            ) : (
              meeting.date
            )}
          </td>
          <td className="p-3 text-sm text-gray-700">
            {isEditing === meeting.id ? (
              <input
                type="text"
                value={editData.attendees || ''}
                onChange={(e) =>
                  setEditData((p) => ({ ...p, attendees: e.target.value }))
                }
                className={styles.formInput}
              />
            ) : (
              editData.attendees
            )}
          </td>
          <td className="p-3 text-sm text-gray-700">
            {isEditing === meeting.id ? (
              <textarea
                value={editData.summary || ''}
                onChange={(e) =>
                  setEditData((p) => ({ ...p, summary: e.target.value }))
                }
                className={styles.formTextarea}
              />
            ) : (
              editData.summary
            )}
          </td>
          <td className="p-3 text-sm text-gray-700">
            {isEditing === meeting.id ? (
              <textarea
                value={editData.todo || ''}
                onChange={(e) =>
                  setEditData((p) => ({ ...p, todo: e.target.value }))
                }
                className={styles.formTextarea}
              />
            ) : (
              editData.todo
            )}
          </td>
          <td className="p-3 text-sm text-gray-700 whitespace-nowrap">
            {isEditing === meeting.id ? (
              <input
                type="date"
                value={editData.nextMeetingDate || ''}
                onChange={(e) =>
                  setEditData((p) => ({
                    ...p,
                    nextMeetingDate: e.target.value,
                  }))
                }
                className={styles.formInput}
              />
            ) : (
              editData.nextMeetingDate
            )}
          </td>
          <td className="p-3 text-right whitespace-nowrap">
            {isEditing === meeting.id ? (
              <button
                onClick={handleSave}
                className="p-2 text-emerald-600 hover:text-emerald-800 bg-emerald-50 rounded-full hover:bg-emerald-100 transition"
              >
                <Save className="w-4 h-4" />
              </button>
            ) : (
              <button
                onClick={() => setIsEditing(meeting.id)}
                className="p-2 text-indigo-600 hover:text-indigo-800 bg-indigo-50 rounded-full hover:bg-indigo-100 transition"
              >
                <Edit className="w-4 h-4" />
              </button>
            )}
          </td>
        </tr>
      );
    };

    const AddScheduleForm = () => {
      const [newSchedule, setNewSchedule] = useState({
        id: crypto.randomUUID(),
        scheduleDate: new Date().toISOString().substring(0, 10),
        personnel: '',
        area: '',
        memo: '',
        resourceContent: '',
        resourceMonth: new Date().toISOString().substring(0, 7),
        resourceAmount: 0,
      });

      const handleAdd = () => {
        if (!newSchedule.scheduleDate || !newSchedule.personnel) {
          alert('請填寫時程與人數/家數/區域。');
          return;
        }
        updatePrivateData({ schedules: [...appData.schedules, newSchedule] });
        setIsAddingSchedule(false);
      };

      return (
        <div className="p-6 border border-amber-200 rounded-xl bg-amber-50/50 backdrop-blur-sm shadow-inner space-y-4 mb-6">
          <h4 className="font-bold text-lg text-amber-800 flex items-center">
            <Plus className="w-5 h-5 mr-2" /> 新增排程
          </h4>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <input
              type="date"
              value={newSchedule.scheduleDate}
              onChange={(e) =>
                setNewSchedule((p) => ({ ...p, scheduleDate: e.target.value }))
              }
              className={styles.formInput}
              title="時程"
            />
            <input
              type="text"
              value={newSchedule.personnel}
              onChange={(e) =>
                setNewSchedule((p) => ({ ...p, personnel: e.target.value }))
              }
              className={styles.formInput}
              placeholder="人員/家數/區域 (e.g. Joe/5家/A區)"
            />
            <input
              type="text"
              value={newSchedule.memo}
              onChange={(e) =>
                setNewSchedule((p) => ({ ...p, memo: e.target.value }))
              }
              className={`${styles.formInput} md:col-span-2`}
              placeholder="待辦&備忘"
            />
          </div>
          <div className="flex flex-wrap gap-2 items-center bg-white/50 p-3 rounded-lg border border-amber-100">
            <span className="text-sm font-semibold text-amber-700 mr-2">
              資源配給:
            </span>
            <input
              type="text"
              value={newSchedule.resourceContent}
              onChange={(e) =>
                setNewSchedule((p) => ({
                  ...p,
                  resourceContent: e.target.value,
                }))
              }
              className={`${styles.formInput} flex-grow min-w-[150px]`}
              placeholder="配給內容"
            />
            <input
              type="month"
              value={newSchedule.resourceMonth}
              onChange={(e) =>
                setNewSchedule((p) => ({ ...p, resourceMonth: e.target.value }))
              }
              className={`${styles.formInput} w-36`}
              title="配給月份"
            />
            <input
              type="number"
              value={newSchedule.resourceAmount}
              onChange={(e) =>
                setNewSchedule((p) => ({
                  ...p,
                  resourceAmount: parseInt(e.target.value) || 0,
                }))
              }
              className={`${styles.formInput} w-28`}
              placeholder="金額"
            />
          </div>
          <div className="flex justify-end space-x-3 pt-2">
            <button
              onClick={() => setIsAddingSchedule(false)}
              className={styles.btnSecondary}
            >
              取消
            </button>
            <button
              onClick={handleAdd}
              className={`${styles.btnPrimary} bg-amber-600 hover:bg-amber-700`}
            >
              <Plus className="w-4 h-4 mr-1" /> 確認新增
            </button>
          </div>
        </div>
      );
    };

    const AddMeetingForm = () => {
      const [newMeeting, setNewMeeting] = useState({
        id: crypto.randomUUID(),
        date: new Date().toISOString().substring(0, 10),
        attendees: '',
        summary: '',
        todo: '',
        nextMeetingDate: '',
        nextAttendees: '',
        nextTopics: '',
      });

      const handleAdd = () => {
        if (!newMeeting.date || !newMeeting.summary) {
          alert('請填寫日期與會議總結。');
          return;
        }
        updatePrivateData({ meetings: [...appData.meetings, newMeeting] });
        setIsAddingMeeting(false);
      };

      return (
        <div className="p-6 border border-blue-200 rounded-xl bg-blue-50/50 backdrop-blur-sm shadow-inner space-y-4 mb-6">
          <h4 className="font-bold text-lg text-blue-800 flex items-center">
            <Plus className="w-5 h-5 mr-2" /> 新增戰勤會議紀錄
          </h4>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <input
              type="date"
              value={newMeeting.date}
              onChange={(e) =>
                setNewMeeting((p) => ({ ...p, date: e.target.value }))
              }
              className={styles.formInput}
              title="日期"
            />
            <input
              type="text"
              value={newMeeting.attendees}
              onChange={(e) =>
                setNewMeeting((p) => ({ ...p, attendees: e.target.value }))
              }
              className={styles.formInput}
              placeholder="與會人員"
            />
            <textarea
              value={newMeeting.summary}
              onChange={(e) =>
                setNewMeeting((p) => ({ ...p, summary: e.target.value }))
              }
              className={`${styles.formTextarea} md:col-span-2`}
              placeholder="會議總結"
              rows="3"
            />
            <textarea
              value={newMeeting.todo}
              onChange={(e) =>
                setNewMeeting((p) => ({ ...p, todo: e.target.value }))
              }
              className={`${styles.formTextarea} md:col-span-2`}
              placeholder="待辦事項"
              rows="2"
            />
          </div>

          <div className="bg-white/60 p-4 rounded-lg border border-blue-100 mt-4">
            <h5 className="font-semibold text-blue-700 mb-3 text-sm uppercase tracking-wider">
              下次會議資訊預告
            </h5>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
              <input
                type="date"
                value={newMeeting.nextMeetingDate}
                onChange={(e) =>
                  setNewMeeting((p) => ({
                    ...p,
                    nextMeetingDate: e.target.value,
                  }))
                }
                className={styles.formInput}
                title="下次開會時間"
              />
              <input
                type="text"
                value={newMeeting.nextAttendees}
                onChange={(e) =>
                  setNewMeeting((p) => ({
                    ...p,
                    nextAttendees: e.target.value,
                  }))
                }
                className={styles.formInput}
                placeholder="下次與會人員"
              />
              <input
                type="text"
                value={newMeeting.nextTopics}
                onChange={(e) =>
                  setNewMeeting((p) => ({ ...p, nextTopics: e.target.value }))
                }
                className={styles.formInput}
                placeholder="下次議題"
              />
            </div>
          </div>

          <div className="flex justify-end space-x-3 pt-2">
            <button
              onClick={() => setIsAddingMeeting(false)}
              className={styles.btnSecondary}
            >
              取消
            </button>
            <button onClick={handleAdd} className={styles.btnPrimary}>
              <Plus className="w-4 h-4 mr-1" /> 確認新增
            </button>
          </div>
        </div>
      );
    };

    const [editingScheduleId, setEditingScheduleId] = useState(null);
    const [editingMeetingId, setEditingMeetingId] = useState(null);

    return (
      <div className="space-y-8 p-6 max-w-7xl mx-auto">
        <div className="flex items-center space-x-3 mb-6">
          <Activity className="w-8 h-8 text-indigo-600" />
          <h2 className="text-3xl font-extrabold text-slate-800 tracking-tight">
            進攻行事曆
          </h2>
        </div>

        {/* 第一大塊: 狀態卡 */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <StatusCard
            title="進攻目標總家數"
            value={totalUnits}
            icon={<Target className="w-6 h-6 text-white" />}
            gradient="from-slate-500 to-slate-700"
          />
          <StatusCard
            title="本牌客戶"
            value={currentClients}
            icon={<CheckCircle className="w-6 h-6 text-white" />}
            gradient="from-emerald-500 to-teal-600"
          />
          <StatusCard
            title="已進攻暫定結案"
            value={settledNonClients}
            icon={<X className="w-6 h-6 text-white" />}
            gradient="from-rose-500 to-pink-600"
          />
        </div>

        {/* 第二大塊: 進攻排程表 */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100">
          <div className="p-6 bg-gradient-to-r from-amber-50 to-orange-50 border-b border-amber-100 flex justify-between items-center flex-wrap gap-4">
            <h3 className="text-xl font-bold text-amber-900 flex items-center">
              <span className="w-2 h-8 bg-amber-500 rounded-full mr-3"></span>
              進攻排程
            </h3>
            <div className="flex space-x-2">
              {!isAddingSchedule && (
                <button
                  onClick={() => setIsAddingSchedule(true)}
                  className={`${styles.btnPrimary} bg-amber-600 hover:bg-amber-700 border-none shadow-amber-200`}
                >
                  <Plus className="w-4 h-4 mr-1" /> 新增排程
                </button>
              )}
              <button
                onClick={() => setScheduleCollapsed((p) => !p)}
                className={`${styles.btnSecondary} text-amber-900 hover:bg-amber-100`}
              >
                {scheduleCollapsed ? (
                  <ChevronsDown className="w-4 h-4 mr-1" />
                ) : (
                  <ChevronsUp className="w-4 h-4 mr-1" />
                )}
                {scheduleCollapsed ? '展開所有' : '收合過期'}
              </button>
              <button
                onClick={deleteSelectedSchedules}
                disabled={selectedScheduleIds.length === 0}
                className={styles.btnDanger}
              >
                <Trash2 className="w-4 h-4 mr-1" /> 刪除
              </button>
              <button onClick={exportSchedules} className={styles.btnInfo}>
                <Download className="w-4 h-4 mr-1" /> 匯出
              </button>
            </div>
          </div>

          <div className="p-4">
            {isAddingSchedule && <AddScheduleForm />}

            <div className="overflow-x-auto rounded-xl border border-gray-200">
              <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50 text-gray-500">
                  <tr>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      選取
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      時程
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      人員/家數/區域
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      待辦&備忘
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      資源內容
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      月份
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      金額
                    </th>
                    <th className="p-4 text-right text-xs font-bold uppercase tracking-wider">
                      操作
                    </th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-100">
                  {filteredSchedules.map((schedule, index) => (
                    <ScheduleRow
                      key={schedule.id}
                      schedule={schedule}
                      index={index}
                      isEditing={editingScheduleId}
                      setIsEditing={setEditingScheduleId}
                    />
                  ))}
                </tbody>
              </table>
              {filteredSchedules.length === 0 && (
                <div className="p-8 text-center text-gray-400 bg-gray-50">
                  <p>尚無排程紀錄，請點擊上方按鈕新增。</p>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* 第三大塊: 戰勤會議紀錄 */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100">
          <div className="p-6 bg-gradient-to-r from-blue-50 to-indigo-50 border-b border-blue-100 flex justify-between items-center flex-wrap gap-4">
            <h3 className="text-xl font-bold text-blue-900 flex items-center">
              <span className="w-2 h-8 bg-blue-600 rounded-full mr-3"></span>
              戰勤會議紀錄
            </h3>
            <div className="flex space-x-2">
              {!isAddingMeeting && (
                <button
                  onClick={() => setIsAddingMeeting(true)}
                  className={styles.btnPrimary}
                >
                  <Plus className="w-4 h-4 mr-1" /> 新增紀錄
                </button>
              )}
              <button
                onClick={deleteSelectedMeetings}
                disabled={selectedMeetingIds.length === 0}
                className={styles.btnDanger}
              >
                <Trash2 className="w-4 h-4 mr-1" /> 刪除
              </button>
              <button onClick={exportMeetings} className={styles.btnInfo}>
                <Download className="w-4 h-4 mr-1" /> 匯出
              </button>
            </div>
          </div>

          <div className="p-4">
            {isAddingMeeting && <AddMeetingForm />}

            <div className="overflow-x-auto rounded-xl border border-gray-200 mt-4">
              <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50 text-gray-500">
                  <tr>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      選取
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      日期
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      與會人員
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      總結
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      待辦
                    </th>
                    <th className="p-4 text-left text-xs font-bold uppercase tracking-wider">
                      下次開會
                    </th>
                    <th className="p-4 text-right text-xs font-bold uppercase tracking-wider">
                      操作
                    </th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-100">
                  {appData.meetings.map((meeting, index) => (
                    <MeetingRow
                      key={meeting.id}
                      meeting={meeting}
                      index={index}
                      isEditing={editingMeetingId}
                      setIsEditing={setEditingMeetingId}
                    />
                  ))}
                </tbody>
              </table>
              {appData.meetings.length === 0 && (
                <div className="p-8 text-center text-gray-400 bg-gray-50">
                  <p>尚無會議紀錄。</p>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  };

  // --- Tab 2: 攻擊準則 ---
  const Tab2Guidelines = () => {
    const { guidelines, talkScripts } = appData.settings;
    const [editingId, setEditingId] = useState(null);
    const [newGuideline, setNewGuideline] = useState({
      title: '',
      content: '',
    });
    const [newTalkScript, setNewTalkScript] = useState({
      title: '',
      content: '',
    });

    const handleUpdateSettings = async (field, value) => {
      await updatePrivateData({ [field]: value });
      setEditingId(null);
    };

    const handleAddGuideline = () => {
      if (newGuideline.title && newGuideline.content) {
        handleUpdateSettings('guidelines', [
          ...guidelines,
          { ...newGuideline, id: crypto.randomUUID() },
        ]);
        setNewGuideline({ title: '', content: '' });
      }
    };

    const handleDeleteGuideline = (id) => {
      handleUpdateSettings(
        'guidelines',
        guidelines.filter((g) => g.id !== id)
      );
    };

    const handleAddTalkScript = () => {
      if (newTalkScript.title && newTalkScript.content) {
        handleUpdateSettings('talkScripts', [
          ...talkScripts,
          { ...newTalkScript, id: crypto.randomUUID() },
        ]);
        setNewTalkScript({ title: '', content: '' });
      }
    };

    const handleDeleteTalkScript = (id) => {
      handleUpdateSettings(
        'talkScripts',
        talkScripts.filter((t) => t.id !== id)
      );
    };

    return (
      <div className="space-y-8 p-6 max-w-7xl mx-auto">
        <div className="flex items-center space-x-3 mb-6">
          <CheckCircle className="w-8 h-8 text-indigo-600" />
          <h2 className="text-3xl font-extrabold text-slate-800">
            攻擊準則與話術庫
          </h2>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
          {/* 1. 核心進攻原則 */}
          <div className="bg-white rounded-2xl shadow-xl overflow-hidden flex flex-col h-full border border-slate-100">
            <div className="p-5 bg-gradient-to-r from-blue-600 to-indigo-700 text-white">
              <h3 className="text-xl font-bold flex items-center">
                <Target className="w-5 h-5 mr-2" /> 核心進攻原則
              </h3>
            </div>
            <div className="p-6 space-y-4 flex-grow bg-slate-50">
              <div className="space-y-4">
                {guidelines.map((item) => (
                  <div
                    key={item.id}
                    className="p-5 bg-white rounded-xl shadow-sm border-l-4 border-blue-500 relative transition hover:shadow-md hover:-translate-y-0.5"
                  >
                    {editingId === item.id ? (
                      <EditBlock
                        item={item}
                        field="guidelines"
                        onSave={handleUpdateSettings}
                        onCancel={() => setEditingId(null)}
                        collection={guidelines}
                      />
                    ) : (
                      <>
                        <p className="font-bold text-lg text-slate-800">
                          {item.title}
                        </p>
                        <p className="text-slate-600 mt-2 leading-relaxed whitespace-pre-wrap">
                          {item.content}
                        </p>
                        <div className="absolute top-3 right-3 flex space-x-1 opacity-50 hover:opacity-100 transition">
                          <button
                            onClick={() => setEditingId(item.id)}
                            className="text-blue-600 hover:bg-blue-50 p-1.5 rounded-full"
                          >
                            <Edit className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => handleDeleteGuideline(item.id)}
                            className="text-red-600 hover:bg-red-50 p-1.5 rounded-full"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                ))}
              </div>
            </div>
            <div className="p-4 bg-white border-t border-slate-100">
              <div className="p-4 border border-dashed border-blue-300 rounded-lg bg-blue-50/50">
                <h4 className="font-semibold mb-3 text-blue-800 text-sm">
                  新增原則
                </h4>
                <input
                  type="text"
                  value={newGuideline.title}
                  onChange={(e) =>
                    setNewGuideline((p) => ({ ...p, title: e.target.value }))
                  }
                  className={`${styles.formInput} mb-2`}
                  placeholder="標題"
                />
                <textarea
                  value={newGuideline.content}
                  onChange={(e) =>
                    setNewGuideline((p) => ({ ...p, content: e.target.value }))
                  }
                  className={`${styles.formTextarea} mb-3`}
                  placeholder="內容"
                  rows="2"
                />
                <button
                  onClick={handleAddGuideline}
                  className={`${styles.btnPrimary} w-full py-2`}
                >
                  <Plus className="w-4 h-4 mr-1" /> 新增原則
                </button>
              </div>
            </div>
          </div>

          {/* 2. 標準話術庫 */}
          <div className="bg-white rounded-2xl shadow-xl overflow-hidden flex flex-col h-full border border-slate-100">
            <div className="p-5 bg-gradient-to-r from-emerald-600 to-teal-700 text-white">
              <h3 className="text-xl font-bold flex items-center">
                <CheckCircle className="w-5 h-5 mr-2" /> 標準話術庫
              </h3>
            </div>
            <div className="p-6 space-y-4 flex-grow bg-slate-50">
              <div className="space-y-4">
                {talkScripts.map((item) => (
                  <div
                    key={item.id}
                    className="p-5 bg-white rounded-xl shadow-sm border-l-4 border-emerald-500 relative transition hover:shadow-md hover:-translate-y-0.5"
                  >
                    {editingId === item.id ? (
                      <EditBlock
                        item={item}
                        field="talkScripts"
                        onSave={handleUpdateSettings}
                        onCancel={() => setEditingId(null)}
                        collection={talkScripts}
                      />
                    ) : (
                      <>
                        <p className="font-bold text-lg text-emerald-800">
                          {item.title}
                        </p>
                        <div className="p-3 mt-3 bg-emerald-50/50 border border-emerald-100 rounded-lg text-slate-700 text-sm leading-relaxed whitespace-pre-wrap">
                          {item.content}
                        </div>
                        <div className="absolute top-3 right-3 flex space-x-1 opacity-50 hover:opacity-100 transition">
                          <button
                            onClick={() => setEditingId(item.id)}
                            className="text-blue-600 hover:bg-blue-50 p-1.5 rounded-full"
                          >
                            <Edit className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => handleDeleteTalkScript(item.id)}
                            className="text-red-600 hover:bg-red-50 p-1.5 rounded-full"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                ))}
              </div>
            </div>
            <div className="p-4 bg-white border-t border-slate-100">
              <div className="p-4 border border-dashed border-emerald-300 rounded-lg bg-emerald-50/50">
                <h4 className="font-semibold mb-3 text-emerald-800 text-sm">
                  新增話術
                </h4>
                <input
                  type="text"
                  value={newTalkScript.title}
                  onChange={(e) =>
                    setNewTalkScript((p) => ({ ...p, title: e.target.value }))
                  }
                  className={`${styles.formInput} mb-2`}
                  placeholder="標題"
                />
                <textarea
                  value={newTalkScript.content}
                  onChange={(e) =>
                    setNewTalkScript((p) => ({ ...p, content: e.target.value }))
                  }
                  className={`${styles.formTextarea} mb-3`}
                  placeholder="話術內容"
                  rows="3"
                />
                <button
                  onClick={handleAddTalkScript}
                  className={`${styles.btnPrimary} w-full py-2 bg-emerald-600 hover:bg-emerald-700`}
                >
                  <Plus className="w-4 h-4 mr-1" /> 新增話術
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const EditBlock = ({ item, field, onSave, onCancel, collection }) => {
    const [editTitle, setEditTitle] = useState(item.title);
    const [editContent, setEditContent] = useState(item.content);

    const handleSave = () => {
      const updatedCollection = collection.map((i) =>
        i.id === item.id ? { ...i, title: editTitle, content: editContent } : i
      );
      onSave(field, updatedCollection);
    };

    return (
      <div className="p-3 bg-amber-50 rounded-lg space-y-3 ring-2 ring-amber-400">
        <input
          type="text"
          value={editTitle}
          onChange={(e) => setEditTitle(e.target.value)}
          className={`${styles.formInput} font-bold text-lg border-amber-300 focus:ring-amber-500`}
        />
        <textarea
          value={editContent}
          onChange={(e) => setEditContent(e.target.value)}
          className={`${styles.formTextarea} border-amber-300 focus:ring-amber-500`}
          rows="4"
        />
        <div className="flex justify-end space-x-2">
          <button onClick={onCancel} className={`${styles.btnSecondary} text-sm`}>
            取消
          </button>
          <button onClick={handleSave} className={`${styles.btnPrimary} text-sm`}>
            <Save className="w-4 h-4 mr-1" /> 儲存
          </button>
        </div>
      </div>
    );
  };

  // --- Tab 3: Targets & Map (MODIFIED FOR IMAGE URL INPUT) ---
  const Tab3TargetsMap = () => {
    const { units, settings } = appData;
    const { areaMap, uploadedMapUrl, equipmentDB } = settings;

    const totalUnits = units.length;
    const currentClients = units.filter(
      (u) => u.attackStatus === 'client'
    ).length;
    const adminUnits = units.filter(
      (u) => u.category === 'Administrative'
    ).length;
    const academicUnits = units.filter((u) => u.category === 'Academic').length;
    const adminSubgroups = units.filter(
      (u) => u.category === 'Administrative' && u.subgroup === '獨立空間'
    ).length;
    const academicSubgroups = units.filter(
      (u) => u.category === 'Academic' && u.subgroup === '獨立空間'
    ).length;

    // Detect incomplete units
    const incompleteUnits = useMemo(() => {
      return units.filter(u => !checkUnitCompleteness(u));
    }, [units]);

    const [isFilterCollapsed, setIsFilterCollapsed] = useState(true);
    const [filter, setFilter] = useState({
      id: '',
      type: '',
      name: '',
      contact: '',
      phone: '',
      brand: '',
      model: '',
    });
    const [selectedUnitIds, setSelectedUnitIds] = useState([]);

    const filteredUnits = useMemo(() => {
      return units.filter((unit) => {
        const equipmentJson = safeParse(unit.equipment);
        const hasMatchingEquipment =
          filter.brand || filter.model
            ? equipmentJson.some(
                (eq) =>
                  (filter.brand === '' || eq.brand.includes(filter.brand)) &&
                  (filter.model === '' || eq.model.includes(filter.model))
              )
            : true;

        return (
          (filter.id === '' || unit.id.includes(filter.id)) &&
          (filter.type === '' || unit.category === filter.type) &&
          (filter.name === '' || unit.name.includes(filter.name)) &&
          (filter.contact === '' ||
            unit.contactName.includes(filter.contact)) &&
          (filter.phone === '' || unit.contactPhone.includes(filter.phone)) &&
          hasMatchingEquipment
        );
      });
    }, [units, filter]);

    const [mapState, setMapState] = useState({
      isDrawing: false,
      start: null,
      current: null,
      areaCodeInput: '',
      uploadedFile: null,
    });

    const handleMapClick = (e) => {
        const rect = e.currentTarget.getBoundingClientRect();
        const x = ((e.clientX - rect.left) / rect.width) * 100;
        const y = ((e.clientY - rect.top) / rect.height) * 100;
  
        if (!mapState.isDrawing) {
          setMapState((p) => ({
            ...p,
            isDrawing: true,
            start: { x, y },
            current: { x, y },
          }));
        } else {
          const { start, areaCodeInput } = mapState;
          if (!areaCodeInput) {
            alert('請輸入區域代表編號。');
            setMapState((p) => ({
              ...p,
              isDrawing: false,
              start: null,
              current: null,
            }));
            return;
          }
  
          const newArea = {
            id: crypto.randomUUID(),
            code: areaCodeInput,
            x1: Math.min(start.x, x),
            y1: Math.min(start.y, y),
            x2: Math.max(start.x, x),
            y2: Math.max(start.y, y),
            unitCount: 0,
          };
  
          updatePrivateData({ areaMap: [...areaMap, newArea] });
          setMapState((p) => ({
            ...p,
            isDrawing: false,
            start: null,
            current: null,
            areaCodeInput: '',
          }));
        }
      };

    const handleMouseMove = (e) => {
        if (!mapState.isDrawing) return;
        const rect = e.currentTarget.getBoundingClientRect();
        const x = ((e.clientX - rect.left) / rect.width) * 100;
        const y = ((e.clientY - rect.top) / rect.height) * 100;
        setMapState((p) => ({ ...p, current: { x, y } }));
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
        <div className="flex justify-between items-center mb-6">
          <div className="flex items-center space-x-3">
            <MapPin className="w-8 h-8 text-indigo-600" />
            <h2 className="text-3xl font-extrabold text-slate-800">戰情地圖與對象總覽</h2>
          </div>
          
          {incompleteUnits.length > 0 && (
            <div className="relative group z-20">
              <button className="flex items-center space-x-2 bg-rose-100 text-rose-700 px-4 py-2 rounded-lg font-bold animate-pulse hover:animate-none hover:bg-rose-200 transition">
                <AlertCircle className="w-5 h-5" />
                <span>{incompleteUnits.length} 筆資料待補齊</span>
                <ChevronsDown className="w-4 h-4" />
              </button>
              <div className="absolute right-0 top-full mt-2 w-64 bg-white rounded-xl shadow-2xl border border-rose-100 hidden group-hover:block overflow-hidden">
                <div className="bg-rose-50 p-3 text-xs font-bold text-rose-800 border-b border-rose-100">
                  異常單位清單 (點擊編輯)
                </div>
                <div className="max-h-60 overflow-y-auto">
                  {incompleteUnits.map(u => (
                    <button 
                      key={u.id}
                      onClick={() => {
                        setEditingUnitId(u.id);
                        setCurrentTab('record');
                        setIsNewUnit(false);
                      }}
                      className="w-full text-left px-4 py-2 text-sm hover:bg-rose-50 text-slate-700 border-b border-gray-50 last:border-0"
                    >
                      {u.name || '(未命名)'}
                      <span className="text-xs text-rose-400 block">缺: {
                        !u.name ? '名稱 ' : ''
                      }{
                        !u.buildingId ? '棟別 ' : ''
                      }{
                        !u.contactName ? '聯絡人 ' : ''
                      }</span>
                    </button>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* 狀態卡 */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-6">
          <StatusCard
            title="總家數"
            value={totalUnits}
            gradient="from-indigo-500 to-purple-600"
            icon={<Building className="w-6 h-6 text-white" />}
          />
          <StatusCard
            title="本牌家數"
            value={currentClients}
            gradient="from-emerald-500 to-teal-500"
            icon={<CheckCircle className="w-6 h-6 text-white" />}
          />
          <StatusCard
            title="行政單位"
            value={adminUnits}
            gradient="from-orange-400 to-red-500"
            icon={<Users className="w-6 h-6 text-white" />}
          />
          <StatusCard
            title="學術單位"
            value={academicUnits}
            gradient="from-sky-500 to-blue-600"
            icon={<Users className="w-6 h-6 text-white" />}
          />
        </div>

        {/* 單位分佈 */}
        <div className="bg-white p-6 rounded-2xl shadow-lg border border-slate-100 flex flex-col md:flex-row gap-8 items-center justify-around bg-gradient-to-br from-white to-indigo-50/30">
          <div className="text-center w-full md:w-1/3 p-4 bg-orange-50 rounded-xl border border-orange-100">
            <p className="font-bold text-xl text-orange-600 mb-2">行政單位</p>
            <div className="flex justify-between items-center text-slate-700 px-4">
              <span>總數: {adminUnits}</span>
              <div className="flex flex-col items-end">
                <span className="font-bold bg-orange-200 px-2 py-1 rounded-md text-orange-800">
                  獨立空間: {adminSubgroups}
                </span>
                <span className="text-[10px] text-orange-600/70 mt-1">(大單位底下獨立空間)</span>
              </div>
            </div>
          </div>
          <div className="text-center w-full md:w-1/3 p-4 bg-sky-50 rounded-xl border border-sky-100">
            <p className="font-bold text-xl text-sky-600 mb-2">學術單位</p>
            <div className="flex justify-between items-center text-slate-700 px-4">
              <span>總數: {academicUnits}</span>
              <div className="flex flex-col items-end">
                <span className="font-bold bg-sky-200 px-2 py-1 rounded-md text-sky-800">
                  獨立空間: {academicSubgroups}
                </span>
                <span className="text-[10px] text-sky-600/70 mt-1">(大單位底下獨立空間)</span>
              </div>
            </div>
          </div>
        </div>

        {/* Map Section */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100">
          <div className="p-5 bg-slate-800 text-white flex justify-between items-center">
            <h3 className="text-xl font-bold flex items-center">
              <MapPin className="w-5 h-5 mr-2" /> 校園地圖戰情室
            </h3>
            <div className="flex space-x-2 items-center">
               {/* Controls */}
               <input
                type="text"
                value={mapState.areaCodeInput}
                onChange={(e) =>
                  setMapState((p) => ({
                    ...p,
                    areaCodeInput: e.target.value,
                  }))
                }
                className={`${styles.formInput} w-24 bg-slate-700 border-slate-600 text-white placeholder-slate-400 focus:ring-slate-500`}
                placeholder="輸入區號"
                disabled={mapState.isDrawing}
              />
              <button
                onClick={() =>
                  setMapState((p) => ({
                    ...p,
                    isDrawing: !p.isDrawing,
                    start: null,
                    current: null,
                  }))
                }
                className={`px-4 py-2 rounded-lg font-medium transition shadow-sm flex items-center justify-center active:scale-95 border-transparent ${
                  mapState.isDrawing
                    ? 'bg-rose-500 text-white hover:bg-rose-600'
                    : 'bg-slate-600 text-white hover:bg-slate-500'
                }`}
                disabled={!uploadedMapUrl}
              >
                {mapState.isDrawing ? '點擊結束' : '圈選區域'}
              </button>
            </div>
          </div>

          {/* Map Display */}
          <div
            className={`relative w-full aspect-[2/1] bg-slate-100 overflow-hidden group ${
              mapState.isDrawing ? 'cursor-crosshair' : 'cursor-default'
            }`}
            onClick={handleMapClick}
            onMouseMove={handleMouseMove}
          >
            {uploadedMapUrl ? (
                <img 
                  src={uploadedMapUrl} 
                  alt="Campus Map" 
                  className="w-full h-full object-cover transition-transform duration-500 ease-out"
                  onError={(e) => {
                    e.target.onerror = null; 
                    // Fallback or alert logic could go here, but keeping it simple
                    // Often Google Drive links might 403 if not public
                  }}
                />
            ) : (
                <div className="w-full h-full flex items-center justify-center text-slate-400">
                    <p>請輸入圖片連結</p>
                </div>
            )}

            {/* Existing Areas Overlays */}
            {areaMap.map((area) => (
               <div
                 key={area.id}
                 className="absolute border-2 border-rose-500 bg-rose-500/20 hover:bg-rose-500/40 transition-all duration-200 group/area shadow-lg"
                 style={{
                   left: `${Math.min(area.x1, area.x2)}%`,
                   top: `${Math.min(area.y1, area.y2)}%`,
                   width: `${Math.abs(area.x2 - area.x1)}%`,
                   height: `${Math.abs(area.y2 - area.y1)}%`,
                 }}
               >
                 <span className="absolute -top-6 left-0 bg-rose-600 text-white text-xs px-2 py-0.5 rounded shadow-sm font-bold whitespace-nowrap z-10">
                   {area.code} ({units.filter(u => u.areaCode === area.code).length})
                 </span>
                 <button
                    className="absolute -top-2 -right-2 p-1 bg-white text-rose-600 rounded-full shadow-md opacity-0 group-hover/area:opacity-100 transition transform hover:scale-110 z-20"
                    onClick={(e) => {
                      e.stopPropagation();
                      updatePrivateData({
                        areaMap: areaMap.filter((a) => a.id !== area.id),
                      });
                    }}
                  >
                    <X className="w-3 h-3" />
                  </button>
               </div>
            ))}
             {/* Drawing Box */}
             {mapState.isDrawing && mapState.start && mapState.current && (
              <div
                className="absolute border-2 border-dashed border-yellow-400 bg-yellow-400/30"
                style={{
                  left: `${Math.min(mapState.start.x, mapState.current.x)}%`,
                  top: `${Math.min(mapState.start.y, mapState.current.y)}%`,
                  width: `${Math.abs(mapState.start.x - mapState.current.x)}%`,
                  height: `${Math.abs(mapState.start.y - mapState.current.y)}%`,
                }}
              ></div>
            )}
          </div>
        </div>

        {/* 客戶概覽與篩選 */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100">
          <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-gray-50/50">
            <h3 className="text-xl font-bold text-slate-800">
              進攻對象概覽{' '}
              <span className="text-sm font-normal text-slate-500 ml-2">
                (共 {filteredUnits.length} 筆)
              </span>
            </h3>
            <button
              onClick={() => setIsFilterCollapsed((p) => !p)}
              className="text-indigo-600 hover:text-indigo-800 text-sm flex items-center bg-indigo-50 px-3 py-1.5 rounded-lg transition"
            >
              {isFilterCollapsed ? '展開篩選' : '收合篩選'}
              {isFilterCollapsed ? (
                <ChevronsDown className="w-4 h-4 ml-1" />
              ) : (
                <ChevronsUp className="w-4 h-4 ml-1" />
              )}
            </button>
          </div>

          <div
            className={`transition-all duration-300 overflow-hidden bg-slate-50 border-b border-slate-100 ${
              isFilterCollapsed ? 'max-h-0' : 'max-h-auto p-6'
            }`}
          >
            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-4">
              <FilterInput
                label="客編 (ID)"
                value={filter.id}
                onChange={(e) =>
                  setFilter((p) => ({ ...p, id: e.target.value }))
                }
              />
              <FilterSelect
                label="類型"
                value={filter.type}
                onChange={(e) =>
                  setFilter((p) => ({ ...p, type: e.target.value }))
                }
              >
                <option value="">全部</option>
                <option value="Administrative">行政</option>
                <option value="Academic">學術</option>
              </FilterSelect>
              <FilterInput
                label="客戶名稱"
                value={filter.name}
                onChange={(e) =>
                  setFilter((p) => ({ ...p, name: e.target.value }))
                }
              />
              <FilterInput
                label="聯絡人"
                value={filter.contact}
                onChange={(e) =>
                  setFilter((p) => ({ ...p, contact: e.target.value }))
                }
              />
              <FilterSelect
                label="設備廠牌"
                value={filter.brand}
                onChange={(e) =>
                  setFilter((p) => ({ ...p, brand: e.target.value }))
                }
              >
                <option value="">全部廠牌</option>
                {[...new Set(equipmentDB.map((e) => e.brand))].map((b) => (
                  <option key={b} value={b}>
                    {b}
                  </option>
                ))}
              </FilterSelect>
              <button
                onClick={() =>
                  setFilter({
                    id: '',
                    type: '',
                    name: '',
                    contact: '',
                    phone: '',
                    brand: '',
                    model: '',
                  })
                }
                className="mt-6 text-sm text-slate-500 hover:text-rose-600 underline"
              >
                清除篩選
              </button>
            </div>
          </div>

          <div className="p-4">
            <div className="flex justify-between items-center mb-4 flex-wrap gap-2">
              <div className="flex space-x-2">
                 <button onClick={downloadImportTemplate} className={`${styles.btnSecondary} py-1.5 text-sm`}>
                    <Download className="w-4 h-4 mr-1"/> 下載匯入範例
                 </button>
                 <label className={`${styles.btnPrimary} py-1.5 text-sm cursor-pointer`}>
                    <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleImport} />
                    <FileUp className="w-4 h-4 mr-1"/> 匯入資料
                 </label>
              </div>
              <div className="flex space-x-2">
                <button
                  onClick={deleteSelectedUnits}
                  disabled={selectedUnitIds.length === 0}
                  className={`${styles.btnDanger} py-1.5 text-sm`}
                >
                  <Trash2 className="w-4 h-4 mr-1" /> 刪除 (
                  {selectedUnitIds.length})
                </button>
                <button
                  onClick={exportUnits}
                  className={`${styles.btnInfo} py-1.5 text-sm`}
                >
                  <Download className="w-4 h-4 mr-1" /> 匯出
                </button>
              </div>
            </div>

            <UnitTable
              units={filteredUnits}
              selectedUnitIds={selectedUnitIds}
              setSelectedUnitIds={setSelectedUnitIds}
              setCurrentTab={setCurrentTab}
              setEditingUnitId={setEditingUnitId}
              setIsNewUnit={setIsNewUnit}
            />
          </div>
        </div>
      </div>
    );
  };

  const FilterInput = ({ label, value, onChange }) => (
    <div className="flex flex-col">
      <label className="text-xs font-bold text-slate-500 mb-1">{label}</label>
      <input
        type="text"
        value={value}
        onChange={onChange}
        className={`${styles.formInput} text-sm`}
      />
    </div>
  );

  const FilterSelect = ({ label, value, onChange, children }) => (
    <div className="flex flex-col">
      <label className="text-xs font-bold text-slate-500 mb-1">{label}</label>
      <select value={value} onChange={onChange} className={`${styles.formSelect} text-sm`}>
        {children}
      </select>
    </div>
  );

  const UnitTable = ({
    units,
    selectedUnitIds,
    setSelectedUnitIds,
    setCurrentTab,
    setEditingUnitId,
    setIsNewUnit,
  }) => {
    return (
      <div className="overflow-x-auto rounded-xl border border-gray-200">
        <table className="min-w-full divide-y divide-gray-200">
          <thead className="bg-gray-50 text-gray-500">
            <tr>
              <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">
                選取
              </th>
              <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">
                類型
              </th>
              <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">
                客戶名稱
              </th>
              <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">
                聯絡人
              </th>
              <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">
                進攻狀態
              </th>
              <th className="p-3 text-right text-xs font-bold uppercase tracking-wider">
                檢視
              </th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-100">
            {units.map((unit) => (
              <tr key={unit.id} className="hover:bg-indigo-50/40 transition">
                <td className="p-3">
                  <input
                    type="checkbox"
                    checked={selectedUnitIds.includes(unit.id)}
                    onChange={() =>
                      setSelectedUnitIds((p) =>
                        p.includes(unit.id)
                          ? p.filter((id) => id !== unit.id)
                          : [...p, unit.id]
                      )
                    }
                    className={styles.checkbox}
                  />
                </td>
                <td className="p-3 text-sm">
                  <span
                    className={`px-2 py-0.5 rounded text-xs font-bold ${
                      unit.category === 'Academic'
                        ? 'bg-sky-100 text-sky-700'
                        : 'bg-orange-100 text-orange-700'
                    }`}
                  >
                    {unit.category === 'Academic' ? '學術' : '行政'}
                  </span>
                </td>
                <td className="p-3 text-sm font-medium text-gray-900">
                  {unit.name}
                  {!checkUnitCompleteness(unit) && (
                    <AlertCircle className="w-4 h-4 text-rose-500 inline ml-1" />
                  )}
                </td>
                <td className="p-3 text-sm text-gray-600">
                  {unit.contactName}
                  <span className="block text-xs text-gray-400">
                    {unit.contactPhone}
                  </span>
                </td>
                <td className="p-3 text-sm">
                  <span
                    className={`px-2 py-1 inline-flex text-xs leading-5 font-bold rounded-full shadow-sm ${
                      unit.attackStatus === 'client'
                        ? 'bg-emerald-100 text-emerald-800'
                        : unit.attackStatus === 'settled_non_client'
                        ? 'bg-rose-100 text-rose-800'
                        : 'bg-amber-100 text-amber-800'
                    }`}
                  >
                    {unit.attackStatus === 'client'
                      ? '本牌客戶'
                      : unit.attackStatus === 'settled_non_client'
                      ? '暫定結案'
                      : '進攻中'}
                  </span>
                </td>
                <td className="p-3 text-right">
                  <button
                    onClick={() => {
                      setCurrentTab('record');
                      setEditingUnitId(unit.id);
                      setIsNewUnit(false);
                    }}
                    className="px-3 py-1.5 text-xs font-medium bg-indigo-50 text-indigo-600 rounded-lg hover:bg-indigo-100 transition border border-indigo-200"
                  >
                    詳細
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {units.length === 0 && (
          <div className="p-8 text-center text-gray-400 bg-gray-50">
            尚無資料。
          </div>
        )}
      </div>
    );
  };

  // --- Tab 4: 紀錄拜訪行為 ---

  const [editingUnitId, setEditingUnitId] = useState(null);
  const [isNewUnit, setIsNewUnit] = useState(false);
  const [newUnitData, setNewUnitData] = useState({});
  const [recordFilter, setRecordFilter] = useState({ type: '', area: '' });

  useEffect(() => {
    const currentUnit = appData.units.find((u) => u.id === editingUnitId);
    if (editingUnitId && currentUnit) {
      setIsNewUnit(false);
      setNewUnitData({
        ...currentUnit,
        equipment: safeParse(currentUnit.equipment),
        history: safeParse(currentUnit.history),
      });
    } else if (isNewUnit) {
      setNewUnitData({
        name: '',
        category: 'Academic',
        subgroup: '',
        buildingId: '',
        attackStatus: 'engaged',
        contactName: '',
        contactPhone: '',
        areaCode: '',
        equipment: [],
        characteristics: [],
        history: [],
      });
    }
  }, [editingUnitId, isNewUnit, appData.units]);

  const handleAddHistory = (newLog) => {
    const logEntry = {
      ...newLog,
      date: new Date().toISOString().substring(0, 10),
      id: crypto.randomUUID(),
    };
    setNewUnitData((prev) => ({
      ...prev,
      history: [...prev.history, logEntry],
    }));
  };

  const handleSaveUnit = async () => {
    if (!newUnitData.name) {
      alert('單位名稱是必填項。');
      return;
    }
    const dataToSave = {
      ...newUnitData,
      subgroup:
        newUnitData.category === 'Academic' ? newUnitData.subgroup : '',
    };

    if (isNewUnit) {
      await addDoc(getUnitCollectionRef(db), {
        ...dataToSave,
        createdAt: new Date().toISOString(),
        equipment: safeStringify(dataToSave.equipment || []),
        history: safeStringify(dataToSave.history || []),
        characteristics: dataToSave.characteristics || [],
      });
    } else {
      await updateUnit(editingUnitId, dataToSave);
    }
    setEditingUnitId(null);
    setIsNewUnit(false);
  };

  const Tab4Record = () => {
    const unitsList = appData.units.filter(
      (unit) =>
        (recordFilter.type === '' || unit.category === recordFilter.type) &&
        (recordFilter.area === '' || unit.areaCode === recordFilter.area)
    );

    const unitTableHeaders = [
      { key: 'category', label: '類型', width: 10 },
      { key: 'areaCode', label: '區域', width: 10 },
      { key: 'name', label: '客戶名稱', width: 25 },
      { key: 'contactName', label: '聯絡人', width: 15 },
      { key: 'contactPhone', label: '電話', width: 15 },
      { key: 'attackStatus', label: '進攻狀態', width: 15 },
    ];
    const exportRecordUnits = () => {
      exportToExcel(
        unitsList,
        '客戶拜訪紀錄清單',
        '客戶清單',
        unitTableHeaders
      );
    };

    if (editingUnitId !== null || isNewUnit) {
      return (
        <UnitRecordView
          newUnitData={newUnitData}
          setNewUnitData={setNewUnitData}
          handleSaveUnit={handleSaveUnit}
          handleAddHistory={handleAddHistory}
          isNewUnit={isNewUnit}
          appData={appData}
          setEditingUnitId={setEditingUnitId}
          setIsNewUnit={setIsNewUnit}
        />
      );
    }

    return (
      <div className="p-6 space-y-6 max-w-7xl mx-auto">
        <div className="flex items-center space-x-3 mb-6">
          <Edit className="w-8 h-8 text-indigo-600" />
          <h2 className="text-3xl font-extrabold text-slate-800">
            拜訪行為紀錄
          </h2>
        </div>

        <div className="bg-white p-6 rounded-2xl shadow-lg border border-slate-100 flex flex-wrap gap-4 items-center">
          <FilterSelect
            label="類型篩選"
            value={recordFilter.type}
            onChange={(e) =>
              setRecordFilter((p) => ({ ...p, type: e.target.value }))
            }
          >
            <option value="">全部類型</option>
            <option value="Administrative">行政</option>
            <option value="Academic">學術</option>
          </FilterSelect>
          <FilterSelect
            label="區域篩選"
            value={recordFilter.area}
            onChange={(e) =>
              setRecordFilter((p) => ({ ...p, area: e.target.value }))
            }
          >
            <option value="">全部區域</option>
            {appData.settings.areaMap.map((a) => (
              <option key={a.code} value={a.code}>
                {a.code}
              </option>
            ))}
          </FilterSelect>

          <div className="flex-grow"></div>

          <button
            onClick={() => {
              setEditingUnitId(null);
              setIsNewUnit(true);
            }}
            className={styles.btnPrimary}
          >
            <Plus className="w-4 h-4 mr-1" /> 新增對象
          </button>
          <button onClick={exportRecordUnits} className={styles.btnInfo}>
            <Download className="w-4 h-4 mr-1" /> 匯出
          </button>
        </div>

        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100 p-4">
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200">
              <thead className="bg-gray-50 text-gray-500">
                <tr>
                  {unitTableHeaders.map((h) => (
                    <th
                      key={h.key}
                      className="p-3 text-left text-xs font-bold uppercase tracking-wider"
                    >
                      {h.label}
                    </th>
                  ))}
                  <th className="p-3 text-right text-xs font-bold uppercase tracking-wider">
                    動作
                  </th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-100">
                {unitsList.map((unit) => (
                  <tr key={unit.id} className="hover:bg-indigo-50/50">
                    <td className="p-3 text-sm text-gray-700">
                      {unit.category}
                    </td>
                    <td className="p-3 text-sm text-gray-700">
                      {unit.areaCode || '-'}
                    </td>
                    <td className="p-3 text-sm font-medium text-gray-900">
                      {unit.name}
                    </td>
                    <td className="p-3 text-sm text-gray-600">
                      {unit.contactName}
                    </td>
                    <td className="p-3 text-sm text-gray-600">
                      {unit.contactPhone}
                    </td>
                    <td className="p-3 text-sm">
                      <span
                        className={`px-2 py-0.5 inline-flex text-xs leading-5 font-bold rounded-full ${
                          unit.attackStatus === 'client'
                            ? 'bg-emerald-100 text-emerald-800'
                            : unit.attackStatus === 'settled_non_client'
                            ? 'bg-rose-100 text-rose-800'
                            : 'bg-amber-100 text-amber-800'
                        }`}
                      >
                        {unit.attackStatus === 'client'
                          ? '本牌客戶'
                          : unit.attackStatus === 'settled_non_client'
                          ? '暫定結案'
                          : '進攻中'}
                      </span>
                    </td>
                    <td className="p-3 text-right">
                      <button
                        onClick={() => setEditingUnitId(unit.id)}
                        className={`${styles.btnPrimary} py-1.5 px-3 text-xs`}
                      >
                        任務/編輯
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {unitsList.length === 0 && (
              <p className="text-center text-gray-400 py-8">
                無符合篩選條件的客戶。
              </p>
            )}
          </div>
        </div>
      </div>
    );
  };

  const UnitRecordView = ({
    newUnitData,
    setNewUnitData,
    handleSaveUnit,
    handleAddHistory,
    isNewUnit,
    appData,
    setEditingUnitId,
    setIsNewUnit,
  }) => {
    const { equipment, characteristics, history } = newUnitData;

    const [equipmentSearch, setEquipmentSearch] = useState({
      brand: '',
      model: '',
    });
    const availableBrands = [
      ...new Set(appData.settings.equipmentDB.map((e) => e.brand)),
    ];
    const availableModels = [
      ...new Set(
        appData.settings.equipmentDB
          .filter((e) => e.brand === equipmentSearch.brand)
          .map((e) => e.model)
      ),
    ];

    return (
      <div className="bg-white p-8 rounded-2xl shadow-2xl space-y-8 max-w-5xl mx-auto my-6 border border-slate-200">
        <div className="flex justify-between items-center border-b pb-6">
          <h3 className="text-2xl font-extrabold text-slate-800">
            {isNewUnit ? '新增進攻對象' : `編輯/紀錄: ${newUnitData.name}`}
          </h3>
          <button
            onClick={() => {
              setEditingUnitId(null);
              setIsNewUnit(false);
            }}
            className="text-gray-400 hover:text-gray-600"
          >
            <X className="w-6 h-6" />
          </button>
        </div>

        {/* Section 1: Basic Unit Info */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <InputGroup label="單位名稱 (必填)">
            <input
              type="text"
              value={newUnitData.name || ''}
              onChange={(e) =>
                setNewUnitData((p) => ({ ...p, name: e.target.value }))
              }
              className={styles.formInput}
            />
          </InputGroup>

          <InputGroup label="棟別">
            <select
              value={newUnitData.buildingId || ''}
              onChange={(e) =>
                setNewUnitData((p) => ({ ...p, buildingId: e.target.value }))
              }
              className={styles.formSelect}
            >
              <option value="">選擇棟別</option>
              {appData.settings.buildings.map((b) => (
                <option key={b.code} value={b.code}>
                  {b.name} ({b.code})
                </option>
              ))}
            </select>
          </InputGroup>

          <InputGroup label="區域編號">
            <select
              value={newUnitData.areaCode || ''}
              onChange={(e) =>
                setNewUnitData((p) => ({ ...p, areaCode: e.target.value }))
              }
              className={styles.formSelect}
            >
              <option value="">無區域</option>
              {appData.settings.areaMap.map((a) => (
                <option key={a.code} value={a.code}>
                  {a.code}
                </option>
              ))}
            </select>
          </InputGroup>

          <InputGroup label="進攻狀態">
            <select
              value={newUnitData.attackStatus || 'engaged'}
              onChange={(e) =>
                setNewUnitData((p) => ({ ...p, attackStatus: e.target.value }))
              }
              className={styles.formSelect}
            >
              <option value="engaged">進攻中</option>
              <option value="settled_non_client">已進攻暫定結案</option>
              <option value="client">本牌客戶</option>
            </select>
          </InputGroup>

          <InputGroup label="承辦姓名">
            <input
              type="text"
              value={newUnitData.contactName || ''}
              onChange={(e) =>
                setNewUnitData((p) => ({ ...p, contactName: e.target.value }))
              }
              className={styles.formInput}
            />
          </InputGroup>

          <InputGroup label="電話">
            <input
              type="text"
              value={newUnitData.contactPhone || ''}
              onChange={(e) =>
                setNewUnitData((p) => ({ ...p, contactPhone: e.target.value }))
              }
              className={styles.formInput}
            />
          </InputGroup>

          <InputGroup label="單位類別">
            <select
              value={newUnitData.category || 'Academic'}
              onChange={(e) =>
                setNewUnitData((p) => ({ ...p, category: e.target.value }))
              }
              className={styles.formSelect}
            >
              <option value="Academic">學術單位</option>
              <option value="Administrative">行政單位</option>
            </select>
          </InputGroup>

          {(newUnitData.category === 'Academic' ||
            newUnitData.category === 'Administrative') && (
            <InputGroup label="獨立空間分組">
              <select
                value={newUnitData.subgroup || ''}
                onChange={(e) =>
                  setNewUnitData((p) => ({ ...p, subgroup: e.target.value }))
                }
                className={styles.formSelect}
              >
                <option value="">一般</option>
                <option value="獨立空間">獨立空間</option>
              </select>
            </InputGroup>
          )}
        </div>

        <div className="border-t border-slate-100 my-6"></div>

        {/* Section 2: Equipment & Characteristics */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
          <div className="bg-indigo-50/50 p-6 rounded-xl border border-indigo-100">
            <h4 className="text-lg font-bold mb-4 text-indigo-800 flex items-center">
              <span className="w-2 h-2 bg-indigo-500 rounded-full mr-2"></span>
              設備清單 ({equipment?.length || 0})
            </h4>
            <EquipmentAdder
              availableBrands={availableBrands}
              availableModels={availableModels}
              machineTypes={appData.settings.machineTypes}
              equipmentDB={appData.settings.equipmentDB}
              onAdd={(eq) =>
                setNewUnitData((p) => ({
                  ...p,
                  equipment: [
                    ...p.equipment,
                    { ...eq, id: crypto.randomUUID() },
                  ],
                }))
              }
              equipmentSearch={equipmentSearch}
              setEquipmentSearch={setEquipmentSearch}
            />
            <EquipmentList
              equipment={equipment}
              setNewUnitData={setNewUnitData}
            />
          </div>
          <div className="bg-amber-50/50 p-6 rounded-xl border border-amber-100">
            <h4 className="text-lg font-bold mb-4 text-amber-800 flex items-center">
              <span className="w-2 h-2 bg-amber-500 rounded-full mr-2"></span>
              客戶特性
            </h4>
            <CharacteristicsEditor
              characteristics={characteristics}
              setNewUnitData={setNewUnitData}
            />
          </div>
        </div>

        <div className="border-t border-slate-100 my-6"></div>

        {/* Section 3: Visit Log */}
        <div className="bg-emerald-50/30 p-6 rounded-xl border border-emerald-100">
          <h4 className="text-lg font-bold mb-4 text-emerald-800 flex items-center">
            <span className="w-2 h-2 bg-emerald-500 rounded-full mr-2"></span>
            記錄本次拜訪行為
          </h4>
          <HistoryLogAdder onAdd={handleAddHistory} />
          <HistoryLogList history={history} />
        </div>

        <div className="flex justify-end space-x-4 pt-6">
          <button
            onClick={() => {
              setEditingUnitId(null);
              setIsNewUnit(false);
            }}
            className="px-6 py-2.5 bg-gray-100 text-gray-600 rounded-lg hover:bg-gray-200 transition font-medium"
          >
            取消
          </button>
          <button
            onClick={handleSaveUnit}
            className={`${styles.btnPrimary} bg-gradient-to-r from-indigo-600 to-blue-600 px-8 py-2.5 hover:shadow-indigo-500/30 hover:-translate-y-0.5`}
          >
            <Save className="w-5 h-5 mr-2" /> 儲存資料
          </button>
        </div>
      </div>
    );
  };

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

  const InputGroup = ({ label, children }) => (
    <div className="flex flex-col space-y-1.5">
      <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">
        {label}
      </label>
      {children}
    </div>
  );

  const EquipmentAdder = ({
    availableBrands,
    availableModels,
    machineTypes,
    equipmentDB,
    onAdd,
    equipmentSearch,
    setEquipmentSearch,
  }) => {
    const [newEq, setNewEq] = useState({ brand: '', model: '', type: '' });
    const handleAdd = () => {
      if (newEq.brand && newEq.model && newEq.type) {
        onAdd(newEq);
        setNewEq({ brand: '', model: '', type: '' });
        setEquipmentSearch({ brand: '', model: '' });
      } else {
        alert('請填寫完整的設備資訊。');
      }
    };

    useEffect(() => {
      const match = equipmentDB.find(
        (e) => e.brand === newEq.brand && e.model === newEq.model
      );
      if (match && newEq.type === '') {
        setNewEq((p) => ({ ...p, type: match.type }));
      }
    }, [newEq.brand, newEq.model, equipmentDB]);

    return (
      <div className="flex flex-col gap-3">
        <div className="flex gap-2">
          <select
            value={newEq.brand}
            onChange={(e) => {
              setNewEq((p) => ({ ...p, brand: e.target.value, model: '' }));
              setEquipmentSearch((p) => ({ ...p, brand: e.target.value }));
            }}
            className={`${styles.formSelect} flex-1`}
          >
            <option value="">選擇廠牌</option>
            {availableBrands.map((b) => (
              <option key={b} value={b}>
                {b}
              </option>
            ))}
          </select>

          <select
            value={newEq.model}
            onChange={(e) => setNewEq((p) => ({ ...p, model: e.target.value }))}
            className={`${styles.formSelect} flex-1`}
          >
            <option value="">選擇型號</option>
            {availableModels.map((m) => (
              <option key={m} value={m}>
                {m}
              </option>
            ))}
          </select>
        </div>
        <div className="flex gap-2">
          <select
            value={newEq.type}
            onChange={(e) => setNewEq((p) => ({ ...p, type: e.target.value }))}
            className={`${styles.formSelect} flex-1`}
          >
            <option value="">選擇類型</option>
            {machineTypes.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>

          <button
            onClick={handleAdd}
            className="p-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition"
          >
            <Plus className="w-5 h-5" />
          </button>
        </div>
      </div>
    );
  };

  const EquipmentList = ({ equipment, setNewUnitData }) => (
    <div className="space-y-2 mt-4 max-h-48 overflow-y-auto pr-1">
      {(equipment || []).map((eq) => (
        <div
          key={eq.id}
          className="flex justify-between items-center p-3 bg-white border border-indigo-100 rounded-lg shadow-sm"
        >
          <div className="text-sm">
            <span className="font-bold text-slate-800">
              {eq.brand} {eq.model}
            </span>
            <span className="block text-xs text-slate-500">{eq.type}</span>
          </div>
          <button
            onClick={() =>
              setNewUnitData((p) => ({
                ...p,
                equipment: p.equipment.filter((e) => e.id !== eq.id),
              }))
            }
            className="text-gray-400 hover:text-red-500 transition"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      ))}
    </div>
  );

  const CharacteristicsEditor = ({ characteristics, setNewUnitData }) => {
    const [newChar, setNewChar] = useState('');
    const handleAdd = () => {
      if (newChar) {
        setNewUnitData((p) => ({
          ...p,
          characteristics: [...p.characteristics, newChar],
        }));
        setNewChar('');
      }
    };
    const handleDelete = (char) => {
      setNewUnitData((p) => ({
        ...p,
        characteristics: p.characteristics.filter((c) => c !== char),
      }));
    };

    return (
      <div className="space-y-3">
        <div className="flex gap-2">
          <input
            type="text"
            value={newChar}
            onChange={(e) => setNewChar(e.target.value)}
            className={`${styles.formInput} flex-grow`}
            placeholder="e.g. 預算緊張"
          />
          <button
            onClick={handleAdd}
            className="p-2 bg-amber-500 text-white rounded-lg hover:bg-amber-600 transition"
          >
            <Plus className="w-5 h-5" />
          </button>
        </div>
        <div className="flex flex-wrap gap-2">
          {(characteristics || []).map((char, index) => (
            <span
              key={index}
              className="flex items-center bg-white border border-amber-200 text-amber-800 text-sm px-3 py-1 rounded-full shadow-sm"
            >
              {char}
              <button
                onClick={() => handleDelete(char)}
                className="ml-2 text-amber-400 hover:text-amber-600"
              >
                <X className="w-3 h-3" />
              </button>
            </span>
          ))}
        </div>
      </div>
    );
  };

  const HistoryLogAdder = ({ onAdd }) => {
    const [newLog, setNewLog] = useState({
      activity: '',
      item: '',
      quantity: 1,
      supplement: '',
    });

    const handleAdd = () => {
      if (newLog.activity) {
        onAdd({
          activity: newLog.activity,
          promotionalItems: newLog.item
            ? [{ item: newLog.item, quantity: newLog.quantity }]
            : [],
          characteristicSupplement: newLog.supplement,
        });
        setNewLog({ activity: '', item: '', quantity: 1, supplement: '' });
      } else {
        alert('請填寫實際行為。');
      }
    };

    return (
      <div className="bg-white p-4 rounded-xl shadow-sm border border-emerald-100 space-y-3">
        <input
          type="text"
          value={newLog.activity}
          onChange={(e) =>
            setNewLog((p) => ({ ...p, activity: e.target.value }))
          }
          className={`${styles.formInput} font-medium`}
          placeholder="行為紀錄 (e.g. 設備demo)"
        />
        <div className="flex gap-2">
          <input
            type="text"
            value={newLog.item}
            onChange={(e) => setNewLog((p) => ({ ...p, item: e.target.value }))}
            className={`${styles.formInput} flex-grow`}
            placeholder="輔銷物"
          />
          <input
            type="number"
            value={newLog.quantity}
            onChange={(e) =>
              setNewLog((p) => ({
                ...p,
                quantity: parseInt(e.target.value) || 1,
              }))
            }
            className={`${styles.formInput} w-20 text-center`}
            min="1"
          />
        </div>
        <textarea
          value={newLog.supplement}
          onChange={(e) =>
            setNewLog((p) => ({ ...p, supplement: e.target.value }))
          }
          className={`${styles.formTextarea} text-sm`}
          placeholder="補充說明..."
          rows="2"
        ></textarea>
        <button
          onClick={handleAdd}
          className="w-full py-2 bg-emerald-50 text-emerald-700 font-bold rounded-lg hover:bg-emerald-100 transition"
        >
          <Plus className="w-4 h-4 mr-1 inline" /> 加入紀錄
        </button>
      </div>
    );
  };

  const HistoryLogList = ({ history }) => (
    <div className="mt-6">
      <h4 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-3">
        歷史紀錄
      </h4>
      <div className="space-y-3 max-h-80 overflow-y-auto pr-1">
        {(history || [])
          .sort((a, b) => new Date(b.date) - new Date(a.date))
          .map((log) => (
            <div
              key={log.id}
              className="p-4 bg-white border border-slate-100 rounded-xl shadow-sm hover:shadow-md transition"
            >
              <div className="flex justify-between items-start mb-1">
                <span className="font-bold text-slate-800">{log.activity}</span>
                <span className="text-xs text-slate-400 font-mono">
                  {log.date}
                </span>
              </div>
              {(log.promotionalItems || []).length > 0 && (
                <div className="flex flex-wrap gap-1 mt-2">
                  {log.promotionalItems.map((item, i) => (
                    <span
                      key={i}
                      className="px-2 py-0.5 bg-yellow-100 text-yellow-800 text-xs rounded-full font-medium"
                    >
                      {item.item} x{item.quantity}
                    </span>
                  ))}
                </div>
              )}
              {log.characteristicSupplement && (
                <p className="text-xs text-slate-500 mt-2 bg-slate-50 p-2 rounded">
                  {log.characteristicSupplement}
                </p>
              )}
            </div>
          ))}
      </div>
    </div>
  );

  // --- Tab 5: 參數設定 ---
  const Tab5Settings = () => {
    const [newBuilding, setNewBuilding] = useState({ name: '', code: '' });
    const [newEquipment, setNewEquipment] = useState({
      brand: '',
      model: '',
      type: '',
    });

    const handleUpdateSettings = async (field, value) => {
      await updatePrivateData({ [field]: value });
    };

    const handleAddBuilding = () => {
      if (newBuilding.name && newBuilding.code) {
        const updatedBuildings = [...appData.settings.buildings, newBuilding];
        handleUpdateSettings('buildings', updatedBuildings);
        setNewBuilding({ name: '', code: '' });
      }
    };

    const handleDeleteBuilding = (code) => {
      const updatedBuildings = appData.settings.buildings.filter(
        (b) => b.code !== code
      );
      handleUpdateSettings('buildings', updatedBuildings);
    };

    const handleAddEquipmentDB = () => {
      if (newEquipment.brand && newEquipment.model && newEquipment.type) {
        const updatedDB = [...appData.settings.equipmentDB, newEquipment];
        handleUpdateSettings('equipmentDB', updatedDB);

        if (!appData.settings.machineTypes.includes(newEquipment.type)) {
          handleUpdateSettings('machineTypes', [
            ...appData.settings.machineTypes,
            newEquipment.type,
          ]);
        }
        setNewEquipment({ brand: '', model: '', type: '' });
      }
    };

    const handleDeleteEquipmentDB = (brand, model) => {
      const updatedDB = appData.settings.equipmentDB.filter(
        (e) => !(e.brand === brand && e.model === model)
      );
      handleUpdateSettings('equipmentDB', updatedDB);
    };

    return (
      <div className="space-y-8 p-6 max-w-7xl mx-auto">
        <div className="flex items-center space-x-3 mb-6">
          <Edit className="w-8 h-8 text-indigo-600" />
          <h2 className="text-3xl font-extrabold text-slate-800">參數設定</h2>
        </div>

        {/* 1. 棟別/校區名稱設定 */}
        <div className="bg-white p-6 rounded-2xl shadow-xl border border-slate-100">
          <h3 className="text-xl font-bold mb-6 text-slate-800 border-b pb-2 flex items-center">
            <Building className="w-5 h-5 mr-2 text-indigo-500" /> 棟別設定
          </h3>
          <div className="flex flex-wrap gap-3 mb-6 bg-slate-50 p-4 rounded-xl">
            <input
              type="text"
              value={newBuilding.name}
              onChange={(e) =>
                setNewBuilding((p) => ({ ...p, name: e.target.value }))
              }
              className={`${styles.formInput} flex-grow`}
              placeholder="名稱 (e.g. 行政大樓)"
            />
            <input
              type="text"
              value={newBuilding.code}
              onChange={(e) =>
                setNewBuilding((p) => ({ ...p, code: e.target.value }))
              }
              className={`${styles.formInput} w-24`}
              placeholder="代號"
            />
            <button
              onClick={handleAddBuilding}
              className={styles.btnPrimary}
            >
              <Plus className="w-4 h-4 mr-1" /> 新增
            </button>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
            {appData.settings.buildings.map((b) => (
              <div
                key={b.code}
                className="flex justify-between items-center p-3 bg-white border border-slate-200 rounded-lg shadow-sm hover:border-indigo-300 transition"
              >
                <p className="text-slate-700 font-medium">
                  {b.name}{' '}
                  <span className="ml-2 px-2 py-0.5 bg-indigo-100 text-indigo-700 text-xs rounded-full">
                    {b.code}
                  </span>
                </p>
                <button
                  onClick={() => handleDeleteBuilding(b.code)}
                  className="text-slate-400 hover:text-red-500 transition"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            ))}
          </div>
        </div>

        {/* 2. 設備資料庫設定 */}
        <div className="bg-white p-6 rounded-2xl shadow-xl border border-slate-100">
          <h3 className="text-xl font-bold mb-6 text-slate-800 border-b pb-2 flex items-center">
            <Activity className="w-5 h-5 mr-2 text-indigo-500" /> 設備資料庫
          </h3>
          <div className="flex flex-wrap gap-3 mb-6 bg-slate-50 p-4 rounded-xl">
            <select
              value={newEquipment.type}
              onChange={(e) =>
                setNewEquipment((p) => ({ ...p, type: e.target.value }))
              }
              className={`${styles.formSelect} flex-1 min-w-[150px]`}
            >
              <option value="">選擇類型</option>
              {appData.settings.machineTypes.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
            <input
              type="text"
              value={newEquipment.brand}
              onChange={(e) =>
                setNewEquipment((p) => ({ ...p, brand: e.target.value }))
              }
              className={`${styles.formInput} flex-1 min-w-[120px]`}
              placeholder="廠牌"
            />
            <input
              type="text"
              value={newEquipment.model}
              onChange={(e) =>
                setNewEquipment((p) => ({ ...p, model: e.target.value }))
              }
              className={`${styles.formInput} flex-1 min-w-[120px]`}
              placeholder="型號"
            />
            <button
              onClick={handleAddEquipmentDB}
              className={styles.btnPrimary}
            >
              <Plus className="w-4 h-4 mr-1" /> 新增
            </button>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-h-96 overflow-y-auto">
            {appData.settings.equipmentDB.map((e, index) => (
              <div
                key={index}
                className="flex justify-between items-center p-3 bg-white border border-slate-200 rounded-lg shadow-sm"
              >
                <div>
                  <span className="font-bold text-slate-800">
                    {e.brand} {e.model}
                  </span>
                  <span className="block text-xs text-slate-500">{e.type}</span>
                </div>
                <button
                  onClick={() => handleDeleteEquipmentDB(e.brand, e.model)}
                  className="text-slate-400 hover:text-red-500 transition"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  };

  const renderTabContent = () => {
    switch (currentTab) {
      case 'calendar': return <Tab1Calendar />;
      case 'guidelines': return <Tab2Guidelines />;
      case 'targets': return <Tab3TargetsMap />;
      case 'record': return <Tab4Record />;
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
                 <div className="text-xs text-slate-500 font-mono">ID: {userId ? String(userId).substring(0, 8) + '...' : 'Guest'}</div>
            </div>
            <nav className="flex space-x-1 overflow-x-auto pb-1 no-scrollbar">
            {navItems.map((item) => (
                <button
                key={item.id}
                onClick={() => setCurrentTab(item.id)}
                className={`relative px-5 py-3 text-sm font-medium transition-all duration-300 rounded-t-lg flex items-center space-x-2 whitespace-nowrap ${currentTab === item.id ? 'text-indigo-600 bg-indigo-50/50' : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'}`}
                >
                {item.icon}
                <span>{item.label}</span>
                {currentTab === item.id && <span className="absolute bottom-0 left-0 w-full h-0.5 bg-indigo-600 rounded-t-full" />}
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

export default App;