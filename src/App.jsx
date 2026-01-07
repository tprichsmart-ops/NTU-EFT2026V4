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
  Link as LinkIcon,
  ZoomIn,
  ZoomOut,
  Undo,
  Check,
  Camera,
  Printer,
  User,
  Phone,
  History,
  Eye,
  FileText,
  Calendar as CalendarIcon,
  Columns,
  Copy,
  MessageSquare
} from 'lucide-react';

// --- 1. 全域設定與工具 (Global Config & Utils) ---

const appId = typeof __app_id !== 'undefined' ? __app_id : 'ntu-strategy-default-app';

const firebaseConfig =
  typeof __firebase_config !== 'undefined' ? JSON.parse(__firebase_config) : {};
const initialAuthToken =
  typeof __initial_auth_token !== 'undefined' ? __initial_auth_token : null;

const safeStringify = (data) => JSON.stringify(data);
const safeParse = (data) => {
  try {
    return JSON.parse(data);
  } catch (e) {
    return data;
  }
};

const downloadImportTemplate = () => {
  if (typeof window.XLSX === 'undefined') {
    alert('Excel 工具尚未載入，請稍候再試');
    return;
  }
  
  const headers = [
    '單位名稱 (必填)', '棟別代號 (必填)', '樓層', '科室別(房號)',
    '承辦姓名1(主)', '電話1(主)', '承辦姓名2', '電話2', '承辦姓名3', '電話3',
    '獨立空間分組 (請填: 獨立空間/一般)', '進攻狀態 (請填: 進攻中/已進攻暫定結案/本牌客戶)', '單位類別 (請填: 行政/學術)'
  ];

  const sampleData = [
    ['範例系辦', 'B2', '3F', '302室', '王小明', '02-33661234', '陳小華', '0912345678', '', '', '獨立空間', '進攻中', '學術'],
    ['範例處室', 'A1', '1F', '註冊組', '李大同', '02-33665678', '', '', '', '', '一般', '本牌客戶', '行政']
  ];

  const ws = window.XLSX.utils.aoa_to_sheet([headers, ...sampleData]);
  ws['!cols'] = [{ wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 20 }, { wch: 10 }];

  const range = { s: { r: 1, c: 0 }, e: { r: 100, c: 12 } };
  if (!ws['!dataValidation']) ws['!dataValidation'] = [];

  const setValidation = (colIndex, list) => {
      for (let r = range.s.r; r <= range.e.r; ++r) {
        ws['!dataValidation'].push({
          sqref: window.XLSX.utils.encode_cell({r: r, c: colIndex}),
          type: 'list',
          operator: 'equal',
          formula1: `"${list}"`,
          showErrorMessage: true,
          errorTitle: '輸入錯誤',
          error: '請從下拉選單中選擇'
        });
      }
  }
  setValidation(10, "獨立空間,一般");
  setValidation(11, "進攻中,已進攻暫定結案,本牌客戶");
  setValidation(12, "行政,學術");

  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, '匯入範例');
  window.XLSX.writeFile(wb, `進攻對象匯入範例_${new Date().toISOString().slice(0,10)}.xlsx`);
};

// --- 2. 樣式 (Styles) ---
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
  buildings: [{ name: '行政大樓', code: 'A1' }, { name: '博雅教學館', code: 'B2' }],
  machineTypes: ['彩色影印機', '黑白影印機', '複合列表機', '單工列表機'],
  equipmentDB: [{ brand: 'HP', model: 'M479fdw', type: '複合列表機' }, { brand: 'Canon', model: 'imageRUNNER DX C357i', type: '彩色影印機' }],
  guidelines: [{ id: 1, title: '學術/行政分組原則', content: '學術單位需確認是否為「獨立空間」。' }, { id: 2, title: '本牌客戶結案原則', content: '一旦確認為本牌 (EIP 資料建立)，則該筆進攻對象結案。' }],
  talkScripts: [{ id: 3, title: '初次拜訪', content: '我們提供節能、高效率的設備，協助貴單位達成綠色採購目標。' }, { id: 4, title: '設備汰換', content: '提供最新的複合機，搭配客製化維護合約，降低運營成本。' }],
  areaMap: [],
  uploadedMapUrl: 'https://drive.google.com/thumbnail?id=1fmrcmaTr3qSeccln8If59g_eoPnDDY4J&sz=w3000',
  customScheduleColumns: []
};

// --- 3. Hooks (Moved to Top) ---
const useExcelExport = () => {
  useEffect(() => {
    if (typeof window.XLSX === 'undefined') {
      const script = document.createElement('script');
      script.src = 'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js';
      document.head.appendChild(script);
    }
  }, []);

  const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  };

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
          let value = row[key];
          if (key.startsWith('customData.') && row.customData) {
              value = row.customData[key.split('.')[1]];
          }
          if (Array.isArray(value)) {
            return value.map((item) => typeof item === 'object' ? JSON.stringify(item) : item).join('; ');
          }
          return value !== undefined ? value : '';
        })
      ),
    ];
    const ws = window.XLSX.utils.aoa_to_sheet(worksheetData);
    ws['!cols'] = columnHeaders.map((h) => ({ wch: h.width || 20 }));
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const wbout = window.XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}_${new Date().toISOString().substring(0, 10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  return exportToExcel;
};

// --- 4. 共用 UI 組件 (UI Components) ---

const StatusCard = ({ title, value, icon, gradient }) => (
  <div className={`relative p-6 rounded-2xl shadow-lg text-white bg-gradient-to-br ${gradient} overflow-hidden transform hover:-translate-y-1 transition duration-300`}>
    <div className="absolute top-0 right-0 p-4 opacity-20 transform scale-150">{icon}</div>
    <p className="text-sm font-medium opacity-90 tracking-wide">{title}</p>
    <p className="text-4xl font-extrabold mt-2 tracking-tight">{value}</p>
  </div>
);

const InputGroup = ({ label, children }) => (
  <div className="flex flex-col space-y-1.5">
    <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">{label}</label>
    {children}
  </div>
);

const FilterSelect = ({ label, value, onChange, children }) => (
  <div className="flex flex-col">
    <label className="text-xs font-bold text-slate-500 mb-1">{label}</label>
    <select value={value} onChange={onChange} className={`${styles.formSelect} text-sm`}>{children}</select>
  </div>
);

const FilterInput = ({ label, value, onChange }) => (
  <div className="flex flex-col">
    <label className="text-xs font-bold text-slate-500 mb-1">{label}</label>
    <input type="text" value={value} onChange={onChange} className={`${styles.formInput} text-sm`}/>
  </div>
);

// --- 5. 業務邏輯組件 (Business Components) ---

const UnitTable = ({ units, selectedUnitIds, setSelectedUnitIds, onViewUnit }) => {
  return (
    <div className="overflow-x-auto rounded-xl border border-gray-200">
      <table className="min-w-full divide-y divide-gray-200">
        <thead className="bg-gray-50 text-gray-500">
          <tr>
            <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">選取</th>
            <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">單位資訊</th>
            <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">位置</th>
            <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">主要聯絡人</th>
            <th className="p-3 text-left text-xs font-bold uppercase tracking-wider">狀態/分組</th>
            <th className="p-3 text-right text-xs font-bold uppercase tracking-wider">預覽</th>
          </tr>
        </thead>
        <tbody className="bg-white divide-y divide-gray-100">
          {units.map((unit) => (
            <tr key={unit.id} className="hover:bg-indigo-50/40 transition">
              <td className="p-3">
                <input type="checkbox" checked={selectedUnitIds.includes(unit.id)} onChange={() => setSelectedUnitIds((p) => p.includes(unit.id) ? p.filter((id) => id !== unit.id) : [...p, unit.id])} className={styles.checkbox} />
              </td>
              <td className="p-3">
                <div className="text-sm font-bold text-gray-900">{unit.name}</div>
                <div className="text-xs text-gray-500">
                  <span className={`inline-block mr-1 px-1.5 rounded ${unit.category === 'Academic' ? 'bg-sky-100 text-sky-700' : 'bg-orange-100 text-orange-700'}`}>{unit.category === 'Academic' ? '學術' : '行政'}</span>
                </div>
              </td>
              <td className="p-3 text-sm text-gray-600">
                <div className="font-medium text-indigo-900">{unit.buildingId || '-'} 棟</div>
                <div className="text-xs">{unit.floor ? `${unit.floor}` : ''} {unit.roomNumber ? `${unit.roomNumber}` : ''}</div>
              </td>
              <td className="p-3 text-sm text-gray-600">
                <div className="font-bold text-slate-700">{unit.contactName1 || unit.contactName || '-'}</div>
                <div className="text-xs text-gray-400">{unit.contactPhone1 || unit.contactPhone || '-'}</div>
              </td>
              <td className="p-3 text-sm">
                <span className={`px-2 py-0.5 inline-flex text-xs leading-5 font-bold rounded-full ${unit.attackStatus === 'client' ? 'bg-emerald-100 text-emerald-800' : unit.attackStatus === 'settled_non_client' ? 'bg-rose-100 text-rose-800' : 'bg-amber-100 text-amber-800'}`}>
                  {unit.attackStatus === 'client' ? '本牌客戶' : unit.attackStatus === 'settled_non_client' ? '暫定結案' : '進攻中'}
                </span>
                {unit.subgroup && <div className="text-xs text-gray-500 mt-1">({unit.subgroup})</div>}
              </td>
              <td className="p-3 text-right">
                <button onClick={() => onViewUnit(unit)} className="px-3 py-1.5 text-xs font-medium bg-white text-indigo-600 rounded-lg hover:bg-indigo-50 transition border border-indigo-200 flex items-center ml-auto">
                  <Eye className="w-3 h-3 mr-1"/> 詳細
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {units.length === 0 && (<div className="p-8 text-center text-gray-400 bg-gray-50">尚無資料。</div>)}
    </div>
  );
};

const UnitPreviewModal = ({ unit, onClose }) => {
    const equipment = safeParse(unit.equipment || '[]');
    const history = safeParse(unit.history || '[]');

    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
             <div className="bg-white rounded-2xl shadow-2xl w-full max-w-5xl max-h-[90vh] overflow-y-auto border border-slate-200">
                <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-indigo-50/50 sticky top-0">
                    <h3 className="text-xl font-bold text-indigo-900 flex items-center">
                        <MapPin className="w-6 h-6 mr-2 text-indigo-600"/> 
                        {unit.name} 
                        <span className="ml-3 text-sm bg-indigo-100 text-indigo-700 px-3 py-1 rounded-full font-medium">{unit.category === 'Academic' ? '學術' : '行政'}</span>
                    </h3>
                    <button onClick={onClose} className="p-2 hover:bg-indigo-100 text-indigo-600 rounded-full transition"><X className="w-6 h-6"/></button>
                </div>

                <div className="p-8 space-y-8">
                    <div className="grid grid-cols-1 md:grid-cols-4 gap-6 bg-white rounded-xl">
                        <div className="md:col-span-1 space-y-4">
                             <div className="bg-gray-50 p-4 rounded-xl border border-gray-100">
                                 <h4 className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-3">位置資訊</h4>
                                 <p className="text-sm text-gray-800 mb-1"><span className="font-semibold">棟別:</span> {unit.buildingId}</p>
                                 <p className="text-sm text-gray-800 mb-1"><span className="font-semibold">樓層:</span> {unit.floor || '-'}</p>
                                 <p className="text-sm text-gray-800 mb-1"><span className="font-semibold">房號:</span> {unit.roomNumber || '-'}</p>
                             </div>
                             <div className="bg-gray-50 p-4 rounded-xl border border-gray-100">
                                 <h4 className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-3">狀態資訊</h4>
                                 <span className={`inline-block px-3 py-1 text-xs font-bold rounded-full ${unit.attackStatus === 'client' ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'}`}>
                                     {unit.attackStatus === 'client' ? '本牌客戶' : unit.attackStatus === 'settled_non_client' ? '暫定結案' : '進攻中'}
                                 </span>
                                 <p className="text-sm text-gray-600 mt-2">分組: {unit.subgroup || '一般'}</p>
                             </div>
                        </div>

                        <div className="md:col-span-3">
                             <div className="bg-slate-50 p-5 rounded-xl border border-slate-200 mb-6">
                                 <h4 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-3 flex items-center"><User className="w-4 h-4 mr-2"/> 聯絡人資訊</h4>
                                 <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                                     <div className="bg-white p-3 rounded-lg shadow-sm border border-slate-100 border-l-4 border-l-indigo-400">
                                         <span className="text-xs text-indigo-500 font-bold mb-1 block">主要負責人</span>
                                         <p className="font-bold text-gray-800">{unit.contactName1 || '-'}</p>
                                         <p className="text-sm text-gray-500">{unit.contactPhone1 || '-'}</p>
                                     </div>
                                     <div className="bg-white p-3 rounded-lg shadow-sm border border-slate-100">
                                         <span className="text-xs text-gray-400 font-bold mb-1 block">第二聯絡人</span>
                                         <p className="font-bold text-gray-700">{unit.contactName2 || '-'}</p>
                                         <p className="text-sm text-gray-500">{unit.contactPhone2 || '-'}</p>
                                     </div>
                                     <div className="bg-white p-3 rounded-lg shadow-sm border border-slate-100">
                                         <span className="text-xs text-gray-400 font-bold mb-1 block">第三聯絡人</span>
                                         <p className="font-bold text-gray-700">{unit.contactName3 || '-'}</p>
                                         <p className="text-sm text-gray-500">{unit.contactPhone3 || '-'}</p>
                                     </div>
                                 </div>
                             </div>

                             <h4 className="text-lg font-bold text-slate-800 mb-4 flex items-center">
                                 <Printer className="w-5 h-5 mr-2 text-indigo-600"/> 設備與履歷預覽
                             </h4>
                             <div className="space-y-4">
                                 {(equipment || []).map((eq, idx) => (
                                     <div key={idx} className="border border-indigo-100 rounded-xl overflow-hidden shadow-sm">
                                         <div className="bg-gradient-to-r from-indigo-50 to-white p-4 flex justify-between items-center border-b border-indigo-50">
                                             <div>
                                                 <span className="font-bold text-slate-800 text-lg">{eq.brand} {eq.model}</span>
                                                 <span className="ml-2 text-xs bg-white border border-indigo-200 text-indigo-600 px-2 py-0.5 rounded-full">{eq.type}</span>
                                             </div>
                                             <div className="text-right text-sm text-slate-600">
                                                 <span className="mr-3 bg-white px-2 py-1 rounded border border-gray-100">方案: <span className="font-semibold text-indigo-600">{eq.plan || '未指定'}</span></span>
                                                 <span className="bg-white px-2 py-1 rounded border border-gray-100">廠商: <span className="font-semibold text-indigo-600">{eq.vendor || '未指定'}</span></span>
                                             </div>
                                         </div>
                                         <div className="p-4 bg-white">
                                             <h5 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2 flex items-center">
                                                <History className="w-3 h-3 mr-1"/> 最近 10 筆相關拜訪
                                             </h5>
                                             <div className="space-y-1">
                                                 {history
                                                     ?.filter(h => h.relatedEquipmentId === eq.id)
                                                     .sort((a, b) => new Date(b.date) - new Date(a.date))
                                                     .slice(0, 10)
                                                     .map(log => (
                                                         <div key={log.id} className="text-sm grid grid-cols-12 gap-2 text-slate-600 hover:bg-slate-50 p-1.5 rounded transition">
                                                             <span className="col-span-3 font-mono text-xs text-slate-400">{log.date}</span>
                                                             <span className="col-span-9 font-medium">{log.activity}</span>
                                                         </div>
                                                     ))
                                                 }
                                                 {(!history?.some(h => h.relatedEquipmentId === eq.id)) && (
                                                     <p className="text-xs text-gray-300 italic py-1">此設備尚無專屬拜訪紀錄</p>
                                                 )}
                                             </div>
                                         </div>
                                     </div>
                                 ))}
                                 {(equipment || []).length === 0 && <p className="text-center text-gray-400 py-4 bg-gray-50 rounded-xl">尚無設備資料</p>}
                             </div>
                        </div>
                    </div>
                </div>
             </div>
        </div>
    );
};

const EquipmentAdder = ({ availableBrands, availableModels, machineTypes, onAdd, equipmentSearch, setEquipmentSearch }) => {
    const [plan, setPlan] = useState('');
    const [vendor, setVendor] = useState('');
    return (
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 bg-white p-4 rounded-lg shadow-sm border border-indigo-100 mb-4">
            <select className={styles.formSelect} value={equipmentSearch.brand} onChange={e => setEquipmentSearch(p => ({...p, brand: e.target.value}))}>
                <option value="">選擇廠牌</option>
                {availableBrands.map(b => <option key={b} value={b}>{b}</option>)}
            </select>
            <select className={styles.formSelect} value={equipmentSearch.model} onChange={e => setEquipmentSearch(p => ({...p, model: e.target.value}))}>
                <option value="">選擇型號</option>
                {availableModels.map(m => <option key={m} value={m}>{m}</option>)}
            </select>
            <input className={styles.formInput} placeholder="目前方案" value={plan} onChange={e => setPlan(e.target.value)} />
            <input className={styles.formInput} placeholder="目前廠商" value={vendor} onChange={e => setVendor(e.target.value)} />
            <button onClick={() => { if(equipmentSearch.brand && equipmentSearch.model) { onAdd({brand: equipmentSearch.brand, model: equipmentSearch.model, plan, vendor, type: '影印機'}); setPlan(''); setVendor(''); } }} className={`${styles.btnPrimary} col-span-2 md:col-span-4`}>新增設備</button>
        </div>
    );
};

const EquipmentList = ({ equipment, setNewUnitData, history }) => (
    <div className="space-y-3">
        {equipment.map(eq => (
            <div key={eq.id} className="p-3 bg-white border border-indigo-200 rounded-lg flex justify-between items-center">
                <div><span className="font-bold">{eq.brand} {eq.model}</span> <span className="text-xs text-gray-500">({eq.plan})</span></div>
                <button onClick={() => setNewUnitData(p => ({...p, equipment: p.equipment.filter(e => e.id !== eq.id)}))} className="text-red-400 p-1"><Trash2 className="w-4 h-4"/></button>
            </div>
        ))}
    </div>
);

const CharacteristicsEditor = ({ characteristics, setNewUnitData }) => {
    const options = ["對價格敏感", "重視售後服務", "偏好特定廠牌", "有自行維修能力", "決策緩慢", "預算充足"];
    const toggle = (val) => {
        const next = characteristics.includes(val) ? characteristics.filter(c => c !== val) : [...characteristics, val];
        setNewUnitData(p => ({...p, characteristics: next}));
    };
    return (
        <div className="flex flex-wrap gap-2">
            {options.map(opt => (
                <button key={opt} onClick={() => toggle(opt)} className={`px-3 py-1 text-xs rounded-full border transition ${characteristics.includes(opt) ? 'bg-amber-500 text-white border-amber-600' : 'bg-white text-amber-700 border-amber-200 hover:bg-amber-50'}`}>{opt}</button>
            ))}
        </div>
    );
};

const HistoryLogAdder = ({ onAdd, equipmentList }) => {
    const [activity, setActivity] = useState('');
    const [relatedId, setRelatedId] = useState('');
    return (
        <div className="space-y-3 mb-4">
            <select className={styles.formSelect} value={relatedId} onChange={e => setRelatedId(e.target.value)}>
                <option value="">選擇關聯設備 (選填)</option>
                {equipmentList.map(e => <option key={e.id} value={e.id}>{e.brand} {e.model}</option>)}
            </select>
            <textarea className={styles.formTextarea} placeholder="輸入拜訪紀錄內容..." value={activity} onChange={e => setActivity(e.target.value)} rows="2" />
            <button onClick={() => { if(activity) { onAdd({activity, relatedEquipmentId: relatedId}); setActivity(''); setRelatedId(''); } }} className={styles.btnPrimary}>紀錄拜訪</button>
        </div>
    );
};

const HistoryLogList = ({ history }) => (
    <div className="space-y-2 mt-4 max-h-60 overflow-y-auto">
        {history.sort((a,b) => new Date(b.date) - new Date(a.date)).map(h => (
            <div key={h.id} className="text-xs p-2 bg-white rounded border border-emerald-100 flex justify-between">
                <span className="text-emerald-700 font-mono">{h.date}</span>
                <span className="flex-1 ml-3 text-gray-700">{h.activity}</span>
            </div>
        ))}
    </div>
);

// --- 6. 分頁組件 (Tab Components) ---

// 6.1 進攻行事曆 Tab
const Tab1Calendar = ({ appData, updatePrivateData, db, userId, setGlobalMessage, exportToExcel }) => {
    const [selectedScheduleIds, setSelectedScheduleIds] = useState([]);
    const [scheduleRows, setScheduleRows] = useState([]);
    const [newColumnName, setNewColumnName] = useState('');
    const [bulkAddCount, setBulkAddCount] = useState(1);
    const [selectedMeetingIds, setSelectedMeetingIds] = useState([]);
    const [isAddingMeeting, setIsAddingMeeting] = useState(false);
    const [editingMeetingId, setEditingMeetingId] = useState(null);
    const [copyModalContent, setCopyModalContent] = useState(null);
    const [uploadFileName, setUploadFileName] = useState('');

    useEffect(() => {
        if(appData.schedules) setScheduleRows(appData.schedules);
    }, [appData.schedules]);

    const customColumns = appData.settings.customScheduleColumns || [];

    const handleAddColumn = () => {
        if (!newColumnName) return;
        const newCol = { id: crypto.randomUUID(), label: newColumnName };
        updatePrivateData({ customScheduleColumns: [...customColumns, newCol] });
        setNewColumnName('');
    };

    const handleDeleteColumn = (colId) => {
        if (confirm('確定刪除此欄位嗎？')) {
            updatePrivateData({ customScheduleColumns: customColumns.filter(c => c.id !== colId) });
        }
    };

    const handleBulkAddSchedules = () => {
        const newRows = Array.from({ length: bulkAddCount }).map(() => ({
            id: crypto.randomUUID(), startDate: '', endDate: '', personnel: '', resourceContent: '', resource1: '', resource2: '', memo: '', customData: {}
        }));
        setScheduleRows(prev => [...prev, ...newRows]);
    };

    const handleScheduleChange = (id, field, value, isCustom = false) => {
        setScheduleRows(prevRows => prevRows.map(row => {
            if (row.id === id) {
                if (isCustom) return { ...row, customData: { ...row.customData, [field]: value } };
                return { ...row, [field]: value };
            }
            return row;
        }));
    };

    const handleSaveSchedules = () => {
        updatePrivateData({ schedules: scheduleRows });
        alert('排程已儲存');
    };

    const handleDeleteSchedules = () => {
        if (selectedScheduleIds.length === 0) return;
        if (confirm(`確定刪除選取的 ${selectedScheduleIds.length} 筆排程？`)) {
            const updatedSchedules = scheduleRows.filter(s => !selectedScheduleIds.includes(s.id));
            setScheduleRows(updatedSchedules);
            updatePrivateData({ schedules: updatedSchedules });
            setSelectedScheduleIds([]);
        }
    };

    const exportSchedules = () => {
        const headers = [
            { key: 'startDate', label: '開始日期', width: 15 },
            { key: 'endDate', label: '結束日期', width: 15 },
            { key: 'personnel', label: '人員/家數/區域', width: 30 },
            { key: 'resourceContent', label: '資源配給內容', width: 20 },
            { key: 'resource1', label: '配給月份/數量(1)', width: 15 },
            { key: 'resource2', label: '配給月份/數量(2)', width: 15 },
            { key: 'memo', label: '待辦&備忘', width: 40 },
            ...customColumns.map(col => ({ key: `customData.${col.id}`, label: col.label, width: 20 }))
        ];
        exportToExcel(scheduleRows, '進攻排程表', '排程', headers);
    };

    const getDayOfWeek = (dateStr) => {
        if (!dateStr) return '';
        const date = new Date(dateStr);
        const days = ['日', '一', '二', '三', '四', '五', '六'];
        return `(${days[date.getDay()]})`;
    };

    const deleteSelectedMeetings = () => {
      if (selectedMeetingIds.length === 0) return;
      const updatedMeetings = appData.meetings.filter(m => !selectedMeetingIds.includes(m.id));
      updatePrivateData({ meetings: updatedMeetings });
      setSelectedMeetingIds([]);
    };

    const exportMeetings = () => {
      const headers = [
        { key: 'date', label: '本次會議日期', width: 15 },
        { key: 'attendees', label: '與會人員', width: 20 },
        { key: 'summary', label: '總結', width: 40 },
        { key: 'nextMeetingDate', label: '下次預計會議', width: 15 },
        { key: 'nextAttendees', label: '下次預訂與會人', width: 20 },
        { key: 'nextTopics', label: '下次議題', width: 30 },
      ];
      exportToExcel(appData.meetings, '戰勤會議紀錄', '會議紀錄', headers);
    };

    const handleFileUpload = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        if (!db || !userId) { alert("資料庫未連線，請重新整理頁面"); return; } 
        if (file.size > 800 * 1024) { alert('檔案過大！請上傳小於 800KB 的檔案。'); return; }

        const finalName = uploadFileName.trim() || file.name;
        const reader = new FileReader();
        reader.onload = async (event) => {
            const base64String = event.target.result;
            try {
                const filesRef = collection(db, 'artifacts', appId, 'users', userId, 'files');
                await addDoc(filesRef, {
                    name: finalName,
                    originalName: file.name,
                    type: file.type,
                    data: base64String,
                    createdAt: new Date().toISOString()
                });
                setGlobalMessage({ text: '檔案上傳成功', type: 'success' });
                setUploadFileName('');
            } catch (err) {
                console.error("Upload failed", err);
                alert("上傳失敗: " + err.message);
            }
        };
        reader.readAsDataURL(file);
    };

    const handleDeleteFile = async (fileId) => {
        if(confirm('確定刪除此檔案？')) {
            try { await deleteDoc(doc(db, 'artifacts', appId, 'users', userId, 'files', fileId)); } catch(e) { console.error(e); }
        }
    };

    const MeetingRow = ({ meeting }) => {
        const isEditing = editingMeetingId === meeting.id;
        const [editData, setEditData] = useState(meeting);
        const handleSave = () => {
            const updated = appData.meetings.map(m => m.id === meeting.id ? editData : m);
            updatePrivateData({ meetings: updated });
            setEditingMeetingId(null);
        };
        const handleCancel = () => { setEditData(meeting); setEditingMeetingId(null); };
        const copyContent = () => {
             const text = `【戰勤會議紀錄】\n📅 本次日期：${meeting.date}\n👥 與會：${meeting.attendees}\n📝 總結：\n${meeting.summary}\n\n📅 下次會議：${meeting.nextMeetingDate}\n👥 下次與會：${meeting.nextAttendees}\n💡 下次議題：\n${meeting.nextTopics}`;
             setCopyModalContent(text);
        };

        return (
            <tr className="hover:bg-indigo-50/30 transition border-b border-gray-100">
                <td className="p-3 align-top"><input type="checkbox" checked={selectedMeetingIds.includes(meeting.id)} onChange={() => setSelectedMeetingIds(p => p.includes(meeting.id) ? p.filter(id=>id!==meeting.id) : [...p, meeting.id])} className={styles.checkbox}/></td>
                <td className="p-2 align-top">{isEditing ? <input type="date" value={editData.date} onChange={e=>setEditData({...editData, date: e.target.value})} className={styles.formInput}/> : meeting.date}</td>
                <td className="p-2 align-top">{isEditing ? <input type="text" value={editData.attendees} onChange={e=>setEditData({...editData, attendees: e.target.value})} className={styles.formInput}/> : meeting.attendees}</td>
                <td className="p-2 align-top">{isEditing ? <textarea value={editData.summary} onChange={e=>setEditData({...editData, summary: e.target.value})} className={styles.formTextarea} rows={3}/> : <div className="whitespace-pre-wrap">{meeting.summary}</div>}</td>
                <td className="p-2 align-top">{isEditing ? <input type="date" value={editData.nextMeetingDate} onChange={e=>setEditData({...editData, nextMeetingDate: e.target.value})} className={styles.formInput}/> : meeting.nextMeetingDate}</td>
                <td className="p-2 align-top">{isEditing ? <input type="text" value={editData.nextAttendees} onChange={e=>setEditData({...editData, nextAttendees: e.target.value})} className={styles.formInput}/> : meeting.nextAttendees}</td>
                <td className="p-2 align-top min-w-[200px]">{isEditing ? <textarea value={editData.nextTopics} onChange={e=>setEditData({...editData, nextTopics: e.target.value})} className={styles.formTextarea} rows={4}/> : <div className="whitespace-pre-wrap">{meeting.nextTopics}</div>}</td>
                <td className="p-2 align-top text-right">
                    {isEditing ? (
                        <div className="flex flex-col gap-1"><button onClick={handleSave} className="p-1 text-emerald-600 bg-emerald-50 rounded"><Save className="w-4 h-4"/></button><button onClick={handleCancel} className="p-1 text-gray-500 bg-gray-50 rounded"><X className="w-4 h-4"/></button></div>
                    ) : (
                        <div className="flex flex-col gap-1"><button onClick={() => setEditingMeetingId(meeting.id)} className="p-1 text-indigo-600 bg-indigo-50 rounded"><Edit className="w-4 h-4"/></button><button onClick={copyContent} className="p-1 text-gray-600 bg-gray-100 rounded" title="複製"><Copy className="w-4 h-4"/></button></div>
                    )}
                </td>
            </tr>
        );
    };

    const AddMeetingForm = () => {
        const [newM, setNewM] = useState({ id: '', date: '', attendees: '', summary: '', nextMeetingDate: '', nextAttendees: '', nextTopics: '' });
        
        // Generate Preview Text Live
        const previewText = `【戰勤會議紀錄】\n` +
            `📅 本次日期：${newM.date || '(未填)'}\n` +
            `👥 與會人員：${newM.attendees || '(未填)'}\n` +
            `📝 會議總結：\n${newM.summary || '(無)'}\n\n` +
            `📅 下次預計：${newM.nextMeetingDate || '(未定)'}\n` +
            `👥 下次與會：${newM.nextAttendees || '(同上)'}\n` +
            `💡 下次議題：\n${newM.nextTopics || '(無)'}`;

        const add = () => {
            if(!newM.date) { alert('請填寫日期'); return; }
            updatePrivateData({ meetings: [...appData.meetings, { ...newM, id: crypto.randomUUID() }] });
            setIsAddingMeeting(false);
        };

        const copyPreview = () => {
            const el = document.createElement('textarea');
            el.value = previewText;
            document.body.appendChild(el);
            el.select();
            document.execCommand('copy');
            document.body.removeChild(el);
            alert('內容已複製！');
        };

        return (
            <div className="p-4 bg-blue-50 rounded-lg mb-4 border border-blue-200">
                <h4 className="font-bold text-blue-800 mb-2">新增會議紀錄</h4>
                <div className="grid grid-cols-2 gap-2 mb-2">
                    <InputGroup label="本次開會日期"><input type="date" className={styles.formInput} value={newM.date} onChange={e=>setNewM({...newM, date: e.target.value})} /></InputGroup>
                    <InputGroup label="與會人員"><input type="text" className={styles.formInput} placeholder="與會人員" value={newM.attendees} onChange={e=>setNewM({...newM, attendees: e.target.value})} /></InputGroup>
                </div>
                <div className="mb-2"><InputGroup label="會議總結"><textarea className={styles.formTextarea} placeholder="總結" rows={3} value={newM.summary} onChange={e=>setNewM({...newM, summary: e.target.value})} /></InputGroup></div>
                <div className="grid grid-cols-3 gap-2 mb-4">
                    <InputGroup label="下次預計日期"><input type="date" className={styles.formInput} value={newM.nextMeetingDate} onChange={e=>setNewM({...newM, nextMeetingDate: e.target.value})} /></InputGroup>
                    <InputGroup label="下次預訂與會"><input type="text" className={styles.formInput} placeholder="人員" value={newM.nextAttendees} onChange={e=>setNewM({...newM, nextAttendees: e.target.value})} /></InputGroup>
                    <InputGroup label="下次議題"><textarea className={styles.formTextarea} placeholder="議題" rows={3} value={newM.nextTopics} onChange={e=>setNewM({...newM, nextTopics: e.target.value})} /></InputGroup>
                </div>
                
                {/* PREVIEW BLOCK */}
                <div className="bg-white border-2 border-dashed border-indigo-200 rounded-lg p-3 mb-4 relative group">
                    <div className="flex justify-between items-center mb-1">
                        <span className="text-xs font-bold text-indigo-500 uppercase">Line 訊息預覽 (確認內容後可直接複製)</span>
                        <button onClick={copyPreview} className="text-xs flex items-center bg-indigo-100 text-indigo-700 px-2 py-1 rounded hover:bg-indigo-200 transition"><Copy className="w-3 h-3 mr-1"/> 一鍵複製</button>
                    </div>
                    <pre className="text-xs text-gray-600 whitespace-pre-wrap font-mono bg-gray-50 p-2 rounded">{previewText}</pre>
                </div>

                <div className="flex justify-end gap-2">
                    <button onClick={()=>setIsAddingMeeting(false)} className={styles.btnSecondary}>取消</button>
                    <button onClick={add} className={styles.btnPrimary}>確認新增</button>
                </div>
            </div>
        )
    };

    return (
      <div className="space-y-12 p-6 max-w-[95%] mx-auto">
        <div className="flex items-center space-x-3 mb-2">
          <Activity className="w-8 h-8 text-indigo-600" />
          <h2 className="text-3xl font-extrabold text-slate-800 tracking-tight">進攻行事曆</h2>
        </div>

        {/* SCHEDULE */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100">
            <div className="p-6 bg-gradient-to-r from-orange-100 to-amber-50 border-b border-orange-200 flex flex-wrap justify-between items-center gap-4">
                <div>
                    <h3 className="text-xl font-bold text-orange-900 flex items-center"><CalendarIcon className="w-5 h-5 mr-2" /> 進攻排程</h3>
                    <p className="text-xs text-orange-700 mt-1">批量新增以快速編輯，編輯後請記得按儲存</p>
                </div>
                <div className="flex items-center gap-2 flex-wrap">
                    <div className="flex items-center bg-white rounded-lg border border-orange-200 p-1">
                        <input type="number" min="1" max="20" value={bulkAddCount} onChange={e=>setBulkAddCount(parseInt(e.target.value)||1)} className="w-12 text-center outline-none text-sm font-bold text-orange-600"/>
                        <button onClick={handleBulkAddSchedules} className="px-3 py-1.5 bg-orange-500 hover:bg-orange-600 text-white text-sm font-medium rounded transition flex items-center"><Plus className="w-4 h-4 mr-1"/> 新增列</button>
                    </div>
                    <button onClick={handleSaveSchedules} className={`${styles.btnPrimary} bg-emerald-600 hover:bg-emerald-700`}><Save className="w-4 h-4 mr-1"/> 儲存變更</button>
                    <button onClick={handleDeleteSchedules} disabled={selectedScheduleIds.length===0} className={styles.btnDanger}><Trash2 className="w-4 h-4 mr-1"/> 刪除選取</button>
                    <button onClick={exportSchedules} className={styles.btnInfo}><Download className="w-4 h-4 mr-1"/> 匯出</button>
                </div>
            </div>
            <div className="px-6 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
                <Columns className="w-4 h-4 text-slate-500"/>
                <span className="text-sm font-bold text-slate-600">擴充欄位:</span>
                <input type="text" placeholder="輸入新欄位標題..." value={newColumnName} onChange={e=>setNewColumnName(e.target.value)} className="px-3 py-1 text-sm border border-slate-300 rounded focus:ring-2 focus:ring-indigo-500 outline-none w-48"/>
                <button onClick={handleAddColumn} className="px-3 py-1 bg-slate-200 hover:bg-slate-300 text-slate-700 text-sm rounded transition">新增</button>
            </div>
            <div className="overflow-x-auto pb-4">
                <table className="min-w-full divide-y divide-gray-200 border-collapse">
                    <thead className="bg-slate-50">
                        <tr>
                            <th className="p-3 w-10 text-center sticky left-0 bg-slate-50 z-10 border-r"><input type="checkbox" className={styles.checkbox} disabled/></th>
                            <th className="p-3 text-left text-xs font-bold text-slate-600 uppercase tracking-wider w-48 bg-slate-100 border-r border-slate-200">時程 (起訖)</th>
                            <th className="p-3 text-left text-xs font-bold text-slate-600 uppercase tracking-wider min-w-[200px] bg-indigo-50 border-r border-indigo-100">人員 / 家數 / 區域</th>
                            <th className="p-3 text-left text-xs font-bold text-slate-600 uppercase tracking-wider w-32 bg-emerald-50 border-r border-emerald-100">資源配給內容</th>
                            <th className="p-3 text-left text-xs font-bold text-slate-600 uppercase tracking-wider w-32 bg-emerald-50 border-r border-emerald-100">配給月份/數量 (1)</th>
                            <th className="p-3 text-left text-xs font-bold text-slate-600 uppercase tracking-wider w-32 bg-emerald-50 border-r border-emerald-100">配給月份/數量 (2)</th>
                            <th className="p-3 text-left text-xs font-bold text-slate-600 uppercase tracking-wider min-w-[250px] bg-amber-50 border-r border-amber-100">待辦 & 備忘</th>
                            {customColumns.map(col => (
                                <th key={col.id} className="p-3 text-left text-xs font-bold text-indigo-600 uppercase tracking-wider min-w-[150px] border-r border-slate-200 group relative">
                                    {col.label} <button onClick={() => handleDeleteColumn(col.id)} className="absolute right-1 top-1/2 -translate-y-1/2 opacity-0 group-hover:opacity-100 text-red-400 hover:text-red-600"><X className="w-3 h-3"/></button>
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody className="bg-white divide-y divide-gray-100">
                        {scheduleRows.map((row) => (
                            <tr key={row.id} className="hover:bg-slate-50 transition group">
                                <td className="p-3 text-center sticky left-0 bg-white group-hover:bg-slate-50 z-10 border-r border-slate-200"><input type="checkbox" checked={selectedScheduleIds.includes(row.id)} onChange={() => setSelectedScheduleIds(p => p.includes(row.id) ? p.filter(id=>id!==row.id) : [...p, row.id])} className={styles.checkbox}/></td>
                                <td className="p-2 border-r border-slate-200">
                                    <div className="flex flex-col gap-1">
                                        <div className="flex items-center text-xs text-slate-500"><span className="w-6">起:</span><input type="date" value={row.startDate || ''} onChange={e=>handleScheduleChange(row.id, 'startDate', e.target.value)} className="bg-transparent outline-none border-b border-transparent focus:border-indigo-300 w-full"/> <span className="ml-1 text-[10px]">{getDayOfWeek(row.startDate)}</span></div>
                                        <div className="flex items-center text-xs text-slate-500"><span className="w-6">訖:</span><input type="date" value={row.endDate || ''} onChange={e=>handleScheduleChange(row.id, 'endDate', e.target.value)} className="bg-transparent outline-none border-b border-transparent focus:border-indigo-300 w-full"/> <span className="ml-1 text-[10px]">{getDayOfWeek(row.endDate)}</span></div>
                                    </div>
                                </td>
                                <td className="p-2 bg-indigo-50/10 border-r border-indigo-100"><textarea value={row.personnel || ''} onChange={e=>handleScheduleChange(row.id, 'personnel', e.target.value)} className="w-full bg-transparent outline-none resize-none text-sm text-slate-700 placeholder-indigo-200" rows={3}/></td>
                                <td className="p-2 bg-emerald-50/10 border-r border-emerald-100"><input type="text" value={row.resourceContent || ''} onChange={e=>handleScheduleChange(row.id, 'resourceContent', e.target.value)} className="w-full bg-transparent outline-none text-sm border-b border-transparent focus:border-emerald-300"/></td>
                                <td className="p-2 bg-emerald-50/10 border-r border-emerald-100"><input type="text" value={row.resource1 || ''} onChange={e=>handleScheduleChange(row.id, 'resource1', e.target.value)} className="w-full bg-transparent outline-none text-sm border-b border-transparent focus:border-emerald-300"/></td>
                                <td className="p-2 bg-emerald-50/10 border-r border-emerald-100"><input type="text" value={row.resource2 || ''} onChange={e=>handleScheduleChange(row.id, 'resource2', e.target.value)} className="w-full bg-transparent outline-none text-sm border-b border-transparent focus:border-emerald-300"/></td>
                                <td className="p-2 bg-amber-50/10 border-r border-amber-100"><textarea value={row.memo || ''} onChange={e=>handleScheduleChange(row.id, 'memo', e.target.value)} className="w-full bg-transparent outline-none resize-none text-sm text-slate-700 placeholder-amber-200" rows={3}/></td>
                                {customColumns.map(col => (
                                    <td key={col.id} className="p-2 border-r border-slate-200"><textarea value={row.customData?.[col.id] || ''} onChange={e=>handleScheduleChange(row.id, col.id, e.target.value, true)} className="w-full bg-transparent outline-none resize-none text-sm" rows={2}/></td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>

        {/* MEETINGS */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100">
            <div className="p-6 bg-gradient-to-r from-blue-100 to-indigo-50 border-b border-blue-200 flex justify-between items-center">
                <h3 className="text-xl font-bold text-blue-900 flex items-center"><Users className="w-5 h-5 mr-2" /> 戰勤會議紀錄</h3>
                <div className="flex gap-2">
                    <button onClick={()=>setIsAddingMeeting(true)} className={styles.btnPrimary}><Plus className="w-4 h-4 mr-1"/> 新增紀錄</button>
                    <button onClick={deleteSelectedMeetings} disabled={selectedMeetingIds.length===0} className={styles.btnDanger}><Trash2 className="w-4 h-4 mr-1"/> 刪除</button>
                    <button onClick={exportMeetings} className={styles.btnInfo}><Download className="w-4 h-4 mr-1"/> 匯出 Excel</button>
                </div>
            </div>
            <div className="p-4">
                {isAddingMeeting && <AddMeetingForm />}
                <div className="overflow-x-auto">
                    <table className="min-w-full divide-y divide-gray-200">
                        <thead className="bg-slate-50 text-slate-500">
                            <tr>
                                <th className="p-3 text-left w-10"></th>
                                <th className="p-3 text-left text-xs font-bold uppercase w-32">本次會議日期</th>
                                <th className="p-3 text-left text-xs font-bold uppercase w-32">與會人員</th>
                                <th className="p-3 text-left text-xs font-bold uppercase min-w-[200px]">總結</th>
                                <th className="p-3 text-left text-xs font-bold uppercase w-32">下次預計會議</th>
                                <th className="p-3 text-left text-xs font-bold uppercase w-32">下次與會人</th>
                                <th className="p-3 text-left text-xs font-bold uppercase min-w-[250px]">下次議題</th>
                                <th className="p-3 text-right text-xs font-bold uppercase w-20">操作</th>
                            </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-gray-100">
                            {appData.meetings.map(meeting => <MeetingRow key={meeting.id} meeting={meeting} />)}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        {/* Copy Modal */}
        {copyModalContent && (
            <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
                <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center"><MessageSquare className="w-5 h-5 mr-2 text-indigo-600"/> 複製到 Line</h3>
                    <textarea readOnly className="w-full h-64 p-3 bg-gray-50 border border-gray-200 rounded-lg text-sm font-mono mb-4 focus:outline-none resize-none" value={copyModalContent}></textarea>
                    <div className="flex justify-end gap-2">
                        <button onClick={()=>setCopyModalContent(null)} className={styles.btnSecondary}>關閉</button>
                        <button onClick={handleCopyText} className={`${styles.btnPrimary} bg-emerald-600 hover:bg-emerald-700`}><Copy className="w-4 h-4 mr-1"/> 複製</button>
                    </div>
                </div>
            </div>
        )}

        {/* ARCHIVES */}
        <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100">
            <div className="p-6 bg-gradient-to-r from-slate-100 to-gray-50 border-b border-slate-200 flex flex-col md:flex-row justify-between items-center gap-4">
                <h3 className="text-xl font-bold text-slate-800 flex items-center"><FileText className="w-5 h-5 mr-2" /> 資料備存 (文件/圖片)</h3>
                <div className="flex items-center gap-2 w-full md:w-auto bg-white p-2 rounded-xl border border-slate-200 shadow-sm">
                    <div className="flex-grow">
                        <label className="text-xs font-bold text-slate-400 ml-1 block mb-1">自訂檔名 (選填)</label>
                        <input type="text" placeholder="留空則使用原檔名" value={uploadFileName} onChange={(e) => setUploadFileName(e.target.value)} className="w-full px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:border-indigo-400 transition" />
                    </div>
                    <div className="relative overflow-hidden group flex-shrink-0 self-end">
                        <button className={`${styles.btnPrimary} bg-slate-700 hover:bg-slate-800 h-10 px-4`}><UploadCloud className="w-4 h-4 mr-2"/> 上傳</button>
                        <input type="file" onChange={handleFileUpload} accept="image/*,application/pdf,.doc,.docx,.xls,.xlsx" className="absolute inset-0 opacity-0 cursor-pointer" />
                    </div>
                </div>
            </div>
            
            <div className="p-6">
                <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-4">
                    {appData.files && appData.files.map(file => (
                        <div key={file.id} className="border border-slate-200 rounded-xl p-4 hover:shadow-md transition bg-slate-50 relative group">
                            <button onClick={()=>handleDeleteFile(file.id)} className="absolute top-2 right-2 p-1.5 text-slate-400 hover:text-red-500 opacity-0 group-hover:opacity-100 transition"><Trash2 className="w-4 h-4"/></button>
                            <div className="flex items-center justify-center h-24 mb-3 bg-white rounded border border-slate-100 overflow-hidden">
                                {file.type.startsWith('image/') ? <img src={file.data} alt={file.name} className="h-full object-contain"/> : <FileText className="w-12 h-12 text-slate-400"/>}
                            </div>
                            <p className="text-sm font-bold text-slate-700 truncate mb-1" title={file.name}>{file.name}</p>
                            <p className="text-xs text-slate-400 mb-3">{new Date(file.createdAt).toLocaleDateString()}</p>
                            <a href={file.data} download={file.name} className="block w-full text-center py-2 bg-white border border-slate-300 text-slate-600 rounded-lg text-sm hover:bg-slate-100 transition">下載</a>
                        </div>
                    ))}
                    {(!appData.files || appData.files.length === 0) && (
                        <div className="col-span-full py-8 text-center text-slate-400 border-2 border-dashed border-slate-200 rounded-xl"><UploadCloud className="w-10 h-10 mx-auto mb-2 opacity-50"/><p>尚無備存資料，請點擊右上角上傳</p><p className="text-xs mt-1 text-slate-300">(支援圖片、PDF、Office文件，限制 800KB 以下)</p></div>
                    )}
                </div>
            </div>
        </div>
      </div>
    );
};

// 6.2 戰情地圖 Tab (Reused Components for other tabs)
const Tab3TargetsMap = ({ appData, setEditingUnitId, setIsNewUnit, setPreviewUnit, setShowAddUnitModal, exportToExcel, deleteUnits, filter, setFilter, viewMode = 'map' }) => {
    const { units, settings } = appData;
    const [selectedUnitIds, setSelectedUnitIds] = useState([]);
    const [zoom, setZoom] = useState(0.7);
    const filtered = units.filter(u => 
        (filter.type === '' || u.category === filter.type) &&
        (filter.area === '' || (!u.areaCode && filter.area==='') || u.areaCode === filter.area)
    );
    const handleDelete = () => { if(confirm('確定刪除？')) { deleteUnits(selectedUnitIds); setSelectedUnitIds([]); }};

    return (
        <div className="p-6 max-w-7xl mx-auto space-y-6">
            <div className="flex justify-between"><h2 className="text-2xl font-bold flex items-center"><MapPin className="w-6 h-6 mr-2"/> {viewMode === 'map' ? '戰情地圖' : '拜訪紀錄總覽'}</h2></div>
            {viewMode === 'map' && (
                <>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <StatusCard title="總家數" value={units.length} gradient="from-indigo-500 to-purple-600" icon={<Building className="w-5 h-5 text-white"/>}/>
                    <StatusCard title="本牌客戶" value={units.filter(u=>u.attackStatus==='client').length} gradient="from-emerald-500 to-teal-500" icon={<CheckCircle className="w-5 h-5 text-white"/>}/>
                    <StatusCard title="行政單位" value={units.filter(u=>u.category==='Administrative').length} gradient="from-orange-400 to-red-500" icon={<Users className="w-5 h-5 text-white"/>}/>
                    <StatusCard title="學術單位" value={units.filter(u=>u.category==='Academic').length} gradient="from-sky-500 to-blue-600" icon={<Users className="w-5 h-5 text-white"/>}/>
                </div>
                <div className="bg-white rounded-xl shadow overflow-hidden h-[600px] relative border border-slate-200">
                    <div className="absolute top-4 right-4 z-10 flex bg-white rounded shadow border"><button onClick={()=>setZoom(z=>Math.max(0.2,z-0.1))} className="p-2 hover:bg-gray-100"><ZoomOut className="w-4 h-4"/></button><span className="px-2 py-2 text-xs font-mono">{Math.round(zoom*100)}%</span><button onClick={()=>setZoom(z=>Math.min(3,z+0.1))} className="p-2 hover:bg-gray-100"><ZoomIn className="w-4 h-4"/></button></div>
                    <div className="w-full h-full overflow-auto bg-slate-100 p-10">
                        <div style={{transform:`scale(${zoom})`, transformOrigin:'top left'}} className="relative inline-block">
                            <img src={settings.uploadedMapUrl} alt="Map" className="block"/>
                            <svg className="absolute inset-0 w-full h-full pointer-events-none" viewBox="0 0 100 100" preserveAspectRatio="none">{settings.areaMap.map(area=><polygon key={area.id} points={area.points?area.points.map(p=>`${p.x},${p.y}`).join(' '):''} fill="rgba(255,0,0,0.2)" stroke="red" strokeWidth="0.5"/>)}</svg>
                        </div>
                    </div>
                </div>
                </>
            )}
            <div className="bg-white p-4 rounded-xl shadow border border-slate-200">
                <div className="flex gap-4 mb-4 flex-wrap items-center">
                    <FilterSelect label="類型" value={filter.type} onChange={e=>setFilter({...filter, type: e.target.value})}><option value="">全部</option><option value="Administrative">行政</option><option value="Academic">學術</option></FilterSelect>
                    <div className="flex-grow"></div>
                    <button onClick={downloadImportTemplate} className={`${styles.btnSecondary} py-1 text-sm`}><Download className="w-4 h-4 mr-1"/> 下載範本</button>
                    <button onClick={handleDelete} disabled={selectedUnitIds.length===0} className={`${styles.btnDanger} py-1 text-sm`}><Trash2 className="w-4 h-4 mr-1"/> 刪除</button>
                    <button onClick={()=>exportToExcel(filtered, '列表', 'Sheet1', [{key:'name',label:'名稱'},{key:'category',label:'類別'}])} className={`${styles.btnInfo} py-1 text-sm`}><Download className="w-4 h-4 mr-1"/> 匯出</button>
                    <button onClick={()=>setIsNewUnit(true)} className={`${styles.btnPrimary} py-1 text-sm`}><Plus className="w-4 h-4 mr-1"/> 新增</button>
                </div>
                <UnitTable units={filtered} selectedUnitIds={selectedUnitIds} setSelectedUnitIds={setSelectedUnitIds} onViewUnit={setPreviewUnit}/>
            </div>
        </div>
    );
};

// 6.3 攻擊準則 Tab
const Tab2Guidelines = ({ appData, updatePrivateData }) => {
    const { guidelines, talkScripts } = appData.settings;
    const [editingId, setEditingId] = useState(null);
    const [newGuideline, setNewGuideline] = useState({ title: '', content: '' });
    const [newTalkScript, setNewTalkScript] = useState({ title: '', content: '' });

    const handleUpdate = (field, value) => { updatePrivateData({ [field]: value }); setEditingId(null); };
    const addItem = (field, newItem, setter) => { if(newItem.title){ handleUpdate(field, [...appData.settings[field], {...newItem, id: crypto.randomUUID()}]); setter({title:'',content:''}); }};
    const deleteItem = (field, id) => { handleUpdate(field, appData.settings[field].filter(i=>i.id!==id)); };

    const EditBlock = ({ item, field, collection }) => {
        const [t, setT] = useState(item.title); const [c, setC] = useState(item.content);
        return (
            <div className="p-3 bg-amber-50 rounded-lg space-y-3 ring-2 ring-amber-400">
                <input type="text" value={t} onChange={e=>setT(e.target.value)} className={`${styles.formInput} font-bold`}/>
                <textarea value={c} onChange={e=>setC(e.target.value)} className={styles.formTextarea} rows={4}/>
                <div className="flex justify-end gap-2"><button onClick={()=>setEditingId(null)} className={styles.btnSecondary}>取消</button><button onClick={()=>handleUpdate(field, collection.map(i=>i.id===item.id?{...i,title:t,content:c}:i))} className={styles.btnPrimary}>儲存</button></div>
            </div>
        );
    };

    return (
        <div className="space-y-8 p-6 max-w-7xl mx-auto">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {/* Guidelines */}
                <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100 flex flex-col">
                    <div className="p-5 bg-gradient-to-r from-blue-600 to-indigo-700 text-white"><h3 className="text-xl font-bold flex items-center"><Target className="w-5 h-5 mr-2"/> 核心進攻原則</h3></div>
                    <div className="p-6 space-y-4 flex-grow bg-slate-50">
                        {guidelines.map(item => (
                            <div key={item.id} className="p-5 bg-white rounded-xl shadow-sm border-l-4 border-blue-500 relative group">
                                {editingId===item.id ? <EditBlock item={item} field="guidelines" collection={guidelines}/> : (
                                    <>
                                        <p className="font-bold text-lg text-slate-800">{item.title}</p>
                                        <p className="text-slate-600 mt-2 whitespace-pre-wrap">{item.content}</p>
                                        <div className="absolute top-3 right-3 flex gap-1 opacity-0 group-hover:opacity-100 transition">
                                            <button onClick={()=>setEditingId(item.id)} className="p-1.5 text-blue-600 bg-blue-50 rounded-full"><Edit className="w-4 h-4"/></button>
                                            <button onClick={()=>deleteItem('guidelines', item.id)} className="p-1.5 text-red-600 bg-red-50 rounded-full"><Trash2 className="w-4 h-4"/></button>
                                        </div>
                                    </>
                                )}
                            </div>
                        ))}
                    </div>
                    <div className="p-4 bg-white border-t border-slate-100">
                        <div className="p-3 bg-blue-50 rounded-lg border border-blue-200">
                            <input placeholder="標題" className={`${styles.formInput} mb-2`} value={newGuideline.title} onChange={e=>setNewGuideline({...newGuideline, title: e.target.value})}/>
                            <textarea placeholder="內容" className={`${styles.formTextarea} mb-2`} rows={2} value={newGuideline.content} onChange={e=>setNewGuideline({...newGuideline, content: e.target.value})}/>
                            <button onClick={()=>addItem('guidelines', newGuideline, setNewGuideline)} className={`${styles.btnPrimary} w-full`}>新增原則</button>
                        </div>
                    </div>
                </div>
                {/* Talk Scripts */}
                <div className="bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100 flex flex-col">
                    <div className="p-5 bg-gradient-to-r from-emerald-600 to-teal-700 text-white"><h3 className="text-xl font-bold flex items-center"><CheckCircle className="w-5 h-5 mr-2"/> 標準話術庫</h3></div>
                    <div className="p-6 space-y-4 flex-grow bg-slate-50">
                        {talkScripts.map(item => (
                            <div key={item.id} className="p-5 bg-white rounded-xl shadow-sm border-l-4 border-emerald-500 relative group">
                                {editingId===item.id ? <EditBlock item={item} field="talkScripts" collection={talkScripts}/> : (
                                    <>
                                        <p className="font-bold text-lg text-emerald-800">{item.title}</p>
                                        <div className="p-3 mt-3 bg-emerald-50 border border-emerald-100 rounded-lg text-slate-700 text-sm whitespace-pre-wrap">{item.content}</div>
                                        <div className="absolute top-3 right-3 flex gap-1 opacity-0 group-hover:opacity-100 transition">
                                            <button onClick={()=>setEditingId(item.id)} className="p-1.5 text-blue-600 bg-blue-50 rounded-full"><Edit className="w-4 h-4"/></button>
                                            <button onClick={()=>deleteItem('talkScripts', item.id)} className="p-1.5 text-red-600 bg-red-50 rounded-full"><Trash2 className="w-4 h-4"/></button>
                                        </div>
                                    </>
                                )}
                            </div>
                        ))}
                    </div>
                    <div className="p-4 bg-white border-t border-slate-100">
                        <div className="p-3 bg-emerald-50 rounded-lg border border-emerald-200">
                            <input placeholder="標題" className={`${styles.formInput} mb-2`} value={newTalkScript.title} onChange={e=>setNewTalkScript({...newTalkScript, title: e.target.value})}/>
                            <textarea placeholder="內容" className={`${styles.formTextarea} mb-2`} rows={2} value={newTalkScript.content} onChange={e=>setNewTalkScript({...newTalkScript, content: e.target.value})}/>
                            <button onClick={()=>addItem('talkScripts', newTalkScript, setNewTalkScript)} className={`${styles.btnPrimary} w-full bg-emerald-600 hover:bg-emerald-700`}>新增話術</button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};

// 6.4 設定 Tab
const Tab5Settings = ({ appData, updatePrivateData }) => {
    const [bName, setBName] = useState(''); const [bCode, setBCode] = useState('');
    return (
        <div className="p-6 max-w-7xl mx-auto">
            <div className="bg-white p-6 rounded-xl shadow border border-slate-200">
                <h3 className="font-bold text-lg mb-4">參數設定</h3>
                <div className="flex gap-2">
                    <input placeholder="棟別名稱" className={styles.formInput} value={bName} onChange={e=>setBName(e.target.value)}/>
                    <input placeholder="代號" className={styles.formInput} value={bCode} onChange={e=>setBCode(e.target.value)}/>
                    <button onClick={()=>{if(bName&&bCode) updatePrivateData({buildings:[...appData.settings.buildings,{name:bName,code:bCode}]}); setBName(''); setBCode('')}} className={styles.btnPrimary}>新增</button>
                </div>
            </div>
        </div>
    );
};

// --- 7. App Main Component ---
const App = () => {
  const [currentTab, setCurrentTab] = useState('targets');
  const [db, setDb] = useState(null);
  const [auth, setAuth] = useState(null);
  const [userId, setUserId] = useState(null);
  const [isLoading, setIsLoading] = useState(true);
  const [globalMessage, setGlobalMessage] = useState({ text: '', type: '' });
  const [appData, setAppData] = useState({ units: [], settings: initialSettings, schedules: [], meetings: [], files: [] });
  
  // Editing States
  const [editingUnitId, setEditingUnitId] = useState(null);
  const [isNewUnit, setIsNewUnit] = useState(false);
  const [newUnitData, setNewUnitData] = useState({});
  const [recordFilter, setRecordFilter] = useState({ type: '', area: '' });
  const [showAddUnitModal, setShowAddUnitModal] = useState(false);
  const [previewUnit, setPreviewUnit] = useState(null); 

  const exportToExcel = useExcelExport();

  // Initialization
  useEffect(() => {
      if (Object.keys(firebaseConfig).length === 0) { setIsLoading(false); return; }
      const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApp();
      const database = getFirestore(app);
      const authentication = getAuth(app);
      setDb(database); setAuth(authentication);

      const initAuth = async () => {
          try {
              if (initialAuthToken) await signInWithCustomToken(authentication, initialAuthToken);
              else await signInAnonymously(authentication);
          } catch (e) {
              console.error(e); setGlobalMessage({ text: '驗證失敗', type: 'error' }); setIsLoading(false);
          }
      };
      initAuth();
      return onAuthStateChanged(authentication, (user) => {
          setUserId(user ? user.uid : null);
          setIsLoading(false);
      });
  }, []);

  // Data Listening
  useEffect(() => {
      if (!db || !userId) return;
      const unsubUnits = onSnapshot(collection(db, 'artifacts', appId, 'public', 'data', 'units'), (snap) => {
          setAppData(p => ({ ...p, units: snap.docs.map(d => ({ id: d.id, ...d.data(), equipment: safeParse(d.data().equipment||'[]'), history: safeParse(d.data().history||'[]'), characteristics: d.data().characteristics||[] })) }));
      });
      const unsubSettings = onSnapshot(doc(db, 'artifacts', appId, 'users', userId, 'settings', 'params'), (snap) => {
          if (snap.exists()) {
              const data = snap.data();
              setAppData(p => ({ ...p, settings: { ...initialSettings, ...data, customScheduleColumns: data.customScheduleColumns || [] }, schedules: data.schedules || [], meetings: data.meetings || [] }));
          } else {
              setDoc(doc(db, 'artifacts', appId, 'users', userId, 'settings', 'params'), { ...initialSettings, schedules: [], meetings: [] });
          }
      });
      const unsubFiles = onSnapshot(collection(db, 'artifacts', appId, 'users', userId, 'files'), (snap) => {
          setAppData(p => ({ ...p, files: snap.docs.map(d => ({ id: d.id, ...d.data() })) }));
      });
      return () => { unsubUnits(); unsubSettings(); unsubFiles(); };
  }, [db, userId]);

  const updatePrivateData = async (fields) => {
      if (!db || !userId) return;
      try { await updateDoc(doc(db, 'artifacts', appId, 'users', userId, 'settings', 'params'), fields); setGlobalMessage({ text: '更新成功', type: 'success' }); }
      catch (e) { console.error(e); setGlobalMessage({ text: '更新失敗', type: 'error' }); }
  };

  const handleUnitUpdate = async (id, data, isNew = false) => {
      if (!db || !userId) return;
      const payload = { ...data, equipment: safeStringify(data.equipment||[]), history: safeStringify(data.history||[]), characteristics: data.characteristics||[] };
      try {
          if (isNew) {
              await addDoc(collection(db, 'artifacts', appId, 'public', 'data', 'units'), { ...payload, createdAt: new Date().toISOString() });
          } else {
              await updateDoc(doc(db, 'artifacts', appId, 'public', 'data', 'units', id), payload);
          }
          setEditingUnitId(null); setIsNewUnit(false);
          setGlobalMessage({ text: '儲存成功', type: 'success' });
      } catch (e) { console.error(e); alert('儲存失敗'); }
  };

  const deleteUnitsFunc = async (ids) => {
      if (!db || !userId) return;
      await Promise.all(ids.map(id => deleteDoc(doc(db, 'artifacts', appId, 'public', 'data', 'units', id))));
  };

  // Render Logic
  const renderContent = () => {
      if (editingUnitId || isNewUnit) {
          return <UnitRecordView newUnitData={newUnitData} setNewUnitData={setNewUnitData} handleSaveUnit={() => handleUnitUpdate(editingUnitId, newUnitData, isNewUnit)} handleAddHistory={(log) => setNewUnitData(p => ({ ...p, history: [...p.history, { ...log, date: new Date().toISOString().substring(0, 10), id: crypto.randomUUID() }] }))} isNewUnit={isNewUnit} appData={appData} setEditingUnitId={setEditingUnitId} setIsNewUnit={setIsNewUnit} />;
      }
      switch (currentTab) {
          case 'targets': return <Tab3TargetsMap appData={appData} setEditingUnitId={setEditingUnitId} setIsNewUnit={setIsNewUnit} setPreviewUnit={setPreviewUnit} setShowAddUnitModal={setShowAddUnitModal} exportToExcel={exportToExcel} db={db} deleteUnits={deleteUnitsFunc} filter={recordFilter} setFilter={setRecordFilter} />;
          case 'calendar': return <Tab1Calendar appData={appData} updatePrivateData={updatePrivateData} db={db} userId={userId} setGlobalMessage={setGlobalMessage} exportToExcel={exportToExcel} />;
          case 'record': return <Tab3TargetsMap appData={appData} setEditingUnitId={setEditingUnitId} setIsNewUnit={setIsNewUnit} setPreviewUnit={setPreviewUnit} setShowAddUnitModal={setShowAddUnitModal} exportToExcel={exportToExcel} db={db} deleteUnits={deleteUnitsFunc} filter={recordFilter} setFilter={setRecordFilter} viewMode="record" />; 
          case 'guidelines': return <Tab2Guidelines appData={appData} updatePrivateData={updatePrivateData} />;
          case 'settings': return <Tab5Settings appData={appData} updatePrivateData={updatePrivateData} />;
          default: return null;
      }
  };

  const navItems = [
      { id: 'targets', label: '戰情地圖', icon: <MapPin className="w-4 h-4"/> },
      { id: 'calendar', label: '行事曆', icon: <Activity className="w-4 h-4"/> },
      { id: 'record', label: '拜訪紀錄', icon: <Edit className="w-4 h-4"/> },
      { id: 'guidelines', label: '攻擊準則', icon: <Target className="w-4 h-4"/> },
      { id: 'settings', label: '設定', icon: <Building className="w-4 h-4"/> },
  ];

  if (isLoading) return <div className="flex h-screen items-center justify-center bg-slate-50 text-slate-500"><Loader className="w-10 h-10 animate-spin text-indigo-600 mb-2"/></div>;

  return (
      <div className="min-h-screen bg-slate-50 font-sans text-slate-900">
          <header className="sticky top-0 z-50 bg-white/80 backdrop-blur border-b border-slate-200 shadow-sm">
              <div className="max-w-7xl mx-auto px-4">
                  <div className="flex justify-between items-center h-16">
                      <div className="flex items-center space-x-2"><Activity className="w-6 h-6 text-indigo-600"/><h1 className="text-xl font-bold text-slate-800">2026 台大攻略戰情室</h1></div>
                      <div className="text-xs text-slate-400 font-mono">{userId ? userId.substring(0,8)+'...' : 'Guest'}</div>
                  </div>
                  <nav className="flex space-x-1 overflow-x-auto no-scrollbar">{navItems.map(i => (
                      <button key={i.id} onClick={() => setCurrentTab(i.id)} className={`px-4 py-2 text-sm font-medium border-b-2 transition ${currentTab === i.id ? 'border-indigo-600 text-indigo-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}>{i.label}</button>
                  ))}</nav>
              </div>
          </header>
          {globalMessage.text && <div className={`fixed top-20 right-6 p-4 rounded-xl shadow-xl z-50 text-white ${globalMessage.type === 'success' ? 'bg-emerald-600' : 'bg-red-600'}`}>{globalMessage.text}<button onClick={() => setGlobalMessage({ text: '', type: '' })} className="ml-2"><X className="w-4 h-4"/></button></div>}
          <main className="py-6">{renderContent()}</main>
          {showAddUnitModal && <AddUnitModal onClose={() => setShowAddUnitModal(false)} onSave={async (d) => { await addDoc(collection(db, 'artifacts', appId, 'public', 'data', 'units'), { ...d, createdAt: new Date().toISOString(), equipment: '[]', history: '[]', characteristics: [] }); setShowAddUnitModal(false); }} appData={appData} />}
          {previewUnit && <UnitPreviewModal unit={previewUnit} onClose={() => setPreviewUnit(null)} />}
      </div>
  );
};

export default App;