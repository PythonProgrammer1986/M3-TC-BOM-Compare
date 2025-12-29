
import React, { useState, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  CheckCircle2, 
  Download, 
  Upload, 
  Trash2,
  Table as TableIcon,
  ChevronRight,
  Info,
  Activity,
  ArrowRight,
  ShieldCheck,
  Search,
  Database,
  Lock,
  Globe,
  Settings2
} from 'lucide-react';
import { TCBOMRow, M3BOMRow, IIMRow, BOMError, ValidationSummary, TCConnection } from './types';
import { validateBOM } from './services/bomValidator';

const App: React.FC = () => {
  const [sourceMode, setSourceMode] = useState<'upload' | 'direct'>('upload');
  const [tcData, setTcData] = useState<TCBOMRow[] | null>(null);
  const [m3Data, setM3Data] = useState<M3BOMRow[] | null>(null);
  const [iimData, setIimData] = useState<IIMRow[] | null>(null);
  
  const [errors, setErrors] = useState<BOMError[]>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [isConnecting, setIsConnecting] = useState(false);

  const [tcConfig, setTcConfig] = useState<TCConnection>({
    url: '',
    user: '',
    pass: '',
    itemId: ''
  });

  const [summary, setSummary] = useState<ValidationSummary>({
    totalErrors: 0,
    task1Errors: 0,
    task2Errors: 0,
    task3Errors: 0,
    quantityErrors: 0,
    status: 'Idle',
    healthScore: 100
  });

  const filteredErrors = useMemo(() => {
    return errors.filter(e => 
      e.partId.toLowerCase().includes(searchTerm.toLowerCase()) || 
      e.parentId.toLowerCase().includes(searchTerm.toLowerCase()) ||
      e.errorType.toLowerCase().includes(searchTerm.toLowerCase())
    );
  }, [errors, searchTerm]);

  const handleFileUpload = (type: 'TC' | 'M3' | 'IIM', file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" }) as any[];

        const cleanedData = jsonData.map(row => {
          const newRow: any = {};
          Object.keys(row).forEach(key => {
            newRow[key.trim()] = String(row[key] ?? "").trim();
          });
          return newRow;
        });

        if (type === 'TC') setTcData(cleanedData);
        if (type === 'M3') setM3Data(cleanedData);
        if (type === 'IIM') setIimData(cleanedData);
      } catch (err) {
        console.error("Error parsing Excel file:", err);
        alert("Could not parse Excel file. Please ensure it's a valid .xlsx or .xls file.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const fetchFromTeamcenter = async () => {
    if (!tcConfig.url || !tcConfig.user || !tcConfig.pass || !tcConfig.itemId) {
      alert("Please fill in all Teamcenter connection fields.");
      return;
    }

    setIsConnecting(true);
    try {
      // In production, this would call the Teamcenter SOA/REST API
      await new Promise(resolve => setTimeout(resolve, 2000));
      alert(`Integration Point Active: Attempting to pull BOM for ${tcConfig.itemId} from ${tcConfig.url}. This feature requires back-end proxy configuration for production use.`);
    } catch (err) {
      console.error("TC Connection Error:", err);
      alert("Failed to connect to Teamcenter. Check URL and Credentials.");
    } finally {
      setIsConnecting(false);
    }
  };

  const runValidation = useCallback(() => {
    if (!tcData || !m3Data || !iimData) return;
    setSummary(prev => ({ ...prev, status: 'Validating' }));
    
    setTimeout(() => {
      const result = validateBOM(tcData, m3Data, iimData);
      setErrors(result.errors);
      setSummary({
        totalErrors: result.errors.length,
        task1Errors: result.errors.filter(e => e.task === 'Task 1').length,
        task2Errors: result.errors.filter(e => e.task === 'Task 2').length,
        task3Errors: result.errors.filter(e => e.task === 'Task 3').length,
        quantityErrors: result.errors.filter(e => e.task === 'Quantity Check').length,
        status: 'Success',
        healthScore: result.healthScore
      });
    }, 600);
  }, [tcData, m3Data, iimData]);

  const reset = () => {
    setTcData(null);
    setM3Data(null);
    setIimData(null);
    setErrors([]);
    setSummary({ totalErrors: 0, task1Errors: 0, task2Errors: 0, task3Errors: 0, quantityErrors: 0, status: 'Idle', healthScore: 100 });
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      <header className="sticky top-0 z-40 bg-white/80 backdrop-blur-md border-b border-slate-200 px-6 py-4 shadow-sm">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-slate-900 rounded-xl flex items-center justify-center text-white shadow-lg">
              <ShieldCheck size={24} />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight">BOM Auditor Pro</h1>
              <div className="flex items-center gap-2 text-xs text-slate-500 font-medium">
                <span className="flex items-center gap-1"><Database size={12} /> TC Direct</span>
                <span className="w-1 h-1 bg-slate-300 rounded-full"></span>
                <span className="flex items-center gap-1"><Activity size={12} /> M3 ERP</span>
              </div>
            </div>
          </div>
          
          <div className="flex items-center gap-3">
            {(tcData || m3Data || iimData) && (
              <button onClick={reset} className="p-2 text-slate-400 hover:text-red-500 transition-colors" title="Reset All Data">
                <Trash2 size={20} />
              </button>
            )}
            <button 
              onClick={() => {
                const ws = XLSX.utils.json_to_sheet(errors);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Audit_Findings");
                XLSX.writeFile(wb, "BOM_Audit_Report.xlsx");
              }}
              className="hidden md:flex items-center gap-2 px-4 py-2 bg-slate-900 text-white rounded-lg hover:bg-slate-800 font-semibold shadow-md transition-all active:scale-95 disabled:opacity-50"
              disabled={errors.length === 0}
            >
              <Download size={18} /> Export Results
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-6 space-y-8 w-full flex-1">
        
        {/* Connection Toggle */}
        <div className="flex justify-center">
          <div className="bg-white p-1 rounded-xl border border-slate-200 flex gap-1 shadow-sm">
            <button 
              onClick={() => setSourceMode('upload')}
              className={`px-4 py-2 rounded-lg text-sm font-bold transition-all ${sourceMode === 'upload' ? 'bg-slate-900 text-white shadow-md' : 'text-slate-500 hover:bg-slate-50'}`}
            >
              <div className="flex items-center gap-2"><Upload size={16} /> XLSX Upload</div>
            </button>
            <button 
              onClick={() => setSourceMode('direct')}
              className={`px-4 py-2 rounded-lg text-sm font-bold transition-all ${sourceMode === 'direct' ? 'bg-slate-900 text-white shadow-md' : 'text-slate-500 hover:bg-slate-50'}`}
            >
              <div className="flex items-center gap-2"><Globe size={16} /> Direct TC Connect</div>
            </button>
          </div>
        </div>

        {summary.status === 'Idle' && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            
            {/* Source Config Panel */}
            <div className="lg:col-span-8 space-y-6">
              {sourceMode === 'upload' ? (
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  <FileUploadSection title="Teamcenter" desc="BOM Export (.xlsx)" data={tcData} onUpload={(f) => handleFileUpload('TC', f)} icon={<Database className="text-blue-500" />} />
                  <FileUploadSection title="M3 ERP" desc="Production Structure" data={m3Data} onUpload={(f) => handleFileUpload('M3', f)} icon={<Activity className="text-indigo-500" />} />
                  <FileUploadSection title="IIM Master" desc="Item Policy Map" data={iimData} onUpload={(f) => handleFileUpload('IIM', f)} icon={<Settings2 className="text-emerald-500" />} />
                </div>
              ) : (
                <div className="bg-white p-8 rounded-3xl border border-slate-200 shadow-sm space-y-6">
                  <div className="flex items-center gap-3 mb-2">
                    <Globe className="text-blue-500" />
                    <h3 className="text-xl font-bold">Teamcenter Credentials</h3>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="space-y-1">
                      <label className="text-xs font-bold text-slate-500 uppercase">Server URL</label>
                      <div className="relative">
                        <Globe className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                        <input 
                          type="text" placeholder="https://tc-server:8080/tc" 
                          className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm"
                          value={tcConfig.url} onChange={e => setTcConfig({...tcConfig, url: e.target.value})}
                        />
                      </div>
                    </div>
                    <div className="space-y-1">
                      <label className="text-xs font-bold text-slate-500 uppercase">Top-Level Item ID</label>
                      <div className="relative">
                        <TableIcon className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                        <input 
                          type="text" placeholder="e.g. PART-10045" 
                          className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm"
                          value={tcConfig.itemId} onChange={e => setTcConfig({...tcConfig, itemId: e.target.value})}
                        />
                      </div>
                    </div>
                    <div className="space-y-1">
                      <label className="text-xs font-bold text-slate-500 uppercase">Username</label>
                      <div className="relative">
                        <Lock className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                        <input 
                          type="text" placeholder="infodba" 
                          className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm"
                          value={tcConfig.user} onChange={e => setTcConfig({...tcConfig, user: e.target.value})}
                        />
                      </div>
                    </div>
                    <div className="space-y-1">
                      <label className="text-xs font-bold text-slate-500 uppercase">Password</label>
                      <div className="relative">
                        <Lock className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                        <input 
                          type="password" placeholder="••••••••" 
                          className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm"
                          value={tcConfig.pass} onChange={e => setTcConfig({...tcConfig, pass: e.target.value})}
                        />
                      </div>
                    </div>
                  </div>
                  <button 
                    onClick={fetchFromTeamcenter}
                    disabled={isConnecting}
                    className="w-full py-4 bg-blue-600 text-white rounded-2xl font-bold flex items-center justify-center gap-3 hover:bg-blue-700 transition-all shadow-lg disabled:bg-slate-300"
                  >
                    {isConnecting ? (
                      <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                    ) : (
                      <>Pull BOM Data from TC <ArrowRight size={18} /></>
                    )}
                  </button>
                </div>
              )}
            </div>

            {/* Side Logic Card */}
            <div className="lg:col-span-4">
              <div className="bg-slate-900 text-white p-6 rounded-3xl shadow-xl h-full flex flex-col justify-between">
                <div>
                  <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
                    <Settings2 className="text-blue-400" />
                    Validation Rules
                  </h3>
                  <ul className="space-y-4 text-sm text-slate-300">
                    <li className="flex gap-3">
                      <div className="w-5 h-5 rounded bg-blue-500/20 text-blue-400 flex items-center justify-center shrink-0">1</div>
                      <p><strong>Missing L1:</strong> Items in Engineering but missing from M3.</p>
                    </li>
                    <li className="flex gap-3">
                      <div className="w-5 h-5 rounded bg-blue-500/20 text-blue-400 flex items-center justify-center shrink-0">2</div>
                      <p><strong>Qty Mismatch:</strong> Detects difference in planned counts.</p>
                    </li>
                    <li className="flex gap-3">
                      <div className="w-5 h-5 rounded bg-blue-500/20 text-blue-400 flex items-center justify-center shrink-0">3</div>
                      <p><strong>Phantom Sync:</strong> Validates Phantom structure expansion.</p>
                    </li>
                  </ul>
                </div>
                <button
                  onClick={runValidation}
                  disabled={!tcData || !m3Data || !iimData}
                  className={`mt-8 w-full py-4 rounded-2xl font-bold flex items-center justify-center gap-2 transition-all ${
                    tcData && m3Data && iimData ? 'bg-white text-slate-900 hover:scale-105' : 'bg-slate-800 text-slate-500 cursor-not-allowed'
                  }`}
                >
                  Start Structural Audit
                </button>
              </div>
            </div>
          </div>
        )}

        {summary.status === 'Success' && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            <div className="lg:col-span-4 space-y-6">
              <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="font-bold text-slate-500 uppercase text-xs">Integrity Score</h3>
                  <Activity size={18} className="text-slate-400" />
                </div>
                <div className="text-5xl font-black text-slate-900">{Math.round(summary.healthScore)}%</div>
                <div className="mt-4 w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                  <div className="h-full bg-slate-900 transition-all duration-1000" style={{ width: `${summary.healthScore}%` }}></div>
                </div>
              </div>

              <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm space-y-4">
                <h4 className="font-bold text-sm uppercase tracking-widest text-slate-400">Findings Breakdown</h4>
                <div className="space-y-2">
                  <BreakdownRow label="Missing L1" val={summary.task1Errors} color="bg-red-500" />
                  <BreakdownRow label="Qty Errors" val={summary.quantityErrors} color="bg-orange-500" />
                  <BreakdownRow label="Phantom Sync" val={summary.task2Errors} color="bg-amber-500" />
                  <BreakdownRow label="Policy" val={summary.task3Errors} color="bg-rose-500" />
                </div>
              </div>
              
              <div className="bg-blue-50 p-6 rounded-3xl border border-blue-100">
                <h4 className="font-bold text-blue-900 mb-2 flex items-center gap-2">
                  <Info size={18} /> Manual Correction Guide
                </h4>
                <p className="text-sm text-blue-700 leading-relaxed">
                  Most errors can be resolved in M3 using transaction <strong>PDS001</strong> (Product Structure) or <strong>MMS001</strong> (Item Master).
                </p>
              </div>
            </div>

            <div className="lg:col-span-8 space-y-6">
              <div className="bg-white rounded-3xl border border-slate-200 shadow-sm flex flex-col h-[700px]">
                <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-slate-50/50 rounded-t-3xl">
                  <div className="relative flex-1 max-w-sm">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
                    <input 
                      type="text" 
                      placeholder="Search results..." 
                      className="w-full pl-10 pr-4 py-2 bg-white border border-slate-200 rounded-xl text-sm shadow-sm"
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                    />
                  </div>
                  <span className="text-sm font-bold text-slate-500">{filteredErrors.length} issues identified</span>
                </div>

                <div className="flex-1 overflow-auto custom-scrollbar">
                  <table className="w-full border-collapse">
                    <thead className="sticky top-0 bg-white z-10 border-b border-slate-100">
                      <tr className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                        <th className="px-6 py-4 text-left">Type</th>
                        <th className="px-6 py-4 text-left">Hierarchy (Parent → Child)</th>
                        <th className="px-6 py-4 text-center">Priority</th>
                        <th className="px-6 py-4 text-right">M3 Transaction</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {filteredErrors.map((error) => (
                        <tr 
                          key={error.id} 
                          className="hover:bg-slate-50 transition-colors"
                        >
                          <td className="px-6 py-4">
                            <div className="text-sm font-bold text-slate-900">{error.errorType}</div>
                            <div className="text-[10px] text-slate-400 font-bold uppercase">{error.task}</div>
                          </td>
                          <td className="px-6 py-4">
                            <div className="flex items-center gap-2 text-sm">
                              <span className="font-mono text-slate-400 truncate max-w-[100px]">{error.parentId}</span>
                              <ArrowRight size={12} className="text-slate-300" />
                              <span className="font-bold text-slate-900">{error.partId}</span>
                            </div>
                            <div className="text-[11px] text-slate-500 mt-0.5">{error.description}</div>
                          </td>
                          <td className="px-6 py-4 text-center">
                            <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase ${
                              error.priority === 'High' ? 'bg-red-100 text-red-700' : 'bg-orange-100 text-orange-700'
                            }`}>
                              {error.priority}
                            </span>
                          </td>
                          <td className="px-6 py-4 text-right">
                            <div className="text-[11px] font-medium text-blue-700 bg-blue-50 px-3 py-1.5 rounded-lg border border-blue-100 inline-block">
                              {error.actionableFix}
                            </div>
                          </td>
                        </tr>
                      ))}
                      {filteredErrors.length === 0 && (
                        <tr>
                          <td colSpan={4} className="py-20 text-center text-slate-400">
                            No issues found matching your search.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
};

const FileUploadSection: React.FC<{ title: string, desc: string, data: any, onUpload: (f: File) => void, icon: React.ReactNode }> = ({ title, desc, data, onUpload, icon }) => (
  <div className={`p-6 bg-white rounded-3xl border-2 transition-all ${data ? 'border-green-500 shadow-green-50' : 'border-slate-100 hover:border-blue-200'}`}>
    <div className="flex items-start justify-between mb-4">
      <div className="w-10 h-10 rounded-xl bg-slate-50 flex items-center justify-center">{icon}</div>
      {data && <div className="p-1 bg-green-100 text-green-600 rounded-full"><CheckCircle2 size={16} /></div>}
    </div>
    <h3 className="font-bold text-slate-900">{title}</h3>
    <p className="text-xs text-slate-500 mb-6">{desc}</p>
    {data ? (
      <div className="flex items-center gap-2 text-green-600 font-bold text-sm bg-green-50 py-2 px-3 rounded-xl border border-green-100">
        <TableIcon size={14} /> {data.length.toLocaleString()} rows
      </div>
    ) : (
      <label className="flex items-center justify-center gap-2 py-3 bg-slate-900 text-white rounded-xl text-xs font-bold cursor-pointer hover:bg-slate-800 transition-all active:scale-95 shadow-md">
        <Upload size={14} /> Browse XLSX
        <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => e.target.files?.[0] && onUpload(e.target.files[0])} />
      </label>
    )}
  </div>
);

const BreakdownRow: React.FC<{ label: string, val: number, color: string }> = ({ label, val, color }) => (
  <div className="flex items-center justify-between text-sm">
    <div className="flex items-center gap-2">
      <div className={`w-2 h-2 rounded-full ${color}`} />
      <span className="text-slate-600">{label}</span>
    </div>
    <span className="font-bold">{val}</span>
  </div>
);

export default App;
