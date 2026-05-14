
import React, { useState } from 'react';
import { 
  FileUp, 
  Trash2, 
  CheckCircle2, 
  Download, 
  ArrowRight, 
  Users, 
  Calendar, 
  ShieldCheck, 
  Loader2,
  Table as TableIcon,
  Filter,
  Trash,
  Search,
  Clock,
  ListFilter
} from 'lucide-react';
import { parseExcel, findMajorityMonth, cleanRosterData, updateNames, exportToExcel, exportCombinedSummary } from './services/excelService';
import { RosterSheet } from './types';

const App: React.FC = () => {
  const [rosterFiles, setRosterFiles] = useState<File[]>([]);
  const [staffFile, setStaffFile] = useState<File | null>(null);
  const [processing, setProcessing] = useState(false);
  const [step, setStep] = useState<'upload' | 'verify' | 'complete'>('upload');
  
  const [cleanedRoster, setCleanedRoster] = useState<RosterSheet[]>([]);
  const [majorityMonth, setMajorityMonth] = useState<string | null>(null);
  
  const [activeTab, setActiveTab] = useState(0);

  const handleRosterUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files);
      setRosterFiles(prev => [...prev, ...newFiles]);
    }
  };

  const removeRosterFile = (index: number) => {
    setRosterFiles(prev => prev.filter((_, i) => i !== index));
  };

  const handleStaffUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setStaffFile(e.target.files[0]);
    }
  };

  const startProcessing = async () => {
    if (rosterFiles.length === 0 || !staffFile) return;

    setProcessing(true);
    try {
      const allRosterSheets: RosterSheet[] = [];
      
      for (const file of rosterFiles) {
        const sheets = await parseExcel(file);
        allRosterSheets.push(...sheets);
      }

      if (allRosterSheets.length === 0) {
        alert("No valid tabs found! Please ensure tabs are named in yyyymmdd format.");
        setProcessing(false);
        return;
      }

      const staffSheets = await parseExcel(staffFile);
      const staffData = staffSheets[0]?.data || [];

      const month = findMajorityMonth(allRosterSheets);
      if (!month) {
        alert("Could not identify dates in Column C.");
        setProcessing(false);
        return;
      }
      setMajorityMonth(month);

      // Clean: Majority Month, Truncate A-H, Truncate Row 72, Remove Empty Tabs
      const cleaned = cleanRosterData(allRosterSheets, month);
      
      if (cleaned.length === 0) {
        alert("No matching data found. All tabs were empty after filtering.");
        setProcessing(false);
        return;
      }

      // Update names in Column F using Staff List Column A
      const finalRoster = updateNames(cleaned, staffData);

      setCleanedRoster(finalRoster);
      setActiveTab(0);
      setStep('verify');
    } catch (error) {
      console.error(error);
      alert("Processing failed. Check file formats.");
    } finally {
      setProcessing(false);
    }
  };

  const handleCellChange = (sheetIndex: number, rowIndex: number, colIndex: number, value: string) => {
    const updated = [...cleanedRoster];
    updated[sheetIndex].data[rowIndex][colIndex] = value;
    setCleanedRoster(updated);
  };

  const handleExportStandard = () => {
    exportToExcel(cleanedRoster, `Processed_Roster_${majorityMonth}.xlsx`);
    setStep('complete');
  };

  const handleExportSummary = () => {
    exportCombinedSummary(cleanedRoster, `Roster_Summary_${majorityMonth}.xlsx`);
    setStep('complete');
  };

  const reset = () => {
    setRosterFiles([]);
    setStaffFile(null);
    setStep('upload');
    setCleanedRoster([]);
    setMajorityMonth(null);
    setActiveTab(0);
  };

  /**
   * Helper to format cell values for the input fields.
   * Ensures Dates don't show as long browser strings.
   */
  const formatCellValue = (cell: any): string => {
    if (cell === null || cell === undefined) return '';
    if (cell instanceof Date) {
      return cell.toLocaleDateString('en-CA'); // YYYY-MM-DD
    }
    return String(cell);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center p-4 md:p-8">
      <header className="w-full max-w-6xl mb-12 text-center">
        <div className="flex items-center justify-center gap-2 mb-2">
          <ShieldCheck className="w-10 h-10 text-indigo-600" />
          <h1 className="text-4xl font-extrabold tracking-tight text-slate-900">Roster Purifier</h1>
        </div>
        <p className="text-slate-500 text-lg">Automated cleaning, truncation, and name expansion.</p>
      </header>

      <main className="w-full max-w-6xl">
        {step === 'upload' && (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            <div className="bg-white rounded-2xl shadow-xl p-8 border border-slate-100 flex flex-col h-full">
              <div className="flex items-center gap-3 mb-6">
                <div className="p-3 bg-indigo-50 rounded-xl">
                  <TableIcon className="w-6 h-6 text-indigo-600" />
                </div>
                <div>
                  <h2 className="text-xl font-bold text-slate-800">1. Upload Rosters</h2>
                  <p className="text-sm text-slate-500">yyyymmdd tabs. Multiple files allowed.</p>
                </div>
              </div>
              
              <div className="flex-1 flex flex-col gap-4">
                <label className={`flex flex-col items-center justify-center border-2 border-dashed rounded-xl cursor-pointer transition-all p-6 ${rosterFiles.length > 0 ? 'border-indigo-400 bg-indigo-50/30' : 'border-slate-200 hover:border-indigo-400 hover:bg-slate-50'}`}>
                  <div className="flex flex-col items-center justify-center">
                    <FileUp className={`w-10 h-10 mb-2 ${rosterFiles.length > 0 ? 'text-indigo-600' : 'text-slate-400'}`} />
                    <span className="text-slate-500 text-sm font-medium">Add Roster Files</span>
                  </div>
                  <input type="file" className="hidden" accept=".xlsx, .xls" multiple onChange={handleRosterUpload} />
                </label>

                {rosterFiles.length > 0 && (
                  <div className="space-y-2 max-h-48 overflow-y-auto pr-2">
                    {rosterFiles.map((file, idx) => (
                      <div key={idx} className="flex items-center justify-between bg-slate-50 p-2 rounded-lg border border-slate-100 group">
                        <div className="flex items-center gap-2 overflow-hidden">
                          <TableIcon className="w-4 h-4 text-indigo-400 shrink-0" />
                          <span className="text-xs font-medium text-slate-700 truncate">{file.name}</span>
                        </div>
                        <button 
                          onClick={() => removeRosterFile(idx)}
                          className="p-1 text-slate-400 hover:text-red-500 transition-colors"
                        >
                          <Trash className="w-4 h-4" />
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>

            <div className="bg-white rounded-2xl shadow-xl p-8 border border-slate-100 flex flex-col h-full">
              <div className="flex items-center gap-3 mb-6">
                <div className="p-3 bg-emerald-50 rounded-xl">
                  <Users className="w-6 h-6 text-emerald-600" />
                </div>
                <div>
                  <h2 className="text-xl font-bold text-slate-800">2. Staff List</h2>
                  <p className="text-sm text-slate-500">Column A (Full Names) expands Roster Col F.</p>
                </div>
              </div>
              
              <label className={`flex-1 flex flex-col items-center justify-center border-2 border-dashed rounded-xl cursor-pointer transition-all ${staffFile ? 'border-emerald-400 bg-emerald-50/30' : 'border-slate-200 hover:border-emerald-400 hover:bg-slate-50'}`}>
                <div className="flex flex-col items-center justify-center py-10 px-4">
                  <FileUp className={`w-12 h-12 mb-4 ${staffFile ? 'text-emerald-600' : 'text-slate-400'}`} />
                  {staffFile ? <span className="text-emerald-700 font-medium text-center">{staffFile.name}</span> : <span className="text-slate-500">Upload Staff List (Excel/CSV)</span>}
                </div>
                <input type="file" className="hidden" accept=".xlsx, .xls, .csv" onChange={handleStaffUpload} />
              </label>
            </div>

            <div className="md:col-span-2 flex justify-center mt-4">
              <button
                onClick={startProcessing}
                disabled={rosterFiles.length === 0 || !staffFile || processing}
                className={`flex items-center gap-2 px-10 py-4 rounded-full font-bold text-lg transition-all ${(rosterFiles.length === 0 || !staffFile || processing) ? 'bg-slate-200 text-slate-400' : 'bg-indigo-600 text-white shadow-lg hover:scale-105'}`}
              >
                {processing ? <Loader2 className="w-6 h-6 animate-spin" /> : <>Clean & Match Names <ArrowRight className="w-6 h-6" /></>}
              </button>
            </div>
          </div>
        )}

        {step === 'verify' && (
          <div className="bg-white rounded-2xl shadow-xl border border-slate-100 overflow-hidden animate-in zoom-in-95">
            <div className="p-6 bg-slate-900 text-white flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div>
                <h2 className="text-2xl font-bold">Human Verification</h2>
                <div className="flex flex-wrap items-center gap-4 mt-1 text-slate-400 text-sm">
                  <div className="flex items-center gap-1"><Calendar className="w-4 h-4" />Majority: {majorityMonth}</div>
                  <div className="flex items-center gap-1"><Search className="w-4 h-4" />Names Expanded</div>
                  <div className="flex items-center gap-1"><Clock className="w-4 h-4" />Numeric Dates</div>
                </div>
              </div>
              <div className="flex flex-wrap gap-3">
                <button onClick={reset} className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-lg flex items-center gap-2 transition-colors">
                  <Trash2 className="w-4 h-4" /> Reset
                </button>
                <button onClick={handleExportStandard} className="px-5 py-2 bg-indigo-500 hover:bg-indigo-600 text-white rounded-lg font-bold flex items-center gap-2 transition-all hover:scale-105">
                  <Download className="w-5 h-5" /> Export Tabs
                </button>
                <button onClick={handleExportSummary} className="px-5 py-2 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg font-bold flex items-center gap-2 transition-all hover:scale-105">
                  <ListFilter className="w-5 h-5" /> Export Summary (C-G)
                </button>
              </div>
            </div>

            <div className="flex border-b border-slate-200 overflow-x-auto bg-slate-50">
              {cleanedRoster.map((sheet, idx) => (
                <button
                  key={idx}
                  onClick={() => setActiveTab(idx)}
                  className={`px-6 py-3 text-sm font-semibold whitespace-nowrap border-b-2 ${activeTab === idx ? 'border-indigo-600 text-indigo-600 bg-white' : 'border-transparent text-slate-500'}`}
                >
                  {sheet.name}
                </button>
              ))}
            </div>

            <div className="overflow-auto max-h-[60vh]">
              <table className="w-full border-collapse">
                <thead className="sticky top-0 bg-slate-100 z-10 shadow-sm">
                  <tr>
                    {['A', 'B', 'Date (Numeric)', 'RAW in', 'RAW out', 'Full Name (F)', 'Change (G)', 'H'].map((col) => (
                      <th key={col} className={`px-4 py-3 text-xs font-bold uppercase tracking-wider border-b ${col.includes('F') || col.includes('Date') || col.includes('RAW') || col.includes('Change') ? 'text-indigo-600 bg-indigo-50' : 'text-slate-400'}`}>{col}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {cleanedRoster[activeTab]?.data.map((row, rIdx) => (
                    <tr key={rIdx} className="hover:bg-slate-50">
                      {row.map((cell, cIdx) => (
                        <td key={cIdx} className={`p-0 border-r last:border-0 ${cIdx >= 2 && cIdx <= 6 ? 'bg-indigo-50/10' : ''}`}>
                          <input
                            type="text"
                            value={formatCellValue(cell)}
                            onChange={(e) => handleCellChange(activeTab, rIdx, cIdx, e.target.value)}
                            className={`w-full px-4 py-2 text-sm bg-transparent border-none focus:ring-2 focus:ring-indigo-400 focus:bg-white ${cIdx === 5 ? 'font-medium text-slate-900' : 'text-slate-600'}`}
                          />
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {step === 'complete' && (
          <div className="flex flex-col items-center justify-center py-20 text-center">
            <div className="w-24 h-24 bg-emerald-100 rounded-full flex items-center justify-center mb-6">
              <CheckCircle2 className="w-12 h-12 text-emerald-600" />
            </div>
            <h2 className="text-3xl font-bold mb-2">Success!</h2>
            <p className="text-slate-500 mb-8">Roster cleaned, names aligned, and dates saved as numeric values.</p>
            <button onClick={reset} className="px-8 py-3 bg-indigo-600 text-white rounded-full font-bold shadow-lg">Process Another</button>
          </div>
        )}
      </main>
    </div>
  );
};

export default App;
