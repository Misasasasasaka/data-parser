
import React, { useState, useCallback, useMemo, useRef, useEffect } from 'react';
import type { RecordData } from './types';

declare const XLSX: any; // Declare XLSX from CDN

// --- Constants ---
const TABLE_HEADERS = ['编号', '哈弗币(M)', '段位', '等级', '保险', '体力', '负重', 'AWM', '6头', '6甲', '特殊皮肤', 'KD', '租金', '押金', '合计', '租期（天）', '上号方式', '比例', '', '备注'];

const KEY_MAP: Record<string, string> = {
  '编号': '编号',
  '哈弗币': '哈弗币(M)',
  '段位': '段位',
  '等级': '等级',
  '保险格数': '保险',
  '体力': '体力',
  '负重': '负重',
  'AWM': 'AWM',
  '6头': '6头',
  '6甲': '6甲',
  '皮肤': '特殊皮肤',
  '特殊皮肤': '特殊皮肤',
  '绝密KD': 'KD',
  '押金': '押金',
  '上号方式': '上号方式',
};

// --- SVG Icons (defined outside components) ---
const DownloadIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
  </svg>
);

const TrashIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
  </svg>
);

const PlusIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
    </svg>
);

const UploadIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
    </svg>
);


// --- UI Components ---

const Header: React.FC = () => (
  <header className="text-center mb-10">
    <h1 className="text-4xl font-bold text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-teal-400">
      Data Parser & Exporter
    </h1>
    <p className="text-slate-400 mt-2">Paste, upload, or manually add structured data, then export to Excel instantly.</p>
  </header>
);

interface InputSectionProps {
  inputText: string;
  setInputText: (text: string) => void;
  onParse: () => void;
  error: string | null;
}

const InputSection: React.FC<InputSectionProps> = ({ inputText, setInputText, onParse, error }) => (
  <div className="bg-slate-800 p-6 rounded-lg shadow-lg mb-6">
    <label htmlFor="data-input" className="block text-sm font-medium text-slate-300 mb-2">
      Paste Data Here
    </label>
    <textarea
      id="data-input"
      value={inputText}
      onChange={(e) => setInputText(e.target.value)}
      placeholder={`Example format:
编号：...
哈弗币：...
段位：...
等级：...
保险格数：...
体力：...
负重：...
AWM：...
6头：...
6甲：...
皮肤：...
绝密KD：...
押金：...
上号方式：...

(You can paste multiple records. Each record must start with '编号'.)`}
      className="w-full h-60 p-3 bg-slate-900 border border-slate-700 rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors duration-200 text-slate-200 placeholder-slate-500 font-mono text-sm"
      aria-label="Data input text area"
    />
    {error && <p className="mt-2 text-sm text-red-400" role="alert">{error}</p>}
    <button
      onClick={onParse}
      className="mt-4 w-full flex items-center justify-center bg-blue-600 hover:bg-blue-700 disabled:bg-blue-800 disabled:cursor-not-allowed text-white font-bold py-2 px-4 rounded-md transition-all duration-200 transform hover:scale-105"
      disabled={!inputText.trim()}
    >
      <PlusIcon />
      Parse & Add to Table
    </button>
  </div>
);

interface ActionsBarProps {
  recordCount: number;
  onExport: () => void;
  onClear: () => void;
  onImport: () => void;
  onAdd: () => void;
}

const ActionsBar: React.FC<ActionsBarProps> = ({ recordCount, onExport, onClear, onImport, onAdd }) => (
    <div className="flex flex-col sm:flex-row justify-between items-center mb-4 gap-4">
        <p className="text-slate-400 text-sm">
            {recordCount > 0 ? `${recordCount} record(s) loaded.` : 'No data loaded.'}
        </p>
        <div className="flex flex-wrap gap-3 justify-center">
             <button
                onClick={onImport}
                className="flex items-center justify-center bg-green-600 hover:bg-green-700 text-white font-semibold py-2 px-4 rounded-md transition-colors duration-200"
            >
                <UploadIcon />
                Import Excel
            </button>
            <button
                onClick={onAdd}
                className="flex items-center justify-center bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-2 px-4 rounded-md transition-colors duration-200"
            >
                <PlusIcon />
                Add Record
            </button>
            <button
                onClick={onExport}
                disabled={recordCount === 0}
                className="flex items-center justify-center bg-teal-600 hover:bg-teal-700 disabled:bg-teal-800 disabled:text-slate-400 disabled:cursor-not-allowed text-white font-semibold py-2 px-4 rounded-md transition-colors duration-200"
            >
                <DownloadIcon />
                Export to Excel
            </button>
            <button
                onClick={onClear}
                disabled={recordCount === 0}
                className="flex items-center justify-center bg-red-600 hover:bg-red-700 disabled:bg-red-800 disabled:text-slate-400 disabled:cursor-not-allowed text-white font-semibold py-2 px-4 rounded-md transition-colors duration-200"
            >
                <TrashIcon />
                Clear All
            </button>
        </div>
    </div>
);


interface DataTableProps {
    records: RecordData[];
    headers: string[];
    onDeleteRecord: (index: number) => void;
}

const getHeaderStyles = (header: string): string => {
    // font-family: 'SimHei' is for '黑体'. Text size 16px is text-base in tailwind.
    const baseStyle = "font-['SimHei'] text-base font-bold";

    if (header === '') {
        return `${baseStyle} bg-slate-600`;
    }

    const blueBgWhiteText = ['编号', '哈弗币(M)', '段位', '等级', '6头', '6甲', '特殊皮肤', 'KD', '租金', '押金', '租期（天）', '上号方式', '备注'];
    const grayBgRedText = ['保险', '体力', '负重'];
    const grayBgBlackText = ['AWM', '合计'];
    const blueBgGreenText = ['比例'];

    if (blueBgWhiteText.includes(header)) {
        return `${baseStyle} bg-blue-600 text-white`;
    }
    if (grayBgRedText.includes(header)) {
        return `${baseStyle} bg-slate-400 text-red-600`;
    }
    if (grayBgBlackText.includes(header)) {
        return `${baseStyle} bg-slate-400 text-black`;
    }
    if (blueBgGreenText.includes(header)) {
        return `${baseStyle} bg-blue-600 text-green-400`;
    }
    // This function is only for data headers, Actions header is styled separately.
    return '';
};

const DataTable: React.FC<DataTableProps> = ({ records, headers, onDeleteRecord }) => {
    if (records.length === 0) {
        return (
            <div className="text-center py-10 bg-slate-800 rounded-lg shadow-inner">
                <p className="text-slate-400">Data will appear here once parsed, imported or added.</p>
            </div>
        );
    }

    return (
        <div className="overflow-x-auto bg-slate-800 rounded-lg shadow-lg">
            <table className="min-w-full text-sm text-left text-slate-300">
                <thead>
                    <tr>
                        {headers.map(header => (
                            <th key={header} scope="col" className={`px-6 py-3 whitespace-nowrap text-center ${getHeaderStyles(header)}`}>
                                {header}
                            </th>
                        ))}
                        <th scope="col" className="px-6 py-3 text-center text-xs text-slate-300 uppercase bg-slate-700">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    {records.map((record, index) => (
                        <tr key={index} className="border-b border-slate-700 hover:bg-slate-700/30 transition-colors duration-150">
                            {headers.map(header => (
                                <td key={`${index}-${header}`} className="px-6 py-4 whitespace-nowrap">
                                    {record[header] || '-'}
                                </td>
                            ))}
                             <td className="px-6 py-4 text-center">
                                <button
                                    onClick={() => onDeleteRecord(index)}
                                    className="text-slate-400 hover:text-red-500 transition-colors"
                                    aria-label={`Delete record ${index + 1}`}
                                    title="Delete record"
                                >
                                    <TrashIcon />
                                </button>
                            </td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};

interface AddRecordModalProps {
    isOpen: boolean;
    onClose: () => void;
    onSubmit: (newRecord: RecordData) => void;
    headers: string[];
}

const AddRecordModal: React.FC<AddRecordModalProps> = ({ isOpen, onClose, onSubmit, headers }) => {
    const [formData, setFormData] = useState<RecordData>({});
    const [dynamicFields, setDynamicFields] = useState<{ id: number; key: string; value: string }[]>([{ id: 1, key: '', value: '' }]);

    useEffect(() => {
        if (!isOpen) return;

        if (headers.length > 0) {
            const initialData = headers.reduce((acc, header) => ({ ...acc, [header]: '' }), {});
            setFormData(initialData);
        } else {
            setDynamicFields([{ id: Date.now(), key: '', value: '' }]);
        }

        const handleEsc = (event: KeyboardEvent) => {
            if (event.key === 'Escape') onClose();
        };
        window.addEventListener('keydown', handleEsc);
        return () => window.removeEventListener('keydown', handleEsc);
    }, [isOpen, headers, onClose]);

    const handleFormChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const { name, value } = e.target;
        setFormData(prev => ({ ...prev, [name]: value }));
    };
    
    const handleDynamicFieldChange = (id: number, field: 'key' | 'value', value: string) => {
        setDynamicFields(fields => fields.map(f => f.id === id ? { ...f, [field]: value } : f));
    };

    const addDynamicField = () => {
        setDynamicFields(fields => [...fields, { id: Date.now(), key: '', value: '' }]);
    };
    
    const removeDynamicField = (id: number) => {
        setDynamicFields(fields => fields.filter(f => f.id !== id));
    };

    const handleSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        if (headers.length > 0) {
            onSubmit(formData);
        } else {
            const newRecord = dynamicFields.reduce((acc, field) => {
                if (field.key.trim()) {
                    acc[field.key.trim()] = field.value.trim();
                }
                return acc;
            }, {} as RecordData);
            if (Object.keys(newRecord).length > 0) {
                onSubmit(newRecord);
            }
        }
    };

    if (!isOpen) return null;

    return (
        <div className="fixed inset-0 bg-black bg-opacity-70 z-50 flex justify-center items-center p-4" onClick={onClose} role="dialog" aria-modal="true">
            <div className="bg-slate-800 rounded-lg shadow-xl w-full max-w-lg p-6" onClick={e => e.stopPropagation()}>
                <h2 className="text-xl font-bold mb-4 text-white">Add New Record</h2>
                <form onSubmit={handleSubmit}>
                    <div className="space-y-4 max-h-[60vh] overflow-y-auto pr-2">
                        {headers.length > 0 ? (
                            headers.map(header => (
                                <div key={header}>
                                    <label htmlFor={header} className="block text-sm font-medium text-slate-300 mb-1">{header || 'Separator'}</label>
                                    <input
                                        type="text"
                                        id={header}
                                        name={header}
                                        onChange={handleFormChange}
                                        className="w-full p-2 bg-slate-900 border border-slate-700 rounded-md focus:ring-2 focus:ring-indigo-500"
                                        disabled={header === ''}
                                    />
                                </div>
                            ))
                        ) : (
                            dynamicFields.map((field, index) => (
                                <div key={field.id} className="flex items-center gap-2">
                                    <input
                                        type="text"
                                        placeholder="Field Name"
                                        value={field.key}
                                        onChange={(e) => handleDynamicFieldChange(field.id, 'key', e.target.value)}
                                        className="flex-1 p-2 bg-slate-900 border border-slate-700 rounded-md"
                                    />
                                    <input
                                        type="text"
                                        placeholder="Value"
                                        value={field.value}
                                        onChange={(e) => handleDynamicFieldChange(field.id, 'value', e.target.value)}
                                        className="flex-1 p-2 bg-slate-900 border border-slate-700 rounded-md"
                                    />
                                     <button type="button" onClick={() => removeDynamicField(field.id)} className="text-slate-400 hover:text-red-500 p-1" aria-label="Remove field">
                                        <TrashIcon />
                                    </button>
                                </div>
                            ))
                        )}
                        {headers.length === 0 && (
                            <button type="button" onClick={addDynamicField} className="text-indigo-400 hover:text-indigo-300 text-sm font-semibold">
                                + Add another field
                            </button>
                        )}
                    </div>
                    <div className="mt-6 flex justify-end gap-4">
                        <button type="button" onClick={onClose} className="py-2 px-4 bg-slate-600 hover:bg-slate-700 rounded-md text-white font-semibold">Cancel</button>
                        <button type="submit" className="py-2 px-4 bg-indigo-600 hover:bg-indigo-700 rounded-md text-white font-semibold">Save Record</button>
                    </div>
                </form>
            </div>
        </div>
    );
};


// --- Main App Component ---

function App() {
  const [records, setRecords] = useState<RecordData[]>([]);
  const [inputText, setInputText] = useState<string>('');
  const [error, setError] = useState<string | null>(null);
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const headers = useMemo(() => {
    return TABLE_HEADERS;
  }, []);


  const parseAndAddData = useCallback(() => {
    setError(null);
    if (!inputText.trim()) {
      setError("Input text cannot be empty.");
      return;
    }

    try {
      const chunks = inputText.trim().split(/^(?=编号[:：])/m).filter(chunk => chunk.trim() !== '');
      if (chunks.length === 0) {
          throw new Error("No valid records found. Each record must start with '编号：'.");
      }

      const newRecords: RecordData[] = chunks.map(chunk => {
        const record: RecordData = TABLE_HEADERS.reduce((acc, header) => ({ ...acc, [header]: '' }), {} as RecordData);
        
        const lines = chunk.trim().split('\n');
        lines.forEach(line => {
          const parts = line.split(/[:：]/, 2);
          if (parts.length === 2) {
            const rawKey = parts[0].trim();
            const value = parts[1].trim();
            const mappedKey = KEY_MAP[rawKey];
            if(mappedKey && mappedKey in record) {
                record[mappedKey] = value;
            }
          }
        });
        return record;
      });
      
      setRecords(prevRecords => [...prevRecords, ...newRecords]);
      setInputText('');
    } catch (e: any) {
      setError(`Parsing failed: ${e.message}`);
    }
  }, [inputText]);

  const exportToExcel = useCallback(() => {
    if (records.length === 0) return;
    try {
        // Manually construct the data for the worksheet in an array-of-arrays format.
        // This gives us direct control over the data and ensures styles are applied correctly.
        const dataForSheet = [
            headers,
            ...records.map(record => headers.map(header => record[header] || ''))
        ];
        
        // Create the worksheet from our array-of-arrays data.
        const worksheet = XLSX.utils.aoa_to_sheet(dataForSheet);

        // --- Define Cell Styles ---
        const baseFont = { name: 'SimHei', sz: 16, bold: true };
        const baseAlignment = { vertical: "center", horizontal: "center", wrapText: true };

        const styles = {
            whiteOnBlue: {
                font: { ...baseFont, color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "4F81BD" }, patternType: "solid" },
                alignment: baseAlignment
            },
            redOnGray: {
                font: { ...baseFont, color: { rgb: "FF0000" } },
                fill: { fgColor: { rgb: "D9D9D9" }, patternType: "solid" },
                alignment: baseAlignment
            },
            blackOnGray: {
                font: { ...baseFont, color: { rgb: "000000" } },
                fill: { fgColor: { rgb: "D9D9D9" }, patternType: "solid" },
                alignment: baseAlignment
            },
            greenOnBlue: {
                font: { ...baseFont, color: { rgb: "00B050" } },
                fill: { fgColor: { rgb: "4F81BD" }, patternType: "solid" },
                alignment: baseAlignment
            },
            blank: {
                fill: { fgColor: { rgb: "FFFFFF" }, patternType: "solid" },
                alignment: baseAlignment
            }
        };

        // Map the styles to each header column by its position (A=0, B=1, etc.)
        const headerStyles = [
            styles.whiteOnBlue, // A
            styles.whiteOnBlue, // B
            styles.whiteOnBlue, // C
            styles.whiteOnBlue, // D
            styles.redOnGray,   // E
            styles.redOnGray,   // F
            styles.redOnGray,   // G
            styles.blackOnGray, // H
            styles.whiteOnBlue, // I
            styles.whiteOnBlue, // J
            styles.whiteOnBlue, // K
            styles.whiteOnBlue, // L
            styles.whiteOnBlue, // M
            styles.whiteOnBlue, // N
            styles.blackOnGray, // O
            styles.whiteOnBlue, // P
            styles.whiteOnBlue, // Q
            styles.greenOnBlue, // R
            styles.blank,       // S
            styles.whiteOnBlue  // T
        ];

        // --- Apply Styles and Formatting ---

        // Loop through the headers to apply styles to the first row of the worksheet.
        if (worksheet['!ref']) { // Check if sheet is not empty
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            for (let C = range.s.c; C <= range.e.c; ++C) {
                if (C >= headerStyles.length) continue;
                const address = XLSX.utils.encode_cell({ c: C, r: 0 }); // r: 0 is the first row
                if (worksheet[address]) {
                    worksheet[address].s = headerStyles[C];
                }
            }
        }
        
        // Set column widths for better readability.
        const colWidths = headers.map(header => ({ wch: header.length > 0 ? Math.max(15, header.length * 1.2) : 5 }));
        worksheet['!cols'] = colWidths;
        
        // Set the height of the header row.
        worksheet['!rows'] = worksheet['!rows'] || [];
        worksheet['!rows'][0] = { hpx: 30 };

        // --- Generate and Download Excel File ---
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Parsed Data");
        XLSX.writeFile(workbook, "parsed_data_export.xlsx");
    } catch (e) {
        setError("Failed to export data to Excel.");
        console.error(e);
    }
  }, [records, headers]);


  const handleFileImport = (event: React.ChangeEvent<HTMLInputElement>) => {
    setError(null);
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = e.target?.result;
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet) as RecordData[];
            setRecords(prev => [...prev, ...json]);
        } catch (err) {
            setError('Failed to parse the Excel file. Please ensure it is a valid format.');
            console.error(err);
        }
    };
    reader.onerror = () => {
        setError('Failed to read the file.');
    };
    reader.readAsArrayBuffer(file);

    if(event.target) event.target.value = '';
  };
  
  const triggerFileImport = () => fileInputRef.current?.click();

  const handleDeleteRecord = (indexToDelete: number) => {
    setRecords(prev => prev.filter((_, index) => index !== indexToDelete));
  };
  
  const handleAddRecord = (newRecord: RecordData) => {
      setRecords(prev => [...prev, newRecord]);
      setIsAddModalOpen(false);
  };

  const clearAllData = useCallback(() => {
    setRecords([]);
  }, []);

  return (
    <div className="min-h-screen font-sans">
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-12">
        <Header />
        <InputSection
            inputText={inputText}
            setInputText={setInputText}
            onParse={parseAndAddData}
            error={error}
        />
        <ActionsBar 
            recordCount={records.length}
            onExport={exportToExcel}
            onClear={clearAllData}
            onImport={triggerFileImport}
            onAdd={() => setIsAddModalOpen(true)}
        />
        <DataTable records={records} headers={headers} onDeleteRecord={handleDeleteRecord} />
        <AddRecordModal 
            isOpen={isAddModalOpen}
            onClose={() => setIsAddModalOpen(false)}
            onSubmit={handleAddRecord}
            headers={headers}
        />
        <input 
            type="file" 
            ref={fileInputRef} 
            onChange={handleFileImport}
            accept=".xlsx, .xls"
            className="hidden"
            aria-hidden="true"
        />
      </main>
    </div>
  );
}

export default App;
