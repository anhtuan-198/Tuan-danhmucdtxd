import React, { useState, useEffect, useRef } from 'react';
import Papa from 'papaparse';
import { Settings, RefreshCw, FilePlus, Database, Download, ExternalLink, AlertCircle, X } from 'lucide-react';
import { EVN_HCMC_LOGO } from "./assets/logo";

const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1KgPTbaGUntJXZTjUs3v_iKBmG7yEltxSK7sjxOkaNK8/export?format=csv&gid=0';
const PROJECTS_CSV_URL = 'https://docs.google.com/spreadsheets/d/1KgPTbaGUntJXZTjUs3v_iKBmG7yEltxSK7sjxOkaNK8/export?format=csv&gid=1152018861'; // Using gid for 'Thông tin theo MCT' if known, or gviz. Let's use gviz to be safe.
const PROJECTS_GVIZ_URL = 'https://docs.google.com/spreadsheets/d/1KgPTbaGUntJXZTjUs3v_iKBmG7yEltxSK7sjxOkaNK8/gviz/tq?tqx=out:csv&sheet=' + encodeURIComponent('Thông tin theo MCT');

const PROXY_URL = `https://corsproxy.io/?${encodeURIComponent(SHEET_CSV_URL)}`;
const PROJECTS_PROXY_URL = `https://corsproxy.io/?${encodeURIComponent(PROJECTS_GVIZ_URL)}`;

const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwneLdbK4axFnelPpWLGeCclV9_GLwrCIm2SglXlDh_Cmq16U3b-t7_ryfpo14mLU2Z/exec';

export default function App() {
  const [data, setData] = useState<string[][]>([]);
  const [projectInfo, setProjectInfo] = useState<Record<string, string>>({});
  const [availableProjects, setAvailableProjects] = useState<Record<string, string>[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [scriptUrl, setScriptUrl] = useState(localStorage.getItem('appsScriptUrl') || DEFAULT_SCRIPT_URL);
  const [showSettings, setShowSettings] = useState(false);
  const [showImportModal, setShowImportModal] = useState(false);
  const [actionLoading, setActionLoading] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');

  const [showDropdown, setShowDropdown] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  const clickCountRef = useRef(0);
  const lastClickTimeRef = useRef(0);

  const handleLogoClick = () => {
    const now = Date.now();
    if (now - lastClickTimeRef.current < 500) {
      clickCountRef.current += 1;
    } else {
      clickCountRef.current = 1;
    }
    lastClickTimeRef.current = now;

    if (clickCountRef.current >= 5) {
      setShowSettings(prev => !prev);
      clickCountRef.current = 0;
    }
  };

  useEffect(() => {
    fetchData();
    
    // Handle click outside for custom dropdown
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setShowDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const fetchData = async (expectedProjectCode?: string, retries = 0) => {
    setLoading(true);
    setError(null);
    try {
      // Add timestamp to prevent caching
      const timestamp = new Date().getTime();
      const urlWithCacheBuster = `${SHEET_CSV_URL}&t=${timestamp}`;
      
      // 1. Fetch CSV data first
      let response;
      let fetchError = null;
      
      try {
        response = await fetch(urlWithCacheBuster);
        if (!response.ok) throw new Error(`Direct fetch failed: ${response.status}`);
      } catch (e) {
        fetchError = e;
        try {
          response = await fetch(`${PROXY_URL}&t=${timestamp}`);
          if (!response.ok) throw new Error(`Proxy fetch failed: ${response.status}`);
        } catch (e2) {
          fetchError = e2;
          try {
            response = await fetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(urlWithCacheBuster)}`);
            if (!response.ok) throw new Error(`AllOrigins fetch failed: ${response.status}`);
          } catch (e3) {
            fetchError = e3;
            // One last try with another proxy
            try {
              response = await fetch(`https://thingproxy.freeboard.io/fetch/${encodeURIComponent(urlWithCacheBuster)}`);
              if (!response.ok) throw new Error(`ThingProxy fetch failed: ${response.status}`);
            } catch (e4) {
              fetchError = e4;
              throw new Error('Không thể kết nối với dữ liệu Google Sheet. Vui lòng kiểm tra kết nối mạng hoặc thử lại sau.');
            }
          }
        }
      }
      
      if (!response || !response.ok) {
        throw new Error('Không thể tải dữ liệu từ Google Sheet.');
      }
      
      const csvText = await response.text();
      
      // 2. Parse CSV immediately to check for project code match
      let csvResults: Papa.ParseResult<string[]> | null = null;
      Papa.parse(csvText, {
        complete: (results) => {
            csvResults = results as Papa.ParseResult<string[]>;
        }
      });
      
      if (!csvResults || !csvResults.data) throw new Error('Failed to parse CSV');
      
      const allRows = csvResults.data as string[][];
      
      // Extract Project Info from rows 7-13 (approx)
      const info: Record<string, string> = {};
      for (let i = 7; i <= 13; i++) {
        if (allRows[i] && allRows[i][0]) {
          info[allRows[i][0].replace(':', '').trim()] = allRows[i][2] || '';
        }
      }

      // Check if the fetched data matches the expected project
      const fetchedProjectCode = info["Mã công trình"];
      
      if (expectedProjectCode && fetchedProjectCode !== expectedProjectCode) {
        if (retries > 0) {
            console.log(`Mismatch: Expected ${expectedProjectCode}, got ${fetchedProjectCode}. Retrying... (${retries} left)`);
            setTimeout(() => fetchData(expectedProjectCode, retries - 1), 1000);
            return;
        }

        console.warn(`Mismatch: Expected ${expectedProjectCode}, got ${fetchedProjectCode}`);
        // Data mismatch - likely Apps Script failed to update or data is empty
        // We show empty state instead of stale data
        setData([]);
        setLoading(false);
        // We don't overwrite projectInfo here because we want to keep what the user selected
        return;
      }
      
      // 3. Match found! Now fetch HTML version to extract links
      let htmlText = '';
      try {
        const htmlUrl = 'https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/htmlview/sheet?headers=true&gid=0';
        let htmlResponse;
        try {
          htmlResponse = await fetch(htmlUrl);
          if (!htmlResponse.ok) throw new Error('Direct fetch failed');
        } catch (e) {
          try {
            htmlResponse = await fetch(`https://corsproxy.io/?${encodeURIComponent(htmlUrl)}`);
            if (!htmlResponse.ok) throw new Error('Proxy fetch failed');
          } catch (e2) {
            htmlResponse = await fetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(htmlUrl)}`);
          }
        }
        if (htmlResponse && htmlResponse.ok) {
          htmlText = await htmlResponse.text();
        }
      } catch (e) {
        console.warn('Could not fetch HTML version for links');
      }
      
      // Parse HTML to extract links if available
      const linkMap = new Map<string, string>();
      if (htmlText) {
        // Parse the table structure
        try {
          const parser = new DOMParser();
          const doc = parser.parseFromString(htmlText, 'text/html');
          const rows = doc.querySelectorAll('table tbody tr');
          
          rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            // Usually column 1 is Tên văn bản
            if (cells.length > 1) {
              const textCell = cells[1];
              const textContent = textCell.textContent?.replace(/\s+/g, ' ').trim();
              
              // Find the link anywhere in the row
              const linkElement = row.querySelector('a');
              
              if (textContent && linkElement && linkElement.href) {
                let href = linkElement.href;
                if (href.includes('google.com/url?')) {
                  try {
                    const urlParams = new URLSearchParams(href.split('?')[1]);
                    const q = urlParams.get('q');
                    if (q) href = q;
                  } catch (e) {
                    // Ignore parsing errors
                  }
                }
                linkMap.set(textContent, href);
              }
            }
          });
          console.log(`Extracted ${linkMap.size} links from HTML`);
        } catch (e) {
          console.warn('Error parsing HTML table', e);
        }
      }
      
      // 4. Apply extracted links to the data
      if (linkMap.size > 0) {
        for (let i = 0; i < allRows.length; i++) {
          if (allRows[i] && allRows[i][5]?.trim().toLowerCase() === 'xem file') {
            const textToMatch = allRows[i][1]?.replace(/\s+/g, ' ').trim();
            if (textToMatch && linkMap.has(textToMatch)) {
              allRows[i][5] = linkMap.get(textToMatch)!;
            }
          }
        }
      }
      
      // Reorder the keys as requested
      const orderedInfo: Record<string, string> = {};
      const desiredOrder = [
        "Tên dự án/công trình",
        "Mã công trình",
        "Chủ Đầu Tư",
        "Địa điểm xây dựng",
        "Đơn vị TV Thiết Kế",
        "Đơn vị TV Giám sát",
        "Đơn vị thi công"
      ];
      
      desiredOrder.forEach(key => {
        if (info[key] !== undefined) {
          orderedInfo[key] = info[key];
        }
      });
      
      // Add any remaining keys
      Object.keys(info).forEach(key => {
        if (orderedInfo[key] === undefined) {
          orderedInfo[key] = info[key];
        }
      });

      setProjectInfo(orderedInfo);
      
      // Fetch available projects
      fetchProjects(orderedInfo);

      // Extract table data (from row 15 onwards)
      // Row 15 is header, 16+ is data
      const tableData = allRows.slice(15).filter(row => row.some(cell => cell.trim() !== ''));
      setData(tableData);
      
      setLoading(false);

    } catch (err: any) {
      setError(err.message);
      setLoading(false);
    }
  };

  const fetchProjects = async (currentProjectInfo: Record<string, string>) => {
    try {
      let response;
      const timestamp = new Date().getTime();
      const urlWithCacheBuster = `${PROJECTS_GVIZ_URL}&t=${timestamp}`;
      
      try {
        response = await fetch(urlWithCacheBuster);
        if (!response.ok) throw new Error('Direct fetch failed');
      } catch (e) {
        try {
          response = await fetch(`${PROJECTS_PROXY_URL}&t=${timestamp}`);
          if (!response.ok) throw new Error('Proxy fetch failed');
        } catch (e2) {
          try {
            response = await fetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(urlWithCacheBuster)}`);
            if (!response.ok) throw new Error('AllOrigins fetch failed');
          } catch (e3) {
            // One last try
            response = await fetch(`https://thingproxy.freeboard.io/fetch/${encodeURIComponent(urlWithCacheBuster)}`);
            if (!response.ok) throw new Error('ThingProxy fetch failed');
          }
        }
      }
      
      const csvText = await response.text();
      Papa.parse(csvText, {
        header: false,
        skipEmptyLines: true,
        complete: (results) => {
          const rows = results.data as string[][];
          const projects: Record<string, string>[] = [];
          
          // The sheet "Thông tin theo MCT" structure:
          // Col 0: STT
          // Col 1: Mã công trình
          // Col 2: Tên công trình
          // Col 3: Địa điểm xây dựng
          // Col 4: Chủ đầu tư
          // Col 5: Đơn vị TV Thiết Kế
          // Col 6: Đơn vị TV Giám sát
          // Col 7: Đơn vị thi công
          
          // Start from row 1 (skipping header)
          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (row && row.length >= 3 && row[2] && row[2].trim() !== "" && row[2] !== "Tên công trình") {
              projects.push({
                "Tên dự án/công trình": row[2].trim(),
                "Mã công trình": row[1]?.trim() || "",
                "Chủ Đầu Tư": row[4]?.trim() || "",
                "Địa điểm xây dựng": row[3]?.trim() || "",
                "Đơn vị TV Thiết Kế": row[5]?.trim() || "",
                "Đơn vị TV Giám sát": row[6]?.trim() || "",
                "Đơn vị thi công": row[7]?.trim() || ""
              });
            }
          }
          
          // Ensure current project is in the list if not already
          const currentName = currentProjectInfo["Tên dự án/công trình"];
          if (currentName && !projects.some(p => p["Tên dự án/công trình"] === currentName)) {
            projects.unshift(currentProjectInfo);
          }
          
          // Remove duplicates by name
          const uniqueProjects = projects.filter((v, i, a) => 
            a.findIndex(t => t["Tên dự án/công trình"] === v["Tên dự án/công trình"]) === i
          );
          
          setAvailableProjects(uniqueProjects);
        }
      });
    } catch (err) {
      console.error("Failed to fetch projects list:", err);
      setAvailableProjects([currentProjectInfo]);
    }
  };

  const saveScriptUrl = (url: string) => {
    setScriptUrl(url);
    localStorage.setItem('appsScriptUrl', url);
  };

  const triggerAction = async (actionName: string, actionId: string, params: Record<string, string> = {}) => {
    if (!scriptUrl) {
      alert('Vui lòng cấu hình Apps Script Web App URL trong phần Cài đặt trước khi thực hiện chức năng này.');
      setShowSettings(true);
      return;
    }

    setActionLoading(actionId);
    try {
      const url = new URL(scriptUrl);
      url.searchParams.append('action', actionId);
      
      // Append any additional parameters
      Object.entries(params).forEach(([key, value]) => {
        url.searchParams.append(key, value);
      });
      
      // We use no-cors because Apps Script Web Apps often have CORS issues unless configured properly.
      // With no-cors, the browser won't let us read the response, but the request DOES reach the server.
      await fetch(url.toString(), { mode: 'no-cors' });
      
      // Wait for Google Sheets to run the script and update its published CSV cache
      // This might take a while depending on how heavy the Apps Script is
      await new Promise(resolve => setTimeout(resolve, 500));
      
      // If we are updating project info, we expect the sheet to reflect this project
      if (actionId === 'updateProjectInfo' && params["Mã công trình"]) {
        // Use retries (10 attempts * 1s = 10s max) to poll for the update
        fetchData(params["Mã công trình"], 10);
      } else {
        fetchData();
      }
      
    } catch (err: any) {
      console.error(err);
      alert(`Lỗi khi thực hiện "${actionName}": Vui lòng kiểm tra lại kết nối hoặc URL Apps Script.`);
    } finally {
      setActionLoading(null);
    }
  };

  // Clean up data for display
  const displayData = data.filter(row => row.some(cell => cell.trim() !== ''));
  const headers = displayData.length > 0 ? displayData[0] : [];
  const allRows = displayData.length > 1 ? displayData.slice(1) : [];
  const rows = allRows.filter(row => row.some(cell => cell.toLowerCase().includes(searchTerm.toLowerCase())));

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex items-center bg-white rounded p-1 cursor-pointer select-none" onClick={handleLogoClick} title="EVN HCMC">
              <img 
                src={EVN_HCMC_LOGO} 
                alt="EVN HCMC" 
                className="h-14 w-auto object-contain"
                referrerPolicy="no-referrer"
              />
            </div>
            <h1 className="text-xl sm:text-2xl font-bold text-blue-900 uppercase tracking-tight">QUẢN LÝ DANH MỤC ĐTXD</h1>
          </div>
          <div className="flex items-center gap-4">
            <button 
              onClick={fetchData}
              className="p-2 text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-full transition-colors"
              title="Làm mới dữ liệu"
            >
              <RefreshCw className={`w-5 h-5 ${loading ? 'animate-spin' : ''}`} />
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Import Modal */}
        {showImportModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
            <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl overflow-hidden flex flex-col h-[90vh]">
              <div className="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                <div className="flex items-center gap-2">
                  <Database className="w-5 h-5 text-blue-600" />
                  <h3 className="font-bold text-slate-800 uppercase">Nhập dữ liệu</h3>
                </div>
                <button 
                  onClick={() => {
                    setShowImportModal(false);
                    fetchData();
                  }}
                  className="p-2 hover:bg-slate-200 rounded-full text-slate-500 hover:text-slate-700 transition-all"
                  title="Đóng cửa sổ"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>
              <div className="flex-1 overflow-hidden relative bg-slate-100">
                <iframe 
                  src={scriptUrl} 
                  className="w-full h-full border-none"
                  title="Apps Script Form"
                />
              </div>
              <div className="px-6 py-4 border-t border-slate-200 flex justify-end bg-slate-50">
                <button 
                  onClick={() => {
                    setShowImportModal(false);
                    fetchData();
                  }}
                  className="px-8 py-2.5 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-bold transition-colors shadow-md"
                >
                  HOÀN TẤT
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Settings Panel */}
        {showSettings && (
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 mb-8 animate-in slide-in-from-top-4 fade-in duration-200">
            <h2 className="text-lg font-medium mb-4 flex items-center gap-2">
              <Settings className="w-5 h-5 text-slate-500" />
              Cấu hình kết nối Google Apps Script
            </h2>
            <div className="bg-amber-50 border border-amber-200 rounded-lg p-4 mb-4 text-sm text-amber-800 flex gap-3">
              <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
              <div>
                <p className="font-medium mb-1">Tại sao cần cấu hình URL này?</p>
                <p className="mb-2">Để các nút chức năng hoạt động, bạn cần triển khai mã Apps Script trên Google Sheet của bạn dưới dạng <strong>Web App</strong> và dán URL vào đây.</p>
                <ol className="list-decimal ml-5 space-y-1">
                  <li>Mở Google Sheet, vào <strong>Tiện ích mở rộng &gt; Apps Script</strong>.</li>
                  <li>Thêm hàm <code>doGet(e)</code> để xử lý các tham số <code>action</code> ('create', 'import', 'export').</li>
                  <li>Chọn <strong>Triển khai &gt; Triển khai mới</strong>, chọn loại <strong>Ứng dụng web</strong>.</li>
                  <li>Quyền truy cập: <strong>Bất kỳ ai</strong>.</li>
                  <li>Sao chép URL Web App và dán vào ô bên dưới.</li>
                </ol>
              </div>
            </div>
            <div className="flex gap-3">
              <input 
                type="url" 
                value={scriptUrl}
                onChange={(e) => saveScriptUrl(e.target.value)}
                placeholder="https://script.google.com/macros/s/.../exec"
                className="flex-1 px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none transition-shadow"
              />
              <button 
                onClick={() => setShowSettings(false)}
                className="px-4 py-2 bg-slate-800 text-white rounded-lg hover:bg-slate-700 font-medium transition-colors"
              >
                Lưu & Đóng
              </button>
            </div>
          </div>
        )}

        {/* Action Buttons */}
        <div className="flex flex-wrap items-center justify-between gap-4 mb-8 bg-white p-6 rounded-xl shadow-sm border border-slate-200">
          <div className="flex items-center gap-4 w-full sm:w-auto">
            <button 
              onClick={() => {
                if (!scriptUrl) {
                  alert('Vui lòng cấu hình Apps Script Web App URL trong phần Cài đặt trước khi thực hiện chức năng này.');
                  setShowSettings(true);
                  return;
                }
                setShowImportModal(true);
              }}
              className="flex items-center justify-center gap-2 px-6 py-3 bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-lg hover:from-blue-700 hover:to-blue-800 font-semibold shadow-md hover:shadow-lg transform hover:-translate-y-0.5 transition-all duration-200 w-full sm:w-auto"
            >
              <FilePlus className="w-5 h-5" />
              <span className="uppercase tracking-wide text-sm">Nhập dữ liệu</span>
            </button>
            
            <button 
              onClick={async () => {
                const filename = `Danh mục hồ sơ - ${projectInfo["Mã công trình"] || ""}.xlsx`;
                const url = "https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/export?format=xlsx&gid=0";
                
                try {
                  // Try multiple proxies for download
                  const proxies = [
                    `https://corsproxy.io/?${encodeURIComponent(url)}`,
                    `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
                    `https://thingproxy.freeboard.io/fetch/${encodeURIComponent(url)}`
                  ];
                  
                  let blob = null;
                  for (const proxy of proxies) {
                    try {
                      const response = await fetch(proxy);
                      if (response.ok) {
                        blob = await response.blob();
                        break;
                      }
                    } catch (e) {
                      console.warn(`Proxy ${proxy} failed`, e);
                    }
                  }
                  
                  if (blob) {
                    const blobUrl = window.URL.createObjectURL(blob);
                    const link = document.createElement('a');
                    link.href = blobUrl;
                    link.download = filename;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    window.URL.revokeObjectURL(blobUrl);
                  } else {
                    throw new Error('All proxies failed');
                  }
                } catch (error) {
                  console.error('Download failed', error);
                  // Fallback to direct link if all proxies fail
                  const link = document.createElement('a');
                  link.href = url;
                  link.target = '_blank';
                  link.click();
                }
              }}
              className="flex items-center justify-center gap-2 px-6 py-3 bg-white text-emerald-700 border-2 border-emerald-600 rounded-lg hover:bg-emerald-50 font-semibold shadow-sm hover:shadow-md transition-all duration-200 w-full sm:w-auto"
            >
              <Download className="w-5 h-5" />
              <span className="uppercase tracking-wide text-sm">Tải xuống danh mục</span>
            </button>
          </div>
        </div>

        {/* Project Info */}
        {Object.keys(projectInfo).length > 0 && (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 mb-8">
            <h2 className="text-lg font-semibold text-slate-800 mb-4 border-b border-slate-100 pb-2">Thông tin dự án</h2>
            <div className="flex flex-col gap-y-3">
              {Object.entries(projectInfo).map(([key, value], idx) => {
                if (key === 'Tên dự án/công trình') {
                  return (
                    <div key={idx} className="flex flex-col sm:flex-row sm:items-center gap-1 sm:gap-3">
                      <span className="text-slate-500 font-medium min-w-[160px]">{key}:</span>
                      <div className="relative flex-1 max-w-4xl" ref={dropdownRef}>
                        <input 
                          type="text" 
                          value={value}
                          onFocus={() => setShowDropdown(true)}
                          onChange={(e) => {
                            const selectedName = e.target.value;
                            setProjectInfo(prev => ({...prev, "Tên dự án/công trình": selectedName}));
                            setShowDropdown(true);
                            
                            const foundProject = availableProjects.find(p => p["Tên dự án/công trình"] === selectedName);
                            if (foundProject) {
                              setProjectInfo(foundProject);
                              setData([]);
                              setLoading(true);
                              setShowDropdown(false);
                              if (scriptUrl) {
                                triggerAction('Cập nhật Thông tin dự án', 'updateProjectInfo', foundProject);
                              } else {
                                alert('Vui lòng cài đặt Apps Script Web App URL (biểu tượng bánh răng góc trên bên phải) để có thể đồng bộ dữ liệu lên Google Sheet.');
                                setLoading(false);
                              }
                            }
                          }}
                          className="w-full pl-3 pr-10 py-1.5 border border-slate-300 rounded-md focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none text-slate-800 font-semibold"
                          placeholder="Nhập từ khoá để tìm kiếm..."
                        />
                        {value && (
                          <button 
                            onClick={() => {
                              setProjectInfo(prev => ({...prev, "Tên dự án/công trình": ""}));
                              setShowDropdown(true);
                            }}
                            className="absolute right-3 top-1/2 -translate-y-1/2 p-1 text-slate-400 hover:text-slate-600"
                            title="Xoá"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        )}
                        
                        {showDropdown && (
                          <div className="absolute z-50 left-0 right-0 mt-1 bg-white border border-slate-200 rounded-md shadow-lg max-h-80 overflow-y-auto">
                            {availableProjects
                              .filter(p => !value || p["Tên dự án/công trình"].toLowerCase().includes(value.toLowerCase()))
                              .map((p, i) => (
                                <div 
                                  key={i}
                                  className="px-4 py-2 hover:bg-emerald-50 cursor-pointer text-sm text-slate-700 border-b border-slate-50 last:border-0 leading-relaxed"
                                  onClick={() => {
                                    setProjectInfo(p);
                                    setData([]);
                                    setLoading(true);
                                    setShowDropdown(false);
                                    if (scriptUrl) {
                                      triggerAction('Cập nhật Thông tin dự án', 'updateProjectInfo', p);
                                    } else {
                                      alert('Vui lòng cài đặt Apps Script Web App URL (biểu tượng bánh răng góc trên bên phải) để có thể đồng bộ dữ liệu lên Google Sheet.');
                                      setLoading(false);
                                    }
                                  }}
                                >
                                  {p["Tên dự án/công trình"]}
                                </div>
                              ))}
                            {availableProjects.filter(p => !value || p["Tên dự án/công trình"].toLowerCase().includes(value.toLowerCase())).length === 0 && (
                              <div className="px-4 py-3 text-sm text-slate-400 italic">Không tìm thấy dự án phù hợp</div>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  );
                }
                return (
                  <div key={idx} className="flex flex-col sm:flex-row sm:items-start gap-1 sm:gap-3">
                    <span className="text-slate-500 font-medium min-w-[160px]">{key}:</span>
                    <span className="text-slate-800 font-semibold">{value}</span>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Data Table */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
          <div className="px-6 py-4 border-b border-slate-200 bg-slate-50 flex flex-col sm:flex-row justify-between items-center gap-4">
            <h2 className="font-semibold text-slate-800">DANH MỤC HỒ SƠ</h2>
            <div className="flex items-center gap-4 w-full sm:w-auto">
              <input
                type="text"
                placeholder="Tìm kiếm..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="px-3 py-1.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none w-full sm:w-64"
              />
              <span className="text-xs font-medium text-slate-500 bg-slate-200 px-2.5 py-1 rounded-full whitespace-nowrap">
                {rows.length} dòng
              </span>
            </div>
          </div>
          
          {loading ? (
            <div className="p-12 flex flex-col items-center justify-center text-slate-500">
              <RefreshCw className="w-8 h-8 animate-spin mb-4 text-emerald-600" />
              <p>Đang tải dữ liệu ...</p>
            </div>
          ) : error ? (
            <div className="p-8 text-center text-red-600">
              <AlertCircle className="w-12 h-12 mx-auto mb-3 opacity-50" />
              <p className="font-medium">Lỗi khi tải dữ liệu</p>
              <p className="text-sm mt-1">{error}</p>
            </div>
          ) : (
            <div className="overflow-x-auto max-h-[600px]">
              <table className="w-full text-sm text-left whitespace-nowrap">
                <thead className="text-xs text-slate-600 uppercase bg-slate-100 sticky top-0 z-10 shadow-sm text-center">
                  <tr>
                    {/* Custom headers based on the CSV structure we saw */}
                    <th className="px-4 py-3 font-semibold border-b border-slate-200">STT</th>
                    <th className="px-4 py-3 font-semibold border-b border-slate-200">TÊN VĂN BẢN</th>
                    <th className="px-4 py-3 font-semibold border-b border-slate-200">SỐ VB</th>
                    <th className="px-4 py-3 font-semibold border-b border-slate-200">NGÀY VB</th>
                    <th className="px-4 py-3 font-semibold border-b border-slate-200">FILE VB</th>
                    <th className="px-4 py-3 font-semibold border-b border-slate-200">CƠ QUAN BAN HÀNH</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {rows.map((row, rowIndex) => {
                    // Skip empty rows or rows that are just section headers if they don't fit well
                    // But we'll try to render them nicely
                    const isSectionHeader = row[1] && !row[3] && !row[4] && !row[5] && !row[6];
                    
                    return (
                      <tr 
                        key={rowIndex} 
                        className={`hover:bg-slate-50 transition-colors ${isSectionHeader ? 'bg-slate-100 font-semibold text-emerald-800' : ''}`}
                      >
                        <td className="px-4 py-3 border-r border-slate-100 text-center text-slate-500">{row[0]}</td>
                        <td className="px-4 py-3 border-r border-slate-100 whitespace-normal min-w-[300px]">{row[1]}</td>
                        <td className="px-4 py-3 border-r border-slate-100 text-center">{row[3]}</td>
                        <td className="px-4 py-3 border-r border-slate-100">{row[4]}</td>
                        <td className="px-4 py-3 border-r border-slate-100">
                          {row[5] && (row[5].startsWith('http://') || row[5].startsWith('https://')) ? (
                            <a 
                              href={row[5]} 
                              target="_blank" 
                              rel="noopener noreferrer"
                              className="inline-flex items-center gap-1 text-blue-600 hover:text-blue-800 hover:underline font-medium"
                            >
                              <ExternalLink className="w-3.5 h-3.5" />
                              Xem file
                            </a>
                          ) : (
                            <span className={row[5]?.trim().toLowerCase() === 'xem file' ? 'text-slate-400 italic' : ''}>
                              {row[5]}
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3 border-r border-slate-100 text-center">{row[6]}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </main>

      {/* Action Loading Overlay */}
      {actionLoading && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center">
          <div className="bg-white rounded-xl shadow-xl p-8 max-w-sm w-full mx-4 flex flex-col items-center text-center">
            <div className="w-12 h-12 border-4 border-emerald-100 border-t-emerald-600 rounded-full animate-spin mb-4"></div>
            <h3 className="text-lg font-semibold text-slate-800 mb-2">Đang xử lý dữ liệu...</h3>
            <p className="text-slate-500 text-sm">
              Hệ thống đang cập nhật...
            </p>
          </div>
        </div>
      )}
    </div>
  );
}
