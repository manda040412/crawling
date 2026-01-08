import React, { useState } from 'react';
import { Upload, Download, Search, CheckCircle, XCircle, Loader } from 'lucide-react';
import * as XLSX from 'xlsx';

const JikiuCrawler = () => {
  const [file, setFile] = useState(null);
  const [data, setData] = useState([]);
  const [results, setResults] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setFile(file);
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    setData(jsonData);
    setResults([]);
  };

  const searchJikiu = async (itemCode) => {
    try {
      // Search for the part
      const searchUrl = `https://www.jikiu.com/catalogue/search?part=${encodeURIComponent(itemCode)}`;
      
      // Note: Due to CORS restrictions, we can't directly fetch from Jikiu
      // This is a demonstration of the structure
      const response = await fetch(searchUrl, {
        method: 'GET',
        headers: {
          'Accept': 'text/html',
        }
      });

      if (!response.ok) {
        return null;
      }

      const html = await response.text();
      
      // Parse HTML to extract data
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      
      // Extract specification data
      const specifications = {};
      const specElements = doc.querySelectorAll('.specification-item');
      specElements.forEach(el => {
        const label = el.querySelector('.label')?.textContent.trim();
        const value = el.querySelector('.value')?.textContent.trim();
        if (label && value) {
          specifications[label] = value;
        }
      });

      // Extract crosses data
      const crosses = [];
      const crossElements = doc.querySelectorAll('.crosses-table tr');
      crossElements.forEach(el => {
        const owner = el.querySelector('td:nth-child(1)')?.textContent.trim();
        const number = el.querySelector('td:nth-child(2)')?.textContent.trim();
        if (owner && number) {
          crosses.push({ owner, number });
        }
      });

      return {
        found: true,
        specifications,
        crosses,
        url: searchUrl
      };
    } catch (error) {
      console.error(`Error searching for ${itemCode}:`, error);
      return null;
    }
  };

  const processData = async () => {
    setIsProcessing(true);
    setProgress({ current: 0, total: data.length });
    const processedResults = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      setProgress({ current: i + 1, total: data.length });

      // Simulate API call delay to avoid overwhelming the server
      await new Promise(resolve => setTimeout(resolve, 1000));

      const itemCode = row.ItemCode || row['Item Code'] || '';
      
      // Mock result for demonstration (since we can't actually crawl due to CORS)
      const result = {
        brand: row.Brand || '',
        itemCode: itemCode,
        carMakerName: row['Car Maker Name'] || row.CarMakerName || '',
        carModelName: row['Car Model Name'] || row.CarModelName || '',
        carChassisName: row['Car Chassis Name'] || row.CarChassisName || '',
        carEngineDescName: row['Car EngineDesc Name'] || row.CarEngineDescName || '',
        carVehicleName: row['Car Vehicle Name'] || row.CarVehicleName || '',
        yearFrom: row['Year From'] || row.YearFrom || '',
        yearTo: row['Year To'] || row.YearTo || '',
        oemNo: row['OEM No.'] || row.OEMNo || '',
        partDescription: row['Part Description'] || row.PartDescription || '',
        aliasName: row['Alias Name'] || row.AliasName || '',
        printDescription: row['Print Description'] || row.PrintDescription || '',
        
        // Jikiu data (would be populated from actual crawling)
        foundInJikiu: Math.random() > 0.5, // Mock: randomly found or not
        jikiuPartNumber: itemCode,
        jikiuUrl: `https://www.jikiu.com/catalogue/search?part=${encodeURIComponent(itemCode)}`,
        
        // Specifications
        conePitch: '',
        coneSizeMm: '',
        threadSize: '',
        overallHeightMm: '',
        diameterMm: '',
        mountingHeightMm: '',
        location: '',
        position: '',
        
        // Crosses
        crosses: []
      };

      processedResults.push(result);
    }

    setResults(processedResults);
    setIsProcessing(false);
  };

  const exportToExcel = () => {
    // Prepare data for export
    const exportData = results.map(r => ({
      'Brand': r.brand,
      'Item Code': r.itemCode,
      'Car Maker Name': r.carMakerName,
      'Car Model Name': r.carModelName,
      'Car Chassis Name': r.carChassisName,
      'Car Engine Desc Name': r.carEngineDescName,
      'Car Vehicle Name': r.carVehicleName,
      'Year From': r.yearFrom,
      'Year To': r.yearTo,
      'OEM No.': r.oemNo,
      'Part Description': r.partDescription,
      'Alias Name': r.aliasName,
      'Print Description': r.printDescription,
      'Found in Jikiu': r.foundInJikiu ? 'YES' : 'NO',
      'Jikiu URL': r.jikiuUrl,
      'Cone Pitch': r.conePitch,
      'Cone Size (mm)': r.coneSizeMm,
      'Thread Size': r.threadSize,
      'Overall Height (mm)': r.overallHeightMm,
      'Diameter (mm)': r.diameterMm,
      'Mounting Height (mm)': r.mountingHeightMm,
      'Location': r.location,
      'Position': r.position,
      'Crosses': r.crosses.map(c => `${c.owner}: ${c.number}`).join('; ')
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Jikiu Crawl Results');
    
    // Auto-size columns
    const maxWidth = 50;
    const colWidths = Object.keys(exportData[0] || {}).map(key => {
      const maxLen = Math.max(
        key.length,
        ...exportData.map(row => String(row[key] || '').length)
      );
      return { wch: Math.min(maxLen + 2, maxWidth) };
    });
    ws['!cols'] = colWidths;

    XLSX.writeFile(wb, 'Jikiu_Crawl_Results.xlsx');
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-8">
      <div className="max-w-7xl mx-auto">
        <div className="bg-white rounded-2xl shadow-2xl p-8">
          <div className="flex items-center justify-between mb-8">
            <h1 className="text-4xl font-bold text-gray-800">
              Jikiu Parts Crawler
            </h1>
            <Search className="w-10 h-10 text-indigo-600" />
          </div>

          {/* Upload Section */}
          <div className="mb-8">
            <label className="flex flex-col items-center justify-center w-full h-48 border-2 border-dashed border-indigo-300 rounded-xl cursor-pointer hover:bg-indigo-50 transition-all">
              <div className="flex flex-col items-center justify-center pt-5 pb-6">
                <Upload className="w-16 h-16 mb-4 text-indigo-500" />
                <p className="mb-2 text-lg font-semibold text-gray-700">
                  {file ? file.name : 'Click to upload Excel file'}
                </p>
                <p className="text-sm text-gray-500">
                  Upload List spare parts-Anugerah Auto.xlsx
                </p>
              </div>
              <input
                type="file"
                className="hidden"
                accept=".xlsx,.xls"
                onChange={handleFileUpload}
              />
            </label>
          </div>

          {/* Data Summary */}
          {data.length > 0 && (
            <div className="mb-8 p-6 bg-indigo-50 rounded-xl">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-lg font-semibold text-gray-800">
                    Loaded: {data.length} items
                  </p>
                  <p className="text-sm text-gray-600 mt-1">
                    Ready to crawl Jikiu.com
                  </p>
                </div>
                <button
                  onClick={processData}
                  disabled={isProcessing}
                  className="px-8 py-3 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 disabled:bg-gray-400 disabled:cursor-not-allowed transition-all flex items-center gap-2"
                >
                  {isProcessing ? (
                    <>
                      <Loader className="w-5 h-5 animate-spin" />
                      Processing...
                    </>
                  ) : (
                    <>
                      <Search className="w-5 h-5" />
                      Start Crawling
                    </>
                  )}
                </button>
              </div>

              {/* Progress Bar */}
              {isProcessing && (
                <div className="mt-4">
                  <div className="flex justify-between text-sm text-gray-600 mb-2">
                    <span>Progress</span>
                    <span>{progress.current} / {progress.total}</span>
                  </div>
                  <div className="w-full bg-gray-200 rounded-full h-3">
                    <div
                      className="bg-indigo-600 h-3 rounded-full transition-all duration-300"
                      style={{ width: `${(progress.current / progress.total) * 100}%` }}
                    />
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Results */}
          {results.length > 0 && (
            <div>
              <div className="flex items-center justify-between mb-6">
                <h2 className="text-2xl font-bold text-gray-800">
                  Results ({results.length} items)
                </h2>
                <button
                  onClick={exportToExcel}
                  className="px-6 py-3 bg-green-600 text-white rounded-lg font-semibold hover:bg-green-700 transition-all flex items-center gap-2"
                >
                  <Download className="w-5 h-5" />
                  Export to Excel
                </button>
              </div>

              {/* Summary Stats */}
              <div className="grid grid-cols-3 gap-4 mb-6">
                <div className="bg-green-50 p-4 rounded-lg">
                  <div className="flex items-center gap-2">
                    <CheckCircle className="w-6 h-6 text-green-600" />
                    <div>
                      <p className="text-sm text-gray-600">Found</p>
                      <p className="text-2xl font-bold text-green-600">
                        {results.filter(r => r.foundInJikiu).length}
                      </p>
                    </div>
                  </div>
                </div>
                <div className="bg-red-50 p-4 rounded-lg">
                  <div className="flex items-center gap-2">
                    <XCircle className="w-6 h-6 text-red-600" />
                    <div>
                      <p className="text-sm text-gray-600">Not Found</p>
                      <p className="text-2xl font-bold text-red-600">
                        {results.filter(r => !r.foundInJikiu).length}
                      </p>
                    </div>
                  </div>
                </div>
                <div className="bg-blue-50 p-4 rounded-lg">
                  <div className="flex items-center gap-2">
                    <Search className="w-6 h-6 text-blue-600" />
                    <div>
                      <p className="text-sm text-gray-600">Total</p>
                      <p className="text-2xl font-bold text-blue-600">
                        {results.length}
                      </p>
                    </div>
                  </div>
                </div>
              </div>

              {/* Results Table */}
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead className="bg-gray-100 border-b-2 border-gray-300">
                    <tr>
                      <th className="px-4 py-3 text-left font-semibold">Item Code</th>
                      <th className="px-4 py-3 text-left font-semibold">Brand</th>
                      <th className="px-4 py-3 text-left font-semibold">Part Description</th>
                      <th className="px-4 py-3 text-center font-semibold">Found in Jikiu</th>
                      <th className="px-4 py-3 text-left font-semibold">Jikiu URL</th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.map((result, idx) => (
                      <tr key={idx} className="border-b hover:bg-gray-50">
                        <td className="px-4 py-3 font-mono">{result.itemCode}</td>
                        <td className="px-4 py-3">{result.brand}</td>
                        <td className="px-4 py-3">{result.partDescription}</td>
                        <td className="px-4 py-3 text-center">
                          {result.foundInJikiu ? (
                            <CheckCircle className="w-5 h-5 text-green-600 mx-auto" />
                          ) : (
                            <XCircle className="w-5 h-5 text-red-600 mx-auto" />
                          )}
                        </td>
                        <td className="px-4 py-3">
                          <a
                            href={result.jikiuUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="text-indigo-600 hover:underline text-xs"
                          >
                            View on Jikiu
                          </a>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Info Notice */}
          <div className="mt-8 p-4 bg-yellow-50 border-l-4 border-yellow-400 rounded">
            <p className="text-sm text-yellow-800">
              <strong>Note:</strong> Due to browser security restrictions (CORS), this demo simulates crawling results. 
              For actual web crawling, you would need to:
              <br />• Use a backend service (Node.js, Python) to fetch data from Jikiu
              <br />• Implement proper rate limiting to respect the website
              <br />• Handle authentication if required
              <br />• Parse HTML responses to extract the needed data
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default JikiuCrawler;