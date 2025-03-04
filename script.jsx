import React, { useState } from 'react';
import * as XLSX from 'sheetjs';
import _ from 'lodash';

const JsonToExcelConverter = () => {
  const [jsonInput, setJsonInput] = useState('');
  const [status, setStatus] = useState('');
  const [loading, setLoading] = useState(false);
  
  // Function to extract server topology from content
  const extractServerTopology = (jsonObj) => {
    if (jsonObj && jsonObj.serversTopology) {
      return jsonObj.serversTopology;
    }
    return null;
  };
  
  // Function to extract non-array values from JSON
  const extractNonArrayValues = (jsonObj) => {
    const nonArrayValues = [];
    
    if (jsonObj) {
      Object.keys(jsonObj).forEach(key => {
        const value = jsonObj[key];
        // Skip array and object values except simple null values
        if (value === null || 
            (typeof value !== 'object' && !Array.isArray(value)) ||
            (typeof value === 'object' && !Array.isArray(value) && value !== null && Object.keys(value).length === 0)) {
          nonArrayValues.push([key, value]);
        }
      });
    }
    
    return nonArrayValues;
  };
  
  // Function to clean up and parse JSON
  const parseJsonInput = (input) => {
    try {
      // Handle JSON with or without the "this is my json" prefix
      let jsonContent = input;
      
      if (input.includes('this is my json')) {
        jsonContent = input.substring(input.indexOf('{'));
      }
      
      // Handle common JSON syntax errors
      // Fix missing quotes and syntax issues that might exist
      const fixedContent = jsonContent
        .replace(/"([^"]*)"([^"]*):\s*"([^"]*)",/g, (match, p1, p2, p3) => {
          if (p2.trim() === '') {
            return `"${p1}": "${p3}",`;
          }
          return match;
        })
        .replace(/"client_id":\s*"([^"]*),/g, '"client_id": "$1",');
      
      return JSON.parse(fixedContent);
    } catch (error) {
      throw new Error(`Invalid JSON: ${error.message}`);
    }
  };
  
  const processJson = () => {
    try {
      setLoading(true);
      setStatus('Processing JSON...');
      
      if (!jsonInput.trim()) {
        throw new Error('Please enter JSON data');
      }
      
      // Parse JSON input
      const jsonObj = parseJsonInput(jsonInput);
      
      // Extract server topology
      const serverTopology = extractServerTopology(jsonObj);
      
      // Extract non-array values
      const nonArrayValues = extractNonArrayValues(jsonObj);
      
      // Create workbook
      const wb = XLSX.utils.book_new();
      
      // Sheet 1: Non-array values
      const nonArraySheet = createNonArraySheet(nonArrayValues);
      XLSX.utils.book_append_sheet(wb, nonArraySheet, 'General Info');
      
      // Sheet 2: Server topology
      const topologySheet = createServerTopologySheet(serverTopology);
      XLSX.utils.book_append_sheet(wb, topologySheet, 'Server Topology');
      
      // Generate Excel file
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      
      // Create Blob and download
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      
      // Create download link
      const link = document.createElement('a');
      link.href = url;
      link.download = 'subscription_data.xlsx';
      link.click();
      
      setStatus('Excel file created successfully!');
      setLoading(false);
    } catch (error) {
      console.error('Error processing JSON:', error);
      setStatus(`Error: ${error.message}`);
      setLoading(false);
    }
  };
  
  const createNonArraySheet = (nonArrayValues) => {
    // Create worksheet with headers
    return XLSX.utils.aoa_to_sheet(
      [['Key', 'Value'], ...nonArrayValues]
    );
  };
  
  const createServerTopologySheet = (serverTopology) => {
    if (!serverTopology) {
      // Return empty sheet if no data
      return XLSX.utils.aoa_to_sheet([['No server topology data found']]);
    }
    
    // Parse the server topology string
    // Format: "[server1-provider-status-type:value,server2-provider-status-type:value]"
    const serversString = serverTopology.substring(1, serverTopology.length - 1);
    const serverEntries = serversString.split(',');
    
    const parsedServers = serverEntries.map(entry => {
      const [serverInfo, value] = entry.split(':');
      const parts = serverInfo.split('-');
      
      // Need at least 4 parts for the required columns
      if (parts.length >= 4) {
        const serverName = parts[0];
        const provider = parts[1];
        const status = parts[2];
        const type = parts[3];
        const updated = value === '1' ? 'Yes' : 'No';
        
        return [serverName, provider, status, type, updated];
      }
      return null;
    }).filter(entry => entry !== null);
    
    // Create headers and add data
    const headers = ['Server Name', 'Provider', 'Status', 'Type', 'Updated?'];
    const data_with_headers = [headers, ...parsedServers];
    
    // Create worksheet
    return XLSX.utils.aoa_to_sheet(data_with_headers);
  };
  
  // Load sample data from file for demo purposes
  const loadSampleData = async () => {
    try {
      setLoading(true);
      setStatus('Loading sample data...');
      
      // Read the file content
      const fileContent = await window.fs.readFile('paste.txt', { encoding: 'utf8' });
      setJsonInput(fileContent);
      
      setStatus('Sample data loaded successfully!');
      setLoading(false);
    } catch (error) {
      console.error('Error loading sample data:', error);
      setStatus(`Error loading sample: ${error.message}`);
      setLoading(false);
    }
  };
  
  return (
    <div className="flex flex-col items-center justify-center p-6 bg-gray-50 rounded-lg shadow-sm">
      <h1 className="text-2xl font-bold mb-6">JSON to Excel Converter</h1>
      
      <div className="bg-blue-50 border border-blue-200 p-4 rounded-md mb-6 w-full">
        <h2 className="text-lg font-semibold mb-2">This tool will:</h2>
        <ul className="list-disc pl-6">
          <li>Convert your JSON data to an Excel file with two sheets</li>
          <li>Sheet 1: All non-array values from the JSON</li>
          <li>Sheet 2: Server topology parsed into a table with columns for server name, provider, status, type, and whether it's updated</li>
        </ul>
      </div>
      
      <div className="w-full mb-4">
        <label htmlFor="json-input" className="block mb-2 font-medium">
          Paste your JSON data:
        </label>
        <textarea
          id="json-input"
          value={jsonInput}
          onChange={(e) => setJsonInput(e.target.value)}
          className="w-full h-64 p-3 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
          placeholder='Example: {"serversTopology": "[server1-provider-status-type:0,server2-provider-status-type:1]", "client_id": "example@test.com", ...}'
        />
      </div>
      
      <div className="flex w-full mb-6 space-x-4">
        <button
          onClick={loadSampleData}
          disabled={loading}
          className="px-4 py-2 bg-gray-600 text-white rounded-md hover:bg-gray-700 focus:outline-none focus:ring-2 focus:ring-gray-500 focus:ring-opacity-50 disabled:opacity-50 disabled:cursor-not-allowed"
        >
          Load Sample Data
        </button>
        
        <button
          onClick={processJson}
          disabled={loading || !jsonInput.trim()}
          className="flex-1 px-6 py-3 bg-blue-600 text-white rounded-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 disabled:opacity-50 disabled:cursor-not-allowed"
        >
          {loading ? 'Processing...' : 'Generate Excel File'}
        </button>
      </div>
      
      {status && (
        <div className={`p-3 rounded-md w-full ${
          status.includes('Error') 
            ? 'bg-red-100 text-red-800 border border-red-200' 
            : 'bg-green-100 text-green-800 border border-green-200'
        }`}>
          {status}
        </div>
      )}
    </div>
  );
};

export default JsonToExcelConverter;