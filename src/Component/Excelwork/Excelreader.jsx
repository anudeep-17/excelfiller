"use client";
import React, {useState} from 'react';
import * as XLSX from 'xlsx';

function Excelreader(props) {
    const [data, setData] = useState([]);

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    readExcel(file);
  };

  const readExcel = (file) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      const data = event.target.result;
      const workbook = XLSX.read(data, { type: 'binary' });
      console.log(workbook);
      // Allow user selection or iterate over sheets
      const sheetName = props.sheetName || workbook.SheetNames[1]; // Use props if provided, otherwise default to first sheet
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      
      setData(jsonData);
    };

    reader.onerror = (error) => {
      console.error('Error reading the file:', error);
    };

    reader.readAsBinaryString(file);
  };
  
  return (
<div>
      <h1>Excel Data</h1>
      <input type="file" onChange={handleFileChange} />
      {data.length > 0 && (
        console.log(data),
        <table className="table">
          <thead>
            <tr>
              {Object.keys(data[0]).map((key) => (
                <th key={key}>{key}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, index) => (
              <tr key={index}>
                {Object.values(row).map((value, index) => (
                  <td key={index}>{value}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

export default Excelreader;
