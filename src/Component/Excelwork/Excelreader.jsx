"use client";
import { Work } from '@mui/icons-material';
import { Box, Button, Typography } from '@mui/material';
import React, {useEffect, useState} from 'react';
import * as XLSX from 'xlsx';
 

function Excelreader(props) {
    const [Workbook, setWorkbook] = useState({});  
    const [Sheets, setSheets] = useState([]);
    const [SelectedSheet, setSelectedSheet] = useState(null);
    const [Filecontent, setFileContent] = useState([]);

   
  const handleFileChange = (event) => {
    const file = event.target.files[0];
    GetWorkoBookDetails(file);
  };

  const GetWorkoBookDetails = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // Now you can access workbook details
      console.log('Workbook Details:', workbook);
      setWorkbook(workbook);
      GetSheetNames(workbook);
    };
    reader.readAsArrayBuffer(file);
  };

  const GetSheetNames = (workbook) => {
    const sheetNames = workbook.SheetNames;
    console.log('Sheet Names:', sheetNames);
    setSheets(sheetNames);
  };

  const handleSheetSelect = (e) => {
    const sheetName = e.target.value;
    setSelectedSheet(sheetName);
    ParseSelectedSheet(sheetName);
  };

  const ParseSelectedSheet = (sheetName) => {
    const worksheet = Workbook.Sheets[sheetName];
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    console.log('Sheet Data:', sheetData);
    setFileContent(sheetData);
  };

  const readExcel = (file) => {
    if(file === undefined) return;

    const reader = new FileReader();

    reader.onload = (event) => {
      const data = event.target.result;
      const workbook = XLSX.read(data, { type: 'binary' });
      setWorkbook(workbook);
      setSheets(workbook.SheetNames);
      const sheetName = SheetNameSelected || workbook.SheetNames[1]; // Use props if provided, otherwise default to first sheet
      setSheetNameSelected(sheetName);
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      setfilecontent(jsonData);
    };

    reader.onerror = (error) => {
      console.error('Error reading the file:', error);
    };

    reader.readAsBinaryString(file);
  };


  
  return (
    // <Box
    //   sx={{
    //     display: 'flex',
    //     flexDirection: 'column',
    //     alignItems: 'center',
    //     justifyContent: 'start',
    //     height: '100vh',
    //   }}
    // >
    //   {/* <div>
    //   <h1>Excel Data</h1>
    //   <input type="file" onChange={handleFileChange} />
    //   {filecontent.length > 0 && (
    //     <table className="table">
    //       <thead>
    //         <tr>
    //           {Object.keys(data[0]).map((key) => (
    //             <th key={key}>{key}</th>
    //           ))}
    //         </tr>
    //       </thead>
    //       <tbody>
    //         {filecontent.map((row, index) => (
    //           <tr key={index}>
    //             {Object.values(row).map((value, index) => (
    //               <td key={index}>{value}</td>
    //             ))}
    //           </tr>
    //         ))}
    //       </tbody>
    //     </table>
    //   )}
    // </div> */}
    //   <Typography
    //     variant="h3"
    //     sx={{ marginBottom: 2 }}
    //   >  
    //   Excel Auto Filler.
    //   </Typography>
 
    //   <input
    //     accept=".xls,.xlsx"
    //     id="file-upload"
    //     multiple
    //     type="file"
    //     onChange={handleFileChange}
    //     style={{ display: 'none' }}
    //   />
    //   <label htmlFor="file-upload">
    //     <Button
    //       variant="contained"
    //       component="span"
    //       sx={{ marginTop: 2 }}
    //     >
    //       Upload file
    //     </Button>
    //   </label>

    // </Box>
    <div>
      <input type="file" onChange={handleFileChange} />
      <div>
        <select onChange={handleSheetSelect}>
          <option value="">Select Sheet</option>
          {Sheets.map((sheet, index) => (
            <option key={index} value={sheet}>
              {sheet}
            </option>
          ))}
        </select>
      </div>
      {SelectedSheet && (
        <div>
          <h3>Selected Sheet: {SelectedSheet}</h3>
          <pre>{JSON.stringify(Filecontent, null, 2)}</pre>
        </div>
      )}
    </div>
  );
}

export default Excelreader;
