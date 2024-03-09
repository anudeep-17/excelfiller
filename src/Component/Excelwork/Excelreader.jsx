"use client";
import { Work } from '@mui/icons-material';
import { Box, Button, Typography } from '@mui/material';
import React, {useEffect, useState} from 'react';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import FormControl from '@mui/material/FormControl';
import Select from '@mui/material/Select';
import TextField from '@mui/material/TextField';
import * as XLSX from 'xlsx';

const Excelreader = () => {
  const [Workbook, setWorkbook] = useState({});
  const [Sheets, setSheets] = useState([]);
  const [SelectedSheet, setSelectedSheet] = useState(null);
  const [Filecontent, setFileContent] = useState([]);
  const [EditRange, setEditRange] = useState({ start: { row: 1, col: 1 }, end: { row: 10, col: 10 } }); 

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    GetWorkoBookDetails(file);
  };

  const GetWorkoBookDetails = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      setWorkbook(workbook);
      GetSheetNames(workbook);
    };
    reader.readAsArrayBuffer(file);
  };

  const GetSheetNames = (workbook) => {
    const sheetNames = workbook.SheetNames;
    setSheets(sheetNames);
  };

  const handleSheetSelect = (e) => {
    const sheetName = e.target.value;
    setSelectedSheet(sheetName);
    // Reset the edit range when a new sheet is selected
    setEditRange({ start: { row: 1, col: 1 }, end: { row: 10, col: 10 } });
    ParseSelectedSheet(sheetName);
  };

  const ParseSelectedSheet = (sheetName) => {
    const worksheet = Workbook.Sheets[sheetName];
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    setFileContent(sheetData);
  };

  const handleCellEdit = (e, rowIndex, cellIndex) => {
    if (
      rowIndex < EditRange.start.row ||
      rowIndex > EditRange.end.row ||
      cellIndex < EditRange.start.col ||
      cellIndex > EditRange.end.col
    ) {
      return;
    }
    const newValue = e.target.innerText;
    const newData = [...Filecontent];
    newData[rowIndex][cellIndex] = newValue;
    setFileContent(newData);
  };

  const handleSaveExcel = () => {
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(Filecontent);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, SelectedSheet);
    XLSX.writeFile(newWorkbook, `Modified_${SelectedSheet}.xlsx`);
  };

  const handleRangeChange = (e, type, index) => {
    const value = parseInt(e.target.value);
    if (isNaN(value)) {
      return;
    }
    const newRange = { ...EditRange };
    newRange[type][index] = value;
    setEditRange(newRange);
  };

  return (
    <Box
      sx={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'start',
        height: '100vh',
      }}
    >
      <Typography variant="h3" sx={{ marginBottom: 2 }}>
        Excel Auto Filler
      </Typography>

      <input
        accept=".xls,.xlsx"
        id="file-upload"
        multiple
        type="file"
        onChange={handleFileChange}
        style={{ display: 'none' }}
      />
      <label htmlFor="file-upload">
        <Button variant="contained" component="span" sx={{ marginTop: 2 }}>
          Upload file
        </Button>
      </label>

      <FormControl sx={{ width: 800, m: 2 }}>
        <InputLabel id="demo-simple-select-label">
          Sheets From Selected Sheet
        </InputLabel>
        <Select
          labelId="demo-simple-select-label"
          id="demo-simple-select"
          value={SelectedSheet || ''}
          label="Sheets"
          onChange={handleSheetSelect}
        >
          {Sheets.map((sheet, index) => (
            <MenuItem key={index} value={sheet}>
              {sheet}
            </MenuItem>
          ))}
        </Select>
      </FormControl>

      {SelectedSheet && (
        <div>
          <h3>Selected Sheet: {SelectedSheet}</h3>
          <FormControl sx={{ width: '100px', marginBottom: 2 }}>
            <InputLabel htmlFor="start-row">Start Row</InputLabel>
            <TextField
              id="start-row"
              type="number"
              value={EditRange.start.row}
              onChange={(e) => handleRangeChange(e, 'start', 'row')}
            />
          </FormControl>
          <FormControl sx={{ width: '100px', marginBottom: 2 }}>
            <InputLabel htmlFor="start-col">Start Column</InputLabel>
            <TextField
              id="start-col"
              type="number"
              value={EditRange.start.col}
              onChange={(e) => handleRangeChange(e, 'start', 'col')}
            />
          </FormControl>
          <FormControl sx={{ width: '100px', marginBottom: 2 }}>
            <InputLabel htmlFor="end-row">End Row</InputLabel>
            <TextField
              id="end-row"
              type="number"
              value={EditRange.end.row}
              onChange={(e) => handleRangeChange(e, 'end', 'row')}
            />
          </FormControl>
          <FormControl sx={{ width: '100px', marginBottom: 2 }}>
            <InputLabel htmlFor="end-col">End Column</InputLabel>
            <TextField
              id="end-col"
              type="number"
              value={EditRange.end.col}
              onChange={(e) => handleRangeChange(e, 'end', 'col')}
            />
          </FormControl>
          <table border="1">
            <tbody>
              {Filecontent.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {row.map((cell, cellIndex) => (
                    <td
                      key={cellIndex}
                      contentEditable={
                        rowIndex >= EditRange.start.row &&
                        rowIndex <= EditRange.end.row &&
                        cellIndex >= EditRange.start.col &&
                        cellIndex <= EditRange.end.col
                      }
                      suppressContentEditableWarning
                      onBlur={(e) => handleCellEdit(e, rowIndex, cellIndex)}
                    >
                      {cell}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          <Button
            variant="contained"
            onClick={handleSaveExcel}
            sx={{ marginTop: 2 }}
          >
            Save Excel
          </Button>
        </div>
      )}
    </Box>
  );
};

export default Excelreader;