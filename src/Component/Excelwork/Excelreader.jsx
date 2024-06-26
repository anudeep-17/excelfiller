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
import Fab from '@mui/material/Fab';
import CachedIcon from '@mui/icons-material/Cached';
import Alert from '@mui/material/Alert';
import CheckIcon from '@mui/icons-material/Check';
import Snackbar from '@mui/material/Snackbar';
import AddIcon from '@mui/icons-material/Add';
import copy from 'copy-to-clipboard';

const Excelreader = () => {
  const [fileName, setFileName] = useState('');
  const [Workbook, setWorkbook] = useState({});
  const [Sheets, setSheets] = useState([]);
  const [SelectedSheet, setSelectedSheet] = useState(null);
  const [Filecontent, setFileContent] = useState([]);
  const [EditRange, setEditRange] = useState({ start: { row: 0, col: 0 }, end: { row: 10, col: 10 } }); 
  const [SaveAs, setSaveAs] = useState('');
  const [OriginalFilecontent, setOriginalFilecontent] = useState([]);
  
  const [startRow, setStartRow] = useState(0);
  const [endRow, setEndRow] = useState(0);
  
  // snackbar openers
  const [open_forupload, setOpen_forUpload] = useState(false);
  const [open_forSave, setOpen_forSave] = useState(false);
  const [open_forReload, setOpen_forReload] = useState(false);
  const [open_forcopy, setOpen_forcopy] = useState(false);

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    setFileName(file.name);
    GetWorkoBookDetails(file);
    setOpen_forUpload(true);
  };

  const handleClose = (event, reason) => {
    if (reason === 'clickaway') {
      return;
    }
    if(open_forupload){
      setOpen_forUpload(false);
    }
    if(open_forSave){
      setOpen_forSave(false);
    }
    if(open_forReload){
      setOpen_forReload(false);
    }
    if(open_forcopy){
      setOpen_forcopy(false);
    }
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
    console.log(sheetName);
    // Reset the edit range when a new sheet is selected
    setEditRange({ start: { row: 0, col: 0 }, end: { row: 90, col: 10 } });
    ParseSelectedSheet(sheetName);
  };

  const ParseSelectedSheet = (sheetName) => {
    const worksheet = Workbook.Sheets[sheetName];
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    setFileContent(sheetData);
    setOriginalFilecontent(JSON.parse(JSON.stringify(sheetData)));
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
    XLSX.writeFile(newWorkbook, `${SaveAs}.xlsx`);
    setOpen_forSave(true);
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

  const ReloadSheet = () => {
    setFileContent(JSON.parse(JSON.stringify(OriginalFilecontent)));
    setOriginalFilecontent(JSON.parse(JSON.stringify(OriginalFilecontent)));
    setOpen_forReload(true);
    setSaveAs('');
  }
  
  const handleCopy = () => {
    if (!Filecontent || Filecontent.length === 0) {
      alert('No file content available');
      return;
    }
  
    let tableRows = [];
    for (let i = EditRange.start.row; i <= EditRange.end.row; i++) {
      if (!Filecontent[i]) {
        continue;
      }
  
      let tableCols = [];
      for (let j = EditRange.start.col; j <= EditRange.end.col; j++) {
        if (!Filecontent[i][j]) {
          break;
        }
        tableCols.push(Filecontent[i][j]);
      }
      if (tableCols.length > 0) {
        tableRows.push(tableCols.join('\t'));
      }
    }
    copy(tableRows.join('\n'));
  };


  const calculateSum = (startRow, endRow, columnNumber) => {
    let sum = 0;
    for (let i = startRow; i <= endRow; i++) {
      sum += Number(Filecontent[i][columnNumber] || 0);
    } 
    alert(`The sum is: ${sum} && The sum over 100: ${(sum/220)*100}`);
 
  };

  return (
    <Box
      sx={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'start',
        height: '100vh',
        '& > :not(style)': { m: 1 } 
      }}
    >

      <Snackbar open={open_forupload} onClose={handleClose} autoHideDuration={6000} anchorOrigin={{ vertical: 'top', horizontal: 'right' }}>
          <Alert icon={<CheckIcon fontSize="inherit" />} severity="success">
            File Upoad successful please select the sheet you want to edit
          </Alert>
      </Snackbar>

      <Typography variant="h3" sx={{ marginBottom: 2 }}>
        Excel Auto Filler
      </Typography>

      {Filecontent && Filecontent.length > 0 ? <Fab variant="extended" size="medium" color="primary" sx={{ position: 'fixed', bottom: '20px', right: '20px' }} onClick={ReloadSheet}>
        <CachedIcon sx={{ mr: 1 }} />
        Reload the Sheet 
      </Fab>: null}

        {Filecontent && Filecontent.length > 0 ? <Fab variant="extended" size="medium" color="primary" sx={{ position: 'fixed', bottom: '80px', right: '20px'}} onClick={()=>{calculateSum(11,39,3)}}>
          <AddIcon sx={{ mr: 1 }} />
          Get Sum
        </Fab>: null
        }

      {Filecontent && Filecontent.length > 0 ? <Fab variant="extended" size="medium" color="primary" sx={{ position: 'fixed', bottom: '150px', right: '20px'}} onClick={handleCopy}>
        <AddIcon sx={{ mr: 1 }} />
        Copy to clipboard
      </Fab>: null
      }

      <Snackbar open={open_forReload} onClose={handleClose} autoHideDuration={6000} anchorOrigin={{ vertical: 'top', horizontal: 'right' }}>
          <Alert icon={<CheckIcon fontSize="inherit" />} severity="success">
            File Reload Successful
          </Alert>
      </Snackbar> 

      <input
        accept=".xls,.xlsx"
        id="file-upload"
        multiple
        type="file"
        onChange={handleFileChange}
        style={{ display: 'none' }}
      />
      <label htmlFor="file-upload">
        <Box
          sx={{
            display: 'flex',
            flexDirection: 'row',
            alignItems: 'center',
            alignContent : 'center',
            justifyContent: 'center',
            padding: 2,
            cursor: 'pointer',
            marginTop: 2 
          }}
        >
         <Button variant="contained" component="span">
          Upload file
        </Button>
        <Typography variant="h6" sx={{ marginLeft: 2}}>
          {fileName? fileName: 'No file selected'}
        </Typography>
        </Box> 
       
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
        <Box
          sx={{
            display: 'flex',
            flexDirection: 'column',
            alignContent:'start',
            justifyContent: 'start',
            marginTop: 2,
          }}
        >
          <Typography variant='h2'>
            Selected Sheet: {SelectedSheet}
          </Typography>
          {/* Colum stuff */}

          <FormControl sx={{ width: '100px', marginBottom: 2 }}>
            <Typography variant="h6">From Column</Typography>
            <TextField
              id="start-col"
              type="number"
              value={EditRange.start.col}
              onChange={(e) => handleRangeChange(e, 'start', 'col')}
            />
          </FormControl>
          <FormControl sx={{ width: '100px', marginBottom: 2 }}>
            <Typography variant="h6">To Column</Typography>
            <TextField
              id="start-col"
              type="number"
              value={EditRange.end.col}
              onChange={(e) => handleRangeChange(e, 'end', 'col')}
            />
          </FormControl>


          {/* Row stuff */}
          <FormControl sx={{ width: '100', marginBottom: 2 }}>
            <Typography variant="h6">From Row</Typography>
            <TextField
              id="start-row"
              type="number"
              value={EditRange.start.row}
              onChange={(e) => handleRangeChange(e, 'start', 'row')}
            />
          </FormControl>
          <FormControl sx={{ width: '100', marginBottom: 2 }}>
            <Typography variant="h6">To Row</Typography>
            <TextField
              id="start-row"
              type="number"
              value={EditRange.end.row}
              onChange={(e) => handleRangeChange(e, 'end', 'row')}
            />
          </FormControl>

        <table border="1">
          <tbody>
            {Filecontent.slice(EditRange.start.row, EditRange.end.row + 1).map((row, rowIndex) => {
              console.log(row);
              return (
                <tr key={rowIndex}>
                  {Array.from({ length: EditRange.end.col - EditRange.start.col + 1 }).map((_, cellIndex) => (
                    <td
                      key={cellIndex}
                      contentEditable
                      onBlur={(e) => handleCellEdit(e, rowIndex + EditRange.start.row, cellIndex + EditRange.start.col)}
                    >
                      {row[cellIndex + EditRange.start.col] !== null ? row[cellIndex + EditRange.start.col] : ''}
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>

        


          <Typography variant="h6" sx={{ marginTop: 2 }}>
            Enter how you want the file to be saved
          </Typography>
          <Box
            sx={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              marginTop: 2,
            }}
          >
            <Typography variant="h6">Save the file as</Typography>
            <TextField id="outlined-basic" label="Save the file as" variant="outlined" onChange={(e) => setSaveAs(e.target.value)} />
          </Box>

          <Button
            variant="contained"
            onClick={handleSaveExcel}
            sx={{ marginTop: 2 }}
          >
            Save Excel
          </Button>

          <Snackbar open={open_forSave} onClose={handleClose} autoHideDuration={6000} anchorOrigin={{ vertical: 'top', horizontal: 'right' }}>
              <Alert icon={<CheckIcon fontSize="inherit" />} severity="success">
                File Saved Successfully, Check Your Downloads
              </Alert>
          </Snackbar>


        </Box>
      )}
    </Box>
  );
};

export default Excelreader;