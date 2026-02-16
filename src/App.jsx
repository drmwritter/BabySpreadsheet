import { useState, useRef, useEffect } from 'react';
import { DataGrid, useGridApiRef } from '@mui/x-data-grid';
import Button from '@mui/material/Button';
import TextField from '@mui/material/TextField';
import InputAdornment from '@mui/material/InputAdornment';
import { createTheme, ThemeProvider } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';
import Box from '@mui/material/Box';
import Menu from '@mui/material/Menu';
import MenuItem from '@mui/material/MenuItem';
import Divider from '@mui/material/Divider';
import Typography from '@mui/material/Typography';
import './App.css';

const lightTheme = createTheme({
  palette: {
    mode: 'light',
    background: {
      default: '#f5f5f5',
      paper: '#ffffff',
    },
  },
});

const DEFAULT_ROW_HEIGHT = 26;

// --- Initial State --- //
const initialColumns = [
  { field: 'A', headerName: 'A', width: 150, editable: true, sortable: false },
  { field: 'B', headerName: 'B', width: 150, editable: true, sortable: false },
  { field: 'C', headerName: 'C', width: 150, editable: true, sortable: false },
  { field: 'D', headerName: 'D', width: 150, editable: true, sortable: false },
  { field: 'E', headerName: 'E', width: 150, editable: true, sortable: false },
  { field: 'F', headerName: 'F', width: 150, editable: true, sortable: false },
  { field: 'G', headerName: 'G', width: 150, editable: true, sortable: false },
  { field: 'H', headerName: 'H', width: 150, editable: true, sortable: false },
  { field: 'I', headerName: 'I', width: 150, editable: true, sortable: false },
  { field: 'J', headerName: 'J', width: 150, editable: true, sortable: false },
];

const initialRows = Array.from({ length: 20 }, (_, index) => ({
  id: index + 1,
  _height: DEFAULT_ROW_HEIGHT,
  ...initialColumns.reduce((acc, col) => ({ ...acc, [col.field]: '' }), {}),
}));

const getInitialState = () => ({
    columns: initialColumns,
    rawData: initialRows
});

// --- Helper Functions --- //

/**
 * Generates a column field name based on an index (0 -> A, 1 -> B, etc.)
 */
const getColumnName = (index) => {
    let name = '';
    let i = index;
    do {
        name = String.fromCharCode(65 + (i % 26)) + name;
        i = Math.floor(i / 26) - 1;
    } while (i >= 0);
    return name;
};

/**
 * Gets the 0-based index from a column name ('A' -> 0, 'B' -> 1, etc.)
 */
const getColumnIndex = (colName) => {
    let index = 0;
    for (let i = 0; i < colName.length; i++) {
        index = index * 26 + (colName.charCodeAt(i) - 65 + 1);
    }
    return index - 1;
};


/**
 * Parses a cell ID (e.g., 'A1') into its column and row parts.
 */
const parseCellId = (cellId) => {
    const match = cellId.match(/([A-Z]+)(\d+)/);
    if (!match) return null;
    return { col: match[1], row: parseInt(match[2], 10) };
};


// --- Main App Component --- //

function App() {
  const [rows, setRows] = useState(initialRows);
  const [columns, setColumns] = useState(initialColumns);
  const [rawData, setRawData] = useState(initialRows);
  const [savedData, setSavedData] = useState(() => JSON.parse(JSON.stringify(getInitialState())));
  const [isModified, setIsModified] = useState(false);

  const [rowSelectionModel, setRowSelectionModel] = useState([]);
  const [columnSelectionModel, setColumnSelectionModel] = useState([]);
  const [cellSelectionModel, setCellSelectionModel] = useState(null);
  const [formulaBarInput, setFormulaBarInput] = useState('');
  const [fileName, setFileName] = useState('spreadsheet.json');
  const [selectedRowHeightInput, setSelectedRowHeightInput] = useState('');
  const fileInputRef = useRef(null);
  const apiRef = useGridApiRef();
  const [anchorEl, setAnchorEl] = useState(null);

  useEffect(() => {
    const currentState = JSON.stringify({ columns, rawData });
    const savedState = JSON.stringify(savedData);
    setIsModified(currentState !== savedState);
  }, [columns, rawData, savedData]);

  const getNumericValue = (val) => {
    // Return numbers directly
    if (typeof val === 'number' && isFinite(val)) return val;
    // For non-strings or empty strings, no numeric value
    if (typeof val !== 'string' || val.trim() === '') return null;
    
    const num = Number(val);
    // If conversion results in a finite number, return it
    if (!isNaN(num) && isFinite(num)) {
      return num;
    }
    
    return null; // Otherwise, no numeric value
  };

  const evaluateFormula = (formula, data, visited = new Set(), idToRowNumberMap) => {
    if (typeof formula !== 'string' || !formula.startsWith('=')) {
        return formula;
    }

    const formulaBody = formula.substring(1).toUpperCase();

    // --- 1. Resolve range-based functions (SUM, AVERAGE, MEDIAN, MEAN) ---
    const rangeFunctionRegex = /(SUM|AVERAGE|MEDIAN|MEAN)\(([A-Z]+\d+):([A-Z]+\d+)\)/g;
    const withFunctionsResolved = formulaBody.replace(rangeFunctionRegex, (match, func, startCell, endCell) => {
        const start = parseCellId(startCell);
        const end = parseCellId(endCell);

        if (!start || !end) return '#ERROR';

        const startCol = getColumnIndex(start.col);
        const endCol = getColumnIndex(end.col);
        const startRow = start.row;
        const endRow = end.row;

        const rangeValues = [];

        // Collect all numeric values in the range
        for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
            for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
                const colName = getColumnName(c);
                const cellId = `${colName}${r}`;

                if (visited.has(cellId)) return '#REF!';
                const rowData = data[r - 1]; // Use 1-based index directly on the data array
                if (rowData && rowData.hasOwnProperty(colName)) {
                    const newVisited = new Set(visited);
                    newVisited.add(cellId);
                    const cellValue = evaluateFormula(rowData[colName], data, newVisited);

                    if (typeof cellValue === 'string' && cellValue.startsWith('#')) {
                        return cellValue; // Propagate error from within the range
                    }
                    const numericValue = getNumericValue(cellValue);
                    if (numericValue !== null) {
                        rangeValues.push(numericValue);
                    }
                }
            }
        }

        // Evaluate the collected values based on the function
        switch (func) {
            case 'SUM':
                return rangeValues.reduce((a, b) => a + b, 0);
            case 'AVERAGE':
            case 'MEAN':
                if (rangeValues.length === 0) return '#DIV/0!';
                const sum = rangeValues.reduce((a, b) => a + b, 0);
                return sum / rangeValues.length;
            case 'MEDIAN':
                if (rangeValues.length === 0) return '#NUM!';
                rangeValues.sort((a, b) => a - b);
                const mid = Math.floor(rangeValues.length / 2);
                return rangeValues.length % 2 !== 0
                    ? rangeValues[mid]
                    : (rangeValues[mid - 1] + rangeValues[mid]) / 2;
            default:
                return '#NAME?'; // Should not happen with the regex
        }
    });

    if (withFunctionsResolved.includes('#')) return withFunctionsResolved;

    // --- 2. Resolve single cell references ---
    const cellRegex = /([A-Z]+)(\d+)/g;
    const withCellsResolved = withFunctionsResolved.replace(cellRegex, (match, col, rowStr) => {
        const row = parseInt(rowStr, 10);
        const cellId = `${col}${row}`;

        if (visited.has(cellId)) return '#REF!';

        const rowData = data[row - 1]; // Use 1-based index directly
        if (rowData && rowData.hasOwnProperty(col)) {
            const newVisited = new Set(visited);
            newVisited.add(cellId);
            const evaluated = evaluateFormula(rowData[col], data, newVisited);

            if (typeof evaluated === 'string' && evaluated.startsWith('#')) {
                return evaluated; // Propagate errors
            }
            
            const numericValue = getNumericValue(evaluated);
            return numericValue !== null ? numericValue : 0;
        }
        return 0; // Cell not found or out of bounds
    });

    if (withCellsResolved.includes('#')) return withCellsResolved;

    // --- 3. Evaluate the final expression ---
    try {
        // eslint-disable-next-line no-new-func
        const result = new Function(`return ${withCellsResolved.replace(/#REF!/g, 'NaN')}`)();
        if (result === '#REF!' || isNaN(result)) return '#REF!';
        if (!isFinite(result)) return "#DIV/0!";
        return result;
    } catch (error) {
        return "#ERROR";
    }
  };

  useEffect(() => {
    const idToRowNumberMap = new Map();
    rawData.forEach((row, index) => {
      idToRowNumberMap.set(row.id, index + 1);
    });

    const newRows = rawData.map(rawRow => {
        const newRow = { ...rawRow };
        Object.keys(newRow).forEach(field => {
            if (field === 'id' || field === '_height') return;
            const currentCellId = `${field}${idToRowNumberMap.get(rawRow.id)}`;
            const formula = rawRow[field];
            newRow[field] = evaluateFormula(formula, rawData, new Set([currentCellId]), idToRowNumberMap);
        });
        return newRow;
    });
    setRows(newRows);
  }, [rawData]);

  useEffect(() => {
    if (cellSelectionModel) {
        const { id, field } = cellSelectionModel;
        const row = rawData.find(r => r.id === id);
        if (row && row.hasOwnProperty(field)) {
            const value = row[field];
            setFormulaBarInput(value === null || value === undefined ? '' : String(value));
        } else {
            setFormulaBarInput('');
        }
    } else {
        setFormulaBarInput('');
    }
  }, [cellSelectionModel, rawData]);

  useEffect(() => {
    if (rowSelectionModel.length > 0) {
      const firstSelectedRowId = rowSelectionModel[0];
      const row = rawData.find(r => r.id === firstSelectedRowId);
      if (row) {
        setSelectedRowHeightInput(row._height || '');
      }
    } else {
      setSelectedRowHeightInput('');
    }
  }, [rowSelectionModel, rawData]);

  const resetGrid = () => {
    const initialState = getInitialState();
    setColumns(initialState.columns);
    setRawData(initialState.rawData);
    setSavedData(JSON.parse(JSON.stringify(initialState)));
    setFileName('spreadsheet.json');
  };

  const handleNew = () => {
    if (isModified) {
        if (!window.confirm("You have unsaved changes. Are you sure you want to create a new spreadsheet and lose your changes?")) {
            return;
        }
    }
    resetGrid();
  };

  const processRowUpdate = (newRow, oldRow) => {
    const updatedRawData = rawData.map(row => (row.id === oldRow.id ? { ...row, ...newRow } : row));
    setRawData(updatedRawData);
    return newRow;
  };

  const addRow = () => {
    const newId = rawData.length > 0 ? Math.max(...rawData.map(r => r.id)) + 1 : 1;
    const newRow = { id: newId, _height: DEFAULT_ROW_HEIGHT, ...columns.reduce((acc, col) => ({ ...acc, [col.field]: '' }), {}) };
    setRawData(currentRawData => [...currentRawData, newRow]);
  };

    const deleteRow = () => {
        if (rowSelectionModel.length === 0) return;

        const oldToNewRowMap = new Map();
        let newRowNum = 1;
        rawData.forEach((row, index) => {
            if (!rowSelectionModel.includes(row.id)) {
                oldToNewRowMap.set(index + 1, newRowNum++);
            }
        });

        const updateFormulasForDeletion = (data, mapping) => {
            const cellRefRegex = /([A-Z]+)(\d+)/g;
            return data.map(row => {
                const newRow = { ...row };
                for (const field in newRow) {
                    const value = newRow[field];
                    if (typeof value === 'string' && value.startsWith('=')) {
                        newRow[field] = value.replace(cellRefRegex, (match, col, rowNumStr) => {
                            const oldRowNum = parseInt(rowNumStr, 10);
                            const newRowNum = mapping.get(oldRowNum);
                            if (newRowNum) {
                                return `${col}${newRowNum}`;
                            } else {
                                return '#REF!';
                            }
                        });
                    }
                }
                return newRow;
            });
        };

        const keptRows = rawData.filter(row => !rowSelectionModel.includes(row.id));
        const updatedKeptRows = updateFormulasForDeletion(keptRows, oldToNewRowMap);

        setRawData(updatedKeptRows);
        setRowSelectionModel([]);
    };

  const addColumn = () => {
    setColumns(currentColumns => {
        const newField = getColumnName(currentColumns.length);
        const newColumn = { field: newField, headerName: newField, width: 150, editable: true, sortable: false };
        setRawData(currentRawData => currentRawData.map(row => ({ ...row, [newField]: '' })));
        return [...currentColumns, newColumn];
    });
  };

    const insertRow = (offset) => { // offset 0 for above, 1 for below
        if (!cellSelectionModel) return;

        const { id: selectedRowId } = cellSelectionModel;
        const selectedRowIndex = rawData.findIndex(r => r.id === selectedRowId);
        if (selectedRowIndex === -1) return;

        const insertionPoint = selectedRowIndex + offset;

        const updateFormulasForRowChange = (data, startIndex, amount) => {
            const cellRefRegex = /([A-Z]+)(\d+)/g;
            return data.map(row => {
                const newRow = { ...row };
                for (const field in newRow) {
                    const value = newRow[field];
                    if (typeof value === 'string' && value.startsWith('=')) {
                        newRow[field] = value.replace(cellRefRegex, (match, col, rowNumStr) => {
                            const rowNum = parseInt(rowNumStr, 10);
                            if (rowNum >= startIndex) {
                                return `${col}${rowNum + amount}`;
                            }
                            return match;
                        });
                    }
                }
                return newRow;
            });
        };

        const dataWithUpdatedFormulas = updateFormulasForRowChange(rawData, insertionPoint + 1, 1);

        const newId = (rawData.length > 0 ? Math.max(...rawData.map(r => r.id)) : 0) + 1;
        const newRowData = { id: newId, _height: DEFAULT_ROW_HEIGHT, ...columns.reduce((acc, col) => ({ ...acc, [col.field]: '' }), {}) };

        const finalRawData = [...dataWithUpdatedFormulas];
        finalRawData.splice(insertionPoint, 0, newRowData);

        setRawData(finalRawData);
        setCellSelectionModel(null);
    };

    const insertColumn = (offset) => { // offset 0 for left, 1 for right
        if (!cellSelectionModel) return;

        const { field: selectedColField } = cellSelectionModel;
        const selectedColIndex = columns.findIndex(c => c.field === selectedColField);
        if (selectedColIndex === -1) return;

        const insertionPoint = selectedColIndex + offset;

        const updateFormulasForColumnChange = (data, startIndex, amount) => {
            const cellRefRegex = /([A-Z]+)(\d+)/g;
            return data.map(row => {
                const newRow = { ...row };
                for (const field in newRow) {
                    const value = newRow[field];
                    if (typeof value === 'string' && value.startsWith('=')) {
                        newRow[field] = value.replace(cellRefRegex, (match, colName, rowNumStr) => {
                            const colIndex = getColumnIndex(colName);
                            if (colIndex >= startIndex) {
                                return `${getColumnName(colIndex + amount)}${rowNumStr}`;
                            }
                            return match;
                        });
                    }
                }
                return newRow;
            });
        };

        const dataWithUpdatedFormulas = updateFormulasForColumnChange(rawData, insertionPoint, 1);

        const newTotalCols = columns.length + 1;
        const newColumns = [];
        const oldToNewFieldMap = new Map();

        let oldColIdx = 0;
        for (let i = 0; i < newTotalCols; i++) {
            const newField = getColumnName(i);
            newColumns.push({ field: newField, headerName: newField, width: 150, editable: true, sortable: false });
            if (i === insertionPoint) continue;
            oldToNewFieldMap.set(columns[oldColIdx].field, newField);
            oldColIdx++;
        }

        const finalRawData = dataWithUpdatedFormulas.map(row => {
            const newRow = { id: row.id };
            for (const [oldField, newField] of oldToNewFieldMap.entries()) {
                if (row.hasOwnProperty(oldField)) {
                    newRow[newField] = row[oldField];
                }
            }
            newRow[getColumnName(insertionPoint)] = '';
            return newRow;
        });

        setColumns(newColumns);
        setRawData(finalRawData);
        setCellSelectionModel(null);
        setColumnSelectionModel([]);
    };


    const deleteSelectedColumns = () => {
        if (columnSelectionModel.length === 0) return;
        if (columns.length - columnSelectionModel.length === 0) {
            console.warn("Cannot delete all columns.");
            return;
        }

        const oldToNewColIndexMap = new Map();
        let newColIndex = 0;
        columns.forEach((col, index) => {
            if (!columnSelectionModel.includes(col.field)) {
                oldToNewColIndexMap.set(index, newColIndex++);
            }
        });

        const updateFormulasForColDeletion = (data, mapping) => {
            const cellRefRegex = /([A-Z]+)(\d+)/g;
            return data.map(row => {
                const newRow = { ...row };
                for (const field in newRow) {
                    const value = newRow[field];
                    if (typeof value === 'string' && value.startsWith('=')) {
                        newRow[field] = value.replace(cellRefRegex, (match, colName, rowNumStr) => {
                            const oldColIndex = getColumnIndex(colName);
                            const newColIndex = mapping.get(oldColIndex);
                            if (newColIndex !== undefined) {
                                return `${getColumnName(newColIndex)}${rowNumStr}`;
                            } else {
                                return '#REF!';
                            }
                        });
                    }
                }
                return newRow;
            });
        };

        const dataWithUpdatedFormulas = updateFormulasForColDeletion(rawData, oldToNewColIndexMap);

        const remainingColumns = columns.filter(col => !columnSelectionModel.includes(col.field));

        const oldToNewFieldMap = new Map();
        const newColumns = remainingColumns.map((col, index) => {
            const newField = getColumnName(index);
            oldToNewFieldMap.set(col.field, newField);
            return { ...col, field: newField, headerName: newField };
        });

        const newRawData = dataWithUpdatedFormulas.map(row => {
            const newRow = { id: row.id };
            for (const [oldField, newField] of oldToNewFieldMap.entries()) {
                if (row.hasOwnProperty(oldField)) {
                    newRow[newField] = row[oldField];
                }
            }
            return newRow;
        });

        setColumns(newColumns);
        setRawData(newRawData);
        setColumnSelectionModel([]);
    };

    const handleSave = () => {
        let finalFileName = fileName.trim();
        if (!finalFileName) {
            finalFileName = 'spreadsheet.json';
        }
        if (!finalFileName.endsWith('.json')) {
            finalFileName += '.json';
        }

        const dataToSave = { columns, rawData };
        const dataString = JSON.stringify(dataToSave, null, 2);
        const blob = new Blob([dataString], { type: 'application/json' });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = finalFileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        setSavedData(JSON.parse(dataString)); // Deep copy the saved state
    };

    const handleLoadClick = () => {
        if (isModified) {
            if (!window.confirm("You have unsaved changes. Are you sure you want to load a new spreadsheet and lose your changes?")) {
                return;
            }
        }
        fileInputRef.current.click();
    };

    const handleFileChange = (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const text = e.target.result;
                const loadedState = JSON.parse(text);
                if (loadedState.columns && loadedState.rawData) {
                    const sanitizedRawData = loadedState.rawData.map(row => ({...row, _height: row._height || DEFAULT_ROW_HEIGHT}));
                    const newState = { ...loadedState, rawData: sanitizedRawData };

                    setColumns(newState.columns);
                    setRawData(newState.rawData);
                    setSavedData(JSON.parse(JSON.stringify(newState)));
                    setFileName(file.name);
                    alert(`Spreadsheet ${file.name} loaded.`);
                } else {
                    alert("Invalid spreadsheet file format.");
                }
            } catch (error) {
                console.error("Error parsing file:", error);
                alert("Failed to load or parse the file.");
            }
        };
        reader.readAsText(file);

        event.target.value = null;
    };


  const handleColumnHeaderClick = (params, event) => {
    const { field } = params.colDef;
    if (field === '#') return;

    const isMultiSelect = event.ctrlKey || event.metaKey;

    setColumnSelectionModel(currentSelection => {
        const isSelected = currentSelection.includes(field);
        if (isMultiSelect) {
            return isSelected
                ? currentSelection.filter(f => f !== field)
                : [...currentSelection, field];
        } else {
            return isSelected && currentSelection.length === 1
                ? []
                : [field];
        }
    });
  };

  const handleCellClick = (params) => {
    if (params.field !== '#') {
      setCellSelectionModel({ id: params.id, field: params.field });
    }
  }

  const handleFormulaBarChange = (event) => {
      setFormulaBarInput(event.target.value);
  };

  const handleFormulaBarKeyDown = (event) => {
      if (event.key === 'Enter' && cellSelectionModel) {
          event.preventDefault();
          const { id, field } = cellSelectionModel;
          
          setRawData(currentRawData =>
              currentRawData.map(row =>
                  row.id === id ? { ...row, [field]: formulaBarInput } : row
              )
          );
          event.target.blur();
      }
  };

  const handleSelectedRowHeightKeyDown = (event) => {
    if (event.key === 'Enter') {
      const newHeight = parseInt(selectedRowHeightInput, 10);
      if (!isNaN(newHeight) && newHeight > 0) {
        setRawData(currentRawData => 
          currentRawData.map(row => 
            rowSelectionModel.includes(row.id) ? { ...row, _height: newHeight } : row
          )
        );
      }
      event.target.blur();
    }
  }

  const handleCellEditStart = (params) => {
    const { id, field } = params;
    const rawRow = rawData.find(r => r.id === id);
    if (rawRow) {
        apiRef.current.setEditCellValue({ id, field, value: rawRow[field] });
    }
  };

  const handleMenuClick = (event) => {
    setAnchorEl(event.currentTarget);
  };

  const getDisplayColumns = () => [
    {
      field: '#',
      headerName: '#',
      width: 90,
      sortable: false,
      filterable: false,
      renderHeader: () => <strong>#</strong>,
      renderCell: (params) => {
        const index = rawData.findIndex(r => r.id === params.id);
        return <strong>{index + 1}</strong>;
      },
    },
    ...columns.map(col => ({
        ...col,
        headerClassName: columnSelectionModel.includes(col.field) 
            ? 'column-header-selected' 
            : ''
    })),
  ];

  const menuOpen = Boolean(anchorEl);

  const handleMenuClose = () => {
    setAnchorEl(null);
  };


  return (
    <ThemeProvider theme={lightTheme}>
      <CssBaseline />
      <Box className="app-container">
        <div className="header">
          <TextField label="Filename" variant="outlined" size="small" value={fileName} onChange={e => setFileName(e.target.value)} sx={{ mr: 1 }} />
          <Box sx={{ display: 'flex', alignItems: 'center', mr: 2, ml: 1 }}>
            <Typography variant="subtitle2" sx={{ color: isModified ? 'error.main' : 'success.main', fontWeight: 'bold' }}>
              {isModified ? 'UNSAVED CHANGES' : 'All changes saved'}
            </Typography>
          </Box>
          <TextField 
            label="Selected Row Height"
            variant="outlined" 
            size="small" 
            type="number"
            value={selectedRowHeightInput}
            onChange={e => setSelectedRowHeightInput(e.target.value)}
            onKeyDown={handleSelectedRowHeightKeyDown}
            disabled={rowSelectionModel.length === 0}
            sx={{ mr: 1, width: 180 }}
          />
          <Button
            id="actions-button"
            aria-controls={menuOpen ? 'actions-menu' : undefined}
            aria-haspopup="true"
            aria-expanded={menuOpen ? 'true' : undefined}
            onClick={handleMenuClick}
            variant="contained"
          >
            Actions
          </Button>
          <Menu
            id="actions-menu"
            anchorEl={anchorEl}
            open={menuOpen}
            onClose={handleMenuClose}
            MenuListProps={{
              'aria-labelledby': 'actions-button',
            }}
          >
            <MenuItem onClick={() => { handleNew(); handleMenuClose(); }}>New</MenuItem>
            <MenuItem onClick={() => { handleSave(); handleMenuClose(); }} disabled={!isModified}>Save</MenuItem>
            <MenuItem onClick={() => { handleLoadClick(); handleMenuClose(); }}>Load</MenuItem>
            <Divider />
            <MenuItem onClick={() => { addRow(); handleMenuClose(); }}>Add Row (End)</MenuItem>
            <MenuItem onClick={() => { addColumn(); handleMenuClose(); }}>Add Column (End)</MenuItem>
            <Divider />
            <MenuItem onClick={() => { insertRow(0); handleMenuClose(); }} disabled={!cellSelectionModel}>Insert Row Above</MenuItem>
            <MenuItem onClick={() => { insertRow(1); handleMenuClose(); }} disabled={!cellSelectionModel}>Insert Row Below</MenuItem>
            <MenuItem onClick={() => { insertColumn(0); handleMenuClose(); }} disabled={!cellSelectionModel}>Insert Column Left</MenuItem>
            <MenuItem onClick={() => { insertColumn(1); handleMenuClose(); }} disabled={!cellSelectionModel}>Insert Column Right</MenuItem>
            <Divider />
            <MenuItem sx={{ color: 'error.main' }} onClick={() => { deleteRow(); handleMenuClose(); }} disabled={rowSelectionModel.length === 0}>Delete Selected Rows</MenuItem>
            <MenuItem sx={{ color: 'error.main' }} onClick={() => { deleteSelectedColumns(); handleMenuClose(); }} disabled={columnSelectionModel.length === 0}>Delete Selected Columns</MenuItem>
          </Menu>
          <input 
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            style={{ display: 'none' }}
            accept=".json,application/json"
          />
        </div>

        <Box sx={{ p: 1, borderBottom: 1, borderColor: 'divider' }}>
          <TextField
              fullWidth
              variant="outlined"
              size="small"
              label={cellSelectionModel ? `Formula for ${cellSelectionModel.field}${rawData.findIndex(r => r.id === cellSelectionModel.id) + 1}` : "Selected Cell Formula"}
              value={formulaBarInput}
              onChange={handleFormulaBarChange}
              onKeyDown={handleFormulaBarKeyDown}
              InputProps={{
                  startAdornment: <InputAdornment position="start"><i>fx</i></InputAdornment>,
              }}
          />
        </Box>

        <Box className="grid-container">
          <DataGrid
            apiRef={apiRef}
            key={`${JSON.stringify(columns)}-${rawData.length}`}
            rows={rows}
            columns={getDisplayColumns()}
            getRowHeight={({ id }) => {
              const row = rawData.find(r => r.id === id);
              return row ? row._height : DEFAULT_ROW_HEIGHT;
            }}
            processRowUpdate={processRowUpdate}
            onProcessRowUpdateError={(error) => console.error(error)}
            onCellEditStart={handleCellEditStart}
            checkboxSelection
            onRowSelectionModelChange={(newSelectionModel) => {
              setRowSelectionModel(newSelectionModel);
            }}
            rowSelectionModel={rowSelectionModel}
            onColumnHeaderClick={handleColumnHeaderClick}
            onCellClick={handleCellClick}
            cellClassName={params => {
                if (cellSelectionModel && params.id === cellSelectionModel.id && params.field === cellSelectionModel.field) {
                    return 'cell-selected';
                }
                return '';
            }}
            disableColumnMenu
            disableRowSelectionOnClick
            hideFooter
            showCellVerticalBorder={true}
          />
        </Box>
      </Box>
    </ThemeProvider>
  );
}

export default App;
