import { useState, useRef, useCallback, useMemo, useEffect } from "react";
import "./App.css";
import { createEngine } from "./engine/core.js";

const TOTAL_ROWS = 50;
const TOTAL_COLS = 50;

export default function App() {
  // Engine instance is created once and reused across renders
  // Note: The engine maintains its own internal state, so React state is only used for UI updates
  const [engine] = useState(() => createEngine(TOTAL_ROWS, TOTAL_COLS));
  const [version, setVersion] = useState(0);
  const [selectedCell, setSelectedCell] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState("");
  // Cell styles are stored separately from engine data
  // Format: { "row,col": { bold: bool, italic: bool, ... } }
  const [cellStyles, setCellStyles] = useState({});
  const cellInputRef = useRef(null);
  const gridRef = useRef(null);

  const [openFilterCol, setOpenFilterCol] = useState(null);

  const [sortingConfig, setSortingConfig] = useState({
    col: null,
    direction: null,
  });

  const [filterConfig, setFilterConfig] = useState({
    col: null,
    values: new Set(),
  });

  const forceRerender = useCallback(() => setVersion((v) => v + 1), []);

  // ────── Cell style helpers ──────

  const getCellStyle = useCallback(
    (row, col) => {
      const key = `${row},${col}`;
      return (
        cellStyles[key] || {
          bold: false,
          italic: false,
          underline: false,
          bg: "white",
          color: "#202124",
          align: "left",
          fontSize: 13,
        }
      );
    },
    [cellStyles],
  );

  const updateCellStyle = useCallback(
    (row, col, updates) => {
      const key = `${row},${col}`;
      setCellStyles((prev) => ({
        ...prev,
        [key]: { ...getCellStyle(row, col), ...updates },
      }));
    },
    [getCellStyle],
  );

  // ────── Cell editing ──────

  const startEditing = useCallback(
    (row, col) => {
      setSelectedCell({ r: row, c: col });
      setEditingCell({ r: row, c: col });
      const cellData = engine.getCell(row, col);
      setEditValue(cellData.raw);
      setTimeout(() => cellInputRef.current?.focus(), 0);
    },
    [engine],
  );

  const commitEdit = useCallback(
    (row, col) => {
      // Only commit if the value actually changed to avoid unnecessary recalculations
      const currentCell = engine.getCell(row, col);
      if (currentCell.raw !== editValue) {
        engine.setCell(row, col, editValue);
        forceRerender();
      }
      setEditingCell(null);
    },
    [engine, editValue, forceRerender],
  );

  const handleCellClick = useCallback(
    (row, col) => {
      if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
        commitEdit(editingCell.r, editingCell.c);
      }
      if (!editingCell || editingCell.r !== row || editingCell.c !== col) {
        startEditing(row, col);
      }
    },
    [editingCell, commitEdit, startEditing],
  );

  // ────── Keyboard navigation ──────

  const handleKeyDown = useCallback(
    (event, row, col) => {
      if (event.key === "Enter") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(Math.min(row + 1, engine.rows - 1), col);
      } else if (event.key === "Tab") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(row, Math.min(col + 1, engine.cols - 1));
      } else if (event.key === "Escape") {
        setEditValue(engine.getCell(row, col).raw);
        setEditingCell(null);
      } else if (event.key === "ArrowDown") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(Math.min(row + 1, engine.rows - 1), col);
      } else if (event.key === "ArrowUp") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(Math.max(row - 1, 0), col);
      } else if (event.key === "ArrowLeft") {
        event.preventDefault();
        commitEdit(row, col);
        if (col > 0) {
          startEditing(row, col - 1);
        } else if (row > 0) {
          startEditing(row - 1, engine.cols - 1);
        }
      } else if (event.key === "ArrowRight") {
        event.preventDefault();
        commitEdit(row, col);
        startEditing(row, Math.min(col + 1, engine.cols - 1));
      }
    },
    [engine, commitEdit, startEditing],
  );

  // ────── Formula bar handlers ──────

  const handleFormulaBarKeyDown = useCallback(
    (event) => {
      if (!editingCell) return;
      handleKeyDown(event, editingCell.r, editingCell.c);
    },
    [editingCell, handleKeyDown],
  );

  const handleFormulaBarFocus = useCallback(() => {
    if (selectedCell && !editingCell) {
      setEditingCell(selectedCell);
      setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw);
    }
  }, [selectedCell, editingCell, engine]);

  const handleFormulaBarChange = useCallback(
    (value) => {
      if (!editingCell && selectedCell) setEditingCell(selectedCell);
      setEditValue(value);
    },
    [editingCell, selectedCell],
  );

  // ────── Undo / Redo ──────

  const handleUndo = useCallback(() => {
    if (engine.undo()) forceRerender();
  }, [engine, forceRerender]);
  const handleRedo = useCallback(() => {
    if (engine.redo()) forceRerender();
  }, [engine, forceRerender]);

  // ────── Formatting toggles ──────

  const toggleBold = useCallback(() => {
    if (!selectedCell) return;
    const style = getCellStyle(selectedCell.r, selectedCell.c);
    updateCellStyle(selectedCell.r, selectedCell.c, { bold: !style.bold });
  }, [selectedCell, getCellStyle, updateCellStyle]);

  const toggleItalic = useCallback(() => {
    if (!selectedCell) return;
    const style = getCellStyle(selectedCell.r, selectedCell.c);
    updateCellStyle(selectedCell.r, selectedCell.c, { italic: !style.italic });
  }, [selectedCell, getCellStyle, updateCellStyle]);

  const toggleUnderline = useCallback(() => {
    if (!selectedCell) return;
    const style = getCellStyle(selectedCell.r, selectedCell.c);
    updateCellStyle(selectedCell.r, selectedCell.c, {
      underline: !style.underline,
    });
  }, [selectedCell, getCellStyle, updateCellStyle]);

  const changeFontSize = useCallback(
    (size) => {
      if (!selectedCell) return;
      updateCellStyle(selectedCell.r, selectedCell.c, { fontSize: size });
    },
    [selectedCell, updateCellStyle],
  );

  const changeAlignment = useCallback(
    (align) => {
      if (!selectedCell) return;
      updateCellStyle(selectedCell.r, selectedCell.c, { align });
    },
    [selectedCell, updateCellStyle],
  );

  const changeFontColor = useCallback(
    (color) => {
      if (!selectedCell) return;
      updateCellStyle(selectedCell.r, selectedCell.c, { color });
    },
    [selectedCell, updateCellStyle],
  );

  const changeBackgroundColor = useCallback(
    (color) => {
      if (!selectedCell) return;
      updateCellStyle(selectedCell.r, selectedCell.c, { bg: color });
    },
    [selectedCell, updateCellStyle],
  );

  // ────── Clear operations ──────

  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return;
    engine.setCell(selectedCell.r, selectedCell.c, "");
    forceRerender();
    // Remove style entry for cleared cell
    // Note: This deletes the style object entirely - if you need to preserve default styles,
    // you may want to set them explicitly rather than deleting
    const key = `${selectedCell.r},${selectedCell.c}`;
    setCellStyles((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
    setEditValue("");
  }, [selectedCell, engine, forceRerender]);

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++) {
      for (let c = 0; c < engine.cols; c++) {
        engine.setCell(r, c, "");
      }
    }
    forceRerender();
    setCellStyles({});
    setSelectedCell(null);
    setEditingCell(null);
    setEditValue("");
  }, [engine, forceRerender]);

  // ────── Row / Column operations ──────

  const insertRow = useCallback(() => {
    if (!selectedCell) return;
    engine.insertRow(selectedCell.r);
    forceRerender();
    setSelectedCell({ r: selectedCell.r + 1, c: selectedCell.c });
  }, [selectedCell, engine, forceRerender]);

  const deleteRow = useCallback(() => {
    if (!selectedCell) return;
    engine.deleteRow(selectedCell.r);
    forceRerender();
    if (selectedCell.r >= engine.rows) {
      setSelectedCell({ r: engine.rows - 1, c: selectedCell.c });
    }
  }, [selectedCell, engine, forceRerender]);

  const insertColumn = useCallback(() => {
    if (!selectedCell) return;
    engine.insertColumn(selectedCell.c);
    forceRerender();
    setSelectedCell({ r: selectedCell.r, c: selectedCell.c + 1 });
  }, [selectedCell, engine, forceRerender]);

  const deleteColumn = useCallback(() => {
    if (!selectedCell) return;
    engine.deleteColumn(selectedCell.c);
    forceRerender();
    if (selectedCell.c >= engine.cols) {
      setSelectedCell({ r: selectedCell.r, c: engine.cols - 1 });
    }
  }, [selectedCell, engine, forceRerender]);

  // ────── Derived state ──────

  const selectedCellStyle = useMemo(() => {
    return selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null;
  }, [selectedCell, getCellStyle]);

  const getColumnLabel = useCallback((col) => {
    let label = "";
    let num = col + 1;
    while (num > 0) {
      num--;
      label = String.fromCharCode(65 + (num % 26)) + label;
      num = Math.floor(num / 26);
    }
    return label;
  }, []);

  const selectedCellLabel = selectedCell
    ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}`
    : "No cell";

  // Formula bar shows the raw formula text, not the computed value
  // When editing, show the current editValue; otherwise show the cell's raw content
  // Note: This is different from the cell display, which shows computed values
  const formulaBarValue = editingCell
    ? editValue
    : selectedCell
      ? engine.getCell(selectedCell.r, selectedCell.c).raw
      : "";

  // ────── View Layer ──────

  const viewRows = useMemo(() => {
    let rows = Array.from({ length: engine.rows }, (_, i) => i).filter(
      (row) => {
        const cell = engine.getCell(row, sortingConfig.col ?? 0);
        return cell.raw !== "" || cell.computed !== null;
      },
    );

    if (filterConfig.col !== null && filterConfig.values.size > 0) {
      rows = rows.filter((row) => {
        const cell = engine.getCell(row, filterConfig.col);
        const value = cell.error ? cell.error : (cell.computed ?? cell.raw);
        return filterConfig.values.has(String(value));
      });
    }

    if (sortingConfig.col !== null && sortingConfig.direction) {
      rows = [...rows].sort((a, b) => {
        const cellA = engine.getCell(a, sortingConfig.col);
        const cellB = engine.getCell(b, sortingConfig.col);

        const getValue = (cell) => {
          if (cell.error) return null;

          const value = cell.computed ?? cell.raw;

          if (value === "" || value === null) return null;

          const num = Number(value);
          return isNaN(num) ? null : num;
        };

        const valA = getValue(cellA);
        const valB = getValue(cellB);

        if (valA === null && valB === null) return 0;
        if (valA === null) return 1;
        if (valB === null) return -1;

        if (sortingConfig.direction === "asc") {
          return valA - valB;
        }

        if (sortingConfig.direction === "desc") {
          return valB - valA;
        }

        return 0;
      });
    }
    return rows;
  }, [engine, version, sortingConfig, filterConfig]);

  const getUniqueColumnValues = (col) => {
    const values = new Set();

    for (let row = 0; row < engine.rows; row++) {
      const cell = engine.getCell(row, col);
      const value = cell.error ? cell.error : (cell.computed ?? cell.raw);

      values.add(String(value));
    }

    return Array.from(values);
  };

  const saveToLocalStorage = () => {
    try {
      const data = {
        rows: engine.rows,
        cols: engine.cols,
        cells: engine.cells,
      };
      localStorage.setItem("spreadsheet", JSON.stringify(data));
    } catch (error) {
      console.error("Save failed", error);
    }
  };


  useEffect(() => {
    const grid = gridRef.current;
    if (!grid) return;

    const handlePaste = (e) => {
      e.preventDefault();
      e.stopPropagation();

      if (editingCell) {
        setEditingCell(null);
      }

      const text = e.clipboardData.getData("text");
      const parsed = parseClipboardData(text);

      if (!selectedCell) return;

      const startRow = selectedCell.r;
      const startCol = selectedCell.c;

      parsed.forEach((rowData, rIndex) => {
        rowData.forEach((value, cIndex) => {
          const targetRow = startRow + rIndex;
          const targetCol = startCol + cIndex;

          if (targetRow < engine.rows && targetCol < engine.cols) {
            engine.setCell(targetRow, targetCol, value.trim());
          }
        });
      });

      forceRerender();
    };

    const handleCopy = (e) => {
      if (!selectedCell) return;

      e.preventDefault();

      const startRow = selectedCell.r;
      const startCol = selectedCell.c;

      let rows = [];

      for (let r = startRow; r < engine.rows; r++) {
        let row = [];
        let isEmptyRow = true;

        for (let c = startCol; c < engine.cols; c++) {
          const cell = engine.getCell(r, c);

          const value = cell?.computed ?? cell?.raw ?? "";

          if (value !== "") isEmptyRow = false;

          row.push(value);
        }

        if (isEmptyRow) break;

        rows.push(row.join("\t"));
      }

      const finalText = rows.join("\n");

      e.clipboardData.setData("text/plain", finalText);
    };

    grid.addEventListener("paste", handlePaste);
    grid.addEventListener("copy", handleCopy);

    return () => {
      grid.removeEventListener("paste", handlePaste);
      grid.removeEventListener("copy", handleCopy);
    };
  }, [selectedCell, engine]);

  const parseClipboardData = (text) => {
    const rows = text.split("\n");
    return rows.map((row) => row.split("\t"));
  };

  // ────── Render ──────

  return (
    <div className="app-wrapper">
      <div className="app-header">
        <h2 className="app-title">📊 Spreadsheet App</h2>
      </div>

      <div className="main-content">
        {/* ── Toolbar ── */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button
              className={`toolbar-btn bold-btn ${selectedCellStyle?.bold ? "active" : ""}`}
              onClick={toggleBold}
              title="Bold"
            >
              B
            </button>
            <button
              className={`toolbar-btn italic-btn ${selectedCellStyle?.italic ? "active" : ""}`}
              onClick={toggleItalic}
              title="Italic"
            >
              I
            </button>
            <button
              className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? "active" : ""}`}
              onClick={toggleUnderline}
              title="Underline"
            >
              U
            </button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size:</span>
            <select
              className="toolbar-select"
              value={selectedCellStyle?.fontSize || 13}
              onChange={(e) => changeFontSize(parseInt(e.target.value))}
            >
              {[8, 10, 11, 12, 13, 14, 16, 18, 20, 24].map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </div>

          <div className="toolbar-group">
            <button
              className={`align-btn ${selectedCellStyle?.align === "left" ? "active" : ""}`}
              onClick={() => changeAlignment("left")}
              title="Align Left"
            >
              ⬤←
            </button>
            <button
              className={`align-btn ${selectedCellStyle?.align === "center" ? "active" : ""}`}
              onClick={() => changeAlignment("center")}
              title="Align Center"
            >
              ⬤
            </button>
            <button
              className={`align-btn ${selectedCellStyle?.align === "right" ? "active" : ""}`}
              onClick={() => changeAlignment("right")}
              title="Align Right"
            >
              ⬤→
            </button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text:</span>
            <input
              type="color"
              value={selectedCellStyle?.color || "#000000"}
              onChange={(e) => changeFontColor(e.target.value)}
              title="Font color"
              style={{
                width: "32px",
                height: "32px",
                border: "1px solid #dadce0",
                cursor: "pointer",
                borderRadius: "4px",
              }}
            />
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Fill:</span>
            <select
              className="toolbar-select"
              value={selectedCellStyle?.bg || "white"}
              onChange={(e) => changeBackgroundColor(e.target.value)}
            >
              <option value="white">White</option>
              <option value="#ffff99">Yellow</option>
              <option value="#99ffcc">Green</option>
              <option value="#ffcccc">Red</option>
              <option value="#cce5ff">Blue</option>
              <option value="#e0ccff">Purple</option>
              <option value="#ffd9b3">Orange</option>
              <option value="#f0f0f0">Gray</option>
            </select>
          </div>

          <div className="toolbar-group">
            <button
              className="toolbar-btn"
              onClick={handleUndo}
              disabled={!engine.canUndo()}
              title="Undo"
            >
              ↶ Undo
            </button>
            <button
              className="toolbar-btn"
              onClick={handleRedo}
              disabled={!engine.canRedo()}
              title="Redo"
            >
              ↷ Redo
            </button>
          </div>

          <div className="toolbar-group">
            <button
              className="toolbar-btn"
              onClick={insertRow}
              title="Insert Row"
            >
              + Row
            </button>
            <button
              className="toolbar-btn"
              onClick={deleteRow}
              title="Delete Row"
            >
              - Row
            </button>
            <button
              className="toolbar-btn"
              onClick={insertColumn}
              title="Insert Column"
            >
              + Col
            </button>
            <button
              className="toolbar-btn"
              onClick={deleteColumn}
              title="Delete Column"
            >
              - Col
            </button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn danger" onClick={clearSelectedCell}>
              ✕ Cell
            </button>
            <button className="toolbar-btn danger" onClick={clearAllCells}>
              ✕ All
            </button>
          </div>
        </div>

        {/* ── Formula Bar ── */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <input
            className="formula-bar-input"
            value={formulaBarValue}
            onChange={(e) => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Select a cell then type, or enter a formula like =SUM(A1:A5)"
          />
        </div>

        {/* ── Grid ── */}
        <div
          ref={gridRef}
          className="grid-scroll"
          style={{ overflow: "visible" }}
          tabIndex={0}
        >
          <table className="grid-table">
            <thead>
              <tr>
                <th className="col-header-blank"></th>
                {Array.from({ length: engine.cols }, (_, colIndex) => (
                  <th
                    key={colIndex}
                    className="col-header"
                    style={{ position: "relative" }}
                  >
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                      }}
                    >
                      {/* SORT CLICK */}
                      <span
                        style={{ cursor: "pointer" }}
                        onClick={() => {
                          setSortingConfig((prev) => {
                            if (prev.col !== colIndex)
                              return { col: colIndex, direction: "asc" };
                            if (prev.direction === "asc")
                              return { col: colIndex, direction: "desc" };
                            return { col: null, direction: null };
                          });
                        }}
                      >
                        {getColumnLabel(colIndex)}
                        {sortingConfig.col === colIndex &&
                          (sortingConfig.direction === "asc"
                            ? " ↑"
                            : sortingConfig.direction === "desc"
                              ? " ↓"
                              : "")}
                      </span>

                      {/* FILTER BUTTON */}
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setOpenFilterCol((prev) =>
                            prev === colIndex ? null : colIndex,
                          );

                          setFilterConfig({
                            col: colIndex,
                            values: new Set(getUniqueColumnValues(colIndex)),
                          });
                        }}
                        style={{
                          border: "none",
                          background: "grey",
                          cursor: "pointer",
                          fontSize: "6px",
                          marginRight: "4px",
                        }}
                      >
                        ⏷
                      </button>
                    </div>

                    {/* DROPDOWN */}
                    {openFilterCol === colIndex && (
                      <div
                        style={{
                          position: "absolute",
                          top: "30px",
                          left: "0px",
                          background: "white",
                          border: "1px solid #ccc",
                          padding: "8px",
                          zIndex: 9999,
                          minWidth: "120px",
                          boxShadow: "0 2px 8px rgba(0,0,0,0.15)",
                        }}
                      >
                        {getUniqueColumnValues(colIndex).map((val) => (
                          <div key={val}>
                            <label>
                              <input
                                type="checkbox"
                                checked={filterConfig.values.has(val)}
                                onChange={(e) => {
                                  const newValues = new Set(
                                    filterConfig.values,
                                  );

                                  if (e.target.checked) newValues.add(val);
                                  else newValues.delete(val);

                                  setFilterConfig({
                                    col: colIndex,
                                    values: newValues,
                                  });
                                }}
                              />
                              {val}
                            </label>
                          </div>
                        ))}
                      </div>
                    )}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {viewRows.map((rowIndex, index) => (
                <tr key={rowIndex}>
                  <td className="row-header">{index + 1}</td>
                  {Array.from({ length: engine.cols }, (_, colIndex) => {
                    const isSelected =
                      selectedCell?.r === rowIndex &&
                      selectedCell?.c === colIndex;
                    const isEditing =
                      editingCell?.r === rowIndex &&
                      editingCell?.c === colIndex;
                    const cellData = engine.getCell(rowIndex, colIndex);
                    const style = cellStyles[`${rowIndex},${colIndex}`] || {};
                    const displayValue = cellData.error
                      ? cellData.error
                      : cellData.computed !== null && cellData.computed !== ""
                        ? String(cellData.computed)
                        : cellData.raw;

                    return (
                      <td
                        key={colIndex}
                        className={`cell ${isSelected ? "selected" : ""}`}
                        style={{ background: style.bg || "white" }}
                        onMouseDown={(e) => {
                          e.preventDefault();
                          handleCellClick(rowIndex, colIndex);
                        }}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            className="cell-input"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={() => commitEdit(rowIndex, colIndex)}
                            onKeyDown={(e) =>
                              handleKeyDown(e, rowIndex, colIndex)
                            }
                            ref={isSelected ? cellInputRef : undefined}
                            style={{
                              fontWeight: style.bold ? "bold" : "normal",
                              fontStyle: style.italic ? "italic" : "normal",
                              textDecoration: style.underline
                                ? "underline"
                                : "none",
                              color: style.color || "#202124",
                              fontSize: (style.fontSize || 13) + "px",
                              textAlign: style.align || "left",
                              background: style.bg || "white",
                            }}
                          />
                        ) : (
                          <div
                            className={`cell-display align-${style.align || "left"} ${cellData.error ? "error" : ""}`}
                            style={{
                              fontWeight: style.bold ? "bold" : "normal",
                              fontStyle: style.italic ? "italic" : "normal",
                              textDecoration: style.underline
                                ? "underline"
                                : "none",
                              color: cellData.error
                                ? "#d93025"
                                : style.color || "#202124",
                              fontSize: (style.fontSize || 13) + "px",
                            }}
                          >
                            {displayValue}
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <p className="footer-hint">
          Click a cell to edit · Enter/Tab/Arrow keys to navigate · Formulas:
          =A1+B1 · =SUM(A1:A5) · =AVG(A1:A5) · =MAX(A1:A5) · =MIN(A1:A5)
        </p>
      </div>
    </div>
  );
}
