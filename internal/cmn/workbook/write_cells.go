// Copyright (C) 2026 Yota Hamada
// SPDX-License-Identifier: GPL-3.0-or-later

package workbook

import (
	"context"
	"fmt"
	"os"
	"sort"
	"strings"
	"time"
)

// CellValue is what one cell of a WriteCells request receives: a value,
// optionally pinned to a type, a formula, or nothing.
type CellValue struct {
	// Value is written as a cell value; ISO date and date-time strings
	// become dates unless Type pins them to string.
	Value any
	// Type pins Value the way a column type does.
	Type ColumnType
	// Formula is written as the cell's formula, with or without a leading =.
	Formula string
	// Clear removes the cell's value and formula and keeps its style.
	Clear bool
}

// ParseCells reads a cells option: a map from a cell address to a scalar,
// null, {value: v, type: t}, or {formula: text}.
func ParseCells(v any) (map[string]CellValue, error) {
	raw, ok := v.(map[string]any)
	if !ok {
		return nil, fmt.Errorf("cells must map cell addresses to values")
	}
	if len(raw) == 0 {
		return nil, fmt.Errorf("cells must not be empty")
	}
	cells := make(map[string]CellValue, len(raw))
	for addr, spec := range raw {
		if strings.TrimSpace(addr) == "" {
			return nil, fmt.Errorf("cells: cell address must not be empty")
		}
		cv, err := parseCellValue(spec)
		if err != nil {
			return nil, fmt.Errorf("cells.%s: %w", addr, err)
		}
		cells[addr] = cv
	}
	return cells, nil
}

func parseCellValue(spec any) (CellValue, error) {
	switch x := spec.(type) {
	case nil:
		return CellValue{Clear: true}, nil
	case string, bool, int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64, float32, float64:
		return CellValue{Value: normalizeScalar(x)}, nil
	case time.Time:
		// YAML reads an unquoted date such as 2026-10-01 as a time.
		return CellValue{Value: x}, nil
	case map[string]any:
		if formula, ok := x["formula"]; ok {
			text, isText := formula.(string)
			if !isText || strings.TrimSpace(text) == "" {
				return CellValue{}, fmt.Errorf("formula must not be empty")
			}
			if len(x) != 1 {
				return CellValue{}, fmt.Errorf("use a scalar, null, {value: v, type: t}, or {formula: text}")
			}
			return CellValue{Formula: strings.TrimSpace(text)}, nil
		}
		value, hasValue := x["value"]
		if !hasValue || len(x) > 2 {
			return CellValue{}, fmt.Errorf("use a scalar, null, {value: v, type: t}, or {formula: text}")
		}
		cv := CellValue{Value: normalizeScalar(value), Clear: value == nil}
		if typeSpec, hasType := x["type"]; hasType {
			text, isText := typeSpec.(string)
			if !isText {
				return CellValue{}, fmt.Errorf("type must be a column type")
			}
			t, err := ParseColumnType(text)
			if err != nil {
				return CellValue{}, fmt.Errorf("type: %w", err)
			}
			cv.Type = t
		} else if len(x) != 1 {
			return CellValue{}, fmt.Errorf("use a scalar, null, {value: v, type: t}, or {formula: text}")
		}
		return cv, nil
	default:
		return CellValue{}, fmt.Errorf("use a scalar, null, {value: v, type: t}, or {formula: text}")
	}
}

// WriteCellsOptions controls WriteCells.
type WriteCellsOptions struct {
	Password string
	// Sheet is the sheet for addresses that do not name one; the first
	// sheet by default.
	Sheet string
	// Cells maps addresses to what they receive. An address is a cell such
	// as B2, Sheet!B2, or 'My Sheet'!B2, or a defined name that refers to
	// one cell.
	Cells map[string]CellValue
	// Output, when set, receives the result while the workbook at path is
	// left as it was, so a template can be filled many times.
	Output  string
	InPlace bool
	DryRun  bool
	Lock    LockOptions
}

// WriteCells writes values and formulas into named cells of an existing
// workbook, keeping every cell's style; a date written into a plain cell
// gains a date number format. The workbook must exist: filling a template
// needs a template, and xlsx.write creates workbooks.
func WriteCells(ctx context.Context, path string, opts WriteCellsOptions) (*WriteResult, error) {
	return withLock(ctx, path, opts.Lock, func() (*WriteResult, error) {
		return writeCellsOnce(ctx, path, opts)
	})
}

func writeCellsOnce(ctx context.Context, path string, opts WriteCellsOptions) (*WriteResult, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	result := &WriteResult{Path: path, DryRun: opts.DryRun, Warnings: []string{}}
	warn := func(msg string) { result.Warnings = append(result.Warnings, msg) }
	warning, err := checkLockFile(path)
	if err != nil {
		return nil, err
	}
	if warning != "" {
		warn(warning)
	}
	if opts.Output != "" {
		if err := CheckExtension(opts.Output); err != nil {
			return nil, fmt.Errorf("output: %w", err)
		}
		if _, statErr := os.Stat(opts.Output); statErr == nil {
			warning, err := checkLockFile(opts.Output)
			if err != nil {
				return nil, err
			}
			if warning != "" {
				warn(warning)
			}
		}
		result.Path = opts.Output
	}
	w, err := open(path, opts.Password)
	if err != nil {
		return nil, err
	}
	defer w.close()

	sheet, err := w.resolveSheet(opts.Sheet)
	if err != nil {
		return nil, err
	}
	result.Sheet = sheet
	result.Changes.Sheet = sheet

	addresses := make([]string, 0, len(opts.Cells))
	for addr := range opts.Cells {
		addresses = append(addresses, addr)
	}
	sort.Strings(addresses)

	var box *region
	changedRows := map[string]map[int]bool{}
	for _, addr := range addresses {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		reg, err := w.parseRange(sheet, addr)
		if err != nil {
			return nil, err
		}
		if reg.C1 != reg.C2 || reg.R1 != reg.R2 {
			return nil, fmt.Errorf("%s: %q is not a single cell", w.base, addr)
		}
		changed, err := w.writeCell(reg.Sheet, reg.C1, reg.R1, opts.Cells[addr])
		if err != nil {
			return nil, err
		}
		if !changed {
			continue
		}
		w.forget(reg.Sheet)
		if changedRows[reg.Sheet] == nil {
			changedRows[reg.Sheet] = map[int]bool{}
		}
		changedRows[reg.Sheet][reg.R1] = true
		result.Changes.CellsChanged++
		if reg.Sheet == sheet {
			if box == nil {
				cell := reg
				box = &cell
			} else {
				box.C1, box.R1 = min(box.C1, reg.C1), min(box.R1, reg.R1)
				box.C2, box.R2 = max(box.C2, reg.C2), max(box.R2, reg.R2)
			}
		}
	}
	for _, rows := range changedRows {
		result.Changes.RowsUpdated += len(rows)
	}
	if box != nil {
		result.Changes.Range = box.String()
	}
	if opts.DryRun {
		return result, nil
	}
	if opts.Output != "" {
		w.path, w.base = opts.Output, Base(opts.Output)
	}
	if err := w.save(opts.InPlace); err != nil {
		return nil, err
	}
	return result, nil
}

// writeCell gives one cell its value, formula, or nothing, and reports
// whether the cell changed. A value equal to what the cell holds, or a
// formula equal to the cell's, leaves the cell alone.
func (w *file) writeCell(sheet string, col, row int, cv CellValue) (bool, error) {
	cell := cellName(col, row)
	existingFormula, err := w.f.GetCellFormula(sheet, cell)
	if err != nil {
		return false, w.cellError(sheet, col, row, err.Error())
	}
	if cv.Formula != "" {
		formula := strings.TrimPrefix(cv.Formula, "=")
		if existingFormula == formula {
			return false, nil
		}
		if err := w.f.SetCellFormula(sheet, cell, formula); err != nil {
			return false, w.cellError(sheet, col, row, err.Error())
		}
		return true, nil
	}
	grid, err := w.grid(sheet)
	if err != nil {
		return false, err
	}
	existing, err := w.cellValue(sheet, col, row, cellAt(grid, col, row), ReadOptions{}, func(string) {})
	if err != nil {
		return false, err
	}
	out, err := outValue(cv.Value, cv.Type, w.date1904)
	if err != nil {
		return false, w.cellError(sheet, col, row, err.Error())
	}
	if cv.Clear || out == nil {
		if existing == nil && existingFormula == "" {
			return false, nil
		}
		// The style stays: an emptied template cell keeps its look.
		if err := w.f.SetCellFormula(sheet, cell, ""); err != nil {
			return false, w.cellError(sheet, col, row, err.Error())
		}
		if err := w.f.SetCellDefault(sheet, cell, ""); err != nil {
			return false, w.cellError(sheet, col, row, err.Error())
		}
		return true, nil
	}
	if existingFormula == "" && sameValue(existing, out) {
		return false, nil
	}
	if err := w.setCell(sheet, col, row, out); err != nil {
		return false, err
	}
	w.styleWrittenCell(sheet, col, row, w.styleAt(sheet, col, row), out, cv.Type)
	return true, nil
}
