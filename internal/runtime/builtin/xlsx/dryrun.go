// Copyright (C) 2026 Yota Hamada
// SPDX-License-Identifier: GPL-3.0-or-later

package xlsx

import (
	"context"
	"errors"
	"fmt"
	"io/fs"
	"os"
	"sort"
	"strings"

	"github.com/dagucloud/dagu/v2/internal/cmn/workbook"
	"github.com/dagucloud/dagu/v2/internal/ir"
	"github.com/dagucloud/dagu/v2/internal/runtime"
)

// dryRunCheck is the xlsx executor's dry-run check: it reports a workbook
// that does not exist, a sheet that is not in it, or header columns a step
// names that are not in the header row. A with field whose value is still
// a reference, such as a step output, is skipped; so is a problem the dry
// run cannot judge, such as a locked workbook. Validation errors are not
// repeated here, since dagu dry reports them before any step is checked.
func dryRunCheck(ctx context.Context, step ir.Step) error {
	cfg, op, err := loadConfig(step, true)
	if err != nil || cfg.deferred["path"] {
		return nil
	}
	workDir := runtime.GetEnv(ctx).WorkingDir
	path, err := resolvePath(workDir, cfg.Path)
	if err != nil {
		return nil
	}
	switch op {
	case opWrite, opAppend:
		// The workbook may be created; only the input file can be checked.
		if !cfg.provided("input") {
			return nil
		}
		input, err := resolvePath(workDir, cfg.Input)
		if err != nil {
			return nil
		}
		if _, err := os.Stat(input); errors.Is(err, fs.ErrNotExist) {
			return fmt.Errorf("field 'with.input': %s: file not found", workbook.Base(input))
		}
		return nil
	case opInfo, opListSheets:
		return workbook.Check(ctx, path, workbook.CheckOptions{Password: cfg.Password})
	case opSheet:
		operation := strings.ToLower(strings.TrimSpace(cfg.Operation))
		if operation == string(workbook.SheetAdd) || cfg.Missing == string(workbook.MissingSkip) || cfg.deferred["sheet"] {
			return workbook.Check(ctx, path, workbook.CheckOptions{Password: cfg.Password})
		}
		return workbook.Check(ctx, path, workbook.CheckOptions{Password: cfg.Password, Sheet: cfg.Sheet})
	case opWriteCells:
		if cfg.deferred["sheet"] {
			return workbook.Check(ctx, path, workbook.CheckOptions{Password: cfg.Password})
		}
		return workbook.Check(ctx, path, workbook.CheckOptions{Password: cfg.Password, Sheet: cfg.Sheet})
	}
	opts := workbook.CheckOptions{Password: cfg.Password, Range: cfg.Range, Header: cfg.header}
	if !cfg.deferred["sheet"] {
		opts.Sheet = cfg.Sheet
	}
	add := func(field string, names ...string) {
		// update_rows resolves its key and set columns exactly, as the run
		// does; the reading operations accept a loose match.
		exact := op == opUpdateRows
		for _, name := range names {
			opts.Columns = append(opts.Columns, workbook.ColumnCheck{Field: field, Name: name, Exact: exact})
		}
	}
	switch op {
	case opRead, opValidate, opConvert:
		for _, sel := range cfg.columns {
			add("columns", sel.Source)
		}
		add("types", sortedNames(cfg.types)...)
		add("where", sortedNames(cfg.Where)...)
		add("required", cfg.Required...)
		add("not_blank", cfg.NotBlank...)
		add("unique", cfg.Unique...)
		add("allowed", sortedNames(cfg.allowed)...)
	case opUpdateRows:
		if key := strings.TrimSpace(cfg.Key); key != "" && key != workbook.RowNumberKey {
			add("key", key)
		}
		add("set", sortedNames(cfg.set)...)
	}
	return workbook.Check(ctx, path, opts)
}

// sortedNames returns a map's keys in order, so warnings read the same
// from run to run.
func sortedNames[V any](m map[string]V) []string {
	names := make([]string, 0, len(m))
	for name := range m {
		names = append(names, name)
	}
	sort.Strings(names)
	return names
}
