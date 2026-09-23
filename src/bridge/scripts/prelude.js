// Shared helpers for every JXA script. The Rust side prepends `const ARGS = {...};`
// and appends one script body that ends in a `respond(() => ...)` call.

class GridskiError extends Error {
  constructor(code, message) {
    super(message);
    this.gridskiCode = code;
  }
}

// Wrap a script body so osascript always prints exactly one JSON envelope.
function respond(fn) {
  try {
    return JSON.stringify({ ok: fn() });
  } catch (e) {
    return JSON.stringify({
      error: {
        code: e.gridskiCode || null,
        number: typeof e.errorNumber === "number" ? e.errorNumber : null,
        message: String((e && e.message) || e),
      },
    });
  }
}

// Never launch Excel as a side effect: only talk to it if it is already running.
function excelApp() {
  const app = Application("Microsoft Excel");
  if (!app.running()) {
    throw new GridskiError("NOT_RUNNING", "Microsoft Excel is not running.");
  }
  return app;
}

function exists(ref) {
  try {
    ref.name();
    return true;
  } catch (e) {
    return false;
  }
}

// Resolve { workbook?, sheet? } to live references, defaulting to whatever is active.
function resolveSheet(excel, args) {
  let wb;
  if (args.workbook) {
    wb = excel.workbooks[args.workbook];
    if (!exists(wb)) throw new GridskiError("NOT_FOUND", `No open workbook named "${args.workbook}".`);
  } else {
    wb = excel.activeWorkbook;
    if (!exists(wb)) throw new GridskiError("NOT_FOUND", "No workbook is active in Excel.");
  }
  let sh;
  if (args.sheet) {
    sh = wb.worksheets[args.sheet];
    if (!exists(sh)) throw new GridskiError("NOT_FOUND", `No sheet named "${args.sheet}" in "${wb.name()}".`);
  } else {
    sh = wb.activeSheet;
  }
  return { wb, sh };
}

function resolveRange(sh, address) {
  const r = sh.ranges[address];
  let resolved;
  try {
    resolved = r.getAddress();
  } catch (e) {
    throw new GridskiError("NOT_FOUND", `"${address}" is not a valid range on sheet "${sh.name()}".`);
  }
  if (resolved.includes(",")) {
    throw new GridskiError("UNSUPPORTED", "Multi-area ranges are not supported; read each area separately.");
  }
  return r;
}

function pad(n) {
  return String(n).padStart(2, "0");
}

// Excel dates arrive as JS Dates at local midnight; format them in local time so
// "2026-01-05" doesn't turn into "2026-01-05T08:00:00Z".
function formatDate(d) {
  const date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  if (d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0) return date;
  return `${date} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

// Excel returns a bare scalar for a single cell; always hand Rust a 2-D array.
function grid(v) {
  const rows = Array.isArray(v) ? v : [[v]];
  return rows.map((row) =>
    (Array.isArray(row) ? row : [row]).map((c) => (c instanceof Date ? formatDate(c) : c === undefined ? null : c))
  );
}

// Read a single-area range, truncating to at most maxCells by dropping trailing rows/columns.
function readRange(wb, sh, range, maxCells, includeFormulas) {
  const totalRows = range.rows.length;
  const totalCols = range.columns.length;
  const cols = Math.min(totalCols, maxCells);
  const rows = Math.min(totalRows, Math.max(1, Math.floor(maxCells / cols)));
  const r = rows === totalRows && cols === totalCols ? range : range.getResize({ rowSize: rows, columnSize: cols });
  return {
    workbook: wb.name(),
    sheet: sh.name(),
    address: r.getAddress(),
    total_rows: totalRows,
    total_cols: totalCols,
    values: grid(r.value()),
    text: grid(r.stringValue()),
    formulas: includeFormulas ? grid(r.formula()) : null,
  };
}

// Quote a string as an AppleScript literal.
function asq(s) {
  return '"' + String(s).replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"';
}

// Some Excel commands (positioning sheets, creating names, save-as) only work from
// AppleScript, so run those snippets in-process instead of through the JXA bridge.
function applescript(source) {
  ObjC.import("Foundation");
  const err = $();
  const result = $.NSAppleScript.alloc.initWithSource(source).executeAndReturnError(err);
  if (result.isNil()) {
    const info = ObjC.deepUnwrap(err) || {};
    const e = new Error(info.NSAppleScriptErrorBriefMessage || info.NSAppleScriptErrorMessage || "AppleScript failed");
    e.errorNumber = info.NSAppleScriptErrorNumber;
    throw e;
  }
  return ObjC.unwrap(result.stringValue);
}

// Run AppleScript with Excel's confirmation dialogs suppressed, so a delete or
// overwrite can't block on a modal nobody will answer.
function excelScript(body) {
  return applescript(
    `tell application "Microsoft Excel"\nset display alerts to false\ntry\n${body}\non error m number n\n` +
      `set display alerts to true\nerror m number n\nend try\nset display alerts to true\nend tell`
  );
}

function fileMtime(path) {
  ObjC.import("Foundation");
  const attrs = $.NSFileManager.defaultManager.attributesOfItemAtPathError(path, null);
  return attrs.isNil() ? null : attrs.fileModificationDate.timeIntervalSince1970;
}

function workbookInfo(excel, wb) {
  const active = exists(excel.activeWorkbook) ? excel.activeWorkbook.name() : null;
  const activeSheet = wb.activeSheet.name();
  return {
    name: wb.name(),
    path: wb.fullName(),
    saved: wb.saved(),
    active: wb.name() === active,
    sheets: wb.worksheets().map((sh) => ({
      name: sh.name(),
      used_range: sh.usedRange.getAddress().replace(/\$/g, ""),
      active: sh.name() === activeSheet,
    })),
  };
}

// Excel can't hold two workbooks with the same file name (compared without case), even
// from different folders; opening or saving as one fails with a bare "Parameter error".
// `self` is the full name of the workbook being saved, which may keep its own name.
function refuseNameClash(excel, path, self) {
  const name = path.split("/").pop().toLowerCase();
  const clash = excel.workbooks().find((wb) => wb.name().toLowerCase() === name && wb.fullName() !== self);
  if (clash) {
    throw new GridskiError(
      "UNSUPPORTED",
      `A workbook named ${clash.name()} is already open (${clash.fullName()}). Close it first — ` +
        "Excel can't hold two workbooks with the same name."
    );
  }
}

function resolveWorkbook(excel, name) {
  return resolveSheet(excel, { workbook: name }).wb;
}

function hexColor(rgb) {
  return "#" + rgb.map((c) => Math.max(0, Math.min(255, c)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function rgbFromHex(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex));
  if (!m) throw new GridskiError("INVALID", `"${hex}" is not a color; use "#RRGGBB".`);
  const n = parseInt(m[1], 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function colLetters(n) {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

// "$B$2:$D$3" → [2, 2, 3, 4] as [top, left, bottom, right].
function bounds(address) {
  const parts = address.replace(/\$/g, "").split("!").pop().split(":");
  const cell = (a) => {
    const m = /^([A-Z]+)(\d+)$/.exec(a);
    const col = m[1].split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    return [Number(m[2]), col];
  };
  const [t, l] = cell(parts[0]);
  const [b, r] = cell(parts[1] || parts[0]);
  return [t, l, b, r];
}

function a1(t, l, b, r) {
  const tl = colLetters(l) + t;
  return t === b && l === r ? tl : `${tl}:${colLetters(r)}${b}`;
}
