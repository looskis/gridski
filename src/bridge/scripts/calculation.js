// Read or change Excel's calculation settings. Iteration is what lets a model with a
// deliberate circular reference (interest on average balances) resolve instead of
// raising a circular-reference warning.
const MODES = { automatic: "calculation automatic", manual: "calculation manual", semiautomatic: "calculation semiautomatic" };

respond(() => {
  const excel = excelApp();
  const applied = [];
  const has = (k) => ARGS[k] !== null && ARGS[k] !== undefined;

  if (has("mode")) {
    if (!MODES[ARGS.mode]) throw new GridskiError("INVALID", "mode is automatic, manual, or semiautomatic.");
    excel.calculation = MODES[ARGS.mode];
    applied.push("mode");
  }
  if (has("iterative")) (excel.iteration = ARGS.iterative), applied.push("iterative");
  if (has("max_iterations")) (excel.maxIterations = ARGS.max_iterations), applied.push("max_iterations");
  if (has("max_change")) (excel.maxChange = ARGS.max_change), applied.push("max_change");
  if (ARGS.recalculate) {
    excelScript("calculate full");
    applied.push("recalculate");
  }

  const mode = String(excel.calculation()).replace("calculation ", "");
  return {
    mode,
    iterative: excel.iteration(),
    max_iterations: excel.maxIterations(),
    max_change: excel.maxChange(),
    applied,
  };
});
