// Excel answers a format query on a multi-cell range only when every cell agrees
// (otherwise null, or an out-of-range color index for fills). So ask about a whole
// block first and split it in half only where cells differ: models format by row
// and block, so this takes a handful of queries per row instead of several per cell.
respond(() => {
  const excel = excelApp();
  const { wb, sh } = resolveSheet(excel, ARGS);
  const range = ARGS.range ? resolveRange(sh, ARGS.range) : sh.usedRange;
  const deadline = Date.now() + ARGS.budget_ms;

  const uniform = (t, l, b, r) => {
    const rg = sh.ranges[a1(t, l, b, r)];
    const font = rg.fontObject;
    if (font.fontColorIndex() === null) return null;
    const style = font.fontStyle();
    if (style === null) return null;
    const numberFormat = rg.numberFormat();
    if (numberFormat === null) return null;
    const fillIndex = rg.interiorObject.colorIndex();
    const noFill = fillIndex === "color index none";
    if (!noFill && !(typeof fillIndex === "number" && fillIndex >= 1 && fillIndex <= 56)) return null;
    return {
      font_color: hexColor(font.color()),
      bold: /bold/i.test(style),
      italic: /italic|oblique/i.test(style),
      fill: noFill ? null : hexColor(rg.interiorObject.color()),
      number_format: numberFormat,
    };
  };

  const blocks = [];
  const unread = [];
  const stack = [bounds(range.getAddress())];
  while (stack.length) {
    const [t, l, b, r] = stack.pop();
    if (Date.now() > deadline) {
      unread.push(a1(t, l, b, r));
      continue;
    }
    const fmt = uniform(t, l, b, r);
    if (fmt) {
      blocks.push({ top: t, left: l, bottom: b, right: r, ...fmt });
    } else if (b > t) {
      const mid = Math.floor((t + b) / 2);
      stack.push([mid + 1, l, b, r], [t, l, mid, r]);
    } else if (r > l) {
      const mid = Math.floor((l + r) / 2);
      stack.push([t, mid + 1, b, r], [t, l, t, mid]);
    } else {
      // A single cell whose format still reads as mixed (e.g. rich text); report what we can.
      blocks.push({ top: t, left: l, bottom: b, right: r, font_color: null, bold: false, italic: false, fill: null, number_format: sh.ranges[a1(t, l, t, l)].numberFormat() });
    }
  }
  return { workbook: wb.name(), sheet: sh.name(), address: range.getAddress(), blocks, unread };
});
