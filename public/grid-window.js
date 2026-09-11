(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PictosGrid = api;
})(globalThis, function () {
  function layout({ count, columns, width, top, viewport, gap = 14, overscan = 2 }) {
    columns = Math.min(6, Math.max(1, Math.round(columns) || 3));
    const size = Math.max(1, (width - gap * (columns - 1)) / columns);
    const stride = size + gap, rows = Math.ceil(count / columns);
    const firstRow = Math.min(Math.max(0, rows - 1), Math.max(0, Math.floor(top / stride) - overscan));
    const lastRow = Math.min(rows, Math.max(firstRow + 1, Math.ceil((top + viewport) / stride) + overscan));
    return { columns, size, stride, height: Math.max(0, rows * stride - gap), start: firstRow * columns, end: Math.min(count, lastRow * columns) };
  }
  return { layout };
});
