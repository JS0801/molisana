/**
 * @NApiVersion 2.1
 * @NScriptType ClientScript
 */
define([], function () {

  // ==== Column index constants (TD indices in each <tr>) ====
  const MONTH_CELL_INDEX = 2;       // “Month of Stock” cell
  const CUBIC_CELL_INDEX = 3;       // “Item Cubic Space” cell (computed)
  const WEIGHT_CELL_INDEX = 4;      // “Weight” cell (computed/display)

  // Source columns in your table (NOT CSV indices)
  const SRC_PER_CUBIC_IDX = 35;     // per-unit cubic (m³ per unit)
  const SRC_PER_WGT_IDX = 33;       // per-unit weight (value)
  const SRC_PER_WGT_UNIT_IDX = 34;  // unit string: 'g', 'gram', 'lb', etc.
  const SRC_PER_AVAIL_IDX = 36;
  const SRC_MONTH_AVG_IDX_1 = 16;   // preferred month avg

  function pageInit(context) {
    const checkboxes = document.querySelectorAll('input[type="checkbox"][name^="row_select_"]');

    checkboxes.forEach(function (checkbox) {
      const rowId = checkbox.name.split('_')[2];
      const qtyInput = document.querySelector(`input[name="qty_input_${rowId}"]`);
      const row = checkbox.closest('tr');
      if (!qtyInput || !row) return;

      // Checkbox change
      checkbox.addEventListener('change', function () {
        if (checkbox.checked) {
          updateRowCells(row, qtyInput);
        } else {
          clearRowCells(row);
        }
        updateTotals();
      });

      // Qty change (only matters if selected)
      qtyInput.addEventListener('input', function () {
        if (checkbox.checked) {
          updateRowCells(row, qtyInput);
          updateTotals();
        }
      });
    });

    // Initial totals (in case anything starts prechecked)
    updateTotals();
  }

  // --- per-row calculation and highlighting ---
  function updateRowCells(row, qtyInput) {
    const qtyOrdered = toNum(qtyInput.value);
    const cells = row.querySelectorAll('td');
    if (cells.length <= WEIGHT_CELL_INDEX) return;

    // Month of stock
    const last3rdValue = qtyOrdered;
    const monthAvg = toNum(cells[SRC_MONTH_AVG_IDX_1]?.textContent);
    console.log(cells)
    
    const monthQty = parseFloat(cells[16]?.textContent) || 0;
    const inTransit = parseFloat(cells[25]?.textContent) || 0;
    const onOrder = parseFloat(cells[30]?.textContent) || 0;
    const avail = parseFloat(cells[SRC_PER_AVAIL_IDX]?.textContent) || 0;

    console.log({
      monthAvg, monthQty, inTransit, onOrder, avail, qtyOrdered
    })
    
    
    const monthCell = cells[MONTH_CELL_INDEX];
    const total = parseFloat(inTransit) + parseFloat(onOrder) + parseFloat(avail) + parseFloat(qtyOrdered)
    console.log(total)
    monthCell.textContent = (total/monthQty).toFixed(2);
    styleCalcCell(monthCell);

    // Cubic (m³) for this row = qty * per-unit cubic
    const perCubic = toNum(cells[SRC_PER_CUBIC_IDX]?.textContent);
    const rowCubic = qtyOrdered * (isNaN(perCubic) ? 0 : perCubic);
    const cubicCell = cells[CUBIC_CELL_INDEX];
    cubicCell.textContent = formatNumber(rowCubic, 2);
    styleCalcCell(cubicCell);

    // Weight for this row (display nice, but totals will recompute from source)
    const perWgtVal = toNum(cells[SRC_PER_WGT_IDX]?.textContent);
    const perWgtUnit = (cells[SRC_PER_WGT_UNIT_IDX]?.textContent || '').toLowerCase().trim();
    const perWgtKg = perUnitWeightToKg(perWgtVal, perWgtUnit);
    const rowWgtKg = qtyOrdered * perWgtKg;

    const weightCell = cells[WEIGHT_CELL_INDEX];
    weightCell.textContent = formatNumber(rowWgtKg, 2) + ' kg';
    styleCalcCell(weightCell);
  }

  function clearRowCells(row) {
    const cells = row.querySelectorAll('td');
    [MONTH_CELL_INDEX, CUBIC_CELL_INDEX, WEIGHT_CELL_INDEX].forEach(idx => {
      const c = cells[idx];
      if (!c) return;
      c.textContent = '';
      c.style.backgroundColor = '';
      c.style.fontWeight = '';
    });
  }

  // --- totals across all checked rows ---
  function updateTotals() {
    let totalCubic = 0; // m³
    let totalWeightKg = 0;

    document.querySelectorAll('input[type="checkbox"][name^="row_select_"]').forEach(cb => {
      if (!cb.checked) return;

      const row = cb.closest('tr');
      const rowId = cb.name.split('_')[2];
      const qtyInput = document.querySelector(`input[name="qty_input_${rowId}"]`);
      if (!row || !qtyInput) return;

      const cells = row.querySelectorAll('td');
      const qty = toNum(qtyInput.value);

      // Recompute from sources to avoid parsing display text
      const perCubic = toNum(cells[SRC_PER_CUBIC_IDX]?.textContent);
      const cubic = qty * (isNaN(perCubic) ? 0 : perCubic);
      totalCubic += cubic;

      const perWgtVal = toNum(cells[SRC_PER_WGT_IDX]?.textContent);
      const perWgtUnit = (cells[SRC_PER_WGT_UNIT_IDX]?.textContent || '').toLowerCase().trim();
      const perWgtKg = perUnitWeightToKg(perWgtVal, perWgtUnit);
      totalWeightKg += qty * perWgtKg;
    });

    // Push to the pills (if present)
    const cubicEl = document.getElementById('totalCubicValue');
    if (cubicEl) cubicEl.textContent = formatNumber(totalCubic, 2);

    const wgtEl = document.getElementById('totalWeightValue');
    if (wgtEl) wgtEl.textContent = formatNumber(totalWeightKg, 2);
  }

  // --- helpers ---
  function perUnitWeightToKg(value, unit) {
    // Normalize to kg per unit
    if (!value || isNaN(value)) return 0;
    if (unit === 'g' || unit === 'gram' || unit === 'grams') return value / 1000;      // g -> kg
    if (unit === 'kg' || unit === 'kilogram' || unit === 'kilograms') return value;    // kg -> kg
    // default treat as pounds (lb)
    return value * 0.453592; // lb -> kg
  }

  function toNum(v) {
    const n = parseFloat((v || '').toString().replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }

  function firstNonNaN() {
    for (let i = 0; i < arguments.length; i++) {
      const v = arguments[i];
      if (!isNaN(v) && v !== 0) return v;
    }
    return NaN;
  }

  function styleCalcCell(el) {
    el.style.backgroundColor = '#e6f7ff';
    el.style.fontWeight = 'bold';
  }

  function formatNumber(n, maxFrac) {
    const num = typeof n === 'number' ? n : toNum(n);
    return num.toLocaleString(undefined, { maximumFractionDigits: maxFrac ?? 2 });
  }

  return {
    pageInit: pageInit
  };
});
