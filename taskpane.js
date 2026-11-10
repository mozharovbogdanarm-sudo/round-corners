/* Скругление углов — панель для PowerPoint (Office.js).
 *
 * Модель: pct = 0..100, где 100 — максимальное скругление (радиус = половина меньшей стороны).
 * У фигур «Прямоугольник: скругленные углы» значение adj = радиус / меньшая_сторона (0..0.5),
 * поэтому adj = pct / 200. В режиме «пт» пользователь вводит радиус в пунктах.
 */

// PowerPoint в этой сборке не отдаёт geometricShapeType, поэтому скругляемые фигуры определяем по имени
// ("Прямоугольник: скругленные углы 3", "Rectangle: Rounded Corners 3", "Rounded Rectangle 3" и т.п.).
function isRoundShape(s) { return /скругл|round/i.test(String((s && s.name) || "")); }

const el = (id) => document.getElementById(id);
const slider = el("slider");
const number = el("number");
const statusEl = el("status");

let unit = "pct";          // "pct" | "pt"
let minSidePt = 100;       // меньшая сторона первой выбранной фигуры, для режима «пт»
let pct = 0;               // текущее значение 0..100
let busy = Promise.resolve();
let debounceTimer = null;
let lastUserAction = 0;

function setStatus(text, isError) {
  statusEl.textContent = text;
  statusEl.className = isError ? "err" : "";
}

/* ---------- UI ---------- */

function radiusPtFromPct(p) { return (p / 100) * (minSidePt / 2); }
function pctFromRadiusPt(r) { return (r / (minSidePt / 2)) * 100; }

function refreshControls() {
  if (unit === "pct") {
    slider.max = 100; slider.step = 1;
    number.max = 100; number.step = 1;
    slider.value = pct;
    number.value = Math.round(pct * 10) / 10;
    el("numUnit").textContent = "%";
    el("valueLabel").textContent = "Скругление, % (100% — максимально)";
  } else {
    const maxPt = Math.max(0.1, minSidePt / 2);
    slider.max = maxPt; slider.step = 0.1;
    number.max = maxPt; number.step = 0.1;
    const r = radiusPtFromPct(pct);
    slider.value = r;
    number.value = Math.round(r * 10) / 10;
    el("numUnit").textContent = "пт";
    el("valueLabel").textContent = "Радиус скругления, пт (максимум " + (Math.round(maxPt * 10) / 10) + ")";
  }
  el("unitPct").classList.toggle("on", unit === "pct");
  el("unitPt").classList.toggle("on", unit === "pt");
  // Превью: 50% высоты демо-прямоугольника (56px) = 28px при pct=100
  el("demo").style.borderRadius = (pct / 100) * 28 + "px";
}

function userSetValue(raw) {
  let v = parseFloat(String(raw).replace(",", "."));
  if (!isFinite(v)) return;
  if (unit === "pct") {
    pct = Math.min(100, Math.max(0, v));
  } else {
    const maxPt = minSidePt / 2;
    pct = Math.min(100, Math.max(0, pctFromRadiusPt(Math.min(maxPt, Math.max(0, v)))));
  }
  lastUserAction = Date.now();
  refreshControls();
  scheduleApply();
}

slider.addEventListener("input", () => userSetValue(slider.value));
number.addEventListener("input", () => {
  if (number.value === "" || number.value === "-") return;
  userSetValue(number.value);
});
number.addEventListener("change", () => { userSetValue(number.value); refreshControls(); });
number.addEventListener("keydown", (e) => { if (e.key === "Enter") { userSetValue(number.value); refreshControls(); } });

document.querySelectorAll(".presets button").forEach((b) => {
  b.addEventListener("click", () => {
    const p = parseFloat(b.dataset.pct);
    pct = p; lastUserAction = Date.now();
    refreshControls(); scheduleApply();
  });
});

el("unitPct").addEventListener("click", () => { unit = "pct"; refreshControls(); });
el("unitPt").addEventListener("click", () => { unit = "pt"; refreshControls(); });

/* ---------- Применение к фигурам ---------- */

function scheduleApply() {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => {
    busy = busy.then(applyToSelection).catch(() => {});
  }, 120);
}

async function applyToSelection() {
  try {
    await PowerPoint.run(async (ctx) => {
      const shapes = ctx.presentation.getSelectedShapes();
      shapes.load("items/type,items/width,items/height,items/name");
      await ctx.sync();

      if (shapes.items.length === 0) { setStatus("Выделите фигуру.", true); return; }
      const geo = shapes.items.filter((s) => s.type === "GeometricShape");
      if (geo.length === 0) { setStatus("Среди выделенного нет автофигур (картинки, таблицы и т.п. не поддерживаются).", true); return; }

      const targets = geo.filter(isRoundShape);
      if (targets.length === 0) {
        setStatus("У фигуры «" + geo[0].name + "» нет скругляемых углов. Замените её: Формат фигуры → Изменить фигуру → «Прямоугольник: скругленные углы».", true);
        return;
      }

      const expected = [];
      targets.forEach((s) => {
        const minSide = Math.max(0.01, Math.min(s.width, s.height));
        let adj;
        if (unit === "pt") {
          adj = Math.min(0.5, radiusPtFromPct(pct) / minSide);
        } else {
          adj = pct / 200;
        }
        expected.push(adj);
        s.adjustments.set(0, adj);
      });
      await ctx.sync();

      const skipped = shapes.items.length - targets.length;
      let report = "Применено к фигурам: " + targets.length + (skipped > 0 ? " (пропущено: " + skipped + ")" : "") + ".";

      // Диагностика: перечитываем значение после записи.
      try {
        const back = targets.map((s) => s.adjustments.get(0));
        await ctx.sync();
        const got = back.map((r) => {
          const v = r && typeof r === "object" && "value" in r ? r.value : r;
          return typeof v === "number" ? v.toFixed(3) : String(v);
        });
        report += " Записано: " + expected.map((v) => v.toFixed(3)).join(", ") +
          ", прочитано: " + got.join(", ") + ".";
      } catch (e) {
        report += " (не удалось перечитать значение: " + (e && e.message ? e.message : e) + ")";
      }
      setStatus(report);
    });
  } catch (e) {
    setStatus("Не удалось изменить скругление: " + (e && e.message ? e.message : e) +
      ". Нужна актуальная версия PowerPoint с поддержкой adjustments.", true);
  }
}

/* ---------- Чтение значения из выделения ---------- */

async function readSelection() {
  try {
    await PowerPoint.run(async (ctx) => {
      const shapes = ctx.presentation.getSelectedShapes();
      shapes.load("items/type,items/width,items/height,items/name");
      await ctx.sync();
      if (shapes.items.length === 0) { setStatus("Выделите фигуру."); return; }

      const geo = shapes.items.filter((s) => s.type === "GeometricShape");
      if (geo.length === 0) { setStatus("Среди выделенного нет автофигур.", true); return; }

      const first = geo[0];
      minSidePt = Math.max(1, Math.min(first.width, first.height));

      if (isRoundShape(first)) {
        const res = first.adjustments.get(0);
        await ctx.sync();
        const v = res && typeof res === "object" && "value" in res ? res.value : res;
        if (typeof v === "number" && isFinite(v)) {
          pct = Math.min(100, Math.max(0, v * 200));
        }
        setStatus("Выделено фигур: " + shapes.items.length + ".");
      } else {
        pct = 0;
        setStatus("Фигура «" + first.name + "» не имеет скругляемых углов. Замените её на «Прямоугольник: скругленные углы».", true);
      }
      refreshControls();
    });
  } catch (e) {
    /* чтение значения — удобство, ошибки здесь не критичны */
  }
}

function onSelectionChanged() {
  // Не перебиваем пользователя, пока он двигает ползунок.
  if (Date.now() - lastUserAction < 700) return;
  busy = busy.then(readSelection).catch(() => {});
}

Office.onReady((info) => {
  refreshControls();
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("Эта надстройка работает только в PowerPoint.", true);
    return;
  }
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
  readSelection();
});
