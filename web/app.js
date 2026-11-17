"use strict";

const content = document.querySelector("#content");
const modal = document.querySelector("#modal");
const form = document.querySelector("#form");
const fields = document.querySelector("#fields");
const formError = document.querySelector("#form-error");
const status = document.querySelector("#status");
const numberFormat = new Intl.NumberFormat("tr-TR", { maximumFractionDigits: 3 });
const unitPriceFormat = new Intl.NumberFormat("tr-TR", {maximumFractionDigits:4});
const currency = new Intl.NumberFormat("tr-TR", { style: "currency", currency: "TRY" });
const definitions = {
  receipt: "Alış girişi", person: "Personele teslim", department: "Departmana teslim",
  return: "İade", adjustment: "Sayım düzeltmesi"
};
const scopes = { person: "Personel", department: "Departman", both: "Her ikisi" };
const modules = { all: "Tüm malzemeler", sarf: "Sarf malzemeleri", kkd: "İş kıyafeti ve KKD" };
const sizeKinds = {size: "Beden", shoe: "Ayakkabı numarası"};
const calculationLabels = {sum: "Toplam", average: "Ortalama", min: "En düşük (Min)", max: "En yüksek (Max)", none: "Hesap yok"};
const calculationColumnsByPage = {
  stok: {"Eldeki miktar": "quantity", "Stok değeri": "money"},
  hareketler: {"Miktar": "quantity", "Stok değeri farkı": "money"}
};
const calculationsByPage = Object.fromEntries(
  Object.keys(calculationColumnsByPage).map(key => [key, {quantity: "sum", money: "sum"}])
);
const pages = {
  stok: ["Stok durumu", "Satın Alma deposundaki güncel ürün ve beden bakiyeleri."],
  hareketler: ["Stok hareketleri", "Alış, teslim, iade ve sayım kayıtları."],
  sayimlar: ["Sayım", "Fiziksel stok sayımlarını hazırlayın, kilitleyin ve onaylayın."],
  urunler: ["Ürün kartları", "Ürün, stok birimi, beden ve teslim kapsamı tanımları."],
  tanimlar: ["Departman ve personel", "Teslim fişlerinde kullanılacak kişi ve departmanlar."],
  raporlar: ["Tüketim raporları", "Ürün ve departman bazında aylık çıkış maliyetleri."],
  excel: ["Excel içe aktar", "Tanımları ve stok hareketlerini tek dosyayla toplu kaydedin."],
  kullanicilar: ["Kullanıcılar ve yetkiler", "Windows hesapları, erişim kapsamı ve son işlemler."],
  "donem-kilidi": ["Dönem kilidi", "Kullanıcıya ve tarih aralığına göre dönem kilidi istisnalarını yönetin."]
};

function localDay() {
  return new Intl.DateTimeFormat("sv-SE", {timeZone: "Europe/Istanbul"}).format(new Date());
}

const today = localDay();
const filters = { module: "all", search: "", from: "", to: "", department: "", person: "", kind: "" };
const report = { from: today.slice(0, 7) + "-01", to: today, department: "", person: "" };
let data;
let activeForm;
let busy = false;
let importFile;
let importPreview;
let personnelDate = "";
let historyRequest = 0;
let countDetails;
const countQuantityDrafts = new Map();
let timelines = new Map();
let authRequired = null;
let session = null;
const isAdmin = () => session?.admin === true;
const demoAccess = () => authRequired === false;
const can = permission => demoAccess() || isAdmin() || Boolean(session?.permissions.includes(permission));
const mayEdit = permission => can(permission) && !session?.department_id;
const adminScreens = new Set(["kullanicilar", "donem-kilidi"]);
const canView = screen => adminScreens.has(screen) ? isAdmin() :
  screen === "excel" ? demoAccess() || isAdmin() :
  screen === "sayimlar" ? can("view.sayimlar") && !session?.department_id : can("view." + screen);
const countId = () => /^#sayimlar\/(\d+)$/.exec(location.hash)?.[1] ?? "";
const countAdmin = () => demoAccess() || isAdmin();
function dateIsOpen(day) {
  return countAdmin() || day >= data.open_period_start ||
    (data.open_periods || []).some(grant => day >= grant.start_day && day <= grant.end_day);
}
function periodLimit() {
  if (countAdmin() || !data?.open_period_start) return "";
  const earliest = (data.open_periods || []).reduce((day, grant) =>
    grant.start_day < day ? grant.start_day : day, data.open_period_start);
  return ` min="${earliest}"`;
}

function esc(value) {
  return String(value ?? "").replace(/[&<>"']/g, char => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[char]));
}

const money = value => currency.format(value / 100);
const qty = value => numberFormat.format(value / 1000);
const variantName = row => [row.color, row.size].filter(Boolean).join(" / ");
const productName = row => [row.name, variantName(row)].filter(Boolean).map(esc).join(" · ");
const page = () => {
  const requested = countId() ? "sayimlar" : location.hash.slice(1) || "stok";
  if (Object.hasOwn(pages, requested) && canView(requested)) return requested;
  return Object.keys(pages).find(canView) || "stok";
};
const options = (items, current = "", blank) =>
  (blank === undefined ? "" : '<option value="">' + esc(blank) + "</option>") +
  items.map(([value, label]) =>
    `<option value="${esc(value)}"${String(value) === String(current) ? " selected" : ""}>${esc(label)}</option>`
  ).join("");
function onDay(type, row, day, recordId = row.id) {
  if (!day) return row;
  return timelines.get(type + "/" + recordId)?.find(version => version.effective_day <= day) ?? row;
}

const departmentOptions = (activeOnly = false, day = "") => data.departments
  .map(row => [row, onDay("departments", row, day)]).filter(([, version]) => !activeOnly || version.active)
  .map(([row, version]) => [row.id, version.name]);
const personOptions = (activeOnly = false, day = "") => data.people.map(row => [row, onDay("people", row, day)])
  .filter(([row, version]) => {
    if (!activeOnly) return true;
    const dept = data.departments.find(item => item.id === version.department_id);
    return version.active && (dept ? onDay("departments", dept, day).active : row.department_active);
  }).map(([row, version]) => [row.id, version.name + " · " + row.code]);
const label = (title, control, css = "") => `<label class="${css}">${title}${control}</label>`;
const input = (name, value = "", more = "") => `<input name="${name}" value="${esc(value)}" ${more}>`;
const select = (name, items, current, blank, more = "") =>
  `<select name="${name}" ${more}>${options(items, current, blank)}</select>`;
const documentPermission = {receipt:"movement.receipt", person:"movement.delivery", department:"movement.delivery",
  return:"movement.return", adjustment:"movement.adjustment"};
const definitionPermission = {"products":"edit.products", "product-update":"edit.products", "variants":"edit.products",
  "product-groups":"edit.products", "product-group-size-update":"edit.products",
  "departments":"edit.departments", "department-update":"edit.departments",
  "people":"edit.people", "person-update":"edit.people", "person-status":"edit.people"};
function canAct(name) {
  if (["user-update", "period-add", "period-revoke"].includes(name)) return isAdmin();
  if (name === "count-create" || name === "count-line") return canView("sayimlar") && can("count.edit");
  if (name === "count-transition") return canView("sayimlar");
  if (name === "receipt" || name === "adjustment") {
    return can(documentPermission[name]) && can("field.cost") &&
      can(name === "receipt" ? "field.supplier" : "field.notes") && !session?.department_id;
  }
  if (documentPermission[name]) return can(documentPermission[name]);
  return mayEdit(definitionPermission[name]);
}

const countStatuses = {draft:"Taslak", locked:"Kilitli", rejected:"Onaylanmadı", approved:"Onaylandı"};
function countStatus(status) {
  return `<span class="pill${status === "locked" || status === "rejected" ? " warning" : ""}">${countStatuses[status]}</span>`;
}

function countWritable(count) {
  return count.status !== "approved" &&
    (countAdmin() || (count.status === "draft" && dateIsOpen(count.day)));
}

const countQuantityKey = (count, variantId) => count.id + "/" + variantId;
const countQuantityValue = line => line.counted_qty === null ? "" : String(line.counted_qty / 1000);

function hasCountQuantityDrafts(id = countId()) {
  return Boolean(id && [...countQuantityDrafts.keys()].some(key => key.startsWith(id + "/")));
}

function countQuantityStepper(value, caption, attributes = "") {
  return `<div class="count-stepper">
    <button type="button" data-count-step="-1" aria-label="Miktarı 1 azalt">−</button>
    ${input("counted_qty", value, `type="number" min="0" step=".001" data-count-step-field
      aria-label="${esc(caption)}" placeholder="Sayılmadı" ${attributes}`)}
    <button type="button" data-count-step="1" aria-label="Miktarı 1 artır">+</button>
  </div>`;
}

function stepCountQuantity(field, step) {
  if (step < 0 && !field.valueAsNumber) return;
  const thousandths = Number.isFinite(field.valueAsNumber) ? Math.round(field.valueAsNumber * 1000) : 0;
  field.value = Math.max(0, thousandths + step * 1000) / 1000;
  field.dispatchEvent(new Event("input", {bubbles: true}));
}

function countQuantityCell(count, line, editable) {
  const amount = line.counted_qty === null ? null : quantityCell(line.counted_qty, line.unit);
  if (!editable) return amount ?? "—";
  const key = countQuantityKey(count, line.variant_id);
  const value = countQuantityDrafts.get(key) ?? countQuantityValue(line);
  const details = [line.name, variantName(line)].filter(Boolean).join(" · ");
  const markup = `<span data-saved-quantity hidden>${amount?.text ?? "—"}</span>
    <div class="count-quantity-editor" data-cell-controls>
      ${countQuantityStepper(value, details + " sayılan miktar", `data-count-quantity="${line.variant_id}"`)}
      <button type="button" data-count-save="${line.variant_id}" aria-label="${esc(details)} miktarını kaydet">Kaydet</button>
    </div>`;
  return amount ? {...amount, text: markup} : markup;
}

function countPage() {
  if (!countId()) {
    const rows = (data.counts || []).map(count => [
      "#" + count.id, esc(count.day), countStatus(count.status), numberFormat.format(count.entered),
      numberFormat.format(count.total), esc(count.owner), esc(count.note || "—"),
      `<a class="table-link" href="#sayimlar/${count.id}">Sayımı aç</a>`
    ]);
    return panel("Sayım listesi", table(["Sayım no","Sayım tarihi","Durum","Sayılan satır","Toplam satır","Oluşturan",
      "Açıklama",actionColumn], rows, "Henüz sayım oluşturulmamış."), action("count-create","+ Yeni sayım",true));
  }
  if (!countDetails || String(countDetails.id) !== countId()) return "<p>Sayım yükleniyor…</p>";
  const count = countDetails;
  const writable = countWritable(count);
  const editable = writable && can("count.edit");
  const canDecide = writable && can("count.approve") && ["draft","locked"].includes(count.status);
  const transition = (op, title, primary = false) =>
    `<button type="button" data-count-transition="${op}" class="${primary ? "primary" : ""}">${title}</button>`;
  const dateLocked = !dateIsOpen(count.day);
  const actions = [
    editable && count.status === "draft" ? transition("lock","Sayımı kilitle") : "",
    countAdmin() && ["locked","rejected"].includes(count.status) ? transition("reopen","Taslağı yeniden aç") : "",
    canDecide ? transition("reject","Reddet") + transition("approve","Onayla ve stoğa işle",true) : ""
  ].join("");
  const entries = count.lines.map(line => [
    esc(line.code), esc(line.name), esc(line.category), esc(line.color),
    line.size_kind === "shoe" ? "" : esc(line.size), line.size_kind === "shoe" ? esc(line.size) : "",
    esc(line.unit), quantityCell(line.expected_qty, line.unit),
    countQuantityCell(count, line, editable),
    line.counted_qty === null ? "—" : quantityCell(line.counted_qty - line.expected_qty, line.unit, true),
    ...(can("field.cost") ? [line.unit_price === null ? "—" : esc(unitPriceFormat.format(line.unit_price / 10000))] : []),
    `<div class="actions">${editable ? `${can("field.cost") ? `<button type="button" data-count-item="${line.variant_id}">Birim maliyet</button>` : ""}
      <button type="button" data-count-delete="${line.variant_id}">Sil</button>` : ""}</div>`
  ]);
  const selectedIds = new Set(count.lines.map(line => line.variant_id));
  const canAdd = editable && data.stock.some(row => !selectedIds.has(row.id));
  const entered = count.lines.filter(line => line.counted_qty !== null).length;
  const imports = editable ? `<div class="count-import filters">
    <a class="table-link" href="/sayim-sablonu.xlsx?id=${count.id}" download>Sayım Excel şablonu</a>
    <label>Sayım Excel dosyası<input id="count-file" type="file" accept=".xlsx"></label>
    <button type="button" id="count-upload">Excel’den miktarları yükle</button>
    </div>` : "";
  return `<section class="panel count-header"><div class="panel-head">
      <div><a href="#sayimlar">Sayım listesi</a><h2>Sayım #${count.id} · ${esc(count.day)}</h2>
      <p>${countStatus(count.status)} · ${entered} / ${count.lines.length} satır sayıldı</p>
      ${count.note ? `<p>${esc(count.note)}</p>` : ""}
      ${dateLocked ? '<p>Sayım tarihi kilitli dönemde.</p>' : ""}
      ${count.documents.length ? `<p>Stok düzeltme fişleri ${
        count.documents.map(doc => "#" + doc.id + " · " + esc(doc.day)).join(", ")}</p>` : ""}
      </div><div class="actions">${actions}</div>
    </div>${imports}</section>` +
    panel("Sayım ürünleri", table(["Ürün kodu","Ürün","Grup","Renk","Beden","Ayakkabı numarası","Birim",
      "Sistem bakiyesi","Sayılan miktar","Fark", ...(can("field.cost") ? ["Birim maliyet TL"] : []),actionColumn],
      entries,"Sayım listesinde ürün kalmadı."),
      editable ? `<div class="actions">${canAdd ? '<button type="button" data-count-item="add">+ Ürün ekle</button>' : ""}
      ${transition("refresh","Sistem bakiyelerini yenile")}</div>` : "");
}
const action = (name, title, primary = false) => canAct(name)
  ? `<button type="button" data-action="${name}" class="${primary ? "primary" : ""}">${title}</button>` : "";
const stat = (name, value, key = "") => `<div class="stat" data-stat="${key}"><span>${name}</span><strong>${value}</strong></div>`;
const quantityCell = (value, unit, signed = false) => ({
  text: (signed && value > 0 ? "+" : "") + qty(value) + " " + esc(unit),
  value: value / 1000, format: "quantity", unit
});
const moneyCell = value => ({text: money(value), value: value / 100, format: "money", unit: ""});
const activeBadge = active => active ? '<span class="pill">Aktif</span>' : '<span class="pill warning">Pasif</span>';
const edit = (type, row) => canAct(type) ?
  `<button type="button" data-action="${type}" data-record-id="${esc(row.id)}" aria-label="${esc(row.name || row.account)} kaydını düzenle">Düzenle</button>` : "";
const historyButton = (type, row) =>
  `<button type="button" data-history="${type}" data-record-id="${row.id}" aria-label="${esc(row.name)} tanım geçmişi">Geçmiş</button>`;
const workStatus = active => active ? '<span class="pill">Çalışıyor</span>' : '<span class="pill warning">Çalışmıyor</span>';
const actionColumn = { label: "İşlem", action: true };
const effectiveField = () => label("Geçerlilik tarihi",
  input("effective_day", localDay(), `type="date" max="${localDay()}"${periodLimit()} required`), "version-date");

function table(columns, values, empty, {calculationColumns = {}, rowKeys = []} = {}) {
  const headings = columns.map(column => typeof column === "string"
    ? `<th scope="col">${column}</th>`
    : `<th scope="col" data-action-column>${column.label}</th>`).join("");
  const body = values.map((row, index) => (rowKeys[index] ? `<tr data-row-key="${esc(rowKeys[index])}">` : "<tr>") +
    row.map(cell => typeof cell === "object"
    ? `<td class="num" data-value="${cell.value}" data-format="${cell.format}" data-unit="${esc(cell.unit)}">${cell.text}</td>`
    : `<td>${cell}</td>`).join("") + "</tr>").join("");
  const quantityIndex = columns.findIndex(title => calculationColumns[title] === "quantity");
  const calculations = Object.keys(calculationColumns).length ? `<tr class="table-calculation-row">${columns.map((title, index) => {
    if (index < quantityIndex) return "";
    const kind = calculationColumns[title];
    const span = index === quantityIndex ? ` colspan="${index + 1}"` : "";
    return kind ? `<td${span}><div class="calculation-tools">
      ${kind === "quantity" ? '<span class="calculation-caption"></span>' : ""}
      ${select("calculation-" + kind, Object.entries(calculationLabels), calculationsByPage[page()][kind], undefined,
        `data-table-calculation="${kind}" aria-label="${title} hesabı"`)}
      <output data-calculation-result="${kind}" aria-live="polite">—</output>
    </div></td>` : "<td></td>";
  }).join("")}</tr>` : "";
  return `<div class="table-wrap" tabindex="0" role="region" aria-label="Sonuç tablosu"><table>
    <thead>${calculations}<tr class="column-headings">${headings}</tr></thead><tbody>${body}</tbody></table></div>` +
    (values.length ? "" : `<p class="empty">${empty}</p>`);
}

function panel(title, body, controls = "") {
  return `<section class="panel"><div class="panel-head"><h2>${title}</h2>${controls}</div>${body}</section>`;
}

function tabs() {
  return `<div class="tabs" aria-label="Malzeme grubu">${Object.entries(modules).map(([key, title]) =>
    `<button type="button" data-module="${key}" aria-pressed="${filters.module === key}">${title}</button>`
  ).join("")}</div>`;
}

function changeable(name, items, current, blank) {
  return select(name, items, current, blank, "data-filter");
}

function searchField() {
  return label("Ürün ara", input("search", filters.search, 'type="search" placeholder="Ürün adı veya kodu" data-filter'), "search");
}

function productTools() {
  return `<div class="product-tools">${searchField()}${action("products", "+ Ürün kartı", true)}</div>`;
}

function matchingModule(row) {
  return filters.module === "all" || row.module === filters.module;
}

function matchingSearch(row) {
  return [row.code, row.name, row.color, row.size].join(" ").toLocaleLowerCase("tr")
    .includes(filters.search.trim().toLocaleLowerCase("tr"));
}

function stockPage() {
  const stock = data.stock.filter(matchingModule);
  const values = stock.filter(matchingSearch).map(row => [
    `<strong>${esc(row.name)}</strong><span class="sub">${esc(row.code)}</span>`,
    esc(row.category), esc(row.color),
    row.size_kind === "shoe" ? "" : esc(row.size), row.size_kind === "shoe" ? esc(row.size) : "",
    quantityCell(row.qty, row.unit),
    ...(can("field.cost") ? [moneyCell(row.value)] : []),
    (!row.active ? '<span class="pill warning">Pasif ürün</span> ' : "") +
    (row.qty === 0 ? '<span class="pill warning">Stok yok</span>' :
      row.qty <= row.minimum ? '<span class="pill warning">Kritik stok</span>' : '<span class="pill">Mevcut</span>')
  ]);
  return `<div class="stats">${stat("Ürün kartı", new Set(stock.map(row => row.product_id)).size)}
    ${stat("Kritik veya tükenen", stock.filter(row => row.qty <= row.minimum).length)}
    ${can("field.cost") ? stat("Toplam stok değeri", money(stock.reduce((total, row) => total + row.value, 0))) : ""}</div>` +
    panel("Depo bakiyesi", table(["Ürün", "Grup", "Renk", "Beden", "Ayakkabı numarası", "Eldeki miktar",
      ...(can("field.cost") ? ["Stok değeri"] : []), "Durum"],
      values, data.products.length ? "Filtreye uyan stok bulunamadı." : "Stok takibine başlamak için bir ürün kartı oluşturun.",
      {calculationColumns: calculationColumnsByPage.stok}),
    productTools());
}

function productsPage() {
  const groups = data.groups.filter(matchingModule);
  const values = data.products.filter(matchingModule).filter(matchingSearch).map(row => {
    const variants = data.stock.filter(variant => variant.product_id === row.id && variantName(variant));
    return [esc(row.code), esc(row.name), esc(row.category), esc(row.unit), scopes[row.scope],
      variants.map(item => esc(variantName(item))).join(", ") || "—", activeBadge(row.active),
      `<div class="actions">${edit("product-update", row)}${historyButton("products", row)}${variants.length && canAct("variants")
        ? `<button data-variant-product="${row.id}">Varyant ekle</button>` : ""}</div>`];
  });
  return panel("Ürün listesi", table(["Kod", "Ürün", "Grup", "Birim", "Teslim kapsamı", "Renk / beden", "Durum", actionColumn],
    values, "Gösterilecek ürün kartı yok."), productTools()) +
    panel("Ürün grupları", table(["Grup", "Kod öneki", "Modül", "Ölçü türü", actionColumn],
      groups.map(row => [esc(row.name), esc(row.prefix), modules[row.module], sizeKinds[row.size_kind] || sizeKinds.size,
        `<div class="actions">${edit("product-group-size-update", row)}${historyButton("product_groups", row)}</div>`]),
      "Otomatik kodlama için Tişört (TSH) gibi bir ürün grubu ekleyin."), action("product-groups", "+ Ürün grubu"));
}

function definitionsPage() {
  return panel("Departmanlar", table(["Departman", "Durum", actionColumn],
    data.departments.map(row => [esc(row.name), activeBadge(row.active),
      `<div class="actions">${edit("department-update", row)}${historyButton("departments", row)}</div>`]),
    "Henüz departman eklenmedi."), action("departments", "+ Departman")) +
    panel("Personel", table(["Sicil", "Ad soyad", "Departman", "Çalışma durumu", actionColumn],
      data.people.map(row => {
        const version = onDay("people", row, personnelDate);
        const dept = data.departments.find(item => item.id === version.department_id);
        return [esc(row.code), esc(version.name), esc(dept ? onDay("departments", dept, personnelDate).name : row.department),
        workStatus(version.active),
        `<div class="actions">${edit("person-update", row)}${historyButton("people", row)}
          ${canAct("person-status") ? `<button type="button" data-action="person-status" data-record-id="${row.id}"
            aria-label="${esc(row.name)} çalışma durumunu değiştir">${row.active ? "Çalışmıyor yap" : "Çalışıyor yap"}</button>` : ""}</div>`];
      }),
      "Henüz personel eklenmedi."),
      `<div class="filters">${label("Çalışma durumu tarihi", input("personnelDate", personnelDate,
        `type="date" max="${localDay()}" data-filter`))}${action("people", "+ Personel")}</div>`);
}

function dateFilters(current) {
  return label("Başlangıç", input("from", current.from, 'type="date" data-filter')) +
    label("Bitiş", input("to", current.to, 'type="date" data-filter'));
}

function recipientFilters(current) {
  return label("Departman", changeable("department", departmentOptions(), current.department, "Tümü")) +
    label("Personel", changeable("person", personOptions(), current.person, "Tümü"));
}

function matches(row, current) {
  return matchingModule(row) && (!current.from || row.day >= current.from) &&
    (!current.to || row.day <= current.to) &&
    (!current.department || String(row.department_id) === current.department) &&
    (!current.person || String(row.person_id) === current.person);
}

function movementsPage() {
  const values = data.movements.filter(row => matches(row, filters) &&
    (!filters.kind || row.kind === filters.kind) && matchingSearch(row)).map(row => [
    esc(row.day) + `<span class="sub">Fiş #${row.document_id}</span>`,
    definitions[row.kind], productName(row),
    esc(row.department || row.supplier || "—") + (row.person
      ? `<span class="sub">${esc(row.person)}${row.person_active === 0 ? " · Çalışmıyor" : " · Çalışıyor"}</span>` : ""),
    quantityCell(row.qty, row.unit, true),
    ...(can("field.cost") ? [moneyCell(row.value)] : []),
    [row.reference, row.purchase_day ? "Alış " + row.purchase_day : "", row.note,
      row.source_id ? "Teslim fişi #" + row.source_id : ""].filter(Boolean).map(esc).join(" · ") || "—"
  ]);
  const controls = `<div class="filters">${dateFilters(filters)}${recipientFilters(filters)}
    ${label("Hareket türü", changeable("kind", Object.entries(definitions), filters.kind, "Tümü"))}${searchField()}</div>`;
  return panel("Hareket listesi", table(["Tarih", "İşlem", "Ürün", "Departman / teslim alan", "Miktar",
    ...(can("field.cost") ? ["Stok değeri farkı"] : []), "Belge ve açıklama"], values,
    "Filtreye uyan hareket bulunamadı.", {calculationColumns: calculationColumnsByPage.hareketler}), controls);
}

function reportsPage() {
  const controls = `<div class="filters">${dateFilters(report)}${recipientFilters(report)}</div>`;
  if (!report.from || !report.to || report.from > report.to) {
    return panel("Rapor dönemi", '<p class="empty">Geçerli bir başlangıç ve bitiş tarihi seçin.</p>', controls);
  }
  const totals = new Map();
  for (const row of data.movements) {
    if (!matches(row, report)) continue;
    const kind = row.source_kind || row.kind;
    if (kind !== "person" && kind !== "department") continue;
    const month = row.day.slice(0, 7);
    const key = JSON.stringify([month, row.department_id, row.department, row.variant_id, row.name, kind]);
    const total = totals.get(key) ?? { row, month, kind, amount: 0, cost: 0 };
    total.amount -= row.qty;
    total.cost -= row.value ?? 0;
    totals.set(key, total);
  }
  const entries = [...totals.values()];
  const values = entries.map(({ row, month, kind, amount, cost }) => [
    month, esc(row.department), definitions[kind], productName(row), quantityCell(amount, row.unit),
    ...(can("field.cost") ? [moneyCell(cost)] : [])
  ]);
  return panel("Rapor dönemi", "", controls) +
    `<div class="stats report-stats" aria-live="polite">
    ${session?.department_id ? "" : stat("Giriş miktarı", "—", "report-received-quantity")}
    ${stat("Tüketim miktarı", "—", "report-net-quantity")}
    ${session?.department_id ? "" : stat("Son kalan bakiye miktarı", "—", "report-ending-quantity")}` +
    (can("field.cost") ? `
    ${session?.department_id ? "" : stat("Depo giriş değeri", "—", "report-receipts")}
    ${stat("Net tüketim maliyeti", "—", "report-consumption")}
    ${stat("Aylık ortalama", "—", "report-average")}` : "") + "</div>" +
    panel("Aylık tüketim dökümü", table(["Ay", "Departman", "Teslim türü", "Ürün", "Net miktar",
      ...(can("field.cost") ? ["Net maliyet"] : [])],
      values, "Seçili dönemde tüketim kaydı yok.",
      {rowKeys: entries.map(({month, row}) => month + "/" + row.variant_id)}));
}

function accountsPage() {
  const entries = (data.accounts || []).map(row => [
    esc(row.account), esc(row.name), row.admin ? "Yönetici" : "Tanımlı yetkiler",
    esc(data.departments.find(department => department.id === row.department_id)?.name || "Tüm departmanlar"),
    activeBadge(row.active), edit("user-update", {...row, id:row.account})
  ]);
  const log = (data.audit || []).map(row => [
    esc(row.recorded_at.replace("T"," ").slice(0,19) + " UTC"), esc(row.account),
    esc(auditLabels[row.action] || row.action), esc(row.record_ref || "—")
  ]);
  return panel("Windows hesapları", table(["Windows hesabı","Görünen ad","Erişim","Kapsam","Durum",actionColumn],
    entries,"Windows hesabı tanımlanmamış."), action("user-update","+ Windows hesabı",true)) +
    panel("Son işlemler", table(["Zaman","Windows hesabı","İşlem türü","Kayıt"],log,"İşlem kaydı yok."));
}

function periodAccounts() {
  return (data.accounts || []).filter(user => user.active && !user.admin);
}

function periodPage() {
  const rows = (data.period_grants || []).map(grant => [
    esc(grant.account), esc(grant.start_day), esc(grant.end_day), activeBadge(!grant.revoked_at),
    esc(grant.reason || "—"), esc(grant.granted_by), esc(grant.revoked_by || "—"),
    grant.revoked_at ? "" : `<button type="button" data-action="period-revoke" data-record-id="${grant.id}">İzni kapat</button>`
  ]);
  const controls = periodAccounts().length ? action("period-add", "+ Tarih aralığı aç", true) :
    '<a class="table-link" href="#kullanicilar">Kullanıcı tanımla</a>';
  return `<p class="period-note">${esc(data.open_period_start)} ve sonraki tarihler mevcut yetkiler kapsamında açıktır.
    Daha eski kayıt tarihleri için kullanıcıya özel izin verebilirsiniz.</p>` +
    panel("Dönem kilidi istisnaları", table(["Windows hesabı","Başlangıç","Bitiş","Durum","Açıklama",
      "Açan yönetici","Kapatan yönetici",actionColumn],rows,"Henüz dönem kilidi istisnası bulunmuyor."),controls);
}

const auditLabels = {
  "period-add":"Dönem kilidi istisnası açıldı", "period-revoke":"Dönem kilidi istisnası kapatıldı",
  login:"Oturum açıldı", logout:"Oturum kapatıldı", "user-update":"Kullanıcı yetkileri kaydedildi",
  "document-receipt":"Alış girişi", "document-person":"Personele teslim", "document-department":"Departmana teslim",
  "document-return":"İade", "document-adjustment":"Sayım düzeltmesi", "excel-import":"Excel içe aktarma",
  products:"Ürün oluşturuldu", "product-update":"Ürün düzenlendi", variants:"Varyant eklendi",
  "product-groups":"Ürün grubu oluşturuldu", departments:"Departman oluşturuldu",
  "product-group-size-update":"Ürün grubu ölçü türü düzenlendi",
  "department-update":"Departman düzenlendi", people:"Personel oluşturuldu",
  "person-update":"Personel düzenlendi", "person-status":"Çalışma durumu güncellendi"
};

function excelPage() {
  return `<section class="panel import-panel">
    <div class="import-content">
      <h2>Excel şablonları</h2>
      <p>Dosyada ürün grubu, departman, personel, ürün kartı, varyant ve stok hareketi için ayrı sayfalar var.
      İhtiyacınız olan sayfalara veri girin ve sütun başlıklarını koruyun.</p>
      <a class="download" href="/sablonlar/toplu-aktarim.xlsx" download>Excel şablonlarını indir</a>
      <label>Excel dosyası (.xlsx)<input id="excel-file" type="file" accept=".xlsx"></label>
      <p id="excel-selected" class="sub"></p>
      <div class="actions">
        <button id="excel-preview" type="button">Dosyayı kontrol et</button>
        <button id="excel-import" class="primary" type="button" disabled>İçe aktar</button>
      </div>
      <div id="excel-result" role="status" aria-live="polite"></div>
    </div></section>`;
}

function showImportResult(result, saved) {
  const resultPanel = document.querySelector("#excel-result");
  if (!resultPanel) return;
  const summary = Object.entries(result.counts).map(([title, count]) => `<li>${esc(title)} · ${count} satır</li>`).join("");
  const codes = result.codes.length ? `<p>${saved ? "Oluşturulan" : "Öngörülen"} ürün kodları</p><ul>${
    result.codes.map(([reference, code]) => `<li>${esc(reference)} → ${esc(code)}</li>`).join("")}</ul>` : "";
  resultPanel.innerHTML = `<p><strong>${saved ? "Aktarım tamamlandı." : "Kontrol tamamlandı. Veriler henüz kaydedilmedi."}</strong></p>
    <ul>${summary}${result.documents ? `<li>${result.documents} stok fişi</li>` : ""}</ul>${codes}`;
}

async function uploadExcel(preview) {
  if (busy || (!preview && !importPreview)) return;
  const panel = document.querySelector("#excel-result");
  if (!importFile) {
    panel.textContent = "Önce bir .xlsx dosyası seçin.";
    return;
  }
  if (!/\.xlsx$/i.test(importFile.name) || importFile.size > 5 * 1024 * 1024) {
    panel.textContent = "En çok 5 MB boyutunda bir .xlsx dosyası seçin.";
    return;
  }
  const file = importFile;
  const check = document.querySelector("#excel-preview");
  const save = document.querySelector("#excel-import");
  const input = document.querySelector("#excel-file");
  busy = check.disabled = save.disabled = input.disabled = true;
  importPreview = false;
  panel.textContent = preview ? "Dosya kontrol ediliyor…" : "Dosya içe aktarılıyor…";
  try {
    const response = await fetch("/api/" + (preview ? "excel-preview" : "excel-import"), {
      method: "POST", headers: {"Content-Type": "application/octet-stream"}, body: file
    });
    const result = await response.json();
    if (!response.ok) {
      panel.textContent = result.error + " Dosyadan kayıt eklenmedi.";
      if (!panel.isConnected) showStatus(result.error);
      return;
    }
    showImportResult(result, !preview);
    importPreview = preview && panel.isConnected;
    if (!preview) {
      if (!panel.isConnected) showStatus("Excel aktarımı tamamlandı.");
      importFile = undefined;
      input.value = "";
      const selected = document.querySelector("#excel-selected");
      if (selected) selected.textContent = "";
      await refresh(false).catch(() => showStatus("Aktarım tamamlandı. Güncel liste alınamadı."));
    }
  } catch (error) {
    const message = "Sunucu yanıtı alınamadı. Tekrar denemeden önce kayıtları kontrol edin.";
    panel.textContent = message;
    if (!panel.isConnected) showStatus(message);
  } finally {
    busy = check.disabled = input.disabled = false;
    save.disabled = !importPreview;
  }
}

function render() {
  if (authRequired === null) return;
  const signedOut = authRequired && !session;
  document.querySelector("nav").hidden = signedOut;
  document.querySelectorAll(".toolbar [data-action]").forEach(button => { button.hidden = signedOut || !canAct(button.dataset.action); });
  const account = document.querySelector("#account-label");
  account.hidden = !session;
  account.textContent = session ? session.name + (session.admin ? " · Yönetici" : "") : "";
  document.querySelector("#logout").hidden = !session;
  if (signedOut) {
    content.innerHTML = `<section class="panel login-panel"><h1>Oturum aç</h1>
      <p>Uygulamada tanımlı şirket hesabınızla devam edin.</p>
      <a class="login-link" href="/api/windows-login">Windows hesabıyla devam et</a>
      <p id="login-error" role="alert" class="sub"></p></section>`;
    return;
  }
  if (!data) return;
  const current = page();
  const [title, description] = pages[current];
  document.querySelectorAll("nav a").forEach(link => {
    link.hidden = !canView(link.hash.slice(1));
    if (link.hash === "#" + current) link.setAttribute("aria-current", "page");
    else link.removeAttribute("aria-current");
  });
  if (!canView(current)) {
    content.innerHTML = '<p>Hesabınıza henüz ekran yetkisi verilmemiş.</p>';
    return;
  }
  const view = { stok: stockPage, hareketler: movementsPage, urunler: productsPage,
    tanimlar: definitionsPage, raporlar: reportsPage, excel: excelPage, kullanicilar: accountsPage,
    sayimlar: countPage, "donem-kilidi": periodPage };
  const movementActions = current === "hareketler" ? action("return", "İade al") : "";
  const printable = !["excel", "kullanicilar", "donem-kilidi"].includes(current) && can("export");
  const actions = `<div class="actions">${movementActions}${printable ? `
    <button type="button" data-export="xlsx">Excel indir</button>
    <button type="button" data-export="pdf">PDF indir</button>
    ${current === "sayimlar" ? '<button type="button" data-export="csv">CSV indir</button>' : ""}` : ""}</div>`;
  const viewKey = current === "sayimlar" ? "sayimlar/" + countId() : current;
  const scrollPositions = content.dataset.view === viewKey
    ? [...content.querySelectorAll(".page-body, .table-wrap")].map(element => [element.scrollLeft, element.scrollTop])
    : [];
  content.dataset.view = viewKey;
  content.innerHTML = `<header class="page-heading">
    <div class="section-heading"><div><h1>${title}</h1><p>${description}</p></div>${actions}</div>
    ${["tanimlar", "excel", "kullanicilar", "sayimlar", "donem-kilidi"].includes(current) ? "" : tabs()}
    </header><div class="page-body">${view[current]()}</div>`;
  tableFilters.prepare(viewKey);
  [...content.querySelectorAll(".page-body, .table-wrap")].forEach((element, index) => {
    const position = scrollPositions[index];
    if (position) element.scrollTo(...position);
  });
}

function syncFilteredTotals() {
  const current = page();
  if (calculationColumnsByPage[current]) showTableCalculations();
  if (!["stok","raporlar"].includes(current)) return;
  const table = tableFilters.tables[0];
  if (!table) return;
  const rows = table.rows.filter(row => !row.element.hidden);
  if (current === "stok") {
    const figures = [...content.querySelectorAll(".stat strong")];
    figures[0].textContent = new Set(rows.map(row => row.element.cells[0].querySelector(".sub").textContent)).size;
    figures[1].textContent = rows.filter(row => /Stok yok|Kritik stok/.test(row.values.at(-1))).length;
    if (figures[2]) figures[2].textContent = money(sumMoneyColumn(table, rows, "Stok değeri"));
  } else {
    syncReportCards(table, rows);
  }
}

function sumMoneyColumn(table, rows, name) {
  const index = table.headers.findIndex(header => header.name === name);
  return rows.reduce((sum, row) => sum + Math.round(Number(row.element.cells[index].dataset.value) * 100), 0);
}

function syncReportCards(table, rows) {
  const restrictedRows = Boolean(table.selected.size || report.department || report.person);
  const visibleProductsByMonth = new Set(rows.map(row => row.element.dataset.rowKey));
  const countMonth = day => Number(day.slice(0, 4)) * 12 + Number(day.slice(5, 7));
  const months = restrictedRows
    ? new Set([...visibleProductsByMonth].map(key => key.slice(0, 7))).size
    : countMonth(report.to) - countMonth(report.from) + 1;
  function update(key, value, title) {
    const element = content.querySelector(`[data-stat="report-${key}"]`);
    if (!element) return;
    if (title) element.querySelector("span").textContent = title;
    element.querySelector("strong").textContent = value;
  }

  const quantityIndex = table.headers.findIndex(header => header.name === "Net miktar");
  const netQuantities = sumByUnit(rows.map(({element}) => {
    const cell = element.cells[quantityIndex];
    return {unit: cell.dataset.unit, qty: Math.round(Number(cell.dataset.value) * 1000)};
  }));
  update("net-quantity", formatQuantities(netQuantities));

  if (can("field.cost")) {
    const net = sumMoneyColumn(table, rows, "Net maliyet");
    update("consumption", money(net));
    update("average", money(months ? net / months : 0), "Aylık ortalama · " + months + " ay");
  }

  if (session?.department_id) return;
  const receipts = data.movements.filter(row =>
    row.kind === "receipt" && matchingModule(row) && row.day >= report.from && row.day <= report.to &&
    (!restrictedRows || visibleProductsByMonth.has(row.day.slice(0, 7) + "/" + row.variant_id)));
  const scope = restrictedRows ? " · Görünen ürünler ve aylar" : "";
  update("received-quantity", formatQuantities(sumByUnit(receipts)), "Giriş miktarı" + scope);
  if (can("field.cost")) {
    update("receipts", money(receipts.reduce((sum, row) => sum + (row.value ?? 0), 0)), "Depo giriş değeri" + scope);
  }

  const variantIds = new Set([...visibleProductsByMonth].map(key => key.split("/")[1]));
  const balanceDay = reportBalanceDay(table, visibleProductsByMonth);
  const balance = data.movements.filter(row => row.day <= balanceDay && matchingModule(row) &&
    (!restrictedRows || variantIds.has(String(row.variant_id))));
  update("ending-quantity", formatQuantities(sumByUnit(balance)), "Son kalan bakiye miktarı · " + balanceDay);
}

function reportBalanceDay(table, visibleProductsByMonth) {
  const monthColumn = table.headers.findIndex(header => header.name === "Ay");
  if (!table.selected.has(monthColumn) || !visibleProductsByMonth.size) return report.to;
  const lastVisibleMonth = [...visibleProductsByMonth].map(key => key.slice(0, 7)).sort().at(-1);
  if (lastVisibleMonth >= report.to.slice(0, 7)) return report.to;
  const monthEnd = new Date(lastVisibleMonth + "-01T00:00:00Z");
  monthEnd.setUTCMonth(monthEnd.getUTCMonth() + 1, 0);
  return monthEnd.toISOString().slice(0, 10);
}

function sumByUnit(rows) {
  const totals = new Map();
  for (const {unit, qty} of rows) {
    totals.set(unit, (totals.get(unit) ?? 0) + qty);
  }
  return totals;
}

function formatQuantities(totals) {
  if (!totals.size) return "—";
  return [...totals].sort(([left], [right]) => left.localeCompare(right, "tr"))
    .map(([unit, amount]) => (qty(amount) + " " + unit).trim()).join("\n");
}

function calculateTable(table, screen) {
  const columns = calculationColumnsByPage[screen];
  if (!table || !columns) return [];
  const visible = table.rows.filter(row => !row.element.hidden);
  return table.headers.flatMap(({name}, index) => {
    const format = columns[name];
    const operation = calculationsByPage[screen][format];
    if (!operation || operation === "none") return [];
    const scale = format === "quantity" ? 1000 : 100;
    const units = new Map();
    for (const row of visible) {
      const cell = row.element.cells[index];
      const unit = cell.dataset.unit;
      const value = Math.round(Number(cell.dataset.value) * scale);
      const stats = units.get(unit) || {sum: 0, count: 0, min: value, max: value};
      stats.sum += value;
      stats.count++;
      stats.min = Math.min(stats.min, value);
      stats.max = Math.max(stats.max, value);
      units.set(unit, stats);
    }
    return [...units].sort(([first], [second]) => first.localeCompare(second, "tr")).map(([unit, stats]) => {
      const result = operation === "average" ? stats.sum / stats.count : stats[operation];
      return {name, operation, format, unit, value: result / scale,
        text: format === "quantity" ? (qty(result) + " " + unit).trim() : money(result)};
    });
  });
}

function showTableCalculations() {
  const table = tableFilters.tables[0];
  const screen = page();
  const figures = calculateTable(table, screen);
  const caption = content.querySelector(".calculation-caption");
  const filtered = table?.selected.size || filters.module !== "all" || filters.search.trim() ||
    (screen === "hareketler" && (filters.from || filters.to || filters.department || filters.person || filters.kind));
  if (caption) caption.textContent = filtered ? "Filtreli satırlar" : "Tüm satırlar";
  for (const output of content.querySelectorAll("[data-calculation-result]")) {
    const matching = figures.filter(figure => figure.format === output.dataset.calculationResult);
    output.textContent = matching.map(figure => figure.text).join(" · ") || "—";
  }
}

function exportPayload() {
  const current = page();
  const sections = [];
  const text = node => (node?.innerText ?? "").split("\n").map(line => line.trim()).filter(Boolean).join("\n");
  const stats = [...content.querySelectorAll(".stat:not([hidden])")].map(stat => [
    {text: text(stat.querySelector("span"))}, {text: text(stat.querySelector("strong"))}
  ]);
  if (stats.length) sections.push({title: "Özet", columns: ["Gösterge", "Değer"], rows: stats});
  const calculations = calculateTable(tableFilters.tables[0], current).map(result => [
    {text: result.name}, {text: calculationLabels[result.operation]},
    {text: result.text, value: result.value, format: result.format, unit: result.unit}
  ]);
  if (calculations.length) sections.push({title: "Sütun hesapları", columns: ["Sütun", "İşlem", "Sonuç"], rows: calculations});
  for (const panel of content.querySelectorAll(".panel")) {
    const table = panel.querySelector("table");
    if (!table) continue;
    const indexes = [...table.tHead.querySelector(".column-headings").cells].flatMap((cell, index) =>
      cell.hasAttribute("data-action-column") ? [] : [[text(cell), index]]);
    const rows = [...table.tBodies[0].rows].filter(row => !row.hidden).map(row => indexes.map(([, index]) => {
      const cell = row.cells[index];
      const saved = cell.querySelector("[data-saved-quantity]");
      const result = {text: saved ? saved.textContent.trim() : text(cell)};
      if (cell.hasAttribute("data-value")) {
        result.value = Number(cell.dataset.value);
        result.format = cell.dataset.format;
        result.unit = cell.dataset.unit;
      }
      return result;
    }));
    sections.push({title: text(panel.querySelector("h2")), columns: indexes.map(([name]) => name),
      rows, empty: text(panel.querySelector(".empty:not([hidden])"))});
  }
  const selected = ["tanimlar","sayimlar"].includes(current) ? [] : [modules[filters.module]];
  if (current === "sayimlar" && countId() && countDetails) {
    selected.push("Sayım #" + countDetails.id, countDetails.day, countStatuses[countDetails.status]);
  }
  for (const control of content.querySelectorAll("[data-filter]")) {
    if (!control.value) continue;
    const value = control.tagName === "SELECT" ? control.selectedOptions[0]?.textContent : control.value;
    selected.push(control.closest("label").firstChild.textContent.trim() + " " + value);
  }
  if (data.demo) selected.unshift("Demo verisi");
  selected.push(...tableFilters.describe());
  const context = selected.join(" · ");
  return {title: pages[current][0], context: context.length > 1500 ? context.slice(0, 1499) + "…" : context, sections};
}

async function downloadExport(button) {
  if (button.disabled) return;
  const extension = button.dataset.export;
  const current = page();
  const payload = exportPayload();
  if (!payload.sections.length) {
    showStatus("Çıktı almak için geçerli bir filtre seçin.");
    return;
  }
  button.disabled = true;
  try {
    const response = await fetch("/api/export-" + extension, {
      method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(payload)
    });
    if (!response.ok) throw new Error((await response.json()).error);
    const url = URL.createObjectURL(await response.blob());
    const link = document.createElement("a");
    link.href = url;
    link.download = "sarf-" + current + "-" + localDay() + "." + extension;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  } catch (error) {
    showStatus(error.message || "Çıktı hazırlanamadı.");
  } finally {
    button.disabled = false;
  }
}

async function request(path, payload) {
  const response = await fetch("/api/" + path, payload ? {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload)
  } : {});
  const result = await response.json();
  if (response.status === 401 && path === "state" && authRequired) {
    session = null;
    data = undefined;
    countQuantityDrafts.clear();
    if (modal.open) modal.close();
    render();
  }
  if (!response.ok) throw new Error(result.error);
  return result;
}

async function refresh(shouldRender = true) {
  data = await request("state");
  if (authRequired) session = data.session;
  timelines = new Map();
  for (const version of data.definition_timeline ?? []) {
    const key = version.entity + "/" + version.entity_id;
    if (!timelines.has(key)) timelines.set(key, []);
    timelines.get(key).push(version);
  }
  document.querySelector("#demo-label").hidden = !data.demo;
  const selectedCount = canView("sayimlar") ? countId() : "";
  if (selectedCount) {
    const details = await request("stock-count?" + new URLSearchParams({id:selectedCount}));
    if (countId() === selectedCount) countDetails = details;
  }
  if (shouldRender) render();
}

async function startApp() {
  try {
    const result = await request("session");
    authRequired = result.required;
    session = result.user;
    if (!authRequired || session) await refresh();
    else render();
  } catch (error) {
    if (!authRequired || session) content.innerHTML = '<p role="alert">Uygulama bağlantısı kurulamadı. Sayfayı yenileyin.</p>';
  }
}

async function windowsLogin(link) {
  if (link.getAttribute("aria-busy") === "true") return;
  link.setAttribute("aria-busy", "true");
  const feedback = document.querySelector("#login-error");
  feedback.textContent = "Windows hesabı doğrulanıyor…";
  try {
    const result = await request("windows-login");
    session = result.user;
    await refresh();
  } catch (error) {
    feedback.textContent = error.message;
  } finally {
    link.removeAttribute("aria-busy");
  }
}

function showStatus(message) {
  status.hidden = false;
  status.textContent = message;
  clearTimeout(showStatus.timer);
  showStatus.timer = setTimeout(() => { status.hidden = true; }, 5500);
}

function definitionsForm(type, productId) {
  const text = (name, title, more = "required") => label(title, input(name, "", `maxlength="250" ${more}`));
  const active = row => label("Durum", select("active", [["1", "Aktif"], ["0", "Pasif"]], row.active));
  if (type === "period-add") {
    const end = new Date(data.open_period_start + "T00:00:00Z");
    end.setUTCDate(end.getUTCDate() - 1);
    const max = end.toISOString().slice(0, 10);
    const accountOptions = periodAccounts().map(user => [user.account, user.name + " · " + user.account]);
    return label("Windows kullanıcısı", select("account", accountOptions, "", "Kullanıcı seçin", "required"), "wide") +
      label("Başlangıç tarihi", input("start_day", "", `type="date" max="${max}" required`)) +
      label("Bitiş tarihi", input("end_day", "", `type="date" max="${max}" required`)) +
      label("Açıklama", input("reason", "", 'maxlength="250" placeholder="İsteğe bağlı"'), "wide");
  }
  if (type === "period-revoke") {
    const grant = data.period_grants.find(row => String(row.id) === productId);
    return input("id", grant.id, 'type="hidden"') +
      `<p class="wide"><strong>${esc(grant.account)}</strong> için ${esc(grant.start_day)} – ${esc(grant.end_day)} yazma izni kapatılacak.</p>`;
  }
  if (type === "count-create") {
    return label("Sayım tarihi", input("day", localDay(), `type="date" max="${localDay()}"${periodLimit()} required`)) +
      label("Açıklama", input("note", "", 'maxlength="250" placeholder="İsteğe bağlı"')) +
      '<p class="wide sub">Sistem bakiyeleri seçilen tarihten alınır. Sayılmayan satırların bakiyesi onayda değişmez.</p>';
  }
  if (type === "count-line" || type === "count-transition") {
    return countForm(type, productId);
  }
  if (type === "product-group-size-update") {
    const row = data.groups.find(group => String(group.id) === productId);
    return input("id", row.id, 'type="hidden"') +
      label("Ürün grubu", input("name_display", row.name, "readonly")) +
      label("Ölçü türü", select("size_kind", Object.entries(sizeKinds), row.size_kind)) + effectiveField();
  }
  if (type === "user-update") {
    const row = data.accounts.find(account => account.account === productId);
    const permissions = Object.entries(data.permission_labels).map(([key, title]) =>
      `<label class="filter-option"><input type="checkbox" name="permission" value="${esc(key)}"${
        row && !row.admin && row.permissions.includes(key) ? " checked" : ""}> ${esc(title)}</label>`).join("");
    return label("Windows hesabı", input("account", row?.account,
      row ? "readonly required" : 'required maxlength="250" placeholder="ETKİALANI\\kullanıcı"')) +
      label("Görünen ad", input("name", row?.name, "required maxlength=250")) +
      label("Departman kapsamı", select("department_id", departmentOptions(), row?.department_id, "Tüm departmanlar")) +
      active(row || {active:1}) +
      `<label class="filter-option wide"><input type="checkbox" name="admin"${row?.admin ? " checked" : ""}> Uygulama yöneticisi</label>
        <fieldset class="permission-grid wide"><legend>Ekran, işlem ve alan yetkileri</legend>${permissions}</fieldset>`;
  }
  if (type === "person-status") {
    const row = data.people.find(item => String(item.id) === productId);
    return input("id", row.id, 'type="hidden"') + input("active", row.active ? "0" : "1", 'type="hidden"') +
      `<p class="wide"><strong>${esc(row.name)}</strong> için çalışma durumu <strong>${row.active ? "Çalışmıyor" : "Çalışıyor"}</strong> olarak kaydedilecek.</p>` + effectiveField();
  }
  if (type === "person-update") {
    const row = data.people.find(item => String(item.id) === productId);
    const departments = data.departments.filter(dept => dept.active || dept.id === row.department_id)
      .map(dept => [dept.id, dept.name]);
    return input("id", row.id, 'type="hidden"') +
      label("Sicil numarası", input("code_display", row.code, "readonly")) +
      label("Ad soyad", input("name", row.name, "required maxlength=250")) +
      label("Departman", select("department_id", departments, row.department_id, "Departman seçin", "required")) +
      label("Çalışma durumu", select("active", [["1", "Çalışıyor"], ["0", "Çalışmıyor"]], row.active)) + effectiveField();
  }
  if (type === "department-update") {
    const row = data.departments.find(item => String(item.id) === productId);
    return input("id", row.id, 'type="hidden"') +
      label("Departman adı", input("name", row.name, "required maxlength=250")) + active(row) + effectiveField();
  }
  if (type === "product-update") {
    const row = data.products.find(item => String(item.id) === productId);
    return input("id", row.id, 'type="hidden"') +
      label("Ürün kodu", input("code_display", row.code, "readonly")) +
      label("Ürün grubu", input("group_display", row.category, "readonly")) +
      label("Ürün / model adı", input("name", row.name, "required maxlength=250")) +
      label(row.has_movements ? "Stok birimi (hareket kaydı var)" : "Stok birimi",
        input("unit", row.unit, row.has_movements ? "readonly" : "required maxlength=250")) +
      label("Teslim kapsamı", select("scope", Object.entries(scopes), row.scope)) + active(row) + effectiveField();
  }
  if (type === "departments") return text("name", "Departman adı");
  if (type === "people") return text("code", "Sicil numarası") + text("name", "Ad soyad") +
    label("Departman", select("department_id", departmentOptions(true), "", "Departman seçin", "required"));
  if (type === "product-groups") return text("name", "Ürün grubu adı", 'required placeholder="Örn. Tişört"') +
    label("Kod öneki", input("prefix", "", 'required pattern="[A-Za-z]{2,8}" maxlength="8" placeholder="Örn. TSH"')) +
    label("Modül", select("module", Object.entries(modules).slice(1), filters.module)) +
    label("Ölçü türü", select("size_kind", Object.entries(sizeKinds), "", "Grup adına göre otomatik"));
  if (type === "variants") return input("product_id", productId, 'type="hidden"') +
    text("size", "Beden veya numara", 'placeholder="Örn. M veya 42"') + text("color", "Renk", 'placeholder="Örn. Lacivert"') +
    label("Kritik stok", input("minimum", 0, 'type="number" min="0" step=".001" required')) + effectiveField();
  const groups = data.groups.filter(matchingModule);
  return label("Ürün grubu", select("group_id", groups.map(row => [row.id, row.name + " · " + row.prefix]),
    groups.length === 1 ? groups[0].id : "", "Grup seçin", "required")) +
    text("name", "Ürün / model adı") +
    `<p class="sub wide">Ürün kodu kaydettiğinizde grup öneki ve sıradaki numarayla oluşur. Renk ve beden gerekiyorsa aşağıda tanımlayın. Diğer renkleri kayıttan sonra aynı karta ekleyebilirsiniz.</p>` +
    text("unit", "Stok birimi", 'required placeholder="Adet, çift, kg…"') +
    label("Teslim kapsamı", select("scope", Object.entries(scopes), "department")) +
    text("sizes", "Bedenler veya numaralar", 'placeholder="Örn. 42, 43, 44"') +
    text("color", "Renk", 'placeholder="Örn. Lacivert"') +
    label("Stok satırı başına kritik miktar", input("minimum", 0, 'type="number" min="0" step=".001" required'));
}

const countTransitions = {
  lock: ["Sayımı kilitle", "Kilitlenen sayımı yalnızca yönetici işleyebilir. Depo bakiyeleri onayla güncellenir."],
  reopen: ["Taslağı yeniden aç", "Sayım yeniden düzenlemeye açılacak."],
  reject: ["Sayımı reddet", "Sayım raporu Onaylanmadı durumunda saklanacak. Depo bakiyesi değişmeyecek."],
  approve: ["Sayımı onayla", "Sayılan satırlardaki stok farkları düzeltme fişi olarak kaydedilecek. Sayımdan sonraki hareketler varsa düzeltme bunların tarihini izler."],
  refresh: ["Sistem bakiyelerini yenile", "Sayım tarihindeki sistem bakiyeleri yeniden alınacak. Girilen sayım miktarları korunacak."]
};

function countForm(type, selection) {
  const count = countDetails;
  const identity = input("id", count.id, 'type="hidden"') + input("revision", count.revision, 'type="hidden"');
  if (type === "count-transition") {
    return identity + input("operation", selection, 'type="hidden"') +
      `<p class="wide">${countTransitions[selection][1]}</p>`;
  }
  if (selection === "add") {
    const ids = new Set(count.lines.map(line => line.variant_id));
    const available = data.stock.filter(row => !ids.has(row.id))
      .map(row => [row.id, [row.code, row.name, variantName(row)].filter(Boolean).join(" · ")]);
    return identity + input("operation", "add", 'type="hidden"') +
      label("Ürün", select("variant_id", available, "", "Ürün seçin", "required"), "wide");
  }
  const [operation, key] = selection.split("/");
  const row = count.lines.find(line => String(line.variant_id) === key);
  const caption = [row.code, row.name, variantName(row)].filter(Boolean).map(esc).join(" · ");
  const fixed = identity + input("operation", operation, 'type="hidden"') + input("variant_id", row.variant_id, 'type="hidden"');
  if (operation === "delete") return fixed + `<p class="wide"><strong>${caption}</strong> sayım listesinden çıkarılacak.</p>`;
  return fixed + `<p class="wide"><strong>${caption}</strong><span class="sub">Sistem bakiyesi ${qty(row.expected_qty)} ${esc(row.unit)}</span></p>` +
    `<div><label for="count-modal-quantity">Sayılan miktar</label>${countQuantityStepper(
      countQuantityDrafts.get(countQuantityKey(count, row.variant_id)) ?? countQuantityValue(row),
      "Sayılan miktar", 'id="count-modal-quantity"')}</div>` +
    (can("field.cost") ? label("Birim maliyet TL", input("unit_price", row.unit_price === null ? "" : row.unit_price / 10000,
      'type="number" min="0" step=".0001" placeholder="Stok maliyetinden al"')) : "");
}

function documentForm(kind, day = localDay()) {
  const isDelivery = ["person", "department"].includes(kind);
  const deliveryDocs = [...new Map(data.movements.filter(row => ["person", "department"].includes(row.kind) && dateIsOpen(row.day))
    .map(row => [row.document_id, row])).values()];
  return `<div class="form-grid">
    ${isDelivery ? label("Teslim türü", select("kind", [["person", definitions.person], ["department", definitions.department]], kind)) : ""}
    ${label("İşlem tarihi", input("day", day, `type="date"${periodLimit()} required`))}
    ${kind === "receipt" ? label("Alış tarihi", input("purchase_day", day, `type="date"${periodLimit()} required`)) +
      label("Tedarikçi", input("supplier", "", "required maxlength=250")) : ""}
    ${kind === "department" ? label("Teslim edilen departman", select("department_id", departmentOptions(true, day), "", "Departman seçin", "required")) : ""}
    ${isDelivery ? label(kind === "department" ? "Teslim alan yetkili" : "Teslim alan personel",
      select("person_id", personOptions(true, day), "", "Kişi seçin", "required")) : ""}
    ${kind === "return" ? label("Teslim fişi", select("source_id", deliveryDocs.map(row =>
      [row.document_id, "#" + row.document_id + " · " + row.day + " · " + (row.person || row.department)]),
      "", "Teslim fişi seçin", "required")) : ""}
    ${label("Belge numarası", input("reference", "", kind === "receipt" ? "required maxlength=250" : "maxlength=250"))}
    ${can("field.notes") ? label(kind === "adjustment" ? "Düzeltme gerekçesi" : "Açıklama", input("note", "",
      kind === "adjustment" ? "required maxlength=250" : "maxlength=250"), "wide") : ""}
    </div><div class="lines" id="lines"></div><button type="button" id="add-line">+ Ürün satırı</button>`;
}

function addLine(preset = {}) {
  const kind = activeForm.kind;
  const day = form.elements.namedItem("day")?.value;
  const source = form.elements.namedItem("source_id")?.value;
  const sourceLines = data.movements.filter(row => String(row.document_id) === source);
  const candidates = data.stock.filter(row => {
    if (kind === "return") return sourceLines.some(line => line.variant_id === row.id);
    const version = onDay("products", row, day, row.product_id);
    if (!version.active && kind !== "adjustment") return false;
    return !["person", "department"].includes(kind) || [kind, "both"].includes(version.scope);
  });
  const hasPrice = kind === "receipt" || kind === "adjustment";
  const line = document.createElement("div");
  line.className = "line" + (hasPrice ? "" : " no-price");
  line.innerHTML = label("Ürün / renk / beden", select("variant_id", candidates.map(row =>
    [row.id, [row.code, onDay("products", row, day, row.product_id).name, variantName(row)].filter(Boolean).join(" · ") +
      " — " + qty(row.qty) + " " + row.unit]),
    preset.variant_id, "Ürün seçin", "required")) +
    label(kind === "adjustment" ? "Sayılan miktar" : "Miktar",
      input("qty", preset.qty, `type="number" min="${kind === "adjustment" ? "0" : ".001"}" step=".001" required`)) +
    (hasPrice ? label(kind === "adjustment" ? "Artış birim TL" : "Birim fiyat (TL)",
      input("unit_price", preset.unit_price, `type="number" min="0" step=".0001" ${kind === "receipt" ? "required" : ""}`)) : "") +
    '<button type="button" data-remove-line aria-label="Ürün satırını kaldır">×</button>';
  document.querySelector("#lines").append(line);
}

function openForm(type, productId) {
  if (!data || busy || !canAct(type)) return;
  if (type === "count-transition" && ["lock", "approve", "reject"].includes(productId) && hasCountQuantityDrafts()) {
    showStatus("Önce değiştirdiğiniz miktarları satırlardaki Kaydet düğmesiyle kaydedin.");
    return;
  }
  if (type === "products" && !data.groups.some(matchingModule)) {
    showStatus("Ürün kartı oluşturmadan önce ürün grubunu tanımlayın.");
    type = "product-groups";
  }
  const isDocument = Object.hasOwn(definitions, type);
  const titles = { products: "Yeni ürün kartı", variants: "Yeni renk / beden", departments: "Yeni departman",
    people: "Yeni personel", "product-groups": "Yeni ürün grubu", "product-group-size-update": "Ürün grubu ölçü türü",
    "product-update": "Ürün kartını düzenle", "department-update": "Departmanı düzenle",
    "person-status": "Personel çalışma durumu", "person-update": "Personeli düzenle",
    "user-update": "Windows hesabı ve yetkileri", "count-create": "Yeni fiziksel sayım",
    "count-line": productId === "add" ? "Sayım listesine ürün ekle" :
      String(productId).startsWith("delete/") ? "Sayım satırını sil" : "Sayım miktarını kaydet",
    "count-transition": countTransitions[productId]?.[0], "period-add": "Tarih aralığı aç",
    "period-revoke": "Dönem iznini kapat" };
  const title = ["person", "department"].includes(type) ? "Malzeme teslimi" : definitions[type] || titles[type];
  activeForm = isDocument ? { route: "documents", kind: type, request_id: crypto.randomUUID() } : { route: type };
  document.querySelector("#modal-title").textContent = title;
  formError.hidden = true;
  fields.innerHTML = isDocument ? documentForm(type) : `<div class="form-grid">${definitionsForm(type, productId)}</div>`;
  if (isDocument) addLine();
  modal.showModal();
}

async function openHistory(type, id) {
  const requestId = ++historyRequest;
  const dialog = document.querySelector("#history-modal");
  const target = document.querySelector("#history-results");
  const collection = type === "product_groups" ? data.groups : data[type];
  const record = collection.find(item => String(item.id) === id);
  document.querySelector("#history-title").textContent = record.name + " · Tanım geçmişi";
  dialog.showModal();
  target.textContent = "Tanım geçmişi yükleniyor…";
  try {
    const {versions} = await request("definition-history?" + new URLSearchParams({entity:type, id}));
    if (requestId !== historyRequest || !dialog.open) return;
    target.innerHTML = versions.map(({effective_day, recorded_at, state}, index) => {
      const dateLabel = effective_day === "0001-01-01" ? "Başlangıç kaydı" : "Geçerlilik " + effective_day;
      const time = new Intl.DateTimeFormat("tr-TR", {dateStyle:"short", timeStyle:"medium", timeZone:"Europe/Istanbul"})
        .format(new Date(recorded_at));
      const rows = [["Ad", esc(state.name)]];
      if (state.code) rows.push([type === "people" ? "Sicil" : "Stok kodu", esc(state.code)]);
      if (state.prefix) rows.push(["Önek", esc(state.prefix)]);
      if (state.size_kind) rows.push(["Ölçü türü", sizeKinds[state.size_kind]]);
      if (state.module) rows.push(["Modül", modules[state.module]]);
      if (state.category) rows.push(["Ürün grubu", esc(state.category)]);
      if (state.unit) rows.push(["Stok birimi", esc(state.unit)]);
      if (state.scope) rows.push(["Teslim kapsamı", scopes[state.scope]]);
      if (state.department) rows.push(["Departman", esc(state.department)]);
      if (Object.hasOwn(state, "active")) rows.push(["Durum", type === "people" ? workStatus(state.active) : activeBadge(state.active)]);
      const variants = state.variants?.map(row => [esc(row.color || "—"), esc(row.size || "—"), qty(row.minimum)]);
      return `<section class="panel"><div class="panel-head"><div>
        <h2>${esc(dateLabel)}${index === 0 ? " · Güncel sürüm" : ""}</h2>
        <p>Kayıt zamanı ${esc(time)}</p></div></div>
        ${table(["Alan", "Değer"], rows, "")}
        ${variants ? table(["Renk", "Beden / numara", "Kritik stok"], variants, "Varyant yok.") : ""}</section>`;
    }).join("");
    tableFilters.mount("history/" + type + "/" + id, target);
  } catch (error) {
    if (requestId === historyRequest && dialog.open) target.textContent = error.message;
  }
}

function readLines() {
  return [...fields.querySelectorAll(".line")].map(line =>
    Object.fromEntries([...line.querySelectorAll("[name]")].map(control => [control.name, control.value])));
}

fields.addEventListener("change", event => {
  const control = event.target;
  if (control.name === "source_id") {
    fields.querySelector("#lines").replaceChildren();
    addLine();
  }
  if (control.name === "kind" || control.name === "day") {
    const values = Object.fromEntries(new FormData(form));
    const lines = readLines();
    if (control.name === "kind") activeForm.kind = control.value;
    fields.innerHTML = documentForm(activeForm.kind, values.day);
    for (const [name, value] of Object.entries(values)) {
      if (["kind", "variant_id", "qty", "unit_price"].includes(name)) continue;
      const field = form.elements.namedItem(name);
      if (field) field.value = value;
    }
    lines.forEach(addLine);
    form.elements.namedItem(control.name)?.focus();
  }
});

form.addEventListener("submit", async event => {
  event.preventDefault();
  if (busy) return;
  const payload = Object.fromEntries(new FormData(form));
  if (activeForm.route === "user-update") {
    payload.admin = form.elements.admin.checked;
    payload.permissions = new FormData(form).getAll("permission");
  }
  if (activeForm.kind) {
    payload.kind = activeForm.kind;
    payload.request_id = activeForm.request_id;
    payload.lines = readLines();
  }
  const submit = form.querySelector('[type="submit"]');
  busy = submit.disabled = true;
  formError.hidden = true;
  try {
    const result = await request(activeForm.route, payload);
    modal.close();
    const route = activeForm.route;
    if (route === "count-line") countQuantityDrafts.delete(countQuantityKey(countDetails, payload.variant_id));
    showStatus(route.startsWith("count-") ? `Sayım #${result.id} güncellendi.` :
      route === "period-add" ? "Dönem izni açıldı." : route === "period-revoke" ? "Dönem izni kapatıldı." :
      result.id ? `Fiş #${result.id} kaydedildi.` :
      result.code ? `${result.code} ürün kartı oluşturuldu.` :
      activeForm.route === "person-status" ? "Çalışma durumu güncellendi." :
      activeForm.route.endsWith("-update") ? "Değişiklikler kaydedildi." : "Kayıt oluşturuldu.");
    if (route === "count-create") location.hash = "sayimlar/" + result.id;
    else await refresh().catch(() => showStatus("Kayıt tamamlandı. Güncel liste alınamadı."));
  } catch (error) {
    formError.textContent = error.message;
    formError.hidden = false;
    formError.focus();
  } finally {
    busy = submit.disabled = false;
  }
});

async function saveCountQuantity(editor) {
  if (busy || !countDetails || !canAct("count-line") || !countWritable(countDetails)) return;
  const field = editor.querySelector("[data-count-quantity]");
  if (!field.reportValidity()) return;
  const button = editor.querySelector("[data-count-save]");
  const count = countDetails;
  const variantId = field.dataset.countQuantity;
  busy = button.disabled = field.disabled = true;
  try {
    await request("count-line", {
      id:count.id, revision:count.revision, variant_id:variantId, operation:"save", counted_qty:field.value
    });
    countQuantityDrafts.delete(countQuantityKey(count, variantId));
    showStatus("Sayım miktarı kaydedildi.");
    await refresh().catch(() => showStatus("Miktar kaydedildi. Güncel liste alınamadı."));
  } catch (error) {
    showStatus(error.message);
  } finally {
    busy = button.disabled = field.disabled = false;
  }
}

async function uploadCount() {
  if (busy || !countDetails) return;
  if (hasCountQuantityDrafts()) {
    showStatus("Excel yüklemeden önce değiştirdiğiniz miktarları kaydedin.");
    return;
  }
  const file = document.querySelector("#count-file").files[0];
  if (!file || !/\.xlsx$/i.test(file.name) || file.size > 5 * 1024 * 1024) {
    showStatus("En çok 5 MB boyutunda bir .xlsx sayım dosyası seçin.");
    return;
  }
  const button = document.querySelector("#count-upload");
  busy = button.disabled = true;
  try {
    const query = new URLSearchParams({id:countDetails.id, revision:countDetails.revision});
    const response = await fetch("/api/count-import?" + query, {
      method:"POST", headers:{"Content-Type":"application/octet-stream"}, body:file
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error);
    showStatus(result.rows + " sayım satırı kaydedildi.");
    await refresh().catch(() => showStatus("Sayım kaydedildi. Güncel liste alınamadı."));
  } catch (error) {
    showStatus(error.message || "Excel dosyası işlenemedi.");
  } finally {
    busy = button.disabled = false;
  }
}

document.addEventListener("click", event => {
  const login = event.target.closest(".login-link");
  if (login) {
    event.preventDefault();
    windowsLogin(login);
    return;
  }
  const button = event.target.closest("button");
  if (!button) return;
  if (button.dataset.countStep) {
    stepCountQuantity(button.closest(".count-stepper").querySelector("input"), Number(button.dataset.countStep));
    return;
  }
  if (button.dataset.countSave) saveCountQuantity(button.closest(".count-quantity-editor"));
  if (button.dataset.countItem) openForm("count-line", button.dataset.countItem === "add" ? "add" : "save/" + button.dataset.countItem);
  if (button.dataset.countDelete) openForm("count-line", "delete/" + button.dataset.countDelete);
  if (button.dataset.countTransition) openForm("count-transition", button.dataset.countTransition);
  if (button.id === "count-upload") uploadCount();
  if (button.id === "logout") {
    request("logout", {}).then(() => {
      session = null;
      data = undefined;
      countQuantityDrafts.clear();
      tableFilters.selections.clear();
      render();
    }).catch(error => showStatus(error.message));
    return;
  }
  if (button.dataset.action) openForm(button.dataset.action, button.dataset.recordId);
  if (button.dataset.history) openHistory(button.dataset.history, button.dataset.recordId);
  if (button.dataset.export) downloadExport(button);
  if (button.id === "excel-preview") uploadExcel(true);
  if (button.id === "excel-import") uploadExcel(false);
  if (button.dataset.variantProduct) openForm("variants", button.dataset.variantProduct);
  if (button.dataset.module) {
    filters.module = button.dataset.module;
    render();
    content.querySelector(`[data-module="${filters.module}"]`)?.focus();
  }
  if (button.id === "add-line" && fields.querySelectorAll(".line").length < 100) addLine();
  if (Object.hasOwn(button.dataset, "removeLine") && fields.querySelectorAll(".line").length > 1) button.closest(".line").remove();
  if (["close-modal", "cancel-modal"].includes(button.id) && !busy) modal.close();
});

modal.addEventListener("cancel", event => { if (busy) event.preventDefault(); });
content.addEventListener("change", event => {
  const calculation = event.target.dataset.tableCalculation;
  if (calculation) {
    calculationsByPage[page()][calculation] = event.target.value;
    showTableCalculations();
    return;
  }
  if (event.target.id !== "excel-file") return;
  importFile = event.target.files[0];
  importPreview = false;
  document.querySelector("#excel-import").disabled = true;
  document.querySelector("#excel-result").textContent = "";
  document.querySelector("#excel-selected").textContent = importFile?.name ?? "";
});
content.addEventListener("input", event => {
  const input = event.target;
  if (input.dataset.countQuantity && countDetails) {
    const key = countQuantityKey(countDetails, input.dataset.countQuantity);
    const line = countDetails.lines.find(row => String(row.variant_id) === input.dataset.countQuantity);
    if (line && countQuantityValue(line) === input.value) countQuantityDrafts.delete(key);
    else countQuantityDrafts.set(key, input.value);
    return;
  }
  if (!Object.hasOwn(input.dataset, "filter")) return;
  const current = page() === "raporlar" ? report : filters;
  if (input.name === "personnelDate") personnelDate = input.value;
  else current[input.name] = input.value;
  const cursor = input.type === "search" ? input.selectionStart : null;
  render();
  const next = content.querySelector(`[name="${input.name}"]`);
  next?.focus();
  if (cursor !== null) next.setSelectionRange(cursor, cursor);
});
document.addEventListener("keydown", event => {
  if (event.target.matches("[data-count-step-field]") && ["ArrowUp", "ArrowDown"].includes(event.key)) {
    event.preventDefault();
    stepCountQuantity(event.target, event.key === "ArrowUp" ? 1 : -1);
    return;
  }
  if (event.key !== "Enter" || !event.target.dataset.countQuantity) return;
  event.preventDefault();
  saveCountQuantity(event.target.closest(".count-quantity-editor"));
});

window.addEventListener("hashchange", () => {
  filters.search = "";
  importFile = undefined;
  importPreview = false;
  render();
  if (countId() && canView("sayimlar") && data) refresh().catch(error => {
    if (countId() && page() === "sayimlar") content.textContent = error.message;
  });
});
const tableFilters = new TableFilters(content, syncFilteredTotals);
document.querySelector("#history-modal").addEventListener("close", () => {
  if (tableFilters.menu.closest("#history-modal")) tableFilters.close(false);
});
startApp();
setInterval(() => {
  if (data && !modal.open && !document.querySelector("#history-modal").open && !tableFilters.editor &&
      !content.contains(document.activeElement) && !document.hidden && !busy && page() !== "excel" &&
      !document.querySelector("#count-file")?.files.length && !hasCountQuantityDrafts()) {
    refresh().catch(() => showStatus("Güncel stok alınamadı."));
  }
}, 5000);
