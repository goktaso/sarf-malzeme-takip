"use strict";

class TableFilters {
  constructor(container, onChange) {
    this.container = container;
    this.onChange = onChange;
    this.menu = document.querySelector("#column-filter");
    this.selections = new Map();
    this.controls = new WeakMap();
    this.tables = [];
    this.editor = null;
    this.collator = new Intl.Collator("tr", { numeric: true });
    this.menu.addEventListener("toggle", () => {
      if (!this.menu.matches(":popover-open")) this.resetEditor();
    });
    document.addEventListener("click", event => {
      const trigger = event.target.closest("[data-column-filter]");
      if (trigger) this.open(trigger);
      const clear = event.target.closest("[data-clear-filters]");
      if (clear) {
        const table = this.controls.get(clear);
        table.selected.clear();
        this.apply(table);
        table.headers.find(header => header.trigger)?.trigger.focus();
      }
    });
    window.addEventListener("resize", () => this.positionMenu());
    window.addEventListener("scroll", () => this.positionMenu(), { capture: true, passive: true });
  }

  prepare(page) {
    this.close(false);
    this.tables = this.mount(page, this.container);
    this.onChange();
  }

  mount(page, container) {
    return [...container.querySelectorAll("table")].map((element, tableIndex) => {
      const panel = element.closest(".panel");
      const title = panel?.querySelector("h2")?.textContent.trim() || "Tablo";
      const headingRow = element.tHead.querySelector(".column-headings");
      const signature = [...headingRow.cells].map(cell => cell.textContent.trim());
      const key = JSON.stringify([page, title, tableIndex, signature]);
      if (!this.selections.has(key)) this.selections.set(key, new Map());
      const headers = [...headingRow.cells].map((cell, column) => {
        const name = cell.textContent.trim();
        if (cell.hasAttribute("data-action-column")) return { name, trigger: null };
        const trigger = document.createElement("button");
        trigger.type = "button";
        trigger.className = "column-trigger";
        trigger.dataset.columnFilter = String(column);
        trigger.setAttribute("aria-label", name + " sütununu filtrele");
        trigger.setAttribute("aria-haspopup", "dialog");
        trigger.setAttribute("aria-controls", this.menu.id);
        trigger.setAttribute("aria-expanded", "false");
        trigger.innerHTML = '<span class="column-arrow" aria-hidden="true"></span>';
        const heading = document.createElement("span");
        heading.className = "column-heading";
        const label = document.createElement("span");
        label.textContent = name;
        heading.append(label, trigger);
        cell.replaceChildren(heading);
        return { name, trigger };
      });
      const rows = [...element.tBodies[0].rows].map(row => ({
        element: row,
        values: [...row.cells].map(cell => this.cellValue(cell))
      }));
      const footer = document.createElement("div");
      footer.className = "column-filter-summary";
      footer.innerHTML = '<span role="status"></span><button type="button">Sütun filtrelerini temizle</button>';
      const clear = footer.querySelector("button");
      clear.setAttribute("data-clear-filters", "");
      const wrapper = element.closest(".table-wrap");
      wrapper.before(footer);
      const empty = document.createElement("p");
      empty.className = "empty";
      empty.textContent = "Sütun filtrelerine uyan kayıt bulunamadı.";
      wrapper.after(empty);
      const table = { title, headers, rows, selected: this.selections.get(key), footer, empty,
        main: container === this.container };
      for (const header of headers) {
        if (header.trigger) this.controls.set(header.trigger, table);
      }
      this.controls.set(clear, table);
      this.apply(table, false);
      return table;
    });
  }

  cellValue(cell) {
    const copy = cell.cloneNode(true);
    for (const element of copy.querySelectorAll("[data-cell-controls]")) element.remove();
    for (const element of copy.querySelectorAll(".sub")) element.prepend(" · ");
    return copy.textContent.replace(/\s+/g, " ").trim();
  }

  matches(row, selected, except = -1) {
    for (const [column, values] of selected) {
      if (column !== except && !values.has(row.values[column])) return false;
    }
    return true;
  }

  apply(table, notify = true) {
    let count = 0;
    for (const row of table.rows) {
      const matches = this.matches(row, table.selected);
      row.element.hidden = !matches;
      if (matches) count++;
    }
    table.headers.forEach(({ name, trigger }, index) => {
      if (!trigger) return;
      const active = table.selected.has(index);
      trigger.classList.toggle("is-filtered", active);
      trigger.setAttribute("aria-label", name + " sütunu filtresi" + (active ? " (uygulanmış)" : ""));
    });
    table.empty.hidden = !table.rows.length || count > 0;
    table.footer.hidden = !table.selected.size;
    table.footer.querySelector("span").textContent = count + " / " + table.rows.length + " kayıt";
    if (notify && table.main) this.onChange();
  }

  open(trigger) {
    if (this.editor?.trigger === trigger && this.menu.matches(":popover-open")) {
      this.close();
      return;
    }
    this.close(false);
    const table = this.controls.get(trigger);
    const column = Number(trigger.dataset.columnFilter);
    const all = new Set(table.rows.map(row => row.values[column]));
    const available = [...new Set(table.rows.filter(row => this.matches(row, table.selected, column))
      .map(row => row.values[column]))].sort(this.collator.compare);
    const selected = table.selected.get(column);
    const draft = selected ? new Set(selected) : new Set(all);
    this.editor = { table, column, trigger, all, available, draft, edited: false, shown: [] };
    const heading = this.container.ownerDocument.createElement("h2");
    heading.id = "column-filter-title";
    heading.textContent = table.headers[column].name + " filtresi";
    const form = document.createElement("form");
    form.innerHTML = `<label>Bu sütunda ara<input type="search" placeholder="Değer ara"></label>
      <label class="filter-option filter-all"><input type="checkbox" data-select-all> Tümünü seç</label>
      <div class="filter-options" role="group" aria-label="Sütun değerleri"></div>
      <div class="filter-actions">
        <button type="button" data-reset-column>Temizle</button>
        <button type="submit" class="primary">Uygula</button>
      </div>`;
    this.menu.replaceChildren(heading, form);
    form.querySelector('input[type="search"]').addEventListener("input", () => this.renderOptions());
    form.addEventListener("change", event => this.changeOption(event.target));
    form.querySelector("[data-reset-column]").addEventListener("click", () => {
      table.selected.delete(column);
      this.apply(table);
      this.close();
    });
    form.addEventListener("submit", event => {
      event.preventDefault();
      const { draft, edited, shown } = this.editor;
      const search = form.querySelector('input[type="search"]').value.trim();
      const values = !edited && search ? new Set(shown) : draft;
      if (edited || search) {
        const allAvailable = values.size === available.length && available.every(value => values.has(value));
        if ((edited && values.size === 0) || allAvailable || [...all].every(value => values.has(value))) {
          table.selected.delete(column);
        }
        else table.selected.set(column, values);
      }
      this.apply(table);
      this.close();
    });
    this.renderOptions();
    trigger.setAttribute("aria-expanded", "true");
    const parent = trigger.closest("dialog") || document.body;
    if (this.menu.parentElement !== parent) parent.append(this.menu);
    this.menu.showPopover();
    this.positionMenu();
    form.querySelector('input[type="search"]').focus();
  }

  renderOptions() {
    const editor = this.editor;
    const query = this.menu.querySelector('input[type="search"]').value.trim().toLocaleLowerCase("tr");
    editor.shown = editor.available.filter(value => (value || "Boş").toLocaleLowerCase("tr").includes(query));
    const options = this.menu.querySelector(".filter-options");
    const elements = editor.shown.map((value, index) => {
      const label = document.createElement("label");
      label.className = "filter-option";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.dataset.optionIndex = String(index);
      checkbox.checked = editor.draft.has(value);
      label.append(checkbox, document.createTextNode(value || "(Boş)"));
      return label;
    });
    if (!elements.length) {
      const empty = document.createElement("p");
      empty.textContent = "Eşleşen değer yok.";
      elements.push(empty);
    }
    options.replaceChildren(...elements);
    this.updateSelectAll();
  }

  changeOption(input) {
    const { available, draft, shown } = this.editor;
    if (input.type !== "checkbox") return;
    const bulk = input.hasAttribute("data-select-all");
    if (bulk) draft.clear();
    const values = bulk ? available : [shown[Number(input.dataset.optionIndex)]];
    for (const value of values) {
      if (input.checked) draft.add(value);
      else draft.delete(value);
    }
    this.editor.edited = true;
    if (bulk) {
      this.menu.querySelectorAll("[data-option-index]").forEach(option => { option.checked = input.checked; });
    }
    this.updateSelectAll();
  }

  updateSelectAll() {
    const { available, draft } = this.editor;
    const checkbox = this.menu.querySelector("[data-select-all]");
    const count = available.filter(value => draft.has(value)).length;
    checkbox.disabled = available.length === 0;
    checkbox.checked = available.length > 0 && count === available.length;
    checkbox.indeterminate = count > 0 && !checkbox.checked;
  }

  positionMenu() {
    if (!this.editor || !this.menu.matches(":popover-open")) return;
    const trigger = this.editor.trigger;
    const rect = trigger.getBoundingClientRect();
    const width = document.documentElement.clientWidth;
    const height = document.documentElement.clientHeight;
    if (rect.bottom < 0 || rect.top > height || rect.right < 0 || rect.left > width) {
      this.close(false);
      return;
    }
    const margin = 8;
    const left = Math.min(Math.max(margin, rect.left), width - this.menu.offsetWidth - margin);
    const below = rect.bottom + margin;
    const above = rect.top - this.menu.offsetHeight - margin;
    const top = below + this.menu.offsetHeight <= height ? below : Math.max(margin, above);
    this.menu.style.left = Math.max(margin, left) + "px";
    this.menu.style.top = top + "px";
  }

  resetEditor() {
    this.editor?.trigger.setAttribute("aria-expanded", "false");
    this.editor = null;
  }

  close(focus = true) {
    const trigger = this.editor?.trigger;
    this.resetEditor();
    if (this.menu.matches(":popover-open")) this.menu.hidePopover();
    if (focus) trigger?.focus();
  }

  describe() {
    return this.tables.flatMap(table => [...table.selected].map(([column, values]) => {
      const entries = [...values].map(value => value || "(Boş)");
      return table.title + " / " + table.headers[column].name + " " +
        (entries.slice(0, 4).join(", ") || "(Seçim yok)") + (entries.length > 4 ? " +" + (entries.length - 4) : "");
    }));
  }
}
