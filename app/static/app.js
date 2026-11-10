(function () {
  "use strict";

  var table = document.querySelector(".osv-table");
  if (table) {
    var rows = Array.prototype.slice.call(table.querySelectorAll("tr.osv-row"));
    var byId = {};
    rows.forEach(function (row) {
      byId[row.dataset.id] = row;
    });

    var isCollapsed = function (id) {
      var row = byId[id];
      return !!row && row.dataset.collapsed === "1";
    };

    var isVisible = function (id) {
      var row = byId[id];
      if (!row) return false;
      var parentId = row.dataset.parent;
      while (parentId) {
        if (isCollapsed(parentId)) return false;
        var parent = byId[parentId];
        parentId = parent ? parent.dataset.parent : "";
      }
      return true;
    };

    var refresh = function () {
      rows.forEach(function (row) {
        row.classList.toggle("is-hidden", !isVisible(row.dataset.id));
        var toggle = row.querySelector("[data-toggle]");
        if (toggle) {
          var collapsed = isCollapsed(row.dataset.id);
          toggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
        }
      });
    };

    table.addEventListener("click", function (event) {
      var toggle = event.target.closest("[data-toggle]");
      if (!toggle) return;
      var row = byId[toggle.dataset.toggle];
      if (!row) return;
      row.dataset.collapsed = row.dataset.collapsed === "1" ? "0" : "1";
      refresh();
    });

    refresh();
  }

  var form = document.getElementById("upload-form");
  if (form) {
    var input = document.getElementById("file");
    var fileList = document.getElementById("file-list");
    var submit = document.getElementById("submit-btn");
    var selected = [];

    var syncInput = function () {
      try {
        var transfer = new DataTransfer();
        selected.forEach(function (file) {
          transfer.items.add(file);
        });
        input.files = transfer.files;
      } catch (error) {
        /* Older browsers keep the file chosen through the file dialog. */
      }

      fileList.innerHTML = "";
      selected.forEach(function (file, index) {
        var item = document.createElement("div");
        item.className = "up-selected__item";

        var name = document.createElement("span");
        name.className = "up-selected__name";
        name.textContent = file.name;

        var size = document.createElement("span");
        size.className = "up-selected__size";
        size.textContent = (file.size / 1024 / 1024).toFixed(2) + " МБ";

        var remove = document.createElement("button");
        remove.type = "button";
        remove.className = "up-selected__remove";
        remove.setAttribute("aria-label", "Убрать файл");
        remove.textContent = "×";
        remove.addEventListener("click", function (event) {
          event.stopPropagation();
          selected.splice(index, 1);
          syncInput();
        });

        item.appendChild(name);
        item.appendChild(size);
        item.appendChild(remove);
        fileList.appendChild(item);
      });

      fileList.hidden = selected.length === 0;
      submit.disabled = selected.length === 0;
    };

    var addFiles = function (filesToAdd) {
      Array.prototype.forEach.call(filesToAdd, function (file) {
        var exists = selected.some(function (item) {
          return (
            item.name === file.name &&
            item.size === file.size &&
            item.lastModified === file.lastModified
          );
        });
        if (!exists) selected.push(file);
      });
      syncInput();
    };

    form.addEventListener("click", function (event) {
      if (event.target.closest("button")) return;
      if (event.target.closest(".up-selected")) return;
      input.click();
    });

    input.addEventListener("change", function () {
      if (input.files.length) addFiles(input.files);
    });

    ["dragenter", "dragover"].forEach(function (name) {
      form.addEventListener(name, function (event) {
        event.preventDefault();
        form.classList.add("is-drag");
      });
    });

    ["dragleave", "drop"].forEach(function (name) {
      form.addEventListener(name, function (event) {
        event.preventDefault();
        form.classList.remove("is-drag");
      });
    });

    form.addEventListener("drop", function (event) {
      if (event.dataTransfer && event.dataTransfer.files.length) {
        addFiles(event.dataTransfer.files);
      }
    });

    form.addEventListener("submit", function (event) {
      if (!selected.length) {
        event.preventDefault();
        return;
      }
      submit.disabled = true;
      submit.textContent = "Загружаем…";
    });

    syncInput();
  }

  var fotDetails = document.querySelector("details.fot-dashboard");
  if (fotDetails) {
    var syncFotDetails = function () {
      document.body.classList.toggle("fot-expanded", fotDetails.open);
    };
    fotDetails.addEventListener("toggle", syncFotDetails);
    syncFotDetails();
  }

  var mappingFile = document.getElementById("mapping-import-file");
  var mappingName = document.getElementById("mapping-import-name");
  if (mappingFile && mappingName) {
    mappingFile.addEventListener("change", function () {
      mappingName.textContent = mappingFile.files.length
        ? mappingFile.files[0].name
        : "Файл не выбран";
    });
    var mappingForm = mappingFile.closest("form");
    if (mappingForm) {
      mappingForm.addEventListener("submit", function (event) {
        if (!mappingFile.files.length) {
          event.preventDefault();
          mappingName.textContent = "Выберите файл";
        }
      });
    }
  }
  document.querySelectorAll("[data-select-all]").forEach(function (master) {
    var group = master.dataset.selectAll;
    master.addEventListener("change", function () {
      document
        .querySelectorAll('input.bulk-check[data-group="' + group + '"]')
        .forEach(function (checkbox) {
          if (master.dataset.visibleOnly) {
            var row = checkbox.closest("tr");
            if (row && row.hidden) return;
          }
          checkbox.checked = master.checked;
        });
    });
  });

  document.querySelectorAll("[data-bulk-apply]").forEach(function (button) {
    var group = button.dataset.bulkApply;
    button.addEventListener("click", function () {
      var bulkInput = document.querySelector(
        '[data-bulk-input="' + group + '"]'
      );
      if (!bulkInput) return;
      var value = bulkInput.value.trim();
      if (!value) {
        bulkInput.focus();
        return;
      }
      document
        .querySelectorAll('input.bulk-check[data-group="' + group + '"]')
        .forEach(function (checkbox) {
          if (!checkbox.checked) return;
          var row = checkbox.closest("tr");
          var target = row ? row.querySelector("input.bulk-target") : null;
          if (target) {
            target.value = value;
            target.dispatchEvent(new Event("input", { bubbles: true }));
          }
        });
    });
  });

  var fotSearch = document.getElementById("fot-search");
  var fotStatus = document.getElementById("fot-status");
  if (fotSearch && fotStatus) {
    var applyFotFilter = function () {
      var query = fotSearch.value.trim().toLowerCase();
      var status = fotStatus.value;
      document.querySelectorAll("tr.fot-row").forEach(function (row) {
        var matchesText =
          !query || (row.dataset.search || "").indexOf(query) !== -1;
        var matchesStatus = !status || row.dataset.status === status;
        row.hidden = !(matchesText && matchesStatus);
      });
    };
    fotSearch.addEventListener("input", applyFotFilter);
    fotStatus.addEventListener("change", applyFotFilter);
  }

  var fotForm = document.getElementById("fot-form");
  var fotDirty = document.getElementById("fot-dirty");
  if (fotForm && fotDirty) {
    fotForm.addEventListener("input", function (event) {
      fotDirty.hidden = false;
      var input = event.target;
      var row = input.closest ? input.closest("tr.fot-row") : null;
      if (row && input.classList.contains("cell-input")) {
        var department = row.cells[2] ? row.cells[2].textContent : "";
        row.dataset.search = (department + " " + input.value).toLowerCase();
      }
    });
    fotForm.addEventListener("change", function () {
      fotDirty.hidden = false;
    });
  }

  document.querySelectorAll("td.editable-cell").forEach(function (cell) {
    cell.addEventListener("click", function (event) {
      if (event.target.closest("input")) return;
      var input = cell.querySelector("input.cell-input");
      if (input) {
        input.focus();
        input.select();
      }
    });
  });

  // Mapping tables: dirty state, draft restore, unload guard, keyboard nav.

  var mappingSaving = false;

  window.addEventListener("beforeunload", function (event) {
    if (mappingSaving) return;
    if (document.querySelector("input.cell-input.is-dirty")) {
      event.preventDefault();
      event.returnValue = "";
    }
  });

  document.querySelectorAll("#unified-form, #fot-form").forEach(function (form) {
    var draftKey = "mappingDraft:" + form.id;
    var stateLabel = document.querySelector(
      '[data-save-state="' + form.id + '"]'
    );
    var buttons = Array.prototype.slice.call(
      document.querySelectorAll('button[form="' + form.id + '"]')
    ).concat(
      Array.prototype.slice.call(form.querySelectorAll('button[type="submit"]'))
    );
    buttons.forEach(function (button) {
      button.dataset.label = button.textContent;
    });

    var updateState = function () {
      var dirty = form.querySelectorAll("input.cell-input.is-dirty").length;
      if (stateLabel) {
        stateLabel.textContent = "Изменено: " + dirty;
        stateLabel.hidden = dirty === 0;
      }
      buttons.forEach(function (button) {
        button.disabled = dirty === 0;
        button.textContent =
          button.dataset.label + (dirty ? " (" + dirty + ")" : "");
      });
    };

    var draft = {};
    try {
      draft = JSON.parse(sessionStorage.getItem(draftKey) || "{}") || {};
    } catch (error) {
      draft = {};
    }
    var restored = false;
    form.querySelectorAll("input.cell-input").forEach(function (input) {
      if (Object.prototype.hasOwnProperty.call(draft, input.name)) {
        if (input.value !== draft[input.name]) restored = true;
        input.value = draft[input.name];
        input.classList.add("is-dirty");
      }
    });

    form.addEventListener("input", function (event) {
      var input = event.target;
      if (!input.classList || !input.classList.contains("cell-input")) return;
      input.classList.add("is-dirty");
      draft[input.name] = input.value;
      try {
        sessionStorage.setItem(draftKey, JSON.stringify(draft));
      } catch (error) {
        /* storage full or blocked — dirty marking still works */
      }
      updateState();
    });

    form.addEventListener("submit", function () {
      mappingSaving = true;
      try {
        sessionStorage.removeItem(draftKey);
      } catch (error) {
        /* nothing to clean up */
      }
    });

    updateState();
    if (restored && stateLabel) {
      stateLabel.textContent += " · восстановлены правки из черновика";
    }
  });

  document.addEventListener("keydown", function (event) {
    var input = event.target;
    if (!input.classList || !input.classList.contains("cell-input")) return;
    var direction = 0;
    if (event.key === "Enter" || event.key === "ArrowDown") direction = 1;
    else if (event.key === "ArrowUp") direction = -1;
    if (!direction) return;
    event.preventDefault();
    var row = input.closest("tr");
    while (row) {
      row = direction > 0 ? row.nextElementSibling : row.previousElementSibling;
      if (row && row.querySelector("input.cell-input")) break;
    }
    if (!row) return;
    var next = row.querySelector("input.cell-input");
    next.focus();
    next.select();
  });

  var importDialog = document.querySelector(".import-dialog");
  if (importDialog) {
    document.querySelectorAll("[data-import-open]").forEach(function (button) {
      button.addEventListener("click", function () {
        if (typeof importDialog.showModal === "function") {
          importDialog.showModal();
        }
      });
    });
    document.querySelectorAll("[data-import-close]").forEach(function (button) {
      button.addEventListener("click", function () {
        importDialog.close();
      });
    });
    importDialog.addEventListener("click", function (event) {
      if (event.target === importDialog) importDialog.close();
    });
  }

  document.querySelectorAll("form.import-form").forEach(function (form) {
    form.addEventListener("submit", function () {
      var fileInput = form.querySelector('input[type="file"]');
      if (!fileInput || !fileInput.files.length) return;
      var button = form.querySelector('button[type="submit"]');
      if (!button) return;
      button.disabled = true;
      button.dataset.label = button.dataset.label || button.textContent;
      button.textContent = "Обработка файла…";
      var status = document.createElement("span");
      status.className = "import-status";
      status.textContent = "Идёт обработка файла, страница обновится сама.";
      form.appendChild(status);
    });
  });

  document.querySelectorAll("form[data-auto-submit]").forEach(function (form) {
    form.querySelectorAll("select").forEach(function (select) {
      select.addEventListener("change", function () {
        form.submit();
      });
    });
    var search = form.querySelector('input[type="search"]');
    if (search) {
      var timer = null;
      search.addEventListener("input", function () {
        if (timer) clearTimeout(timer);
        timer = setTimeout(function () {
          form.submit();
        }, 500);
      });
    }
  });
})();
