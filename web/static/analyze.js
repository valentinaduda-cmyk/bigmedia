(function () {
  var form = document.querySelector("form[data-analyze-url]");
  if (!form) return;
  var url = form.getAttribute("data-analyze-url");
  var fileInput = form.querySelector('input[type="file"]');
  if (!fileInput) return;
  var banner = document.getElementById("analyze-warnings");

  function payload() {
    var fd = new FormData();
    for (var i = 0; i < fileInput.files.length; i++) fd.append("files", fileInput.files[i]);
    form.querySelectorAll("input[name], select[name]").forEach(function (el) {
      if (el.type === "file") return;
      if (el.type === "checkbox") { if (el.checked) fd.append(el.name, el.value || "on"); }
      else fd.append(el.name, el.value);
    });
    return fd;
  }

  function rebuildSelect(sel, options, suggested) {
    var want = options.indexOf(suggested) !== -1 ? suggested
             : (options.indexOf(sel.value) !== -1 ? sel.value : options[0]);
    sel.innerHTML = "";
    options.forEach(function (opt) {
      var o = document.createElement("option");
      o.value = opt;
      o.textContent = opt;
      if (opt === want) o.selected = true;
      sel.appendChild(o);
    });
  }

  function apply(data) {
    var suggestions = data.suggestions || {};
    var sheetNames = (data.sheets || []).map(function (s) { return s.name; });

    if ((data.headers || []).length) {
      form.querySelectorAll('select[data-source="headers"]').forEach(function (sel) {
        rebuildSelect(sel, data.headers, suggestions[sel.name]);
      });
    }
    if (sheetNames.length) {
      form.querySelectorAll('select[data-source="sheets"]').forEach(function (sel) {
        rebuildSelect(sel, sheetNames, suggestions[sel.name]);
      });
    }

    var cats = suggestions.categories;
    var counts = (data.annotations || {}).categories || {};
    form.querySelectorAll('input[name="categories"]').forEach(function (box) {
      if (cats) box.checked = cats.indexOf(box.value) !== -1;
      var n = counts[box.value];
      var span = box.parentElement.querySelector(".category-name");
      if (span && typeof n === "number") {
        span.textContent = box.value + " (" + n + (n === 1 ? " clip" : " clips") + ")";
      }
    });

    Object.keys(suggestions).forEach(function (key) {
      if (key === "categories") return;
      var el = form.querySelector('[name="' + key + '"]');
      if (el && el.tagName !== "SELECT" && el.type !== "checkbox") el.value = suggestions[key];
    });

    if (banner) {
      var msgs = data.warnings || [];
      banner.textContent = msgs.join("  •  ");
      banner.hidden = msgs.length === 0;
    }
  }

  function analyze() {
    if (!fileInput.files.length) return;
    fetch(url, { method: "POST", body: payload() })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (data) { if (data) apply(data); })
      .catch(function () { console.warn("analyze request failed"); });
  }

  fileInput.addEventListener("change", analyze);
  form.querySelectorAll("select[data-analyze]").forEach(function (sel) {
    sel.addEventListener("change", analyze);
  });
})();
