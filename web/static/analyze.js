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

  function setBanner(msgs) {
    if (!banner) return;
    banner.textContent = (msgs || []).join("  •  ");
    banner.hidden = !msgs || msgs.length === 0;
  }

  // The plain-text label for a field, for warning messages. Falls back to
  // the input name when there's no wrapping <label>.
  function fieldLabel(el) {
    var lab = el.closest && el.closest("label");
    if (lab) {
      var parts = [];
      lab.childNodes.forEach(function (n) { if (n.nodeType === 3) parts.push(n.textContent); });
      var t = parts.join(" ").replace(/\s+/g, " ").trim();
      if (t) return t;
    }
    return el.name;
  }

  // Rebuild a <select>'s options from `options`, keeping a valid current
  // selection or honoring a server suggestion. When neither is possible the
  // value is silently substituted to options[0]; push a note into `subs` so
  // apply() can surface it in the warning banner (spec decision 6: warn,
  // fall back, keep Run enabled).
  function rebuildSelect(sel, options, suggested, subs) {
    var prev = sel.value;
    var want;
    if (options.indexOf(suggested) !== -1) {
      want = suggested;
    } else if (options.indexOf(prev) !== -1) {
      want = prev;
    } else {
      want = options[0];
      if (prev && prev !== want && subs) {
        subs.push(fieldLabel(sel) + " '" + prev + "' not in this file — using '" + want + "'");
      }
    }
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
    var subs = [];

    if ((data.headers || []).length) {
      form.querySelectorAll('select[data-source="headers"]').forEach(function (sel) {
        rebuildSelect(sel, data.headers, suggestions[sel.name], subs);
      });
    }
    if (sheetNames.length) {
      form.querySelectorAll('select[data-source="sheets"]').forEach(function (sel) {
        rebuildSelect(sel, sheetNames, suggestions[sel.name], subs);
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

    setBanner((data.warnings || []).concat(subs));
  }

  // isRerun guards against an infinite loop: a programmatic <select> rebuild
  // changes the selection without firing a `change` event, so the counts we
  // just applied were computed for the value we *sent*, not the one now
  // shown. Re-run analysis exactly once in that case.
  function analyze(isRerun) {
    if (!fileInput.files.length) return;
    var sent = {};
    form.querySelectorAll("select[data-analyze]").forEach(function (sel) { sent[sel.name] = sel.value; });

    fetch(url, { method: "POST", body: payload() })
      .then(function (r) {
        if (!r.ok) {
          console.warn("analyze request failed: HTTP " + r.status);
          setBanner(["Could not analyze the file — using defaults."]);
          return null;
        }
        return r.json();
      })
      .then(function (data) {
        if (!data) return;
        apply(data);
        if (isRerun) return;
        var changed = false;
        form.querySelectorAll("select[data-analyze]").forEach(function (sel) {
          if (sent[sel.name] !== undefined && sent[sel.name] !== sel.value) changed = true;
        });
        if (changed) analyze(true);
      })
      .catch(function () {
        console.warn("analyze request failed");
        setBanner(["Could not analyze the file — using defaults."]);
      });
  }

  fileInput.addEventListener("change", function () { analyze(false); });
  form.querySelectorAll("select[data-analyze]").forEach(function (sel) {
    sel.addEventListener("change", function () { analyze(false); });
  });
})();
