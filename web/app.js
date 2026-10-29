// 월별 개강 업무 도구 - 웹 버전.
//
// 처리 로직은 exe와 똑같이 scripts/*.py 를 그대로 쓰고, 브라우저 안의 Python(Pyodide)에서 돌린다.
// 고른 파일은 이 브라우저 안에서만 읽고 쓰며 어떤 서버로도 보내지 않는다(CDN에서는 Pyodide·pandas·openpyxl
// 프로그램만 받아온다). 파일 선택·폴더 저장·최근 파일 이력은 Chrome/Edge의 File System Access API를 쓴다.
(function () {
  const PY_MODULES = [
    "pricing", "process_apply_list", "generate_company_files", "build_gaegang_sheet",
    "tel_list", "add_applications", "run_steps", "added_mapping",
  ];
  // openpyxl은 Pyodide 기본 패키지에 없어서 PyPI에서 버전을 고정해 받는다.
  const PYPI_WHEELS = [
    "https://files.pythonhosted.org/packages/c1/8b/5fe2cc11fee489817272089c4203e679c63b570a5aaeb18d852ae3cbba6a/et_xmlfile-2.0.0-py3-none-any.whl",
    "https://files.pythonhosted.org/packages/c0/da/977ded879c29cbd04de313843e76868e6e13408a94ed6b987245dc7c8506/openpyxl-3.1.5-py2.py3-none-any.whl",
  ];
  const MAX_RECENT_FILES = 3;
  const FILE_TYPES = [{
    description: "엑셀 CSV 파일",
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "application/vnd.ms-excel": [".xls"],
      "text/csv": [".csv"],
    },
  }];

  const state = {}; // slot -> FileSystemFileHandle
  const rows = {};
  let pyodide = null;

  // ----- 최근 파일 이력 (IndexedDB, 칸마다 따로, 파일 핸들만 저장하고 내용은 저장하지 않음) -----
  function openDb() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open("월별개강업무도구", 1);
      req.onupgradeneeded = () => req.result.createObjectStore("recent");
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function getRecentFiles(slot) {
    try {
      const db = await openDb();
      return await new Promise((resolve, reject) => {
        const req = db.transaction("recent").objectStore("recent").get(slot);
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => reject(req.error);
      });
    } catch (e) {
      return [];
    }
  }

  async function rememberRecentFile(slot, handle) {
    try {
      const kept = [];
      for (const h of await getRecentFiles(slot)) {
        if (!(await h.isSameEntry(handle))) kept.push(h);
      }
      const handles = [handle, ...kept].slice(0, MAX_RECENT_FILES);
      const db = await openDb();
      await new Promise((resolve, reject) => {
        const tx = db.transaction("recent", "readwrite");
        tx.objectStore("recent").put(handles, slot);
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
    } catch (e) {
      // 이력 저장 실패는 작업에 영향이 없으므로 무시한다.
    }
  }

  async function ensurePermission(handle, mode) {
    if ((await handle.queryPermission({ mode })) === "granted") return true;
    return (await handle.requestPermission({ mode })) === "granted";
  }

  // ----- 파일 선택 줄 -----
  function setFieldValue(slot, handle) {
    state[slot] = handle || null;
    rows[slot].input.value = handle ? handle.name : "";
    if (slot === "added") onAddedFileChanged(handle);
    if (["apply", "refund", "general"].includes(slot) && handle) onBaseFileChanged();
  }

  function closeAllRecentMenus() {
    Object.values(rows).forEach(({ recentMenu }) => recentMenu.classList.add("hidden"));
  }

  function renderRecentMenu(slot, handles) {
    const { recentMenu } = rows[slot];
    recentMenu.innerHTML = "";
    if (!handles || handles.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.textContent = "최근 선택한 파일 없음";
      recentMenu.appendChild(empty);
      return;
    }
    handles.forEach((handle) => {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "item";
      item.textContent = handle.name;
      item.title = handle.name;
      item.addEventListener("click", async (ev) => {
        ev.stopPropagation();
        recentMenu.classList.add("hidden");
        try {
          if (!(await ensurePermission(handle, "read"))) return;
          await handle.getFile();
        } catch (e) {
          showModal("알려드립니다", `파일을 열 수 없습니다. 옮겨졌거나 지워졌을 수 있습니다.\n${handle.name}`);
          return;
        }
        await rememberRecentFile(slot, handle);
        setFieldValue(slot, handle);
      });
      recentMenu.appendChild(item);
    });
  }

  document.querySelectorAll(".file-row").forEach((row) => {
    const slot = row.dataset.slot;
    const input = row.querySelector("input[type=text]");
    const pickBtn = row.querySelector(".btn.pick");
    const recentBtn = row.querySelector(".btn.recent");
    const recentMenu = row.querySelector(".recent-menu");
    rows[slot] = { row, input, pickBtn, recentBtn, recentMenu };
    state[slot] = null;

    pickBtn.addEventListener("click", async () => {
      let handles;
      try {
        handles = await window.showOpenFilePicker({ types: FILE_TYPES });
      } catch (e) {
        return; // 취소
      }
      await rememberRecentFile(slot, handles[0]);
      setFieldValue(slot, handles[0]);
    });

    recentBtn.addEventListener("click", async (ev) => {
      ev.stopPropagation();
      const wasOpen = !recentMenu.classList.contains("hidden");
      closeAllRecentMenus();
      if (wasOpen) return;
      renderRecentMenu(slot, await getRecentFiles(slot));
      recentMenu.classList.remove("hidden");
    });
  });

  document.addEventListener("click", () => closeAllRecentMenus());

  function clearFileFields() {
    Object.keys(rows).forEach((slot) => setFieldValue(slot, null));
    pasteText.value = "";
    resetMapping();
    document.getElementById("date-start").value = "";
    document.getElementById("date-end").value = "";
  }

  const logEl = document.getElementById("log");
  function appendLog(text) {
    logEl.textContent += (logEl.textContent ? "\n" : "") + text;
    logEl.scrollTop = logEl.scrollHeight;
  }
  function clearLog() {
    logEl.textContent = "";
  }

  // Python 로그에 나오는 가상 파일 시스템 입력 경로(/work/in/<칸>/)는 빼고 파일 이름만 보여준다.
  function showPythonLogs(logs) {
    logs.forEach((text) => appendLog(text.replace(/\/work\/in\/[^/]+\//g, "")));
  }

  const modalOverlay = document.getElementById("modal-overlay");
  const modalTitle = document.getElementById("modal-title");
  const modalMessage = document.getElementById("modal-message");
  const modalCancel = document.getElementById("modal-cancel");
  let modalResolve = null;
  function closeModal(answer) {
    modalOverlay.classList.add("hidden");
    if (modalResolve) modalResolve(answer);
    modalResolve = null;
  }
  document.getElementById("modal-ok").addEventListener("click", () => closeModal(true));
  modalCancel.addEventListener("click", () => closeModal(false));

  function showModal(title, message) {
    modalTitle.textContent = title;
    modalMessage.textContent = message;
    modalCancel.classList.add("hidden");
    modalOverlay.classList.remove("hidden");
  }

  // 확인/취소를 고르게 하고, 확인이면 true를 돌려준다.
  function showConfirm(title, message) {
    showModal(title, message);
    modalCancel.classList.remove("hidden");
    return new Promise((resolve) => { modalResolve = resolve; });
  }

  // ----- Pyodide 가상 파일 시스템 <-> 실제 파일 -----
  // 입력: /work/in/<칸>/<파일이름>, 결과: /work/out/ 아래에 만든 뒤 고른 폴더로 그대로 복사한다.
  async function copyInputs(slots) {
    pyodide.runPython("import shutil, os\nshutil.rmtree('/work', ignore_errors=True)\nos.makedirs('/work/out')");
    const paths = {};
    for (const slot of slots) {
      const handle = state[slot];
      if (!handle) continue;
      const file = await handle.getFile();
      const dir = `/work/in/${slot}`;
      pyodide.FS.mkdirTree(dir);
      paths[slot] = `${dir}/${file.name}`;
      pyodide.FS.writeFile(paths[slot], new Uint8Array(await file.arrayBuffer()));
    }
    return paths;
  }

  async function writeFileTo(fileHandle, bytes) {
    const writable = await fileHandle.createWritable();
    await writable.write(bytes);
    await writable.close();
  }

  async function copyOutputs(fsDir, dirHandle) {
    for (const name of pyodide.FS.readdir(fsDir)) {
      if (name === "." || name === "..") continue;
      const fsPath = `${fsDir}/${name}`;
      if (pyodide.FS.isDir(pyodide.FS.stat(fsPath).mode)) {
        await copyOutputs(fsPath, await dirHandle.getDirectoryHandle(name, { create: true }));
      } else {
        await writeFileTo(await dirHandle.getFileHandle(name, { create: true }), pyodide.FS.readFile(fsPath));
      }
    }
  }

  // 화면 갱신(로그·안내)이 먼저 그려지도록 한 박자 쉰 뒤 Python을 돌린다.
  function nextFrame() {
    return new Promise((resolve) => setTimeout(resolve, 50));
  }

  // ----- ⑤ 추가 신청 칸 맞추기 (scripts/added_mapping.py) -----
  // 서식이 제각각인 추가 신청(파일 또는 붙여넣은 표)에서 시트·칸 이름 줄을 고르고, 원본 신청명단 칸마다 쓸 칸을 고른다.
  // 읽은 내용은 Pyodide 메모리(_map_sheets)에만 두고, 파일·붙여넣은 내용을 바꾸거나 화면을 비우면 지운다.
  const mapEl = document.getElementById("mapping");
  const mapStatus = document.getElementById("map-status");
  const mapBody = document.getElementById("map-body");
  const mapSheet = document.getElementById("map-sheet");
  const mapHeader = document.getElementById("map-header");
  const mapPreview = document.getElementById("map-preview");
  const mapFields = document.getElementById("map-fields");
  const pasteBox = document.getElementById("paste-box");
  const pasteText = document.getElementById("paste-text");
  // source: 칸 맞추기 화면이 어떤 내용을 읽었는지("file" | "paste"), 아직 없으면 null
  const mapState = {
    ready: false, source: null, sheet: null, headerRow: 1, columns: [], mapping: {}, auto: new Set(), fixedText: {},
  };
  let addedMode = "file"; // ⑤ 추가 신청을 넣는 방법: "file" | "paste"
  let mapColumns = null; // { all: 원본 신청명단 칸, required: 꼭 맞춰야 하는 칸 }

  function pyToJs(proxy) {
    const value = proxy.toJs({ dict_converter: Object.fromEntries });
    if (proxy.destroy) proxy.destroy();
    return value;
  }

  function missingRequiredColumns() {
    if (!mapColumns) return [];
    return mapColumns.required.filter((c) => mapState.mapping[c] === undefined)
      .map((c) => (mapState.fixedText[c] !== undefined ? `${c}(직접 입력한 날짜를 확인해주세요)` : c));
  }

  function resetMapping() {
    clearReview();
    mapState.ready = false;
    mapState.source = null;
    if (pyodide) pyodide.runPython("globals().pop('_map_sheets', None)");
    mapEl.classList.add("hidden");
  }

  // readSheets: _map_sheets를 채우는 함수. 읽는 사이에 다른 내용으로 바뀌었으면 false를 돌려준다.
  async function loadMapping(source, readSheets, failMessage) {
    mapEl.classList.remove("hidden");
    mapBody.classList.add("hidden");
    mapStatus.textContent = "읽는 중입니다...";
    await nextFrame();
    try {
      if (!(await readSheets())) return;
      const sheets = pyToJs(pyodide.runPython("list(_map_sheets)"));
      mapSheet.innerHTML = "";
      sheets.forEach((name) => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        mapSheet.appendChild(option);
      });
      mapState.sheet = sheets[0];
      mapState.headerRow = guessHeaderRow();
      mapHeader.value = mapState.headerRow;
      refreshMapping();
      mapStatus.textContent = "";
      mapBody.classList.remove("hidden");
      mapState.source = source;
      mapState.ready = true;
    } catch (e) {
      mapStatus.textContent = failMessage + "\n" + e;
    }
  }

  async function onAddedFileChanged(handle) {
    resetMapping();
    if (!handle) return;
    if (!pyodide) {
      mapEl.classList.remove("hidden");
      mapBody.classList.add("hidden");
      mapStatus.textContent = "처리 도구를 불러오는 중입니다. 준비되면 칸 맞추기 화면이 나옵니다.";
      return;
    }
    await loadMapping("file", async () => {
      const file = await handle.getFile();
      const path = `/map/${file.name}`;
      pyodide.FS.mkdirTree("/map");
      pyodide.FS.writeFile(path, new Uint8Array(await file.arrayBuffer()));
      try {
        pyodide.globals.set("_map_path", path);
        pyodide.runPython("import added_mapping\n_map_sheets = added_mapping.read_sheets(_map_path)");
      } finally {
        pyodide.FS.unlink(path);
      }
      return state.added === handle;
    }, "파일을 읽지 못했습니다. 엑셀(xlsx) 또는 CSV 파일인지 확인해주세요.");
  }

  // ⑤ 넣는 방법(파일 / 붙여넣기)을 바꾸면 다른 쪽에 넣어 둔 내용은 비운다(둘을 섞어 쓰지 않음).
  document.querySelectorAll(".source-toggle .toggle").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (btn.dataset.mode === addedMode) return;
      addedMode = btn.dataset.mode;
      document.querySelectorAll(".source-toggle .toggle").forEach((b) => b.classList.toggle("active", b === btn));
      rows.added.row.classList.toggle("hidden", addedMode !== "file");
      pasteBox.classList.toggle("hidden", addedMode !== "paste");
      setFieldValue("added", null);
      pasteText.value = "";
      resetMapping();
    });
  });

  document.getElementById("paste-read").addEventListener("click", async () => {
    const text = pasteText.value;
    if (!text.trim()) {
      showModal("알려드립니다", "붙여넣은 내용이 없습니다. 엑셀·메일의 표를 복사해서 붙여넣어 주세요.");
      return;
    }
    if (!text.includes("\t")) {
      showModal("알려드립니다", "붙여넣은 내용이 표 형식이 아닙니다(칸 구분이 없음).\n"
        + "엑셀이나 메일의 표를 칸 이름 줄과 함께 복사해서 붙여넣어 주세요. "
        + "메일 본문 문장·카톡 같은 글은 다음 단계(브라우저 안 LLM)에서 지원할 예정입니다.");
      return;
    }
    if (!pyodide) {
      showModal("잠시만요", "처리 도구를 불러오는 중입니다. 잠시 후 다시 눌러주세요.");
      return;
    }
    resetMapping();
    await loadMapping("paste", async () => {
      pyodide.globals.set("_map_text", text);
      try {
        pyodide.runPython("import added_mapping\n_map_sheets = added_mapping.read_pasted(_map_text)");
      } finally {
        pyodide.runPython("globals().pop('_map_text', None)");
      }
      return pasteText.value === text;
    }, "붙여넣은 표를 읽지 못했습니다.");
  });

  // 표를 읽은 뒤 붙여넣은 내용을 고치면 칸 맞추기를 다시 해야 한다.
  pasteText.addEventListener("input", () => {
    if (mapState.source === "paste") resetMapping();
  });

  // 위쪽 줄 중 원본 칸 이름과 가장 많이 맞는 줄을 칸 이름 줄로 골라 둔다(사용자가 숫자로 바꿀 수 있음).
  function guessHeaderRow() {
    pyodide.globals.set("_map_sheet", mapState.sheet);
    return pyodide.runPython("added_mapping.guess_header_row(_map_sheets, _map_sheet)");
  }

  // 시트·칸 이름 줄이 바뀔 때마다 미리보기와 칸 목록을 다시 만들고, 칸 고르기는 자동 추천으로 되돌린다.
  function refreshMapping() {
    pyodide.globals.set("_map_sheet", mapState.sheet);
    pyodide.globals.set("_map_header", mapState.headerRow);
    const info = pyToJs(pyodide.runPython("added_mapping.sheet_info(_map_sheets, _map_sheet, _map_header)"));
    mapHeader.max = Math.max(1, info.row_count);
    mapState.columns = info.columns;
    mapState.mapping = { ...info.suggested };
    mapState.auto = new Set(Object.keys(info.suggested)); // 자동으로 고른 칸(사용자가 바꾸면 빠짐)
    mapState.fixedText = {}; // "직접 입력"을 고른 칸 -> 입력한 글자
    renderPreview(info.first_rows);
    renderFields();
  }

  function renderPreview(firstRows) {
    mapPreview.innerHTML = "";
    firstRows.forEach((cells, i) => {
      const tr = document.createElement("tr");
      if (i + 1 === mapState.headerRow) tr.className = "header-row";
      const no = document.createElement("td");
      no.className = "rowno";
      no.textContent = i + 1;
      tr.appendChild(no);
      cells.forEach((text) => {
        const td = document.createElement("td");
        td.textContent = text;
        td.title = text;
        tr.appendChild(td);
      });
      mapPreview.appendChild(tr);
    });
  }

  // 교육시작일·교육종료일은 파일에 칸이 없으면 "직접 입력"으로 날짜를 정해 모든 줄에 같이 넣을 수 있다.
  const FIXED_DATE_FIELDS = ["교육시작일", "교육종료일"];
  const FIXED = "fixed";

  function renderFields() {
    mapFields.innerHTML = "";
    mapColumns.all.forEach((target) => {
      const required = mapColumns.required.includes(target);
      const canFix = FIXED_DATE_FIELDS.includes(target);
      const field = document.createElement("label");
      field.className = "mapping-field";
      const name = document.createElement("span");
      name.textContent = required ? `${target} *` : target;
      const select = document.createElement("select");
      const none = document.createElement("option");
      none.value = "";
      none.textContent = required ? "(골라주세요)" : "(없음 - 비워 둠)";
      select.appendChild(none);
      if (canFix) {
        const fixed = document.createElement("option");
        fixed.value = FIXED;
        fixed.textContent = "직접 입력 (모든 줄에 같은 날짜)";
        select.appendChild(fixed);
      }
      mapState.columns.forEach((col) => {
        const option = document.createElement("option");
        option.value = String(col.index);
        option.textContent = col.label;
        select.appendChild(option);
      });
      const autoTag = document.createElement("span");
      autoTag.className = "auto-tag";
      autoTag.textContent = "자동";
      autoTag.title = "칸 이름이 비슷해서 자동으로 골랐습니다. 맞는지 확인해주세요.";
      name.appendChild(autoTag);

      const fixedBox = document.createElement("div");
      fixedBox.className = "fixed-date hidden";
      const fixedInput = document.createElement("input");
      fixedInput.type = "text";
      fixedInput.inputMode = "numeric";
      fixedInput.placeholder = "예: 2026-11-02";
      fixedInput.autocomplete = "off";
      fixedInput.value = mapState.fixedText[target] || "";
      fixedBox.appendChild(fixedInput);
      if (canFix) attachCalendar(fixedInput);

      const currentValue = () => (mapState.fixedText[target] !== undefined ? FIXED
        : mapState.mapping[target] === undefined ? "" : String(mapState.mapping[target]));
      const update = () => {
        field.classList.toggle("missing", required && mapState.mapping[target] === undefined);
        autoTag.classList.toggle("hidden", !mapState.auto.has(target));
        fixedBox.classList.toggle("hidden", mapState.fixedText[target] === undefined);
        fixedInput.classList.toggle("invalid", Boolean(fixedInput.value.trim()) && !parseDateText(fixedInput.value));
      };
      // 직접 입력한 날짜는 읽을 수 있을 때만 칸 맞추기 값이 된다(못 읽으면 아직 안 고른 것으로 봄).
      const applyFixed = () => {
        mapState.fixedText[target] = fixedInput.value;
        const parsed = parseDateText(fixedInput.value);
        if (parsed) {
          fixedInput.value = parsed;
          mapState.fixedText[target] = parsed;
          mapState.mapping[target] = parsed;
        } else {
          delete mapState.mapping[target];
        }
      };
      select.value = currentValue();
      select.addEventListener("change", async () => {
        const chosen = select.value;
        select.value = currentValue();
        if (!(await confirmDiscardReview())) return;
        select.value = chosen;
        mapState.auto.delete(target);
        delete mapState.fixedText[target];
        delete mapState.mapping[target];
        if (chosen === FIXED) {
          applyFixed();
        } else if (chosen !== "") {
          mapState.mapping[target] = Number(chosen);
        }
        update();
        if (chosen === FIXED) fixedInput.focus();
      });
      fixedInput.addEventListener("input", update);
      fixedInput.addEventListener("change", async () => {
        if (!(await confirmDiscardReview())) {
          fixedInput.value = mapState.fixedText[target] || "";
          update();
          return;
        }
        applyFixed();
        update();
      });
      update();
      field.append(name, select, fixedBox);
      mapFields.appendChild(field);
    });
  }


  mapSheet.addEventListener("change", async () => {
    const chosen = mapSheet.value;
    mapSheet.value = mapState.sheet;
    if (!(await confirmDiscardReview())) return;
    mapSheet.value = chosen;
    mapState.sheet = mapSheet.value;
    mapState.headerRow = guessHeaderRow();
    mapHeader.value = mapState.headerRow;
    refreshMapping();
  });

  mapHeader.addEventListener("change", async () => {
    const row = parseInt(mapHeader.value, 10);
    mapHeader.value = mapState.headerRow;
    if (!(row >= 1 && row <= Number(mapHeader.max))) return;
    if (!(await confirmDiscardReview())) return;
    mapHeader.value = row;
    mapState.headerRow = row;
    refreshMapping();
  });

  // ----- 확인 표: 칸 맞추기 결과를 반영 전에 보여 주고 고치게 한다 (scripts/added_mapping.py) -----
  // 검사는 기존 규칙(필수 칸·날짜·커리큘럼 과정명·중복·제외 규칙)을 미리 보여 주는 것뿐이다.
  // 줄마다 values(지금 값)와 original(확인 표를 만들 때 값, 직접 추가한 줄은 null),
  // auto(비슷한 과정명으로 자동 추정한 칸 -> 파일에 적혀 있던 값)를 들고 있다.
  const reviewEl = document.getElementById("review");
  const reviewTable = document.getElementById("review-table");
  const reviewSummary = document.getElementById("review-summary");
  const courseList = document.getElementById("course-list");
  const reviewState = { built: false, dirty: false, columns: [], rows: [], deleted: 0, checks: [] };
  let recheckTimer = null;

  function clearReview() {
    clearTimeout(recheckTimer);
    Object.assign(reviewState, { built: false, dirty: false, columns: [], rows: [], deleted: 0, checks: [] });
    reviewEl.classList.add("hidden");
    reviewTable.innerHTML = "";
    if (pyodide) pyodide.runPython("globals().pop('_review_ctx', None)");
  }

  // 확인 표를 고친 뒤 칸 맞추기를 바꾸면 고친 내용이 사라지므로 먼저 물어본다.
  async function confirmDiscardReview() {
    if (reviewState.built && reviewState.dirty
        && !(await showConfirm("알려드립니다", "확인 표에서 고친 내용이 사라집니다. 계속할까요?"))) {
      return false;
    }
    clearReview();
    return true;
  }

  // ①②③ 파일로 검사 기준(원본 아이디·과정명, 커리큘럼 과정명)을 읽는다.
  async function loadReviewContext() {
    const paths = {};
    pyodide.FS.mkdirTree("/review");
    try {
      for (const slot of ["apply", "refund", "general"]) {
        const file = await state[slot].getFile();
        paths[slot] = `/review/${slot}_${file.name}`;
        pyodide.FS.writeFile(paths[slot], new Uint8Array(await file.arrayBuffer()));
      }
      pyodide.globals.set("_review_paths", pyodide.toPy(paths));
      pyodide.runPython("_review_ctx = added_mapping.load_check_context("
        + "_review_paths['apply'], _review_paths['refund'], _review_paths['general'])\n"
        + "globals().pop('_review_paths', None)");
    } finally {
      Object.values(paths).forEach((path) => pyodide.FS.unlink(path));
    }
    courseList.innerHTML = "";
    pyToJs(pyodide.runPython("added_mapping.course_names(_review_ctx)")).forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      courseList.appendChild(option);
    });
  }

  document.getElementById("review-build").addEventListener("click", async () => {
    if (!(state.apply && state.refund && state.general)) {
      showModal("알려드립니다", "확인 표는 ① 신청명단 / ② 환급과정 커리큘럼 / ③ 비환급과정 커리큘럼과 비교해 검사합니다. "
        + "①②③ 파일을 먼저 골라주세요.");
      return;
    }
    const missing = missingRequiredColumns();
    if (missing.length) {
      showModal("알려드립니다", `칸 맞추기에서 다음 칸을 골라주세요: ${missing.join(", ")}`);
      return;
    }
    if (!(await confirmDiscardReview())) return;
    try {
      await loadReviewContext();
      pyodide.globals.set("_map_mapping", pyodide.toPy(mapState.mapping));
      const rows = pyToJs(pyodide.runPython(
        "added_mapping.build_review_rows(_map_sheets, _map_sheet, _map_header, _map_mapping, _review_ctx)"));
      reviewState.columns = pyToJs(pyodide.runPython("added_mapping.review_columns(_map_mapping)"));
      // auto: 비슷한 과정명으로 자동 추정한 칸 -> 파일에 적혀 있던 값
      reviewState.rows = rows.map(({ values, auto }) => ({ values, original: { ...values }, auto }));
      reviewState.built = true;
      renderReview();
    } catch (e) {
      clearReview();
      showModal("알려드립니다", "확인 표를 만들지 못했습니다.\n" + e);
    }
  });

  document.getElementById("review-add").addEventListener("click", () => {
    reviewState.rows.push({ values: Object.fromEntries(reviewState.columns.map((c) => [c, ""])), original: null, auto: {} });
    reviewState.dirty = true;
    renderReview();
    const inputs = reviewTable.querySelectorAll("tr:last-child input");
    if (inputs.length) inputs[0].focus();
  });

  function renderReview() {
    reviewTable.innerHTML = "";
    const head = document.createElement("tr");
    ["#", ...reviewState.columns, "확인 내용", ""].forEach((name) => {
      const th = document.createElement("th");
      th.textContent = name;
      head.appendChild(th);
    });
    reviewTable.appendChild(head);
    reviewState.rows.forEach((row, i) => {
      const tr = document.createElement("tr");
      const no = document.createElement("td");
      no.className = "rowno";
      no.textContent = i + 1;
      tr.appendChild(no);
      reviewState.columns.forEach((col) => {
        const td = document.createElement("td");
        td.dataset.col = col;
        const input = document.createElement("input");
        input.type = "text";
        input.value = row.values[col];
        // 신청과정명은 직접 고치거나 커리큘럼 과정명 목록에서 골라 넣는다(자동으로 바꾸지 않음).
        if (col === "신청과정명") input.setAttribute("list", "course-list");
        input.addEventListener("input", () => {
          row.values[col] = input.value;
          reviewState.dirty = true;
          clearTimeout(recheckTimer);
          recheckTimer = setTimeout(recheckReview, 250);
        });
        td.appendChild(input);
        tr.appendChild(td);
      });
      const messages = document.createElement("td");
      messages.className = "messages";
      tr.appendChild(messages);
      const del = document.createElement("td");
      const delBtn = document.createElement("button");
      delBtn.type = "button";
      delBtn.className = "delete";
      delBtn.title = "이 줄 지우기";
      delBtn.textContent = "✕";
      delBtn.addEventListener("click", () => {
        reviewState.rows.splice(reviewState.rows.indexOf(row), 1);
        if (row.original) reviewState.deleted += 1;
        reviewState.dirty = true;
        renderReview();
      });
      del.appendChild(delBtn);
      tr.appendChild(del);
      reviewTable.appendChild(tr);
    });
    reviewEl.classList.remove("hidden");
    recheckReview();
  }

  // 지금 표 내용으로 다시 검사해 노란 칸·확인 내용·요약을 바꾼다(입력 칸은 그대로 둬서 고치던 자리를 유지).
  function recheckReview() {
    clearTimeout(recheckTimer);
    if (!reviewState.built) return;
    const values = pyodide.toPy(reviewState.rows);
    try {
      pyodide.globals.set("_review_rows", values);
      reviewState.checks = pyToJs(pyodide.runPython(
        "_checks = added_mapping.check_review_rows(_review_rows, _review_ctx)\n"
        + "globals().pop('_review_rows', None)\n_checks"));
    } finally {
      values.destroy();
    }
    const trs = reviewTable.querySelectorAll("tr");
    reviewState.checks.forEach((check, i) => {
      const tr = trs[i + 1];
      tr.querySelectorAll("td[data-col]").forEach((td) => td.classList.toggle("warn", check.cells.includes(td.dataset.col)));
      tr.querySelector("td.messages").textContent = check.messages.join("\n");
    });
    const flagged = reviewState.checks.filter((c) => c.messages.length).length;
    reviewSummary.textContent = `${reviewState.rows.length}건 · 확인필요 ${flagged}건`;
  }

  // ①②③을 바꾸면 검사 기준이 바뀌므로 다시 읽어서 검사한다.
  async function onBaseFileChanged() {
    if (!reviewState.built) return;
    if (!(state.apply && state.refund && state.general)) return;
    try {
      await loadReviewContext();
      recheckReview();
    } catch (e) {
      showModal("알려드립니다", "확인 표 검사 기준(①②③ 파일)을 다시 읽지 못했습니다.\n" + e);
    }
  }

  // ----- 실행 -----
  const saveTelCheckbox = document.getElementById("save-tel");
  const dateStartInput = document.getElementById("date-start");
  const dateEndInput = document.getElementById("date-end");
  const runAllBtn = document.getElementById("run-all");
  const runAddedBtn = document.getElementById("run-added");

  function setButtonsDisabled(disabled) {
    runAllBtn.disabled = disabled;
    runAddedBtn.disabled = disabled;
  }

  // 교육시작일 기간은 키보드로 입력한다: 2026-11-01 / 2026.11.01 / 2026/11/1 / 20261101.
  // 읽으면 "YYYY-MM-DD"를, 비어 있으면 ""를, 읽지 못하면 null을 돌려준다.
  function parseDateText(text) {
    const value = text.trim();
    if (!value) return "";
    const match = value.match(/^(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})\.?$/) || value.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (!match) return null;
    const [y, m, d] = match.slice(1).map(Number);
    const date = new Date(y, m - 1, d);
    if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) return null;
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }

  // 날짜 글자 칸 옆에 달력 버튼을 붙인다. 달력에서 고르면 2026-11-02 모양으로 글자 칸에 들어간다.
  function attachCalendar(textInput) {
    const picker = document.createElement("input");
    picker.type = "date";
    picker.className = "calendar-picker";
    picker.tabIndex = -1;
    picker.setAttribute("aria-hidden", "true");
    const button = document.createElement("button");
    button.type = "button";
    button.className = "calendar-btn";
    button.title = "달력에서 고르기";
    button.setAttribute("aria-label", "달력에서 고르기");
    button.textContent = "📅";
    button.addEventListener("click", (ev) => {
      ev.preventDefault();
      picker.value = parseDateText(textInput.value) || "";
      try {
        picker.showPicker();
      } catch (e) {
        picker.focus();
      }
    });
    picker.addEventListener("change", () => {
      if (!picker.value) return;
      textInput.value = picker.value;
      textInput.dispatchEvent(new Event("input", { bubbles: true }));
      textInput.dispatchEvent(new Event("change", { bubbles: true }));
    });
    textInput.after(button, picker);
  }

  [dateStartInput, dateEndInput].forEach((input) => {
    input.addEventListener("blur", () => {
      const parsed = parseDateText(input.value);
      if (parsed) input.value = parsed;
    });
    attachCalendar(input);
  });

  function validate(requiredSlots, missingMessage) {
    if (requiredSlots.some((slot) => !state[slot])) return missingMessage;
    if (saveTelCheckbox.checked && !state.credit) {
      return "도쿄일렉트론코리아 별도 파일을 만들려면 ④ 학점 파일을 골라야 합니다.";
    }
    const start = parseDateText(dateStartInput.value);
    const end = parseDateText(dateEndInput.value);
    if (start === null || end === null) {
      return "교육시작일 기간을 날짜로 읽지 못했습니다. 2026-11-01처럼 연-월-일로 입력해주세요.";
    }
    if (start && end && start > end) {
      return "교육시작일 기간이 올바르지 않습니다. 시작일이 종료일보다 늦습니다.";
    }
    return null;
  }

  async function pickOutputDirectory() {
    try {
      return await window.showDirectoryPicker({ mode: "readwrite" });
    } catch (e) {
      if (e.name === "AbortError") return null; // 취소
      showModal("알려드립니다", "저장 폴더 선택 창을 열지 못했습니다. 버튼을 한 번 더 눌러주세요.");
      return null;
    }
  }

  // 입력 확인 -> 저장 폴더 선택 -> Python 실행 -> 결과를 고른 폴더에 저장.
  // writeBackSlot: Python이 그 칸의 입력 파일을 직접 갱신하는 경우(⑥ 기존 개강내역), 갱신된 내용을 원래 파일에 덮어쓴다.
  // 성공하면 저장한 폴더 이름을, 취소·실패면 null을 돌려준다.
  async function runJob({ requiredSlots, missingMessage, python, writeBackSlot, extra }) {
    const error = validate(requiredSlots, missingMessage);
    if (error) {
      showModal("알려드립니다", error);
      return null;
    }
    const writeBackHandle = writeBackSlot && state[writeBackSlot];
    if (writeBackHandle && !(await ensurePermission(writeBackHandle, "readwrite"))) {
      showModal("알려드립니다", `${writeBackHandle.name} 파일을 갱신하려면 파일 수정 권한을 허용해야 합니다.`);
      return null;
    }
    const dirHandle = await pickOutputDirectory();
    if (!dirHandle) return null;

    clearLog();
    appendLog("처리 중입니다. 끝날 때까지 이 창을 닫지 마세요...");
    await nextFrame();
    const logs = [];
    let ns = null;
    try {
      const paths = await copyInputs(Object.keys(state));
      ns = pyodide.toPy({
        p: paths,
        save_tel: saveTelCheckbox.checked,
        start_date: parseDateText(dateStartInput.value) || null,
        end_date: parseDateText(dateEndInput.value) || null,
        log: (text) => logs.push(text),
        ...extra,
      });
      pyodide.runPython("import os\nos.chdir('/work/out')");
      pyodide.runPython(python, { globals: ns });
      await copyOutputs("/work/out", dirHandle);
      if (writeBackHandle) {
        await writeFileTo(writeBackHandle, pyodide.FS.readFile(paths[writeBackSlot]));
      }
      clearLog();
      showPythonLogs(logs);
      return dirHandle.name;
    } catch (e) {
      clearLog();
      showPythonLogs(logs);
      appendLog("[오류] 처리 중 문제가 생겼습니다:\n" + e);
      showModal("알려드립니다", "처리 중 오류가 났습니다. 아래 처리 결과의 [오류] 내용을 확인해주세요.");
      return null;
    } finally {
      if (ns) ns.destroy();
      // 신청명단 내용을 다음 실행까지 들고 있지 않도록 가상 파일 시스템을 비운다.
      pyodide.runPython("import shutil, os\nos.chdir('/')\nshutil.rmtree('/work', ignore_errors=True)");
    }
  }

  runAllBtn.addEventListener("click", async () => {
    setButtonsDisabled(true);
    try {
      const dirName = await runJob({
        requiredSlots: ["apply", "refund", "general"],
        missingMessage: "신청명단 / 환급과정 커리큘럼 / 비환급과정 커리큘럼, 3개 파일을 모두 골라야 합니다.",
        python: [
          "import run_steps",
          "run_steps.run_all(p['apply'], p['refund'], p['general'], p.get('credit'), '.', save_tel,",
          "                  start_date=start_date, end_date=end_date, log=log)",
        ].join("\n"),
      });
      if (!dirName) return;
      appendLog(`\n모든 작업이 끝났습니다. 저장 위치: ${dirName} 폴더`);
      clearFileFields();
      showModal("완료", `처리가 끝났습니다.\n저장 위치: ${dirName} 폴더\n\n`
        + "노란색으로 표시된 '확인필요' 줄은 직접 확인해주세요.");
    } finally {
      setButtonsDisabled(false);
    }
  });

  runAddedBtn.addEventListener("click", async () => {
    setButtonsDisabled(true);
    try {
      const hasExisting = Boolean(state.existing_gaegang);
      const missingMessage = "① 신청명단(원본) / ② 환급과정 커리큘럼 / ③ 비환급과정 커리큘럼 / "
        + "⑤ 추가 신청(파일 또는 붙여넣기)을 모두 넣어야 합니다. ⑥ 기존 개강내역 파일은 선택사항입니다.";
      const hasAdded = addedMode === "file" ? Boolean(state.added) : mapState.source === "paste";
      if (!(state.apply && state.refund && state.general && hasAdded)) {
        showModal("알려드립니다", addedMode === "paste" && pasteText.value.trim() && !mapState.source
          ? "붙여넣은 내용 아래의 \"표 읽기\"를 먼저 눌러주세요." : missingMessage);
        return;
      }
      if (!reviewState.built) {
        showModal("알려드립니다", "칸 맞추기 아래의 \"확인 표 만들기\"를 먼저 눌러 내용을 확인해주세요.");
        return;
      }
      if (!reviewState.rows.length) {
        showModal("알려드립니다", "확인 표에 반영할 줄이 없습니다.");
        return;
      }
      recheckReview();
      const flagged = reviewState.checks.filter((c) => c.messages.length).length;
      if (flagged && !(await showConfirm("확인필요가 있습니다",
        `확인 표에 확인필요 ${flagged}건이 있습니다. 결과 파일에도 확인필요로 표시됩니다.\n그대로 반영할까요?`))) {
        return;
      }
      const dirName = await runJob({
        requiredSlots: ["apply", "refund", "general"],
        missingMessage,
        python: [
          "import add_applications, added_mapping",
          "# 확인 표에서 확정한 내용을 표준 양식으로 바꿔 넘긴다(직접 고친 칸·추가한 줄은 확인필요 메모).",
          "added_path = added_mapping.save_reviewed(review_rows, '/work/in/added_standard.xlsx',",
          "                                         deleted=review_deleted, log=log)",
          "add_applications.run(p['apply'], added_path, p['refund'], p['general'], p.get('existing_gaegang', ''), '.',",
          "                     credit_path=p.get('credit') if save_tel else None, log=log,",
          "                     start_date=start_date, end_date=end_date)",
        ].join("\n"),
        // exe와 똑같이 ⑥ 기존 개강내역 파일은 그 파일 자체를 갱신한다.
        writeBackSlot: "existing_gaegang",
        extra: { review_rows: reviewState.rows, review_deleted: reviewState.deleted },
      });
      if (!dirName) return;
      const gaegangMsg = hasExisting ? "개강내역은 고른 파일을 갱신했습니다."
        : "⑥을 고르지 않아 개강내역은 새 파일로 만들었고, 교육회차는 비어 있습니다.";
      appendLog(`\n추가 신청 반영이 끝났습니다. 저장 위치: ${dirName} 폴더`);
      clearFileFields();
      showModal("완료", `추가 신청 반영이 끝났습니다.\n저장 위치: ${dirName} 폴더\n\n`
        + `${gaegangMsg} 노란색 '확인필요' 줄은 직접 확인해주세요.`);
    } finally {
      setButtonsDisabled(false);
    }
  });

  // ----- 시작: 브라우저 확인 + Pyodide 준비 -----
  setButtonsDisabled(true);

  async function init() {
    if (!window.showOpenFilePicker || !window.showDirectoryPicker) {
      logEl.textContent = "";
      showModal("알려드립니다", "이 도구는 Chrome 또는 Edge(PC)에서 열어주세요. 지금 브라우저는 파일 저장 기능을 지원하지 않습니다.");
      return;
    }
    try {
      pyodide = await loadPyodide();
      await pyodide.loadPackage(["pandas", "micropip"]);
      await pyodide.pyimport("micropip").install(PYPI_WHEELS, { deps: false });
      pyodide.FS.mkdirTree("/app");
      for (const name of PY_MODULES) {
        const res = await fetch(`../scripts/${name}.py`, { cache: "no-cache" });
        if (!res.ok) throw new Error(`${name}.py 를 불러오지 못했습니다 (${res.status})`);
        pyodide.FS.writeFile(`/app/${name}.py`, await res.text());
      }
      pyodide.runPython("import sys\nsys.path.insert(0, '/app')");
      const [all, required] = pyToJs(pyodide.runPython(
        "import added_mapping\n[added_mapping.REQUIRED_COLUMNS, added_mapping.REQUIRED_FOR_ADDED]"));
      mapColumns = { all, required };
      clearLog();
      setButtonsDisabled(false);
      if (addedMode === "file" && state.added) onAddedFileChanged(state.added); // 준비 전에 ⑤를 골라둔 경우
    } catch (e) {
      clearLog();
      appendLog("[오류] 처리 도구를 불러오지 못했습니다. 인터넷 연결을 확인하고 새로고침해주세요.\n" + e);
    }
  }

  init();
})();
