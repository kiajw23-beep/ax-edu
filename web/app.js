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
  document.getElementById("modal-ok").addEventListener("click", () => {
    modalOverlay.classList.add("hidden");
  });

  function showModal(title, message) {
    modalTitle.textContent = title;
    modalMessage.textContent = message;
    modalOverlay.classList.remove("hidden");
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

  // ----- ⑤ 추가 신청 파일 칸 맞추기 (scripts/added_mapping.py) -----
  // 서식이 제각각인 추가 신청 파일에서 시트·칸 이름 줄을 고르고, 원본 신청명단 칸마다 쓸 칸을 고른다.
  // 읽은 내용은 Pyodide 메모리(_map_sheets)에만 두고, 파일을 바꾸거나 화면을 비우면 지운다.
  const mapEl = document.getElementById("mapping");
  const mapStatus = document.getElementById("map-status");
  const mapBody = document.getElementById("map-body");
  const mapSheet = document.getElementById("map-sheet");
  const mapHeader = document.getElementById("map-header");
  const mapPreview = document.getElementById("map-preview");
  const mapFields = document.getElementById("map-fields");
  const mapState = { ready: false, sheet: null, headerRow: 1, columns: [], mapping: {} };
  let mapColumns = null; // { all: 원본 신청명단 칸, required: 꼭 맞춰야 하는 칸 }

  function pyToJs(proxy) {
    const value = proxy.toJs({ dict_converter: Object.fromEntries });
    if (proxy.destroy) proxy.destroy();
    return value;
  }

  function missingRequiredColumns() {
    return mapColumns ? mapColumns.required.filter((c) => mapState.mapping[c] === undefined) : [];
  }

  async function onAddedFileChanged(handle) {
    mapState.ready = false;
    if (pyodide) pyodide.runPython("globals().pop('_map_sheets', None)");
    if (!handle) {
      mapEl.classList.add("hidden");
      return;
    }
    mapEl.classList.remove("hidden");
    mapBody.classList.add("hidden");
    if (!pyodide) {
      mapStatus.textContent = "처리 도구를 불러오는 중입니다. 준비되면 칸 맞추기 화면이 나옵니다.";
      return;
    }
    mapStatus.textContent = "파일을 읽는 중입니다...";
    await nextFrame();
    try {
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
      if (state.added !== handle) return; // 읽는 사이에 다른 파일을 골랐음
      const sheets = pyToJs(pyodide.runPython("list(_map_sheets)"));
      mapSheet.innerHTML = "";
      sheets.forEach((name) => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        mapSheet.appendChild(option);
      });
      mapState.sheet = sheets[0];
      mapState.headerRow = 1;
      mapHeader.value = 1;
      refreshMapping();
      mapStatus.textContent = "";
      mapBody.classList.remove("hidden");
      mapState.ready = true;
    } catch (e) {
      mapStatus.textContent = "파일을 읽지 못했습니다. 엑셀(xlsx) 또는 CSV 파일인지 확인해주세요.\n" + e;
    }
  }

  // 시트·칸 이름 줄이 바뀔 때마다 미리보기와 칸 목록을 다시 만들고, 칸 고르기는 자동 추천으로 되돌린다.
  function refreshMapping() {
    pyodide.globals.set("_map_sheet", mapState.sheet);
    pyodide.globals.set("_map_header", mapState.headerRow);
    const info = pyToJs(pyodide.runPython("added_mapping.sheet_info(_map_sheets, _map_sheet, _map_header)"));
    mapHeader.max = Math.max(1, info.row_count);
    mapState.columns = info.columns;
    mapState.mapping = { ...info.suggested };
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

  function renderFields() {
    mapFields.innerHTML = "";
    mapColumns.all.forEach((target) => {
      const required = mapColumns.required.includes(target);
      const field = document.createElement("label");
      field.className = "mapping-field";
      const name = document.createElement("span");
      name.textContent = required ? `${target} *` : target;
      const select = document.createElement("select");
      const none = document.createElement("option");
      none.value = "";
      none.textContent = required ? "(골라주세요)" : "(없음 - 비워 둠)";
      select.appendChild(none);
      mapState.columns.forEach((col) => {
        const option = document.createElement("option");
        option.value = String(col.index);
        option.textContent = col.label;
        select.appendChild(option);
      });
      const update = () => field.classList.toggle("missing", required && mapState.mapping[target] === undefined);
      if (mapState.mapping[target] !== undefined) select.value = String(mapState.mapping[target]);
      select.addEventListener("change", () => {
        if (select.value === "") delete mapState.mapping[target];
        else mapState.mapping[target] = Number(select.value);
        update();
      });
      update();
      field.append(name, select);
      mapFields.appendChild(field);
    });
  }

  mapSheet.addEventListener("change", () => {
    mapState.sheet = mapSheet.value;
    mapState.headerRow = 1;
    mapHeader.value = 1;
    refreshMapping();
  });

  mapHeader.addEventListener("change", () => {
    const row = parseInt(mapHeader.value, 10);
    if (!(row >= 1 && row <= Number(mapHeader.max))) {
      mapHeader.value = mapState.headerRow;
      return;
    }
    mapState.headerRow = row;
    refreshMapping();
  });

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

  function validate(requiredSlots, missingMessage) {
    if (requiredSlots.some((slot) => !state[slot])) return missingMessage;
    if (saveTelCheckbox.checked && !state.credit) {
      return "도쿄일렉트론코리아 별도 파일을 만들려면 ④ 학점 파일을 골라야 합니다.";
    }
    const start = dateStartInput.value;
    const end = dateEndInput.value;
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
        start_date: dateStartInput.value || null,
        end_date: dateEndInput.value || null,
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
      if (state.added) {
        if (!mapState.ready) {
          showModal("알려드립니다", "⑤ 추가 신청 파일의 칸 맞추기가 아직 준비되지 않았습니다. 잠시 후 다시 눌러주세요.");
          return;
        }
        const missing = missingRequiredColumns();
        if (missing.length) {
          showModal("알려드립니다", `칸 맞추기에서 다음 칸을 골라주세요: ${missing.join(", ")}`);
          return;
        }
      }
      const dirName = await runJob({
        requiredSlots: ["apply", "refund", "general", "added"],
        missingMessage: "① 신청명단(원본) / ② 환급과정 커리큘럼 / ③ 비환급과정 커리큘럼 / "
          + "⑤ 추가 신청 파일을 모두 골라야 합니다. ⑥ 기존 개강내역 파일은 선택사항입니다.",
        python: [
          "import add_applications, added_mapping",
          "p['added'] = added_mapping.save_standard(p['added'], map_sheet, map_header, map_mapping,",
          "                                         '/work/in/added_standard.xlsx', log=log)",
          "add_applications.run(p['apply'], p['added'], p['refund'], p['general'], p.get('existing_gaegang', ''), '.',",
          "                     credit_path=p.get('credit') if save_tel else None, log=log,",
          "                     start_date=start_date, end_date=end_date)",
        ].join("\n"),
        // exe와 똑같이 ⑥ 기존 개강내역 파일은 그 파일 자체를 갱신한다.
        writeBackSlot: "existing_gaegang",
        extra: { map_sheet: mapState.sheet, map_header: mapState.headerRow, map_mapping: mapState.mapping },
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
      if (state.added) onAddedFileChanged(state.added); // 준비 전에 ⑤를 골라둔 경우
    } catch (e) {
      clearLog();
      appendLog("[오류] 처리 도구를 불러오지 못했습니다. 인터넷 연결을 확인하고 새로고침해주세요.\n" + e);
    }
  }

  init();
})();
