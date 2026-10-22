(function () {
  const state = {};
  const rows = {};

  function fileName(path) {
    const parts = String(path).split(/[\\/]/);
    return parts[parts.length - 1];
  }

  function setFieldValue(slot, path) {
    state[slot] = path || "";
    rows[slot].input.value = path || "";
  }

  function closeAllRecentMenus() {
    Object.values(rows).forEach(({ recentMenu }) => recentMenu.classList.add("hidden"));
  }

  function renderRecentMenu(slot, paths) {
    const { recentMenu } = rows[slot];
    recentMenu.innerHTML = "";
    if (!paths || paths.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.textContent = "최근 선택한 파일 없음";
      recentMenu.appendChild(empty);
      return;
    }
    paths.forEach((path) => {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "item";
      item.textContent = fileName(path);
      item.title = path;
      item.addEventListener("click", async (ev) => {
        ev.stopPropagation();
        await callApi(() => window.pywebview.api.use_recent_file(slot, path));
        setFieldValue(slot, path);
        recentMenu.classList.add("hidden");
      });
      recentMenu.appendChild(item);
    });
  }

  async function callApi(promiseFactory) {
    if (!window.pywebview || !window.pywebview.api) {
      showModal("잠시만요", "아직 준비 중입니다. 잠시 후 다시 시도해주세요.");
      return null;
    }
    return promiseFactory();
  }

  document.querySelectorAll(".file-row").forEach((row) => {
    const slot = row.dataset.slot;
    const input = row.querySelector("input[type=text]");
    const pickBtn = row.querySelector(".btn.pick");
    const recentBtn = row.querySelector(".btn.recent");
    const recentMenu = row.querySelector(".recent-menu");
    rows[slot] = { row, input, pickBtn, recentBtn, recentMenu };
    state[slot] = "";

    pickBtn.addEventListener("click", async () => {
      const path = await callApi(() => window.pywebview.api.pick_file(slot));
      if (path) {
        setFieldValue(slot, path);
      }
    });

    recentBtn.addEventListener("click", async (ev) => {
      ev.stopPropagation();
      const wasOpen = !recentMenu.classList.contains("hidden");
      closeAllRecentMenus();
      if (wasOpen) return;
      const paths = await callApi(() => window.pywebview.api.get_recent_files(slot));
      renderRecentMenu(slot, paths || []);
      recentMenu.classList.remove("hidden");
    });
  });

  document.addEventListener("click", () => closeAllRecentMenus());

  window.clearFileFields = function () {
    Object.keys(rows).forEach((slot) => setFieldValue(slot, ""));
  };

  const logEl = document.getElementById("log");
  window.appendLog = function (text) {
    logEl.textContent += (logEl.textContent ? "\n" : "") + text;
    logEl.scrollTop = logEl.scrollHeight;
  };
  window.clearLog = function () {
    logEl.textContent = "";
  };

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

  const saveTelCheckbox = document.getElementById("save-tel");
  const runAllBtn = document.getElementById("run-all");
  const runAddedBtn = document.getElementById("run-added");

  function setButtonsDisabled(disabled) {
    runAllBtn.disabled = disabled;
    runAddedBtn.disabled = disabled;
  }

  runAllBtn.addEventListener("click", async () => {
    setButtonsDisabled(true);
    try {
      const result = await callApi(() => window.pywebview.api.run_all({ ...state }, saveTelCheckbox.checked));
      if (result && result.message) {
        showModal(result.ok ? "완료" : "알려드립니다", result.message);
      }
    } finally {
      setButtonsDisabled(false);
    }
  });

  runAddedBtn.addEventListener("click", async () => {
    setButtonsDisabled(true);
    try {
      const result = await callApi(() => window.pywebview.api.run_added({ ...state }, saveTelCheckbox.checked));
      if (result && result.message) {
        showModal(result.ok ? "완료" : "알려드립니다", result.message);
      }
    } finally {
      setButtonsDisabled(false);
    }
  });

  setButtonsDisabled(true);
  window.addEventListener("pywebviewready", () => setButtonsDisabled(false));
})();
