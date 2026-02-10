const input = document.getElementById("sheetUrl");
const useTaskIdCheckbox = document.getElementById("useTaskIdForOfferSelection");
const generateProductLinkCheckbox = document.getElementById("generateProductLink");
const customTaskIdInput = document.getElementById("customTaskId");
const taskFilterWarning = document.getElementById("taskFilterWarning");
const saveBtn = document.getElementById("save");
const saveStatus = document.getElementById("saveStatus");

let saveStatusTimer = null;

function setSaveStatus(message, type) {
  if (!saveStatus) return;
  saveStatus.textContent = message || "";
  saveStatus.classList.remove("ok", "error");
  if (type === "ok" || type === "error") {
    saveStatus.classList.add(type);
  }
  if (saveStatusTimer) {
    clearTimeout(saveStatusTimer);
    saveStatusTimer = null;
  }
  if (message) {
    saveStatusTimer = setTimeout(() => {
      saveStatus.textContent = "";
      saveStatus.classList.remove("ok", "error");
      saveStatusTimer = null;
    }, 2500);
  }
}

function updateTaskFilterWarning() {
  const customTaskId = customTaskIdInput.value.trim();
  const shouldShow = useTaskIdCheckbox.checked && Boolean(customTaskId);
  taskFilterWarning.classList.toggle("show", shouldShow);
}

chrome.storage.sync.get(
  ["sheetUrl", "useTaskIdForOfferSelection", "generateProductLink", "customTaskId"],
  (data) => {
    if (data && data.sheetUrl) {
      input.value = data.sheetUrl;
    }
    useTaskIdCheckbox.checked = data && typeof data.useTaskIdForOfferSelection === "boolean"
      ? data.useTaskIdForOfferSelection
      : true;
    generateProductLinkCheckbox.checked = data && typeof data.generateProductLink === "boolean"
      ? data.generateProductLink
      : true;
    customTaskIdInput.value = data && data.customTaskId
      ? String(data.customTaskId).trim()
      : "";
    updateTaskFilterWarning();
  }
);

customTaskIdInput.addEventListener("input", () => {
  customTaskIdInput.value = customTaskIdInput.value.replace(/[^\d]/g, "");
  updateTaskFilterWarning();
});
useTaskIdCheckbox.addEventListener("change", updateTaskFilterWarning);

saveBtn.onclick = () => {
  const url = input.value.trim();
  const customTaskId = customTaskIdInput.value.trim();
  updateTaskFilterWarning();
  saveBtn.disabled = true;
  setSaveStatus("Сохраняем...", "");
  chrome.storage.sync.set({
    sheetUrl: url,
    useTaskIdForOfferSelection: useTaskIdCheckbox.checked,
    generateProductLink: generateProductLinkCheckbox.checked,
    customTaskId
  }, () => {
    saveBtn.disabled = false;
    if (chrome.runtime.lastError) {
      setSaveStatus(`Ошибка: ${chrome.runtime.lastError.message}`, "error");
      return;
    }
    setSaveStatus("Сохранено", "ok");
  });
};
