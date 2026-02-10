const input = document.getElementById("sheetUrl");
const useTaskIdCheckbox = document.getElementById("useTaskIdForOfferSelection");
const generateProductLinkCheckbox = document.getElementById("generateProductLink");
const customTaskIdInput = document.getElementById("customTaskId");
const taskFilterWarning = document.getElementById("taskFilterWarning");
const saveBtn = document.getElementById("save");

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
  chrome.storage.sync.set({
    sheetUrl: url,
    useTaskIdForOfferSelection: useTaskIdCheckbox.checked,
    generateProductLink: generateProductLinkCheckbox.checked,
    customTaskId
  });
};
