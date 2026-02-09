const input = document.getElementById("sheetUrl");
const saveBtn = document.getElementById("save");

chrome.storage.sync.get("sheetUrl", (data) => {
  if (data && data.sheetUrl) {
    input.value = data.sheetUrl;
  }
});

saveBtn.onclick = () => {
  const url = input.value;
  chrome.storage.sync.set({ sheetUrl: url });
};
