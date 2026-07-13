const saveButton = document.getElementById("save-order");
const status = document.getElementById("status");

function showStatus(message, type) {
  status.textContent = message;
  status.className = type || "";
}

function getCurrentOrderHtml() {
  const orderNumberNode = document.querySelector('[data-testid="orderNumber"]');
  const productBlocks = document.querySelectorAll('[data-testid="productBlock"]');

  if (!orderNumberNode || productBlocks.length === 0) {
    return {
      ok: false,
      message: "This is not an AliExpress order page. Open one order and try again."
    };
  }

  const orderNumber = orderNumberNode.textContent.replace(/\D/g, "");
  if (!orderNumber) {
    return {
      ok: false,
      message: "The order number could not be read from this page."
    };
  }

  return {
    ok: true,
    orderNumber,
    html: "<!DOCTYPE html>\n" + document.documentElement.outerHTML
  };
}

saveButton.addEventListener("click", async () => {
  saveButton.disabled = true;
  showStatus("Reading the current page...");

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const [result] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: getCurrentOrderHtml
    });

    const data = result.result;
    if (!data.ok) {
      showStatus(data.message, "error");
      return;
    }

    const file = new Blob([data.html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(file);

    try {
      await chrome.downloads.download({
        url,
        filename: `AliExpress-order-${data.orderNumber}.html`,
        saveAs: true,
        conflictAction: "uniquify"
      });
      showStatus("Saved. You can now run the Python script.", "success");
    } finally {
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
  } catch (error) {
    showStatus("Chrome could not read this page. Open one AliExpress order and try again.", "error");
    console.error(error);
  } finally {
    saveButton.disabled = false;
  }
});
