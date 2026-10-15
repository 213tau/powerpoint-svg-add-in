Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        const btn = document.getElementById("insertSvgButton");
        if (btn) {
            btn.onclick = insertSvgFromClipboard;
        }
    }
});

async function insertSvgFromClipboard() {
    const statusDiv = document.getElementById("statusMessage");
    statusDiv.textContent = "Reading clipboard...";
    statusDiv.className = "";

    try {
        // 1. Read text from clipboard
        const svgText = await navigator.clipboard.readText();
        
        // 2. Validate content
        if (!svgText || !svgText.trim().toLowerCase().includes("<svg")) {
            statusDiv.textContent = "Error: Clipboard does not contain valid <svg> markup.";
            statusDiv.className = "error";
            return;
        }

        statusDiv.textContent = "Inserting vector into slide...";

        // 3. Inject into PowerPoint as XmlSvg
        Office.context.document.setSelectedDataAsync(
            svgText,
            {
                coercionType: Office.CoercionType.XmlSvg,
                imageLeft: 100,
                imageTop: 100,
                imageWidth: 300
            },
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    statusDiv.textContent = "Failed: " + asyncResult.error.message;
                    statusDiv.className = "error";
                } else {
                    statusDiv.textContent = "SVG vector inserted successfully!";
                    statusDiv.className = "success";
                }
            }
        );
    } catch (err) {
        console.error("Clipboard error:", err);
        statusDiv.textContent = "Permission denied or unable to access clipboard.";
        statusDiv.className = "error";
    }
}