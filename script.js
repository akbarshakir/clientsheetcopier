
let accessToken = null;
let tokenClient = null;

const $ = id => document.getElementById(id);

const GOOGLE_DOCS_MIME = "application/vnd.google-apps.document";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DOC_MIME = "application/msword";
const ODT_MIME = "application/vnd.oasis.opendocument.text";
const RTF_MIME = "application/rtf";

function showMessage(message, type = "success") {
  const box = $("message");
  box.textContent = message;
  box.className = `message ${type}`;
}

function setConnected(email) {
  $("connectionStatus").textContent = "Connected";
  $("connectionStatus").className = "status online";
  $("connectBtn").textContent = email;
  $("createBtn").disabled = false;
}

function waitForGoogle() {
  return new Promise((resolve, reject) => {
    let attempts = 0;

    const timer = setInterval(() => {
      if (window.google?.accounts?.oauth2) {
        clearInterval(timer);
        resolve();
        return;
      }

      attempts++;

      if (attempts > 50) {
        clearInterval(timer);
        reject(new Error("Google Identity Services could not load."));
      }
    }, 200);
  });
}

async function apiRequest(url, options = {}) {
  if (!accessToken) {
    throw new Error("Please connect your Google Account first.");
  }

  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...options.headers
    }
  });

  if (response.status === 401) {
    accessToken = null;
    $("createBtn").disabled = true;
    $("connectionStatus").textContent = "Reconnect Required";
    $("connectionStatus").className = "status offline";

    throw new Error("Google session expired. Please reconnect.");
  }

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));

    throw new Error(
      error.error?.message || `Google API Error: ${response.status}`
    );
  }

  return response;
}

// Google Account Connection
async function connectGoogle() {
  try {
    await waitForGoogle();

    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: APP_CONFIG.googleClientId,
      scope: APP_CONFIG.scopes,

      callback: async response => {
        if (response.error) {
          showMessage(response.error, "error");
          return;
        }

        accessToken = response.access_token;

        try {
          const userResponse = await apiRequest(
            "https://www.googleapis.com/oauth2/v3/userinfo"
          );

          const user = await userResponse.json();

          if (
            !user.email ||
            user.email.toLowerCase() !== APP_CONFIG.myEmail.toLowerCase()
          ) {
            accessToken = null;
            $("createBtn").disabled = true;
            $("connectionStatus").textContent = "Wrong Account";
            $("connectionStatus").className = "status offline";

            showMessage(
              "Please connect using " + APP_CONFIG.myEmail,
              "error"
            );
            return;
          }

          setConnected(user.email);
          showMessage("Google account connected successfully.");

        } catch (error) {
          accessToken = null;
          showMessage(error.message, "error");
        }
      }
    });

    tokenClient.requestAccessToken({ prompt: "consent" });

  } catch (error) {
    showMessage(error.message, "error");
  }
}

// Extract ID from Google Docs or Drive URL
function extractDocumentId(url) {
  const patterns = [
    /\/document\/d\/([a-zA-Z0-9_-]+)/,
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/open\?id=([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/
  ];

  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match) return match[1];
  }

  throw new Error(
    "Invalid link. Please paste a Google Docs or Google Drive file URL."
  );
}

// Get Original File Details
async function getOriginalDocument(fileId) {
  const url =
    `https://www.googleapis.com/drive/v3/files/${fileId}` +
    `?fields=id,name,mimeType,owners(emailAddress),webViewLink` +
    `&supportsAllDrives=true`;

  const response = await apiRequest(url);
  return response.json();
}

// Create Native Google Docs Copy
async function createDocumentCopy(fileId, name) {
  const url =
    `https://www.googleapis.com/drive/v3/files/${fileId}/copy` +
    `?supportsAllDrives=true&fields=id,name,mimeType,webViewLink`;

  const response = await apiRequest(url, {
    method: "POST",
    body: JSON.stringify({ name })
  });

  return response.json();
}

// Download Word File
async function downloadOriginalFile(fileId) {
  const url =
    `https://www.googleapis.com/drive/v3/files/${fileId}` +
    `?alt=media&supportsAllDrives=true`;

  const response = await apiRequest(url);
  return response.blob();
}

// Convert Word or ODT File into Google Docs
async function convertToGoogleDocs(fileId, name, mimeType) {
  const allowedTypes = [
    DOCX_MIME,
    DOC_MIME,
    ODT_MIME,
    RTF_MIME
  ];

  if (!allowedTypes.includes(mimeType)) {
    throw new Error(
      "Unsupported file format. Please use Google Docs, DOCX, DOC, ODT or RTF."
    );
  }

  showMessage("Downloading original document for conversion...");

  const originalBlob = await downloadOriginalFile(fileId);

  if (!originalBlob.size) {
    throw new Error("The original document is empty or could not be downloaded.");
  }

  showMessage("Converting document into Google Docs format...");

  const boundary = "LINKIFY_CONVERSION_BOUNDARY_739184";

  const metadata = {
    name,
    mimeType: GOOGLE_DOCS_MIME
  };

  const multipartBody = new Blob([
    `--${boundary}\r\n`,
    "Content-Type: application/json; charset=UTF-8\r\n\r\n",
    JSON.stringify(metadata),
    `\r\n--${boundary}\r\n`,
    `Content-Type: ${originalBlob.type || mimeType}\r\n\r\n`,
    originalBlob,
    `\r\n--${boundary}--`
  ], {
    type: `multipart/related; boundary=${boundary}`
  });

  const uploadUrl =
    "https://www.googleapis.com/upload/drive/v3/files" +
    "?uploadType=multipart&supportsAllDrives=true" +
    "&fields=id,name,mimeType,webViewLink";

  const response = await apiRequest(uploadUrl, {
    method: "POST",
    headers: {
      "Content-Type": `multipart/related; boundary=${boundary}`
    },
    body: multipartBody
  });

  const converted = await response.json();

  if (converted.mimeType !== GOOGLE_DOCS_MIME) {
    throw new Error("Google did not return a native Google Docs file.");
  }

  return converted;
}

// Replace Original Owner Email in Document Text
async function replaceOwnerEmail(copyId, ownerEmail) {
  const url =
    `https://docs.googleapis.com/v1/documents/${copyId}:batchUpdate`;

  const body = {
    requests: [
      {
        replaceAllText: {
          containsText: {
            text: ownerEmail,
            matchCase: false
          },
          replaceText: APP_CONFIG.myEmail
        }
      }
    ]
  };

  const response = await apiRequest(url, {
    method: "POST",
    body: JSON.stringify(body)
  });

  return response.json();
}

// Enable Anyone With Link Viewer Access
async function enablePublicViewer(copyId) {
  const url =
    `https://www.googleapis.com/drive/v3/files/${copyId}/permissions` +
    `?sendNotificationEmail=false&supportsAllDrives=true`;

  const response = await apiRequest(url, {
    method: "POST",
    body: JSON.stringify({
      type: "anyone",
      role: "reader"
    })
  });

  return response.json();
}

// Delete Incomplete Copy if an Operation Fails
async function deleteCopy(copyId) {
  try {
    await apiRequest(
      `https://www.googleapis.com/drive/v3/files/${copyId}` +
      `?supportsAllDrives=true`,
      { method: "DELETE" }
    );
  } catch (error) {
    console.error("Could not delete incomplete copy:", error);
  }
}

// Main Copy Workflow
async function createCopy() {
  const originalUrl = $("docUrl").value.trim();
  const customName = $("newName").value.trim();

  if (!originalUrl) {
    showMessage("Please paste the original document URL.", "error");
    return;
  }

  if (!accessToken) {
    showMessage("Connect your Google Account first.", "error");
    return;
  }

  const button = $("createBtn");
  button.disabled = true;
  button.textContent = "Processing...";

  let copyId = null;

  try {
    $("result").classList.add("hidden");

    showMessage("Checking original document...");

    const sourceId = extractDocumentId(originalUrl);
    const original = await getOriginalDocument(sourceId);

    const ownerEmail = original.owners?.find(
      owner => owner.emailAddress
    )?.emailAddress || null;

    if (
      ownerEmail &&
      ownerEmail.toLowerCase() === APP_CONFIG.myEmail.toLowerCase()
    ) {
      throw new Error(
        "This document already belongs to your LINKIFY account."
      );
    }

    const newName =
      customName || `LINKIFY - ${original.name} - Copy`;

    let copy;

    if (original.mimeType === GOOGLE_DOCS_MIME) {
      showMessage("Creating native Google Docs copy...");

      copy = await createDocumentCopy(sourceId, newName);

    } else {
      showMessage("Word or Office document detected. Starting automatic conversion...");

      copy = await convertToGoogleDocs(
        sourceId,
        newName,
        original.mimeType
      );
    }

    copyId = copy.id;

    if (copy.mimeType !== GOOGLE_DOCS_MIME) {
      throw new Error("The new file is not a native Google Docs document.");
    }

    let replacements = 0;
    let replacementMessage = "";

    if (ownerEmail) {
      showMessage("Checking and replacing original owner's email...");

      const replaceResult = await replaceOwnerEmail(
        copyId,
        ownerEmail
      );

      replacements =
        replaceResult.replies?.[0]?.replaceAllText?.occurrencesChanged ?? 0;

      replacementMessage = replacements > 0
        ? `${replacements} occurrence(s) of the detected owner email were replaced.`
        : "The detected owner email was not found in the document text.";
    } else {
      replacementMessage =
        "Google did not provide the original owner's email. Email replacement was skipped.";
    }

    showMessage("Setting Anyone with the link Viewer access...");

    await enablePublicViewer(copyId);

    const shareLink =
      `https://docs.google.com/document/d/${copyId}/edit`;

    saveHistory({
      name: newName,
      link: shareLink,
      date: new Date().toLocaleString(),
      replacements
    });

    $("resultLink").value = shareLink;
    $("openLink").href = shareLink;
    $("replacementInfo").textContent = replacementMessage;

    $("result").classList.remove("hidden");

    showMessage(
      "Google Docs copy created and Viewer sharing enabled successfully.",
      "success"
    );

    $("docUrl").value = "";
    $("newName").value = "";

  } catch (error) {
    if (copyId) {
      await deleteCopy(copyId);
    }

    showMessage(error.message, "error");

  } finally {
    button.disabled = !accessToken;
    button.textContent = "Create Copy & Generate Link";
  }
}

// Local Browser History
function saveHistory(item) {
  const history = JSON.parse(
    localStorage.getItem("linkifyDocsHistory") || "[]"
  );

  history.unshift(item);

  localStorage.setItem(
    "linkifyDocsHistory",
    JSON.stringify(history.slice(0, 50))
  );

  renderHistory();
}

function escapeHTML(value) {
  return String(value).replace(/[&<>"']/g, char => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  })[char]);
}

function renderHistory() {
  const history = JSON.parse(
    localStorage.getItem("linkifyDocsHistory") || "[]"
  );

  const box = $("history");

  if (!history.length) {
    box.innerHTML = '<p class="muted">No copies created yet.</p>';
    return;
  }

  box.innerHTML = history.map(item => `
    <div class="history-item">
      <strong>${escapeHTML(item.name)}</strong>
      <a href="${escapeHTML(item.link)}" target="_blank" rel="noopener">
        Open Google Document
      </a>
      <small>
        ${escapeHTML(item.date)}
        | ${Number(item.replacements) || 0} replacements
      </small>
    </div>
  `).join("");
}

// Copy Shareable Link
async function copyResultLink() {
  const link = $("resultLink").value;

  try {
    await navigator.clipboard.writeText(link);
    $("copyBtn").textContent = "Link Copied!";
  } catch {
    $("resultLink").select();
    document.execCommand("copy");
    $("copyBtn").textContent = "Link Copied!";
  }

  setTimeout(() => {
    $("copyBtn").textContent = "Copy Shareable Link";
  }, 2000);
}

// Clear Local History
function clearHistory() {
  if (!confirm("Delete local copy history?")) return;

  localStorage.removeItem("linkifyDocsHistory");
  renderHistory();
}

// Initialize
document.addEventListener("DOMContentLoaded", () => {
  $("connectBtn").addEventListener("click", connectGoogle);
  $("createBtn").addEventListener("click", createCopy);
  $("copyBtn").addEventListener("click", copyResultLink);
  $("clearHistory").addEventListener("click", clearHistory);

  renderHistory();
});
