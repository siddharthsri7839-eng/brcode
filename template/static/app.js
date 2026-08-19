/* ═══════════════════════════════════════════════════════════════════════════
   InvenScan — Frontend Application Logic (Blue & White Theme + 3 Roles)
   ═══════════════════════════════════════════════════════════════════════════ */

const API = "/api";
let currentItemId = null;
let capturedImageData = null;
let currentUser = null;

// ─── Initialize ──────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", async () => {
  const page = document.body.dataset.page;

  // Check auth user
  try {
    const res = await fetch(`${API}/auth/me`);
    if (res.ok) {
      const data = await res.json();
      if (data.authenticated) currentUser = data.user;
    }
  } catch (e) {
    console.warn("Auth check skipped:", e);
  }

  // Load categories dynamically for forms and filters
  loadCategories();

  if (page === "dashboard") {
    loadDashboardItems();
    setupDashboardFilters();
  } else if (page === "form") {
    setupFormPage();
  } else if (page === "users") {
    loadUsers();
  } else if (page === "barcode") {
    loadBarcodeGallery();
  } else if (page === "categories") {
    loadCategoriesPage();
  }
});

// ─── Toast Notifications ─────────────────────────────────────────────────────
function showToast(message, type = "info", duration = 4000) {
  const container = document.getElementById("toast-container");
  if (!container) return;

  const icons = { success: "✅", error: "❌", info: "ℹ️" };
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.innerHTML = `<span>${icons[type] || "ℹ️"}</span><span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateX(100%)";
    toast.style.transition = "all 0.3s ease";
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function formatCurrency(val) {
  return `₹${parseFloat(val || 0).toFixed(2)}`;
}

function formatDate(dateStr) {
  if (!dateStr) return "—";
  try {
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? dateStr : d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
  } catch {
    return dateStr;
  }
}

function getExpiryStatus(expDateStr) {
  if (!expDateStr) return { cls: "", label: "—" };
  const exp = new Date(expDateStr);
  const now = new Date();
  const diff = Math.ceil((exp - now) / (1000 * 60 * 60 * 24));

  if (diff < 0)  return { cls: "badge-red",   label: `Expired ${Math.abs(diff)}d ago` };
  if (diff < 30) return { cls: "badge-amber", label: `Expires in ${diff}d` };
  return { cls: "badge-green", label: formatDate(expDateStr) };
}


// ═══════════════════════════════════════════════════════════════════════════════
//  DASHBOARD PAGE LOGIC
// ═══════════════════════════════════════════════════════════════════════════════

async function loadDashboardItems() {
  const tbody = document.getElementById("inventory-tbody");
  const q = document.getElementById("search-input")?.value || "";
  const cat = document.getElementById("category-filter")?.value || "";

  try {
    const res = await fetch(`${API}/items?q=${encodeURIComponent(q)}&category=${encodeURIComponent(cat)}`);
    const data = await res.json();
    const items = data.items || [];

    // Stats
    const totalEl = document.getElementById("stat-total");
    const valEl = document.getElementById("stat-value");
    const lowEl = document.getElementById("stat-lowstock");
    if (totalEl) totalEl.textContent = data.stats?.total_items || items.length;
    if (valEl) valEl.textContent = formatCurrency(data.stats?.total_value || 0);
    if (lowEl) lowEl.textContent = data.stats?.low_stock || 0;

    if (!tbody) return;
    if (items.length === 0) {
      tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:2.5rem;color:var(--text-muted);">No inventory items found. Click "+ Add / Scan Item" to add stock.</td></tr>`;
      return;
    }

    const isReceiver = currentUser && currentUser.role === "receiver";

    tbody.innerHTML = items.map(item => {
      const expStatus = getExpiryStatus(item.exp_date);
      const stockBadge = (item.quantity < 5) ? "badge-red" : (item.quantity < 10 ? "badge-amber" : "badge-green");

      return `
        <tr id="item-row-${item.id}">
          <td>
            <div style="font-weight:700;color:var(--text-primary);font-size:0.92rem;">${escapeHtml(item.item_name)}</div>
            <div style="display:flex;gap:6px;align-items:center;margin-top:3px;">
              ${item.category ? `<span class="badge badge-blue" style="font-size:0.65rem;">${escapeHtml(item.category)}</span>` : ''}
              ${item.batch_no ? `<span class="badge badge-purple" style="font-size:0.65rem;">Batch: ${escapeHtml(item.batch_no)}</span>` : ''}
            </div>
          </td>
          <td>
            <span class="badge ${stockBadge}">${item.quantity} ${escapeHtml(item.unit || "pcs")}</span>
          </td>
          <td style="font-weight:700;color:var(--blue-700);">${formatCurrency(item.sell_price || item.purchase_price)}</td>
          <td style="color:var(--text-secondary);font-size:0.85rem;">${formatDate(item.mfg_date)}</td>
          <td><span class="badge ${expStatus.cls}">${expStatus.label}</span></td>
          <td>
            <span class="badge badge-blue font-mono" style="font-size:0.75rem;">
              ${escapeHtml(item.barcode_id || "—")}
            </span>
          </td>
          <td>
            <div style="display:flex;gap:6px;">
              <a href="/form/${item.id}" class="btn btn-ghost btn-sm" title="Edit Item">✏️ Edit</a>
              <button onclick="openBarcodeModal('${item.id}', '${escapeHtml(item.barcode_id || '')}', '${escapeHtml(item.item_name)}')" class="btn btn-secondary btn-sm" title="View Barcode">📊 Barcode</button>
              ${!isReceiver ? `<button onclick="deleteItem(${item.id})" class="btn btn-danger btn-sm" title="Delete">🗑️</button>` : ''}
            </div>
          </td>
        </tr>
      `;
    }).join("");

  } catch (err) {
    console.error(err);
    if (tbody) tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:2rem;color:var(--red);">Failed to load inventory items.</td></tr>`;
  }
}

function setupDashboardFilters() {
  const search = document.getElementById("search-input");
  const cat = document.getElementById("category-filter");
  let timer;

  if (search) {
    search.addEventListener("input", () => {
      clearTimeout(timer);
      timer = setTimeout(loadDashboardItems, 250);
    });
  }
  if (cat) cat.addEventListener("change", loadDashboardItems);
}

async function deleteItem(id) {
  if (!confirm("Are you sure you want to delete this inventory item?")) return;
  try {
    const res = await fetch(`${API}/items/${id}`, { method: "DELETE" });
    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || "Failed to delete");
    showToast("Item deleted successfully", "success");
    loadDashboardItems();
  } catch (e) {
    showToast(e.message, "error");
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
//  FORM PAGE LOGIC (PHOTO SCAN & AUTO-FILL)
// ═══════════════════════════════════════════════════════════════════════════════

function setupFormPage() {
  const path = window.location.pathname;
  const match = path.match(/\/form\/(\d+)/);
  if (match) {
    currentItemId = match[1];
    loadItemForEdit(currentItemId);
  }

  setupCaptureZone();
}

async function loadItemForEdit(id) {
  try {
    const res = await fetch(`${API}/items/${id}`);
    const item = await res.json();
    if (item.error) throw new Error(item.error);

    fillForm(item);
    showBarcodePanel(id, item.barcode_id, item.item_name);

    document.getElementById("form-title").textContent = "✏️ Edit Inventory Item";
    document.getElementById("form-submit-btn").textContent = "💾 Update Item & Barcode";
  } catch (err) {
    showToast("Failed to load item: " + err.message, "error");
  }
}

function fillForm(fields, isOCR = false) {
  const map = {
    item_name:  "f-item-name",
    sell_price: "f-sell-price",
    mrp:        "f-sell-price",
    category:   "f-category",
    quantity:   "f-quantity",
    unit:       "f-unit",
    mfg_date:   "f-mfg-date",
    exp_date:   "f-exp-date",
    batch_no:   "f-batch-no",
    barcode_id: "f-barcode-id",
    notes:      "f-notes",
  };

  for (const [key, elId] of Object.entries(map)) {
    if (fields[key] !== undefined && fields[key] !== "" && fields[key] !== null) {
      const el = document.getElementById(elId);
      if (el) {
        el.value = fields[key];
        if (isOCR) {
          el.style.borderColor = "var(--blue-600)";
          el.style.backgroundColor = "#EFF6FF";
          setTimeout(() => {
            el.style.borderColor = "";
            el.style.backgroundColor = "";
          }, 2000);
        }
      }
    }
  }
}



let capturedImages = [];

function setupCaptureZone() {
  const zone = document.getElementById("capture-zone");
  if (!zone) return;

  zone.addEventListener("dragover", e => { e.preventDefault(); zone.classList.add("dragging"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragging"));
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("dragging");
    const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith("image/"));
    if (files.length > 0) handleMultipleFiles(files);
  });
}

function openCamera() {
  const input = document.getElementById("camera-input");
  if (input) {
    input.accept = "image/*";
    input.capture = "environment";
    input.click();
  }
}

function openFilePicker() {
  const input = document.getElementById("camera-input");
  if (input) {
    input.removeAttribute("capture");
    input.accept = "image/*";
    input.click();
  }
}

async function onImageSelected(event) {
  const files = Array.from(event.target.files || []);
  if (files.length > 0) {
    await handleMultipleFiles(files);
  }
  event.target.value = "";
}

function compressImage(file, maxDimension = 1400, quality = 0.88) {
  return new Promise((resolve) => {
    const img = new Image();
    const reader = new FileReader();
    reader.onload = e => {
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDimension || height > maxDimension) {
          if (width > height) {
            height = Math.round((height * maxDimension) / width);
            width = maxDimension;
          } else {
            width = Math.round((width * maxDimension) / height);
            height = maxDimension;
          }
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

async function handleMultipleFiles(files) {
  for (const file of files) {
    const b64 = await compressImage(file);
    capturedImages.push(b64);
  }
  renderMultiPreviews();
  showToast(`📸 ${files.length} photo(s) added! Total: ${capturedImages.length} photos ready.`, "info");
}

function renderMultiPreviews() {
  const container = document.getElementById("multi-preview-container");
  const grid = document.getElementById("multi-preview-grid");
  const badge = document.getElementById("photo-count-badge");
  const heading = document.getElementById("preview-heading");
  const actions = document.getElementById("capture-actions-row");
  const placeholder = document.getElementById("capture-placeholder");

  if (!container || !grid) return;

  if (capturedImages.length === 0) {
    container.style.display = "none";
    if (actions) actions.classList.add("hidden");
    if (badge) badge.style.display = "none";
    if (placeholder) placeholder.style.display = "block";
    return;
  }

  container.style.display = "block";
  if (actions) actions.classList.remove("hidden");
  if (badge) {
    badge.textContent = `${capturedImages.length} Photo${capturedImages.length > 1 ? 's' : ''}`;
    badge.style.display = "inline-flex";
  }
  if (heading) heading.textContent = `Uploaded Photos (${capturedImages.length})`;

  grid.innerHTML = capturedImages.map((imgSrc, idx) => `
    <div class="multi-thumb-card">
      <img src="${imgSrc}" alt="Photo ${idx + 1}" />
      <span class="multi-thumb-badge">#${idx + 1}</span>
      <button type="button" onclick="removePhoto(${idx})" class="multi-thumb-remove" title="Remove photo">✕</button>
    </div>
  `).join("");
}

function removePhoto(index) {
  capturedImages.splice(index, 1);
  renderMultiPreviews();
  showToast("Photo removed", "info");
}

async function runOCR() {
  if (capturedImages.length === 0) {
    showToast("Please upload at least 1 product photo first", "error");
    return;
  }

  const btn = document.getElementById("btn-ocr");
  const statusEl = document.getElementById("ocr-status");
  const chips = document.getElementById("ai-chips");

  if (btn) { btn.disabled = true; btn.textContent = `🔄 Analyzing ${capturedImages.length} Photos...`; }
  if (statusEl) {
    statusEl.style.display = "block";
    statusEl.style.background = "#EFF6FF";
    statusEl.style.color = "var(--blue-700)";
    statusEl.textContent = `🤖 AI Vision OCR scanning ${capturedImages.length} photos and merging data...`;
  }

  try {
    const res = await fetch(`${API}/ocr`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ images: capturedImages })
    });

    const resText = await res.text();
    let data;
    try {
      data = JSON.parse(resText);
    } catch {
      throw new Error(`Server returned (${res.status}): ${resText.substring(0, 120)}`);
    }

    if (!res.ok || data.error) throw new Error(data.error || "OCR extraction failed");

    const fields = data.fields || {};
    fillForm(fields, true);

    if (chips) {
      const badges = [];
      if (fields.item_name) badges.push(`🏷️ ${fields.item_name.substring(0, 30)}`);
      if (fields.sell_price || fields.mrp) badges.push(`💰 MRP: ₹${fields.sell_price || fields.mrp}`);
      if (fields.mfg_date) badges.push(`🏭 MFD: ${fields.mfg_date}`);
      if (fields.exp_date) badges.push(`⏰ EXP: ${fields.exp_date}`);
      if (fields.batch_no) badges.push(`🔢 Batch: ${fields.batch_no}`);
      if (fields.quantity) badges.push(`📦 Qty: ${fields.quantity} ${fields.unit || ''}`);
      chips.innerHTML = badges.map(b => `<div class="ai-chip">${escapeHtml(b)}</div>`).join("");
    }

    if (statusEl) {
      statusEl.style.background = "var(--green-light)";
      statusEl.style.color = "#065F46";
      statusEl.textContent = `✨ Auto-filled ${data.filled || 4} fields from ${data.images_scanned || capturedImages.length} photos!`;
    }
    showToast(`✅ Extracted data merged from ${capturedImages.length} photos!`, "success");

  } catch (err) {
    console.error(err);
    if (statusEl) {
      statusEl.style.background = "var(--red-light)";
      statusEl.style.color = "#991B1B";
      statusEl.textContent = `❌ ${err.message}`;
    }
    showToast("Scan error: " + err.message, "error");
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = "⚡ Scan All Photos & Auto-Fill Form"; }
  }
}

async function saveItem(event) {
  event.preventDefault();

  const payload = {
    item_name:  document.getElementById("f-item-name")?.value?.trim(),
    sell_price: document.getElementById("f-sell-price")?.value,
    category:   document.getElementById("f-category")?.value,
    quantity:   document.getElementById("f-quantity")?.value || 0,
    unit:       document.getElementById("f-unit")?.value || "pcs",
    mfg_date:   document.getElementById("f-mfg-date")?.value,
    exp_date:   document.getElementById("f-exp-date")?.value,
    batch_no:   document.getElementById("f-batch-no")?.value?.trim(),
    barcode_id: document.getElementById("f-barcode-id")?.value?.trim(),
    notes:      document.getElementById("f-notes")?.value?.trim(),
  };

  if (!payload.item_name) {
    showToast("Item Name is required", "error");
    document.getElementById("f-item-name")?.focus();
    return;
  }

  const btn = document.getElementById("form-submit-btn");
  if (btn) { btn.disabled = true; btn.textContent = "⏳ Saving..."; }

  try {
    const url = currentItemId ? `${API}/items/${currentItemId}` : `${API}/items`;
    const method = currentItemId ? "PUT" : "POST";

    const res = await fetch(url, {
      method,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || "Save failed");

    const item = data.item;
    currentItemId = item.id;

    showToast(currentItemId ? "✅ Item saved & barcode generated!" : "✅ Item updated!", "success");
    showBarcodePanel(item.id, item.barcode_id, item.item_name);

    if (item.barcode_id) {
      document.getElementById("f-barcode-id").value = item.barcode_id;
    }
  } catch (err) {
    showToast("Save error: " + err.message, "error");
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = "💾 Save Item & Generate Barcode"; }
  }
}

function showBarcodePanel(id, val, name) {
  const panel = document.getElementById("barcode-panel");
  const img   = document.getElementById("barcode-img");
  const valEl = document.getElementById("barcode-display-value");

  if (!panel || !img) return;
  img.src = `${API}/generate-barcode/${id}?t=${Date.now()}`;
  if (valEl) valEl.textContent = val || `INV${String(id).padStart(6, '0')}`;
  panel.style.display = "block";
}

function downloadBarcode() {
  const img = document.getElementById("barcode-img");
  const val = document.getElementById("barcode-display-value")?.textContent || "barcode";
  if (!img || !img.src) return;
  const a = document.createElement("a");
  a.href = img.src;
  a.download = `${val}.png`;
  a.click();
}


// ═══════════════════════════════════════════════════════════════════════════════
//  USER MANAGEMENT (ADMIN ROLE)
// ═══════════════════════════════════════════════════════════════════════════════

async function loadUsers() {
  const tbody = document.getElementById("users-tbody");
  if (!tbody) return;

  try {
    const res = await fetch(`${API}/users`);
    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || "Failed to load users");

    const users = data.users || [];
    tbody.innerHTML = users.map(u => {
      const roleCls = `role-${u.role}`;
      return `
        <tr>
          <td class="font-mono" style="font-weight:700;">#${u.id}</td>
          <td style="font-weight:700;color:var(--text-primary);">${escapeHtml(u.username)}</td>
          <td>${escapeHtml(u.full_name || '—')}</td>
          <td><span class="user-role-badge ${roleCls}">${escapeHtml(u.role).replace('_', ' ')}</span></td>
          <td style="color:var(--text-secondary);font-size:0.85rem;">${formatDate(u.created_at)}</td>
          <td>
            <div style="display:flex;gap:6px;">
              <button onclick='openEditUserModal(${JSON.stringify(u)})' class="btn btn-ghost btn-sm">✏️ Edit</button>
              <button onclick="deleteUserAccount(${u.id}, '${escapeHtml(u.username)}')" class="btn btn-danger btn-sm">🗑️ Delete</button>
            </div>
          </td>
        </tr>
      `;
    }).join("");

  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:2rem;color:var(--red);">${err.message}</td></tr>`;
  }
}

function openCreateUserModal() {
  document.getElementById("user-modal-title").textContent = "Create New User";
  document.getElementById("modal-user-id").value = "";
  document.getElementById("m-username").value = "";
  document.getElementById("m-username").disabled = false;
  document.getElementById("m-fullname").value = "";
  document.getElementById("m-role").value = "receiver";
  document.getElementById("m-password").value = "";
  document.getElementById("m-password").required = true;
  document.getElementById("pwd-required-star").style.display = "inline";
  document.getElementById("pwd-hint").style.display = "none";
  document.getElementById("user-modal").style.display = "flex";
}

function openEditUserModal(u) {
  document.getElementById("user-modal-title").textContent = `Edit User: ${u.username}`;
  document.getElementById("modal-user-id").value = u.id;
  document.getElementById("m-username").value = u.username;
  document.getElementById("m-username").disabled = true;
  document.getElementById("m-fullname").value = u.full_name || "";
  document.getElementById("m-role").value = u.role;
  document.getElementById("m-password").value = "";
  document.getElementById("m-password").required = false;
  document.getElementById("pwd-required-star").style.display = "none";
  document.getElementById("pwd-hint").style.display = "block";
  document.getElementById("user-modal").style.display = "flex";
}

function closeUserModal() {
  document.getElementById("user-modal").style.display = "none";
}

async function saveUserForm(e) {
  e.preventDefault();
  const id       = document.getElementById("modal-user-id").value;
  const username = document.getElementById("m-username").value.trim();
  const full_name= document.getElementById("m-fullname").value.trim();
  const role     = document.getElementById("m-role").value;
  const password = document.getElementById("m-password").value.trim();

  const url    = id ? `${API}/users/${id}` : `${API}/users`;
  const method = id ? "PUT" : "POST";
  const body   = { username, full_name, role };
  if (password) body.password = password;

  try {
    const res = await fetch(url, {
      method,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });

    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || "User operation failed");

    showToast(id ? "User updated successfully!" : "New user created!", "success");
    closeUserModal();
    loadUsers();
  } catch (err) {
    showToast(err.message, "error");
  }
}

async function deleteUserAccount(id, uname) {
  if (!confirm(`Delete user account "${uname}"? This cannot be undone.`)) return;
  try {
    const res = await fetch(`${API}/users/${id}`, { method: "DELETE" });
    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || "Delete failed");
    showToast("User deleted", "success");
    loadUsers();
  } catch (err) {
    showToast(err.message, "error");
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
//  BARCODE CENTER LOGIC
// ═══════════════════════════════════════════════════════════════════════════════

async function loadBarcodeGallery() {
  const tbody = document.getElementById("barcode-gallery-tbody");
  if (!tbody) return;

  try {
    const res = await fetch(`${API}/items`);
    const data = await res.json();
    const items = data.items || [];

    if (items.length === 0) {
      tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:2rem;color:var(--text-muted)">No items in database yet.</td></tr>`;
      return;
    }

    tbody.innerHTML = items.map(item => {
      const bc = item.barcode_id || `INV${String(item.id).padStart(6, '0')}`;
      return `
        <tr>
          <td style="font-weight:700;">${escapeHtml(item.item_name)}</td>
          <td class="font-mono">${escapeHtml(bc)}</td>
          <td>
            <img src="${API}/generate-barcode-value/${encodeURIComponent(bc)}" style="max-height:50px;background:#fff;padding:4px;border:1px solid var(--border);border-radius:4px;" />
          </td>
          <td>
            <a href="${API}/generate-barcode-value/${encodeURIComponent(bc)}" download="${bc}.png" class="btn btn-primary btn-sm">⬇️ Download</a>
          </td>
        </tr>
      `;
    }).join("");
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--red)">Failed to load barcodes.</td></tr>`;
  }
}

function generateCustomBarcode() {
  const val = document.getElementById("custom-barcode-val")?.value?.trim();
  if (!val) { showToast("Enter a barcode value first", "error"); return; }

  const preview = document.getElementById("custom-barcode-preview");
  const img     = document.getElementById("custom-bc-img");
  const dl      = document.getElementById("custom-bc-dl");

  const url = `${API}/generate-barcode-value/${encodeURIComponent(val)}`;
  img.src = url;
  dl.href = url;
  dl.download = `${val}.png`;
  preview.style.display = "block";
}

function openBarcodeScanner() {
  const input = document.getElementById("barcode-scan-input");
  if (input) input.click();
}

async function onBarcodeScanImage(event) {
  const file = event.target.files[0];
  if (!file) return;

  showToast("🔍 Scanning barcode...", "info");

  const reader = new FileReader();
  reader.onload = async e => {
    try {
      const res = await fetch(`${API}/scan-barcode`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ image: e.target.result })
      });
      const data = await res.json();
      if (data.found && data.item) {
        showToast(`✅ Found: ${data.item.item_name}`, "success");
        if (window.location.pathname.includes("/form")) {
          fillForm(data.item, true);
        }
      } else {
        showToast(data.message || "Barcode not recognized in database", "error");
      }
    } catch (err) {
      showToast("Scan error: " + err.message, "error");
    }
  };
  reader.readAsDataURL(file);
  event.target.value = "";
}

// Modal helper
function openBarcodeModal(id, val, name) {
  const modal = document.getElementById("barcode-modal");
  const img   = document.getElementById("modal-barcode-img");
  const title = document.getElementById("modal-item-title");
  const valEl = document.getElementById("modal-barcode-val");

  if (!modal || !img) return;
  title.textContent = name || "Product Barcode";
  img.src = `${API}/generate-barcode/${id}?t=${Date.now()}`;
  valEl.textContent = val || `INV${String(id).padStart(6, '0')}`;
  modal.style.display = "flex";
}

function closeBarcodeModal() {
  const modal = document.getElementById("barcode-modal");
  if (modal) modal.style.display = "none";
}

function downloadModalBarcode() {
  const img = document.getElementById("modal-barcode-img");
  const val = document.getElementById("modal-barcode-val")?.textContent || "barcode";
  if (!img || !img.src) return;
  const a = document.createElement("a");
  a.href = img.src;
  a.download = `${val}.png`;
  a.click();
}

// ─── Dynamic Category Management ───────────────────────────────────────────
async function loadCategories(selectedCategory = null) {
  try {
    const res = await fetch(`${API}/categories`);
    const data = await res.json();
    const categories = data.categories || [];

    // 1. Populate form select
    const formSelect = document.getElementById("f-category");
    if (formSelect) {
      const currentVal = selectedCategory || formSelect.value;
      formSelect.innerHTML = `<option value="">Select Category...</option>` +
        categories.map(c => `<option value="${escapeHtml(c.name)}">${c.icon || '📦'} ${escapeHtml(c.name)}</option>`).join("");
      if (currentVal) formSelect.value = currentVal;
    }

    // 2. Populate dashboard filter select
    const filterSelect = document.getElementById("category-filter");
    if (filterSelect) {
      const filterVal = filterSelect.value;
      filterSelect.innerHTML = `<option value="">All Categories</option>` +
        categories.map(c => `<option value="${escapeHtml(c.name)}">${c.icon || '📦'} ${escapeHtml(c.name)}</option>`).join("");
      if (filterVal) filterSelect.value = filterVal;
    }
  } catch (err) {
    console.warn("Failed to load categories:", err);
  }
}

function openCategoryModal() {
  const modal = document.getElementById("category-modal");
  const input = document.getElementById("new-cat-name");
  if (modal) {
    modal.style.display = "flex";
    if (input) { input.value = ""; input.focus(); }
  }
}

function closeCategoryModal() {
  const modal = document.getElementById("category-modal");
  if (modal) modal.style.display = "none";
}

async function saveNewCategory(e) {
  e.preventDefault();
  const nameInput = document.getElementById("new-cat-name");
  const iconInput = document.getElementById("new-cat-icon");
  const name = nameInput ? nameInput.value.trim() : "";
  const icon = iconInput ? iconInput.value.trim() : "📦";

  if (!name) {
    showToast("Category name is required", "error");
    return;
  }

  try {
    const res = await fetch(`${API}/categories`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, icon })
    });

    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || "Failed to create category");

    showToast(`✅ Category "${name}" added!`, "success");
    closeCategoryModal();
    await loadCategories(name);
  } catch (err) {
    showToast(err.message, "error");
  }
}

async function loadCategoriesPage() {
  const tbody = document.getElementById("categories-page-tbody");
  const countEl = document.getElementById("cat-total-count");
  if (!tbody) return;

  try {
    const res = await fetch(`${API}/categories`);
    const data = await res.json();
    const categories = data.categories || [];

    if (countEl) countEl.textContent = `${categories.length} Categories`;

    if (categories.length === 0) {
      tbody.innerHTML = `<tr><td colspan="3" style="text-align:center;padding:2rem;color:var(--text-muted)">No categories found. Use the form on the left to add one!</td></tr>`;
      return;
    }

    const isAdmin = currentUser && currentUser.role === "admin";

    tbody.innerHTML = categories.map(c => `
      <tr>
        <td style="font-size:1.4rem;text-align:center;width:60px;">${c.icon || '📦'}</td>
        <td style="font-weight:700;color:var(--text-primary);font-size:0.95rem;">${escapeHtml(c.name)}</td>
        <td style="width:100px;">
          ${isAdmin ? `<button onclick="deleteCategory(${c.id}, '${escapeHtml(c.name)}')" class="btn btn-danger btn-sm" title="Delete category">🗑️ Delete</button>` : '<span style="color:var(--text-muted);font-size:0.75rem;">Default</span>'}
        </td>
      </tr>
    `).join("");

  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="3" style="text-align:center;padding:2rem;color:var(--red)">Failed to load categories: ${err.message}</td></tr>`;
  }
}

async function savePageCategory(e) {
  e.preventDefault();
  const name = document.getElementById("page-cat-name")?.value.trim();
  const icon = document.getElementById("page-cat-icon")?.value.trim() || "📦";

  if (!name) {
    showToast("Category name is required", "error");
    return;
  }

  try {
    const res = await fetch(`${API}/categories`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, icon })
    });

    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || "Failed to create category");

    showToast(`✅ Category "${name}" created!`, "success");
    document.getElementById("page-cat-name").value = "";
    loadCategoriesPage();
    loadCategories();
  } catch (err) {
    showToast(err.message, "error");
  }
}

async function deleteCategory(id, name) {
  if (!confirm(`Delete category "${name}"? Existing items with this category will not be deleted.`)) return;

  try {
    const res = await fetch(`${API}/categories/${id}`, { method: "DELETE" });
    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || "Delete failed");

    showToast(`Category "${name}" deleted`, "success");
    loadCategoriesPage();
    loadCategories();
  } catch (err) {
    showToast(err.message, "error");
  }
}

// ─── Excel / CSV Bulk Import ────────────────────────────────────────────────
async function handleExcelImport(event) {
  const file = event.target.files[0];
  if (!file) return;

  showToast(`⏳ Importing ${file.name}...`, "info");
  const formData = new FormData();
  formData.append("file", file);

  try {
    const res = await fetch(`${API}/import/excel`, {
      method: "POST",
      body: formData
    });
    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || "Import failed");

    showToast(`✅ Successfully imported ${data.imported_count || 0} items!`, "success");
    if (typeof loadDashboardItems === "function") {
      loadDashboardItems();
    }
  } catch (err) {
    showToast("Import error: " + err.message, "error");
  } finally {
    event.target.value = "";
  }
}
