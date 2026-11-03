// ============ BEEM CONFIG ============
// WEKA FUNGUO ZAKO MPYA HAPA (badilisha XXXX)
const BEEM_API_KEY = "0fb90a5174db09ad";      // WEKA KEY YAKO HAPA
const BEEM_SECRET_KEY = "MWRhMGI5MTljZjQ1ZjZhMjYzZThkYzRiZTM2YzU2ODA4MjQyMzZhZjU3YjIwNmI2Nzc0YTBmMDRlNjk3MzM4Yg== ";   // WEKA SECRET YAKO HAPA

// ============ FIREBASE CONFIG ============
const firebaseConfig = {
  apiKey: "AIzaSyBhh3nIzAWyvT47HQ-hh6umjqoCMYBI1Lk",
  authDomain: "johcards-2db9b.firebaseapp.com",
  projectId: "johcards-2db9b",
  storageBucket: "johcards-2db9b.firebasestorage.app",
  messagingSenderId: "955656442066",
  appId: "1:955656442066:web:3f0d0335191d67eac61d0b"
};
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

// ============ NGUZO ZA ZIADA (kutoka Excel) ============
// Kila nguzo baada ya jina na namba inakuwa alama ya ujumbe, mfano:
// nguzo "kiasi" -> [KIASI], nguzo "kodi" -> [KODI]
let columnHeaders = [];

// Kama hakuna nguzo za Excel, tunaruhusu safu ya tatu iwe [KODI] (kama zamani)
function headersZaSasa(){
  return columnHeaders.length ? columnHeaders : ['KODI'];
}

function escapeRegExp(s){
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ============ EXCEL UPLOAD ============
function handleExcelUpload(event){
  const file = event.target.files[0];
  if(!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try{
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(sheet);
      const headerRow = (XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] || []).map(h => String(h));

      if(rows.length === 0){
        alert("Faili halina data yoyote!");
        return;
      }

      const keys = headerRow.filter(h => h.trim());
      const findKey = (list) => keys.find(k => list.includes(k.toLowerCase().trim()));

      const nameKey = findKey(['jina', 'majina', 'name', 'first name', 'first_name', 'full name', 'jina kamili']);
      const lastNameKey = findKey(['last name', 'last_name', 'jina la pili']);
      const phoneKey = findKey(['namba ya simu', 'namba za simu', 'namba', 'simu', 'phone', 'phone number', 'mobile', 'phone_number']);

      // Nguzo zote zilizobaki ni nguzo za ziada (kiasi, ahadi, baki, kodi, aina ya kadi...)
      const extraKeys = keys.filter(k => k !== nameKey && k !== lastNameKey && k !== phoneKey);
      columnHeaders = extraKeys.map(k => k.trim().toUpperCase());

      const lines = [];
      rows.forEach(row => {
        let name = nameKey ? String(row[nameKey] ?? '').trim() : '';
        if(lastNameKey && row[lastNameKey]){
          name = `${name} ${String(row[lastNameKey]).trim()}`.trim();
        }
        const phone = phoneKey ? String(row[phoneKey] ?? '').trim() : '';
        if(!name && !phone) return;

        const extras = extraKeys.map(k => String(row[k] ?? '').trim());
        if(extras.length){
          // Tab inatumika kutenganisha, kwa sababu kiasi kinaweza kuwa na koma (1,000,000)
          lines.push([name, phone, ...extras].join('\t'));
        }else{
          lines.push(`${name}${phone ? ', ' + phone : ''}`);
        }
      });

      document.getElementById('recipientsInput').value = lines.join('\n');
      const alama = ['[JINA]'].concat(columnHeaders.map(h => `[${h}]`)).join(' ');
      document.getElementById('listStatusMsg').textContent =
        `Watu ${lines.length} wamepakiwa kutoka Excel. Alama unazoweza kutumia: ${alama}`;
    }catch(err){
      console.error(err);
      alert('Imeshindikana kusoma faili la Excel: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ============ ORODHA ZILIZOHIFADHIWA (Firestore) ============
async function loadSavedListsDropdown(){
  const select = document.getElementById('savedListsSelect');
  if(!select) return;
  try{
    const snap = await db.collection('contactLists').orderBy('createdAt', 'desc').get();
    select.innerHTML = '<option value="">— Chagua orodha —</option>';
    snap.forEach(doc => {
      const opt = document.createElement('option');
      opt.value = doc.id;
      opt.textContent = doc.data().name;
      select.appendChild(opt);
    });
  }catch(err){
    console.error('Imeshindikana kupakia orodha zilizohifadhiwa:', err);
  }
}

async function saveCurrentList(){
  const listName = document.getElementById('listNameInput').value.trim();
  const recipients = parseRecipients();
  const statusMsg = document.getElementById('listStatusMsg');

  if(!listName){
    alert('Andika jina la orodha kwanza (mfano: Ustawi Kidegembye).');
    return;
  }
  if(recipients.length === 0){
    alert('Hakuna majina wala namba za kuhifadhi - andika kwenye sehemu ya 2 kwanza.');
    return;
  }

  statusMsg.textContent = 'Inahifadhi...';
  try{
    await db.collection('contactLists').add({
      name: listName,
      columns: columnHeaders,
      recipients: recipients,
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    statusMsg.textContent = `Orodha "${listName}" imehifadhiwa (watu ${recipients.length}).`;
    document.getElementById('listNameInput').value = '';
    await loadSavedListsDropdown();
  }catch(err){
    console.error(err);
    statusMsg.textContent = 'Hitilafu: ' + err.message;
  }
}

async function loadSavedList(){
  const select = document.getElementById('savedListsSelect');
  const listId = select.value;
  if(!listId) return;

  try{
    const doc = await db.collection('contactLists').doc(listId).get();
    if(!doc.exists) return;
    const data = doc.data();

    // Nguzo za ziada: mpya (columns) au za zamani (code -> KODI)
    let cols = Array.isArray(data.columns) ? data.columns.slice() : [];
    const anyCode = data.recipients.some(r => r.code || (r.fields && r.fields.KODI));
    if(!cols.length && anyCode) cols = ['KODI'];
    columnHeaders = cols;

    const lines = data.recipients.map(r => {
      const f = r.fields || {};
      if(cols.length){
        const vals = cols.map(c => (c in f) ? f[c] : (c === 'KODI' ? (r.code || '') : ''));
        return [r.name, r.phone || '', ...vals].join('\t');
      }
      return r.phone ? `${r.name}, ${r.phone}` : r.name;
    });

    document.getElementById('recipientsInput').value = lines.join('\n');
    const alama = ['[JINA]'].concat(cols.map(h => `[${h}]`)).join(' ');
    document.getElementById('listStatusMsg').textContent =
      `Orodha "${data.name}" imepakiwa (watu ${lines.length}). Alama: ${alama}`;
  }catch(err){
    console.error(err);
    alert('Imeshindikana kupakia orodha: ' + err.message);
  }
}

// ============ SENDER SELECTOR ============
function onSenderChange(){
  const select = document.getElementById('senderSelect');
  const customRow = document.getElementById('customSenderRow');
  if(customRow) customRow.style.display = select.value === 'custom' ? 'flex' : 'none';
}

function getSelectedSender(){
  const select = document.getElementById('senderSelect');
  if(!select) return 'INFO';
  if(select.value === 'custom'){
    return document.getElementById('customSender').value.trim().toUpperCase();
  }
  return select.value;
}

// ============ HESABU HERUFI ============
function updateCharCount(){
  const msgEl = document.getElementById('messageInput');
  if(!msgEl) return;
  const len = msgEl.value.length;
  let parts, limit;
  if(len <= 160){
    parts = len === 0 ? 0 : 1;
    limit = 160;
  }else{
    parts = Math.ceil(len / 153);
    limit = parts * 153;
  }
  const charCountEl = document.getElementById('charCount');
  if(charCountEl){
    charCountEl.textContent = `Herufi: ${len}/${limit} (SMS ${parts || 1})`;
  }
}
const msgInput = document.getElementById('messageInput');
if(msgInput){
  msgInput.addEventListener('input', updateCharCount);
  updateCharCount();
}

// ============ PARSE RECIPIENTS ============
// Mstari: jina, simu  (au kwa Excel yenye nguzo za ziada: jina <tab> simu <tab> nguzo1 <tab> nguzo2 ...)
function parseRecipients(){
  const raw = document.getElementById('recipientsInput').value;
  const lines = raw.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const headers = headersZaSasa();

  return lines.map(line => {
    const parts = (line.includes('\t') ? line.split('\t') : line.split(',')).map(p => p.trim());
    const fields = {};
    headers.forEach((h, i) => { fields[h] = parts[2 + i] || ''; });
    return { name: parts[0] || '', phone: parts[1] || '', fields };
  }).filter(r => r.name.length > 0);
}

function normalizePhone(phone){
  let clean = phone.replace(/\D/g, '');
  if(clean.startsWith('0')){
    clean = '255' + clean.slice(1);
  }
  return clean;
}

// ============ TUMA SMS MOJA ============
// NB: Beem HAIZUII maombi kutoka kivinjari (hakuna CORS) - proxy ya nje
// (kama corsproxy.io) SIYO lazima na inaweza kuondoa "Authorization" header,
// ndiyo maana ilisababisha SMS zote kushindwa bila credits kupungua.
async function sendOneSMS(sender, phone, message){
  const authHeader = 'Basic ' + btoa(`${BEEM_API_KEY}:${BEEM_SECRET_KEY}`);

  try{
    const response = await fetch('https://apisms.beem.africa/v1/send', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': authHeader
      },
      body: JSON.stringify({
        source_addr: sender,
        encoding: 0,
        message: message,
        recipients: [{ recipient_id: "1", dest_addr: normalizePhone(phone) }]
      })
    });

    let data = null;
    try{ data = await response.json(); }catch(e){ data = null; }
    console.log("Beem API Response:", data);

    const ok = response.ok && !(data && data.successful === false);
    return { ok, data };
  }catch(err){
    console.error("SMS Send Error:", err);
    return { ok: false, error: err.message };
  }
}

// Sababu ya kushindwa (kutoka Beem au mtandao)
function sababuYaKushindwa(result){
  if(result.error) return result.error;
  const d = result.data || {};
  const s = d.message || d.error || d.description || d.msg || (d.code ? 'code ' + d.code : '');
  return String(s || '').slice(0, 60);
}

// ============ TUMA KWA WOTE ============
async function sendBulkSMS(){
  const sender = getSelectedSender();
  if(!sender){
    alert('Chagua au andika Sender Name kwanza.');
    return;
  }

  const allRecipients = parseRecipients();
  const validRecipients = allRecipients.filter(r => r.phone && normalizePhone(r.phone).length >= 10);

  if(validRecipients.length === 0){
    alert('Hakuna wapokeaji wenye namba za simu halali! Jaza namba za simu kwanza.');
    return;
  }

  const messageTemplate = document.getElementById('messageInput').value.trim();
  if(!messageTemplate){
    alert('Andika ujumbe kwanza.');
    return;
  }

  // Alama zilizotumika kwenye ujumbe, mfano [JINA], [KIASI]
  const used = [...new Set(
    (messageTemplate.match(/\[([^\]]+)\]/g) || []).map(s => s.slice(1, -1).trim().toUpperCase())
  )];
  const known = ['JINA', ...headersZaSasa()];
  const unknown = used.filter(u => !known.includes(u));
  if(unknown.length){
    alert(
      'Alama hizi hazijulikani: ' + unknown.map(u => `[${u}]`).join(', ') + '.\n\n' +
      'Alama zinazopatikana sasa: ' + known.map(k => `[${k}]`).join(' ') + '\n\n' +
      'Pakia Excel yenye nguzo hizo, au ondoa alama hizo kwenye ujumbe.'
    );
    return;
  }
  const needed = used.filter(u => u !== 'JINA');

  const sendBtn = document.getElementById('sendBtn');
  const statusMsg = document.getElementById('statusMsg');
  const resultsSection = document.getElementById('resultsSection');
  const resultsTable = document.getElementById('resultsTable');

  sendBtn.disabled = true;
  resultsSection.style.display = 'block';
  resultsTable.innerHTML = '';

  let sent = 0, failed = 0;

  for(let i = 0; i < validRecipients.length; i++){
    const r = validRecipients[i];
    statusMsg.textContent = `Inatuma ${i + 1}/${validRecipients.length}...`;

    // Ujumbe unahitaji alama (mf. kiasi) lakini mtu huyu hana: usitume
    const missing = needed.filter(h => !(r.fields[h] || '').length);
    if(missing.length){
      failed++;
      const rowMissing = document.createElement('div');
      rowMissing.className = 'result-row';
      rowMissing.innerHTML = `
        <span class="r-name">${escapeHtml(r.name)}</span>
        <span class="r-phone">${escapeHtml(r.phone)}</span>
        <span class="r-status fail">✕ Hana ${escapeHtml(missing.join(', '))}</span>
      `;
      resultsTable.appendChild(rowMissing);
      continue;
    }

    let personalizedMessage = messageTemplate.replace(/\[JINA\]/gi, () => r.name);
    needed.forEach(h => {
      const re = new RegExp('\\[\\s*' + escapeRegExp(h) + '\\s*\\]', 'gi');
      personalizedMessage = personalizedMessage.replace(re, () => r.fields[h]);
    });

    const result = await sendOneSMS(sender, r.phone, personalizedMessage);

    let sababu = '';
    if(result.ok){ sent++; } else {
      failed++;
      sababu = sababuYaKushindwa(result);
      console.error(`SMS kwa ${r.name} (${r.phone}) imeshindwa:`, result.data || result.error);
    }

    const row = document.createElement('div');
    row.className = 'result-row';
    row.innerHTML = `
      <span class="r-name">${escapeHtml(r.name)}</span>
      <span class="r-phone">${escapeHtml(r.phone)}</span>
      <span class="r-status ${result.ok ? 'ok' : 'fail'}" title="${escapeHtml(sababu)}">${result.ok ? '✓ Imetumwa' : '✕ Imeshindwa' + (sababu ? ': ' + escapeHtml(sababu) : '')}</span>
    `;
    resultsTable.appendChild(row);
  }

  statusMsg.textContent = `Kamili: ${sent} zimetumwa, ${failed} zimeshindwa.`;
  sendBtn.disabled = false;
}

function escapeHtml(str){
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ============ INIT ============
loadSavedListsDropdown();
