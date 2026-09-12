// ============ BEEM CONFIG ============
const BEEM_API_KEY = "cd9893f5aadb3b83";
const BEEM_SECRET_KEY = "NzAzMjhlYmZhMzI3YmY1OGMwNDU3ZDBjNTQ1Mjg0MzJjYjIxZDZmOWZlYmVlNmFlZWM3ZjlmYTA5MDMwNDA2NA==";

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

      if(rows.length === 0){
        alert("Faili halina data yoyote!");
        return;
      }

      const lines = rows.map(row => {
        // Find name column dynamically
        const nameKey = Object.keys(row).find(k => 
          ['jina', 'name', 'first name', 'first_name', 'full name'].includes(k.toLowerCase().trim())
        );
        const lastNameKey = Object.keys(row).find(k => 
          ['last name', 'last_name', 'jina la pili'].includes(k.toLowerCase().trim())
        );
        
        // Find phone column dynamically
        const phoneKey = Object.keys(row).find(k => 
          ['namba', 'phone', 'phone number', 'mobile', 'namba ya simu', 'phone_number'].includes(k.toLowerCase().trim())
        );

        let name = nameKey ? row[nameKey] : '';
        if(lastNameKey && row[lastNameKey]){
          name = `${name} ${row[lastNameKey]}`.trim();
        }
        const phone = phoneKey ? row[phoneKey] : '';

        return `${name}${phone ? ', ' + phone : ''}`;
      }).filter(line => line.trim().length > 0);

      document.getElementById('recipientsInput').value = lines.join('\n');
      document.getElementById('listStatusMsg').textContent = `Watu ${lines.length} wamepakiwa kutoka Excel.`;
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
  const recipients = parseRecipients(); // Sasa inakubali hata ikiwa na majina pekee
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
    const lines = data.recipients.map(r => r.phone ? `${r.name}, ${r.phone}` : r.name);
    document.getElementById('recipientsInput').value = lines.join('\n');
    document.getElementById('listStatusMsg').textContent = `Orodha "${data.name}" imepakiwa (watu ${lines.length}).`;
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
  if(select.value === 'custom'){
    return document.getElementById('customSender').value.trim().toUpperCase();
  }
  return select.value;
}

// ============ HESABU HERUFI ============
function updateCharCount(){
  const len = document.getElementById('messageInput').value.length;
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
function parseRecipients(){
  const raw = document.getElementById('recipientsInput').value;
  const lines = raw.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  
  return lines.map(line => {
    const parts = line.split(',').map(p => p.trim());
    const name = parts[0] || '';
    const phone = parts[1] || '';
    return { name, phone };
  }).filter(r => r.name.length > 0); // Inakubali mradi jina lipo
}

function normalizePhone(phone){
  let clean = phone.replace(/\D/g, '');
  if(clean.startsWith('0')){
    clean = '255' + clean.slice(1);
  }
  return clean;
}

// ============ TUMA SMS MOJA ============
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
    const data = await response.json();
    return { ok: response.ok, data };
  }catch(err){
    return { ok: false, error: err.message };
  }
}

// ============ TUMA KWA WOTE ============
async function sendBulkSMS(){
  const sender = getSelectedSender();
  if(!sender){
    alert('Chagua au andika Sender Name kwanza.');
    return;
  }

  const allRecipients = parseRecipients();
  // Filter walio na namba za simu pekee kabla ya kutuma SMS
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

    const personalizedMessage = messageTemplate.replace(/\[JINA\]/g, r.name);
    const result = await sendOneSMS(sender, r.phone, personalizedMessage);

    if(result.ok){ sent++; } else { failed++; }

    const row = document.createElement('div');
    row.className = 'result-row';
    row.innerHTML = `
      <span class="r-name">${escapeHtml(r.name)}</span>
      <span class="r-phone">${escapeHtml(r.phone)}</span>
      <span class="r-status ${result.ok ? 'ok' : 'fail'}">${result.ok ? '✓ Imetumwa' : '✕ Imeshindwa'}</span>
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
