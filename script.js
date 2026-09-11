// ============ BEEM CONFIG ============
// !! MUHIMU !! Hii ni kwa MAJARIBIO/MATUMIZI YAKO BINAFSI TU (siyo ukurasa wa umma).
// Usishiriki link ya Pen hii na mtu yeyote asiyekusudiwa, kwa sababu funguo zinaonekana
// wazi kwenye msimbo huu.
const BEEM_API_KEY = "cd9893f5aadb3b83";
const BEEM_SECRET_KEY = "NzAzMjhlYmZhMzI3YmY1OGMwNDU3ZDBjNTQ1Mjg0MzJjYjIxZDZmOWZlYmVlNmFlZWM3ZjlmYTA5MDMwNDA2NA==";

// ============ FIREBASE CONFIG (sawa na miradi mingine ya JOHCARDS) ============
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
      const rows = XLSX.utils.sheet_to_json(sheet); // inatumia safu ya kwanza kama headers

      const lines = rows.map(row => {
        // Kubali majina tofauti ya column (Jina/jina/Name, Namba/namba/Phone)
        const name = row['Jina'] || row['jina'] || row['Name'] || row['name'] || '';
        const phone = row['Namba'] || row['namba'] || row['Phone'] || row['phone'] || '';
        return `${name}, ${phone}`;
      }).filter(line => line.trim() !== ',');

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
  const recipients = parseRecipients();
  const statusMsg = document.getElementById('listStatusMsg');

  if(!listName){
    alert('Andika jina la orodha kwanza (mfano: Wazazi wa Kidegembye).');
    return;
  }
  if(recipients.length === 0){
    alert('Hakuna majina/namba za kuhifadhi - jaza sehemu ya 2 kwanza.');
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
    const lines = data.recipients.map(r => `${r.name}, ${r.phone}`);
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
  customRow.style.display = select.value === 'custom' ? 'flex' : 'none';
}

function getSelectedSender(){
  const select = document.getElementById('senderSelect');
  if(select.value === 'custom'){
    return document.getElementById('customSender').value.trim().toUpperCase();
  }
  return select.value;
}

// ============ HESABU HERUFI (kama Beem/mitandao halisi) ============
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
  document.getElementById('charCount').textContent =
    `Herufi: ${len}/${limit} (SMS ${parts || 1})`;
}
document.getElementById('messageInput').addEventListener('input', updateCharCount);
updateCharCount(); // onyesha 0 mara ukurasa unapopakia

// ============ PARSE RECIPIENTS ============
function parseRecipients(){
  const raw = document.getElementById('recipientsInput').value;
  const lines = raw.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  return lines.map(line => {
    const parts = line.split(',').map(p => p.trim());
    return { name: parts[0] || '', phone: parts[1] || '' };
  }).filter(r => r.name && r.phone);
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

  const recipients = parseRecipients();
  if(recipients.length === 0){
    alert('Andika majina na namba za wapokeaji angalau mmoja.');
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

  for(let i = 0; i < recipients.length; i++){
    const r = recipients[i];
    statusMsg.textContent = `Inatuma ${i + 1}/${recipients.length}...`;

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