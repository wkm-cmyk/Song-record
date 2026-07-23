<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Song-record 出版物記錄庫</title>
  <style>
    * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    body { background-color: #f5f7fa; margin: 0; padding: 20px; color: #333; }
    .container { max-width: 1100px; margin: 0 auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
    h1 { color: #1a56db; font-size: 24px; margin-bottom: 20px; }
    .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 15px; }
    input { width: 100%; padding: 10px 14px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; outline: none; }
    input:focus { border-color: #1a56db; }
    .btn-group { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; }
    .btn { padding: 10px 18px; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; }
    .btn:hover { opacity: 0.9; }
    .btn-green { background-color: #10b981; color: white; }
    .btn-blue { background-color: #0284c7; color: white; }
    .btn-toggle { background-color: #6366f1; color: white; padding: 6px 12px; font-size: 13px; }
    .btn-danger { background-color: #ef4444; color: white; padding: 5px 10px; font-size: 12px; }
    .btn-edit { background-color: #f59e0b; color: white; padding: 5px 10px; font-size: 12px; }
    .search-bar { width: 100%; padding: 12px 16px; border: 1px solid #3b82f6; border-radius: 8px; margin-bottom: 20px; font-size: 15px; }
    
    .publisher-card { border: 1px solid #e5e7eb; border-radius: 8px; margin-bottom: 16px; overflow: hidden; background: #fff; }
    .publisher-header { background-color: #f8fafc; padding: 14px 20px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #e5e7eb; }
    .publisher-title { font-size: 16px; font-weight: 700; color: #1e293b; display: flex; align-items: center; gap: 8px; }
    .book-count { font-size: 12px; background: #e0e7ff; color: #3730a3; padding: 2px 8px; border-radius: 12px; font-weight: normal; }
    
    .publisher-content { display: none; padding: 0; }
    .publisher-content.show { display: block; }
    
    /* 均勻分頁與滿版設定 */
    table { width: 100%; table-layout: fixed; border-collapse: collapse; }
    th, td { text-align: left; padding: 12px 10px; border-bottom: 1px solid #e5e7eb; border-right: 1px solid #f1f5f9; font-size: 14px; word-break: break-word; }
    th:last-child, td:last-child { border-right: none; }
    th { background-color: #f8fafc; color: #475569; font-weight: 600; font-size: 13px; }
    
    /* 設定每一欄的平均比例 */
    .col-seq { width: 7%; text-align: center; }
    .col-book { width: 22%; }
    .col-zh { width: 22%; }
    .col-en { width: 22%; }
    .col-composer { width: 17%; }
    .col-action { width: 10%; text-align: center; }

    .action-btn-group { display: flex; flex-direction: column; gap: 4px; align-items: center; }
    .seq-badge { background: #f3f4f6; color: #4b5563; font-weight: 600; padding: 3px 6px; border-radius: 4px; font-size: 12px; display: inline-block; }
    .status-msg { text-align: center; color: #ef4444; font-weight: 600; margin: 15px 0; display: none; }
    .loading { text-align: center; color: #64748b; padding: 20px; }
  </style>
</head>
<body>

<div class="container">
  <h1>📚 出版物記錄與搜尋庫</h1>

  <form id="recordForm" onsubmit="handleFormSubmit(event)">
    <input type="hidden" id="editId">
    <div class="form-grid">
      <input type="text" id="publisher" placeholder="出版社 (Publisher) *" required>
      <input type="text" id="bookName" placeholder="書名 (Book Name) *" required>
      <input type="text" id="songNameZh" placeholder="中文歌名 (選填)">
      <input type="text" id="songNameEn" placeholder="英文歌名 (選填)">
      <input type="text" id="composer" placeholder="作曲家 (Composer) (選填)">
    </div>
    
    <div class="btn-group">
      <button type="submit" id="submitBtn" class="btn btn-green">➕ 新增記錄</button>
      <button type="button" onclick="exportToExcel()" class="btn btn-blue">📊 匯出至 Excel</button>
    </div>
  </form>

  <input type="text" id="searchInput" class="search-bar" placeholder="🔍 搜尋出版社、書名、中英文歌名或作曲家..." oninput="filterData()">

  <div id="statusMsg" class="status-msg">❌ 資料載入失敗，請檢查網路連線或 API 網址。</div>

  <div id="publisherList">
    <div class="loading">⏳ 資料載入中...</div>
  </div>
</div>

<script>
  const API_URL = "https://script.google.com/macros/s/AKfycbwe_NyV5IEWyDrJFU98jt9lT2Pltb4Vi4ZNdjUYOybovj66B4GbEz1RF_taPtmS0pbU/exec";
  let allRecords = [];

  async function loadData() {
    try {
      document.getElementById('statusMsg').style.display = 'none';
      const res = await fetch(API_URL);
      allRecords = await res.json();
      renderGroupedData(allRecords);
    } catch (err) {
      console.error(err);
      document.getElementById('statusMsg').style.display = 'block';
      document.getElementById('publisherList').innerHTML = '<div style="text-align:center; color:#ef4444; padding:20px;">資料載入失敗</div>';
    }
  }

  function renderGroupedData(records) {
    const container = document.getElementById('publisherList');
    if (records.length === 0) {
      container.innerHTML = '<div style="text-align:center; padding:20px; color:#64748b;">尚無記錄數據</div>';
      return;
    }

    const grouped = {};
    records.forEach(item => {
      const pub = item.publisher || "未分類出版社";
      if (!grouped[pub]) {
        grouped[pub] = [];
      }
      grouped[pub].push(item);
    });

    const publishers = Object.keys(grouped).sort((a, b) => 
      a.localeCompare(b, ['zh-HK', 'zh-TW', 'en'], { numeric: true, sensitivity: 'base' })
    );

    let html = '';
    publishers.forEach((pubName, index) => {
      const pubRecords = grouped[pubName];
      const cardId = `pub-content-${index}`;
      const btnId = `pub-btn-${index}`;

      html += `
        <div class="publisher-card">
          <div class="publisher-header">
            <div class="publisher-title">
              🏢 ${escapeHtml(pubName)}
              <span class="book-count">${pubRecords.length} 筆記錄</span>
            </div>
            <button id="${btnId}" onclick="togglePublisher('${cardId}', '${btnId}')" class="btn btn-toggle">
              ▼ 顯示書名與內容
            </button>
          </div>
          <div id="${cardId}" class="publisher-content">
            <table>
              <thead>
                <tr>
                  <th class="col-seq">次序</th>
                  <th class="col-book">書名</th>
                  <th class="col-zh">中文歌名</th>
                  <th class="col-en">英文歌名</th>
                  <th class="col-composer">作曲家</th>
                  <th class="col-action">操作</th>
                </tr>
              </thead>
              <tbody>
                ${pubRecords.map(item => `
                  <tr>
                    <td class="col-seq"><span class="seq-badge">#${item.seqNumber || '-'}</span></td>
                    <td class="col-book"><strong>${escapeHtml(item.bookName)}</strong></td>
                    <td class="col-zh">${escapeHtml(item.songNameZh)}</td>
                    <td class="col-en">${escapeHtml(item.songNameEn)}</td>
                    <td class="col-composer">${escapeHtml(item.composer)}</td>
                    <td class="col-action">
                      <div class="action-btn-group">
                        <button onclick="editRecord('${item.id}')" class="btn btn-edit">編輯</button>
                        <button onclick="deleteRecord('${item.id}')" class="btn btn-danger">刪除</button>
                      </div>
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;
    });

    container.innerHTML = html;
  }

  function togglePublisher(contentId, btnId) {
    const content = document.getElementById(contentId);
    const btn = document.getElementById(btnId);
    if (content.classList.contains('show')) {
      content.classList.remove('show');
      btn.innerHTML = '▼ 顯示書名與內容';
    } else {
      content.classList.add('show');
      btn.innerHTML = '▲ 隱藏內容';
    }
  }

  async function handleFormSubmit(e) {
    e.preventDefault();
    const editId = document.getElementById('editId').value;
    const payload = {
      action: editId ? 'edit' : 'add',
      id: editId,
      publisher: document.getElementById('publisher').value.trim(),
      bookName: document.getElementById('bookName').value.trim(),
      songNameZh: document.getElementById('songNameZh').value.trim(),
      songNameEn: document.getElementById('songNameEn').value.trim(),
      composer: document.getElementById('composer').value.trim()
    };

    const submitBtn = document.getElementById('submitBtn');
    submitBtn.disabled = true;
    submitBtn.innerText = "處理中...";

    try {
      await fetch(API_URL, {
        method: 'POST',
        body: JSON.stringify(payload)
      });
      resetForm();
      loadData();
    } catch (err) {
      alert("儲存失敗：" + err.message);
    } finally {
      submitBtn.disabled = false;
    }
  }

  function editRecord(id) {
    const item = allRecords.find(r => String(r.id) === String(id));
    if (!item) return;
    document.getElementById('editId').value = item.id;
    document.getElementById('publisher').value = item.publisher;
    document.getElementById('bookName').value = item.bookName;
    document.getElementById('songNameZh').value = item.songNameZh;
    document.getElementById('songNameEn').value = item.songNameEn;
    document.getElementById('composer').value = item.composer;
    
    const submitBtn = document.getElementById('submitBtn');
    submitBtn.innerText = "✏️ 儲存修改";
    submitBtn.className = "btn btn-edit";
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  async function deleteRecord(id) {
    if (!confirm("確定要刪除這筆記錄嗎？")) return;
    try {
      await fetch(API_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'delete', id: id })
      });
      loadData();
    } catch (err) {
      alert("刪除失敗：" + err.message);
    }
  }

  function resetForm() {
    document.getElementById('recordForm').reset();
    document.getElementById('editId').value = '';
    const submitBtn = document.getElementById('submitBtn');
    submitBtn.innerText = "➕ 新增記錄";
    submitBtn.className = "btn btn-green";
  }

  function filterData() {
    const query = document.getElementById('searchInput').value.toLowerCase();
    if (!query) {
      renderGroupedData(allRecords);
      return;
    }

    const filtered = allRecords.filter(r => 
      (r.publisher && r.publisher.toLowerCase().includes(query)) ||
      (r.bookName && r.bookName.toLowerCase().includes(query)) ||
      (r.songNameZh && r.songNameZh.toLowerCase().includes(query)) ||
      (r.songNameEn && r.songNameEn.toLowerCase().includes(query)) ||
      (r.composer && r.composer.toLowerCase().includes(query))
    );

    renderGroupedData(filtered);
    document.querySelectorAll('.publisher-content').forEach(el => el.classList.add('show'));
    document.querySelectorAll('.btn-toggle').forEach(btn => btn.innerHTML = '▲ 隱藏內容');
  }

  function exportToExcel() {
    if (allRecords.length === 0) {
      alert("目前沒有可匯出的資料！");
      return;
    }
    let csvContent = "\uFEFF次序,出版社,書名,中文歌名,英文歌名,作曲家\n";
    allRecords.forEach(r => {
      csvContent += `"${r.seqNumber || ''}","${r.publisher}","${r.bookName}","${r.songNameZh}","${r.songNameEn}","${r.composer}"\n`;
    });
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.setAttribute("download", `出版物記錄庫_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  loadData();
</script>

</body>
</html>
