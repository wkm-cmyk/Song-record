
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Song-record 出版物記錄庫</title>
  <style>
    * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    body { background-color: #f5f7fa; margin: 0; padding: 20px; color: #333; }
    .container { max-width: 900px; margin: 0 auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
    h1 { color: #1a56db; font-size: 24px; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }
    .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 15px; }
    input { width: 100%; padding: 10px 14px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; outline: none; transition: border-color 0.2s; }
    input:focus { border-color: #1a56db; }
    .btn-group { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; }
    .btn { padding: 10px 18px; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; transition: opacity 0.2s; display: inline-flex; align-items: center; gap: 6px; }
    .btn:hover { opacity: 0.9; }
    .btn-green { background-color: #10b981; color: white; }
    .btn-blue { background-color: #0284c7; color: white; }
    .btn-danger { background-color: #ef4444; color: white; padding: 4px 8px; font-size: 12px; }
    .btn-edit { background-color: #f59e0b; color: white; padding: 4px 8px; font-size: 12px; }
    .search-bar { width: 100%; padding: 12px 16px; border: 1px solid #3b82f6; border-radius: 8px; margin-bottom: 20px; font-size: 15px; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th, td { text-align: left; padding: 12px; border-bottom: 1px solid #e5e7eb; font-size: 14px; }
    th { background-color: #f8fafc; color: #475569; font-weight: 600; }
    .status-msg { text-align: center; color: #ef4444; font-weight: 600; margin: 15px 0; display: none; }
    .loading { text-align: center; color: #64748b; padding: 20px; }
  </style>
</head>
<body>

<div class="container">
  <h1>📚 出版物記錄與搜尋庫</h1>

  <!-- 輸入表單 -->
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

  <!-- 搜尋框 -->
  <input type="text" id="searchInput" class="search-bar" placeholder="🔍 搜尋書名、中英文歌名、作曲家或出版社..." oninput="filterData()">

  <!-- 狀態提示 -->
  <div id="statusMsg" class="status-msg">❌ 資料載入失敗，請檢查網路連線或 API 網址。</div>

  <!-- 資料表格 -->
  <table>
    <thead>
      <tr>
        <th>出版社</th>
        <th>書名</th>
        <th>中文歌名</th>
        <th>英文歌名</th>
        <th>作曲家</th>
        <th>操作</th>
      </tr>
    </thead>
    <tbody id="tableBody">
      <tr><td colspan="6" class="loading">⏳ 資料載入中...</td></tr>
    </tbody>
  </table>
</div>

<script>
  // ⚠️ 請替換為您的 Google Apps Script Web App URL
  const API_URL = "YOUR_GAS_API_URL";
  let allRecords = [];

  // 1. 初始化讀取資料
  async function loadData() {
    try {
      document.getElementById('statusMsg').style.display = 'none';
      const res = await fetch(API_URL);
      allRecords = await res.json();
      renderTable(allRecords);
    } catch (err) {
      console.error(err);
      document.getElementById('statusMsg').style.display = 'block';
      document.getElementById('tableBody').innerHTML = '<tr><td colspan="6" style="text-align:center; color:#ef4444;">資料載入失敗</td></tr>';
    }
  }

  // 2. 渲染表格數據
  function renderTable(records) {
    const tbody = document.getElementById('tableBody');
    if (records.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;">尚無記錄數據</td></tr>';
      return;
    }

    tbody.innerHTML = records.map(item => `
      <tr>
        <td>${escapeHtml(item.publisher)}</td>
        <td>${escapeHtml(item.bookName)}</td>
        <td>${escapeHtml(item.songNameZh)}</td>
        <td>${escapeHtml(item.songNameEn)}</td>
        <td>${escapeHtml(item.composer)}</td>
        <td>
          <button onclick="editRecord('${item.id}')" class="btn btn-edit">編輯</button>
          <button onclick="deleteRecord('${item.id}')" class="btn btn-danger">刪除</button>
        </td>
      </tr>
    `).join('');
  }

  // 3. 表單提交（新增或編輯）
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

  // 4. 編輯點擊填入
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
  }

  // 5. 刪除記錄
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

  // 6. 重置表單
  function resetForm() {
    document.getElementById('recordForm').reset();
    document.getElementById('editId').value = '';
    const submitBtn = document.getElementById('submitBtn');
    submitBtn.innerText = "➕ 新增記錄";
    submitBtn.className = "btn btn-green";
  }

  // 7. 前端即時搜尋
  function filterData() {
    const query = document.getElementById('searchInput').value.toLowerCase();
    const filtered = allRecords.filter(r => 
      (r.publisher && r.publisher.toLowerCase().includes(query)) ||
      (r.bookName && r.bookName.toLowerCase().includes(query)) ||
      (r.songNameZh && r.songNameZh.toLowerCase().includes(query)) ||
      (r.songNameEn && r.songNameEn.toLowerCase().includes(query)) ||
      (r.composer && r.composer.toLowerCase().includes(query))
    );
    renderTable(filtered);
  }

  // 8. 匯出 Excel (CSV 格式)
  function exportToExcel() {
    if (allRecords.length === 0) {
      alert("目前沒有可匯出的資料！");
      return;
    }
    let csvContent = "\uFEFF出版社,書名,中文歌名,英文歌名,作曲家\n";
    allRecords.forEach(r => {
      csvContent += `"${r.publisher}","${r.bookName}","${r.songNameZh}","${r.songNameEn}","${r.composer}"\n`;
    });
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.setAttribute("download", `出版物記錄庫_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // XSS 防護輔助函數
  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  // 網頁開啟時自動載入
  loadData();
</script>

</body>
</html>
