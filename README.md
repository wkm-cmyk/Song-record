
<html lang="zh-HK">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>出版物記錄與搜尋庫</title>
    <!-- 引入 SheetJS 函式庫以支援匯出 Excel -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; color: #333; }
        .container { max-width: 900px; margin: auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h1, h2 { color: #2c3e50; }
        .form-group { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px; }
        .form-group input { flex: 1 1 180px; padding: 10px; border: 1px solid #ccc; border-radius: 4px; }
        button { padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
        .btn-add { background-color: #27ae60; color: white; }
        .btn-add:disabled { background-color: #95a5a6; cursor: not-allowed; }
        .btn-save-edit { background-color: #e67e22; color: white; }
        .btn-cancel-edit { background-color: #95a5a6; color: white; }
        .btn-export { background-color: #2980b9; color: white; }
        
        /* 操作按鈕容器：強制上下（垂直）排列並靠右 */
        .btn-action-group {
            display: flex;
            flex-direction: column;
            gap: 6px;
            align-items: flex-end;
            justify-content: center;
        }

        /* 修改按鈕樣式 */
        .btn-edit { 
            background-color: #f39c12; 
            color: white; 
            padding: 5px 10px; 
            font-size: 0.85em; 
            border-radius: 4px; 
            cursor: pointer; 
            border: none !important; 
            white-space: nowrap !important; 
            display: inline-block;
        }
        .btn-edit:hover { background-color: #d68910; }

        /* 刪除按鈕樣式 */
        .btn-delete { 
            background-color: #e74c3c; 
            color: white; 
            padding: 5px 10px; 
            font-size: 0.85em; 
            border-radius: 4px; 
            cursor: pointer; 
            border: none !important; 
            white-space: nowrap !important; 
            display: inline-block;
        }
        .btn-delete:hover { background-color: #c0392b; }
        .btn-delete:disabled { background-color: #f5b7b1; cursor: not-allowed; }
        
        .search-bar { width: 100%; padding: 12px; margin: 20px 0; border: 2px solid #3498db; border-radius: 4px; font-size: 16px; box-sizing: border-box; }
        .publisher-group { margin-bottom: 30px; border: 1px solid #e1e8ed; border-radius: 6px; overflow: hidden; background: #fff; width: 100%; }
        .publisher-title { background-color: #34495e; color: white; padding: 12px 15px; margin: 0; font-size: 1.1em; font-weight: 600; }
        
        /* 強制表格 100% 撐滿整個外框 */
        .publisher-group table { 
            display: table !important;
            width: 100% !important; 
            min-width: 100% !important;
            border-collapse: collapse !important; 
            table-layout: fixed !important; 
            margin: 0 !important;
            border: none !important;
        }

        .publisher-group thead { display: table-header-group !important; width: 100% !important; }
        .publisher-group tbody { display: table-row-group !important; width: 100% !important; }
        .publisher-group tr { display: table-row !important; width: 100% !important; }

        /* 徹底清空直行網格線，僅留水平底線 */
        .publisher-group th, 
        .publisher-group td { 
            display: table-cell !important;
            padding: 12px 15px !important; 
            text-align: left !important; 
            border: none !important; 
            border-bottom: 1px solid #e1e8ed !important; 
            word-wrap: break-word !important;
            vertical-align: middle !important;
        }

        .publisher-group th { 
            background-color: #f8f9fa !important; 
            font-weight: 600 !important; 
            color: #2c3e50 !important; 
        }

        /* 最右側按鈕欄位靠右對齊 */
        .publisher-group th:last-child, 
        .publisher-group td:last-child { 
            text-align: right !important; 
            padding-right: 20px !important;
        }

        tr:hover { background-color: #f8f9fa; }
        tr:last-child td { border-bottom: none !important; }
    </style>
</head>
<body>

<div class="container">
    <h1>📚 出版物記錄與搜尋庫</h1>

    <!-- 輸入表單 -->
    <div class="form-group">
        <input type="text" id="publisher" list="publisherList" placeholder="出版社 (Publisher) *" required autocomplete="off">
        <datalist id="publisherList"></datalist>

        <input type="text" id="bookName" list="bookNameList" placeholder="書名 (Book Name) *" required autocomplete="off">
        <datalist id="bookNameList"></datalist>

        <input type="text" id="songNameZh" list="songNameZhList" placeholder="中文歌名 (選填)" autocomplete="off">
        <datalist id="songNameZhList"></datalist>

        <input type="text" id="songNameEn" list="songNameEnList" placeholder="英文歌名 (選填)" autocomplete="off">
        <datalist id="songNameEnList"></datalist>

        <input type="text" id="composer" list="composerList" placeholder="作曲家 (Composer) (選填)" autocomplete="off">
        <datalist id="composerList"></datalist>

        <button id="submitBtn" class="btn-add" onclick="handleSubmit()">➕ 新增記錄</button>
        <button id="cancelBtn" class="btn-cancel-edit" onclick="cancelEdit()" style="display: none;">✖️ 取消</button>
        <button class="btn-export" onclick="exportToExcel()">📊 匯出至 Excel</button>
    </div>

    <input type="text" id="searchInput" class="search-bar" placeholder="🔍 搜尋書名、中英文歌名、作曲家或出版社..." onkeyup="renderRecords()">

    <div id="resultsContainer"></div>
</div>

<script>
    const API_URL = 'https://script.google.com/macros/s/AKfycbxLNIX5-v2gyojTBAylf93lMpOxBM4B2G2xw5CG8ZfhHLw9YViVRE6W5y211u--N4H9/exec';
    
    let records = [];
    let editingId = null; // 當前正在修改的記錄 ID

    async function loadRecords() {
        const container = document.getElementById('resultsContainer');
        container.innerHTML = '<p style="text-align:center; color:#2980b9;">⏳ 正在從雲端載入資料，請稍候...</p>';
        try {
            const response = await fetch(API_URL);
            records = await response.json();
            updateDatalists();
            renderRecords();
        } catch (error) {
            console.error('載入失敗:', error);
            container.innerHTML = '<p style="text-align:center; color:red;">❌ 資料載入失敗，請檢查網路連線或 API 網址。</p>';
        }
    }

    function updateDatalists() {
        const fields = [
            { id: 'publisherList', key: 'publisher' },
            { id: 'bookNameList', key: 'bookName' },
            { id: 'songNameZhList', key: 'songNameZh' },
            { id: 'songNameEnList', key: 'songNameEn' },
            { id: 'composerList', key: 'composer' }
        ];

        fields.forEach(field => {
            const datalist = document.getElementById(field.id);
            if (!datalist) return;

            const uniqueValues = [...new Set(
                records
                    .map(r => r[field.key])
                    .filter(val => val && val.trim() !== '' && val !== '-')
            )];

            datalist.innerHTML = uniqueValues
                .map(val => `<option value="${val.replace(/"/g, '&quot;')}"></option>`)
                .join('');
        });
    }

    // 處理「新增」或「儲存修改」
    async function handleSubmit() {
        const publisher = document.getElementById('publisher').value.trim();
        const bookName = document.getElementById('bookName').value.trim();
        const songNameZh = document.getElementById('songNameZh').value.trim();
        const songNameEn = document.getElementById('songNameEn').value.trim();
        const composer = document.getElementById('composer').value.trim();
        const btn = document.getElementById('submitBtn');

        if (!publisher || !bookName) {
            alert("請至少輸入「出版社」與「書名」！");
            return;
        }

        const isEdit = editingId !== null;
        const payload = { 
            action: isEdit ? 'edit' : 'add', 
            id: editingId,
            publisher, bookName, songNameZh, songNameEn, composer 
        };

        btn.textContent = '⏳ 處理中...';
        btn.disabled = true;

        try {
            const response = await fetch(API_URL, {
                method: 'POST',
                body: JSON.stringify(payload),
                headers: { 'Content-Type': 'text/plain;charset=utf-8' }
            });
            const result = await response.json();

            if (isEdit) {
                // 更新本機資料
                const idx = records.findIndex(r => r.id === editingId);
                if (idx !== -1) {
                    records[idx] = { ...records[idx], publisher, bookName, songNameZh, songNameEn, composer };
                }
            } else {
                // 新增至本機資料
                payload.id = result.id;
                records.push(payload);
            }

            clearForm();
            updateDatalists();
            renderRecords();
            // 已完全移除原本的 alert 提醒！
        } catch (error) {
            console.error('儲存失敗:', error);
            alert("❌ 儲存失敗，請重試。");
        } finally {
            resetSubmitBtn();
        }
    }

    // 開始修改：將資料帶入輸入框
    function startEdit(id) {
        const item = records.find(r => r.id === id);
        if (!item) return;

        editingId = id;
        document.getElementById('publisher').value = item.publisher || '';
        document.getElementById('bookName').value = item.bookName || '';
        document.getElementById('songNameZh').value = item.songNameZh || '';
        document.getElementById('songNameEn').value = item.songNameEn || '';
        document.getElementById('composer').value = item.composer || '';

        const submitBtn = document.getElementById('submitBtn');
        submitBtn.textContent = '💾 儲存修改';
        submitBtn.className = 'btn-save-edit';
        
        document.getElementById('cancelBtn').style.display = 'inline-block';

        // 平滑滾動到頂部表單
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    // 取消修改
    function cancelEdit() {
        clearForm();
        resetSubmitBtn();
    }

    function clearForm() {
        editingId = null;
        document.getElementById('publisher').value = '';
        document.getElementById('bookName').value = '';
        document.getElementById('songNameZh').value = '';
        document.getElementById('songNameEn').value = '';
        document.getElementById('composer').value = '';
        document.getElementById('cancelBtn').style.display = 'none';
    }

    function resetSubmitBtn() {
        const submitBtn = document.getElementById('submitBtn');
        submitBtn.textContent = '➕ 新增記錄';
        submitBtn.className = 'btn-add';
        submitBtn.disabled = false;
    }

    async function deleteRecord(id, btnElement) {
        if (!confirm("⚠️ 確定要刪除這筆記錄嗎？(此動作無法復原)")) {
            return; 
        }
        btnElement.textContent = '⏳...';
        btnElement.disabled = true;

        try {
            await fetch(API_URL, {
                method: 'POST',
                body: JSON.stringify({ action: 'delete', id: id }),
                headers: { 'Content-Type': 'text/plain;charset=utf-8' }
            });
            records = records.filter(record => record.id !== id);
            
            if (editingId === id) cancelEdit();

            updateDatalists();
            renderRecords();
        } catch (error) {
            console.error('刪除失敗:', error);
            alert("❌ 刪除失敗，請檢查網路。");
            btnElement.textContent = '🗑️ 刪除';
            btnElement.disabled = false;
        }
    }

    function customSort(a, b) {
        const strA = a || '';
        const strB = b || '';
        return strA.localeCompare(strB, 'zh-HK', { numeric: true, collation: 'stroke' });
    }

    function renderRecords() {
        const searchTerm = document.getElementById('searchInput').value.toLowerCase();
        const container = document.getElementById('resultsContainer');
        container.innerHTML = '';

        const filteredRecords = records.filter(record => {
            return (record.publisher && record.publisher.toLowerCase().includes(searchTerm)) ||
                   (record.bookName && record.bookName.toLowerCase().includes(searchTerm)) ||
                   (record.songNameZh && record.songNameZh.toLowerCase().includes(searchTerm)) ||
                   (record.songNameEn && record.songNameEn.toLowerCase().includes(searchTerm)) ||
                   (record.composer && record.composer.toLowerCase().includes(searchTerm));
        });

        const groupedData = {};
        filteredRecords.forEach(record => {
            if (!groupedData[record.publisher]) {
                groupedData[record.publisher] = [];
            }
            groupedData[record.publisher].push(record);
        });

        const sortedPublishers = Object.keys(groupedData).sort(customSort);

        if (sortedPublishers.length === 0) {
            container.innerHTML = '<p style="text-align:center; color:#7f8c8d;">找不到符合的記錄。</p>';
            return;
        }

        sortedPublishers.forEach(publisher => {
            const groupDiv = document.createElement('div');
            groupDiv.className = 'publisher-group';
            
            const title = document.createElement('h2');
            title.className = 'publisher-title';
            title.textContent = `🏛️ ${publisher}`;
            groupDiv.appendChild(title);

            groupedData[publisher].sort((a, b) => {
                let cmp = customSort(a.bookName, b.bookName);
                if (cmp !== 0) return cmp;
                cmp = customSort(a.songNameZh, b.songNameZh);
                if (cmp !== 0) return cmp;
                return customSort(a.songNameEn, b.songNameEn);
            });

            let currentBook = '';
            let seq = 0;
            const itemsWithSeq = groupedData[publisher].map(item => {
                if (item.bookName !== currentBook) {
                    currentBook = item.bookName;
                    seq = 1;
                } else {
                    seq++;
                }
                return { ...item, seq };
            });

            const table = document.createElement('table');
            table.innerHTML = `
                <thead>
                    <tr>
                        <th style="width: 8%;">次序</th>
                        <th style="width: 24%;">書名</th>
                        <th style="width: 21%;">中文歌名</th>
                        <th style="width: 21%;">英文歌名</th>
                        <th style="width: 15%;">作曲家</th>
                        <th style="width: 11%;"></th>
                    </tr>
                </thead>
                <tbody>
                    ${itemsWithSeq.map(item => `
                        <tr>
                            <td style="color: #7f8c8d; font-weight: bold;">${item.seq}</td>
                            <td>${item.bookName}</td>
                            <td>${item.songNameZh || '-'}</td>
                            <td>${item.songNameEn || '-'}</td>
                            <td>${item.composer || '-'}</td>
                            <td>
                                <div class="btn-action-group">
                                    <button class="btn-edit" onclick="startEdit('${item.id}')">✏️ 修改</button>
                                    <button class="btn-delete" onclick="deleteRecord('${item.id}', this)">🗑️ 刪除</button>
                                </div>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            `;
            groupDiv.appendChild(table);
            container.appendChild(groupDiv);
        });
    }

    function exportToExcel() {
        if (records.length === 0) {
            alert("沒有資料可以匯出！");
            return;
        }

        const sortedRecords = [...records].sort((a, b) => {
            let cmp = customSort(a.publisher, b.publisher);
            if (cmp !== 0) return cmp;
            cmp = customSort(a.bookName, b.bookName);
            if (cmp !== 0) return cmp;
            cmp = customSort(a.songNameZh, b.songNameZh);
            if (cmp !== 0) return cmp;
            return customSort(a.songNameEn, b.songNameEn);
        });

        let currentPublisher = '';
        let currentBook = '';
        let seq = 0;

        const dataToExport = sortedRecords.map(r => {
            if (r.publisher !== currentPublisher || r.bookName !== currentBook) {
                currentPublisher = r.publisher;
                currentBook = r.bookName;
                seq = 1;
            } else {
                seq++;
            }
            return {
                "次序": seq,
                "出版社 (Publisher)": r.publisher,
                "書名 (Book Name)": r.bookName,
                "中文歌名 (Chinese Song)": r.songNameZh,
                "英文歌名 (English Song)": r.songNameEn,
                "作曲家 (Composer)": r.composer
            };
        });

        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "出版記錄");
        XLSX.writeFile(workbook, "出版物記錄.xlsx");
    }

    window.onload = loadRecords;
</script>
</body>
</html>
