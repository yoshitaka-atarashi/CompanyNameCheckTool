// グローバル変数
let selectedFiles = [];
let currentAction = 'replace';
let DEFAULT_KEYWORDS = [];

// DOM要素
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const filesList = document.getElementById('filesList');
const filesUl = document.getElementById('filesUl');
const changeFilesBtn = document.getElementById('changeFilesBtn');
const presetButtons = document.getElementById('presetButtons');
const selectedKeywordsList = document.getElementById('selectedKeywordsList');
const newKeywordInput = document.getElementById('newKeyword');
const actionRadios = document.querySelectorAll('input[name="action"]');
const detectBtn = document.getElementById('detectBtn');
const previewBtn = document.getElementById('previewBtn');
const executeBtn = document.getElementById('executeBtn');
const resultsSection = document.getElementById('resultsSection');
const resultContent = document.getElementById('resultContent');
const loading = document.getElementById('loading');
const errorAlert = document.getElementById('errorAlert');

// イベントリスナー設定
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});
uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    handleFilesSelect(e.dataTransfer.files);
});

fileInput.addEventListener('change', (e) => {
    handleFilesSelect(e.target.files);
    // 次のファイル選択時に change イベントが発火するようにリセット
    e.target.value = '';
});

changeFilesBtn.addEventListener('click', () => {
    fileInput.click();
});

// 初期化：デフォルトキーワードから UI を生成
function initializePresetKeywords() {
    DEFAULT_KEYWORDS = CONFIG.default_keywords;
    
    // プリセットボタンを動的に生成
    presetButtons.innerHTML = '';
    DEFAULT_KEYWORDS.forEach(keyword => {
        const label = document.createElement('label');
        label.className = 'preset-checkbox';
        label.innerHTML = `
            <input type="checkbox" class="keyword-checkbox" value="${keyword}">
            <span>${keyword}</span>
        `;
        presetButtons.appendChild(label);
    });
    
    // チェックボックスのイベントリスナーを再設定
    document.querySelectorAll('.keyword-checkbox').forEach(checkbox => {
        checkbox.addEventListener('change', updateSelectedKeywords);
    });
    
    // すべてをチェック
    initializeKeywords();
    
    // デフォルト置換テキストを設定
    newKeywordInput.value = CONFIG.default_replacement;
}

// キーワードチェックボックスの動作
// キーワードチェックボックスは initializePresetKeywords で設定されます

// デフォルトすべてをチェック
function initializeKeywords() {
    const checkboxes = document.querySelectorAll('.keyword-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = true;
    });
    updateSelectedKeywords();
}

// 選択済みキーワードを更新
function updateSelectedKeywords() {
    const checkboxes = document.querySelectorAll('.keyword-checkbox');
    const selected = Array.from(checkboxes)
        .filter(cb => cb.checked)
        .map(cb => cb.value);
    
    selectedKeywordsList.textContent = selected.length > 0 
        ? selected.join(', ') 
        : '(なし)';
}

actionRadios.forEach(radio => {
    radio.addEventListener('change', (e) => {
        currentAction = e.target.value;
        updateButtonVisibility();
    });
});

detectBtn.addEventListener('click', detectKeywords);
previewBtn.addEventListener('click', previewChanges);
executeBtn.addEventListener('click', executeChanges);

// ファイル選択処理（既存ファイルに追加）
function handleFilesSelect(files) {
    if (!files || files.length === 0) return;

    const validFiles = [];
    for (let file of files) {
        if (file.name.endsWith('.pptx') || file.name.endsWith('.ppt')) {
            // 重複チェック
            const isDuplicate = selectedFiles.some(f => f.name === file.name && f.size === file.size);
            if (!isDuplicate) {
                validFiles.push(file);
            }
        } else {
            showError(`"${file.name}" はPPTX/PPT形式ではありません`);
        }
    }

    if (validFiles.length > 0) {
        // 既存ファイルに追加（置き換えではなく）
        selectedFiles = selectedFiles.concat(validFiles);
        console.log('Selected files:', selectedFiles);
        displayFilesList();
        resultsSection.style.display = 'none';
        clearError();
    }
}

// ファイルリストを表示
function displayFilesList() {
    filesUl.innerHTML = '';
    selectedFiles.forEach((file, index) => {
        const li = document.createElement('li');
        li.innerHTML = `
            <span class="file-name">${index + 1}. ${file.name}</span>
            <button type="button" class="remove-file" data-index="${index}" title="削除">×</button>
        `;
        filesUl.appendChild(li);

        // 削除ボタンのイベント
        li.querySelector('.remove-file').addEventListener('click', (e) => {
            e.preventDefault();
            const idx = parseInt(e.target.getAttribute('data-index'));
            selectedFiles.splice(idx, 1);
            if (selectedFiles.length === 0) {
                filesList.style.display = 'none';
                uploadArea.style.display = 'block';
            } else {
                displayFilesList();
            }
        });
    });

    uploadArea.style.display = 'none';
    filesList.style.display = 'block';
}

// キーワード検出
async function detectKeywords() {
    if (!validateInputs('detect')) return;

    if (selectedFiles.length === 0) {
        showError('ファイルをアップロードしてください');
        return;
    }

    const selectedKeywords = Array.from(document.querySelectorAll('.keyword-checkbox'))
        .filter(cb => cb.checked)
        .map(cb => cb.value);

    if (selectedKeywords.length === 0) {
        showError('少なくとも1つのキーワードを選択してください');
        return;
    }

    let allResults = [];
    let totalKeywordCount = 0;

    try {
        showLoading(true);

        for (let file of selectedFiles) {
            const formData = new FormData();
            formData.append('file', file);
            formData.append('keywords', JSON.stringify(selectedKeywords));

            const response = await fetch('/api/detect', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (!response.ok) {
                showError(data.error || 'エラーが発生しました');
                return;
            }

            allResults.push({
                filename: file.name,
                data: data
            });

            totalKeywordCount += data.total_count;
        }

        displayDetectResults(allResults, totalKeywordCount);

    } catch (error) {
        showError('通信エラー: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 変更プレビュー
async function previewChanges() {
    if (!validateInputs(currentAction)) return;

    if (selectedFiles.length === 0) {
        showError('ファイルをアップロードしてください');
        return;
    }

    const selectedKeywords = Array.from(document.querySelectorAll('.keyword-checkbox'))
        .filter(cb => cb.checked)
        .map(cb => cb.value);

    if (selectedKeywords.length === 0) {
        showError('少なくとも1つのキーワードを選択してください');
        return;
    }

    let totalBefore = 0;
    let totalAfter = 0;
    let totalModified = 0;

    try {
        showLoading(true);

        for (let file of selectedFiles) {
            const formData = new FormData();
            formData.append('file', file);
            formData.append('keywords', JSON.stringify(selectedKeywords));
            formData.append('action', currentAction);
            if (currentAction === 'replace') {
                formData.append('new_keyword', newKeywordInput.value);
            }

            const response = await fetch('/api/preview', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (!response.ok) {
                showError(data.error || 'エラーが発生しました');
                return;
            }

            totalBefore += data.before.count;
            totalAfter += data.after.count;
            totalModified += data.modified_shapes;
        }

        displayPreviewResults({
            before: { count: totalBefore },
            after: { count: totalAfter },
            modified_shapes: totalModified,
            action: currentAction,
            file_count: selectedFiles.length
        });

    } catch (error) {
        showError('通信エラー: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 変更実行・ファイルダウンロード
async function executeChanges() {
    if (!validateInputs(currentAction)) return;

    if (selectedFiles.length === 0) {
        showError('ファイルをアップロードしてください');
        return;
    }

    const selectedKeywords = Array.from(document.querySelectorAll('.keyword-checkbox'))
        .filter(cb => cb.checked)
        .map(cb => cb.value);

    if (selectedKeywords.length === 0) {
        showError('少なくとも1つのキーワードを選択してください');
        return;
    }

    try {
        showLoading(true);

        for (let file of selectedFiles) {
            const formData = new FormData();
            formData.append('file', file);
            formData.append('keywords', JSON.stringify(selectedKeywords));
            formData.append('action', currentAction);
            if (currentAction === 'replace') {
                formData.append('new_keyword', newKeywordInput.value);
            }

            const response = await fetch('/api/replace', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const data = await response.json();
                showError(data.error || 'エラーが発生しました');
                return;
            }

            // ファイルをダウンロード
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `modified_${file.name}`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        }

        displaySuccessMessage(selectedFiles.length);

    } catch (error) {
        showError('通信エラー: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 検出結果を表示
function displayDetectResults(allResults, totalCount) {
    resultContent.innerHTML = '';

    const statsHtml = `
        <div class="stats">
            <div class="stat-box">
                <div class="stat-value">${totalCount}</div>
                <div class="stat-label">検出されたキーワード数</div>
            </div>
            <div class="stat-box">
                <div class="stat-value">${selectedFiles.length}</div>
                <div class="stat-label">処理したファイル数</div>
            </div>
        </div>
    `;

    let detailsHtml = '';
    allResults.forEach(result => {
        if (result.data.results.length > 0) {
            detailsHtml += `<h3 style="margin-top: 20px;">${result.filename} の詳細</h3>`;
            result.data.results.forEach(item => {
                detailsHtml += `
                    <div class="result-item">
                        <div class="result-header">スライド ${item.slide}</div>
                        <div class="result-details">
                            <p><strong>検出数:</strong> ${item.count}</p>
                            <p><strong>テキスト:</strong> ${escapeHtml(item.text)}</p>
                        </div>
                    </div>
                `;
            });
        }
    });

    resultContent.innerHTML = statsHtml + detailsHtml;
    resultsSection.style.display = 'block';
}

// プレビュー結果を表示
function displayPreviewResults(data) {
    resultContent.innerHTML = '';

    const actionLabel = {
        'delete': '削除',
        'replace': '置換'
    }[data.action] || '操作';

    const html = `
        <div class="alert alert-info">
            ${actionLabel}操作のプレビュー結果です (${data.file_count}ファイル)
        </div>
        <div class="stats">
            <div class="stat-box">
                <div class="stat-value">${data.before.count}</div>
                <div class="stat-label">処理前のキーワード数</div>
            </div>
            <div class="stat-box">
                <div class="stat-value">${data.after.count}</div>
                <div class="stat-label">処理後のキーワード数</div>
            </div>
            <div class="stat-box">
                <div class="stat-value">${data.before.count - data.after.count}</div>
                <div class="stat-label">削除されるキーワード数</div>
            </div>
            <div class="stat-box">
                <div class="stat-value">${data.modified_shapes}</div>
                <div class="stat-label">変更されるテキストボックス数</div>
            </div>
        </div>
    `;

    resultContent.innerHTML = html;
    resultsSection.style.display = 'block';
}

// 成功メッセージを表示
function displaySuccessMessage(fileCount) {
    resultContent.innerHTML = `
        <div class="alert alert-success">
            <strong>✓ 処理が完了しました</strong><br>
            ${fileCount}個のファイルがダウンロードされました
        </div>
    `;
    resultsSection.style.display = 'block';
}

// バリデーション
function validateInputs(action) {
    clearError();

    if (selectedFiles.length === 0) {
        showError('ファイルをアップロードしてください');
        return false;
    }

    const selectedKeywords = Array.from(document.querySelectorAll('.keyword-checkbox'))
        .filter(cb => cb.checked)
        .map(cb => cb.value);

    if (selectedKeywords.length === 0) {
        showError('少なくとも1つのキーワードを選択してください');
        return false;
    }

    if (action === 'replace' && !newKeywordInput.value.trim()) {
        showError('置換先のキーワードを入力してください');
        return false;
    }

    return true;
}

// ボタンの表示/非表示を更新
function updateButtonVisibility() {
    if (currentAction === 'detect') {
        detectBtn.style.display = 'inline-block';
        previewBtn.style.display = 'none';
        executeBtn.style.display = 'none';
    } else {
        detectBtn.style.display = 'none';
        previewBtn.style.display = 'inline-block';
        executeBtn.style.display = 'inline-block';
    }
}

// エラー表示
function showError(message) {
    errorAlert.textContent = message;
    errorAlert.style.display = 'block';
}

function clearError() {
    errorAlert.style.display = 'none';
    errorAlert.textContent = '';
}

// ローディング表示
function showLoading(show) {
    loading.style.display = show ? 'flex' : 'none';
}

// HTML エスケープ
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// 初期化
document.addEventListener('DOMContentLoaded', () => {
    initializePresetKeywords();
    updateButtonVisibility();
});
