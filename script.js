/* script.js (v6 - Weighting Feature Added) */
console.log("script.js: スクリプト開始");

// --- 設定値オブジェクト ---
const config = {
    ESTAT_APP_ID: "767c024416535e276a90d2d9469eeb5acf7e552f",
    STATS_DATA_ID: "0003443840",
    CACHE_KEY: `eStatDataCache_0003443840`,
    CACHE_DURATION_MS: 24 * 60 * 60 * 1000, // 1日
    AGE_CATEGORY_KEY: '@cat03',
    GENDER_CATEGORY_KEY: '@cat02',
    POPULATION_TYPE_KEY: '@cat04',
    POPULATION_TYPE_CODE: '001',
    VALUE_KEY: '$',
    TIME_KEY: '@time',
    UNIT_MULTIPLIER: 1000,
    ageCodeMap: { '01004': '15-19', '01005': '20-24', '01006': '25-29', '01007': '30-34', '01008': '35-39', '01009': '40-44', '01010': '45-49', '01011': '50-54', '01012': '55-59', '01013': '60-64', '01014': '65-69' },
    genderCodeMap: { '001': 'male', '002': 'female' },
    // ★ パネルデータをハードコード ★
    panels5: { '15-19': { male: 451, female: 1362 }, '20-24': { male: 833, female: 1939 }, '25-29': { male: 1584, female: 3261 }, '30-34': { male: 2604, female: 4152 }, '35-39': { male: 4197, female: 5018 }, '40-44': { male: 6549, female: 6245 }, '45-49': { male: 9747, female: 7021 }, '50-54': { male: 12569, female: 7875 }, '55-59': { male: 11741, female: 5790 }, '60-64': { male: 9778, female: 3903 }, '65-69': { male: 7119, female: 2353 } },
    warningThreshold: 0.7
};
config.panels10 = { '15-19': config.panels5['15-19'] || { male: 0, female: 0 }, '20-29': { male: (config.panels5['20-24']?.male || 0) + (config.panels5['25-29']?.male || 0), female: (config.panels5['20-24']?.female || 0) + (config.panels5['25-29']?.female || 0) }, '30-39': { male: (config.panels5['30-34']?.male || 0) + (config.panels5['35-39']?.male || 0), female: (config.panels5['30-34']?.female || 0) + (config.panels5['35-39']?.female || 0) }, '40-49': { male: (config.panels5['40-44']?.male || 0) + (config.panels5['45-49']?.male || 0), female: (config.panels5['40-44']?.female || 0) + (config.panels5['45-49']?.female || 0) }, '50-59': { male: (config.panels5['50-54']?.male || 0) + (config.panels5['55-59']?.male || 0), female: (config.panels5['50-54']?.female || 0) + (config.panels5['55-59']?.female || 0) }, '60-69': { male: (config.panels5['60-64']?.male || 0) + (config.panels5['65-69']?.male || 0), female: (config.panels5['60-64']?.female || 0) + (config.panels5['65-69']?.female || 0) } };
console.log("Using hardcoded panels5:", config.panels5); console.log("Calculated panels10:", config.panels10);
const ageGroups5 = Object.values(config.ageCodeMap); // config参照に変更

// --- グローバル変数 ---
let rawApiDataCache = null;
let availableTimeCodes = [];
let selectedTimeCode = "";
let calculatedRatios5 = null;
let calculatedRatios10 = null;
let isDataReady = false;
let previouslyFocusedElement;
let currentSimulationResults = null; // シミュレーション結果を保持

// DOM要素
let checkerForm, resultModal, modalContent, errorMessage, step2Element, submitButton, loadingSpinner, buttonText;
let targetSampleInput, genderFilterSelect, ageMinInput, ageMaxInput, ageModeSelect, responseRateInput;
let targetSampleError, ageMinError, ageMaxError, ageRangeError, responseRateError;
let timeSelectorElement, timeSelectorContainer;
let calculateWeightsButtonElement; // ウェイト計算ボタン用

// --- キャッシュ処理 ---
function getCachedData() {
    const cached = localStorage.getItem(config.CACHE_KEY); if (!cached) { console.log("getCachedData: キャッシュなし"); return null; } try { const parsed = JSON.parse(cached); const now = Date.now(); if (now - parsed.timestamp < config.CACHE_DURATION_MS) { console.log("getCachedData: 有効なキャッシュあり"); return parsed.data; } else { console.log("getCachedData: キャッシュ期限切れ"); localStorage.removeItem(config.CACHE_KEY); return null; } } catch (e) { console.error("getCachedData: キャッシュの解析に失敗", e); localStorage.removeItem(config.CACHE_KEY); return null; }
}
function saveDataToCache(dataToCache) {
    try { const cacheData = { timestamp: Date.now(), data: dataToCache }; localStorage.setItem(config.CACHE_KEY, JSON.stringify(cacheData)); console.log("saveDataToCache: データをキャッシュに保存しました"); } catch (e) { console.error("saveDataToCache: キャッシュの保存に失敗", e); }
}

// --- APIデータ取得・処理 ---
async function fetchDataFromApi() {
    console.log("fetchDataFromApi: 開始");
    if (config.ESTAT_APP_ID === "YOUR_APP_ID_HERE" || !config.ESTAT_APP_ID) { throw new Error("APIキー (appId) が config オブジェクトに設定されていません。"); }
    const apiUrl = `https://api.e-stat.go.jp/rest/3.0/app/json/getStatsData?appId=${config.ESTAT_APP_ID}&lang=J&statsDataId=${config.STATS_DATA_ID}&metaGetFlg=Y&cntGetFlg=N&explanationGetFlg=Y&annotationGetFlg=Y&sectionHeaderFlg=1&replaceSpChars=0`;
    console.log("fetchDataFromApi: Fetching URL:", apiUrl);
    const response = await fetch(apiUrl);
    if (!response.ok) { throw new Error(`API通信エラー: ${response.status} ${response.statusText}`); }
    const data = await response.json();
    console.log("fetchDataFromApi: APIデータ取得成功");
    return data;
}

function processRawApiData(apiResponseData) {
    console.log("processRawApiData: 開始"); const valueDataAll = apiResponseData?.GET_STATS_DATA?.STATISTICAL_DATA?.DATA_INF?.VALUE; if (!valueDataAll || !Array.isArray(valueDataAll)) { throw new Error("APIレスポンス形式不正: 人口データ(VALUE)が見つかりません。"); } console.log("processRawApiData: valueDataAll 件数:", valueDataAll.length); let timeCodes = []; try { console.log("processRawApiData: 時点コード抽出開始"); const timeCodeSet = new Set(); valueDataAll.forEach(item => { if (item && typeof item === 'object' && config.TIME_KEY in item) { if (item[config.TIME_KEY]) { timeCodeSet.add(item[config.TIME_KEY]); } else { console.warn("processRawApiData: 空の @time コード:", item); } } else { console.warn("processRawApiData: 不正なデータ項目:", item); } }); timeCodes = [...timeCodeSet].sort().reverse(); console.log("processRawApiData: 利用可能な時点コード:", timeCodes); if (timeCodes.length === 0) { console.warn("processRawApiData: 利用可能な時点コードが見つかりませんでした。"); } } catch (e) { console.error("processRawApiData: 時点コードの抽出エラー", e); throw new Error("APIデータから時点コードの抽出に失敗しました。"); } return { valueDataAll, timeCodes };
}

function populateTimeSelector(timeCodes) {
    console.log("populateTimeSelector: 開始", timeCodes); if (!timeSelectorElement) { console.error("populateTimeSelector: timeSelectorElement が見つかりません。"); return; } timeSelectorElement.innerHTML = ''; if (!timeCodes || timeCodes.length === 0) { console.warn("populateTimeSelector: timeCodes が空です。"); timeSelectorElement.disabled = true; const option = document.createElement('option'); option.value = ""; option.textContent = "データなし"; timeSelectorElement.appendChild(option); if (timeSelectorContainer) timeSelectorContainer.style.display = 'block'; return; } console.log("populateTimeSelector: オプション生成開始"); timeCodes.forEach((code, index) => { const option = document.createElement('option'); option.value = code; try { const year = code.substring(0, 4); const month = code.substring(6, 8); option.textContent = `${year}年${parseInt(month, 10)}月`; console.log(`populateTimeSelector: オプション追加: ${option.textContent} (value=${code})`); } catch(e) { console.error("populateTimeSelector: 日付解析エラー", code, e); option.textContent = code; } timeSelectorElement.appendChild(option); if (index === 0) { option.selected = true; selectedTimeCode = code; console.log("populateTimeSelector: デフォルト選択:", selectedTimeCode); } }); timeSelectorElement.disabled = false; console.log("populateTimeSelector: timeSelectorElement.disabled を false に設定"); if (timeSelectorContainer) timeSelectorContainer.style.display = 'block'; console.log("populateTimeSelector: 完了");
}

function calculateRatiosForTime(timeCode) {
    console.log(`calculateRatiosForTime: 開始 (${timeCode})`);
    isDataReady = false;
    calculatedRatios5 = null;
    calculatedRatios10 = null;
    if (!rawApiDataCache) { console.error("calculateRatiosForTime: APIデータキャッシュがありません"); displayInitializationError("内部エラー: キャッシュデータが見つかりません。"); disableForm("データエラー"); return; }

    try {
        const valueData = rawApiDataCache.filter(item => item[config.TIME_KEY] === timeCode);
        if (valueData.length === 0) { throw new Error(`対象時点 (${timeCode}) のデータが見つかりません。`); }

        const population = {};
        let totalPopulation = 0;
        valueData.forEach(item => {
            if (item[config.POPULATION_TYPE_KEY] !== config.POPULATION_TYPE_CODE) return;
            const ageRangeCode = item[config.AGE_CATEGORY_KEY];
            const genderCode = item[config.GENDER_CATEGORY_KEY];
            const countStr = item[config.VALUE_KEY];
            const ageRange = config.ageCodeMap[ageRangeCode];
            const gender = config.genderCodeMap[genderCode];
            const count = (countStr && countStr !== "-") ? parseInt(countStr, 10) * config.UNIT_MULTIPLIER : 0;
            if (ageRange && gender && !isNaN(count) && count >= 0) {
                if (!population[ageRange]) population[ageRange] = { male: 0, female: 0 };
                population[ageRange][gender] += count;
                totalPopulation += count;
            }
        });
        console.log(`calculateRatiosForTime (${timeCode}): 集計後人口データ:`, population);
        console.log(`calculateRatiosForTime (${timeCode}): 集計後総人口:`, totalPopulation);
        if (totalPopulation <= 0) { throw new Error(`データ処理エラー: 集計後の総人口(${timeCode})が0以下です。フィルタリング条件やAPIデータを確認してください。`); }

        calculatedRatios5 = {};
        for (const range in config.ageCodeMap) { const ageRangeKey = config.ageCodeMap[range]; if (!calculatedRatios5[ageRangeKey]) calculatedRatios5[ageRangeKey] = { male: 0, female: 0 }; if (population[ageRangeKey]) calculatedRatios5[ageRangeKey] = { male: (population[ageRangeKey].male || 0) / totalPopulation, female: (population[ageRangeKey].female || 0) / totalPopulation }; }
        console.log(`calculateRatiosForTime (${timeCode}): 計算後構成比 (5歳刻み):`, calculatedRatios5);

        calculatedRatios10 = { '15-19': calculatedRatios5['15-19'] || { male: 0, female: 0 }, '20-29': { male: 0, female: 0 }, '30-39': { male: 0, female: 0 }, '40-49': { male: 0, female: 0 }, '50-59': { male: 0, female: 0 }, '60-69': { male: 0, female: 0 } };
        for (const range5 in calculatedRatios5) { const [minAge] = range5.split('-').map(Number); const ratio = calculatedRatios5[range5]; if (minAge >= 20 && minAge <= 29) { calculatedRatios10['20-29'].male += ratio.male; calculatedRatios10['20-29'].female += ratio.female; } else if (minAge >= 30 && minAge <= 39) { calculatedRatios10['30-39'].male += ratio.male; calculatedRatios10['30-39'].female += ratio.female; } else if (minAge >= 40 && minAge <= 49) { calculatedRatios10['40-49'].male += ratio.male; calculatedRatios10['40-49'].female += ratio.female; } else if (minAge >= 50 && minAge <= 59) { calculatedRatios10['50-59'].male += ratio.male; calculatedRatios10['50-59'].female += ratio.female; } else if (minAge >= 60 && minAge <= 69) { calculatedRatios10['60-69'].male += ratio.male; calculatedRatios10['60-69'].female += ratio.female; } }
        console.log(`calculateRatiosForTime (${timeCode}): 計算後構成比 (10歳刻み):`, calculatedRatios10);

        isDataReady = true;
        enableForm();
        console.log(`calculateRatiosForTime (${timeCode}): データ処理完了`);
    } catch (error) {
        console.error(`calculateRatiosForTime (${timeCode}): エラー発生:`, error);
        displayInitializationError(`人口データの処理に失敗しました (${timeCode}): ${error.message}`);
        isDataReady = false;
        disableForm("データ処理エラー");
    }
}

// --- UI更新・制御 ---
function displayInitializationError(message) {
    console.error("Initialization Error:", message);
    if (errorMessage) {
        errorMessage.textContent = `初期化エラー: ${message}`;
        errorMessage.classList.remove('hidden');
    }
    disableForm("初期化エラー"); // ボタンテキスト変更
}
function disableForm(buttonTextOverride = "データロード中...") {
    if (submitButton) { submitButton.disabled = true; if (buttonText) buttonText.textContent = buttonTextOverride; submitButton.classList.add('opacity-50', 'cursor-not-allowed'); }
    [targetSampleInput, genderFilterSelect, ageMinInput, ageMaxInput, ageModeSelect, responseRateInput, timeSelectorElement].forEach(el => { if (el) el.disabled = true; });
}
function enableForm() {
    if (!isDataReady) { console.warn("enableForm: データ準備未完了"); disableForm("データエラー"); return; }
    console.log("enableForm: フォームを有効化");
    if (submitButton) { submitButton.disabled = false; if (buttonText) buttonText.textContent = '判定を実行'; submitButton.classList.remove('opacity-50', 'cursor-not-allowed'); }
    [targetSampleInput, genderFilterSelect, ageMinInput, ageMaxInput, ageModeSelect, responseRateInput, timeSelectorElement].forEach(el => { if (el) { el.disabled = false; if (el.id === 'timeSelector') { console.log("enableForm: timeSelectorElement の disabled を false に設定"); } } });
}

// --- ヘルパー関数, バリデーション関数, モーダル関連関数 ---
function displayInputError(errorElement, message, inputElement) { if (!errorElement || !inputElement) return; errorElement.textContent = message; if (message) { inputElement.classList.add('input-error'); inputElement.setAttribute('aria-invalid', 'true'); errorElement.style.display = 'block'; } else { inputElement.classList.remove('input-error'); inputElement.removeAttribute('aria-invalid'); errorElement.style.display = 'none'; } }
function toggleLoading(isLoading) { if (!submitButton || !loadingSpinner || !buttonText) { console.warn("toggleLoading: Button elements not found"); return; } if (isLoading) { submitButton.disabled = true; loadingSpinner.classList.remove('hidden'); buttonText.textContent = '判定中...'; submitButton.classList.add('opacity-50', 'cursor-not-allowed'); } else { submitButton.disabled = false; loadingSpinner.classList.add('hidden'); buttonText.textContent = '判定を実行'; submitButton.classList.remove('opacity-50', 'cursor-not-allowed'); } }
function getColorClass(status) { switch (status) { case "十分": return "bg-green-100 text-green-800"; case "注意": return "bg-yellow-100 text-yellow-800"; case "困難": return "bg-red-100 text-red-800"; default: return "bg-gray-100 text-gray-800"; } }
function validateTargetSample() { if (!targetSampleInput) return false; const value = parseInt(targetSampleInput.value, 10); let message = ''; if (isNaN(value) || value <= 0) { message = "回収希望数は1以上の数値を入力してください。"; } displayInputError(targetSampleError, message, targetSampleInput); return !message; }
function validateAgeMin() { if (!ageMinInput) return false; const value = parseInt(ageMinInput.value, 10); let message = ''; if (isNaN(value) || value < 15 || value > 69) { message = "最小年齢は15〜69の範囲で入力してください。"; } displayInputError(ageMinError, message, ageMinInput); validateAgeRange(); return !message; }
function validateAgeMax() { if (!ageMaxInput) return false; const value = parseInt(ageMaxInput.value, 10); let message = ''; if (isNaN(value) || value < 15 || value > 69) { message = "最大年齢は15〜69の範囲で入力してください。"; } displayInputError(ageMaxError, message, ageMaxInput); validateAgeRange(); return !message; }
function validateAgeRange() { if (!ageMinInput || !ageMaxInput) return false; const ageMin = parseInt(ageMinInput.value, 10); const ageMax = parseInt(ageMaxInput.value, 10); let message = ''; if (!isNaN(ageMin) && !isNaN(ageMax) && ageMin > ageMax) { message = "最小年齢は最大年齢以下に設定してください。"; } if (ageRangeError) { ageRangeError.textContent = message; if(message) ageRangeError.style.display = 'block'; else ageRangeError.style.display = 'none'; } return !message; }
function validateResponseRate() { if (!responseRateInput) return false; const value = parseFloat(responseRateInput.value); let message = ''; if (isNaN(value) || value < 1 || value > 100) { message = "想定回答率は1〜100%の範囲で入力してください。"; } displayInputError(responseRateError, message, responseRateInput); return !message; }
function validateForm() { console.log("validateForm: 開始"); if (errorMessage) { errorMessage.classList.add("hidden"); errorMessage.textContent = ''; } else { console.error("validateForm: errorMessage element not found!"); } const isSampleValid = validateTargetSample(); const isAgeMinValid = validateAgeMin(); const isAgeMaxValid = validateAgeMax(); const isAgeRangeValid = validateAgeRange(); const isRateValid = validateResponseRate(); const isValid = isSampleValid && isAgeMinValid && isAgeMaxValid && isAgeRangeValid && isRateValid; console.log("validateForm: 結果 =", isValid); return isValid; }
function handleModalFocus(e) { if (!resultModal || resultModal.classList.contains('hidden')) return; const focusableElements = Array.from( resultModal.querySelectorAll( 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])' ) ).filter(el => el.offsetParent !== null); if (focusableElements.length === 0) return; const firstElement = focusableElements[0]; const lastElement = focusableElements[focusableElements.length - 1]; if (e.key === 'Tab') { if (e.shiftKey) { if (document.activeElement === firstElement) { lastElement.focus(); e.preventDefault(); } } else { if (document.activeElement === lastElement) { firstElement.focus(); e.preventDefault(); } } } }
function openModal() { console.log("openModal: 開始"); if (!resultModal) { console.error("openModal: resultModal element not found!"); return; } previouslyFocusedElement = document.activeElement; resultModal.classList.remove("hidden"); document.body.classList.add("modal-open"); document.addEventListener('keydown', handleModalFocus); setTimeout(() => { const closeButton = document.getElementById('closeModalButton'); if(closeButton) { closeButton.focus(); } else { console.warn("openModal: Close button not found for focusing"); } }, 100); if (step2Element) { const step2Circle = step2Element.querySelector('.w-6'); const step2Text = step2Element.querySelector('span'); if (step2Circle && step2Text) { step2Circle.classList.remove("bg-gray-200", "text-gray-500"); step2Circle.classList.add("bg-green-600", "text-white"); step2Text.classList.remove("text-gray-400"); step2Text.classList.add("text-green-600"); } }
    // ★ ウェイト計算ボタンを表示 ★
    if (calculateWeightsButtonElement) {
        calculateWeightsButtonElement.classList.remove('hidden');
    }
}
window.closeModal = function() { console.log("closeModal: 開始"); if (!resultModal) return; resultModal.classList.add("hidden"); document.body.classList.remove("modal-open"); document.removeEventListener('keydown', handleModalFocus); if (previouslyFocusedElement) { previouslyFocusedElement.focus(); } }
window.downloadCSV = function() {
    console.log("downloadCSV: 開始");
    if (!modalContent) return;
    // ★ テーブル特定ロジック変更 ★
    const requiredTable = modalContent.querySelector('#requiredTable'); // IDで特定
    const deficitTable = modalContent.querySelector('#deficitTable');   // IDで特定
    if (!requiredTable) { console.warn("downloadCSV: Required table not found."); return; }

    let csv = "\uFEFF";
    csv += '"必要数・実績数・ウェイト"\n'; // テーブルタイトル変更
    // ヘッダー行: 必要数テーブルから取得し、「実績数」「ウェイト」を追加
    const headers = Array.from(requiredTable.querySelectorAll('thead th'))
                         .map(th => `"${th.textContent.trim().replace(/"/g, '""')}"`);
    // 最後の「小計」の前に「実績数」「ウェイト」を挿入
    headers.splice(headers.length -1, 0, '"実績数"', '"ウェイト"');
    csv += headers.join(",") + "\r\n";

    // データ行: 必要数テーブルの行をループ
    const rows = Array.from(requiredTable.querySelectorAll('tbody tr'));
    rows.forEach(tr => {
        if (tr.classList.contains('total-row')) return; // 合計行はスキップ

        const cells = Array.from(tr.querySelectorAll('td'));
        const rowData = [];
        cells.forEach((cell, index) => {
             rowData.push(`"${cell.textContent.trim().replace(/"/g, '""')}"`); // 年齢帯、男女別必要数、小計

             // 男女別必要数の後 (index 1 と 2 の後、または性別が片方なら index 1 の後) に実績数とウェイトを追加
             if (index > 0 && index < headers.length - 3) { // 年齢帯の後、小計の前のセル
                 const collectedInput = cell.querySelector('.collected-count-input');
                 const weightOutput = cell.querySelector('.weight-output');
                 rowData.push(`"${collectedInput ? collectedInput.value : ''}"`); // 実績数
                 rowData.push(`"${weightOutput ? weightOutput.textContent.trim() : ''}"`); // ウェイト
             }
        });
        csv += rowData.join(",") + "\r\n";
    });

     // 合計行の処理 (必要に応じて) - 現在は実績数やウェイトの合計は含めない

    // 不足数テーブル
    if (deficitTable) {
        csv += '\r\n"不足数テーブル"\n';
        const headersDeficit = Array.from(deficitTable.querySelectorAll('thead th'))
                               .map(th => `"${th.textContent.trim().replace(/"/g, '""')}"`)
                               .join(",");
        csv += headersDeficit + "\r\n";
        const rowsDeficit = Array.from(deficitTable.querySelectorAll('tbody tr'));
        rowsDeficit.forEach(tr => {
            const cells = Array.from(tr.querySelectorAll('td'));
            const row = cells.map(cell => `"${cell.textContent.trim().replace(/"/g, '""')}"`).join(",");
            csv += row + "\r\n";
        });
    }

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const now = new Date();
    const formattedDate = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
    a.download = `回収可否判定_${formattedDate}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    console.log("downloadCSV: 完了");
}

// --- シミュレーション関連関数 (リファクタリング) ---
function getUserInputs() {
    if (!validateForm()) { console.log("getUserInputs: バリデーション失敗"); if(errorMessage) { errorMessage.textContent = "入力内容を確認してください。"; errorMessage.classList.remove("hidden"); } return null; }
    console.log("getUserInputs: バリデーション成功");
    return {
        sample: parseInt(targetSampleInput.value, 10),
        gender: genderFilterSelect.value,
        ageMin: parseInt(ageMinInput.value, 10),
        ageMax: parseInt(ageMaxInput.value, 10),
        ageMode: ageModeSelect.value,
        responseRate: parseFloat(responseRateInput.value) / 100
    };
}
function performSimulation(inputs) {
    console.log("performSimulation: 計算開始", inputs);
    const { sample, gender, ageMin, ageMax, ageMode, responseRate } = inputs;
    const gendersToProcess = gender === "両方" ? ["male", "female"] : [gender === "男性" ? "male" : "female"];
    const correctionFactor = 1.0;
    const [ratiosData, panelsData] = ageMode === "5"
        ? [calculatedRatios5, config.panels5] // config参照
        : [calculatedRatios10, config.panels10]; // config参照

    if (!ratiosData) { throw new Error(`構成比データ (${ageMode === "5" ? '5歳刻み' : '10歳刻み'}) が読み込まれていません。`); }

    const targetAgeRanges = Object.keys(ratiosData).filter(range => {
        const rangeParts = range.split('-'); if (rangeParts.length !== 2) return false;
        const [minAgeInRange, maxAgeInRange] = rangeParts.map(Number); if (isNaN(minAgeInRange) || isNaN(maxAgeInRange)) return false;
        return ageMax >= minAgeInRange && ageMin <= maxAgeInRange;
    }).sort();
    console.log("performSimulation: 対象年齢範囲:", targetAgeRanges);

    let totalWeight = 0;
    const cellsData = [];
    targetAgeRanges.forEach(range => {
        gendersToProcess.forEach(g => {
            if (ratiosData[range]?.[g] !== undefined && panelsData[range]?.[g] !== undefined) {
                const weight = ratiosData[range][g] * correctionFactor;
                const validWeight = Math.max(0, weight);
                const panelAvailable = Math.floor(panelsData[range][g] * responseRate); // config参照
                cellsData.push({ range, gender: g, weight: validWeight, panel: panelAvailable });
                totalWeight += validWeight;
            } else { console.warn(`データ不足 (Ratio or Panel): range=${range}, gender=${g}`); }
        });
    });
    console.log("performSimulation: 総ウェイト:", totalWeight); console.log("performSimulation: 初期セルデータ:", cellsData);
    if (cellsData.length === 0) { throw new Error("対象となる年齢・性別のデータが見つかりませんでした。"); }
    if (totalWeight < 0) { console.warn("performSimulation: totalWeight is negative."); throw new Error("構成比の合計が負の値になりました。"); }

    let totalRequiredCalculated = 0;
    cellsData.forEach(cell => { const rawRequired = (totalWeight > 0) ? (sample * (cell.weight / totalWeight)) : (sample / cellsData.length); cell.floor = Math.floor(rawRequired); cell.remainder = rawRequired - cell.floor; totalRequiredCalculated += cell.floor; });
    let remainingSamples = sample - totalRequiredCalculated; remainingSamples = Math.max(0, Math.round(remainingSamples * 1e6) / 1e6);
    cellsData.sort((a, b) => b.remainder - a.remainder);
    cellsData.forEach(cell => { cell.required = cell.floor + (remainingSamples > 1e-9 ? 1 : 0); if (remainingSamples > 1e-9) { remainingSamples--; } });
    let finalTotalRequired = cellsData.reduce((sum, cell) => sum + cell.required, 0); let diff = sample - finalTotalRequired;
    if (diff !== 0 && cellsData.length > 0) { console.warn(`Rounding adjustment: diff = ${diff}`); cellsData.sort((a, b) => a.remainder - b.remainder); for (let i = 0; i < Math.abs(diff); i++) { if (diff > 0) { cellsData[i % cellsData.length].required++; } else if (cellsData[i % cellsData.length].required > 0) { cellsData[i % cellsData.length].required--; } } finalTotalRequired = cellsData.reduce((sum, cell) => sum + cell.required, 0); console.log(`Adjusted total: ${finalTotalRequired}`); }
    console.log("performSimulation: 必要数計算後セルデータ:", cellsData);

    const totalsRequired = {}; const totalsDeficit = {}; gendersToProcess.forEach(g => { totalsRequired[g] = 0; totalsDeficit[g] = 0; }); let grandTotalDeficit = 0; const comments = [];
    cellsData.forEach(cell => {
        const req = cell.required; const panel = cell.panel; const deficit = Math.max(req - panel, 0); let statu
