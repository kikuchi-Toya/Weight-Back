/* weightback.js - Weight Back Aggregation Tool */
console.log("weightback.js: スクリプト開始");

// --- 設定値オブジェクト ---
const config = {
    ESTAT_APP_ID: "767c024416535e276a90d2d9469eeb5acf7e552f", // ★ ご自身のIDに要変更 ★
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
    // warningThreshold はこのツールでは直接使用しない
};
const ageGroups5 = Object.values(config.ageCodeMap);

// --- グローバル変数 ---
let rawApiDataCache = null;
let availableTimeCodes = [];
let selectedTimeCode = "";
let calculatedRatios5 = null; // 5歳刻みの構成比のみ使用
let isDataReady = false;
let previouslyFocusedElement;

// DOM要素
let sampleInputForm, sampleInputTableBody, timeSelectorElement, timeSelectorContainer;
let calculateButton, loadingSpinner, buttonText, resultsSection, resultsArea, errorMessage;
let totalMaleElement, totalFemaleElement, grandTotalElement;
let downloadResultsCsvButton; // 結果CSVダウンロードボタン用

// --- キャッシュ処理 ---
function getCachedData() { /* (変更なし、前回のコードと同じ) */
    const cached = localStorage.getItem(config.CACHE_KEY); if (!cached) { console.log("getCachedData: キャッシュなし"); return null; } try { const parsed = JSON.parse(cached); const now = Date.now(); if (now - parsed.timestamp < config.CACHE_DURATION_MS) { console.log("getCachedData: 有効なキャッシュあり"); return parsed.data; } else { console.log("getCachedData: キャッシュ期限切れ"); localStorage.removeItem(config.CACHE_KEY); return null; } } catch (e) { console.error("getCachedData: キャッシュの解析に失敗", e); localStorage.removeItem(config.CACHE_KEY); return null; }
}
function saveDataToCache(dataToCache) { /* (変更なし、前回のコードと同じ) */
    try { const cacheData = { timestamp: Date.now(), data: dataToCache }; localStorage.setItem(config.CACHE_KEY, JSON.stringify(cacheData)); console.log("saveDataToCache: データをキャッシュに保存しました"); } catch (e) { console.error("saveDataToCache: キャッシュの保存に失敗", e); }
}

// --- APIデータ取得・処理 ---
async function fetchDataFromApi() { /* (変更なし、前回のコードと同じ) */
    console.log("fetchDataFromApi: 開始"); if (config.ESTAT_APP_ID === "YOUR_APP_ID_HERE" || !config.ESTAT_APP_ID) { throw new Error("APIキー (appId) が config オブジェクトに設定されていません。"); } const apiUrl = `https://api.e-stat.go.jp/rest/3.0/app/json/getStatsData?appId=${config.ESTAT_APP_ID}&lang=J&statsDataId=${config.STATS_DATA_ID}&metaGetFlg=Y&cntGetFlg=N&explanationGetFlg=Y&annotationGetFlg=Y&sectionHeaderFlg=1&replaceSpChars=0`; console.log("fetchDataFromApi: Fetching URL:", apiUrl); const response = await fetch(apiUrl); if (!response.ok) { throw new Error(`API通信エラー: ${response.status} ${response.statusText}`); } const data = await response.json(); console.log("fetchDataFromApi: APIデータ取得成功"); return data;
}
function processRawApiData(apiResponseData) { /* (変更なし、前回のコードと同じ) */
    console.log("processRawApiData: 開始"); const valueDataAll = apiResponseData?.GET_STATS_DATA?.STATISTICAL_DATA?.DATA_INF?.VALUE; if (!valueDataAll || !Array.isArray(valueDataAll)) { throw new Error("APIレスポンス形式不正: 人口データ(VALUE)が見つかりません。"); } console.log("processRawApiData: valueDataAll 件数:", valueDataAll.length); let timeCodes = []; try { console.log("processRawApiData: 時点コード抽出開始"); const timeCodeSet = new Set(); valueDataAll.forEach(item => { if (item && typeof item === 'object' && config.TIME_KEY in item) { if (item[config.TIME_KEY]) { timeCodeSet.add(item[config.TIME_KEY]); } else { console.warn("processRawApiData: 空の @time コード:", item); } } else { console.warn("processRawApiData: 不正なデータ項目:", item); } }); timeCodes = [...timeCodeSet].sort().reverse(); console.log("processRawApiData: 利用可能な時点コード:", timeCodes); if (timeCodes.length === 0) { console.warn("processRawApiData: 利用可能な時点コードが見つかりませんでした。"); } } catch (e) { console.error("processRawApiData: 時点コードの抽出エラー", e); throw new Error("APIデータから時点コードの抽出に失敗しました。"); } return { valueDataAll, timeCodes };
}
function populateTimeSelector(timeCodes) { /* (変更なし、前回のコードと同じ) */
    console.log("populateTimeSelector: 開始", timeCodes); if (!timeSelectorElement) { console.error("populateTimeSelector: timeSelectorElement が見つかりません。"); return; } timeSelectorElement.innerHTML = ''; if (!timeCodes || timeCodes.length === 0) { console.warn("populateTimeSelector: timeCodes が空です。"); timeSelectorElement.disabled = true; const option = document.createElement('option'); option.value = ""; option.textContent = "データなし"; timeSelectorElement.appendChild(option); if (timeSelectorContainer) timeSelectorContainer.style.display = 'block'; return; } console.log("populateTimeSelector: オプション生成開始"); timeCodes.forEach((code, index) => { const option = document.createElement('option'); option.value = code; try { const year = code.substring(0, 4); const month = code.substring(6, 8); option.textContent = `${year}年${parseInt(month, 10)}月`; console.log(`populateTimeSelector: オプション追加: ${option.textContent} (value=${code})`); } catch(e) { console.error("populateTimeSelector: 日付解析エラー", code, e); option.textContent = code; } timeSelectorElement.appendChild(option); if (index === 0) { option.selected = true; selectedTimeCode = code; console.log("populateTimeSelector: デフォルト選択:", selectedTimeCode); } }); timeSelectorElement.disabled = false; console.log("populateTimeSelector: timeSelectorElement.disabled を false に設定"); if (timeSelectorContainer) timeSelectorContainer.style.display = 'block'; console.log("populateTimeSelector: 完了");
}

function calculateRatiosForTime(timeCode) {
    console.log(`calculateRatiosForTime: 開始 (${timeCode})`);
    isDataReady = false; // Reset flag
    calculatedRatios5 = null; // Reset calculated data
    calculatedRatios10 = null; // Reset calculated data

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
        if (totalPopulation <= 0) { throw new Error(`データ処理エラー: 集計後の総人口(${timeCode})が0以下です。`); }

        // ★ 5歳刻みの構成比のみ計算・保持 ★
        calculatedRatios5 = {};
        for (const range in config.ageCodeMap) {
             const ageRangeKey = config.ageCodeMap[range];
             if (!calculatedRatios5[ageRangeKey]) calculatedRatios5[ageRangeKey] = { male: 0, female: 0 };
             if (population[ageRangeKey]) calculatedRatios5[ageRangeKey] = { male: (population[ageRangeKey].male || 0) / totalPopulation, female: (population[ageRangeKey].female || 0) / totalPopulation };
        }
        console.log(`calculateRatiosForTime (${timeCode}): 計算後構成比 (5歳刻み):`, calculatedRatios5);
        // calculatedRatios10 は不要

        isDataReady = true;
        enableForm(); // 構成比計算が成功したらフォーム有効化
        console.log(`calculateRatiosForTime (${timeCode}): データ処理完了`);

    } catch (error) {
        console.error(`calculateRatiosForTime (${timeCode}): エラー発生:`, error);
        displayInitializationError(`人口構成比の計算に失敗しました (${timeCode}): ${error.message}`);
        isDataReady = false;
        disableForm("データ処理エラー");
    }
}

// --- UI更新・制御 ---
function displayInitializationError(message) {
    console.error("Initialization Error:", message);
    if (errorMessage) { errorMessage.textContent = message; errorMessage.classList.remove('hidden'); }
    disableForm("初期化エラー");
}
function disableForm(buttonTextOverride = "読込中...") { // ボタンテキスト変更
    if (calculateButton) { calculateButton.disabled = true; if (buttonText) buttonText.textContent = buttonTextOverride; calculateButton.classList.add('opacity-50', 'cursor-not-allowed'); }
    if (timeSelectorElement) timeSelectorElement.disabled = true;
    // Disable all input fields in the table
    document.querySelectorAll('#sampleInputTableBody input').forEach(input => input.disabled = true);
}
function enableForm() {
    if (!isDataReady) { console.warn("enableForm: データ準備未完了"); disableForm("データエラー"); return; }
    console.log("enableForm: フォームを有効化");
    if (calculateButton) { calculateButton.disabled = false; if (buttonText) buttonText.textContent = 'ウェイトバック集計を実行'; calculateButton.classList.remove('opacity-50', 'cursor-not-allowed'); }
    if (timeSelectorElement) timeSelectorElement.disabled = false;
     // Enable all input fields in the table
    document.querySelectorAll('#sampleInputTableBody input').forEach(input => input.disabled = false);
}

/**
 * 回収数入力テーブルを生成
 */
function createSampleInputTable() {
    if (!sampleInputTableBody) return;
    sampleInputTableBody.innerHTML = ''; // Clear placeholder

    ageGroups5.forEach(ageRange => {
        const row = sampleInputTableBody.insertRow();
        row.insertCell().textContent = ageRange;

        const cellMale = row.insertCell();
        const inputMale = document.createElement('input');
        inputMale.type = 'number';
        inputMale.id = `sample_${ageRange}_male`;
        inputMale.min = "0";
        inputMale.value = "0";
        inputMale.className = "w-full px-1 py-1 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500";
        inputMale.addEventListener('input', updateTotalCounts); // 入力時に合計を更新
        cellMale.appendChild(inputMale);

        const cellFemale = row.insertCell();
        const inputFemale = document.createElement('input');
        inputFemale.type = 'number';
        inputFemale.id = `sample_${ageRange}_female`;
        inputFemale.min = "0";
        inputFemale.value = "0";
        inputFemale.className = "w-full px-1 py-1 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500";
        inputFemale.addEventListener('input', updateTotalCounts); // 入力時に合計を更新
        cellFemale.appendChild(inputFemale);
    });
    updateTotalCounts(); // 初期合計を計算
}

/**
 * 入力された回収数の合計を計算して表示
 */
function updateTotalCounts() {
    let totalMale = 0;
    let totalFemale = 0;
    ageGroups5.forEach(ageRange => {
        const maleInput = document.getElementById(`sample_${ageRange}_male`);
        const femaleInput = document.getElementById(`sample_${ageRange}_female`);
        totalMale += parseInt(maleInput?.value || 0, 10);
        totalFemale += parseInt(femaleInput?.value || 0, 10);
    });
    if (totalMaleElement) totalMaleElement.textContent = totalMale;
    if (totalFemaleElement) totalFemaleElement.textContent = totalFemale;
    if (grandTotalElement) grandTotalElement.textContent = totalMale + totalFemale;
}

/**
 * フォームから実際の回収数を取得
 * @returns {object|null} 回収数データオブジェクト、またはエラー時 null
 */
function getActualSampleCounts() {
    const actualCounts = {};
    let grandTotal = 0;
    let isValid = true;
    ageGroups5.forEach(ageRange => {
        actualCounts[ageRange] = {};
        ['male', 'female'].forEach(gender => {
            const inputElement = document.getElementById(`sample_${ageRange}_${gender}`);
            const value = parseInt(inputElement?.value || 0, 10);
            if (isNaN(value) || value < 0) {
                console.error(`Invalid input for ${ageRange} ${gender}: ${inputElement?.value}`);
                 if(errorMessage) {
                     errorMessage.textContent = `入力エラー: ${ageRange} ${gender === 'male' ? '男性':'女性'} の値が不正です。0以上の数値を入力してください。`;
                     errorMessage.classList.remove('hidden');
                 }
                 isValid = false;
            }
            actualCounts[ageRange][gender] = value;
            grandTotal += value;
        });
    });

    if (!isValid) return null;

     if (grandTotal <= 0) {
         if(errorMessage) {
             errorMessage.textContent = "入力エラー: 回収数の合計が0です。数値を入力してください。";
             errorMessage.classList.remove('hidden');
         }
         return null;
     }
     if(errorMessage) errorMessage.classList.add('hidden'); // エラーなければ隠す
    return { counts: actualCounts, total: grandTotal };
}

// --- ウェイトバック計算 ---
/**
 * ウェイトバック集計を実行
 * @param {object} actualSampleData - getActualSampleCountsの戻り値 { counts: object, total: number }
 * @param {object} targetRatios - calculateRatiosForTimeの戻り値 (calculatedRatios5)
 * @returns {object} 計算結果 { weights: object, weightedCounts: object, totalWeighted: number }
 */
function calculateWeightback(actualSampleData, targetRatios) {
    console.log("calculateWeightback: 開始", actualSampleData, targetRatios);
    const actualCounts = actualSampleData.counts;
    const actualTotal = actualSampleData.total;
    const weights = {};
    const weightedCounts = {};
    let totalWeighted = 0;

    if (!targetRatios || actualTotal <= 0) {
        throw new Error("ウェイトバック計算に必要なデータが不足しています。");
    }

    ageGroups5.forEach(ageRange => {
        weights[ageRange] = {};
        weightedCounts[ageRange] = {};
        ['male', 'female'].forEach(gender => {
            const targetRatio = targetRatios[ageRange]?.[gender] || 0;
            const actualCount = actualCounts[ageRange]?.[gender] || 0;
            const actualRatio = actualTotal > 0 ? actualCount / actualTotal : 0;

            let weight = 0;
            if (actualRatio > 0 && targetRatio > 0) {
                // 基本ウェイト = 目標構成比 / 実績構成比
                weight = targetRatio / actualRatio;
            } else if (targetRatio > 0 && actualCount === 0) {
                // 目標はあるが実績がない場合 -> ウェイト定義不能（または大きな値やエラー）
                console.warn(`ウェイト計算不能: ${ageRange} ${gender} - 目標構成比 ${targetRatio.toFixed(4)} に対して実績回収数が0です。`);
                weight = 0; // または他の処理 (例: このセルを除外)
            } else {
                // 目標構成比が0の場合など
                weight = 0;
            }

            weights[ageRange][gender] = weight;
            const weightedCount = actualCount * weight;
            weightedCounts[ageRange][gender] = weightedCount;
            totalWeighted += weightedCount;
        });
    });

    // 結果の正規化（任意）：ウェイト計が実績計と一致するように調整
    if (totalWeighted > 0 && actualTotal > 0) {
        const scalingFactor = actualTotal / totalWeighted;
        console.log("calculateWeightback: 正規化係数:", scalingFactor);
        totalWeighted = 0; // 再計算
        ageGroups5.forEach(ageRange => {
            ['male', 'female'].forEach(gender => {
                // weights[ageRange][gender] *= scalingFactor; // ウェイト自体を調整する場合
                weightedCounts[ageRange][gender] *= scalingFactor; // 集計値のみ調整する場合
                totalWeighted += weightedCounts[ageRange][gender];
            });
        });
    }

    console.log("calculateWeightback: 計算後のウェイト:", weights);
    console.log("calculateWeightback: 計算後のウェイト後回収数:", weightedCounts);
    console.log("calculateWeightback: ウェイト後合計:", totalWeighted);

    return { weights, weightedCounts, totalWeighted };
}

/**
 * ウェイトバック結果をHTMLテーブルで表示
 * @param {object} results - calculateWeightbackの戻り値
 * @param {object} actualSampleData - getActualSampleCountsの戻り値
 * @param {object} targetRatios - calculateRatiosForTimeの戻り値
 */
function displayWeightbackResults(results, actualSampleData, targetRatios) {
    console.log("displayWeightbackResults: 開始");
    if (!resultsArea) return;

    const { weights, weightedCounts, totalWeighted } = results;
    const actualCounts = actualSampleData.counts;
    const actualTotal = actualSampleData.total;

    let html = `
        <p class="text-sm mb-4">人口構成比（${selectedTimeCode.substring(0,4)}年${parseInt(selectedTimeCode.substring(6,8),10)}月）に合わせてウェイトバック集計した結果です。</p>
        <div class="overflow-x-auto">
            <table class="min-w-full border-collapse border border-gray-300 text-sm text-center">
                <thead class="bg-gray-100">
                    <tr>
                        <th rowspan="2" class="border px-2 py-1 align-middle">年齢階級</th>
                        <th colspan="2" class="border px-2 py-1">実績回収数</th>
                        <th colspan="2" class="border px-2 py-1">目標構成比</th>
                        <th colspan="2" class="border px-2 py-1">ウェイト値</th>
                        <th colspan="2" class="border px-2 py-1">ウェイト後回収数</th>
                    </tr>
                    <tr>
                        <th class="border px-1 py-1">男性</th><th class="border px-1 py-1">女性</th>
                        <th class="border px-1 py-1">男性</th><th class="border px-1 py-1">女性</th>
                        <th class="border px-1 py-1">男性</th><th class="border px-1 py-1">女性</th>
                        <th class="border px-1 py-1">男性</th><th class="border px-1 py-1">女性</th>
                    </tr>
                </thead>
                <tbody>
    `;

    let totalActualMale = 0, totalActualFemale = 0;
    let totalWeightedMale = 0, totalWeightedFemale = 0;

    ageGroups5.forEach(ageRange => {
        const actualM = actualCounts[ageRange]?.male || 0;
        const actualF = actualCounts[ageRange]?.female || 0;
        const targetRatioM = (targetRatios[ageRange]?.male || 0) * 100;
        const targetRatioF = (targetRatios[ageRange]?.female || 0) * 100;
        const weightM = weights[ageRange]?.male || 0;
        const weightF = weights[ageRange]?.female || 0;
        const weightedM = weightedCounts[ageRange]?.male || 0;
        const weightedF = weightedCounts[ageRange]?.female || 0;

        totalActualMale += actualM;
        totalActualFemale += actualF;
        totalWeightedMale += weightedM;
        totalWeightedFemale += weightedF;

        html += `
            <tr>
                <td class="border px-2 py-1 font-medium">${ageRange}</td>
                <td class="border px-2 py-1 text-right">${actualM}</td>
                <td class="border px-2 py-1 text-right">${actualF}</td>
                <td class="border px-2 py-1 text-right">${targetRatioM.toFixed(2)}%</td>
                <td class="border px-2 py-1 text-right">${targetRatioF.toFixed(2)}%</td>
                <td class="border px-2 py-1 text-right">${weightM.toFixed(3)}</td>
                <td class="border px-2 py-1 text-right">${weightF.toFixed(3)}</td>
                <td class="border px-2 py-1 text-right font-semibold">${weightedM.toFixed(1)}</td>
                <td class="border px-2 py-1 text-right font-semibold">${weightedF.toFixed(1)}</td>
            </tr>
        `;
    });

     html += `
                </tbody>
                <tfoot class="bg-gray-50 font-bold">
                    <tr>
                        <td class="border px-2 py-1 text-right">実績 合計</td>
                        <td class="border px-2 py-1 text-right">${totalActualMale}</td>
                        <td class="border px-2 py-1 text-right">${totalActualFemale}</td>
                        <td colspan="4"></td> {/* Target Ratio and Weight totals are meaningless */}
                        <td class="border px-2 py-1 text-right">${totalWeightedMale.toFixed(1)}</td>
                        <td class="border px-2 py-1 text-right">${totalWeightedFemale.toFixed(1)}</td>
                    </tr>
                     <tr>
                        <td class="border px-2 py-1 text-right">総合計</td>
                        <td colspan="2" class="border px-2 py-1 text-center">${actualTotal}</td>
                        <td colspan="4"></td>
                        <td colspan="2" class="border px-2 py-1 text-center">${totalWeighted.toFixed(1)}</td>
                    </tr>
                </tfoot>
            </table>
        </div>
    `;

    resultsArea.innerHTML = html;
    if(resultsSection) resultsSection.classList.remove('hidden'); // 結果セクションを表示
    console.log("displayWeightbackResults: 結果表示完了");
}


// --- イベントハンドラ ---
/**
 * 集計実行ボタンのクリックハンドラ
 */
function handleCalculateSubmit(event) {
    event.preventDefault(); // Prevent form submission
    console.log("handleCalculateSubmit: 開始");

    if (!isDataReady || !calculatedRatios5) {
        console.warn("handleCalculateSubmit: データ未準備");
        if(errorMessage) { errorMessage.textContent = "人口構成比データが準備できていません。"; errorMessage.classList.remove('hidden'); }
        return;
    }

    const actualSampleData = getActualSampleCounts();
    if (!actualSampleData) {
        console.warn("handleCalculateSubmit: 回収数入力エラー");
        return; // エラーメッセージは getActualSampleCounts 内で表示
    }

    toggleLoading(true); // 計算中表示
    if(errorMessage) errorMessage.classList.add('hidden');
    if(resultsSection) resultsSection.classList.add('hidden'); // 古い結果を隠す

    // 少し待ってUIを更新
    setTimeout(() => {
        try {
            const results = calculateWeightback(actualSampleData, calculatedRatios5);
            displayWeightbackResults(results, actualSampleData, calculatedRatios5);
        } catch (error) {
            console.error("handleCalculateSubmit: 計算または表示エラー:", error);
            if(errorMessage) { errorMessage.textContent = `集計エラー: ${error.message || '詳細不明'}`; errorMessage.classList.remove('hidden'); }
        } finally {
            toggleLoading(false); // 計算完了表示
        }
    }, 50);
}

/**
 * 結果CSVダウンロード処理 (新規追加)
 */
function downloadResultsCSV() {
    console.log("downloadResultsCSV: 開始");
    const resultsTable = resultsArea.querySelector('table');
    if (!resultsTable) {
        console.warn("downloadResultsCSV: 結果テーブルが見つかりません。");
        alert("ダウンロードする結果がありません。"); // Use alert for simplicity here
        return;
    }

    let csv = "\uFEFF"; // BOM for Excel

    // Header Rows
    const headerRows = Array.from(resultsTable.querySelectorAll('thead tr'));
    headerRows.forEach(tr => {
        const cells = Array.from(tr.querySelectorAll('th'));
        // Handle colspan and rowspan for header (simplified: just get text)
        const row = cells.map(cell => `"${(cell.textContent || "").replace(/"/g, '""')}"`).join(",");
        csv += row + "\r\n";
    });


    // Body Rows
    const bodyRows = Array.from(resultsTable.querySelectorAll('tbody tr'));
    bodyRows.forEach(tr => {
        const cells = Array.from(tr.querySelectorAll('td'));
        const row = cells.map(cell => `"${(cell.textContent || "").replace(/"/g, '""')}"`).join(",");
        csv += row + "\r\n";
    });

     // Footer Rows
    const footerRows = Array.from(resultsTable.querySelectorAll('tfoot tr'));
    footerRows.forEach(tr => {
        const cells = Array.from(tr.querySelectorAll('td'));
         // Handle colspan for footer (simplified: just get text)
        const row = cells.map(cell => `"${(cell.textContent || "").replace(/"/g, '""')}"`).join(",");
        csv += row + "\r\n";
    });


    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const selectedOption = timeSelectorElement.options[timeSelectorElement.selectedIndex];
    const timeText = selectedOption ? selectedOption.textContent.replace(/[^0-9]/g,'') : 'unknown_date'; // YYYYMM
    a.download = `ウェイトバック集計結果_${timeText}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    console.log("downloadResultsCSV: 完了");
}


// --- 初期化・イベントリスナー設定 ---
async function initializePage() {
    console.log("initializePage: 開始");
    disableForm("データ準備中...");

    let cachedApiData = getCachedData();
    let apiDataToProcess;

    if (cachedApiData) {
        console.log("initializePage: キャッシュデータを使用");
        apiDataToProcess = cachedApiData;
    } else {
        try {
            const apiResponseData = await fetchDataFromApi();
            const { valueDataAll } = processRawApiData(apiResponseData);
            apiDataToProcess = valueDataAll;
            saveDataToCache(apiDataToProcess);
            console.log("initializePage: APIからデータを取得・保存");
        } catch (error) {
            console.error("initializePage: データ取得フェーズでエラー:", error);
            displayInitializationError(`初期データ取得エラー: ${error.message}`);
            return;
        }
    }

    rawApiDataCache = apiDataToProcess; // Keep raw data for time changes

    try {
        const timeCodes = [...new Set(apiDataToProcess.map(item => item[config.TIME_KEY]))].sort().reverse();
        availableTimeCodes = timeCodes;
        populateTimeSelector(availableTimeCodes); // ドロップダウン生成

        // 入力テーブル生成
        createSampleInputTable();

        if (selectedTimeCode) {
            calculateRatiosForTime(selectedTimeCode); // 初回構成比計算
        } else if (availableTimeCodes.length > 0) {
            selectedTimeCode = availableTimeCodes[0];
            calculateRatiosForTime(selectedTimeCode);
        } else {
             console.error("initializePage: 利用可能なデータ時点が見つかりません。");
             displayInitializationError("エラー: 利用可能なデータ時点が見つかりません。");
        }
    } catch(error) {
         console.error("initializePage: 時間選択、テーブル生成、または初期比率計算でエラー:", error);
         displayInitializationError(`データ処理エラー: ${error.message}`);
    }
}

document.addEventListener('DOMContentLoaded', () => {
    console.log("DOMContentLoaded: イベント発火");

    // DOM要素取得
    sampleInputForm = document.getElementById("sampleInputForm");
    sampleInputTableBody = document.getElementById("sampleInputTableBody");
    timeSelectorElement = document.getElementById("timeSelector");
    timeSelectorContainer = document.getElementById("timeSelectorContainer");
    calculateButton = document.getElementById("calculateWeightback");
    loadingSpinner = document.getElementById("loadingSpinner");
    buttonText = document.getElementById("buttonText");
    resultsSection = document.getElementById("resultsSection");
    resultsArea = document.getElementById("resultsArea");
    errorMessage = document.getElementById("errorMessage");
    totalMaleElement = document.getElementById("totalMale");
    totalFemaleElement = document.getElementById("totalFemale");
    grandTotalElement = document.getElementById("grandTotal");
    downloadResultsCsvButton = document.getElementById('downloadResultsCsvButton'); // 結果CSVボタン

    // 必須要素チェック
    const essentialElements = [sampleInputForm, sampleInputTableBody, timeSelectorElement, timeSelectorContainer, calculateButton, loadingSpinner, buttonText, resultsSection, resultsArea, errorMessage, totalMaleElement, totalFemaleElement, grandTotalElement, downloadResultsCsvButton];
    if (essentialElements.some(el => !el)) {
        console.error("DOMContentLoaded: 必須DOM要素の一部が見つかりません。HTMLを確認してください。");
        const body = document.querySelector('body');
        if(body && !document.getElementById('init-error-msg')) {
            const initError = document.createElement('div'); initError.id = 'init-error-msg'; initError.textContent = "ページの初期化に失敗しました。必須要素が見つかりません。"; initError.style.color = 'red'; initError.style.padding = '10px'; initError.style.border = '1px solid red'; initError.style.margin = '10px'; body.prepend(initError);
        }
        return;
    }
    console.log("DOMContentLoaded: 全ての必須DOM要素を取得完了");

    // 初期化処理
    if (timeSelectorContainer) timeSelectorContainer.style.display = 'none'; // Initially hide
    initializePage(); // データ取得・初期化開始

    // イベントリスナー設定
    if (sampleInputForm) sampleInputForm.addEventListener("submit", handleCalculateSubmit);

    if (timeSelectorElement) {
        timeSelectorElement.addEventListener('change', (event) => {
            const newTimeCode = event.target.value;
            if (newTimeCode && newTimeCode !== selectedTimeCode) {
                console.log("Time selector changed:", newTimeCode);
                selectedTimeCode = newTimeCode;
                disableForm("構成比 再計算中..."); // ボタンを無効化＆テキスト変更
                if(resultsSection) resultsSection.classList.add('hidden'); // 古い結果を隠す
                if(errorMessage) errorMessage.classList.add('hidden'); // 古いエラーを隠す

                setTimeout(() => {
                    calculateRatiosForTime(selectedTimeCode); // 再計算実行 (完了後に enableForm が呼ばれる)
                }, 50);
            }
        });
    }
    console.log("DOMContentLoaded: 初期化・イベントリスナー設定完了");
});

console.log("weightback.js: スクリプト終端");

