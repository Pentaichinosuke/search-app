// ============================================================
// 配列検索
// script.js
// ============================================================

// ============================================================
// 弾の設定
// ============================================================
// 新しい弾を追加するときは、ここに1つ追加するだけです。
// HTMLのメニューを変更する必要はありません。
// ============================================================

const ROUNDS = [
    {
        name: "VR2",
        file: "data_VR2.xlsx",
        redNumbers: [1, 13, 25, 29, 37, 42, 46, 51]
    },
    {
        name: "VR3",
        file: "data_VR3.xlsx",
        redNumbers: [7, 14, 24, 40, 44, 51, 54, 58]
    },
    {
        name: "VR4",
        file: "data_VR4.xlsx",
        redNumbers: [1, 5, 25, 40, 46, 49, 53, 58]
    }
];


// ============================================================
// 共通設定
// ============================================================

const GREEN_MIN = 61;
const GREEN_MAX = 70;

const PINK_TEXT = "P";


// ============================================================
// 現在の状態
// ============================================================

let currentRound = "VR4";


// ============================================================
// DOM
// ============================================================

let sheetsContainer;
let searchInput;
let searchButton;
let resetButton;

let menuButton;
let sideMenu;
let overlay;

let currentRoundDisplay;
let roundMenu;


// ============================================================
// ページ読み込み後に初期化
// ============================================================

document.addEventListener("DOMContentLoaded", () => {

    // DOMを取得
    sheetsContainer =
        document.getElementById("sheetsContainer");

    searchInput =
        document.getElementById("searchInput");

    searchButton =
        document.getElementById("searchBtn");

    resetButton =
        document.getElementById("resetBtn");

    menuButton =
        document.getElementById("menuBtn");

    sideMenu =
        document.getElementById("sideMenu");

    overlay =
        document.getElementById("menuOverlay");

    currentRoundDisplay =
        document.getElementById("currentRound");

    roundMenu =
        document.getElementById("roundMenu");


    // 必要なHTML要素があるか確認
    if (
        !sheetsContainer ||
        !searchInput ||
        !searchButton ||
        !resetButton ||
        !menuButton ||
        !sideMenu ||
        !overlay ||
        !currentRoundDisplay ||
        !roundMenu
    ) {
        console.error(
            "必要なHTML要素が見つかりません。"
        );

        return;
    }


    // 弾メニューを作成
    createRoundMenu();


    // ========================================================
    // 検索ボタン
    // ========================================================

    searchButton.addEventListener(
        "click",
        searchData
    );


    // ========================================================
    // リセットボタン
    // ========================================================

    resetButton.addEventListener(
        "click",
        resetSearch
    );


    // ========================================================
    // Enterキーで検索
    // ========================================================

    searchInput.addEventListener(
        "keydown",
        (event) => {

            if (event.key === "Enter") {
                searchData();
            }

        }
    );


    // ========================================================
    // ハンバーガーメニュー
    // ========================================================

    menuButton.addEventListener(
        "click",
        toggleMenu
    );


    // メニュー外側をクリック
    overlay.addEventListener(
        "click",
        closeMenu
    );


    // Escキーで閉じる
    document.addEventListener(
        "keydown",
        (event) => {

            if (event.key === "Escape") {
                closeMenu();
            }

        }
    );


    // ========================================================
    // 初期表示
    // ========================================================

    loadRound(currentRound);

});


// ============================================================
// 弾メニューを自動生成
// ============================================================

function createRoundMenu() {

    roundMenu.innerHTML = "";


    ROUNDS.forEach((round) => {

        const button =
            document.createElement("button");


        button.type = "button";

        button.className = "round-btn";

        button.textContent = round.name;


        // 弾をクリックしたとき
        button.addEventListener(
            "click",
            () => {

                loadRound(round.name);

                closeMenu();

            }
        );


        roundMenu.appendChild(button);

    });


    updateRoundMenu();

}


// ============================================================
// 弾を切り替える
// ============================================================

function loadRound(roundName) {

    const round =
        ROUNDS.find(
            (item) =>
                item.name === roundName
        );


    if (!round) {

        console.error(
            "指定された弾が見つかりません:",
            roundName
        );

        return;
    }


    // 現在の弾を変更
    currentRound = round.name;


    // 画面右上の表示を変更
    currentRoundDisplay.textContent =
        currentRound;


    // 検索をリセット
    searchInput.value = "";

    clearSearchHighlight();


    // メニューの選択状態を更新
    updateRoundMenu();


    // Excelを読み込む
    loadExcel(
        round.file,
        round
    );

}


// ============================================================
// メニューの選択状態を更新
// ============================================================

function updateRoundMenu() {

    const buttons =
        roundMenu.querySelectorAll(
            ".round-btn"
        );


    buttons.forEach((button) => {

        button.classList.toggle(
            "active",
            button.textContent === currentRound
        );

    });

}


// ============================================================
// Excel読み込み
// ============================================================

async function loadExcel(
    fileName,
    round
) {

    sheetsContainer.innerHTML =
        '<p class="loading-message">' +
        'データを読み込んでいます...' +
        '</p>';


    try {

        // Excelファイルを取得
        const response =
            await fetch(fileName);


        if (!response.ok) {

            throw new Error(
                `Excelファイルを読み込めませんでした: ${fileName}`
            );

        }


        // ArrayBufferへ変換
        const arrayBuffer =
            await response.arrayBuffer();


        // Excelを読み込む
        const workbook =
            XLSX.read(
                arrayBuffer,
                {
                    type: "array"
                }
            );


        // 既存の表示を削除
        sheetsContainer.innerHTML = "";


        // 各シートを処理
        workbook.SheetNames.forEach(
            (sheetName) => {

                const worksheet =
                    workbook.Sheets[sheetName];


                // Excel → 配列
                const data =
                    XLSX.utils.sheet_to_json(
                        worksheet,
                        {
                            header: 1,
                            defval: ""
                        }
                    );


                // シートを表示
                renderSheet(
                    sheetName,
                    data,
                    round.redNumbers
                );

            }
        );


    } catch (error) {

        console.error(error);


        sheetsContainer.innerHTML = `
            <p class="error-message">
                データの読み込みに失敗しました。<br>
                ${escapeHtml(error.message)}
            </p>
        `;

    }

}


// ============================================================
// シート表示
// ============================================================

function renderSheet(
    sheetName,
    data,
    redNumbers
) {

    // シート全体
    const sheetWrapper =
        document.createElement("section");

    sheetWrapper.className =
        "sheet";


    // シート名
    const sheetTitle =
        document.createElement("h2");

    sheetTitle.textContent =
        sheetName;


    // テーブル
    const table =
        document.createElement("table");


    // ========================================================
    // 各行
    // ========================================================

    data.forEach(
        (row, rowIndex) => {

            const tr =
                document.createElement("tr");


            // =================================================
            // 各列
            // =================================================

            row.forEach(
                (cell, colIndex) => {

                    const element =
                        document.createElement(
                            rowIndex === 0
                                ? "th"
                                : "td"
                        );


                    const text =
                        String(cell).trim();


                    const value =
                        Number(cell);


                    // =========================================
                    // table列
                    // =========================================
                    // 0列目のtableを非表示
                    // =========================================

                    if (colIndex === 0) {

                        element.style.display =
                            "none";

                    }


                    // =========================================
                    // row列
                    // =========================================
                    // 表示するが、
                    // 検索・色付けの対象にはしない
                    // =========================================

                    if (colIndex === 1) {

                        element.classList.add(
                            "row-column"
                        );

                        element.dataset.searchable =
                            "false";

                    }


                    // =========================================
                    // col1
                    // =========================================

                    if (colIndex === 2) {

                        element.classList.add(
                            "col1-column"
                        );

                    }


                    // =========================================
                    // col2
                    // =========================================

                    if (colIndex === 3) {

                        element.classList.add(
                            "col2-column"
                        );

                    }


                    // セルの文字
                    element.textContent =
                        text;


                    // =========================================
                    // 検索対象
                    // =========================================
                    // row列とヘッダーは対象外
                    // =========================================

                    if (
                        rowIndex !== 0 &&
                        colIndex !== 1
                    ) {

                        element.dataset.searchable =
                            "true";

                    } else {

                        element.dataset.searchable =
                            "false";

                    }


                    // =========================================
                    // 赤色
                    // =========================================
                    // row列は絶対に赤くしない
                    // =========================================

                    if (
                        rowIndex !== 0 &&
                        colIndex !== 1 &&
                        Number.isFinite(value) &&
                        redNumbers.includes(value)
                    ) {

                        element.classList.add(
                            "red-number"
                        );

                    }


                    // =========================================
                    // 緑色
                    // =========================================
                    // 61～70
                    // row列は色付けしない
                    // =========================================

                    if (
                        rowIndex !== 0 &&
                        colIndex !== 1 &&
                        Number.isFinite(value) &&
                        value >= GREEN_MIN &&
                        value <= GREEN_MAX
                    ) {

                        element.classList.add(
                            "green-number"
                        );

                    }


                    // =========================================
                    // P / p
                    // =========================================
                    // row列は色付けしない
                    // =========================================

                    if (
                        rowIndex !== 0 &&
                        colIndex !== 1 &&
                        text.toUpperCase() === PINK_TEXT
                    ) {

                        element.classList.add(
                            "pink-number"
                        );

                    }


                    // セルを行に追加
                    tr.appendChild(element);

                }
            );


            // 行をテーブルに追加
            table.appendChild(tr);

        }
    );


    // シート名を追加
    sheetWrapper.appendChild(
        sheetTitle
    );


    // テーブルを追加
    sheetWrapper.appendChild(
        table
    );


    // 画面に追加
    sheetsContainer.appendChild(
        sheetWrapper
    );

}


// ============================================================
// 検索
// ============================================================

function searchData() {

    const keyword =
        searchInput.value.trim();


    // 以前の検索結果を解除
    clearSearchHighlight();


    // 空欄なら終了
    if (keyword === "") {
        return;
    }


    // 検索対象のセル
    const cells =
        document.querySelectorAll(
            'td[data-searchable="true"]'
        );


    // 一つずつ確認
    cells.forEach(
        (cell) => {

            const text =
                cell.textContent.trim();


            // 完全一致
            if (text === keyword) {

                cell.classList.add(
                    "search-hit"
                );

            }

        }
    );

}


// ============================================================
// リセット
// ============================================================

function resetSearch() {

    // 入力欄を空にする
    searchInput.value = "";


    // 検索結果を解除
    clearSearchHighlight();

}


// ============================================================
// 検索結果の色を解除
// ============================================================

function clearSearchHighlight() {

    document
        .querySelectorAll(".search-hit")
        .forEach(
            (cell) => {

                cell.classList.remove(
                    "search-hit"
                );

            }
        );

}


// ============================================================
// ハンバーガーメニュー
// ============================================================

function toggleMenu() {

    const isOpen =
        sideMenu.classList.contains("open");


    if (isOpen) {

        closeMenu();

    } else {

        openMenu();

    }

}


// ============================================================
// メニューを開く
// ============================================================

function openMenu() {

    sideMenu.classList.add("open");

    overlay.classList.add("show");


    menuButton.setAttribute(
        "aria-expanded",
        "true"
    );


    menuButton.setAttribute(
        "aria-label",
        "弾のメニューを閉じる"
    );

}


// ============================================================
// メニューを閉じる
// ============================================================

function closeMenu() {

    sideMenu.classList.remove("open");

    overlay.classList.remove("show");


    if (menuButton) {

        menuButton.setAttribute(
            "aria-expanded",
            "false"
        );


        menuButton.setAttribute(
            "aria-label",
            "弾のメニューを開く"
        );

    }

}


// ============================================================
// エラーメッセージ用
// ============================================================

function escapeHtml(text) {

    return String(text)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");

}