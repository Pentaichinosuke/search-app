// ============================================================
// 配列検索
// script.js
// ============================================================


// ============================================================
// 弾の設定
// ============================================================
//
// 新しい弾を追加するときは、ここに追加するだけです。
// 例：
// {
//     name: "VR5",
//     file: "data_VR5.xlsx",
//     redNumbers: [1, 2, 3, 4]
// }
//
// HTML側のメニューを変更する必要はありません。
// ============================================================

const ROUNDS = [
    {
        name: "VR2",
        file: "data_VR2.xlsx",

        // VR2で赤色にする数字
        redNumbers: [
            1, 13, 25, 29, 37, 42, 46, 51
        ]
    },

    {
        name: "VR3",
        file: "data_VR3.xlsx",

        // VR3で赤色にする数字
        redNumbers: [
            7, 14, 24, 40, 44, 51, 54, 58
        ]
    },

    {
        name: "VR4",
        file: "data_VR4.xlsx",

        // VR4で赤色にする数字
        redNumbers: [
            1, 5, 25, 40, 46, 49, 53, 58
        ]
    }
];


// ============================================================
// 共通の色設定
// ============================================================

// 緑色にする数字の範囲
const GREEN_MIN = 61;
const GREEN_MAX = 70;

// ピンク色にする文字
const PINK_TEXT = "P";


// ============================================================
// 現在の状態
// ============================================================

// 初期表示する弾
let currentRound = "VR4";

// 初期表示するExcelファイル
let currentFile = "data_VR4.xlsx";


// ============================================================
// DOMの取得
// ============================================================

const sheetsContainer = document.getElementById("sheetsContainer");
const searchInput = document.getElementById("searchInput");
const searchButton = document.getElementById("searchButton");
const resetButton = document.getElementById("resetButton");

const menuButton = document.getElementById("menuButton");
const sideMenu = document.getElementById("sideMenu");
const overlay = document.getElementById("overlay");

const currentRoundDisplay =
    document.getElementById("currentRound");


// ============================================================
// ページ読み込み時
// ============================================================

document.addEventListener("DOMContentLoaded", () => {

    // 弾のメニューを自動生成
    createRoundMenu();

    // 初期表示
    loadExcel(currentFile, currentRound);

});


// ============================================================
// 弾のメニューを自動生成
// ============================================================
//
// ROUNDSに弾を追加すると、ここで自動的にボタンが作られます。
// ============================================================

function createRoundMenu() {

    // メニュー内の弾ボタンを入れる場所
    const roundMenu = document.getElementById("roundMenu");

    if (!roundMenu) {
        console.error("roundMenu が見つかりません。");
        return;
    }

    // 一度空にする
    roundMenu.innerHTML = "";

    // ROUNDSの内容からボタンを作る
    ROUNDS.forEach(round => {

        const button = document.createElement("button");

        button.textContent = round.name;

        button.classList.add("round-button");

        // 現在選択中の弾
        if (round.name === currentRound) {
            button.classList.add("active");
        }

        // クリック時
        button.addEventListener("click", () => {

            // 選択された弾に変更
            currentRound = round.name;
            currentFile = round.file;

            // Excelを読み込む
            loadExcel(currentFile, currentRound);

            // 検索欄をリセット
            searchInput.value = "";

            // 検索結果の色もリセット
            clearSearchHighlight();

            // メニューを閉じる
            closeMenu();

            // ボタンの選択状態を更新
            updateRoundMenu();

        });

        roundMenu.appendChild(button);

    });

}


// ============================================================
// 弾メニューの選択状態を更新
// ============================================================

function updateRoundMenu() {

    const buttons =
        document.querySelectorAll(".round-button");

    buttons.forEach(button => {

        if (button.textContent === currentRound) {
            button.classList.add("active");
        } else {
            button.classList.remove("active");
        }

    });

}


// ============================================================
// Excelファイルを読み込む
// ============================================================

async function loadExcel(fileName, roundName) {

    try {

        // 「読み込み中」を表示
        sheetsContainer.innerHTML =
            "<p>データを読み込んでいます...</p>";

        // 現在の弾を表示
        if (currentRoundDisplay) {
            currentRoundDisplay.textContent = roundName;
        }

        // Excelファイルを取得
        const response = await fetch(fileName);

        if (!response.ok) {
            throw new Error(
                `Excelファイルを読み込めませんでした: ${fileName}`
            );
        }

        // ArrayBufferに変換
        const arrayBuffer = await response.arrayBuffer();

        // XLSXで読み込み
        const workbook =
            XLSX.read(arrayBuffer, {
                type: "array"
            });

        // 表示部分を空にする
        sheetsContainer.innerHTML = "";

        // 現在の弾の設定を取得
        const roundData =
            ROUNDS.find(round => round.name === roundName);

        // 赤色にする数字
        const redNumbers =
            roundData ? roundData.redNumbers : [];

        // Excelの各シートを処理
        workbook.SheetNames.forEach(sheetName => {

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
                redNumbers
            );

        });

    } catch (error) {

        console.error(error);

        sheetsContainer.innerHTML = `
            <p class="error-message">
                データの読み込みに失敗しました。<br>
                ${error.message}
            </p>
        `;

    }

}


// ============================================================
// シートを画面に表示
// ============================================================

function renderSheet(
    sheetName,
    data,
    redNumbers
) {

    // シート全体
    const sheetWrapper =
        document.createElement("div");

    sheetWrapper.classList.add("sheet-wrapper");


    // シート名
    const sheetTitle =
        document.createElement("h2");

    sheetTitle.textContent = sheetName;

    sheetWrapper.appendChild(sheetTitle);


    // テーブル
    const table =
        document.createElement("table");

    table.classList.add("data-table");


    // ========================================================
    // 各行
    // ========================================================

    data.forEach((row, rowIndex) => {

        const tr =
            document.createElement("tr");


        // ====================================================
        // 各列
        // ====================================================

        row.forEach((cell, colIndex) => {

            const td =
                document.createElement(
                    rowIndex === 0
                        ? "th"
                        : "td"
                );


            // セルの文字
            const text =
                String(cell).trim();


            // 数値として取得
            const value =
                Number(cell);


            // =================================================
            // table列
            // =================================================
            //
            // 0列目の「table」は非表示
            // =================================================

            if (colIndex === 0) {
                td.style.display = "none";
            }


            // =================================================
            // row列
            // =================================================
            //
            // rowは表示するが検索対象にはしない
            // =================================================

            if (colIndex === 1) {

                td.classList.add("row-column");

                // 検索対象外
                td.dataset.searchable = "false";

            }


            // =================================================
            // col1 / col2
            // =================================================

            if (colIndex === 2) {

                td.classList.add("col1-column");

            }

            if (colIndex === 3) {

                td.classList.add("col2-column");

            }


            // =================================================
            // 通常のセル
            // =================================================

            td.textContent = text;


            // =================================================
            // データ検索対象
            // =================================================
            //
            // row列以外を検索対象にする
            // =================================================

            if (rowIndex !== 0 && colIndex !== 1) {

                td.dataset.searchable = "true";

            }


            // =================================================
            // ヘッダー行
            // =================================================

            if (rowIndex === 0) {

                td.dataset.searchable = "false";

            }


            // =================================================
            // 赤色
            // =================================================
            //
            // 弾ごとに設定された数字を赤色にする
            // =================================================

            if (
                rowIndex !== 0 &&
                !isNaN(value) &&
                redNumbers.includes(value)
            ) {

                td.classList.add("red-number");

            }


            // =================================================
            // 緑色
            // =================================================
            //
            // 61～70を緑色にする
            // =================================================

            if (
                rowIndex !== 0 &&
                !isNaN(value) &&
                value >= GREEN_MIN &&
                value <= GREEN_MAX
            ) {

                td.classList.add("green-number");

            }


            // =================================================
            // P / p
            // =================================================
            //
            // P または p のセルをピンク色にする
            // =================================================

            if (
                rowIndex !== 0 &&
                text.toUpperCase() === PINK_TEXT
            ) {

                td.classList.add("pink-number");

            }


            // =================================================
            // セルを追加
            // =================================================

            tr.appendChild(td);

        });


        // 行を追加
        table.appendChild(tr);

    });


    // テーブルを追加
    sheetWrapper.appendChild(table);

    // シートを画面に追加
    sheetsContainer.appendChild(sheetWrapper);

}


// ============================================================
// 検索
// ============================================================

function searchData() {

    // 検索文字
    const keyword =
        searchInput.value.trim();

    // まず以前の検索結果を解除
    clearSearchHighlight();


    // 検索文字が空なら終了
    if (keyword === "") {
        return;
    }


    // 検索対象セル
    const cells =
        document.querySelectorAll(
            'td[data-searchable="true"]'
        );


    // 各セルを確認
    cells.forEach(cell => {

        const text =
            cell.textContent.trim();


        // 数字検索
        if (text === keyword) {

            cell.classList.add("search-hit");

        }

    });

}


// ============================================================
// 検索結果の色を解除
// ============================================================

function clearSearchHighlight() {

    const cells =
        document.querySelectorAll(
            ".search-hit"
        );

    cells.forEach(cell => {

        cell.classList.remove(
            "search-hit"
        );

    });

}


// ============================================================
// 検索ボタン
// ============================================================

if (searchButton) {

    searchButton.addEventListener(
        "click",
        searchData
    );

}


// ============================================================
// Enterキーで検索
// ============================================================

if (searchInput) {

    searchInput.addEventListener(
        "keydown",
        event => {

            if (event.key === "Enter") {

                searchData();

            }

        }
    );

}


// ============================================================
// リセットボタン
// ============================================================

if (resetButton) {

    resetButton.addEventListener(
        "click",
        () => {

            // 入力欄を空にする
            searchInput.value = "";

            // 検索結果を解除
            clearSearchHighlight();

        }
    );

}


// ============================================================
// ハンバーガーメニューを開く
// ============================================================

if (menuButton) {

    menuButton.addEventListener(
        "click",
        openMenu
    );

}


// ============================================================
// メニューを開く
// ============================================================

function openMenu() {

    if (sideMenu) {
        sideMenu.classList.add("open");
    }

    if (overlay) {
        overlay.classList.add("show");
    }

}


// ============================================================
// メニューを閉じる
// ============================================================

function closeMenu() {

    if (sideMenu) {
        sideMenu.classList.remove("open");
    }

    if (overlay) {
        overlay.classList.remove("show");
    }

}


// ============================================================
// 背景をクリックしてメニューを閉じる
// ============================================================

if (overlay) {

    overlay.addEventListener(
        "click",
        closeMenu
    );

}


// ============================================================
// Escapeキーでメニューを閉じる
// ============================================================

document.addEventListener(
    "keydown",
    event => {

        if (event.key === "Escape") {

            closeMenu();

        }

    }
);