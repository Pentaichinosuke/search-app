// ==============================
// 色を付ける数字
// ==============================

const RED_NUMBERS = [
    7,
    14,
    24,
    40,
    44,
    51,
    54,
    58
];


// 61～70は緑色
const GREEN_NUMBERS = [
    61,
    62,
    63,
    64,
    65,
    66,
    67,
    68,
    69,
    70
];


// ==============================
// 現在選択されている弾
// ==============================

let currentFile = "data_VR4.xlsx";
let currentRound = "VR4";


// ==============================
// Excelを読み込む関数
// ==============================

function loadExcel(fileName, roundName) {

    const container =
        document.getElementById("sheetsContainer");

    // 以前の表を消す
    container.innerHTML = "";

    // 現在の弾を表示
    document.getElementById(
        "currentRound"
    ).textContent = roundName;


    fetch(fileName)

    .then(response => {

        if (!response.ok) {
            throw new Error(
                "Excelファイルを読み込めませんでした"
            );
        }

        return response.arrayBuffer();

    })

    .then(data => {

        const workbook =
            XLSX.read(data);


        workbook.SheetNames.forEach(
            sheetName => {

                const worksheet =
                    workbook.Sheets[sheetName];


                const rows =
                    XLSX.utils.sheet_to_json(
                        worksheet,
                        {
                            header: 1
                        }
                    );


                // ==============================
                // Sheet部分
                // ==============================

                const section =
                    document.createElement(
                        "div"
                    );

                section.className =
                    "sheet";


                const title =
                    document.createElement(
                        "h2"
                    );

                title.textContent =
                    sheetName;


                section.appendChild(
                    title
                );


                // ==============================
                // table
                // ==============================

                const table =
                    document.createElement(
                        "table"
                    );


                rows.forEach(
                    (row, rowIndex) => {

                        const tr =
                            document.createElement(
                                "tr"
                            );


                        row.forEach(
                            (cell, colIndex) => {

                                // table列を非表示
                                if (colIndex === 0) {
                                    return;
                                }


                                const element =
                                    document.createElement(
                                        rowIndex === 0
                                            ? "th"
                                            : "td"
                                    );


                                element.textContent =
                                    cell ?? "";


                                // ==========================
                                // row列
                                // ==========================

                                if (colIndex === 1) {

                                    element.classList.add(
                                        "row-column"
                                    );

                                }


                                // ==========================
                                // col1・col2
                                // ==========================

                                if (
                                    rowIndex > 0 &&
                                    colIndex >= 2
                                ) {

                                    const text =
                                        String(
                                            cell
                                        ).trim();


                                    const value =
                                        Number(
                                            cell
                                        );


                                    // 検索対象
                                    element.dataset.searchable =
                                        "true";


                                    // ======================
                                    // P / p → ピンク
                                    // ======================

                                    if (
                                        text.toUpperCase()
                                        === "P"
                                    ) {

                                        element.classList.add(
                                            "pink-number"
                                        );

                                    }


                                    // ======================
                                    // 赤色
                                    // ======================

                                    if (
                                        RED_NUMBERS.includes(
                                            value
                                        )
                                    ) {

                                        element.classList.add(
                                            "red-number"
                                        );

                                    }


                                    // ======================
                                    // 緑色
                                    // ======================

                                    if (
                                        GREEN_NUMBERS.includes(
                                            value
                                        )
                                    ) {

                                        element.classList.add(
                                            "green-number"
                                        );

                                    }

                                }


                                tr.appendChild(
                                    element
                                );

                            }
                        );


                        table.appendChild(
                            tr
                        );

                    }
                );


                section.appendChild(
                    table
                );

                container.appendChild(
                    section
                );

            }
        );

    })

    .catch(error => {

        console.error(error);

        container.innerHTML =
            "<p>Excelファイルを読み込めませんでした。</p>";

    });

}


// ==============================
// 最初にVR4を読み込む
// ==============================

loadExcel(
    currentFile,
    currentRound
);


// ==============================
// ハンバーガーメニュー
// ==============================

const menuBtn =
    document.getElementById(
        "menuBtn"
    );

const sideMenu =
    document.getElementById(
        "sideMenu"
    );

const menuOverlay =
    document.getElementById(
        "menuOverlay"
    );


function openMenu() {

    sideMenu.classList.add(
        "open"
    );

    menuOverlay.classList.add(
        "show"
    );

}


function closeMenu() {

    sideMenu.classList.remove(
        "open"
    );

    menuOverlay.classList.remove(
        "show"
    );

}


menuBtn.addEventListener(
    "click",
    openMenu
);


menuOverlay.addEventListener(
    "click",
    closeMenu
);


// ==============================
// 弾の切り替え
// ==============================

document
.querySelectorAll(".round-btn")
.forEach(button => {

    button.addEventListener(
        "click",
        () => {

            const fileName =
                button.dataset.file;

            const roundName =
                button.textContent.trim();


            currentFile =
                fileName;

            currentRound =
                roundName;


            // Excelを読み込み直す
            loadExcel(
                currentFile,
                currentRound
            );


            // メニューを閉じる
            closeMenu();


            // 検索欄をクリア
            document
                .getElementById(
                    "searchInput"
                )
                .value = "";

        }
    );

});


// ==============================
// 検索
// ==============================

document
.getElementById("searchBtn")
.addEventListener(
    "click",
    () => {

        const input =
            document.getElementById(
                "searchInput"
            );


        const target =
            Number(
                input.value
            );


        if (isNaN(target)) {
            return;
        }


        // 現在表示している弾の
        // 検索対象セルだけ検索
        document
            .querySelectorAll(
                'td[data-searchable="true"]'
            )
            .forEach(td => {

                if (
                    Number(
                        td.textContent
                    ) === target
                ) {

                    td.classList.add(
                        "search-hit"
                    );

                }

            });

    }
);


// ==============================
// リセット
// ==============================

document
.getElementById("resetBtn")
.addEventListener(
    "click",
    () => {

        document
            .querySelectorAll(
                ".search-hit"
            )
            .forEach(td => {

                td.classList.remove(
                    "search-hit"
                );

            });


        document
            .getElementById(
                "searchInput"
            )
            .value = "";

    }
);