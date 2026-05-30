const RED_NUMBERS = [
    1,
    13,
    25,
    29,
    37,
    42,
    46,
    51
];

fetch("data_VR2.xlsx")
.then(response => response.arrayBuffer())
.then(data => {

    const workbook = XLSX.read(data);

    const container =
        document.getElementById("sheetsContainer");

    workbook.SheetNames.forEach(sheetName => {

        const worksheet =
            workbook.Sheets[sheetName];

        const rows =
            XLSX.utils.sheet_to_json(
                worksheet,
                {header:1}
            );

        const section =
            document.createElement("div");

        section.className = "sheet";

        const title =
            document.createElement("h2");

        title.textContent = sheetName;

        section.appendChild(title);

        const table =
            document.createElement("table");

        rows.forEach((row,rowIndex)=>{

            const tr =
                document.createElement("tr");

            row.forEach((cell,colIndex)=>{

                // table列を非表示
                if(colIndex === 0){
                    return;
                }

                const element =
                    document.createElement(
                        rowIndex===0 ? "th" : "td"
                    );

                element.textContent =
                    cell ?? "";

                // row列
                if(colIndex === 1){
                    element.classList.add(
                        "row-column"
                    );
                }

                // col1,col2のみ赤色判定
                if(
                    rowIndex > 0 &&
                    colIndex >= 2
                ){

                    const value =
                        Number(cell);

                    element.dataset.searchable =
                        "true";

                    if(
                        RED_NUMBERS.includes(value)
                    ){
                        element.classList.add(
                            "red-number"
                        );
                    }
                }

                tr.appendChild(element);

            });

            table.appendChild(tr);

        });

        section.appendChild(table);

        container.appendChild(section);

    });

});

document
.getElementById("searchBtn")
.addEventListener("click",()=>{

    const target =
        Number(
            document
            .getElementById("searchInput")
            .value
        );

    if(isNaN(target)){
        return;
    }

    document
    .querySelectorAll(
        'td[data-searchable="true"]'
    )
    .forEach(td=>{

        if(
            Number(td.textContent)
            === target
        ){
            td.classList.add(
                "search-hit"
            );
        }

    });

});

document
.getElementById("resetBtn")
.addEventListener("click",()=>{

    document
    .querySelectorAll(".search-hit")
    .forEach(td=>{

        td.classList.remove(
            "search-hit"
        );

    });

    document
    .getElementById("searchInput")
    .value = "";

});