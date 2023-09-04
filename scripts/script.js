
const exportButton = document.getElementById("button_export");
const table = document.getElementById("sheet");
const currentCellPosition = document.getElementById("current_cell_position");



exportButton.addEventListener('click', () => {

    const table = document.getElementById("sheet");

    // 새로운 테이블 생성
    const newTable = document.createElement('table');

    // rowIndex와 colIndex로부터 input 데이터 추출하여 새로운 테이블에 추가
    for (let rowIndex = 1; rowIndex < table.rows.length; rowIndex++) {
        const newRow = document.createElement('tr');

        for (let colIndex = 1; colIndex < table.rows[rowIndex].cells.length; colIndex++) {
            const newCell = document.createElement('td');
            newCell.textContent = table.rows[rowIndex].cells[colIndex].firstElementChild.value;
            newRow.appendChild(newCell);
        }

        newTable.appendChild(newRow);
    }

    /* table로 worksheet 생성 */
    const worksheet = XLSX.utils.table_to_book(newTable, {sheet: 'sheet-1'});
    /*
    xlsx파일로 다운로드하지 않고 xls파일로 다운로드(구글시트 호환성 문제)
    */
    XLSX.writeFile(worksheet, 'MyTable.xls');
});


for (let rowIndex = 1; rowIndex < table.rows.length; rowIndex++) {
    for (let colIndex = 1; colIndex < table.rows[rowIndex].cells.length; colIndex++) {

        const cellInput = table.rows[rowIndex].cells[colIndex].firstChild;
        cellInput.paramRow = rowIndex;
        cellInput.paramCol = colIndex;
        cellInput.addEventListener("focus", focusOnCell);
        cellInput.addEventListener("blur", blurOutCell);
    }
}

// 내부 cell에 focus 됐을 때 사용
function focusOnCell(event) {
    // console.log("cursor on");

    const rowIndex = event.currentTarget.paramRow;
    const colIndex = event.currentTarget.paramCol;

    const colTableHeader = table.rows[0].cells[colIndex];
    const rowTableHeader = table.rows[rowIndex].cells[0];

    currentCellPosition.textContent = colTableHeader.textContent + rowTableHeader.textContent;

    colTableHeader.style.backgroundColor = "#AED8E6";
    colTableHeader.style.color = "white";

    rowTableHeader.style.backgroundColor = "#AED8E6";
    rowTableHeader.style.color = "white";
}

// 내부 cell에서 (blur)focusOut 됐을 때 사용
function blurOutCell(event) {
    // console.log("cursor off");

    const rowIndex = event.currentTarget.paramRow;
    const colIndex = event.currentTarget.paramCol;

    const colTableHeader = table.rows[0].cells[colIndex];
    const rowTableHeader = table.rows[rowIndex].cells[0];

    currentCellPosition.textContent = "";

    colTableHeader.style.backgroundColor = "#DEDEDE";
    colTableHeader.style.color = "black";

    rowTableHeader.style.backgroundColor = "#DEDEDE";
    rowTableHeader.style.color = "black";
}