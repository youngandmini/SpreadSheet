
const exportButton = document.getElementById("button_export");

const table = document.getElementById("sheet");


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

