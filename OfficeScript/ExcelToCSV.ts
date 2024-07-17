// Excel方眼紙をcsvに変換するOfficeスクリプトです。想定するExcelファイルはSample.xlsxをご確認ください。
function main(workbook: ExcelScript.Workbook) {
    // アクティブなワークシートを取得します
    const sheet = workbook.getActiveWorksheet();

    // シートの範囲を取得します
    const range = sheet.getUsedRange();
    const values = range.getValues();

    // "No"という値が初めて現れる行のインデックスを特定する
    let startRowIndex = -1;
    for (let i = 0; i < values.length; i++) {
        if (values[i].includes("No")) {
            startRowIndex = i;
            break;
        }
    }

    if (startRowIndex === -1) {
        console.log("Noが見つかりませんでした。");
        return;
    }

    // CSVとして保存する文字列を作成します
    let csvContent = "";

    // "No"が出現する行から始まる行を走査し、コンマで区切られた文字列に変換します
    for (let i = startRowIndex; i < values.length; i++) {
        // 空白や無効なエントリをフィルタリング
        const processedRow = values[i].filter(cell => cell != null && cell !== '');

        // コンマの連続を取り除く
        let row = processedRow.join(",").replace(/,{2,}/g, ','); // 連続するコンマを1つに置換

        // 行がカンマだけでない場合のみ追加
        if (row.replace(/,/g, '').trim() !== '') {
            csvContent += row + "\r\n"; // 改行コードを追加
        }
    }

    // 結果をログに出力します。
    console.log(csvContent);

}
