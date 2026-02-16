//基本的なOfficeスクリプト（TypeScript）コード一覧よりコードを取得して作成

function main(workbook: ExcelScript.Workbook) {
  // ====== 触るシートを指定（今回は「Before」）======
  const sheet = workbook.getWorksheet("Before");

  // ====== データがどこまであるか（最終行）を探す ======
  // 使われている範囲（UsedRange）がなければ何もしない
  const used = sheet.getUsedRange();
  if (!used) return;

  // X列（件名）が入っている最終行を探して、データの終わりを決める
  // ※X列は見出しで「件名」の列（この列が最後まで入っている前提）
  // getRangeByIndexes(開始行,開始列,取得行数,取得列数)
  //　⇒取得行数を得る”used.getRowCount()”を引数に追加
  //　xValuesでxRangeの範囲(X列)のデータを取得（2次元配列　行×入力データ）
  const xColIndex = 23; // X列=24列目 → 0始まりで23
  const xRange = sheet.getRangeByIndexes(0, xColIndex, used.getRowCount(), 1);
  const xValues = xRange.getValues();

  let lastRow = 1; // Excelの最後の行番号（1始まり）i=0は見出し行なので除外
    // 最後の行からカウントダウンループ
  for (let i = xValues.length - 1; i >= 1; i--) { 
    // i行目の1列目の値を取り出して変数ｖに代入
    const v = xValues[i][0];
    // ｖが空欄でなければ行に値が入っていると判断
    if (v !== null && v !== "") {
      lastRow = i + 1; // iは0始まりなので1始まりへ変換。
      //これで最後に値が入って言う行番号をLastRowにセット
      break;
      //最後に値が入っている行が見つかったのでループを終了
    }
  }

 // ====== 金額が安い案件の資料Ａ資料Ｂ列を黒塗りする ======
  //====== ルール　======
  // Z列（推定金額）が2000000未満なら、
  // E列/F列（資料A/資料B）の「塗りつぶし色」を黒にする 
  for (let r = 2; r <= lastRow; r++) { // 2行目からデータ（1行目は見出し）
    const zRaw = sheet.getRange(`Z${r}`).getValue();

    // セルの値が「数値」でも「文字列」でも比較できるように数値化
    const z = (typeof zRaw === "number") ? zRaw : parseFloat(String(zRaw));

    // 数値に変換できない場合はスキップ（空欄や文字など）
    if (isNaN(z)) continue;

    if (z < 2000000) {
      // getFont().setColor(...) → 文字色
      // getFill().setColor(...) → 塗りつぶし（背景色）
      sheet.getRange(`E${r}`).getFormat().getFill().setColor("black");
      sheet.getRange(`F${r}`).getFormat().getFill().setColor("black");
    }
  }
}
