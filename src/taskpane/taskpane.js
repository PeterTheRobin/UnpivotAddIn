Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('button').onclick = function () {
      unpivotRange();
    }
  }
});


async function unpivotRange() {
  await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    range.load("address");
    range.load("values");
    range.load("columnCount");
    range.load("rowCount");

    let sheets = context.workbook.worksheets;
    let unpivot_sheet = sheets.add("unpivot");

    await context.sync();

    let unpivot_column_count = 3;
    let unpivot_row_count = (range.rowCount - 1) * (range.columnCount - 1) + 1;
    let unpivot_column = getColumnLetter(unpivot_column_count - 1);
    let unpivot_range = unpivot_sheet.getRange(`A1:${unpivot_column}${unpivot_row_count}`);

    var unpivot_header = ['type', 'attribute', 'value'];

    var attributes = range.values[0].slice(1, range.values[0].length);
    var types = [];
    var values = [];

    for (let i = 1; i < range.values.length; i++) {
      types.push(range.values[i][0]);
      values.push(range.values[i].slice(1, range.values[i].length));
    };

    var unpivot_data = [];

    for (let i = 0; i < types.length; i++){
      for (let j = 0; j < attributes.length; j++){
        unpivot_data.push([types[i], attributes[j], values[i][j]])
      };
    };

    unpivot_data.unshift(unpivot_header);

    unpivot_range.values = unpivot_data;
  });
};



// UTILITY FUNCTIONS

function getColumnLetter(i) {
  const m = i % 26;
  const c = String.fromCharCode(65 + m);
  const r = i - m;
  return r > 0
    ? `${getColumnLetter((r - 1) / 26)}${c}`
    : `${c}`
}

