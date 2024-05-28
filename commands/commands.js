Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

// /**
//  * Shows a notification when the add-in command is executed.
//  * @param event {Office.AddinCommands.Event}
//  */
// function action(event) {
//   const message = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true,
//   };

//   // Show a notification message.
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

//   // Be sure to indicate when the add-in command function is complete.
//   event.completed();
// }


async function unpivotRange(event) {
  await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
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

  event.completed();

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

// Register with Office.js
Office.actions.associate("unpivotRange", unpivotRange);

