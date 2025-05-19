/**
 * Determines the lottery order
 * Sort school names sheet A1:D9 by column A in ascending order and assign random lottery order (1-8) to column C
 */
function determineLotteryOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws1 = ss.getSheetByName("校名");
  const ws2 = ss.getSheetByName("抽選");
  const ws3 = ss.getSheetByName("テーブル");
  
  // Initialize array
  const T = Array(9).fill(0); // 0-8 index array with 9 elements
  
  // Sort (starting from row 2, excluding header row)
  ws1.getRange("A2:D9").sort({column: 1, ascending: true});
  
  // Determine lottery order
  for (let L = 1; L <= 8; L++) {
    while (true) {
      const I = Math.floor((8 - 1 + 1) * Math.random() + 1); // Random number 1-8
      if (T[I] === 0) {
        ws1.getRange(L + 1, 3).setValue(I);
        T[I] = 1;
        break;
      }
    }
  }
}

/**
 * Automatic lottery for team competition
 * Place 1st to 4th ranked schools and 5th or lower ranked schools from each prefecture by lottery
 */
function autoLotteryForTeamCompetition() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws1 = ss.getSheetByName("校名");
  const ws2 = ss.getSheetByName("抽選");
  const ws3 = ss.getSheetByName("テーブル");
  
  // Initialize arrays
  const TokenMei = Array(9).fill("");
  const TokenNumber = Array(9).fill(0);
  const SankaNumber = Array(9).fill(0);
  const Token1_Flag = Array(9).fill(0);
  const Token2_Flag = Array(9).fill(0);
  const Token3_Flag = Array(9).fill(0);
  const Token4_Flag = Array(9).fill(0);
  const Token4_Flag2 = Array(9).fill(0);
  const Token5_Flag = Array(17).fill(0);
  const SideNumber = Array(9).fill(0);
  const ZoneNumber = Array(9).fill(0);
  
  // Sort (starting from row 2, excluding header row)
  ws1.getRange("A2:D9").sort([
    {column: 4, ascending: false},
    {column: 3, ascending: true}
  ]);
  
  // Get lottery order
  for (let L = 1; L <= 8; L++) {
    TokenMei[ws1.getRange(L + 1, 3).getValue()] = ws1.getRange(L + 1, 2).getValue();
    TokenNumber[L] = ws1.getRange(L + 1, 3).getValue();  // Lottery order for 5th place and below
    SankaNumber[L] = ws1.getRange(L + 1, 4).getValue();  // Lottery order for 5th place and below
  }
  
  // Clear table sheet
  ws3.getRange("A1:C16").clearContent();
  ws3.getRange("A21:C36").clearContent();
  ws3.getRange("A41:C56").clearContent();
  
  // Lottery for 1st place schools from each prefecture
  for (let L = 1; L <= 8; L++) {
    while (true) {
      const I = Math.floor((8 - 1 + 1) * Math.random() + 1);
      if (Token1_Flag[I] === 0) {
        switch (I) {
          case 1:
            ws3.getRange(1, 1).setValue(TokenMei[L] + "1");
            ws3.getRange(21, 1).setValue(L);
            break;
          case 2:
            ws3.getRange(4, 1).setValue(TokenMei[L] + "1");
            ws3.getRange(24, 1).setValue(L);
            break;
          case 3:
            ws3.getRange(5, 1).setValue(TokenMei[L] + "1");
            ws3.getRange(25, 1).setValue(L);
            break;
          case 4:
            ws3.getRange(8, 1).setValue(TokenMei[L] + "1");
            ws3.getRange(28, 1).setValue(L);
            break;
          case 5:
            ws3.getRange(9, 1).setValue(TokenMei[L] + "1");
            ws3.getRange(29, 1).setValue(L);
            break;
          case 6:
            ws3.getRange(12, 1).setValue(TokenMei[L] + "1");
            ws3.getRange(32, 1).setValue(L);
            break;
          case 7:
            ws3.getRange(13, 1).setValue(TokenMei[L] + "1");
            ws3.getRange(33, 1).setValue(L);
            break;
          case 8:
            ws3.getRange(16, 1).setValue(TokenMei[L] + "1");
            ws3.getRange(36, 1).setValue(L);
            break;
        }
        Token1_Flag[I] = 1;
        break;
      }
    }
  }
  
  // Lottery for 2nd place schools from each prefecture
  SideNumber[ws3.getRange(21, 1).getValue()] = 8;
  SideNumber[ws3.getRange(24, 1).getValue()] = 8;
  SideNumber[ws3.getRange(25, 1).getValue()] = 8;
  SideNumber[ws3.getRange(28, 1).getValue()] = 8;
  SideNumber[ws3.getRange(29, 1).getValue()] = 4;
  SideNumber[ws3.getRange(32, 1).getValue()] = 4;
  SideNumber[ws3.getRange(33, 1).getValue()] = 4;
  SideNumber[ws3.getRange(36, 1).getValue()] = 4;
  
  for (let L = 1; L <= 8; L++) {
    while (true) {
      const I = Math.floor(4 * Math.random() + SideNumber[L] - 3);
      if (Token2_Flag[I] === 0) {
        switch (I) {
          case 1:
            ws3.getRange(2, 1).setValue(TokenMei[L] + "2");
            ws3.getRange(22, 1).setValue(L);
            break;
          case 2:
            ws3.getRange(3, 1).setValue(TokenMei[L] + "2");
            ws3.getRange(23, 1).setValue(L);
            break;
          case 3:
            ws3.getRange(6, 1).setValue(TokenMei[L] + "2");
            ws3.getRange(26, 1).setValue(L);
            break;
          case 4:
            ws3.getRange(7, 1).setValue(TokenMei[L] + "2");
            ws3.getRange(27, 1).setValue(L);
            break;
          case 5:
            ws3.getRange(10, 1).setValue(TokenMei[L] + "2");
            ws3.getRange(30, 1).setValue(L);
            break;
          case 6:
            ws3.getRange(11, 1).setValue(TokenMei[L] + "2");
            ws3.getRange(31, 1).setValue(L);
            break;
          case 7:
            ws3.getRange(14, 1).setValue(TokenMei[L] + "2");
            ws3.getRange(34, 1).setValue(L);
            break;
          case 8:
            ws3.getRange(15, 1).setValue(TokenMei[L] + "2");
            ws3.getRange(35, 1).setValue(L);
            break;
        }
        Token2_Flag[I] = 1;
        break;
      }
    }
  }
  
  // Lottery for 3rd place schools from each prefecture
  ZoneNumber[ws3.getRange(22, 1).getValue()] = 4;
  ZoneNumber[ws3.getRange(23, 1).getValue()] = 4;
  ZoneNumber[ws3.getRange(26, 1).getValue()] = 2;
  ZoneNumber[ws3.getRange(27, 1).getValue()] = 2;
  ZoneNumber[ws3.getRange(30, 1).getValue()] = 8;
  ZoneNumber[ws3.getRange(31, 1).getValue()] = 8;
  ZoneNumber[ws3.getRange(34, 1).getValue()] = 6;
  ZoneNumber[ws3.getRange(35, 1).getValue()] = 6;
  
  for (let L = 1; L <= 8; L++) {
    while (true) {
      const I = Math.floor(2 * Math.random() + ZoneNumber[L] - 1);
      if (Token3_Flag[I] === 0) {
        switch (I) {
          case 1:
            ws3.getRange(2, 2).setValue(TokenMei[L] + "3");
            ws3.getRange(22, 2).setValue(L);
            break;
          case 2:
            ws3.getRange(3, 2).setValue(TokenMei[L] + "3");
            ws3.getRange(23, 2).setValue(L);
            break;
          case 3:
            ws3.getRange(6, 2).setValue(TokenMei[L] + "3");
            ws3.getRange(26, 2).setValue(L);
            break;
          case 4:
            ws3.getRange(7, 2).setValue(TokenMei[L] + "3");
            ws3.getRange(27, 2).setValue(L);
            break;
          case 5:
            ws3.getRange(10, 2).setValue(TokenMei[L] + "3");
            ws3.getRange(30, 2).setValue(L);
            break;
          case 6:
            ws3.getRange(11, 2).setValue(TokenMei[L] + "3");
            ws3.getRange(31, 2).setValue(L);
            break;
          case 7:
            ws3.getRange(14, 2).setValue(TokenMei[L] + "3");
            ws3.getRange(34, 2).setValue(L);
            break;
          case 8:
            ws3.getRange(15, 2).setValue(TokenMei[L] + "3");
            ws3.getRange(35, 2).setValue(L);
            break;
        }
        Token3_Flag[I] = 1;
        break;
      }
    }
  }
  
  // Lottery for 4th place schools from each prefecture
  ZoneNumber[ws3.getRange(21, 1).getValue()] = 4;
  ZoneNumber[ws3.getRange(24, 1).getValue()] = 4;
  ZoneNumber[ws3.getRange(25, 1).getValue()] = 2;
  ZoneNumber[ws3.getRange(28, 1).getValue()] = 2;
  ZoneNumber[ws3.getRange(29, 1).getValue()] = 8;
  ZoneNumber[ws3.getRange(32, 1).getValue()] = 8;
  ZoneNumber[ws3.getRange(33, 1).getValue()] = 6;
  ZoneNumber[ws3.getRange(36, 1).getValue()] = 6;
  
  for (let L = 1; L <= 8; L++) {
    while (true) {
      const I = Math.floor(2 * Math.random() + ZoneNumber[L] - 1);
      if (Token4_Flag[I] === 0) {
        switch (I) {
          case 1:
            ws3.getRange(1, 2).setValue(TokenMei[L] + "4");
            ws3.getRange(21, 2).setValue(L);
            break;
          case 2:
            ws3.getRange(4, 2).setValue(TokenMei[L] + "4");
            ws3.getRange(24, 2).setValue(L);
            break;
          case 3:
            ws3.getRange(5, 2).setValue(TokenMei[L] + "4");
            ws3.getRange(25, 2).setValue(L);
            break;
          case 4:
            ws3.getRange(8, 2).setValue(TokenMei[L] + "4");
            ws3.getRange(28, 2).setValue(L);
            break;
          case 5:
            ws3.getRange(9, 2).setValue(TokenMei[L] + "4");
            ws3.getRange(29, 2).setValue(L);
            break;
          case 6:
            ws3.getRange(12, 2).setValue(TokenMei[L] + "4");
            ws3.getRange(32, 2).setValue(L);
            break;
          case 7:
            ws3.getRange(13, 2).setValue(TokenMei[L] + "4");
            ws3.getRange(33, 2).setValue(L);
            break;
          case 8:
            ws3.getRange(16, 2).setValue(TokenMei[L] + "4");
            ws3.getRange(36, 2).setValue(L);
            break;
        }
        Token4_Flag[I] = 1;
        break;
      }
    }
  }
  
  // Lottery for 5th place and below schools
  ws2.getRange(1, 26).setValue("");
  let A = 0;
  do {
    lotteryForFifthPlace();
    A = ws2.getRange(1, 26).getValue();
  } while (A !== 16);
  
  // Determine match order
  determineMatchOrder();
}

/**
 * Sort school names
 * Sort school names sheet A1:D9 by column D in descending order and column C in ascending order
 */
function sortSchoolNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("校名");
  
  // Sort (column D=4 descending, column C=3 ascending, starting from row 2, excluding header row)
  ws.getRange("A2:D9").sort([
    {column: 4, ascending: false},
    {column: 3, ascending: true}
  ]);
}

/**
 * Lottery for 5th place and below
 * Place 5th place and below schools by lottery
 */
function lotteryForFifthPlace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws1 = ss.getSheetByName("校名");
  const ws2 = ss.getSheetByName("抽選");
  const ws3 = ss.getSheetByName("テーブル");
  
  // Initialize arrays
  const TokenMei = Array(9).fill("");
  const TokenNumber = Array(9).fill(0);
  const SankaNumber = Array(9).fill(0);
  const Token5_Flag = Array(17).fill(0);
  const SeigenNum = Array(11).fill(0);
  
  // Sort (starting from row 2, excluding header row)
  ws1.getRange("A2:D9").sort([
    {column: 4, ascending: false},
    {column: 3, ascending: true}
  ]);
  
  // Get lottery order
  let L = 0;
  do {
    L = L + 1;
    TokenMei[ws1.getRange(L + 1, 3).getValue()] = ws1.getRange(L + 1, 2).getValue();
    TokenNumber[L] = ws1.getRange(L + 1, 3).getValue();  // Lottery order for 5th place and below
    SankaNumber[L] = ws1.getRange(L + 1, 4).getValue();  // Lottery order for 5th place and below
  } while (L < 8);
  
  // Clear table sheet
  ws3.getRange("C1:C16").clearContent();
  ws3.getRange("C21:C36").clearContent();
  
  // Restrictions based on number of participating schools
  SeigenNum[5] = 2;  // 5 schools
  SeigenNum[6] = 2;
  SeigenNum[7] = 3;
  SeigenNum[8] = 3;
  SeigenNum[9] = 4;
  SeigenNum[10] = 4; // 10 schools
  
  let Cnt = 0;
  L = 0;
  let A = 0;
  
  do {
    L = L + 1;
    // Start processing if there are 5 or more participating schools
    if (SankaNumber[L] > 4) {
      for (let N = 5; N <= SankaNumber[L]; N++) {
        do {
          Cnt = Cnt + 1;
          if (Cnt === 1600) break;
          
          // Decide which of the 16 blocks to place in using random numbers
          const I = Math.floor(16 * Math.random() + 1);
          
          if (Token5_Flag[I] === 0) {
            // Check if there are no schools from the same prefecture in the same a-p group
            if (ws3.getRange(50 + TokenNumber[L], 6 + I).getValue() === 0) {
              // Condition: up to 3 schools in the same zone
              if (ws3.getRange(30 + TokenNumber[L], 6 + Math.floor((I - 1) / 4) + 1).getValue() < 3) {
                if (ws3.getRange(20 + TokenNumber[L], 6 + Math.floor((I - 1) / 8) + 1).getValue() <= SeigenNum[SankaNumber[L]]) {
                  // VBA's StrConv function for full-width to half-width conversion is omitted as GAS doesn't have a direct equivalent
                  ws3.getRange(I, 3).setValue(TokenMei[TokenNumber[L]] + N);
                  ws3.getRange(20 + I, 3).setValue(TokenNumber[L]);
                  Token5_Flag[I] = 1;
                  A = A + 1;
                  break;
                }
              }
            }
          }
        } while (true);
      }
    }
  } while (L < 8);
  
  ws2.getRange(1, 26).setValue(A);
}

/**
 * Determine match order
 * Determine the order of matches
 */
function determineMatchOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws3 = ss.getSheetByName("テーブル");
  
  // Clear table sheet
  ws3.getRange("A41:C56").clearContent();
  
  // Determine match order
  for (let L = 1; L <= 16; L++) {
    const I = Math.floor(3 * Math.random() + 1);
    ws3.getRange(40 + L, 1).setValue(I);
  }
  
  for (let L = 1; L <= 16; L++) {
    while (true) {
      const I = Math.floor(3 * Math.random() + 1);
      if (ws3.getRange(40 + L, 1).getValue() !== I) {
        ws3.getRange(40 + L, 2).setValue(I);
        break;
      }
    }
  }
  
  for (let L = 1; L <= 16; L++) {
    while (true) {
      const I = Math.floor(3 * Math.random() + 1);
      if (ws3.getRange(40 + L, 1).getValue() !== I && ws3.getRange(40 + L, 2).getValue() !== I) {
        ws3.getRange(40 + L, 3).setValue(I);
        break;
      }
    }
  }
}

/**
 * Creates a custom menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('抽選ツール')
      .addItem('初期化', 'initializeSheets')
      .addSeparator()
      .addItem('抽選順決定', 'determineLotteryOrder')
      .addItem('団体戦自動抽選', 'autoLotteryForTeamCompetition')
      // .addItem('並べ替え', 'sortSchoolNames')
      // .addItem('抽選5位', 'lotteryForFifthPlace')
      // .addItem('試合順決定', 'determineMatchOrder')
      .addToUi();
}

/**
 * Initialize sheets before lottery
 * Clear lottery order, table sheet, and reset data to initial state
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws1 = ss.getSheetByName("校名");
  const ws2 = ss.getSheetByName("抽選");
  const ws3 = ss.getSheetByName("テーブル");
  
  // Clear lottery order in school names sheet (column C)
  for (let i = 2; i <= 9; i++) {
    ws1.getRange(i, 3).clearContent();
  }
  
  // Copy values from I1:K16 to A1:C16
  const sourceRange = ws3.getRange(1, 9, 16, 3); // I1:K16
  const targetRange = ws3.getRange(1, 1, 16, 3); // A1:C16
  sourceRange.copyTo(targetRange);
  
  // Clear A41:C56
  ws3.getRange(41, 1, 16, 3).clearContent(); // A41:C56
  
  // Clear lottery result cell
  ws2.getRange(1, 26).clearContent();
  
  // Sort school names by column A (default order), excluding header row
  ws1.getRange(2, 1, 8, 4).sort({column: 1, ascending: true});
  
  // Show confirmation message
  SpreadsheetApp.getUi().alert('初期化が完了しました。抽選を開始できます。');
}
