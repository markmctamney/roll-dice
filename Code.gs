// Load Lodash library for object compare
const _ = LodashGS.load();

const DEBUG = true;

const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
const activeSheet = activeSpreadsheet.getActiveSheet()
const sheetName = activeSheet.getSheetName();
const lastRow = activeSheet.getLastRow();
console.log(`Running script in active sheet ${sheetName}...`)

const DATA_START_ROW_NUM = 6; // Row number where data (not just headers start)
const RAND_COL_NUM = 62; // Column number of generated rand() formula (hidden column BJ);

const COURT_ROW_NUMS = {
  // [courtNumber]: rowNumber
  1: 6, 
  2: 14, 
  3: 22, 
  4: 30, 
  5: 38,
}; // Hard-coded row numbers for each Court

const OUTPUT_RUNTIME_ROW_NUM = 3; // Column number to output most recent script runtime for debugging

const MAX_DESIRED_WIN_PROBABILITY = 0.6;
const MIN_DESIRED_WIN_PROBABILITY = 0.4;

// Count the number of non-empty rows we need to generate random numbers for
const RAND_OUTPUT_NUM_ROWS = activeSheet.getRange(DATA_START_ROW_NUM, RAND_COL_NUM - 1, lastRow).getValues().filter(([value], index, arr) => value ? value !== '' : false).length;

/**
 * Make this an indexed object so we can dynamically get game index without hard-coding, e.g. `gameIndexes[gameNumber]` where gameNumber is an integer 1-8
 */ 
const OUTPUT_PLAYER_GAME_RAND_COL_NUMS = {
  // [gameNumber: number]: number, // number value is column number
  1: RAND_COL_NUM + 2, // game1,
  2: RAND_COL_NUM + 4, // game2,
  3: RAND_COL_NUM + 6, // game3,
  4: RAND_COL_NUM + 8, // game4,
  5: RAND_COL_NUM + 10, // game5,
  6: RAND_COL_NUM + 12, // game6,
  7: RAND_COL_NUM + 14, // game7,
  8: RAND_COL_NUM + 17, // game8 (has extra column in front of it so odd, not even)
}

const GAME_DATA_COL_NUMS = {
  // [gameNumber]: dataColNum,
  1: 6, // var g1column = 6
  2: 11, // var g2column = 11
  3: 16,
  4: 21,
  5: 26,
  6: 31,
  7: 36,
  8: 41,
}

/**
 * @param {number & keyof typeof OUTPUT_PLAYER_GAME_RAND_COL_NUMS} gameNumber
 */
function getRandOutputColNum(gameNumber) {
  return OUTPUT_PLAYER_GAME_RAND_COL_NUMS[gameNumber];
}

/**
 * Helper function to wrap other functions and measure elapsed time, handle errors, etc.
 * 
 * @param {keyof this & string} functionName - Name of function to run with debugging.
 */ 
function withDebugging(functionName) {
  /**
  * @param {any[]} args
  */
  function withDebugWrapper(...args) {
    const startTimeMS = Date.now();
    const func = functionName in this && this[functionName];

    // Make sure valid function name was passed as parameter
    if (typeof func !== 'function') throw new Error('withDebugging can only be called with a function');

    if (DEBUG) console.log(`Running function ${functionName} at ${startTimeMS}`);

    try {
      func(...args);
    } catch (err) {
      console.error(`Error running ${functionName}: ${err}`)
    }

    const endTimeMS = Date.now();
    const elapsedTime = (endTimeMS - startTimeMS) / 1000;

    if (DEBUG) console.log(`Finished running function ${functionName} at ${endTimeMS}. Total execution time: ${elapsedTime}s`);
  }

  return withDebugWrapper;
}

/**
 * @param {number & keyof typeof OUTPUT_PLAYER_GAME_RAND_COL_NUMS} gameNumber
 */
function rollDiceOptimize(gameNumber) {
  const dataColNum = GAME_DATA_COL_NUMS[gameNumber];
  const winColNum = dataColNum - 1;
  const runtimeColNum = dataColNum - 2;
  const optimizeStartMS = Date.now();

  const getWinCount = () => activeSheet.getRange(DATA_START_ROW_NUM, winColNum, RAND_OUTPUT_NUM_ROWS).getValues().filter(([value], index, arr) => value ? value !== 1 && value !== 0 : false).length;
  const getWinPsByCourt = (courtNumber) => activeSheet.getRange(COURT_ROW_NUMS[courtNumber], [dataColNum]).getValue();
  const getCourtsWinPs = () => ({
    1: getWinPsByCourt(1),
    2: getWinPsByCourt(2),
    3: getWinPsByCourt(3),
    4: getWinPsByCourt(4),
    5: getWinPsByCourt(5),
  });

  const winCount = getWinCount();
  if (winCount > 0) {
    Browser.msgBox(`Cannot run rollDiceOptimize for Game ${gameNumber.toString()} in active sheet ${sheetName}. Some of the games appear to be completed. The script will only write data for games that have no data in the "Win" column to prevent accidentally overwriting data for completed and finalized games. If you wish to proceed, remove all 1's and 0's from that column and try again.`);
    return;
  }

  let courtsWinPs = getCourtsWinPs();
  let updatedCourtsWinPs = {}; // Second object to be used to compare and make sure data is changing so we're not winding up in an infinite loop
  let loopCounter = 0;

  /**
   * Helper function to avoid copying and pasting duplicate logic
   * 
   * @param {number & keyof typeof COURT_ROW_NUMS} courtNumber Court number (e.g. 1-5)
   * @returns boolean
   */
  const isOutsideDesiredWinProbability = (courtNumber) => {
    const winPsByCourt = getWinPsByCourt(courtNumber); 
    const winProbComparison = winPsByCourt > MAX_DESIRED_WIN_PROBABILITY || winPsByCourt < MIN_DESIRED_WIN_PROBABILITY;
    return winProbComparison;
  }

  const isDeepEqual = (obj1, obj2) => _.eq(obj1, obj2); // JSON.stringify(obj1) === JSON.stringify(obj2);

  // console.log({isDeepEqual: isDeepEqual(courtsWinPs, updatedCourtsWinPs), court1: isOutsideDesiredWinProbality(1), court2: isOutsideDesiredWinProbality(2), courtsWinPs, updatedCourtsWinPs})
  //Check if win probabilities are a) updating and b) still outside of bounds
  while (!isDeepEqual(courtsWinPs, updatedCourtsWinPs) && 
    (isOutsideDesiredWinProbability(1) || isOutsideDesiredWinProbability(2) || isOutsideDesiredWinProbability(3) || isOutsideDesiredWinProbability(4) || isOutsideDesiredWinProbability(5))
  ) {
    loopCounter += 1;
    
    // Check for Win column data
    // If any 1's or 0's, assume these games already happened and don't overwrite anything
    rollDice(gameNumber);

    // This may be contributing to doc getting locked up, so using flush() instead of sleep
    // @link https://stackoverflow.com/a/43444080
    // SpreadsheetApp.flush();

    //Check the Win percentage on each court
    updatedCourtsWinPs = getCourtsWinPs();
    // console.log('compare', {courtsWinPs, updatedCourtsWinPs});
  }  

  const elapsed = (Date.now() - optimizeStartMS) / 1000;
  const avgIterationDuration = loopCounter > 0 ? (elapsed / loopCounter) : undefined;
  console.log(`Team optimization complete for Game ${gameNumber} after ${elapsed.toFixed(0)}s. Tried ${loopCounter} different variations${avgIterationDuration ? ` with an average duration per iteration of ${avgIterationDuration.toFixed(1)}s` : ''}.`);

  // Write elapsed execution time to designated cell in Row 3 of the sheet
  activeSheet.getRange(OUTPUT_RUNTIME_ROW_NUM,runtimeColNum).setValue(elapsed);
}

const rollDice1Optimize = () => withDebugging('rollDiceOptimize')(1);
const rollDice2Optimize = () => withDebugging('rollDiceOptimize')(2);
const rollDice3Optimize = () => withDebugging('rollDiceOptimize')(3);
const rollDice4Optimize = () => withDebugging('rollDiceOptimize')(4);
const rollDice5Optimize = () => withDebugging('rollDiceOptimize')(5);
const rollDice6Optimize = () => withDebugging('rollDiceOptimize')(6);
const rollDice7Optimize = () => withDebugging('rollDiceOptimize')(7);
const rollDice8Optimize = () => withDebugging('rollDiceOptimize')(8);

/**
 * Generate random numbers for each player for a particular game, and
 * write the values to the hidden column for the appropriate game (BL, BN, BP, ...)
 * 
 * @param {number & keyof typeof OUTPUT_PLAYER_GAME_RAND_COL_NUMS} gameNumber
 * @todo Add overwrite flag and skip ahead if data already exists
 */
function rollDice(gameNumber){
  // Calculate random value in JS code instead of spreadsheet
  const targetRange = activeSheet.getRange(
    DATA_START_ROW_NUM,                 // Range starting row number
    getRandOutputColNum(gameNumber),    // Range starting column number
    RAND_OUTPUT_NUM_ROWS,               // Number of rows to include in range
  ); // Range to write random values to, e.g.= BL6:BL51
  // Auto-populate an array of length ${RAND_OUTPUT_NUM_ROWS}, each populated with a unique random number generated (between 0 and 1)
  const randValues = Array.apply(null, Array(RAND_OUTPUT_NUM_ROWS)).map(() => [Math.random()]);
  targetRange.setValues(randValues);
}

// Functions linked to individual dice buttons in doc
const rollDice1 = () => withDebugging('rollDice')(1);
const rollDice2 = () => withDebugging('rollDice')(2);
const rollDice3 = () => withDebugging('rollDice')(3);
const rollDice4 = () => withDebugging('rollDice')(4);
const rollDice5 = () => withDebugging('rollDice')(5);
const rollDice6 = () => withDebugging('rollDice')(6);
const rollDice7 = () => withDebugging('rollDice')(7);
const rollDice8 = () => withDebugging('rollDice')(8);
