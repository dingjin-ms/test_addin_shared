/* global clearInterval, console, setInterval */

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
export function add(first, second) {
  return first + second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
export function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
export function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
export function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000 * 10);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
export function logMessage(message) {
  console.log(message);

  return message;
}

// Write my own functions starting here

/**
 * Return random int - 10
 * @customfunction
 * @returns {number} Return random int - 10.
 */
function returnInt() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnInt.');
  return Math.floor(Math.random() * 10);
}

/**
 * Return random int - 100
 * @customfunction
 * @returns {number} Return random int - 100.
 */
function returnIntPromise() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnIntPromise.');
  return new Promise(function (resolve) {
      setTimeout(function () {
          resolve(Math.floor(Math.random() * 100));
      }, 1000);
  });
}

/**
 * Return 42
 * @customfunction
 * @returns {number} Return 42.
 */
function return42Promise() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call return42Promise.');
  return new Promise(function (resolve) {
      setTimeout(function () {
          resolve(42);
      }, 1000);
  });
}

/**
 * Return random int - 1000
 * @customfunction
 * @volatile
 * @returns {number} Return random int - 1000.
 */
function returnIntVolatile() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnIntVolatile.');
  return Math.floor(Math.random() * 1000)
}

/**
 * Return 42
 * @customfunction
 * @volatile
 * @returns {number} Return 42.
 */
function return42Volatile() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call return42Volatile.');
  return 42;
}

/**
 * Return current time
 * @customfunction
 * @volatile
 * @returns {string} Return current time.
 */
function returnStringVolatile() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringVolatile.');
  return currentTime();
}

/**
 * Return current time
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
export function returnStringStream5m(invocation) {
  let result = currentTime();
  invocation.setResult(result);
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringStream5m#1.');
  const timer = setInterval(() => {
    //var timestamp = new Date().toISOString();
    result = currentTime();
    invocation.setResult(result);
    console.log(`[${result}] `, 'Call returnStringStream5m#2.');
  }, 1000 * 60 * 5);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Return current time
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
export function returnStringStream1s(invocation) {
  let result = currentTime();
  invocation.setResult(result);
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringStream1s#1.');
  const timer = setInterval(() => {
    //var timestamp = new Date().toISOString();
    result = currentTime();
    invocation.setResult(result);
    console.log(`[${result}] `, 'Call returnStringStream1s#2.');
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Return current time
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
export function returnTestStringStream1s(invocation) {
  invocation.setResult(123);
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnTestStringStream1s#1.');
  const timer = setInterval(() => {
    var timestamp = new Date().toISOString();
    invocation.setResult(123);
    console.log(`[${timestamp}] `, 'Call returnTestStringStream1s#2.');
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Return current time
 * @customfunction
 * @returns {string[][]} Return current time.
 */
export function returnStringDynamicArray() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringDynamicArray.');
  return [[currentTime(), currentTime(), currentTime()]];
}

/**
 * Return test strings
 * @customfunction
 * @returns {string[][]} Return test strings.
 */
export function returnStringTestDynamicArray() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringTestDynamicArray.');
  return [["1", "2", "3"]];
}

/**
 * Return nested range
 * @customfunction
 * @param {string[][]} values Multiple ranges of values.
 * @returns {string} Return nested range.
 */
export function returnStringNested(range) {
  var cell = range[0];
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringNested.');
  return cell;
}

/**
 * Wait before returning current time
 * @customfunction
 * @returns {string} Return current time.
 */
export function returnStringWait() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}] `, 'Call returnStringWait.');
  var num = 0;
  for (var i = 0; i < 100000000000; i++) {
    num++;
  }
  return timestamp;
}

/**
 * Wait before returning current time
 * @customfunction
 * @volatile
 * @returns {string} Return current time.
 */
export function returnStringWaitVolatile() {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}]`, 'Call returnStringWaitVolatile.');
  var num = 0;
  for (var i = 0; i < 100000000000; i++) {
    num++;
  }
  return timestamp;
}

/**
 * Take a number as the input value and return a double as the output.
 * @customfunction
 * @returns A formatted number value.
 */
function returnDoubleCellValue() {
    return {
        type: "Double",
        basicValue: 10,
        numberFormat: "0.00%"
    }
}

/**
 * Input a string to output
 * @customfunction
 * @returns {string} Return input.
 */
export function inputString(str) {
  var timestamp = new Date().toISOString();
  console.log(`[${timestamp}]`, 'Call inputString.');
  return str;
}

/**
 * @customfunction
 * @param {number} first
 * @param {number} second
 * @param {number} [third]
 * @returns {number}
 */
function inputIntOptional(first, second, third) {
  if (third === null) third = 0;
  return first + second + third;
}

/**
 * @customfunction
 * @param {number[]} values
 * @returns {number}
 */
function inputIntRepeating(values) {
  const sum = values.reduce((a, b) => a + b, 0);
  return sum / values.length;
}

/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function inputRangeParams(values) {
  let highest = Number.MIN_SAFE_INTEGER,
    secondHighest = Number.MIN_SAFE_INTEGER;
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  if (secondHighest === Number.MIN_SAFE_INTEGER) {
    secondHighest = null; // No second highest found
  }
  return secondHighest;
}

/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function intputInvocation(first, second, invocation) {
  const address = invocation.address;
  return address;
}

/**
 * This function will call the write API to write "Hello" to A1.
 * @customfunction
 * @returns {string} 
 */
async function callWriteApi() {
  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("A1");
    range.values = [["Hello"]];
    await context.sync();
  });

  return "test";
}

/** 
 * This function will call the read API to read the value from A1.
 * @customfunction
 * @returns {string} 
 */
async function callReadApi() {
  var result = "Initial value";
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("A1");
    range.load("values");
    await context.sync();
    console.log(range.values[0][0]);
    result = range.values[0][0] || "No value found";
  });

  return result;
}

/**
 * This function will call the read API to read the value from A1 10 times using context.sync() 10 times.
 * @customfunction
 * @returns {string} 
 */
async function callContextSync10Times() {
  var result = "Initial value";
  await Excel.run(async (context) => {
    for (let i = 0; i < 10; i++) {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getRange("A1");
      range.load("values");
      await context.sync();
      result = range.values[0][0] || "No value found";
    }
  });
  return result;
}

/**
 * Wait for a specified number of seconds.
 * @customfunction
 * @param {number} seconds The number of seconds to wait.
 * @returns {string} A message indicating the wait is over.
 */
async function waitXSecondsPromise(seconds) {
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve(`Waited ${seconds} seconds`);
    }, seconds * 1000);
  });
}

/**
 * This function will call the write API to write "Hello" to A1.
 * @customfunction
 * @supportSync
 * @returns {string} 
 */
async function syncCallWriteApi() {
  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("A1");
    range.values = [["Hello"]];
    await context.sync();
  });

  return "test";
}

/** 
 * This function will call the read API to read the value from A1.
 * @customfunction
 * @supportSync
 * @returns {string} 
 */
async function syncCallReadApi() {
  var result = "Initial value";
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("A1");
    range.load("values");
    await context.sync();
    console.log(range.values[0][0]);
    result = range.values[0][0] || "No value found";
  });

  return result;
}
