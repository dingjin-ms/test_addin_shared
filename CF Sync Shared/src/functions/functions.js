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
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 
 */
async function syncCallWriteApi(invocation) {
  const context = new Excel.RequestContext();
  context.setInvocation(invocation);

  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let range = sheet.getRange("A1");
  range.values = [["Hello"]];
  await context.sync();

  return "test";
}

/**
 * This function will call the write API to write "Hello" to A1.
 * @customfunction
 * @supportSync
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 
 */
async function syncCallWriteApiCatch(invocation) {
  const context = new Excel.RequestContext();
  context.setInvocation(invocation);

  try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getRange("A1");
      range.values = [["Hello"]];
      await context.sync();
  } catch (error) {
      console.error(error);
      return "test-error";
  }

  return "test";
}

/** 
 * This function will call the read API to read the value from A1.
 * @customfunction
 * @supportSync
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns {string} 
 */
async function syncCallReadApi(invocation) {
  const context = new Excel.RequestContext();
  context.setInvocation(invocation);

  var result = "Initial value";
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let range = sheet.getRange("A1");
  range.load("values");
  await context.sync();
  console.log(range.values[0][0]);
  result = range.values[0][0] || "No value found";

  return result;
}

/**
 * Search for products that match a given substring. Try =SCRIPTLAB.DATATYPESCUSTOMFUNCTIONS.PRODUCTSEARCH("chef", false).
 * @customfunction
 * @param {string} query The string to search for in the sample JSON data.
 * @param {boolean} [completeMatch] Define whether the search should be a match of the whole product name or part of the product name. If omitted, completeMatch = false.
 * @return {Promise<any[][]>} Search results as one or more data type entity values.
 */
async function productSearch(query, completeMatch) {
  // This function searches a set of sample JSON data for the string entered in the
  // custom function, and then returns the search result as one or more entity values.

  // Set up an error to use if a matching product doesn't exist in the JSON data.
  const notAvailableError = {
    type: "Error",
    errorType: "NotAvailable",
  };

  // Search the sample JSON data for matching product names.
  try {
    if (completeMatch === null) {
      completeMatch = false;
    }

    console.log(`Searching for ${query}...`);
    const searchResult = await searchProduct(query, completeMatch);

    // If the search result is empty, return the error.
    if (searchResult.length == 0) {
      return [[notAvailableError]];
    }

    // Create product entities for each of the products returned in the search result.
    const entities = searchResult.map((product) => [makeProductEntity(product)]);
    return entities;
  } catch (error) {
    console.error(error);
  }
}

// Helper function to create entities from product properties.
function makeProductEntity(product) {
  const entity = {
    type: "Entity",
    text: product.productName,
    properties: {
      "Product ID": {
        type: "String",
        basicValue: product.productID.toString() || "",
      },
      "Product Name": {
        type: "String",
        basicValue: product.productName || "",
      },
      "Quantity Per Unit": {
        type: "String",
        basicValue: product.quantityPerUnit || "",
      },
      // Add Unit Price as a formatted number.
      "Unit Price": {
        type: "Double",
        basicValue: product.unitPrice,
        numberFormat: "$* #,##0.00",
      },
      Discontinued: {
        type: "Boolean",
        basicValue: product.discontinued || false,
      },
    },
    layouts: {
      card: {
        title: { property: "Product Name" },
        sections: [
          {
            layout: "List",
            properties: ["Product ID"],
          },
          {
            layout: "List",
            title: "Quantity and price",
            collapsible: true,
            collapsed: false,
            properties: ["Quantity Per Unit", "Unit Price"],
          },
          {
            layout: "List",
            title: "Additional information",
            collapsed: true,
            properties: ["Discontinued"],
          },
        ],
      },
    },
  };

  // Add image property to the entity and then add it to the card layout.
  if (product.productImage) {
    entity.properties["Image"] = {
      type: "WebImage",
      address: product.productImage || "",
    };
    entity.layouts.card.mainImage = { property: "Image" };
  }

  return entity;
}

// Helper function to search the sample JSON product data.
function searchProduct(query, completeMatch) {
  const queryUpperCase = query.toUpperCase();
  if (completeMatch === true) {
    return products.filter((p) => p.productName.toUpperCase() === queryUpperCase);
  } else {
    return products.filter((p) => p.productName.toUpperCase().indexOf(queryUpperCase) >= 0);
  }
}

/** Sample JSON product data. */
const products = [
  {
    productID: 1,
    productName: "Chai",
    supplierID: 1,
    categoryID: 1,
    quantityPerUnit: "10 boxes x 20 bags",
    unitPrice: 18,
    discontinued: false,
    productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/0/04/Masala_Chai.JPG/320px-Masala_Chai.JPG",
  },
  {
    productID: 2,
    productName: "Chang",
    supplierID: 1,
    categoryID: 1,
    quantityPerUnit: "24 - 12 oz bottles",
    unitPrice: 19,
    discontinued: false,
    productImage: "",
  },
  {
    productID: 3,
    productName: "Aniseed Syrup",
    supplierID: 1,
    categoryID: 2,
    quantityPerUnit: "12 - 550 ml bottles",
    unitPrice: 10,
    discontinued: false,
    productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/8/81/Maltose_syrup.jpg/185px-Maltose_syrup.jpg",
  },
  {
    productID: 4,
    productName: "Chef Anton's Cajun Seasoning",
    supplierID: 2,
    categoryID: 2,
    quantityPerUnit: "48 - 6 oz jars",
    unitPrice: 22,
    discontinued: false,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/8/82/Kruidenmengeling-spice.jpg/193px-Kruidenmengeling-spice.jpg",
  },
  {
    productID: 5,
    productName: "Chef Anton's Gumbo Mix",
    supplierID: 2,
    categoryID: 2,
    quantityPerUnit: "36 boxes",
    unitPrice: 21.35,
    discontinued: true,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/4/45/Okra_in_a_Bowl_%28Unsplash%29.jpg/180px-Okra_in_a_Bowl_%28Unsplash%29.jpg",
  },
  {
    productID: 6,
    productName: "Grandma's Boysenberry Spread",
    supplierID: 3,
    categoryID: 2,
    quantityPerUnit: "12 - 8 oz jars",
    unitPrice: 25,
    discontinued: false,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/1/10/Making_cranberry_sauce_-_in_the_jar.jpg/90px-Making_cranberry_sauce_-_in_the_jar.jpg",
  },
  {
    productID: 7,
    productName: "Uncle Bob's Organic Dried Pears",
    supplierID: 3,
    categoryID: 7,
    quantityPerUnit: "12 - 1 lb pkgs.",
    unitPrice: 30,
    discontinued: false,
    productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/f/fd/DriedPears.JPG/120px-DriedPears.JPG",
  },
  {
    productID: 8,
    productName: "Northwoods Cranberry Sauce",
    supplierID: 3,
    categoryID: 2,
    quantityPerUnit: "12 - 12 oz jars",
    unitPrice: 40,
    discontinued: false,
    productImage:
      "https://upload.wikimedia.org/wikipedia/commons/thumb/0/07/Making_cranberry_sauce_-_stovetop.jpg/90px-Making_cranberry_sauce_-_stovetop.jpg",
  },
  {
    productID: 9,
    productName: "Mishi Kobe Niku",
    supplierID: 4,
    categoryID: 6,
    quantityPerUnit: "18 - 500 g pkgs.",
    unitPrice: 97,
    discontinued: true,
    productImage: "",
  },
  {
    productID: 10,
    productName: "Ikura",
    supplierID: 4,
    categoryID: 8,
    quantityPerUnit: "12 - 200 ml jars",
    unitPrice: 31,
    discontinued: false,
    productImage: "",
  },
];

const categories = [
  {
    categoryID: 1,
    categoryName: "Beverages",
    description: "Soft drinks, coffees, teas, beers, and ales",
  },
  {
    categoryID: 2,
    categoryName: "Condiments",
    description: "Sweet and savory sauces, relishes, spreads, and seasonings",
  },
  {
    categoryID: 3,
    categoryName: "Confections",
    description: "Desserts, candies, and sweet breads",
  },
  {
    categoryID: 4,
    categoryName: "Dairy Products",
    description: "Cheeses",
  },
  {
    categoryID: 5,
    categoryName: "Grains/Cereals",
    description: "Breads, crackers, pasta, and cereal",
  },
  {
    categoryID: 6,
    categoryName: "Meat/Poultry",
    description: "Prepared meats",
  },
  {
    categoryID: 7,
    categoryName: "Produce",
    description: "Dried fruit and bean curd",
  },
  {
    categoryID: 8,
    categoryName: "Seafood",
    description: "Seaweed and fish",
  },
];

const suppliers = [
  {
    supplierID: 1,
    companyName: "Exotic Liquids",
    contactName: "Charlotte Cooper",
    contactTitle: "Purchasing Manager",
  },
  {
    supplierID: 2,
    companyName: "New Orleans Cajun Delights",
    contactName: "Shelley Burke",
    contactTitle: "Order Administrator",
  },
  {
    supplierID: 3,
    companyName: "Grandma Kelly's Homestead",
    contactName: "Regina Murphy",
    contactTitle: "Sales Representative",
  },
  {
    supplierID: 4,
    companyName: "Tokyo Traders",
    contactName: "Yoshi Nagase",
    contactTitle: "Marketing Manager",
    address: "9-8 Sekimai Musashino-shi",
  },
  {
    supplierID: 5,
    companyName: "Cooperativa de Quesos 'Las Cabras'",
    contactName: "Antonio del Valle Saavedra",
    contactTitle: "Export Administrator",
  },
  {
    supplierID: 6,
    companyName: "Mayumi's",
    contactName: "Mayumi Ohno",
    contactTitle: "Marketing Representative",
  },
  {
    supplierID: 7,
    companyName: "Pavlova, Ltd.",
    contactName: "Ian Devling",
    contactTitle: "Marketing Manager",
  },
  {
    supplierID: 8,
    companyName: "Specialty Biscuits, Ltd.",
    contactName: "Peter Wilson",
    contactTitle: "Sales Representative",
  },
];
