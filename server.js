// @ts-check

/**
 * Based on https://stackoverflow.com/a/8809472 (Public Domain / MIT)
 * @typedef {string&{__uuidBrand:never}} UUID
 */
function generateUUID() {
  var d = new Date().getTime(); // Timestamp
  var d2 = // Time in microseconds since page-load or 0 if unsupported
    typeof performance !== "undefined" && performance.now
      ? performance.now() * 1000
      : 0;
  return /** @type {UUID} */ (
    "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
      var r = Math.random() * 16; // random number between 0 and 16
      if (d > 0) {
        // Use timestamp until depleted
        r = (d + r) % 16 | 0;
        d = Math.floor(d / 16);
      } else {
        // Use microseconds since page-load if supported
        r = (d2 + r) % 16 | 0;
        d2 = Math.floor(d2 / 16);
      }
      return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
    })
  );
}

/** @type {{[key:UUID]:unknown}} */
var pool = {};
/**
 * @param {unknown} obj
 */
function addToPool(obj) {
  var uuid = generateUUID();
  while (pool.hasOwnProperty(uuid)) {
    uuid = generateUUID();
  }
  pool[uuid] = obj;
  return uuid;
}

/**
 * @typedef {UUID&{__excelApplicationBrand:never}} ExcelApplicationID
 * @typedef {UUID&{__excelWorkbooksBrand:never}} ExcelWorkbooksID
 * @typedef {UUID&{__excelWorkbookBrand:never}} ExcelWorkbookID
 * @typedef {UUID&{__excelWorksheetsBrand:never}} ExcelWorksheetsID
 * @typedef {UUID&{__excelWorksheetBrand:never}} ExcelWorksheetID
 * @typedef {UUID&{__excelRangeBrand:never}} ExcelRangeID
 */

/**
 * @overload
 * @param {"Excel.Application"} progId
 * @returns {ExcelApplicationID}
 */
/**
 * @param {string} progId
 * @returns {UUID}
 */
function wscriptCreateObject(progId) {
  return addToPool(WScript.CreateObject(progId));
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.application.visible
 * @param {ExcelApplicationID} excelApplicationId
 * @param {boolean} value
 */
function excelApplicationSetVisible(excelApplicationId, value) {
  // @ts-ignore
  return (pool[excelApplicationId].Visible = value);
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.application.displayalerts
 * @param {ExcelApplicationID} excelApplicationId
 * @param {boolean} value
 */
function excelApplicationSetDisplayAlerts(excelApplicationId, value) {
  // @ts-ignore
  return (pool[excelApplicationId].DisplayAlerts = value);
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.application.quit
 * @param {ExcelApplicationID} excelApplicationId
 */
function excelApplicationQuit(excelApplicationId) {
  // @ts-ignore
  return pool[excelApplicationId].Quit();
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.application.workbooks
 * @param {ExcelApplicationID} excelApplicationId
 * @returns {ExcelWorkbooksID}
 */
function excelApplicationGetWorkbooks(excelApplicationId) {
  return /** @type {ExcelWorkbooksID} */ (
    // @ts-ignore
    addToPool(pool[excelApplicationId].Workbooks)
  );
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.workbooks.add
 * @param {ExcelWorkbooksID} excelWorkbooksId
 * @returns {ExcelWorkbookID}
 */
function excelWorkbooksAdd(excelWorkbooksId) {
  return /** @type {ExcelWorkbookID} */ (
    // @ts-ignore
    addToPool(pool[excelWorkbooksId].Add())
  );
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.workbooks.count
 * @param {ExcelWorkbooksID} excelWorkbooksId
 * @returns {number}
 */
function excelWorkbooksGetCount(excelWorkbooksId) {
  // @ts-ignore
  return pool[excelWorkbooksId].Count;
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.workbooks.item
 * @param {ExcelWorkbooksID} excelWorkbooksId
 * @param {number} index FIXME
 * @returns {ExcelWorkbookID}
 */
function excelWorkbooksItem(excelWorkbooksId, index) {
  return /** @type {ExcelWorkbookID} */ (
    // @ts-ignore
    addToPool(pool[excelWorkbooksId].Item(index))
  );
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.workbook.worksheets
 * @param {ExcelWorkbookID} excelWorkbookId
 * @returns {ExcelWorksheetsID}
 */
function excelWorkbookGetWorksheets(excelWorkbookId) {
  return /** @type {ExcelWorksheetsID} */ (
    // @ts-ignore
    addToPool(pool[excelWorkbookId].Worksheets)
  );
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.worksheets.count
 * @type {(excelWorksheetsId:ExcelWorksheetsID)=>number}
 */
// @ts-ignore
var excelWorksheetsGetCount = excelWorkbooksGetCount;

/**
 * https://learn.microsoft.com/office/vba/api/excel.worksheets.item
 * @type {(excelWorksheetsId:ExcelWorksheetsID,index:number)=>ExcelWorksheetID}
 */
// @ts-ignore
var excelWorksheetsItem = excelWorkbooksItem;

/**
 * https://learn.microsoft.com/office/vba/api/excel.worksheet.cells
 * @param {ExcelWorksheetID} excelWorksheetId
 * @param {number} row
 * @param {number} column
 * @returns {ExcelRangeID}
 */
function excelWorksheetCells(excelWorksheetId, row, column) {
  return /** @type {ExcelRangeID} */ (
    // @ts-ignore
    addToPool(pool[excelWorksheetId].Cells(row, column))
  );
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.range.value
 * @param {ExcelRangeID} excelRangeId
 * @param {string} value FIXME
 */
function excelRangeSetValue(excelRangeId, value) {
  // @ts-ignore
  return (pool[excelRangeId].Value = value);
}

/**
 * https://learn.microsoft.com/office/vba/api/excel.range.value
 * @param {ExcelRangeID} excelRangeId
 * @returns {string} FIXME
 */
function excelRangeGetValue(excelRangeId) {
  // @ts-ignore
  return pool[excelRangeId].Value;
}

function test() {
  var excel = wscriptCreateObject("Excel.Application");
  try {
    excelApplicationSetVisible(excel, true);

    var workbooks = excelApplicationGetWorkbooks(excel);
    excelWorkbooksAdd(workbooks);
    if (excelWorkbooksGetCount(workbooks) !== 1) {
      throw new Error(
        "Assertion failed: excelWorkbooksGetCount(workbooks) !== 1"
      );
    }

    var workbook = excelWorkbooksItem(workbooks, 1);
    var worksheets = excelWorkbookGetWorksheets(workbook);
    if (excelWorksheetsGetCount(worksheets) !== 1) {
      throw new Error(
        "Assertion failed: excelWorksheetsGetCount(worksheets) !== 1"
      );
    }
    var worksheet = excelWorksheetsItem(worksheets, 1);

    var cell = excelWorksheetCells(worksheet, 1, 1);
    excelRangeSetValue(cell, "foo");
    if (excelRangeGetValue(cell) !== "foo") {
      throw new Error('Assertion failed: excelRangeGetValue(cell) !== "foo"');
    }
  } finally {
    excelApplicationSetDisplayAlerts(excel, false);
    excelApplicationQuit(excel);
  }
}

/**
 * Based on the code published on MDN between [March 5, 2014](https://web.archive.org/web/20140305223108/https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/JSON)? and [March 27, 2018](https://web.archive.org/web/20180327225742/https://developer.mozilla.org/en-us/docs/Web/JavaScript/Reference/Global_Objects/JSON)?, and is considered to be in the public domain (CC0). ([About the license](https://developer.mozilla.org/en-US/docs/MDN/Writing_guidelines/Attrib_copyright_license#code_samples))
 */
Function("return this")().JSON = {
  parse: function (sJSON) {
    return eval("(" + sJSON + ")");
  },
  stringify: (function () {
    var toString = Object.prototype.toString;
    var hasOwnProperty = Object.prototype.hasOwnProperty;
    var isArray =
      Array.isArray ||
      function (a) {
        return toString.call(a) === "[object Array]";
      };
    var escMap = {
      '"': '\\"',
      "\\": "\\\\",
      "\b": "\\b",
      "\f": "\\f",
      "\n": "\\n",
      "\r": "\\r",
      "\t": "\\t"
    };
    var escFunc = function (m) {
      return (
        escMap[m] || "\\u" + (m.charCodeAt(0) + 0x10000).toString(16).substr(1)
      );
    };
    var escRE = /[\\"\u0000-\u001F\u2028\u2029]/g;
    return function stringify(value) {
      if (value == null) {
        return "null";
      } else if (typeof value === "number") {
        return isFinite(value) ? value.toString() : "null";
      } else if (typeof value === "boolean") {
        return value.toString();
      } else if (typeof value === "object") {
        if (typeof value.toJSON === "function") {
          return stringify(value.toJSON());
        } else if (isArray(value)) {
          var res = "[";
          for (var i = 0; i < value.length; i++)
            res += (i ? ", " : "") + stringify(value[i]);
          return res + "]";
        } else if (toString.call(value) === "[object Object]") {
          var tmp = [];
          for (var k in value) {
            // in case "hasOwnProperty" has been shadowed
            if (hasOwnProperty.call(value, k))
              tmp.push(stringify(k) + ": " + stringify(value[k]));
          }
          return "{" + tmp.join(", ") + "}";
        }
      }
      return '"' + value.toString().replace(escRE, escFunc) + '"';
    };
  })()
};

/**
 * https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L9
 * @constant
 * @type {"2.0"}
 */
var JSONRPC_VERSION = "2.0";
/**
 * https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L8
 * @constant
 * @type {"2024-11-05"}
 */
var LATEST_PROTOCOL_VERSION = "2024-11-05";

/**
 * https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L83
 * @constant
 * @type {-32700}
 */
var PARSE_ERROR = -32700;
/**
 * https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L84
 * @constant
 * @type {-32600}
 */
var INVALID_REQUEST = -32600;
/**
 * https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L85
 * @constant
 * @type {-32601}
 */
var METHOD_NOT_FOUND = -32601;
/**
 * https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L86
 * @constant
 * @type {-32602}
 */
var INVALID_PARAMS = -32602;
/**
 * https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L87
 * @constant
 * @type {-32603}
 */
var INTERNAL_ERROR = -32603;

/**
 * @param {string} level
 * @param {string} message
 */
function log(level, message) {
  WScript.StdErr.WriteLine(
    "[" + new Date().toString() + "]\t[" + level + "]\t" + message
  );
}

/**
 * @param {unknown} id
 * @param {Record<string,unknown>} response
 */
function respond(id, response) {
  response.jsonrpc = JSONRPC_VERSION;
  /**
   * FIXME
   * https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L56
   */
  response.id = typeof id === "number" || typeof id === "string" ? id : "";

  var str = JSON.stringify(response);
  WScript.StdOut.WriteLine(str);
  log("debug", str);
}

/**
 * @constant
 * @type {{name:"jscript-excel-mcp-server";version:"0.1.0";}}
 */
var SERVER_INFO = {
  name: "jscript-excel-mcp-server",
  version: "0.1.0"
};

/**
 * @constant
 */
var TOOLS = [
  {
    name: "wscript_create_object",
    inputSchema: {
      type: "object",
      properties: {
        progId: { type: "string" }
      },
      required: ["progId"]
    }
  },
  {
    name: "excel_application_set_visible",
    inputSchema: {
      type: "object",
      properties: {
        excelApplicationId: { type: "string" },
        value: { type: "boolean" }
      },
      required: ["excelApplicationId", "value"]
    }
  },
  {
    name: "excel_application_set_display_alerts",
    inputSchema: {
      type: "object",
      properties: {
        excelApplicationId: { type: "string" },
        value: { type: "boolean" }
      },
      required: ["excelApplicationId", "value"]
    }
  },
  {
    name: "excel_application_quit",
    inputSchema: {
      type: "object",
      properties: {
        excelApplicationId: { type: "string" }
      },
      required: ["excelApplicationId"]
    }
  },
  {
    name: "excel_application_get_workbooks",
    inputSchema: {
      type: "object",
      properties: {
        excelApplicationId: { type: "string" }
      },
      required: ["excelApplicationId"]
    }
  },
  {
    name: "excel_workbooks_add",
    inputSchema: {
      type: "object",
      properties: {
        excelWorkbooksId: { type: "string" }
      },
      required: ["excelWorkbooksId"]
    }
  },
  {
    name: "excel_workbooks_get_count",
    inputSchema: {
      type: "object",
      properties: {
        excelWorkbooksId: { type: "string" }
      },
      required: ["excelWorkbooksId"]
    }
  },
  {
    name: "excel_workbooks_item",
    inputSchema: {
      type: "object",
      properties: {
        excelWorkbooksId: { type: "string" },
        index: { type: "number" }
      },
      required: ["excelWorkbooksId", "index"]
    }
  },
  {
    name: "excel_workbook_get_worksheets",
    inputSchema: {
      type: "object",
      properties: {
        excelWorkbookId: { type: "string" }
      },
      required: ["excelWorkbookId"]
    }
  },
  {
    name: "excel_worksheets_get_count",
    inputSchema: {
      type: "object",
      properties: {
        excelWorksheetsId: { type: "string" }
      },
      required: ["excelWorksheetsId"]
    }
  },
  {
    name: "excel_worksheets_item",
    inputSchema: {
      type: "object",
      properties: {
        excelWorksheetsId: { type: "string" },
        index: { type: "number" }
      },
      required: ["excelWorksheetsId", "index"]
    }
  },
  {
    name: "excel_worksheet_cells",
    inputSchema: {
      type: "object",
      properties: {
        excelWorksheetId: { type: "string" },
        row: { type: "number" },
        column: { type: "number" }
      },
      required: ["excelWorksheetId", "row", "column"]
    }
  },
  {
    name: "excel_range_set_value",
    inputSchema: {
      type: "object",
      properties: {
        excelRangeId: { type: "string" },
        value: { type: "string" }
      },
      required: ["excelRangeId", "value"]
    }
  },
  {
    name: "excel_range_get_value",
    inputSchema: {
      type: "object",
      properties: {
        excelRangeId: { type: "string" }
      },
      required: ["excelRangeId"]
    }
  }
];
/**
 * @constant
 */
var HANDLERS = {
  wscript_create_object: function (args) {
    return wscriptCreateObject(args["progId"]);
  },
  excel_application_set_visible: function (args) {
    return excelApplicationSetVisible(
      args["excelApplicationId"],
      args["value"]
    );
  },
  excel_application_set_display_alerts: function (args) {
    return excelApplicationSetDisplayAlerts(
      args["excelApplicationId"],
      args["value"]
    );
  },
  excel_application_quit: function (args) {
    return excelApplicationQuit(args["excelApplicationId"]);
  },
  excel_application_get_workbooks: function (args) {
    return excelApplicationGetWorkbooks(args["excelApplicationId"]);
  },
  excel_workbooks_add: function (args) {
    return excelWorkbooksAdd(args["excelWorkbooksId"]);
  },
  excel_workbooks_get_count: function (args) {
    return excelWorkbooksGetCount(args["excelWorkbooksId"]);
  },
  excel_workbooks_item: function (args) {
    return excelWorkbooksItem(args["excelWorkbooksId"], args["index"]);
  },
  excel_workbook_get_worksheets: function (args) {
    return excelWorkbookGetWorksheets(args["excelWorkbookId"]);
  },
  excel_worksheets_get_count: function (args) {
    return excelWorksheetsGetCount(args["excelWorksheetsId"]);
  },
  excel_worksheets_item: function (args) {
    return excelWorksheetsItem(args["excelWorksheetsId"], args["index"]);
  },
  excel_worksheet_cells: function (args) {
    return excelWorksheetCells(
      args["excelWorksheetId"],
      args["row"],
      args["column"]
    );
  },
  excel_range_set_value: function (args) {
    return excelRangeSetValue(args["excelRangeId"], args["value"]);
  },
  excel_range_get_value: function (args) {
    return excelRangeGetValue(args["excelRangeId"]);
  }
};

(function main() {
  while (!WScript.StdIn.AtEndOfStream) {
    var line = WScript.StdIn.ReadLine();
    /** @type {Record<string,unknown>} */
    var request = {};
    try {
      request = JSON.parse(line);
    } catch (e) {
      respond(request.id, {
        error: {
          code: PARSE_ERROR,
          message: "PARSE_ERROR: " + (e instanceof Error && e.message) || ""
        }
      });
      continue;
    }

    // https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L21
    if (typeof request.method !== "string") {
      respond(request.id, {
        error: {
          code: INVALID_REQUEST,
          message: "INVALID_REQUEST: " + JSON.stringify(request)
        }
      });
      continue;
    }

    switch (request.method) {
      case "initialize":
        // https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L163
        respond(request.id, {
          result: {
            protocolVersion: LATEST_PROTOCOL_VERSION,
            capabilities: { tools: {} },
            serverInfo: SERVER_INFO
          }
        });
        break;

      case "tools/list":
        // https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L635-L644
        respond(request.id, {
          result: {
            tools: TOOLS
          }
        });
        break;

      case "tools/call":
        // https://github.com/modelcontextprotocol/modelcontextprotocol/blob/c87a0da6d8c2436d56a6398023c80b0562224454/schema/2024-11-05/schema.ts#L646-L678
        if (
          typeof request.params !== "object" ||
          request.params === null ||
          typeof request.params["name"] !== "string"
        ) {
          respond(request.id, {
            error: {
              code: INVALID_PARAMS,
              message: "INVALID_PARAMS: " + JSON.stringify(request)
            }
          });
          break;
        }

        try {
          var ret = HANDLERS[request.params["name"]](
            request.params["arguments"]
          );
          respond(request.id, {
            result: {
              content: [
                {
                  type: "text",
                  text: JSON.stringify(ret)
                }
              ]
            }
          });
        } catch (e) {
          respond(request.id, {
            result: {
              content: [
                {
                  type: "text",
                  message: (e instanceof Error && e.message) || ""
                }
              ],
              isError: true
            }
          });
        }
        break;

      default:
        respond(request.id, {
          error: {
            code: METHOD_NOT_FOUND,
            message: "METHOD_NOT_FOUND: " + JSON.stringify(request)
          }
        });
    }
  }
})();
