#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListResourcesRequestSchema,
  ListToolsRequestSchema,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import XLSX from 'xlsx';
import fs from 'fs';


const server = new Server(
  {
    name: "spreadsheet-query-server",
    version: "0.1.0",
  },
  {
    capabilities: {
      resources: {},
      tools: {},
    },
  },
);

const spreadsheetPath = "/Users/marvinirwin/Downloads/80285249-a89f-4398-a305-4611e71fca9e/accounting.xlsm";

if (!fs.existsSync(spreadsheetPath)) {
  process.exit(1);
}

// Load workbook
let workbook;
try {
  workbook = XLSX.readFile(spreadsheetPath);
} catch (error) {
  process.exit(1);
}

// Create resource base URL
const resourceBaseUrl = new URL('file://' + spreadsheetPath);

server.setRequestHandler(ListResourcesRequestSchema, async () => {
  return {
    resources: workbook.SheetNames.map((sheetName) => ({
      uri: new URL(`${sheetName}/schema`, resourceBaseUrl).href,
      mimeType: "application/json",
      name: `"${sheetName}" sheet schema`,
    })),
  };
});

server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
  const resourceUrl = new URL(request.params.uri);
  const pathComponents = resourceUrl.pathname.split("/");
  const schema = pathComponents.pop();
  const sheetName = pathComponents.pop();
  
  if (schema !== "schema") {
    throw new Error("Invalid resource URI");
  }
  
  if (!workbook.SheetNames.includes(sheetName)) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }
  
  const worksheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  
  const headers = [];
  const firstRow = range.s.r;
  const lastRow = range.e.r;
  
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cellAddress = XLSX.utils.encode_cell({ r: firstRow, c });
    const cellValue = worksheet[cellAddress]?.v || '';
    headers.push(cellValue);
  }
  
  const schemaInfo = headers.map((header, index) => {
    let dataType = 'unknown';
    if (lastRow > firstRow) {
      const cellAddress = XLSX.utils.encode_cell({ r: firstRow + 1, c: range.s.c + index });
      const cell = worksheet[cellAddress];
      if (cell) {
        if (typeof cell.v === 'number') dataType = 'number';
        else if (typeof cell.v === 'string') dataType = 'string';
        else if (typeof cell.v === 'boolean') dataType = 'boolean';
      }
    }
    
    return {
      column_name: header,
      data_type: dataType
    };
  });
  
  return {
    contents: [
      {
        uri: request.params.uri,
        mimeType: "application/json",
        text: JSON.stringify(schemaInfo, null, 2),
      },
    ],
  };
});

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: "querySheet",
        description: "Query data from a specific sheet in the spreadsheet",
        inputSchema: {
          type: "object",
          properties: {
            sheetName: { type: "string" },
            limit: { type: "number", optional: true }
          },
        },
      },
      {
        name: "customFilterSheet",
        description: "Filter spreadsheet data using a custom JavaScript filter function",
        inputSchema: {
          type: "object",
          properties: {
            sheetName: { type: "string" },
            filterCode: { type: "string", description: "JavaScript filter function that takes a row object as input and returns a boolean" },
            sumColumn: { type: "string", optional: true, description: "Column name to sum after filtering (optional)" },
            limit: { type: "number", optional: true }
          },
          required: ["sheetName", "filterCode"]
        },
      },
      {
        name: "listSheets",
        description: "List all available sheets in the spreadsheet",
        inputSchema: {
          type: "object",
          properties: {
            random_string: { type: "string", description: "Dummy parameter for no-parameter tools" }
          },
          required: ["random_string"]
        },
      },
      {
        name: "getLastRow",
        description: "Get the last row of data from a specific sheet in the spreadsheet",
        inputSchema: {
          type: "object",
          properties: {
            sheetName: { type: "string" }
          },
          required: ["sheetName"]
        },
      },
      {
        name: "getRowCount",
        description: "Get the total number of rows in a specific sheet",
        inputSchema: {
          type: "object",
          properties: {
            sheetName: { type: "string" }
          },
          required: ["sheetName"]
        },
      },
      {
        name: "getDistinctValues",
        description: "Get distinct values from a specified column in the sheet",
        inputSchema: {
          type: "object",
          properties: {
            sheetName: { type: "string" },
            columnName: { type: "string" }
          },
          required: ["sheetName", "columnName"]
        },
      },
      {
        name: "reduceColumn",
        description: "Apply a reduce operation on a column (like SQL GROUP BY with aggregation)",
        inputSchema: {
          type: "object",
          properties: {
            sheetName: { type: "string" },
            groupByColumn: { type: "string", description: "Column to group by (similar to SQL DISTINCT/GROUP BY)" },
            reduceCode: { type: "string", description: "JavaScript code block for reduction. Must include a return statement. You can use if statements, loops, and other JavaScript features. The code has access to: accumulator (current value), row (current row data), and getDateFromExcel(excelDate) helper to convert dates." },
            initialValue: { type: "string", description: "Initial value for the reducer (as JSON string)" }
          },
          required: ["sheetName", "groupByColumn", "reduceCode", "initialValue"]
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  if (request.params.name === "querySheet") {
    const { sheetName, limit = 50 } = request.params.arguments as { sheetName: string, limit?: number };
    
    if (!workbook.SheetNames.includes(sheetName)) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    const limitedData = jsonData.slice(0, Math.min(limit, 50));
    
    return {
      content: [{ type: "text", text: JSON.stringify(limitedData, null, 2) }],
      isError: false,
    };
  } else if (request.params.name === "customFilterSheet") {
    interface SpreadsheetRow {
      [key: string]: any;
    }
    
    const { sheetName, filterCode, sumColumn, limit = 50 } = request.params.arguments as { 
      sheetName: string, 
      filterCode: string,
      sumColumn?: string,
      limit?: number
    };
    
    if (!workbook.SheetNames.includes(sheetName)) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet) as SpreadsheetRow[];
    
    // Helper function for working with Excel dates
    const getDateFromExcel = (excelDate: number): Date | null => {
      if (!excelDate) return null;
      // Excel dates are stored as number of days since 1900-01-01
      const excelEpoch = new Date(1900, 0, 1);
      return new Date(excelEpoch.getTime() + (excelDate - 1) * 86400000);
    };
    
    try {
      // Create a filter function from the provided code string
      // This is safe in a controlled environment but would need security review for production
      const filterFunction = new Function('row', 'getDateFromExcel', `
        try {
          return (${filterCode});
        } catch (error) {
          console.error("Error in filter function:", error);
          return false;
        }
      `) as (row: SpreadsheetRow, getDateFromExcel: (date: number) => Date | null) => boolean;
      
      // Apply the filter function to each row
      const filteredData = jsonData.filter(row => filterFunction(row, getDateFromExcel));
      
      // Apply limit if needed
      const limitedFilteredData = filteredData.slice(0, Math.min(limit, 50));
      
      // Calculate sum if requested
      interface FilterResult {
        matchingRows: SpreadsheetRow[];
        count: number;
        sum?: number;
        sumColumn?: string;
      }
      
      let result: FilterResult = {
        matchingRows: limitedFilteredData,
        count: filteredData.length
      };
      
      if (sumColumn && filteredData.some(row => row[sumColumn] !== undefined)) {
        const sum = filteredData.reduce((acc, row) => {
          // Only add numeric values to the sum
          const value = row[sumColumn];
          return acc + (typeof value === 'number' ? value : 0);
        }, 0);
        
        result = {
          ...result,
          sum,
          sumColumn
        };
      }
      
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        isError: false,
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      return {
        content: [{ type: "text", text: `Error executing filter: ${errorMessage}` }],
        isError: true,
      };
    }
  } else if (request.params.name === "listSheets") {
    return {
      content: [{ type: "text", text: JSON.stringify({ sheets: workbook.SheetNames }, null, 2) }],
      isError: false,
    };
  } else if (request.params.name === "getLastRow") {
    const { sheetName } = request.params.arguments as { sheetName: string };
    
    if (!workbook.SheetNames.includes(sheetName)) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    if (jsonData.length === 0) {
      return {
        content: [{ type: "text", text: JSON.stringify({ error: "Sheet is empty" }, null, 2) }],
        isError: false,
      };
    }
    
    // Find the last row that has a valid Transaction ID (not a formula or text description)
    let lastDataRowIndex = jsonData.length - 1;
    while (lastDataRowIndex >= 0) {
      const row = jsonData[lastDataRowIndex];
      const transactionId = row["Transaction ID"];
      
      // Check if this is a normal data row (Transaction ID should be a number)
      if (typeof transactionId === 'number' || 
          (typeof transactionId === 'string' && !isNaN(Number(transactionId)) && !transactionId.startsWith("="))) {
        break;
      }
      
      lastDataRowIndex--;
    }
    
    if (lastDataRowIndex < 0) {
      return {
        content: [{ type: "text", text: JSON.stringify({ error: "No valid data rows found" }, null, 2) }],
        isError: false,
      };
    }
    
    const lastDataRow = jsonData[lastDataRowIndex];
    
    return {
      content: [{ type: "text", text: JSON.stringify(lastDataRow, null, 2) }],
      isError: false,
    };
  } else if (request.params.name === "getRowCount") {
    const { sheetName } = request.params.arguments as { sheetName: string };
    
    if (!workbook.SheetNames.includes(sheetName)) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    return {
      content: [{ type: "text", text: JSON.stringify({ rowCount: jsonData.length }, null, 2) }],
      isError: false,
    };
  } else if (request.params.name === "getDistinctValues") {
    const { sheetName, columnName } = request.params.arguments as { sheetName: string, columnName: string };
    
    if (!workbook.SheetNames.includes(sheetName)) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet) as Record<string, unknown>[];
    
    // Check if column exists
    if (jsonData.length > 0 && !jsonData.some(row => columnName in row)) {
      return {
        content: [{ type: "text", text: JSON.stringify({ error: `Column '${columnName}' not found` }, null, 2) }],
        isError: false,
      };
    }
    
    // Get all values from the column
    const allValues = jsonData.map(row => row[columnName]).filter(value => value !== undefined && value !== null);
    
    // Create a Set to get unique values
    const uniqueValuesSet = new Set<string>();
    allValues.forEach(value => {
      // Convert to string to handle different types consistently
      if (value !== undefined && value !== null && value !== '') {
        uniqueValuesSet.add(String(value).trim());
      }
    });
    
    // Convert Set back to array and sort alphabetically
    const distinctValues = Array.from(uniqueValuesSet).sort();
    
    return {
      content: [{ type: "text", text: JSON.stringify({ 
        column: columnName,
        distinctValues: distinctValues.slice(0, 50),
        count: distinctValues.length,
        limitApplied: distinctValues.length > 50
      }, null, 2) }],
      isError: false,
    };
  } else if (request.params.name === "reduceColumn") {
    interface SpreadsheetRow {
      [key: string]: any;
    }
    
    const { sheetName, groupByColumn, reduceCode, initialValue } = request.params.arguments as { 
      sheetName: string, 
      groupByColumn: string,
      reduceCode: string,
      initialValue: string
    };
    
    if (!workbook.SheetNames.includes(sheetName)) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet) as SpreadsheetRow[];
    
    // Check if column exists
    if (jsonData.length > 0 && !jsonData.some(row => groupByColumn in row)) {
      return {
        content: [{ type: "text", text: JSON.stringify({ error: `Column '${groupByColumn}' not found` }, null, 2) }],
        isError: false,
      };
    }
    
    try {
      // Parse initial value from JSON string
      const parsedInitialValue = JSON.parse(initialValue);
      
      // Helper function for working with Excel dates
      const getDateFromExcel = (excelDate: number): Date | null => {
        if (!excelDate) return null;
        const excelEpoch = new Date(1900, 0, 1);
        return new Date(excelEpoch.getTime() + (excelDate - 1) * 86400000);
      };
      
      // Create the reducer function
      const reduceFunction = new Function('accumulator', 'row', 'getDateFromExcel', `
        try {
          ${reduceCode}
        } catch (error) {
          console.error("Error in reduce function:", error);
          return accumulator;
        }
      `) as (accumulator: any, row: SpreadsheetRow, getDateFromExcel: (date: number) => Date | null) => any;
      
      // Group by unique values in the specified column
      const groupedData: Record<string, any> = {};
      
      // First group the data
      jsonData.forEach(row => {
        if (row[groupByColumn] === undefined || row[groupByColumn] === null) return;
        
        const groupValue = String(row[groupByColumn]).trim();
        if (!groupedData[groupValue]) {
          // Create a deep copy of the initial value for each group
          groupedData[groupValue] = JSON.parse(JSON.stringify(parsedInitialValue));
        }
        
        groupedData[groupValue] = reduceFunction(groupedData[groupValue], row, getDateFromExcel);
      });
      
      // Convert to array format for return
      const resultArray = Object.entries(groupedData).map(([key, value]) => ({
        [groupByColumn]: key,
        result: value
      }));
      
      // Apply limit of 50
      const limitedResults = resultArray.slice(0, 50);
      
      return {
        content: [{ type: "text", text: JSON.stringify({
          groups: limitedResults,
          totalGroups: resultArray.length,
          limitApplied: resultArray.length > 50
        }, null, 2) }],
        isError: false,
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      return {
        content: [{ type: "text", text: `Error executing reduce: ${errorMessage}` }],
        isError: true,
      };
    }
  }
  
  throw new Error(`Unknown tool: ${request.params.name}`);
});

async function runServer() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

runServer().catch(console.error);