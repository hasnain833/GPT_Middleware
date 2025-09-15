# Excel GPT Middleware API Documentation

This document provides comprehensive API documentation for the Excel GPT Middleware, including all endpoints, request/response formats, and usage examples.

## Base URL

```
http://localhost:3000/api/excel
```

## Authentication

All API endpoints require authentication using Azure AD Client Credentials Flow. The middleware handles token acquisition automatically.

### Headers

```http
Content-Type: application/json
x-user-id: user@company.com (optional, for audit logging)
```

## Rate Limits

- **General**: 100 requests per 15 minutes
- **Write Operations**: 20 requests per 5 minutes
- **Authentication**: 5 attempts per 15 minutes

## Endpoints

### Health Check Endpoints

#### GET /health
Basic health check endpoint.

**Response:**
```json
{
  "status": "healthy",
  "service": "excel-gpt-middleware",
  "timestamp": "2024-01-15T10:35:00Z",
  "uptime": 123.456,
  "version": "1.0.0",
  "checks": {
    "service": "healthy",
    "azure_auth": "healthy",
    "graph_api": "unknown",
    "memory": "healthy",
    "disk": "healthy"
  },
  "system": {
    "memory": {
      "rss": 45,
      "heapTotal": 25,
      "heapUsed": 15,
      "external": 2
    },
    "nodeVersion": "v18.17.0",
    "platform": "win32",
    "arch": "x64"
  }
}
```

#### GET /health/detailed
Detailed health check with dependency status.

**Response:**
```json
{
  "status": "healthy",
  "service": "excel-gpt-middleware",
  "timestamp": "2024-01-15T10:35:00Z",
  "uptime": 123.456,
  "version": "1.0.0",
  "checks": {
    "service": "healthy",
    "azure_auth": "healthy",
    "graph_api": "unknown",
    "memory": "healthy",
    "disk": "healthy"
  },
  "system": {
    "memory": {
      "rss": 45,
      "heapTotal": 25,
      "heapUsed": 15,
      "external": 2
    },
    "nodeVersion": "v18.17.0",
    "platform": "win32",
    "arch": "x64"
  }
}
```

### Excel Operations

#### GET /api/excel/workbooks
Get all accessible Excel workbooks from OneDrive and SharePoint.

**Response:**
```json
{
  "success": true,
  "data": {
    "workbooks": [
      {
        "id": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
        "name": "Sales_Report_2024.xlsx",
        "webUrl": "https://company.sharepoint.com/sites/team/Shared%20Documents/Sales_Report_2024.xlsx",
        "location": "SharePoint - Team Site",
        "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
        "lastModified": "2024-01-15T10:30:00Z",
        "size": 2048576
      }
    ],
    "count": 1
  },
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### GET /api/excel/worksheets
Get all worksheets in a specific workbook.

**Query Parameters:**
- `driveId` (required): Drive ID of the workbook
- `itemId` (required): Item ID of the workbook

**Example:**
```http
GET /api/excel/worksheets?driveId=b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd&itemId=01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU
```

**Response:**
```json
{
  "success": true,
  "data": {
    "worksheets": [
      {
        "id": "{00000000-0001-0000-0000-000000000000}",
        "name": "Sales Data",
        "position": 0,
        "visibility": "Visible"
      },
      {
        "id": "{00000000-0001-0000-0000-000000000001}",
        "name": "Summary",
        "position": 1,
        "visibility": "Visible"
      }
    ],
    "count": 2,
    "workbookId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU"
  },
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### POST /api/excel/read
Read data from a specific Excel range.

**Request Body:**
```json
{
  "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
  "itemId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
  "worksheetId": "{00000000-0001-0000-0000-000000000000}",
  "range": "A1:C10"
}
```

**Response:**
```json
{
  "success": true,
  "operation": "READ_RANGE",
  "data": {
    "range": "Sales Data!A1:C10",
    "values": [
      ["Product", "Quantity", "Revenue"],
      ["Widget A", 100, 5000],
      ["Widget B", 150, 7500],
      ["Widget C", 75, 3750]
    ],
    "formulas": [
      ["Product", "Quantity", "Revenue"],
      ["Widget A", 100, "=B2*50"],
      ["Widget B", 150, "=B3*50"],
      ["Widget C", 75, "=B4*50"]
    ],
    "text": [
      ["Product", "Quantity", "Revenue"],
      ["Widget A", "100", "5000"],
      ["Widget B", "150", "7500"],
      ["Widget C", "75", "3750"]
    ],
    "dimensions": {
      "rows": 4,
      "columns": 3
    }
  },
  "metadata": {
    "workbookId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
    "worksheetId": "{00000000-0001-0000-0000-000000000000}",
    "requestedRange": "A1:C10"
  },
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### POST /api/excel/write
Write data to a specific Excel range.

**Request Body:**
```json
{
  "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
  "itemId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
  "worksheetId": "{00000000-0001-0000-0000-000000000000}",
  "range": "A5:C5",
  "values": [
    ["Widget D", 200, 10000]
  ]
}
```

**Response:**
```json
{
  "success": true,
  "operation": "WRITE_RANGE",
  "data": {
    "range": "Sales Data!A5:C5",
    "values": [
      ["Widget D", 200, 10000]
    ],
    "dimensions": {
      "rows": 1,
      "columns": 3
    }
  },
  "metadata": {
    "workbookId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
    "worksheetId": "{00000000-0001-0000-0000-000000000000}",
    "requestedRange": "A5:C5",
    "cellsModified": 3
  },
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### POST /api/excel/read-table
Read data from an Excel table.

**Request Body:**
```json
{
  "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
  "itemId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
  "worksheetId": "{00000000-0001-0000-0000-000000000000}",
  "tableName": "SalesTable"
}
```

**Response:**
```json
{
  "success": true,
  "operation": "READ_TABLE",
  "data": {
    "table": {
      "id": "{5C28C4B8-4F35-4B58-88C6-5D7E6F8A9B0C}",
      "name": "SalesTable",
      "address": "Sales Data!A1:C10"
    },
    "headers": ["Product", "Quantity", "Revenue"],
    "rows": [
      ["Widget A", 100, 5000],
      ["Widget B", 150, 7500],
      ["Widget C", 75, 3750],
      ["Widget D", 200, 10000]
    ],
    "allValues": [
      ["Product", "Quantity", "Revenue"],
      ["Widget A", 100, 5000],
      ["Widget B", 150, 7500],
      ["Widget C", 75, 3750],
      ["Widget D", 200, 10000]
    ],
    "dimensions": {
      "rows": 5,
      "columns": 3
    }
  },
  "metadata": {
    "workbookId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
    "worksheetId": "{00000000-0001-0000-0000-000000000000}",
    "tableName": "SalesTable"
  },
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### POST /api/excel/add-table-rows
Add new rows to an Excel table.

**Request Body:**
```json
{
  "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
  "itemId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
  "worksheetId": "{00000000-0001-0000-0000-000000000000}",
  "tableName": "SalesTable",
  "rows": [
    ["Widget E", 125, 6250],
    ["Widget F", 300, 15000]
  ]
}
```

**Response:**
```json
{
  "success": true,
  "operation": "ADD_TABLE_ROWS",
  "data": {
    "rowsAdded": 2,
    "result": {
      "index": 4,
      "values": [
        ["Widget E", 125, 6250],
        ["Widget F", 300, 15000]
      ]
    }
  },
  "metadata": {
    "workbookId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
    "worksheetId": "{00000000-0001-0000-0000-000000000000}",
    "tableName": "SalesTable"
  },
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### POST /api/excel/batch
Perform multiple Excel operations in a single request.

**Request Body:**
```json
{
  "operations": [
    {
      "type": "read_range",
      "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
      "itemId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
      "worksheetId": "{00000000-0001-0000-0000-000000000000}",
      "range": "A1:A10"
    },
    {
      "type": "write_range",
      "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
      "itemId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
      "worksheetId": "{00000000-0001-0000-0000-000000000001}",
      "range": "B1:B1",
      "values": [["Updated Value"]]
    }
  ]
}
```

**Response:**
```json
{
  "success": true,
  "operation": "BATCH_OPERATIONS",
  "data": {
    "results": [
      {
        "index": 0,
        "operation": "read_range",
        "success": true,
        "data": {
          "range": "Sales Data!A1:A10",
          "values": [["Product"], ["Widget A"], ["Widget B"], ["Widget C"]],
          "rowCount": 4,
          "columnCount": 1
        }
      },
      {
        "index": 1,
        "operation": "write_range",
        "success": true,
        "data": {
          "range": "Summary!B1:B1",
          "values": [["Updated Value"]],
          "rowCount": 1,
          "columnCount": 1
        }
      }
    ],
    "errors": [],
    "summary": {
      "total": 2,
      "successful": 2,
      "failed": 0
    }
  },
  "timestamp": "2024-01-15T10:35:00Z"
}
```

## Error Responses

### Standard Error Format

All errors follow this format:

```json
{
  "error": "Error Type",
  "message": "Human-readable error message",
  "timestamp": "2024-01-15T10:35:00Z"
}
```

### Common Error Types

#### 400 Bad Request - Validation Error
```json
{
  "error": "Validation failed",
  "message": "Request data is invalid",
  "details": [
    {
      "field": "range",
      "message": "\"range\" is required",
      "value": null
    }
  ],
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### 401 Unauthorized - Authentication Error
```json
{
  "error": "Authentication failed",
  "message": "Unable to authenticate with Microsoft Graph API",
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### 403 Forbidden - Permission Error
```json
{
  "error": "Access denied",
  "message": "Write access denied: No write permission",
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### 404 Not Found - Resource Error
```json
{
  "error": "Not found",
  "message": "Requested resource not found",
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### 429 Too Many Requests - Rate Limit Error
```json
{
  "error": "Rate limit exceeded",
  "message": "Too many requests, please try again later",
  "retryAfter": 900,
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### 500 Internal Server Error
```json
{
  "error": "Internal server error",
  "message": "An unexpected error occurred",
  "timestamp": "2024-01-15T10:35:00Z"
}
```

#### 502 Bad Gateway - Graph API Error
```json
{
  "error": "Graph API error",
  "message": "Microsoft Graph service error",
  "timestamp": "2024-01-15T10:35:00Z"
}
```

## Audit Logs

### GET /api/excel/logs
Get audit logs for Excel operations.

**Query Parameters:**
- `startDate` (optional): Filter logs from this date (ISO 8601 format)
- `endDate` (optional): Filter logs to this date (ISO 8601 format)
- `operation` (optional): Filter by operation type (READ, WRITE, etc.)
- `user` (optional): Filter by user ID
- `limit` (optional): Maximum number of logs to return (default: 100)

**Response:**
```json
{
  "success": true,
  "data": {
    "logs": [
      {
        "id": "audit-001",
        "timestamp": "2024-01-15T10:30:00Z",
        "operation": "READ",
        "user": "user@company.com",
        "workbookId": "sample-workbook-id",
        "worksheetId": "Sheet1",
        "range": "A1:C10",
        "success": true,
        "requestId": "req-123",
        "ipAddress": "192.168.1.100"
      }
    ],
    "filters": {
      "startDate": null,
      "endDate": null,
      "operation": null,
      "user": null,
      "limit": 100
    },
    "count": 1,
    "totalCount": 1
  },
  "timestamp": "2024-01-15T10:30:00Z"
}
```

## Data Types and Validation

### Range Format
Excel ranges must follow these patterns:
- Single cell: `A1`, `B5`, `Z100`
- Cell range: `A1:C10`, `B2:F20`
- Column range: `A:C`, `B:B`
- Row range: `1:5`, `10:15`

### Values Array
Data values must be provided as a 2D array:
```json
{
  "values": [
    ["Header 1", "Header 2", "Header 3"],
    ["Value 1", "Value 2", "Value 3"],
    ["Value 4", "Value 5", "Value 6"]
  ]
}
```

### Supported Data Types
- **String**: Text values
- **Number**: Integer or decimal numbers
- **Boolean**: `true` or `false`
- **Date**: ISO 8601 format (`2024-01-15T10:35:00Z`)
- **Formula**: Excel formulas (e.g., `=SUM(A1:A10)`)
- **Null**: Empty cells

## Best Practices

### 1. Batch Operations
Use batch operations for multiple related changes:
```json
{
  "operations": [
    {"type": "read_range", ...},
    {"type": "write_range", ...}
  ]
}
```

### 2. Error Handling
Always check the `success` field and handle errors appropriately:
```javascript
if (response.success) {
  // Process data
  const values = response.data.values;
} else {
  // Handle error
  console.error(response.error, response.message);
}
```

### 3. Rate Limiting
Implement exponential backoff for rate limit errors:
```javascript
async function makeRequest(url, data, retries = 3) {
  try {
    return await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
  } catch (error) {
    if (error.status === 429 && retries > 0) {
      const delay = Math.pow(2, 3 - retries) * 1000;
      await new Promise(resolve => setTimeout(resolve, delay));
      return makeRequest(url, data, retries - 1);
    }
    throw error;
  }
}
```

### 4. Large Data Sets
For large data sets, consider:
- Breaking into smaller chunks
- Using pagination where available
- Implementing progress tracking

### 5. Security
- Never log sensitive data
- Use HTTPS in production
- Implement proper authentication
- Validate all inputs

## Integration Examples

### GPT Integration
Example of how a GPT would interact with the middleware:

```javascript
// GPT function to read Excel data
async function readExcelData(workbookId, worksheetName, range) {
  const response = await fetch('/api/excel/read', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      driveId: process.env.DRIVE_ID,
      itemId: workbookId,
      worksheetId: worksheetName,
      range: range
    })
  });
  
  const data = await response.json();
  
  if (data.success) {
    return data.data.values;
  } else {
    throw new Error(`Failed to read Excel data: ${data.message}`);
  }
}

// GPT function to update Excel data
async function updateExcelData(workbookId, worksheetName, range, values) {
  const response = await fetch('/api/excel/write', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      driveId: process.env.DRIVE_ID,
      itemId: workbookId,
      worksheetId: worksheetName,
      range: range,
      values: values
    })
  });
  
  const data = await response.json();
  
  if (!data.success) {
    throw new Error(`Failed to update Excel data: ${data.message}`);
  }
  
  return data.data;
}
```

## Monitoring and Logging

### Request Logging
All requests are automatically logged with:
- Request ID
- Method and URL
- Response status
- Duration
- IP address
- User agent

### Audit Logging
All Excel operations are logged with:
- User ID
- Timestamp
- Operation type
- Workbook/worksheet/range
- Old and new values (for writes)
- Success/failure status

### Health Monitoring
Monitor these endpoints for system health:
- `/health` - Basic health check
- `/health/detailed` - Comprehensive health status
- `/health/ready` - Readiness probe
- `/health/live` - Liveness probe

## Limits and Quotas

### Request Limits
- Maximum request size: 10MB
- Maximum cells per request: 10,000
- Maximum batch operations: 20
- Maximum table rows per request: 1,000

### Excel Limits
- Maximum Excel rows: 1,048,576
- Maximum Excel columns: 16,384
- Maximum cell content: 32,767 characters
- Maximum worksheet name: 31 characters
- Maximum table name: 255 characters

### Microsoft Graph Limits
- Rate limit: 10,000 requests per hour
- File size limit: 250MB
- Request timeout: 30 seconds
