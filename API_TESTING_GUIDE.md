# Excel GPT Middleware - API Testing Guide

## 🚀 Server Status: RUNNING ✅

The Excel GPT Middleware is successfully running on `http://localhost:3000`

## 📊 Test Results Summary

### ✅ Working Endpoints
- **Health Check**: `GET /health` - Server health status
- **Detailed Health**: `GET /health/detailed` - Comprehensive system status  
- **Root Endpoint**: `GET /` - Service information and available endpoints
- **API Documentation**: `GET /api/docs` - Complete API reference

### 🔒 Protected Endpoints (Require Authentication)
- **Excel Workbooks**: `GET /api/excel/workbooks`
- **Excel Worksheets**: `GET /api/excel/worksheets`
- **Read Excel Range**: `POST /api/excel/read`
- **Write Excel Range**: `POST /api/excel/write`
- **Read Excel Table**: `POST /api/excel/read-table`
- **Add Table Rows**: `POST /api/excel/add-table-rows`
- **Batch Operations**: `POST /api/excel/batch`

## 🧪 Manual Testing Commands

### 1. Test Health Endpoints
```powershell
# Basic health check
Invoke-RestMethod -Uri "http://localhost:3000/health" -Method GET

# Detailed health with system info
Invoke-RestMethod -Uri "http://localhost:3000/health/detailed" -Method GET
```

### 2. Test Service Information
```powershell
# Get service info and available endpoints
Invoke-RestMethod -Uri "http://localhost:3000/" -Method GET

# Get API documentation
Invoke-RestMethod -Uri "http://localhost:3000/api/docs" -Method GET
```

### 3. Test Excel APIs (Will show authentication requirement)
```powershell
# Test workbooks endpoint (requires Azure AD auth)
try {
    Invoke-RestMethod -Uri "http://localhost:3000/api/excel/workbooks" -Method GET
} catch {
    Write-Host "Expected: Authentication required - $($_.Exception.Message)"
}

# Test read endpoint with sample data
$readData = @{
    driveId = "sample-drive-id"
    itemId = "sample-workbook-id"
    worksheetId = "sample-worksheet-id"
    range = "A1:C3"
} | ConvertTo-Json

try {
    Invoke-RestMethod -Uri "http://localhost:3000/api/excel/read" -Method POST -Body $readData -ContentType "application/json"
} catch {
    Write-Host "Expected: Authentication required - $($_.Exception.Message)"
}
```

## 🔧 Configuration for Real Usage

To use the Excel APIs with real data, configure these in your `.env` file:

```env
# Azure AD Configuration (Required for Excel operations)
AZURE_TENANT_ID=your-tenant-id-from-azure-portal
AZURE_CLIENT_ID=your-client-id-from-app-registration
AZURE_CLIENT_SECRET=your-client-secret-from-app-registration

# Optional API Key for additional security
API_KEY=your-secure-api-key

# Production settings
NODE_ENV=production
LOG_LEVEL=info
```

## 📝 Sample GPT Integration

Here's how a GPT would interact with the middleware:

### Example 1: Read Sales Data
```json
{
  "action": "Read sales data from Excel",
  "request": {
    "method": "POST",
    "url": "http://localhost:3000/api/excel/read",
    "headers": {
      "Content-Type": "application/json",
      "x-api-key": "your-api-key"
    },
    "body": {
      "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
      "itemId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
      "worksheetId": "{00000000-0001-0000-0000-000000000000}",
      "range": "A1:D10"
    }
  }
}
```

### Example 2: Update Inventory
```json
{
  "action": "Update inventory quantities",
  "request": {
    "method": "POST", 
    "url": "http://localhost:3000/api/excel/write",
    "headers": {
      "Content-Type": "application/json",
      "x-api-key": "your-api-key"
    },
    "body": {
      "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
      "itemId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
      "worksheetId": "{00000000-0001-0000-0000-000000000000}",
      "range": "B2:B5",
      "values": [
        [150],
        [200], 
        [75],
        [300]
      ]
    }
  }
}
```

### Example 3: Add New Records to Table
```json
{
  "action": "Add new products to sales table",
  "request": {
    "method": "POST",
    "url": "http://localhost:3000/api/excel/add-table-rows", 
    "headers": {
      "Content-Type": "application/json",
      "x-api-key": "your-api-key"
    },
    "body": {
      "driveId": "b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd",
      "itemId": "01BYE5RZ6QN3ZWBTUQOJFZYDG5OH6HWJKU",
      "worksheetId": "{00000000-0001-0000-0000-000000000000}",
      "tableName": "SalesTable",
      "rows": [
        ["Widget X", 100, 5000, "2024-01-15"],
        ["Widget Y", 150, 7500, "2024-01-15"]
      ]
    }
  }
}
```

## 🛡️ Security Features Demonstrated

1. **API Key Validation**: All Excel endpoints require valid API key
2. **Input Validation**: Requests are validated for proper format
3. **Rate Limiting**: Protection against abuse (100 req/15min general, 20 req/5min writes)
4. **Error Handling**: Comprehensive error responses with proper HTTP status codes
5. **Audit Logging**: All operations are logged with timestamps and user info

## 📈 Monitoring Endpoints

- **Basic Health**: `GET /health` - Quick health check
- **Detailed Health**: `GET /health/detailed` - System metrics and dependency status
- **Readiness Probe**: `GET /health/ready` - Container readiness
- **Liveness Probe**: `GET /health/live` - Container liveness

## 🔄 Next Steps for Production

1. **Configure Azure AD**:
   - Create app registration in Azure Portal
   - Add Microsoft Graph permissions
   - Generate client secret
   - Update `.env` file

2. **Set Up Real Excel Files**:
   - Upload Excel files to SharePoint/OneDrive
   - Get drive IDs and item IDs
   - Configure permissions

3. **Test with Real Data**:
   - Use actual workbook/worksheet IDs
   - Test read/write operations
   - Verify audit logging

4. **Deploy to Production**:
   - Use HTTPS
   - Configure proper environment variables
   - Set up monitoring and logging
   - Implement backup procedures

## 🎯 Current Status

✅ **Server Running**: Port 3000  
✅ **Health Endpoints**: Working  
✅ **API Documentation**: Available  
✅ **Security Middleware**: Active  
✅ **Error Handling**: Implemented  
✅ **Rate Limiting**: Configured  
⚠️ **Azure AD**: Needs configuration for Excel operations  
⚠️ **API Key**: Needs configuration for security  

The middleware is ready for Azure AD configuration and real Excel file integration!
