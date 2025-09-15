# Excel GPT Middleware

A secure middleware solution that connects custom GPTs to Excel files stored on SharePoint/OneDrive using Microsoft Graph API with Azure AD authentication.

## ✅ Status: Production Ready

This middleware is fully functional and tested with:
- ✅ Azure AD Client Credentials authentication
- ✅ Microsoft Graph API integration
- ✅ SharePoint/OneDrive read/write permissions
- ✅ Comprehensive error handling and validation
- ✅ Rate limiting and security measures
- ✅ Audit logging and monitoring

## Features

- **Azure AD Authentication**: Automatic service-to-service authentication
- **Excel Integration**: Full read/write access to Excel ranges and tables
- **SharePoint/OneDrive Access**: Works with organizational files
- **Security**: Built-in rate limiting, validation, and audit logging
- **Error Handling**: Comprehensive error responses and logging
- **Production Ready**: Includes monitoring, health checks, and graceful shutdown

## Project Structure

```
excel-gpt-middleware/
├── src/
│   ├── auth/
│   │   ├── azureAuth.js          # Azure AD authentication
│   │   └── middleware.js         # Authentication middleware
│   ├── controllers/
│   │   ├── excelController.js    # Excel operations controller
│   │   └── healthController.js   # Health check endpoints
│   ├── services/
│   │   ├── excelService.js       # Excel API service
│   │   ├── graphService.js       # Microsoft Graph service
│   │   └── auditService.js       # Audit logging service
│   ├── middleware/
│   │   ├── validation.js         # Request validation
│   │   ├── errorHandler.js       # Error handling
│   │   └── rateLimiter.js        # Rate limiting
│   ├── config/
│   │   ├── database.js           # Database configuration
│   │   ├── logger.js             # Winston logger setup
│   │   └── permissions.js        # Role-based permissions
│   ├── routes/
│   │   ├── excel.js              # Excel API routes
│   │   └── health.js             # Health check routes
│   ├── utils/
│   │   ├── helpers.js            # Utility functions
│   │   └── constants.js          # Application constants
│   └── server.js                 # Main server file
├── tests/
│   ├── unit/
│   └── integration/
├── docs/
│   ├── API.md                    # API documentation
│   ├── SETUP.md                  # Setup instructions
│   └── DEPLOYMENT.md             # Deployment guide
├── logs/                         # Log files directory
├── .env.example                  # Environment variables template
├── .gitignore
├── package.json
└── README.md
```

## Quick Start

1. **Configure Azure AD**: Set up your organizational Azure AD app with required permissions
2. **Environment Setup**: Copy `.env.example` to `.env` and add your Azure AD credentials:
   ```bash
   cp .env.example .env
   # Edit .env with your Azure AD values
   ```
3. **Install & Run**:
   ```bash
   npm install
   npm start
   ```
4. **Verify**: Server runs on `http://localhost:3000`

## Environment Files

- **`.env`** - Your actual Azure AD credentials (not in git)
- **`.env.example`** - Template showing required variables

## API Endpoints

### Excel Operations
- `GET /api/excel/workbooks` - List accessible Excel workbooks
- `GET /api/excel/worksheets` - Get worksheets in a workbook
- `POST /api/excel/read` - Read data from Excel ranges
- `POST /api/excel/write` - Write data to Excel ranges
- `POST /api/excel/read-table` - Read data from Excel tables
- `POST /api/excel/add-table-rows` - Add rows to Excel tables
- `POST /api/excel/batch` - Perform batch operations

### Health & Monitoring
- `GET /health` - Basic health check
- `GET /health/detailed` - Detailed system status
- `GET /api/docs` - API documentation

## Documentation

- [Client Setup Guide](CLIENT_SETUP_GUIDE.md) - Easy setup instructions for clients
- [API Documentation](docs/API.md) - Complete API reference
- [Deployment Guide](docs/DEPLOYMENT.md) - Production deployment instructions

## Authentication

This middleware uses **Azure AD Client Credentials Flow** for automatic authentication:
- No API keys required
- No JWT tokens needed  
- Service-to-service authentication handled automatically
- Enterprise-grade security through Azure AD

## Security Features

- Azure AD Client Credentials authentication
- Automatic token management and refresh
- Rate limiting (100 requests per 15 minutes)
- Input validation and sanitization
- Comprehensive audit logging
- CORS protection and security headers

## Integration Example

```javascript
// Simple API call - no authentication headers needed
const response = await fetch('http://localhost:3000/api/excel/workbooks');
const workbooks = await response.json();

// Read Excel data
const data = await fetch('http://localhost:3000/api/excel/read', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    driveId: "your-drive-id",
    itemId: "your-workbook-id",
    worksheetId: "Sheet1", 
    range: "A1:C10"
  })
});
```

## License

MIT License - see LICENSE file for details
