/**
 * Excel Controller
 * Handles HTTP requests for Excel operations
 */

const excelService = require('../services/excelService');
const auditService = require('../services/auditService');
const logger = require('../config/logger');
const { catchAsync } = require('../middleware/errorHandler');

class ExcelController {
    /**
     * Get all accessible workbooks
     */
    getWorkbooks = catchAsync(async (req, res) => {
        const auditContext = auditService.createAuditContext(req);
        
        const workbooks = await excelService.getWorkbooks(req.accessToken, auditContext);
        
        res.status(200).json({
            success: true,
            data: {
                workbooks: workbooks,
                count: workbooks.length
            },
            timestamp: new Date().toISOString()
        });
    });

    /**
     * Get worksheets in a workbook
     */
    getWorksheets = catchAsync(async (req, res) => {
        const { driveId, itemId } = req.query;
        const auditContext = auditService.createAuditContext(req);
        
        const worksheets = await excelService.getWorksheets(
            req.accessToken, 
            driveId, 
            itemId, 
            auditContext
        );
        
        res.status(200).json({
            success: true,
            data: {
                worksheets: worksheets,
                count: worksheets.length,
                workbookId: itemId
            },
            timestamp: new Date().toISOString()
        });
    });

    /**
     * Read data from Excel range
     */
    readRange = catchAsync(async (req, res) => {
        const { driveId, itemId, worksheetId, range } = req.body;
        const auditContext = auditService.createAuditContext(req);
        
        const data = await excelService.readRange({
            accessToken: req.accessToken,
            driveId,
            itemId,
            worksheetId,
            range,
            auditContext
        });
        
        res.status(200).json({
            success: true,
            operation: 'READ_RANGE',
            data: {
                range: data.address,
                values: data.values,
                formulas: data.formulas,
                text: data.text,
                dimensions: {
                    rows: data.rowCount,
                    columns: data.columnCount
                }
            },
            metadata: {
                workbookId: itemId,
                worksheetId: worksheetId,
                requestedRange: range
            },
            timestamp: new Date().toISOString()
        });
    });

    /**
     * Write data to Excel range
     */
    writeRange = catchAsync(async (req, res) => {
        const { driveId, itemId, worksheetId, range, values } = req.body;
        const auditContext = auditService.createAuditContext(req);
        
        const data = await excelService.writeRange({
            accessToken: req.accessToken,
            driveId,
            itemId,
            worksheetId,
            range,
            values,
            auditContext
        });
        
        res.status(200).json({
            success: true,
            operation: 'WRITE_RANGE',
            data: {
                range: data.address,
                values: data.values,
                dimensions: {
                    rows: data.rowCount,
                    columns: data.columnCount
                }
            },
            metadata: {
                workbookId: itemId,
                worksheetId: worksheetId,
                requestedRange: range,
                cellsModified: data.rowCount * data.columnCount
            },
            timestamp: new Date().toISOString()
        });
    });

    /**
     * Read data from Excel table
     */
    readTable = catchAsync(async (req, res) => {
        const { driveId, itemId, worksheetId, tableName } = req.body;
        const auditContext = auditService.createAuditContext(req);
        
        const data = await excelService.readTable({
            accessToken: req.accessToken,
            driveId,
            itemId,
            worksheetId,
            tableName,
            auditContext
        });
        
        res.status(200).json({
            success: true,
            operation: 'READ_TABLE',
            data: {
                table: {
                    id: data.id,
                    name: data.name,
                    address: data.address
                },
                headers: data.headers,
                rows: data.rows,
                allValues: data.values,
                dimensions: {
                    rows: data.rowCount,
                    columns: data.columnCount
                }
            },
            metadata: {
                workbookId: itemId,
                worksheetId: worksheetId,
                tableName: tableName
            },
            timestamp: new Date().toISOString()
        });
    });

    /**
     * Add rows to Excel table
     */
    addTableRows = catchAsync(async (req, res) => {
        const { driveId, itemId, worksheetId, tableName, rows } = req.body;
        const auditContext = auditService.createAuditContext(req);
        
        const result = await excelService.addTableRows({
            accessToken: req.accessToken,
            driveId,
            itemId,
            worksheetId,
            tableName,
            rows,
            auditContext
        });
        
        res.status(201).json({
            success: true,
            operation: 'ADD_TABLE_ROWS',
            data: {
                rowsAdded: rows.length,
                result: result
            },
            metadata: {
                workbookId: itemId,
                worksheetId: worksheetId,
                tableName: tableName
            },
            timestamp: new Date().toISOString()
        });
    });

    /**
     * Batch operations - perform multiple Excel operations in sequence
     */
    batchOperations = catchAsync(async (req, res) => {
        const { operations } = req.body;
        const auditContext = auditService.createAuditContext(req);
        
        if (!Array.isArray(operations) || operations.length === 0) {
            return res.status(400).json({
                success: false,
                error: 'Invalid operations array',
                timestamp: new Date().toISOString()
            });
        }

        const results = [];
        const errors = [];

        for (let i = 0; i < operations.length; i++) {
            const operation = operations[i];
            
            try {
                let result;
                
                switch (operation.type) {
                    case 'READ_range':
                        result = await excelService.readRange({
                            accessToken: req.accessToken,
                            driveId: operation.driveId,
                            itemId: operation.itemId,
                            worksheetId: operation.worksheetId,
                            range: operation.range,
                            auditContext
                        });
                        break;
                        
                    case 'write_range':
                        result = await excelService.writeRange({
                            accessToken: req.accessToken,
                            driveId: operation.driveId,
                            itemId: operation.itemId,
                            worksheetId: operation.worksheetId,
                            range: operation.range,
                            values: operation.values,
                            auditContext
                        });
                        break;
                        
                    case 'read_table':
                        result = await excelService.readTable({
                            accessToken: req.accessToken,
                            driveId: operation.driveId,
                            itemId: operation.itemId,
                            worksheetId: operation.worksheetId,
                            tableName: operation.tableName,
                            auditContext
                        });
                        break;
                        
                    case 'add_table_rows':
                        result = await excelService.addTableRows({
                            accessToken: req.accessToken,
                            driveId: operation.driveId,
                            itemId: operation.itemId,
                            worksheetId: operation.worksheetId,
                            tableName: operation.tableName,
                            rows: operation.rows,
                            auditContext
                        });
                        break;
                        
                    default:
                        throw new Error(`Unknown operation type: ${operation.type}`);
                }
                
                results.push({
                    index: i,
                    operation: operation.type,
                    success: true,
                    data: result
                });
                
            } catch (error) {
                logger.error(`Batch operation ${i} failed:`, error);
                errors.push({
                    index: i,
                    operation: operation.type,
                    error: error.message
                });
                
                // Continue with other operations unless it's a critical error
                if (error.message.includes('Authentication')) {
                    break; // Stop if authentication fails
                }
            }
        }

        const response = {
            success: errors.length === 0,
            operation: 'BATCH_OPERATIONS',
            data: {
                results: results,
                errors: errors,
                summary: {
                    total: operations.length,
                    successful: results.length,
                    failed: errors.length
                }
            },
            timestamp: new Date().toISOString()
        };

        // Return 207 Multi-Status if there were partial failures
        const statusCode = errors.length > 0 && results.length > 0 ? 207 : 
                          errors.length === 0 ? 200 : 400;
        
        res.status(statusCode).json(response);
    });

    /**
     * Get Excel file metadata
     */
    getFileMetadata = catchAsync(async (req, res) => {
        const { driveId, itemId } = req.query;
        const auditContext = auditService.createAuditContext(req);
        
        // This would typically get file metadata from Graph API
        // For now, return basic structure
        res.status(200).json({
            success: true,
            data: {
                fileId: itemId,
                driveId: driveId,
                // Additional metadata would be fetched from Graph API
            },
            timestamp: new Date().toISOString()
        });
    });

    /**
     * Get audit logs
     */
    getLogs = catchAsync(async (req, res) => {
        const { startDate, endDate, operation, user, limit = 100 } = req.query;
        const auditContext = auditService.createAuditContext(req);
        
        // In a real implementation, this would query a database or log files
        // For now, we'll return a sample response showing the log structure
        const logs = {
            logs: [
                {
                    id: "audit-001",
                    timestamp: new Date().toISOString(),
                    operation: "READ",
                    user: auditContext.user,
                    workbookId: "sample-workbook-id",
                    worksheetId: "Sheet1",
                    range: "A1:C10",
                    success: true,
                    requestId: auditContext.requestId,
                    ipAddress: auditContext.ipAddress
                }
            ],
            filters: {
                startDate: startDate || null,
                endDate: endDate || null,
                operation: operation || null,
                user: user || null,
                limit: parseInt(limit)
            },
            count: 1,
            totalCount: 1
        };
        
        auditService.logSystemEvent({
            event: 'AUDIT_LOG_REQUEST',
            details: { filters: logs.filters, requestedBy: auditContext.user }
        });
        
        res.status(200).json({
            success: true,
            data: logs,
            timestamp: new Date().toISOString()
        });
    });
}

module.exports = new ExcelController();
