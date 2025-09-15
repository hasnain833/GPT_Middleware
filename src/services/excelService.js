/**
 * Excel Service
 * High-level service for Excel operations with permission checking and validation
 */

const graphService = require('./graphService');
const auditService = require('./auditService');
const logger = require('../config/logger');
const permissions = require('../config/permissions');

class ExcelService {
    constructor() {
        this.graphService = graphService;
        this.auditService = auditService;
    }

    /**
     * Get all accessible workbooks
     * @param {string} accessToken - Access token
     * @param {Object} auditContext - Audit context
     * @returns {Promise<Array>} List of workbooks
     */
    async getWorkbooks(accessToken, auditContext) {
        try {
            const workbooks = await this.graphService.getWorkbooks(accessToken, auditContext);
            
            // Filter workbooks based on permissions (if configured)
            const filteredWorkbooks = workbooks.filter(workbook => 
                permissions.canAccessWorkbook(auditContext.user, workbook.id)
            );

            return filteredWorkbooks;
        } catch (error) {
            logger.error('Excel service - failed to get workbooks:', error);
            throw error;
        }
    }

    /**
     * Get worksheets in a workbook
     * @param {string} accessToken - Access token
     * @param {string} driveId - Drive ID
     * @param {string} itemId - Item ID
     * @param {Object} auditContext - Audit context
     * @returns {Promise<Array>} List of worksheets
     */
    async getWorksheets(accessToken, driveId, itemId, auditContext) {
        try {
            // Check workbook access permission
            if (!permissions.canAccessWorkbook(auditContext.user, itemId)) {
                throw new Error('Access denied to workbook');
            }

            const worksheets = await this.graphService.getWorksheets(accessToken, driveId, itemId, auditContext);
            
            // Filter worksheets based on permissions
            const filteredWorksheets = worksheets.filter(worksheet => 
                permissions.canAccessWorksheet(auditContext.user, itemId, worksheet.id)
            );

            return filteredWorksheets;
        } catch (error) {
            logger.error('Excel service - failed to get worksheets:', error);
            throw error;
        }
    }

    /**
     * Read data from Excel range with permission checking
     * @param {Object} params - Parameters
     * @returns {Promise<Object>} Range data
     */
    async readRange(params) {
        const { accessToken, driveId, itemId, worksheetId, range, auditContext } = params;

        try {
            // Check permissions
            const hasPermission = permissions.canReadRange(
                auditContext.user, 
                itemId, 
                worksheetId, 
                range
            );

            if (!hasPermission.allowed) {
                auditService.logPermissionCheck({
                    ...auditContext,
                    workbookId: itemId,
                    worksheetId: worksheetId,
                    range: range,
                    requestedPermission: 'READ',
                    granted: false,
                    reason: hasPermission.reason
                });
                throw new Error(`Read access denied: ${hasPermission.reason}`);
            }

            // Log successful permission check
            auditService.logPermissionCheck({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                range: range,
                requestedPermission: 'READ',
                granted: true
            });

            // Perform the read operation
            const data = await this.graphService.readRange(
                accessToken, driveId, itemId, worksheetId, range, auditContext
            );

            return data;
        } catch (error) {
            logger.error('Excel service - failed to read range:', error);
            throw error;
        }
    }

    /**
     * Write data to Excel range with permission checking
     * @param {Object} params - Parameters
     * @returns {Promise<Object>} Updated range data
     */
    async writeRange(params) {
        const { accessToken, driveId, itemId, worksheetId, range, values, auditContext } = params;

        try {
            // Validate input data
            if (!Array.isArray(values) || values.length === 0) {
                throw new Error('Values must be a non-empty array');
            }

            // Check permissions
            const hasPermission = permissions.canWriteRange(
                auditContext.user, 
                itemId, 
                worksheetId, 
                range
            );

            if (!hasPermission.allowed) {
                auditService.logPermissionCheck({
                    ...auditContext,
                    workbookId: itemId,
                    worksheetId: worksheetId,
                    range: range,
                    requestedPermission: 'WRITE',
                    granted: false,
                    reason: hasPermission.reason
                });
                throw new Error(`Write access denied: ${hasPermission.reason}`);
            }

            // Log successful permission check
            auditService.logPermissionCheck({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                range: range,
                requestedPermission: 'WRITE',
                granted: true
            });

            // Perform the write operation
            const data = await this.graphService.writeRange(
                accessToken, driveId, itemId, worksheetId, range, values, auditContext
            );

            return data;
        } catch (error) {
            logger.error('Excel service - failed to write range:', error);
            throw error;
        }
    }

    /**
     * Read data from Excel table with permission checking
     * @param {Object} params - Parameters
     * @returns {Promise<Object>} Table data
     */
    async readTable(params) {
        const { accessToken, driveId, itemId, worksheetId, tableName, auditContext } = params;

        try {
            // Check permissions
            const hasPermission = permissions.canReadTable(
                auditContext.user, 
                itemId, 
                worksheetId, 
                tableName
            );

            if (!hasPermission.allowed) {
                auditService.logPermissionCheck({
                    ...auditContext,
                    workbookId: itemId,
                    worksheetId: worksheetId,
                    table: tableName,
                    requestedPermission: 'READ_TABLE',
                    granted: false,
                    reason: hasPermission.reason
                });
                throw new Error(`Table read access denied: ${hasPermission.reason}`);
            }

            // Log successful permission check
            auditService.logPermissionCheck({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                requestedPermission: 'READ_TABLE',
                granted: true
            });

            // Perform the read operation
            const data = await this.graphService.readTable(
                accessToken, driveId, itemId, worksheetId, tableName, auditContext
            );

            return data;
        } catch (error) {
            logger.error('Excel service - failed to read table:', error);
            throw error;
        }
    }

    /**
     * Add rows to Excel table with permission checking
     * @param {Object} params - Parameters
     * @returns {Promise<Object>} Result
     */
    async addTableRows(params) {
        const { accessToken, driveId, itemId, worksheetId, tableName, rows, auditContext } = params;

        try {
            // Validate input data
            if (!Array.isArray(rows) || rows.length === 0) {
                throw new Error('Rows must be a non-empty array');
            }

            // Check permissions
            const hasPermission = permissions.canWriteTable(
                auditContext.user, 
                itemId, 
                worksheetId, 
                tableName
            );

            if (!hasPermission.allowed) {
                auditService.logPermissionCheck({
                    ...auditContext,
                    workbookId: itemId,
                    worksheetId: worksheetId,
                    table: tableName,
                    requestedPermission: 'WRITE_TABLE',
                    granted: false,
                    reason: hasPermission.reason
                });
                throw new Error(`Table write access denied: ${hasPermission.reason}`);
            }

            // Log successful permission check
            auditService.logPermissionCheck({
                ...auditContext,
                workbookId: itemId,
                worksheetId: worksheetId,
                table: tableName,
                requestedPermission: 'WRITE_TABLE',
                granted: true
            });

            // Perform the write operation
            const data = await this.graphService.addTableRows(
                accessToken, driveId, itemId, worksheetId, tableName, rows, auditContext
            );

            return data;
        } catch (error) {
            logger.error('Excel service - failed to add table rows:', error);
            throw error;
        }
    }

    /**
     * Validate range format
     * @param {string} range - Range string (e.g., 'A1:C10')
     * @returns {boolean} True if valid
     */
    validateRangeFormat(range) {
        // Basic range validation regex
        const rangeRegex = /^[A-Z]+\d+:[A-Z]+\d+$|^[A-Z]+\d+$|^[A-Z]+:[A-Z]+$|^\d+:\d+$/;
        return rangeRegex.test(range);
    }

    /**
     * Parse range to get dimensions
     * @param {string} range - Range string
     * @returns {Object} Range dimensions
     */
    parseRange(range) {
        try {
            // This is a simplified parser - production would need more robust parsing
            const parts = range.split(':');
            if (parts.length === 1) {
                // Single cell
                return { startCell: parts[0], endCell: parts[0], isSingleCell: true };
            } else if (parts.length === 2) {
                return { startCell: parts[0], endCell: parts[1], isSingleCell: false };
            }
            throw new Error('Invalid range format');
        } catch (error) {
            logger.error('Failed to parse range:', error);
            throw new Error(`Invalid range format: ${range}`);
        }
    }
}

module.exports = new ExcelService();
