const axios = require('axios')
const { stringParameters } = require('../actions/utils')
const { getEntraAccessToken } = require('./azure-auth')
const { Logger } =  require('./logger')
const utilityLogger = new Logger()
let loadedHeaderRow = null, loadedWorkSheetId = null, loadedTableId = null, loadedAccessToken = null, loadedFilePath = null

/**
 * Set file path to read from SharePoint
 * 
 * @param {string} filePath 
 */
function setFilePathToRead(filePath) {
    loadedFilePath = filePath
}

/**
 * Set access token for the api calls
 * 
 * @param {string} accessToken 
 */
function setAccessToken(accessToken) {
    loadedAccessToken = accessToken
}

/**
 * Get access token for the api calls
 * 
 * @returns {string}
 * @throws {Error} if access token is not set
 */
function getAccessToken() {
    if (!loadedAccessToken) {
        throw new Error('Access token is not set yet. Please set the access token before using it.')
    }
    return loadedAccessToken
}

/**
 * Get file path to read from SharePoint
 * 
 * @returns {string}
 * @throws {Error} if file path is not set
 */
function getFilePathToRead() {
    if (!loadedFilePath) {
        throw new Error('File path is not set yet. Please set the file path before using it.')
    }
    return loadedFilePath
}

/**
 * Set the logger instance for the utility
 * 
 * @param {object} logger
 * @returns {void} 
 */
function setUtilityLogger(logger) {
    utilityLogger.setLoggerInstance(logger)
}

/**
 * Get the directory full path for SharePoint
 * 
 * @param {array} params
 * @param {string} contentDirName
 * @returns {string}
 */
function getDirectoryPath(params, contentDirName) {
    return params.SHAREPOINT_DIRECTORY_PATH_FROM_ROOT + '/' + contentDirName + '_promotions'
}

/**
 * Get the file name to read from SharePoint
 * 
 * @param {string|null} siteCode
 * @returns {string}
 */
function getFileNameToRead(siteCode = null) {
    return `${siteCode}-promotions.xlsx`
}

/**
 * Get first active worksheet id (in alpha-numerical form)
 * 
 * @returns {string}
 * @throws {Error} if no worksheet found in the excel sheet or api call failed
 */
async function getFirstActiveWorksheetId()
{   
    if (loadedWorkSheetId) {
        return loadedWorkSheetId
    }

    const requestHeaders = {
        'Authorization': `Bearer ${getAccessToken()}`,
        'Content-Type': 'application/json'
    }
    const response = await axios.get(getFilePathToRead() + `/workbook/worksheets/?$select=id,visibility`, { headers: requestHeaders })
    let worksheets =  response.data?.value || []
    if (worksheets.length === 0) {
        utilityLogger.debug(`No worksheet found in the excel sheet or api call failed. Response ${stringParameters(response)}`)
        throw new Error('No worksheet found in the excel sheet or api call failed. Please check the excel sheet and try again.')
    }
    worksheets = worksheets.filter(worksheet => worksheet.visibility === 'Visible')
    loadedWorkSheetId = worksheets[0].id
    return loadedWorkSheetId
}

/**
 * Get first table id from the worksheet
 * 
 * @returns {string}
 * @throws {Error} if no table found in the excel sheet or api call failed
 */
async function getFirstTable()
{
    if (loadedTableId) {
        return loadedTableId
    }
    const getWorksheetId = await getFirstActiveWorksheetId()
    const requestHeaders = {
        'Authorization': `Bearer ${getAccessToken()}`,
        'Content-Type': 'application/json',
    }
    const response = await axios.get(getFilePathToRead() + `/workbook/worksheets/${getWorksheetId}/tables?$select=id`, { headers: requestHeaders })
    let tables = response.data?.value || []
    if (tables.length === 0) {
        utilityLogger.debug(`No table found in the excel sheet or api call failed. Response ${stringParameters(response)}`)
        throw new Error('No table found in the excel sheet or api call failed. Please check the excel sheet and try again.')
    }
    loadedTableId =  tables[0].id
    return loadedTableId
}

/**
 * Get header names from the excel sheet along with column index
 * 
 * @returns {array}
 * @throws {Error} if no columns found in the excel sheet or api call failed
 */
async function getHeaderRow() {
    if (loadedHeaderRow) {
        return loadedHeaderRow
    }
    const getTableId = await getFirstTable()
    const getWorksheetId = await getFirstActiveWorksheetId()
    const requestHeaders = {
        'Authorization': `Bearer ${getAccessToken()}`,
        'Content-Type': 'application/json',
    }
    const apiUrl = getFilePathToRead() + `/workbook/worksheets/${getWorksheetId}/tables/${getTableId}/columns?$select=index,name`
    const response = await axios.get(apiUrl, { headers: requestHeaders })
    const columns = response.data?.value || []
    if (columns.length === 0) {
        utilityLogger.debug(`No columns found in the excel sheet or api call failed. Response ${stringParameters(response)}`)
        throw new Error('No columns found in the excel sheet or api call failed. Please check the excel sheet and try again.')
    }
    columns.forEach((element)=> {
        loadedHeaderRow[element.name] = element.index
    })
    return loadedHeaderRow
}

/**
 * Find column index by header name in the excel sheet
 * 
 * @param {string} headerName 
 * @returns {number}
 */
async function findColumnIndexByHeader(headerName) {
    const loadedHeaderRow = await getHeaderRow()
    return loadedHeaderRow[headerName] || -1
}

/**
 * Find row index by column name and value in the excel sheet
 * 
 * @param {string} colName 
 * @param {any} valueToMatch 
 * @returns {number}
 * @throws {Error} if column not found in the excel sheet
 */
async function findRowIndexByID(colName, valueToMatch) {
    const getTableId = await getFirstTable()
    const getWorksheetId = await getFirstActiveWorksheetId()
    const loadedHeaderRow = await getHeaderRow()
    const columnIndex = loadedHeaderRow[colName]
    if (columnIndex === -1) {
        utilityLogger.debug(`Column ${colName} not found in the excel sheet`)
        throw new Error(`Column ${colName} not found in the excel sheet`)
    }
    const apiUrl = getFilePathToRead() + `/workbook/worksheets/${getWorksheetId}/tables/${getTableId}/columns/itemAt(index=${columnIndex})?$select=values`
    const response = await axios.get(apiUrl, { headers: {
            'Authorization': `Bearer ${getAccessToken()}`,
            'Content-Type': 'application/json',
        } 
    })
    const colValues = response.data?.value || []
    const rowIndex = colValues.findIndex((element)=> parseInt(element[0]) === parseInt(valueToMatch))
    return rowIndex - 1
}

/**
 * Parse boolean value to integer
 * 
 * @param {string | boolean} value 
 * @returns {number}
 */
function parseBoolToInt(value) {
    if (typeof(value) === 'string'){
        value = value.trim().toLowerCase()
    }
    switch(value){
        case true:
        case "true":
        case 1:
        case "1":
        case "on":
        case "yes":
            return 1
        default:
            return 0
    }
}


/**
 * Get Excel sheet columns to update
 * 
 * @returns array
 */
function getSheetColumnsToUpdate() {
    return [
        'schedule_id',
        'rule_id',
        'rule_name',
        'coupon_type',
        'description_en',
        'description_ar',
        'short_terms_and_conditions_en',
        'short_terms_and_conditions_ar',
        'url_key',
        'channel_web',
        'channel_app',
        'start_date',
        'end_date',
        'status',
    ]
}

module.exports = {
    getEntraAccessToken,
    setUtilityLogger,
    getFileNameToRead,
    getDirectoryPath,
    setAccessToken,
    setFilePathToRead
}