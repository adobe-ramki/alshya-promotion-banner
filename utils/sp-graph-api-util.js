const axios = require('axios')
const { stringParameters } = require('../actions/utils')
const { getEntraAccessToken } = require('./azure-auth')
const { Logger } =  require('./logger')
const brandMappingJson = require('./config/brand-mapping.json')
const storeCodeMappingJson = require('./config/store-code-mapping.json')
const utilityLogger = new Logger()
let loadedSiteId = null, loadedHeaderRow = {}, loadedStore = {}, loadedWorkSheetId = null, loadedTableId = null, loadedAccessToken = null, loadedFilePath = null

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
 * @param {string} storeCode
 * @returns {string}
 */
function getDirectoryPath(params, storeCode) {
    storeCode = storeCodeMappingJson[storeCode] || storeCode
    if (typeof storeCode === 'object') {
        loadedStore = storeCode
        storeCode = storeCode.code
    }
    return params.SHAREPOINT_DIRECTORY_PATH_FROM_ROOT + '/' + storeCode
}

/**
 * Set the site id for the SharePoint
 * 
 * @param {object} params 
 * @returns 
 */
async function getSiteId(params)
{
    if (loadedSiteId) {
        return loadedSiteId
    }

    if (!params.brand) {
        utilityLogger.debug('Brand is not set in the params. Please set the brand before using it.')
        throw new Error('Brand is not set in the params. Please set the brand before using it.')
    }
    const brandPath = brandMappingJson[params.brand] || params.brand
    const urlKey = brandPath.urlKey || brandPath
    const brandSiteId = brandPath.siteId || null
    if (brandSiteId) {
        loadedSiteId = brandSiteId
        return loadedSiteId
    }
    const requestHeaders = {
        'Authorization': `Bearer ${getAccessToken()}`, 
        'Content-Type': 'application/json',
    }
    const apiUrl =  `${params.MICROSOFT_GRAPH_BASE_URL}/sites/${params.SHAREPOINT_HOST_NAME}:/sites/AXP/${urlKey}?$select=id`
    const response = await axios.get(apiUrl, { headers: requestHeaders })
    const siteId = response.data?.id || null
    if (!siteId) {
        utilityLogger.debug(`Site id not found in the response. Response ${stringParameters(response)}`)
        throw new Error('Site id not found in the response. Please check the site brand path mapping.')
    }
    loadedSiteId = siteId
    return loadedSiteId
}

/**
 * Get file id from SharePoint
 * 
 * @param {string} filePath 
 * @returns {string}
 */
async function getFileIdFromSharePoint(filePath) {
    const requestHeaders = {
        'Authorization': `Bearer ${getAccessToken()}`,
        'Content-Type': 'application/json',
    }
    const apiUrl = `${filePath}?$select=id`
    const response = await axios.get(apiUrl, { headers: requestHeaders })
    const fileId = response.data?.id || null
    if (!fileId) {
        utilityLogger.debug(`File id not found in the response. Response ${stringParameters(response)}`)
        throw new Error('File id not found in the sharepoint. Please check the file path.')
    }
    return fileId
}

/**
 * Get file item id from SharePoint
 * 
 * @param {object} params 
 * @param {string} siteCode 
 * @param {string} filePathPrefix 
 * @returns 
 */
async function getFileItemId(params, siteCode, filePathPrefix) {
    const getFileNameToRead = params.FILE_NAME_TO_READ
    let fileDirPrefix = filePathPrefix + `drive/root:/` + getDirectoryPath(params, siteCode) + '/'
    const itemId = await getFileIdFromSharePoint(fileDirPrefix + getFileNameToRead)
    return filePathPrefix + `/drive/items/${itemId}`
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

    if (typeof loadedStore === 'object' && typeof loadedStore.sheetId !== 'undefined' && loadedStore.sheetId !== '') {
        loadedWorkSheetId = loadedStore.sheetId
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
    if (loadedHeaderRow.length > 0) {
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
    const colValues = response.data?.values || []
    colValues.shift()
    const rowIndex = colValues.findIndex((element)=> parseInt(element[0]) === parseInt(valueToMatch))
    return rowIndex + 1
}

/**
 * Get row data by index from SharePoint Shet
 * 
 * @param {number} rowIndex 
 * @returns 
 */
async function getRowDataByIndex(rowIndex) {
    if (rowIndex < 1) {
        return null
    }
    const getTableId = await getFirstTable()
    const getWorksheetId = await getFirstActiveWorksheetId()
    const apiUrl = getFilePathToRead() + `/workbook/worksheets/${getWorksheetId}/tables/${getTableId}/rows/itemAt(index=${rowIndex})?$select=values`
    const response = await axios.get(apiUrl, { headers: {
            'Authorization': `Bearer ${getAccessToken()}`,
            'Content-Type': 'application/json',
        }
    })
    return response.data?.values[0] || null
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

/**
 * Save row into SharePoint table
 * 
 * @param {object} jsonData
 * @param {null | number} rowIndex
 */
async function saveRowData(jsonData, rowIndex = null) {
    let hasSaved = false
    const getWorksheetId = await getFirstActiveWorksheetId()
    const getTableId = await getFirstTable()
    const headers = { headers: {
            'Authorization': `Bearer ${getAccessToken()}`,
            'Content-Type': 'application/json',
        }
    }
    const postData = prepareDataForUpdate(jsonData)
    if (getSheetColumnsToUpdate().length !== postData.length) {
        throw new Error("Number of columns mismatched in post data")
    }
    const apiUrl = getFilePathToRead() + `/workbook/worksheets/${getWorksheetId}/tables/${getTableId}/rows`
    if (rowIndex === null) {
        const response = await axios.post(apiUrl, {
            values: [postData]
        }, headers)
        hasSaved = response.status === 201 ? true: false
    } else {
        apiUrl = apiUrl + `/itemAt(index=${rowIndex})`
        const response = await axios.patch(apiUrl, {
            values: [postData]
        }, headers)
        hasSaved = response.status === 200 ? true: false
    }
    return hasSaved ? true : false
}

/**
 * Prepare Data before updating
 * 
 * @param {object} elem 
 * @returns 
 */
function prepareDataForUpdate(elem) {
    let data = []
    const keysToUpdate = getSheetColumnsToUpdate()
    keysToUpdate.forEach((key, index) => {
        if (typeof elem[key] !== 'undefined') {
            data.push((key === 'status'? parseBoolToInt(elem[key]): elem[key]))
        }
    })
    return data
}

/**
 * Create or update rows in SharePoint sheet
 * 
 * @param {object} jsonData 
 * @returns 
 */
async function createOrUpdateRows(jsonData) {
    // check if row exists
    const scheduleId = parseInt(jsonData.schedule_id)
    const rowId = await findRowIndexByID('schedule_id', scheduleId)
    if (rowId > 0) {
        return await saveRowData(jsonData, rowId)
    } else {
        return await saveRowData(jsonData)
    }
}

/**
 * Delete row from SharePoint sheet
 * 
 * @param scheduleId
 * @returns {boolean}
 */
async function deleteRow(scheduleId) {
    const headers = { headers: {
            'Authorization': `Bearer ${getAccessToken()}`,
            'Content-Type': 'application/json',
        }
    }
    const getWorksheetId = await getFirstActiveWorksheetId()
    const getTableId = await getFirstTable()
    scheduleId = parseInt(scheduleId)
    const rowId = await findRowIndexByID('schedule_id', scheduleId)
    if (rowId < 0) {
        utilityLogger.info(`Schedule id ${scheduleId} does not exist`)
    }
    const apiUrl = getFilePathToRead() + `/workbook/worksheets/${getWorksheetId}/tables/${getTableId}/rows/${rowId}`
    const response = await axios.delete(apiUrl, headers)
    utilityLogger.info(`Deletion for schedule_id ${scheduleId}, ${stringParameters(response)}`)
    return response.status === 200 ? true: false
}

/**
 * Deactivate row from SharePoint sheet by changing status to 0
 * 
 * @param scheduleId
 * @returns {boolean}
 */
async function deactivateRow(scheduleId) {
    scheduleId = parseInt(scheduleId)
    const rowId = await findRowIndexByID('schedule_id', scheduleId)
    if (rowId < 0) {
        throw Error(`Invalid schedule id provided no entries found scheduleId -> ${scheduleId}`)
    }
    let rowValues = await getRowDataByIndex(rowId)
    const getWorksheetId = await getFirstActiveWorksheetId()
    const getTableId = await getFirstTable()
    const headers = { headers: {
            'Authorization': `Bearer ${getAccessToken()}`,
            'Content-Type': 'application/json',
        }
    }
    rowValues[await findColumnIndexByHeader('status')] = 0
    const apiUrl = getFilePathToRead() + `/workbook/worksheets/${getWorksheetId}/tables/${getTableId}/rows/itemAt(index=${rowId})`
    const response = await axios.patch(apiUrl, {
        values: [rowValues]
    }, headers)
    return response.status === 200 ? true: false
}

module.exports = {
    getEntraAccessToken,
    setUtilityLogger,
    getDirectoryPath,
    setAccessToken,
    setFilePathToRead,
    getSiteId,
    createOrUpdateRows,
    deleteRow,
    deactivateRow,
    getFileItemId
}