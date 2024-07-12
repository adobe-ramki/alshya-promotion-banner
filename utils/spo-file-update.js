
const axios = require('axios')
const ExcelJS = require('exceljs')
const { stringParameters } = require('../actions/utils')
const { getEntraAccessToken } = require('./azure-auth')
const { Logger } =  require('./logger');
const utilityLogger = new Logger();

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
 * Find the row in the Excel sheet by the primary column value
 *
 * @param {object} worksheet
 * @param {string|number} candidateID
 * @param columnIndex
 * @returns {object | null}
 */
function findRowByID(worksheet, candidateID, columnIndex) {
    let targetRow = null
    worksheet.eachRow((row, rowNumber) => {
        if (row.getCell(columnIndex).value === candidateID) {
            targetRow = row
            return;
        }
    })
    return targetRow
}

/**
 *
 * @param worksheet
 * @param headerName
 * @returns {Promise<number>}
 */
async function findColumnIndexByHeader(worksheet, headerName) {
    const headerRow = worksheet.getRow(1); // Assuming headers are in the first row
    let columnIndex = -1;
    headerRow.eachCell((cell, colNumber) => {
        if (cell.value === headerName) {
            columnIndex = colNumber;
        }
    });
    return columnIndex;
}

/**
 * Set the cell value from the row
 * 
 * @param {object} row 
 * @param {number} cellNumber
 * @param {any} value 
 */
function setCellValue(row, cellNumber, value) {
    const cell = row.getCell(cellNumber)
    cell.value = value
}

/**
 * Parse boolean value to integer
 * @param {string | boolean} value 
 * @returns {number}
 */
function parseBoolToInt(value) {
    if (typeof(value) === 'string'){
        value = value.trim().toLowerCase();
    }
    switch(value){
        case true:
        case "true":
        case 1:
        case "1":
        case "on":
        case "yes":
            return 1;
        default:
            return 0;
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
 * Download file from OneDrive for modification
 * 
 * @param {String} accessToken 
 * @param {String} filePathToRead 
 * @returns {any}
 */
async function downloadFileFromOneDrive(accessToken, filePathToRead) {
    const endpoint = filePathToRead+`?$select=@microsoft.graph.downloadUrl`
    const headers = {
        'Authorization': `Bearer ${accessToken}`
    }
    try {
        const response = await axios.get(endpoint, { headers: headers})
        const downloadUrl = response.data['@microsoft.graph.downloadUrl']
        const fileResponse = await axios.get(downloadUrl, { responseType: 'arraybuffer' })
        return fileResponse.data
    } catch (error) {
        utilityLogger.debug(`Error: ${stringParameters(error)}`)
        return []
    }
}

/**
 * Upload file to OneDrive
 * 
 * @param {string} accessToken 
 * @param {any} fileData 
 * @param {string} filePathToRead 
 * @returns {Promise<void>}
 */
async function uploadFileToOneDrive(accessToken, fileData, filePathToRead) {
    const headers = {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Prefer': 'bypass-shared-lock'
    }
    const originalEndpoint = filePathToRead+`:/content`
    retryFileUpload(fileData, headers, originalEndpoint)
}

/**
 * Upload file to Sharepoint and retry if file is locked until 30 seconds in the interval of 500ms
 * 
 * @param {any} fileData 
 * @param {object} requestHeaders 
 * @param {string} originalEndpoint 
 */
function retryFileUpload(fileData, requestHeaders, originalEndpoint) {
    let count = 0
    const delayInMiliSec = 500
    const tryUptoSec = 30
    const maxRetries = ((tryUptoSec * 1000) / delayInMiliSec)
    const retryInterval = setInterval(
        () => {
            void( 
                async() => {
                    let shouldRetry = false
                    try {
                        await axios.put(originalEndpoint, fileData, { headers: requestHeaders })
                    } catch (error) {
                        if (error.response && error.response.data && error.response.data.error && error.response.data.error.message.includes("locked")) {
                            utilityLogger.info("Current file is locked. Retrying...")
                            shouldRetry = true
                        } else {
                            utilityLogger.debug(`Error uploading the file:${stringParameters(error?.response?.data)}`)
                        }
                    }
                    count++; 
                    if(count===maxRetries || !shouldRetry) {
                        clearInterval(retryInterval);
                    } 
                })();
            }, 
            delayInMiliSec
        );
}

/**
 * Put file to OneDrive return true if success, false otherwise
 * 
 * @param {any} fileData 
 * @param {object} headers 
 * @param {string} filePathToRead 
 * @returns {boolean}
 */
async function putFileToOneDrive(fileData, headers, filePathToRead) {
    try {
        await axios.put(filePathToRead, fileData, { headers: headers })
    } catch (uploadError) {
        utilityLogger.debug(`Error uploading the file:${stringParameters(uploadError.response.data)}`)
        return false
    }
    return true
}

/**
 * Delete file from OneDrive: Use this only if required and file is locked
 * Return true if success, false otherwise
 * 
 * @param {object} headers 
 * @param {string} filePathToRead 
 * @returns {boolean}
 */
async function deleteOnlyIfLocked(headers, filePathToRead) {
    try {
        await axios.delete(filePathToRead, { headers: headers })
    } catch (deleteError) {
        utilityLogger.debug(`Error deleting the locked file:${stringParameters(deleteError.response.data)}`)
        return false
    }
    return true
}

/**
 * Load Excel workbook from file data
 * 
 * @param {buffer} fileData
 * @returns {object}
 */
async function loadExcelWorkBook(fileData)
{
    const workbook = new ExcelJS.Workbook()
    return await workbook.xlsx.load(fileData)
}

/**
 * Upload Excel worksheet to OneDrive
 * 
 * @param {object} workbook 
 * @param {string} filePathToRead 
 * @param {string} accessToken 
 * @returns {Promise<void>}
 */
async function uploadExcelWorkBook(workbook, filePathToRead, accessToken)
{
    const buffer = await workbook.xlsx.writeBuffer()
    await uploadFileToOneDrive(accessToken, buffer, filePathToRead)
    utilityLogger.info(`Successfully saved to OneDrive`)
}

/**
 * Update Excel sheet with new data or create new rows
 * 
 * @param {string} accessToken 
 * @param {array} items 
 * @param {string} filePathToRead 
 * @returns {Promise<void>}
 */
async function updateExcel(accessToken, items, filePathToRead) {
    const fileData = await downloadFileFromOneDrive(accessToken, filePathToRead)
    if (!fileData || fileData.length === 0) {
        utilityLogger.info("Error. Something went wrong...")
        return
    }
    utilityLogger.info("OneDrive data fetched!")
    const workbook = await loadExcelWorkBook(fileData)
    const worksheet = workbook.getWorksheet(1)
    const keysToUpdate = getSheetColumnsToUpdate()
    for (const elem of items) {
        const row = findRowByID(worksheet, elem.schedule_id)
        if (row) {
            utilityLogger.info("..updating row")
            keysToUpdate.forEach((key, index) => {
                setCellValue(row, key, elem[key])
            })
        } else {
            utilityLogger.info("..adding row")
            let data = []
            keysToUpdate.forEach((key, index) => {
                data.push((key === 'status'? parseBoolToInt(elem[key]): elem[key]))
            })
            const newRow = worksheet.addRow(data)
        }
    }
    try {
        await uploadExcelWorkBook(workbook, filePathToRead, accessToken)
    } catch (error) {
        utilityLogger.debug(`Error writing to file:${stringParameters(error)}`)
    }
}

/**
 * Deactivate rows from Excel sheet
 * 
 * @param {array} items 
 * @param {string} filePathToRead 
 * @returns {Promise<void>}
 */
async function removeFromExcel(accessToken, items, filePathToRead) {
    const fileData = await downloadFileFromOneDrive(accessToken, filePathToRead)
    if (!fileData || fileData.length === 0) {
        utilityLogger.info("Error. Something went wrong...")
        return
    }
    utilityLogger.info("OneDrive data fetched!")
    const workbook = await loadExcelWorkBook(fileData)
    const worksheet = workbook.getWorksheet(1)
    const scheduleIdColumnIndex = await findColumnIndexByHeader(worksheet, 'schedule_id');
    if (scheduleIdColumnIndex === -1) {
        throw new Error('Column "schedule_id" or "status" not found');
    }
    
    for (let elem of items) {
        const row = findRowByID(worksheet, elem.schedule_id, scheduleIdColumnIndex)
        if (row) {
            utilityLogger.info("..deleting row")
            worksheet.spliceRows(row._number, 1);
        } else {
            utilityLogger.info("Can not find row to delete")
        }
    }
    try {
        await uploadExcelWorkBook(workbook, filePathToRead, accessToken)

    } catch (error) {
        utilityLogger.debug(`Error writing to file:${stringParameters(error)}`)
    }
}

module.exports = {
    getEntraAccessToken,
    updateExcel,
    removeFromExcel,
    setUtilityLogger,
    getFileNameToRead,
    getDirectoryPath
}