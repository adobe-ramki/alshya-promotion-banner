/*
Copyright 2022 Adobe. All rights reserved.
This file is licensed to you under the Apache License, Version 2.0 (the "License")
you may not use this file except in compliance with the License. You may obtain a copy
of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under
the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
OF ANY KIND, either express or implied. See the License for the specific language
governing permissions and limitations under the License.
*/
const { Core } = require('@adobe/aio-sdk')
const { stringParameters, checkMissingRequestInputs } = require('../../../utils')
const { validateData } = require('../delete/validator')
const { HTTP_INTERNAL_ERROR, HTTP_BAD_REQUEST } = require('../../../constants')
const { actionSuccessResponse, actionErrorResponse } = require('../../../responses')
const {
    getSiteId,
    setUtilityLogger,
    getFileItemId,
    getDirectoryPath,
    setFilePathToRead,
    setAccessToken,
    getEntraAccessToken,
    deactivateRow
} = require('../../../../utils/sp-graph-api-util')

/**
 * This action is on charge of deleting staging content of sales rule information in Adobe commerce to external one drive excel sheet
 *
 * @returns {object} returns response object with status code, request data received and response of the invoked action
 * @param {object} params - includes the env params, type and the data of the event
 */
async function main (params) {
    const logger = Core.Logger('sale-rule-commerce-delete', { level: params.LOG_LEVEL || 'info' })
    setUtilityLogger(logger)
    logger.info('Start processing request')
    logger.debug(`Received params: ${stringParameters(params)}`)

    try {
        const dataObject = params.data?.data?.value?.salesRule || {}
        const requiredParams = ['data.website', 'data.schedule_id', 'data.brand']
        const errorMessage = checkMissingRequestInputs({data: dataObject}, requiredParams, [])
        if (errorMessage) {
        logger.error(`Invalid request parameters: ${stringParameters(params)}`)
        return actionErrorResponse(HTTP_BAD_REQUEST, `Invalid request parameters: ${errorMessage}`)
        }
        const validationResult = validateData(dataObject)
        if (validationResult.success === false) {
            return actionErrorResponse(HTTP_BAD_REQUEST, validationResult.message)
        }
        const websiteCodes = dataObject.website.split(',').filter(i => i)
        if (websiteCodes.length === 0 && websiteCodes.length === 0) {
            return actionSuccessResponse("No changes to update")
        }
        const accessToken = await getEntraAccessToken(params)
        setAccessToken(accessToken)
        params.brand = dataObject.brand
        const loadedSiteId = await getSiteId(params)
        const filePathPrefix = `${params.MICROSOFT_GRAPH_BASE_URL}/sites('${loadedSiteId}')/`
        //remove from sheet
        for(let siteCode of websiteCodes) {
            let filePathToRead = await getFileItemId(params, siteCode, filePathPrefix)
            setFilePathToRead(filePathToRead)
            await deactivateRow(dataObject.schedule_id)
        }
        logger.debug('Process finished successfully')
        return actionSuccessResponse('Data synced successfully')
    } catch (error) {
        logger.error(`Error processing the request: ${error.message}`)
        return actionErrorResponse(HTTP_INTERNAL_ERROR, error.message)
    }
}

exports.main = main
