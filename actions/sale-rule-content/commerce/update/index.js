/*
Copyright 2022 Adobe. All rights reserved.
This file is licensed to you under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License. You may obtain a copy
of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under
the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
OF ANY KIND, either express or implied. See the License for the specific language
governing permissions and limitations under the License.
*/
const { Core } = require('@adobe/aio-sdk')
const { stringParameters } = require('../../../utils')
const { validateData } = require('./validator')
const { HTTP_INTERNAL_ERROR, HTTP_BAD_REQUEST } = require('../../../constants')
const { actionSuccessResponse, actionErrorResponse } = require('../../../responses')
const { getEntraAccessToken, updateExcel, removeFromExcel, setUnitilityLogger } = require('../../../../utils/spo-file-update')
/**
 * This action is on charge of sending created/updated staging content of sales rule information in Adobe commerce to external one drive excel sheet
 *
 * @returns {object} returns response object with status code, request data received and response of the invoked action
 * @param {object} params - includes the env params, type and the data of the event
 */
async function main (params) {
  const logger = Core.Logger('product-commerce-consumer', { level: logLevel || 'info' })
  setUnitilityLogger(logger)
  logger.info('Start processing request')
  logger.debug(`Received params: ${stringParameters(params)}`)

  try {
    const dataObject = params?.data?.value?.salesRule || params?.salesRule || params?.data?.salesRule || {};
    const validationResult = validateData(dataObject)
    if (validationResult.success === false) {
        return actionErrorResponse(HTTP_BAD_REQUEST, validationResult.message)
    }
    const oldwebsiteCodes = dataObject.pre_website.split(',');
    const newwebsiteCodes = dataObject.post_website.split(',');
    const brandCode = dataObject.brand;
    let removeFromWebsites = oldwebsiteCodes.filter(x => !newwebsiteCodes.includes(x));
    if (removeFromWebsites.length === 0 && newwebsiteCodes.length === 0) {
        return actionSuccessResponse("No changes to update")
    }
    const accessToken = await getEntraAccessToken();
    const filePathPrefix = `${params.MICROSOFT_GRAPH_BASE_URL}/sites/${params.ENTRA_SITE_ID}/drive/root:/${getDirectoryPath(params, brandCode)}/`;
    //add or update into sheet
    const rowsData = [dataObject];
    for(let siteCode of newwebsiteCodes) {
        let filePathToRead =  filePathPrefix + getFileNameToRead(siteCode);
        await updateExcel(accessToken, rowsData, filePathToRead);
    }
    //remove/deactivate from sheet
    for(let siteCode of removeFromWebsites) {
        let filePathToRead = filePathPrefix + getFileNameToRead(siteCode);
        await removeFromExcel(accessToken, rowsData, filePathToRead);
    }
  
    logger.debug('Process finished successfully')
    return actionSuccessResponse('Data synced successfully')
  } catch (error) {
    logger.error(`Error processing the request: ${error.message}`)
    return actionErrorResponse(HTTP_INTERNAL_ERROR, error.message)
  }
}

exports.main = main
