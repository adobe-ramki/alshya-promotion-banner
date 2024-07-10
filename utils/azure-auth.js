const axios = require('axios');
/**
 * Generate access token to connect with ENTRA->Sharepoint
 * 
 * @param {object} params includes env parameters
 * @returns {string} returns the access token
 * @throws {Error} in case of any failure
 */
async function getEntraAccessToken (params)
{
    const tokenEndpoint = params.ENTRA_TOKEN_ENDPOINT.replace('{{Tenent_ID}}', params.ENTRA_TENANT_ID);
    const requestData = {
        client_id: params.ENTRA_CLIENT_ID,
        scope: params.ENTRA_AUTH_SCOPE,
        client_secret: params.ENTRA_CLIENT_SECRET,
        grant_type: params.ENTRA_AUTH_GRANT_TYPE
    };
    const response = await axios.post(tokenEndpoint, new URLSearchParams(requestData));
    return response.data.access_token;
}
module.exports = {
    getEntraAccessToken
}