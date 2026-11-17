const encodeUriComponent = require('encodeUriComponent');
const Firestore = require('Firestore');
const getAllEventData = require('getAllEventData');
const getRequestHeader = require('getRequestHeader');
const getTimestampMillis = require('getTimestampMillis');
const getType = require('getType');
const JSON = require('JSON');
const makeString = require('makeString');
const Math = require('Math');
const sendHttpRequest = require('sendHttpRequest');
const sha256Sync = require('sha256Sync');

/*==============================================================================
==============================================================================*/

const eventData = getAllEventData();
const useOptimisticScenario = isUIFieldTrue(data.useOptimisticScenario);

if (shouldExitEarly(data, eventData)) {
  return data.gtmOnSuccess();
}

if (!hasConversionIdentifier(data)) {
  return data.gtmOnSuccess();
}

const firebaseOptions = {};
if (data.firebaseProjectId) firebaseOptions.projectId = data.firebaseProjectId;

if (data.authFlow === 'stape') {
  sendConversionRequest(
    getData('StapeAccessToken', 'StapeDeveloperToken'),
    handleStapeResponse
  );
} else {
  Firestore.read(data.firebasePath, firebaseOptions).then(
    (result) =>
      sendConversionRequest(
        getData(result.data.access_token, data.developerToken),
        getHandleResponse(result.data.refresh_token, false)
      ),
    () => updateAccessToken(data.refreshToken)
  );
}

if (useOptimisticScenario) {
  return data.gtmOnSuccess();
}

/*==============================================================================
Vendor related functions
==============================================================================*/

function sendConversionRequest(postBody, responseCallback) {
  sendHttpRequest(
    getRequestUrl(),
    (statusCode, headers, body) => {
      responseCallback(statusCode, headers, body);
    },
    { headers: getConversionRequestHeaders(), method: 'POST' },
    postBody
  );
}

function handleStapeResponse(statusCode, headers, body) {
  finishTag(isSuccessfulResponse(statusCode, body));
}

function getHandleResponse(refreshToken, isRetryWithFreshToken) {
  return (statusCode, headers, body) => {
    if (!isRetryWithFreshToken && isAuthenticationError(statusCode, body)) {
      return updateAccessToken(refreshToken);
    }

    finishTag(isSuccessfulResponse(statusCode, body));
  };
}

function isAuthenticationError(statusCode, body) {
  if (statusCode === 401) return true;

  const responseBody = makeString(body || '');
  return [
    'AuthenticationTokenExpired',
    'Authentication token expired',
    'InvalidCredentials'
  ].some((authErrorMarker) => responseBody.indexOf(authErrorMarker) !== -1);
}

function isSuccessfulResponse(statusCode, body) {
  if (statusCode < 200 || statusCode >= 400) return false;

  // The API answers some errors with HTTP 200: a SOAP fault, or a BatchError in PartialErrors.
  const responseBody = makeString(body || '');
  return responseBody.indexOf(':Fault>') === -1 && responseBody.indexOf('<BatchError') === -1;
}

function getConversionRequestHeaders() {
  return {
    'Content-Type': 'text/xml; charset=utf-8',
    SOAPAction: 'ApplyOfflineConversions'
  };
}

function updateAccessToken(storedRefreshToken) {
  // The refresh token stored in Firestore goes stale when the user re-authorises and pastes a new
  // one in the UI, so the UI token is tried when the stored one fails.
  const refreshTokens = [];
  if (storedRefreshToken) refreshTokens.push(storedRefreshToken);
  if (data.refreshToken && refreshTokens.indexOf(data.refreshToken) === -1) {
    refreshTokens.push(data.refreshToken);
  }

  requestAccessToken(refreshTokens);
}

function requestAccessToken(refreshTokens) {
  if (refreshTokens.length === 0) return finishTag(false);

  const authUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token/';
  const authBody =
    'refresh_token=' +
    enc(refreshTokens[0]) +
    '&client_id=' +
    enc(data.clientId) +
    '&client_secret=' +
    enc(data.clientSecret) +
    '&grant_type=refresh_token&scope=https%3A%2F%2Fads.microsoft.com%2Fmsads.manage+offline_access&tenant=common';

  sendHttpRequest(
    authUrl,
    (statusCode, headers, body) => {
      const tokens = safeJsonParse(body);
      if (statusCode < 200 || statusCode >= 400 || !tokens || !tokens.access_token) {
        return requestAccessToken(refreshTokens.slice(1));
      }

      Firestore.write(data.firebasePath, tokens, firebaseOptions).then(
        () => {
          sendConversionRequest(
            getData(tokens.access_token, data.developerToken),
            getHandleResponse(tokens.refresh_token, true)
          );
        },
        () => finishTag(false)
      );
    },
    {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      method: 'POST'
    },
    authBody
  );
}

function getRequestUrl() {
  if (data.authFlow === 'own') {
    const apiVersion = '13';

    return (
      'https://campaign.api.bingads.microsoft.com/Api/Advertiser/CampaignManagement/V' +
      apiVersion +
      '/CampaignManagementService.svc?singleWsdl'
    );
  }

  const containerKey = data.containerKey.split(':');
  const containerZone = containerKey[0];
  const containerIdentifier = containerKey[1];
  const containerApiKey = containerKey[2];
  const containerDefaultDomainEnd = containerKey[3] || 'io';

  return (
    'https://' +
    enc(containerIdentifier) +
    '.' +
    enc(containerZone) +
    '.stape.' +
    enc(containerDefaultDomainEnd) +
    '/stape-api/' +
    enc(containerApiKey) +
    '/v2/microsoft-ads/auth-proxy'
  );
}

function getMicrosoftClickId(clickId) {
  if (!clickId) return clickId;

  // The UET _uetmsclkid cookie stores the click ID as "_uet" + MSCLKID.
  const cookiePrefix = '_uet';
  const value = makeString(clickId).trim();
  return value.indexOf(cookiePrefix) === 0 ? value.substring(cookiePrefix.length) : value;
}

function hasConversionIdentifier(data) {
  return !!(
    getMicrosoftClickId(data.microsoftClickId) ||
    hashData('hashedEmailAddress', data.hashedEmailAddress) ||
    hashData('hashedPhoneNumber', data.hashedPhoneNumber)
  );
}

function getData(accessToken, developerToken) {
  const email = hashData('hashedEmailAddress', data.hashedEmailAddress);
  const phone = hashData('hashedPhoneNumber', data.hashedPhoneNumber);
  const value = data.conversionValue || eventData.value;
  const conversionCurrencyCode =
    data.conversionCurrencyCode || eventData.currencyCode || eventData.currency || '';
  const externalAttributionCredit = data.externalAttributionCredit;
  const externalAttributionModel = data.externalAttributionModel;
  const microsoftClickId = getMicrosoftClickId(data.microsoftClickId);

  // prettier-ignore
  return '' +
    '<s:Envelope xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">\n' +
    '    <s:Header xmlns="https://bingads.microsoft.com/CampaignManagement/v13">\n' +
    '        <Action mustUnderstand="1">ApplyOfflineConversions</Action>\n' +
    '        <AuthenticationToken i:nil="false">' + accessToken + '</AuthenticationToken>\n' +
    '        <CustomerAccountId i:nil="false">' + data.customerAccountId + '</CustomerAccountId>\n' +
    '        <CustomerId i:nil="false">' + data.customerId + '</CustomerId>\n' +
    '        <DeveloperToken i:nil="false">' + developerToken + '</DeveloperToken>\n' +
    '    </s:Header>\n' +
    '    <s:Body>\n' +
    '        <ApplyOfflineConversionsRequest xmlns="https://bingads.microsoft.com/CampaignManagement/v13">\n' +
    '            <OfflineConversions i:nil="false">\n' +
    '                <OfflineConversion>\n' +
    '                    <ConversionCurrencyCode i:nil="false">'+conversionCurrencyCode+'</ConversionCurrencyCode>\n' +
    '                    <ConversionName i:nil="false">'+data.conversionName+'</ConversionName>\n' +
    '                    <ConversionTime>'+(data.conversionTime ? data.conversionTime : getConversionDateTime())+'</ConversionTime>\n' +
                         (value ? '<ConversionValue i:nil="false">'+value+'</ConversionValue>\n' : '') +
                         (externalAttributionCredit ? '<ExternalAttributionCredit i:nil="false">'+externalAttributionCredit+'</ExternalAttributionCredit>\n' : '') +
                         (externalAttributionModel ? '<ExternalAttributionModel i:nil="false">'+externalAttributionModel+'</ExternalAttributionModel>\n' : '') +
                         (email ? '<HashedEmailAddress i:nil="false">'+email+'</HashedEmailAddress>\n' : '') +
                         (phone ? '<HashedPhoneNumber i:nil="false">'+phone+'</HashedPhoneNumber>\n' : '') +
                         (microsoftClickId ? '<MicrosoftClickId i:nil="false">'+microsoftClickId+'</MicrosoftClickId>\n' : '') +
    '                </OfflineConversion>\n' +
    '            </OfflineConversions>\n' +
    '        </ApplyOfflineConversionsRequest>\n' +
    '    </s:Body>\n' +
    '</s:Envelope>';
}

/*==============================================================================
Helpers
==============================================================================*/

function shouldExitEarly(data, eventData) {
  if (!isConsentGivenOrNotRequired(data, eventData)) return true;

  const url = getUrl(eventData);
  if (url && url.lastIndexOf('https://gtm-msr.appspot.com/', 0) === 0) return true;

  return false;
}

function getUrl(eventData) {
  return eventData.page_location || getRequestHeader('referer') || eventData.page_referrer;
}

function isConsentGivenOrNotRequired(data, eventData) {
  if (data.adStorageConsent !== 'required') return true;
  if (eventData.consent_state) return !!eventData.consent_state.ad_storage;
  const xGaGcs = eventData['x-ga-gcs'] || ''; // x-ga-gcs is a string like "G110"
  return xGaGcs[2] === '1';
}

function finishTag(isSuccess) {
  if (useOptimisticScenario) return;
  if (isSuccess) data.gtmOnSuccess();
  else data.gtmOnFailure();
}

function isUIFieldTrue(field) {
  return [true, 'true'].indexOf(field) !== -1;
}

function isValidValue(value) {
  const valueType = getType(value);
  if (valueType === 'null' || valueType === 'undefined' || value !== value) return false;
  return value !== '' && value !== 'undefined' && value !== 'null';
}

function safeJsonParse(body) {
  if (getType(body) !== 'string') return undefined;

  const trimmedBody = body.trim();
  const firstChar = trimmedBody.charAt(0);
  const lastChar = trimmedBody.charAt(trimmedBody.length - 1);
  const looksLikeJson =
    (firstChar === '{' && lastChar === '}') || (firstChar === '[' && lastChar === ']');
  if (!looksLikeJson) return undefined;

  return JSON.parse(trimmedBody);
}

function enc(data) {
  if (['null', 'undefined'].indexOf(getType(data)) !== -1) data = '';
  return encodeUriComponent(makeString(data));
}

function getConversionDateTime() {
  return convertTimestampToISO(getTimestampMillis());
}

function isHashed(value) {
  if (!value) {
    return false;
  }

  return makeString(value).match('^[A-Fa-f0-9]{64}$') !== null;
}

function hashData(key, value) {
  // The API takes a single email and phone per conversion: use the first valid item of an array.
  if (getType(value) === 'array') value = value.filter(isValidValue)[0];
  if (!isValidValue(value) || getType(value) === 'object') return undefined;

  if (isHashed(value)) return makeString(value);

  value = makeString(value).trim().toLowerCase();

  if (key === 'hashedPhoneNumber') {
    value = value.split(' ').join('').split('-').join('').split('(').join('').split(')').join('');
  }

  if (!value) return undefined;

  return sha256Sync(value, { outputEncoding: 'hex' });
}

function convertTimestampToISO(timestamp) {
  const msPerDay = 86400000;
  const isLeapYear = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const pad = (value, length) => {
    let result = makeString(value);
    while (result.length < length) result = '0' + result;
    return result;
  };

  let days = Math.floor(timestamp / msPerDay);
  let msOfDay = timestamp - days * msPerDay;

  let year = 1970;
  while (days >= (isLeapYear(year) ? 366 : 365)) {
    days = days - (isLeapYear(year) ? 366 : 365);
    year++;
  }

  const daysInMonth = [31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  let month = 0;
  while (days >= daysInMonth[month]) {
    days = days - daysInMonth[month];
    month++;
  }

  const hours = Math.floor(msOfDay / 3600000);
  msOfDay = msOfDay - hours * 3600000;
  const minutes = Math.floor(msOfDay / 60000);
  msOfDay = msOfDay - minutes * 60000;
  const seconds = Math.floor(msOfDay / 1000);
  const milliseconds = msOfDay - seconds * 1000;

  return (
    year +
    '-' +
    pad(month + 1, 2) +
    '-' +
    pad(days + 1, 2) +
    'T' +
    pad(hours, 2) +
    ':' +
    pad(minutes, 2) +
    ':' +
    pad(seconds, 2) +
    '.' +
    pad(milliseconds, 3) +
    'Z'
  );
}
