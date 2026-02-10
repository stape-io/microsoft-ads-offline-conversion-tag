const encodeUriComponent = require('encodeUriComponent');
const Firestore = require('Firestore');
const getAllEventData = require('getAllEventData');
const getContainerVersion = require('getContainerVersion');
const getRequestHeader = require('getRequestHeader');
const getTimestampMillis = require('getTimestampMillis');
const getType = require('getType');
const JSON = require('JSON');
const logToConsole = require('logToConsole');
const makeString = require('makeString');
const Math = require('Math');
const sendHttpRequest = require('sendHttpRequest');
const sha256Sync = require('sha256Sync');

/*==============================================================================
==============================================================================*/

const eventData = getAllEventData();
const isLoggingEnabled = determinateIsLoggingEnabled();
const traceId = getRequestHeader('trace-id');
const firebaseOptions = {};

if (!isConsentGivenOrNotRequired(data, eventData)) {
  return data.gtmOnSuccess();
}
if (!data.microsoftClickId && !data.hashedEmailAddress && !data.hashedPhoneNumber) {
  return data.gtmOnSuccess();
}

if (data.authFlow === 'stape') {
  return sendConversionRequest(
    getData('StapeAccessToken', 'StapeDeveloperToken'),
    handleStapeResponse
  );
}

if (data.firebaseProjectId) firebaseOptions.projectId = data.firebaseProjectId;

Firestore.read(data.firebasePath, firebaseOptions).then(
  (result) => {
    const postBody = getData(result.data.access_token, data.developerToken);

    return sendConversionRequest(postBody, getHandleResponse(result.data.refresh_token));
  },
  () => updateAccessToken(data.refreshToken)
);

/*==============================================================================
Vendor related functions
==============================================================================*/

function sendConversionRequest(postBody, responseCallback) {
  const postUrl = getUrl();

  if (isLoggingEnabled) {
    logToConsole(
      JSON.stringify({
        Name: 'MicrosoftAdsOfflineConversion',
        Type: 'Request',
        TraceId: traceId,
        EventName: data.conversionName,
        RequestMethod: 'POST',
        RequestUrl: postUrl,
        RequestBody: postBody
      })
    );
  }

  sendHttpRequest(
    postUrl,
    (statusCode, headers, body) => {
      if (isLoggingEnabled) {
        logToConsole(
          JSON.stringify({
            Name: 'MicrosoftAdsOfflineConversion',
            Type: 'Response',
            TraceId: traceId,
            EventName: data.conversionName,
            ResponseStatusCode: statusCode,
            ResponseHeaders: headers,
            ResponseBody: body
          })
        );
      }

      responseCallback(statusCode, headers, body);
    },
    { headers: getConversionRequestHeaders(), method: 'POST' },
    postBody
  );
}

function handleStapeResponse(statusCode, headers, body) {
  if (statusCode >= 200 && statusCode < 400) {
    data.gtmOnSuccess();
  } else {
    data.gtmOnFailure();
  }
}

function getHandleResponse(refreshToken) {
  return (statusCode, headers, body) => {
    if (statusCode >= 200 && statusCode < 400) {
      if (body.indexOf('Authentication token expired') !== -1) {
        updateAccessToken(refreshToken);
      } else {
        data.gtmOnSuccess();
      }
    } else if (statusCode === 401) {
      updateAccessToken(refreshToken);
    } else {
      data.gtmOnFailure();
    }
  };
}

function getConversionRequestHeaders() {
  return {
    'Content-Type': 'text/xml; charset=utf-8',
    SOAPAction: 'ApplyOfflineConversions'
  };
}

function updateAccessToken(refreshToken) {
  const authUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token/';
  const authBody =
    'refresh_token=' +
    enc(refreshToken || data.refreshToken) +
    '&client_id=' +
    enc(data.clientId) +
    '&client_secret=' +
    enc(data.clientSecret) +
    '&grant_type=refresh_token&scope=https%3A%2F%2Fads.microsoft.com%2Fmsads.manage+offline_access&tenant=common';

  if (isLoggingEnabled) {
    logToConsole(
      JSON.stringify({
        Name: 'MicrosoftAdsOfflineConversion',
        Type: 'Request',
        TraceId: traceId,
        EventName: 'Auth',
        RequestMethod: 'POST',
        RequestUrl: authUrl
      })
    );
  }

  sendHttpRequest(
    authUrl,
    (statusCode, headers, body) => {
      if (isLoggingEnabled) {
        logToConsole(
          JSON.stringify({
            Name: 'MicrosoftAdsOfflineConversion',
            Type: 'Response',
            TraceId: traceId,
            EventName: 'Auth',
            ResponseStatusCode: statusCode,
            ResponseHeaders: headers
          })
        );
      }

      if (statusCode >= 200 && statusCode < 400) {
        let bodyParsed = JSON.parse(body);

        Firestore.write(data.firebasePath, bodyParsed, firebaseOptions).then(() => {
          sendConversionRequest(
            getData(bodyParsed.access_token, data.developerToken),
            getHandleResponse(bodyParsed.refresh_token)
          );
        }, data.gtmOnFailure);
      } else {
        data.gtmOnFailure();
      }
    },
    {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      method: 'POST'
    },
    authBody
  );
}

function getUrl() {
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

function getData(accessToken, developerToken) {
  const eventData = getAllEventData();

  const email = data.hashedEmailAddress;
  const phone = data.hashedPhoneNumber;
  const value = data.conversionValue || eventData.value;

  let conversionCurrencyCode = '';
  if (data.conversionCurrencyCode) conversionCurrencyCode = data.conversionCurrencyCode;
  else if (eventData.currencyCode) conversionCurrencyCode = eventData.currencyCode;
  else if (eventData.currency) conversionCurrencyCode = eventData.currency;

  let externalAttributionCredit = data.externalAttributionCredit;
  let externalAttributionModel = data.externalAttributionModel;
  let microsoftClickId = data.microsoftClickId;

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
                         (email ? '<HashedEmailAddress i:nil="false">'+hashData('hashedEmailAddress', email)+'</HashedEmailAddress>\n' : '') +
                         (phone ? '<HashedPhoneNumber i:nil="false">'+hashData('hashedPhoneNumber', phone)+'</HashedPhoneNumber>\n' : '') +
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

function isConsentGivenOrNotRequired(data, eventData) {
  if (data.adStorageConsent !== 'required') return true;
  if (eventData.consent_state) return !!eventData.consent_state.ad_storage;
  const xGaGcs = eventData['x-ga-gcs'] || ''; // x-ga-gcs is a string like "G110"
  return xGaGcs[2] === '1';
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
  if (!value) {
    return value;
  }

  const type = getType(value);

  if (type === 'undefined' || value === 'undefined') {
    return undefined;
  }

  if (type === 'object') {
    return value.map((val) => {
      return hashData(key, val);
    });
  }

  if (isHashed(value)) {
    return value;
  }

  value = makeString(value).trim().toLowerCase();

  if (key === 'hashedPhoneNumber') {
    value = value.split(' ').join('').split('-').join('').split('(').join('').split(')').join('');
  }

  return sha256Sync(value, { outputEncoding: 'hex' });
}

function convertTimestampToISO(timestamp) {
  const secToMs = function (s) {
    return s * 1000;
  };
  const minToMs = function (m) {
    return m * secToMs(60);
  };
  const hoursToMs = function (h) {
    return h * minToMs(60);
  };
  const daysToMs = function (d) {
    return d * hoursToMs(24);
  };
  const format = function (value) {
    return value >= 10 ? value.toString() : '0' + value;
  };
  const fourYearsInMs = daysToMs(365 * 4 + 1);
  let year = 1970 + Math.floor(timestamp / fourYearsInMs) * 4;
  timestamp = timestamp % fourYearsInMs;

  while (true) {
    let isLeapYear = !(year % 4);
    let nextTimestamp = timestamp - daysToMs(isLeapYear ? 366 : 365);
    if (nextTimestamp < 0) {
      break;
    }
    timestamp = nextTimestamp;
    year = year + 1;
  }

  const daysByMonth =
    year % 4 === 0
      ? [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
      : [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

  let month = 0;
  for (let i = 0; i < daysByMonth.length; i++) {
    let msInThisMonth = daysToMs(daysByMonth[i]);
    if (timestamp > msInThisMonth) {
      timestamp = timestamp - msInThisMonth;
    } else {
      month = i + 1;
      break;
    }
  }
  let date = Math.ceil(timestamp / daysToMs(1));
  timestamp = timestamp - daysToMs(date - 1);
  let hours = Math.floor(timestamp / hoursToMs(1));
  timestamp = timestamp - hoursToMs(hours);
  let minutes = Math.floor(timestamp / minToMs(1));
  timestamp = timestamp - minToMs(minutes);
  let sec = Math.floor(timestamp / secToMs(1));

  return (
    year +
    '-' +
    format(month) +
    '-' +
    format(date) +
    'T' +
    format(hours) +
    ':' +
    format(minutes) +
    ':' +
    format(sec) +
    '.1111111Z'
  );
}

function determinateIsLoggingEnabled() {
  const containerVersion = getContainerVersion();
  const isDebug = !!(
    containerVersion &&
    (containerVersion.debugMode || containerVersion.previewMode)
  );

  if (!data.logType) {
    return isDebug;
  }

  if (data.logType === 'no') {
    return false;
  }

  if (data.logType === 'debug') {
    return isDebug;
  }

  return data.logType === 'always';
}
