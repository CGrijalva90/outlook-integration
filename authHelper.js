const config = require('./config');

const credentials = {
  client: {
    id: config.API_KEY,
    secret: config.SECRET_KEY
  },
  auth: {
    tokenHost: 'https://login.microsoftonline.com',
    authorizePath: 'common/oauth2/v2.0/authorize',
    tokenPath: 'common/oauth2/v2.0/token'
  }
};
const oauth2 = require('simple-oauth2').create(credentials);

const redirectUri = 'http://localhost:8000/authorize';

// The scopes the app requires
const scopes = [
  'openid',
  'offline_access',
  'User.Read',
  'Mail.Read',
  'Calendars.Read',
  'Contacts.Read'
];

function getAuthUrl() {
  const returnVal = oauth2.authorizationCode.authorizeURL({
    redirect_uri: redirectUri,
    scope: scopes.join(' ')
  });
  console.log(`Generated auth url: ${returnVal}`);
  return returnVal;
}

async function getTokenFromCode(auth_code) {
  const result = await oauth2.authorizationCode.getToken({
    code: auth_code,
    redirect_uri: redirectUri,
    scope: scopes.join(' ')
  });

  const token = oauth2.accessToken.create(result);
  console.log('Token created: ', token.token);
  return token;
}

function refreshAccessToken(refreshToken) {
  return oauth2.accessToken.create({ refresh_token: refreshToken }).refresh();
}

exports.getAuthUrl = getAuthUrl;
exports.getTokenFromCode = getTokenFromCode;
exports.refreshAccessToken = refreshAccessToken;
