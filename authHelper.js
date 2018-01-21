const keys = require('./config');
const { API_KEY } = keys;

const clientId = keys.API_KEY;
const clientSecret = keys.SECRET_KEY;
const redirectUri = 'http://localhost:8000/authorize';

const scopes = [
  'openid',
  'profile',
  'offline_access',
  'https://outlook.office.com/calendars.readwrite'
];

const credentials = {
  client: {
    id: clientId,
    secret: clientSecret
  },
  auth: {
    tokenHost: 'https://login.microsoftonline.com/common',
    tokenPath: '/oauth2/v2.0/token',
    authorizePath: '/oauth2/v2.0/authorize'
  }
};

const oauth2 = require('simple-oauth2').create(credentials);

module.exports = {
  getAuthUrl: () => {
    const returnVal = oauth2.authCode.authorizeURL({
      redirect_uri: redirectUri,
      scope: scopes.join(' ')
    });
    console.log('');
    console.log(`Generated auth url: ${returnVal}`);
    return returnVal;
  },

  getTokenFromCode(authCode, callback, request, response) {
    oauth2.authCode.getToken(
      {
        code: authCode,
        redirect_uri: redirectUri,
        scope: scopes.join(' ')
      },
      (error, result) => {
        if (error) {
          console.log('Access token error: ', error.message);
          callback(request, response, error, null);
        } else {
          const token = oauth2.accessToken.create(result);
          console.log('');
          console.log('Token created: ', token.token);
          callback(request, response, null, token);
        }
      }
    );
  },

  getEmailFromIdToken(idToken) {
    // JWT is in three parts, separated by a '.'
    const tokenParts = idToken.split('.');

    // Token content is in the second part, in urlsafe base64
    const encodedToken = Buffer.from(
      tokenParts[1].replace('-', '+').replace('_', '/'),
      'base64'
    );

    const decodedToken = encodedToken.toString();

    const jwt = JSON.parse(decodedToken);

    // Email is in the preferred_username field
    return jwt.preferred_username;
  },

  getTokenFromRefreshToken(refreshToken, callback, request, response) {
    const token = oauth2.accessToken.create({
      refreshToken,
      expires_in: 0
    });
    token.refresh((error, result) => {
      if (error) {
        console.log('Refresh token error: ', error.message);
        callback(request, response, error, null);
      } else {
        console.log('New token: ', result.token);
        callback(request, response, null, result);
      }
    });
  }
};
