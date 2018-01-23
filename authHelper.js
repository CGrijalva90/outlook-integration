// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var config = require('./config');
var credentials = {
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
var oauth2 = require('simple-oauth2').create(credentials);

var redirectUri = 'http://localhost:8000/authorize';

// The scopes the app requires
var scopes = [
  'openid',
  'offline_access',
  'User.Read',
  'Mail.Read',
  'Calendars.Read',
  'Contacts.Read'
];

function getAuthUrl() {
  var returnVal = oauth2.authorizationCode.authorizeURL({
    redirect_uri: redirectUri,
    scope: scopes.join(' ')
  });
  console.log('Generated auth url: ' + returnVal);
  return returnVal;
}

// function getTokenFromCode(auth_code, callback, response) {
//   var token;
//   oauth2.authorizationCode.getToken(
//     {
//       code: auth_code,
//       redirect_uri: redirectUri,
//       scope: scopes.join(' ')
//     },
//     function(error, result) {
//       if (error) {
//         console.log('Access token error: ', error.message);
//         callback(response, error, null);
//       } else {
//         token = oauth2.accessToken.create(result);
//         console.log('Token created: ', token.token);
//         callback(response, null, token);
//       }
//     }
//   );
// }

function getTokenFromCode(auth_code, callback, response) {
  var token;
  oauth2.authorizationCode
    .getToken({
      code: auth_code,
      redirect_uri: redirectUri,
      scope: scopes.join(' ')
    })
    .then(result => {
      token = oauth2.accessToken.create(result);
      console.log(token, token.token);
    })
    .catch(error => {
      console.log('Access Token Error', error.message);
    });
}

// oauth2.authorizationCode
//   .getToken(tokenConfig)
//   .then(result => {
//     const accessToken = oauth2.accessToken.create(result);
//   })
//   .catch(error => {
//     console.log('Access Token Error', error.message);
//   });

function refreshAccessToken(refreshToken, callback) {
  var tokenObj = oauth2.accessToken.create({ refresh_token: refreshToken });
  tokenObj.refresh(callback);
}

exports.getAuthUrl = getAuthUrl;
exports.getTokenFromCode = getTokenFromCode;
exports.refreshAccessToken = refreshAccessToken;
