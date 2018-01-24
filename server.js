const express = require('express');
const app = express();
const authHelper = require('./authHelper');
const bodyParser = require('body-parser');
const keys = require('./config');
const session = require('express-session');
const microsoftGraph = require('@microsoft/microsoft-graph-client');
const { API_KEY } = keys;

// const handle = {};
// handle['/mail'] = mail;

app.set('view engine', 'ejs');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
// app.use(
//   session({
//     secret: '0dc529ba-5051-4cd6-8b67-c9a901bb8bdf',
//     resave: false,
//     saveUninitialized: false
//   })
// );

app.get('/', (req, res) => {
  res.render('home', { link: authHelper.getAuthUrl() });
});

// Set up middleware:

// Set CORS
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', '*');
  if (req.method === 'OPTIONS') {
    res.header('Access-Control-Allow-Methods', 'PUT, POST, PATCH, DELETE, GET');
    res.status(200).json({});
  }
  next();
});

app.get('/authorize', (req, res) => {
  console.log("Request handler 'authorize' was called.");

  const { code } = req.query;

  // The authorization code is passed as a query parameter
  console.log(`Code: ${code}`);
  processAuthCode(res, code);
});

// oAuth  process
async function processAuthCode(response, code) {
  let token, email;

  try {
    token = await authHelper.getTokenFromCode(code);
  } catch (error) {
    console.log('Access token error: ', error.message);
    response.writeHead(200, { 'Content-Type': 'text/html' });
    response.write(`<p>ERROR: ${error}</p>`);
    response.end();
    return;
  }

  try {
    email = await getUserEmail(token.token.access_token);
  } catch (error) {
    console.log(`getUserEmail returned an error: ${error}`);
    response.write(`<p>ERROR: ${error}</p>`);
    response.end();
    return;
  }

  const cookies = [
    `node-tutorial-token=${token.token.access_token};Max-Age=4000`,
    `node-tutorial-refresh-token=${token.token.refresh_token};Max-Age=4000`,
    `node-tutorial-token-expires=${token.token.expires_at.getTime()};Max-Age=4000`,
    `node-tutorial-email=${email ? email : ''}';Max-Age=4000`
  ];
  response.setHeader('Set-Cookie', cookies);
  response.redirect(302, '/calendar');
}

// Function to retrieve email in order to find user
async function getUserEmail(token) {
  // Create a Graph client
  const client = microsoftGraph.Client.init({
    authProvider: done => {
      // First parameter is is error handling if no token is retrieved
      // If token is available then return the token
      done(null, token);
    }
  });

  const res = await client.api('/me').get();

  // Outlook.com users have userPrincipalName instead of a mail attribute
  return res.mail ? res.mail : res.userPrincipalName;
}


// Function to retrieve specific value from cookies
function getValueFromCookie(valueName, cookie) {
  if (cookie.includes(valueName)) {
    let start = cookie.indexOf(valueName) + valueName.length + 1;
    let end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}


// Function to retrieve stored access token
async function getAccessToken(request, response) {
  const expiration = new Date(
    parseFloat(
      getValueFromCookie('node-tutorial-token-expires', request.headers.cookie)
    )
  );
  // Check if token is still valid
  if (expiration <= new Date()) {
    // refresh token
    console.log('TOKEN EXPIRED, REFRESHING');
    const refresh_token = getValueFromCookie(
      'node-tutorial-refresh-token',
      request.headers.cookie
    );
    const newToken = await authHelper.refreshAccessToken(refresh_token);

    const cookies = [
      `node-tutorial-token=${token.token.access_token};Max-Age=4000`,
      `node-tutorial-refresh-token=${token.token.refresh_token};Max-Age=4000`,
      `node-tutorial-token-expires=${token.token.expires_at.getTime()};Max-Age=4000`
    ];
    response.setHeader('Set-Cookie', cookies);
    return newToken.token.access_token;
  }
  // Return cached token
  return getValueFromCookie('node-tutorial-token', request.headers.cookie);
}


// Calendar route returning JSON data of user's events
app.get('/calendar', async (request, response) => {
  const token = getValueFromCookie(
    'node-tutorial-token',
    request.headers.cookie
  );
  console.log('Token found in cookie: ', token);
  const email = getValueFromCookie(
    'node-tutorial-email',
    request.headers.cookie
  );
  console.log('Email found in cookie: ', email);

  if (token) {
    // Create a Graph client
    const client = microsoftGraph.Client.init({
      authProvider: done => {
        // Just return the token
        done(null, token);
      }
    });
    try {
      // Get the 10 events with the greatest start date
      const res = await client
        .api('/me/events')
        .header('X-AnchorMailbox', email)
        .top(10)
        .select('subject,start,end,attendees')
        .orderby('start/dateTime DESC')
        .get();

      console.log(res.value);
      response.status(200).json(res.value)

    } catch (err) {
      console.log(`getEvents returned an error: ${err}`);
      response.write(`<p>ERROR: ${err}</p>`);
    }
  } else {
    response.writeHead(200, { 'Content-Type': 'text/html' });
    response.write('<p> No token found in cookie!</p>');
  }
  response.end();
});



// Error handling middleware:
// 404 for unkown routes
app.use((req, res, next) => {
  const error = new Error('Not found');
  error.status = 404;
  next(error);
});

// 500 for internal server errors
app.use((error, req, res, next) => {
  res.status(error.status || 500);
  res.json({
    error: {
      message: error.message
    }
  });
});

const port = process.env.PORT || 8000;

app.listen(port, () => {
  console.log(`Listening on port ${port}...`);
});
