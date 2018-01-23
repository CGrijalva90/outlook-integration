const express = require('express');
const app = express();
const authHelper = require('./authHelper');
const bodyParser = require('body-parser');
const outlookRoutes = require('./api/outlook-routes');
const keys = require('./config');
const session = require('express-session');
const microsoftGraph = require('@microsoft/microsoft-graph-client');
const { API_KEY } = keys;

// const handle = {};
// handle['/mail'] = mail;

app.set('view engine', 'ejs');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(
  session({
    secret: '0dc529ba-5051-4cd6-8b67-c9a901bb8bdf',
    resave: false,
    saveUninitialized: false
  })
);

app.get('/', (req, res) => {
  res.redirect('/outlook');
});

app.get('/login', (req, res) => {
  res.send(authHelper.getAuthUrl);
});

// Set up middleware:

// CORS protection
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', '*');
  if (req.method === 'OPTIONS') {
    res.header('Access-Control-Allow-Methods', 'PUT, POST, PATCH, DELETE, GET');
    res.status(200).json({});
  }
  next();
});

// Outlook routes (calendar and mail will be implemented)
app.use('/outlook', outlookRoutes);

app.get('/authorize', (req, res) => {
  console.log("Request handler 'authorize' was called.");

  const { code } = req.query;
  // The authorization code is passed as a query parameter

  console.log(`Code: ${code}`);
  processAuthCode(res, code);
});

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
  response.writeHead(302, { Location: 'http://localhost:8000/mail' });
  response.end();
}

async function getUserEmail(token) {
  // Create a Graph client
  const client = microsoftGraph.Client.init({
    authProvider: done => {
      // Just return the token
      done(null, token);
    }
  });

  // Get the Graph /Me endpoint to get user email address
  const res = await client.api('/me').get();

  // Office 365 users have a mail attribute
  // Outlook.com users do not, instead they have
  // userPrincipalName
  return res.mail ? res.mail : res.userPrincipalName;
}

function getValueFromCookie(valueName, cookie) {
  if (cookie.includes(valueName)) {
    let start = cookie.indexOf(valueName) + valueName.length + 1;
    let end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}

async function getAccessToken(request, response) {
  const expiration = new Date(
    parseFloat(
      getValueFromCookie('node-tutorial-token-expires', request.headers.cookie)
    )
  );

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

app.get('/mail', async function mail(request, response) {
  let token;

  try {
    token = await getAccessToken(request, response);
  } catch (error) {
    response.writeHead(200, { 'Content-Type': 'text/html' });
    response.write('<p> No token found in cookie!</p>');
    response.end();
    return;
  }

  console.log('Token found in cookie: ', token);
  const email = getValueFromCookie(
    'node-tutorial-email',
    request.headers.cookie
  );
  console.log('Email found in cookie: ', email);

  response.writeHead(200, { 'Content-Type': 'text/html' });
  response.write('<div><h1>Your inbox</h1></div>');

  // Create a Graph client
  const client = microsoftGraph.Client.init({
    authProvider: done => {
      // Just return the token
      done(null, token);
    }
  });

  try {
    // Get the 10 newest messages
    const res = await client
      .api('/me/mailfolders/inbox/messages')
      .header('X-AnchorMailbox', email)
      .top(10)
      .select('subject,from,receivedDateTime,isRead')
      .orderby('receivedDateTime DESC')
      .get();

    console.log(`getMessages returned ${res.value.length} messages.`);
    response.write(
      '<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>'
    );
    res.value.forEach(message => {
      console.log('  Subject: ' + message.subject);
      const from = message.from ? message.from.emailAddress.name : 'NONE';
      response.write(
        `<tr><td>${from}` +
          `</td><td>${message.isRead ? '' : '<b>'} ${message.subject} ${
            message.isRead ? '' : '</b>'
          }` +
          `</td><td>${message.receivedDateTime.toString()}</td></tr>`
      );
    });

    response.write('</table>');
  } catch (err) {
    console.log(`getMessages returned an error: ${err}`);
    response.write(`<p>ERROR: ${err}</p>`);
  }

  response.end();
});

// app.get('/mail', async(mail));

// Error handling middleware
// app.use((req, res, next) => {
//   const error = new Error('Not found');
//   error.status = 404;
//   next(error);
// });

// app.use((error, req, res, next) => {
//   res.status(error.status || 500);
//   res.json({
//     error: {
//       message: error.message
//     }
//   });
// });

const port = process.env.PORT || 8000;

app.listen(port);
