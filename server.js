const express = require('express');
const app = express();
const authHelper = require('./authHelper');
const bodyParser = require('body-parser');
const outlookRoutes = require('./api/outlook-routes');
const pages = require('./pages');
const keys = require('./config');
const { API_KEY } = keys;

app.set('view engine', 'ejs');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.get('/', (req, res) => {
  res.send('Greetings from the index page!');
});

app.get('/login', (req, res) => {
  res.send(authHelper.getAuthUrl);
});

const tokenReceived = (req, res, error, token) => {
  if (error) {
    console.log(`ERROR getting token: ${error}`);
    res.send(`ERROR getting token: ${error}`);
  } else {
    // save tokens in session
    req.session.access_token = token.token.access_token;
    req.session.refresh_token = token.token.refresh_token;
    req.session.email = authHelper.getEmailFromIdToken(token.token.id_token);
    res.redirect('/logincomplete');
  }
};

app.get('/authorize', (req, res) => {
  const authCode = req.query.code;
  if (authCode) {
    console.log('');
    console.log(`Retrieved authCode in /authorize ${authCode}`);
    authHelper.getTokenFromCode(authCode, tokenReceived, req, res);
  } else {
    res.redirect('/');
  }
});

app.get('/logincomplete', (req, res) => {
  const { access_token } = req.session;
  const { refresh_token } = req.session;
  const { email } = req.session;

  if (access_token === undefined || refresh_token === undefined) { // eslint-disable-line
    console.log('/logincomplete called while not logged in');
    res.redirect('/');
    return;
  }

  res.send(pages.loginCompletePage(email));
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

// Error handling middleware
app.use((req, res, next) => {
  const error = new Error('Not found');
  error.status = 404;
  next(error);
});

app.use((error, req, res, next) => {
  res.status(error.status || 500);
  res.json({
    error: {
      message: error.message
    }
  });
});

const port = process.env.PORT || 8000;

app.listen(port);
