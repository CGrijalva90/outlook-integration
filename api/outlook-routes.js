const express = require('express');
const router = express.Router();
const authHelper = require('../authHelper');

router.get('/', (req, res) => {
  res.render('home', { link: authHelper.getAuthUrl() });
});

module.exports = router;
