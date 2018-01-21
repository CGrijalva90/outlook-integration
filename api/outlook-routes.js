const express = require('express');
const router = express.Router();

router.get('/', (req, res) => {
  res.send('Coming from the outlook routes!');
});

module.exports = router;
