var express = require('express');
var router = express.Router();

var documentController = require('../Controllers/document2');

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

router.get('/document',documentController.document2);

module.exports = router;

