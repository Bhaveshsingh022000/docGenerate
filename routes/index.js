var express = require('express');
var router = express.Router();

var spccProposalDocumentController = require('../Controllers/SpccProposalDocument');
var api150TankInspectionController = require('../Controllers/API-510ExtrnalTankInspectionProposalDocument');

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

router.get('/spccproposal',spccProposalDocumentController.proposalDocument);
router.get('/api150tankinspection',api150TankInspectionController.proposalDocument);

module.exports = router;
