'use strict';
var cheerio = require('cheerio');
var async = require('async');
var sleep = require('sleep');
var _ = require('lodash');
var Excel = require('exceljs');
var co = require('co');
var pages = 0;
var workbook = new Excel.Workbook();
var worksheet = workbook.addWorksheet('RESULTS1');
var Url ="https://14.139.56.15/scheme13/studentresult/index.asp";
var request = require('request').defaults({
  strictSSL: false,
  rejectUnauthorized: false
});

worksheet.columns = [{
	header:'Name',
	key: 'name',
	width: 30 
} , {
  header:'Rollno',
  key: 'rollno',
  width: 30 
}, {
	header:'CGPA',
	key: 'cg',
	width: 30 
} , {
	header:'SGPA',
	key: 'sg',
	width: 30 
}
];

function FetchUrls (){
  co(function*() {
		var name,cg,sg,saveObject,rollno;// console.log("haan");
    for( var i1 = 13101 ; i1<13190 ; i1++ ){
  		var options = { 
  			method: 'POST',
    			url: 'https://14.139.56.15/scheme13/studentresult/details.asp',
   			headers: {
   			  'postman-token': 'a5bf99c2-0244-b64b-a77f-c7de59cba996',
       		'cache-control': 'no-cache',accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
       		'content-type': 'application/x-www-form-urlencoded',
       		'set-cookie': 'ASPSESSIONIDSCTASSQQ=EBPHBFGAPCMHGFPCHPEMANOF' 
        },
    		form: { RollNumber: i1.toString(), B1: 'Submit' }
      };

  		request(options, function (error, response, body) {
        if (error) throw new Error(error);
    		var $ = cheerio.load(body, {
          xmlMode: true
        });
        // console.log(body);
        var i = 0 ;
        $('div[align="center"]').each(function(){
          i++;
          if(i==2){
          	var ccc = 0 ;
          	var c1=$(this).find('td').each(function(){
          		ccc ++ ;
          		if(ccc == 2){
          			name = $(this).find('div').text().trim();
          			// console.log(name);
          		}
              if(ccc == 4){
                rollno = $(this).find('div').text().trim();
                // console.log(name);
              }
          	});
          }
          if(i==24){
          	var ss = 0 ;
          	var a = $(this).find('td').each(function(){
            	ss++;
            	// console.log(ss);
            	if(ss==5){
                 sg = $(this).text();              
              }
              if(ss==7){
            		var nn = $(this).text();
                var ttt = nn.split('=');
                cg = ttt[1];
            	}
            });
          }
        });
        saveObject = {
          name:name,
          rollno:rollno,
          cg:cg,
          sg:sg
        };
        console.log(saveObject);
        worksheet.addRow(saveObject).commit();;
        workbook.xlsx.writeFile('RESULT_Civil_SEM_6th.xlsx').then(function() {
          console.log('Row added & Saved');
        });
      });
    }    
	});
}

var init = function () {

	co(function*() {
	FetchUrls();
   
	});
}

console.log("Starting  scraping...");
init();