'use strict'
// required NPM modules
let express = require("express");
let app = express();
let sql = require('mssql');
let Handlebars = require('handlebars');
let csv = require('csv');
let nodemailer = require("nodemailer");
let nodeSSPI = require('node-sspi');
let moment = require('moment');

// config for database
var config = {
    user: 'xxxxxxxxx',
    password: 'xxxxxxxxx',
    server: 'xxxxxxxxx', 
    database: 'xxxxxxxxx', 
    "options": {
    "encrypt": true,
    "enableArithAbort": true
},};
    
app.set('port', process.env.PORT || 3000);
app.use(express.static(__dirname + '/views')); // allows direct navigation to static files
app.use(require("body-parser").urlencoded({extended: true})); // parse form submissions

//windows authentication with node-sspi module
app.use(function (req, res, next) {
  var nodeSSPIObj = new nodeSSPI({
    retrieveGroups: true
  })
  nodeSSPIObj.authenticate(req, res, function(err){
    res.finished || next()
  })
})

// setting handlebars has view engine and requiring express-handlebars module
let handlebars =  require("express-handlebars")
.create({ defaultLayout: "main"});
app.engine("handlebars", handlebars.engine);
app.set("view engine", "handlebars");

// global variables for dates
app.locals.copyright =  new Date().getFullYear();
app.locals.fullyear =  moment().format('L');
app.locals.weekAgo =  moment().subtract(7, 'days').calendar();
app.locals.weekAgoReformat = app.locals.weekAgo[6] + app.locals.weekAgo[7] + app.locals.weekAgo[8] + app.locals.weekAgo[9] + "-" + app.locals.weekAgo[0] + app.locals.weekAgo[1] + "-" + app.locals.weekAgo[3] + app.locals.weekAgo[4];

app.locals.fullyearReformat = app.locals.fullyear[6] + app.locals.fullyear[7] + app.locals.fullyear[8] + app.locals.fullyear[9] + "-" + app.locals.fullyear[0] + app.locals.fullyear[1] + "-" + app.locals.fullyear[3] + app.locals.fullyear[4];

app.locals.yearAgo = (moment().format('YYYY') -1) + "-" + app.locals.fullyear[0] + app.locals.fullyear[1] + "-" + app.locals.fullyear[3] + app.locals.fullyear[4];

// send email when error
async function main(user, err, pathinfo) {

    process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
    
  // create reusable transporter object using the default SMTP transport
  let transporter = nodemailer.createTransport({
    host: "xxxxxxxxx.xxxxxxxxx.com",
    port: 25,
    secure: false, // true for 465, false for other ports
      tls: { secureProtocol: "TLSv1_method" }
  });

let uname = "xxxxxxxxx@xxxxxxxxx.com, " + user;
  // send mail with defined transport object
  let info = await transporter.sendMail({
    from: '"Consumer Affairs App" <xxxxxxxxx@xxxxxxxxx.com>', // sender address
    to: uname, // list of receivers
    subject: "Consumer Affairs App: An error has occured.", // Subject line
    html: "An error has just occurred in the Consumer Affairs application, and your application administrator has been notified. Details are below. <br> <br> <hr> <br> Username: " + user + " <br> Date: " + moment().format('L') + "<br> Time: " + moment().format('LT') + "<br> Error: <b style='color:red'>"  + err + "</b> <br> URL where error occurred: " + pathinfo + "<br><br> <i>This is an automated message. Please do not reply.</i>" // body
  });
}

// for Handlebars Index value
Handlebars.registerHelper("inc", function(value, options){
    return parseInt(value) + 1;
});

// array for dropdowns
let arr = [0,1,2,3,4,5,6,7,8,9,10];

// return all plant names and numbers
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'select plant, plantname from tblplant';
     
return pool.query(querycode);
}).then(result => {   
 let plantNums = "";
        try{
global.plantNums = result.recordset;
        } catch {
global.plantNums = "";
}
});    
 
// return all report codes
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'SELECT reportcode FROM tblReportCode';
     
return pool.query(querycode);
}).then(result => {   
 let reportCodes = "";
        try{
global.reportCodes = result.recordset;
        } catch {
global.reportCodes = "";
}
});  
    
// return all how received
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'SELECT howreceived FROM tblHowReceived';
     
return pool.query(querycode);
}).then(result => {   
 let howReceived = "";
        try{
global.howReceived = result.recordset;
        } catch {
global.howReceived = "";
}
});
            
// return all UPC codes, GTINs and descriptions
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'SELECT * FROM tblProductInfo';
     
return pool.query(querycode);
}).then(result => {   
 let prodInfo = "";
        try{
global.prodInfo = result.recordset;
        } catch {
global.prodInfo = "";
}
});            
       
// return all salutations
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'SELECT salutation FROM tblSalutation';
     
return pool.query(querycode);
}).then(result => {   
 let salutations = "";
        try{
global.salutations = result.recordset;
        } catch {
global.salutations = "";
}
});
        
// return all authorized users
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'SELECT UserName FROM tblAuthUsers';
     
return pool.query(querycode);
}).then(result => {   
 let authUser = "";
        try{
global.authUser = result.recordset;
        } catch {
global.authUser = "";
}
});
                 
// return all countries
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'SELECT Country FROM tblcountry';
     
return pool.query(querycode);
}).then(result => {   
 let countries = "";
        try{
global.countries = result.recordset;
        } catch {
global.countries = "";
}
});         
     
// return all states and provinces
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'SELECT stateorprovince FROM tblState';
     
return pool.query(querycode);
}).then(result => {   
 let stateOrProvince = "";
        try{
global.stateOrProvince = result.recordset;
        } catch {
global.stateOrProvince = "";
}
});           
    
// return all CA Agent names
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'select caagent from tblcaagent';
     
return pool.query(querycode);
}).then(result => {   
 let caAgent = "";
        try{
global.caAgent = result.recordset;
        } catch {
global.caAgent = "";
}
});
   
// return all complaints
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'SELECT status FROM tblStatus';
     
return pool.query(querycode);
}).then(result => {   
 let complaintStatus = "";
        try{
global.complaintStatus = result.recordset;
        } catch {
global.complaintStatus = "";
}
});

// Detail page
app.get('/detail', function(req, res) { 
    
app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
    
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'select [CustomerID],[ReceiveDate],[DOI],[CAAgent],[ProductName],[UPC],[PackSize],UBD,[Time],[Plant],[Complaint],[ReportCode],[RetailerCustomer],[HowReceived],[ProductComments],[Salutation],[FirstName],[MiddleInitial],[LastName],[CompanyName],[Address1],[Address2],[City],[ST],[Zip],[Country],[PhoneNumber],[CellNumber],[EmailAddress],[ConsumerNotes],[StatusOfComplaint],[FoodService],[GovAgency],[CustNum],[GCNumber],[TypeofComp],[AmtofComp],[Coupon55],[Coupon400],[Coupon800],[Coupon65CAN],[Coupon900CAN],[AmtofCoupon],[CAResponse],[CANumber] from tblConsumers where CANumber = ' + req.query.q; 
    
return pool.query(querycode);
}).then(result => {   
    
  res.render('detail', {pageName: "details", title: "Details", data: result.recordset[0], queryString: result.recordset[0].FirstName + " " + result.recordset[0].LastName}); 

}).catch(err => {
    console.log(err);
});
});

app.get('/', function(req, res) { 
    
app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
    
new sql.ConnectionPool(config).connect().then(pool => {    
    
let querycode = "";
    
if(req.query.fromdate){
querycode = 'select * from tblConsumers where (ReceiveDate between ' + "'" + req.query.fromdate + "'" + ' and ' + "'" + req.query.todate + "'" + ') and (firstname like ' + "'" + req.query.fname + "'" + ' or lastname like ' + "'" + req.query.lname + "'" + ' or phonenumber like ' +  "'" + req.query.phone + "'" + ' or emailaddress like ' +  "'" + req.query.email + "'" + ' or CANumber like ' +  "'" + req.query.customerid + "'" + ') order by ReceiveDate desc';
}
    
return pool.query(querycode);
}).then(result => {
    
res.render('search', {data: result.recordset, pageName: "search", title: "Search", queryString: req.query.fromdate});
    
}).catch(err => {
    console.log(err);
});
});

// Reports page
app.get('/reportresult', function(req, res) { 
    
app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

 // return count of plant records for last week   
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'select count(*) as cnt from tblConsumers where ReceiveDate > ' + "'" + app.locals.weekAgoReformat + "'";
     
return pool.query(querycode);
}).then(result => {   
 let reportCount = "";
global.reportCount = result.recordset[0].cnt;
});
 
 // return count of most recurring ReportCount plant records for last week   
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'select top 1 tblconsumers.reportcode, codedescription, count(*) as cnt from tblconsumers join tblreportcode on tblconsumers.reportcode = tblreportcode.reportcode  where receivedate > ' + "'" + app.locals.weekAgoReformat + "'" + ' and plant = ' + "'" + req.query.plant + "'" + ' group by tblconsumers.reportcode, codedescription  order by count(*) desc';
     
return pool.query(querycode);
}).then(result => {   
 let reportCodeCount = "";
 let reportCodeName = "";
let reportCodeDescription = "";
    try{
global.reportCodeCount = result.recordset[0].cnt;
global.reportCodeName = result.recordset[0].reportcode;
global.reportCodeDescription = result.recordset[0].codedescription;
    } catch(err) {
  reportCodeCount = "";
  reportCodeName = "";     
reportCodeDescription = "";
    }
});

// return plant name  
new sql.ConnectionPool(config).connect().then(pool => {
    
let querycode = 'select plantname from tblplant where plant = ' + "'" + req.query.plant + "'";
     
return pool.query(querycode);
}).then(result => {   
 let plantName = "";
    try{
global.plantName = result.recordset[0].plantname;
    } catch(err) {
  plantName = "";
    }
});
    
new sql.ConnectionPool(config).connect().then(pool => {    
        
let querycode = '';
        
if(req.query.plant){
querycode = 'select * from tblConsumers where Plant = ' + "'" + req.query.plant + "'" + ' and ReceiveDate > ' + "'" + app.locals.weekAgoReformat + "' order by ReceiveDate desc";
}
    
return pool.query(querycode);
}).then(result => {   

try {
res.render('reportresult', {pageName: "reports", title: "Reports", data: result.recordset, queryString: req.query.plant, reportCount: reportCount, reportPercent: ((result.recordset.length/reportCount) * 100).toFixed(1), reportCodeCount: reportCodeCount, reportCodeName: reportCodeName, reportCodePercent: ((reportCodeCount/result.recordset.length) * 100).toFixed(1), reportCodeDescription: reportCodeDescription, plantName: plantName, plantNums: plantNums})
} catch (err) {
     res.render('reportresult', {pageName: "reports", title: "Reports", plantNums: plantNums, queryString: req.query.plant}); 
}
    
}).catch(err => {
    console.log(err);
});
});

// Reports page
app.get('/reports', function(req, res) { 
    
    app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
    
    global.isUserAuthorized = '';
    
    for(var i=0; i<global.authUser.length; i++) {
        if((req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") == global.authUser[i].UserName.toLowerCase()) {
            global.isUserAuthorized = true;
        }
    }
    
if(isUserAuthorized){
    
new sql.ConnectionPool(config).connect().then(pool => {    
        
let querycode = '';
        
if(req.query.plant){
querycode = 'select * from tblConsumers where Plant = ' + "'" + req.query.plant + "'" + ' and ReceiveDate > ' + "'" + app.locals.weekAgoReformat + "' order by ReceiveDate desc";
}
    
return pool.query(querycode);
}).then(result => {   

res.render('reports', {pageName: "reports", title: "Reports", plantNums: plantNums, data: result.recordset}); 
    
}).catch(err => {
    console.log(err);
});
        } else {
            res.render('unauth', {uname: (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com", pageName: "Unauthorized", title: "Unauthorized"})
        }
});
  
// Update page
app.use('/update', function(req,res) {
app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
    
     if (req.method === 'POST') {    
new sql.ConnectionPool(config).connect().then(pool => {    
    
global.qc = '';
let DOI = '';
DOI = '1900-01-01';
DOI = req.body.DOI;
         
 if (req.body.ReportCodes == 'NC' || (req.body.ReportCodes >= 5 && req.body.ReportCodes < 6) || (req.body.ReportCodes >= 8 && req.body.ReportCodes < 10) ){
    global.complaint = 'No';
    } else {
    global.complaint = 'Yes';    
    }
     
global.qc = 'UPDATE [tblConsumers] SET [ReceiveDate] = ' + "'" + req.body.ReceiveDate + "'" + ',[DOI] = ' + "'" + DOI + "'" + ',[CAAgent] = ' + "'" + req.body.CAAgent + "'" + ',[ProductName] = ' + "'" + req.body.ProductName.replace(/'/g, "''") + "'"  + ',[UPC] = ' + "'" + req.body.UPC + "'"  + ',[PackSize] = ' + "'" + req.body.PackSize.replace(/'/g, "''") + "'"  + ',[UBD] = ' + "'" + req.body.UBD + "'"  + ',[Time] = ' + "'" + req.body.Time + "'"  + ',[Plant] = ' + "'" + req.body.PlantNum + "'"  + ',[Complaint] = ' + "'" + complaint + "'" + ',[ReportCode] = ' + "'" + req.body.ReportCodes + "'"  + ',[RetailerCustomer] =' + "'" + req.body.RetailerCustomer.replace(/'/g, "''") + "'"  + ',[HowReceived] = ' + "'" + req.body.HowReceived.replace(/'/g, "''") + "'" + ',[ProductComments] =' + "'" + req.body.ProductComments.replace(/'/g, "''").replace(/  /g, "") + "'" + ',[Salutation] = ' + "'" + req.body.Salutation + "'" + ',[FirstName] = ' + "'" + req.body.FirstName + "'" + ',[MiddleInitial] = ' + "'" + req.body.MiddleInitial + "'" + ',[LastName] = ' + "'" + req.body.LastName + "'" + ',[CompanyName] = ' + "'" + req.body.CompanyName.replace(/'/g, "''") + "'" + ',[Address1] = ' + "'" + req.body.Address1.replace(/'/g, "''") + "'" + ',[Address2] = ' + "'" + req.body.Address2.replace(/'/g, "''") + "'" + ',[City] = ' + "'" + req.body.City + "'" + ',[ST] = ' + "'" + req.body.State + "'" + ',[ZIP] = ' + "'" + req.body.ZIP + "'" + ',[Country] = ' + "'" + req.body.Country + "'" + ',[PhoneNumber] = ' + "'" + req.body.PhoneNumber.replace(/'/g, "''") + "'" + ',[CellNumber] = ' + "'" + req.body.CellNumber + "'" + ',[EmailAddress] = ' + "'" + req.body.EmailAddress + "'"  + ',[ConsumerNotes] =' + "'" + req.body.ConsumerNotes.replace(/'/g, "''").replace(/  /g, "") + "'" + ',[StatusOfComplaint] = ' + "'" + req.body.StatusOfComplaint + "'" + ',[FoodService] = ' + "'" + req.body.FoodService + "'" + ',[GovAgency] = ' + "'" + req.body.GovAgency + "'" + ',[CustNum] = ' + "'" + req.body.CustNum + "'" + ',[GCNumber] = ' + "'" + req.body.GCNumber + "'" + ',[TypeofComp] = ' + "'" + req.body.TypeofComp + "'" + ',[AmtofComp] = ' + "'" + req.body.AmtofComp + "'" + ',[Coupon55] = ' + "'" + req.body.Coupon55 + "'"  + ',[Coupon400] = ' + "'" + req.body.Coupon400 + "'" + ',[Coupon800] = ' + "'" + req.body.Coupon800 + "'"  + ',[Coupon65CAN] = ' + "'" + req.body.Coupon65CAN + "'"   + ',[Coupon900CAN] = ' + "'" + req.body.Coupon900CAN + "'"  + ',[AmtofCoupon] = ' + "'" + req.body.AmtofCoupon + "'" + ',[CAResponse] = ' + "'" + req.body.CAResponse.replace(/'/g, "''") + "'" + ' where CustomerID = ' + req.query.custid;

return pool.query(qc);
}).then(result => {
res.render('success', {record: req.query.can, stat: "updated", pageName: "Success", title: "Success"});
    
}).catch(err => {
res.render('failure', {record: req.query.can, err: err, pageName: "Error", title: "Error", stat: "updated"});
main(((req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com"), err, (req.protocol + '://' + req.get('host') + req.originalUrl)).catch(console.error);
}); } else {
    
 new sql.ConnectionPool(config).connect().then(pool => {

let querycode = 'select [CustomerID],[ReceiveDate],[DOI],[CAAgent],[ProductName],[UPC],[PackSize], [UBD],[Time],[Plant],[Complaint],[ReportCode],[RetailerCustomer],[HowReceived],[ProductComments],[Salutation],[FirstName],[MiddleInitial],[LastName],[CompanyName],[Address1],[Address2],[City],[ST],[Zip],[Country],[PhoneNumber],[CellNumber],[EmailAddress],[ConsumerNotes],[StatusOfComplaint],[FoodService],[GovAgency],[CustNum],[GCNumber],[TypeofComp],[AmtofComp],[Coupon55],[Coupon400],[Coupon800],[Coupon65CAN],[Coupon900CAN],[AmtofCoupon],[CAResponse],[CANumber] from tblConsumers where CustomerID = ' + req.query.custid;
     
return pool.query(querycode);
}).then(result => {

 res.render('update', {pageName: "update", title: "Update", data: result.recordset[0], queryString: result.recordset[0].FirstName + " " + result.recordset[0].LastName, caAgent: caAgent, plantNums: plantNums, arr: arr, reportCodes: reportCodes, howReceived: howReceived, salutations: salutations, countries: countries, stateOrProvince: stateOrProvince, complaintStatus: complaintStatus}); 
         
}).catch(err => {
    console.log(err); 
}); }
}); 
    
// Create new record page
app.use('/create', function(req,res) {
    
app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

if(req.method === 'POST'){
            
    if (req.body.GovAgency == 'on'){
    global.govAgency = 'Yes';
    } else {
    global.govAgency = 'No';    
    }
    
    if (req.body.FoodService == 'on'){
    global.foodService = 'Yes';
    } else {
    global.foodService = 'No';    
    }
    
 if (req.body.ReportCode == 'NC' || (req.body.ReportCode >= 5 && req.body.ReportCode < 6) || (req.body.ReportCode >= 8 && req.body.ReportCode < 10) ){
    global.complaint = 'No';
    } else {
    global.complaint = 'Yes';    
    }
    
    let querycode = '';
    let DOI = '';
    DOI = '1900-01-01';
    DOI = req.body.DOI;
    
// insert new record into tblconsumers table
new sql.ConnectionPool(config).connect().then(pool => {   
    
querycode = 'INSERT INTO [tblConsumers] VALUES (' + "'" + req.body.ReceiveDate + "'," + "'" + DOI + "'," + "'" + req.body.CAAgent + "'," + "'" + req.body.ProductName.replace(/'/g, "''") + "'," + "'" + req.body.UPC + "'," + "'" + req.body.PackSize.replace(/'/g, "''") + "'," + "'" + req.body.UBD + "'," + "'" + req.body.Time + "'," + "'" + req.body.PlantNum + "'," + "'" + complaint + "'," + "'" + req.body.ReportCode + "'," + "'" + req.body.RetailerCustomer.replace(/'/g, "''") + "'," + "'" + req.body.HowReceived.replace(/'/g, "''") + "'," + "'" + req.body.ProductComments.replace(/'/g, "''") + "'," + "'" + req.body.Salutation + "'," + "'" + req.body.FirstName + "'," + "'" + req.body.MiddleInitial + "'," + "'" + req.body.LastName + "'," + "'" + req.body.CompanyName.replace(/'/g, "''") + "'," + "'" + req.body.Address1.replace(/'/g, "''") + "'," + "'" + req.body.Address2.replace(/'/g, "''") + "'," + "'" + req.body.City.replace(/'/g, "''") + "'," + "'" + req.body.StateorProvince + "'," + "'" + req.body.Zip + "'," + "'" + req.body.Country + "'," + "'" + req.body.PhoneNumber + "'," + "'" + req.body.CellNumber + "'," + "'" + req.body.EmailAddress + "'," + "'" + req.body.ConsumerNotes.replace(/'/g, "''") + "'," + "'" + req.body.StatusOfComplaint + "'," + "'" + foodService + "'," + "'" + govAgency + "'," + "'" + req.body.CustNum + "'," + "'" + req.body.GCNumber + "'," + "'" + req.body.TypeofComp.replace(/'/g, "''") + "'," + "'" + req.body.AmtofComp + "'," + "'" + req.body.Coupon55 + "'," + "'" + req.body.Coupon400 + "'," + "'" + req.body.Coupon800 + "'," + "'" + req.body.Coupon65 + "'," + "'" + req.body.Coupon900 + "'," + "'" + req.body.AmtofCoupon + "'," + "'" + req.body.CAResponse.replace(/'/g, "''") + "'," + null + ')';

return pool.query(querycode);
}).then(result => { 
    
res.render('success', {record: req.query.can, stat: "created", pageName: "Success", title: "Success"});
    
}).catch(err => {
    res.render('failure', {err: err, pageName: "Error", title: "Error", stat: "created"});
    main((req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, ""), err, (req.protocol + '://' + req.get('host') + req.originalUrl)).catch(console.error);
});
  } else {
 res.render('create', {pageName: "create", title: "New Record", caAgent: caAgent, plantNums: plantNums, arr: arr, reportCodes: reportCodes, howReceived: howReceived, salutations: salutations, countries: countries, stateOrProvince: stateOrProvince, complaintStatus: complaintStatus, prodInfo: prodInfo}); 
 } 
});

//pulls sql records for weekly plant report and puts them into csv format
app.get('/csv', function(req,res) {
    let connection = new sql.ConnectionPool(config, function (err) {
		if (err) {
			console.log(err.message);
		} else {

		//You can pipe mssql request as per docs
		var request = new sql.Request(connection);
		request.stream = true;
		request.query('select [CANumber] as CA#, LEFT(CONVERT(VARCHAR, ReceiveDate, 120), 10) as "Date Received", [ProductName] as Product, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', [PackSize] as Size, [UBD], [Time] ,[Plant], [Complaint],' + "('''' + tblConsumers.ReportCode+ '''') as Code" + ', [RetailerCustomer] as Customer, [HowReceived] as "How Received", [ProductComments] as "Product Comments" from tblConsumers where Plant = ' + "'" + req.query.plant + "'" + ' and ReceiveDate > ' + "'" + app.locals.weekAgoReformat + "'" + ' and ((tblConsumers.Complaint) =' + "'" + "Yes" + "'" + ') order by ReceiveDate desc');

		var stringifier = csv.stringify({header: true});
		//Then simply call attachment and pipe it to the response
        let filename = 'Weekly_Report_For_Plant-' + req.query.plant + '.csv';
		res.attachment(filename);
		request.pipe(stringifier).pipe(res);
        }
	});
})  

// queries page
app.get('/queries', function(req,res) {
    
    app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";

    global.isUserAuthorized = '';
    
    for(var i=0; i<global.authUser.length; i++) {
      if((req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") == global.authUser[i].UserName.toLowerCase()) {
            global.isUserAuthorized = true;
        }
    }
    
if(isUserAuthorized){
    if(req.query.qname == 'Reports By Code'){
        global.rc = true;
    } else {
        global.rc = false;
    }
    
    if(req.query.qname =='Reports by UPC'){
        global.upc = true;
    } else {
        global.upc = false;
    }
    let dateReformFrom = '';
    let dateReformTo = '';
    let rptResult = '';
    let rptResultSec = '';
    
    if(req.query.fromdate){
     dateReformFrom = req.query.fromdate[5] + req.query.fromdate[6] + "-" + req.query.fromdate[8] + req.query.fromdate[9] + "-" + req.query.fromdate[0] + req.query.fromdate[1] + req.query.fromdate[2] + req.query.fromdate[3];
     dateReformTo = req.query.todate[5] + req.query.todate[6] + "-" + req.query.todate[8] + req.query.todate[9] + "-" + req.query.todate[0] + req.query.todate[1] + req.query.todate[2] + req.query.todate[3]
    }
    
    if (req.query.qname == 'Reports By Code'){
        rptResult = true;
    } else {
        rptResult = false;
    }   
    
    if (req.query.qname == 'Reports by UPC'){
        rptResultSec = true;
    } else {
        rptResultSec = false;
    }
    
  res.render('queries', {pageName: "queries", title: "Queries", qs: dateReformFrom, qsto: dateReformTo, qname: req.query.qname, rptcode: req.query.ReportCode, upccode: req.query.UPC, rc: rc, reportCodes: reportCodes, upc: upc, fromdate: req.query.fromdate, todate: req.query.todate, rptResult: rptResult, rptResultSec: rptResultSec});
    } else {
        res.render('unauth', {uname: (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com", pageName: "Unauthorized", title: "Unauthorized"})
    }
});
    
    //pulls sql records for all the different queries
app.get('/csvqueries', function(req,res) {
    let connection = new sql.ConnectionPool(config, function (err) {
		if (err) {
			console.log(err.message);
		} else {

		//You can pipe mssql request as per docs
		var request = new sql.Request(connection);
		request.stream = true;
            global.query = '';
                   
                    if(req.query.qname == 'Reports By Date') {
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.ProductName,' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ')) ORDER BY tblConsumers.ReceiveDate;';
        }
            
        if(req.query.qname == 'Reports By Code') {
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (tblConsumers.ReceiveDate Between ' + "'" + req.query.fromdate + "'" +  ' And ' + "'" + req.query.todate + "'" + ') AND ReportCode = ' + "'" + req.query.rpt + "'" + ' ORDER BY tblConsumers.Plant';
        }
            
        if(req.query.qname == 'FO Only') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" +  ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ReportCode) Like ' + "'" + "4.%" + "'" + ' or (tblConsumers.ReportCode) Like ' + "'" + "3.%" + "'" + ')) ORDER BY tblConsumers.ReportCode';
        }  

        if(req.query.qname == 'Preference Only') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" +  ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ReportCode) Like ' + "'" + "11.%" + "'" + ')) ORDER BY tblConsumers.Time, tblConsumers.ReportCode;';
        }
         
        if(req.query.qname == 'Feedback Only') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" +  ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ReportCode) Like ' + "'" + "8.%" + "'" + ')) ORDER BY tblConsumers.Time, tblConsumers.ReportCode;';
        }
        
        if(req.query.qname == 'Packaging Only') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.Plant, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" +  ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ReportCode) Like ' + "'" + "7.%" + "'" + ')) ORDER BY tblConsumers.Time, tblConsumers.ReportCode;';
        }
    
        if(req.query.qname == 'Gift Card Accnts') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.ProductName, tblConsumers.[RetailerCustomer], tblConsumers.ProductComments, tblConsumers.FirstName, tblConsumers.LastName, tblConsumers.City, tblConsumers.ST, tblConsumers.Zip, tblConsumers.GCNumber, tblConsumers.AmtofComp FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" +  ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.GCNumber) Is Not Null)) ORDER BY tblConsumers.ReceiveDate;';
        }
         
        if(req.query.qname == 'Albertsons Safeway') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" +  ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.UPC) Like ' + "'" + "%21130.%" + "'" + ' Or (tblConsumers.UPC) Like ' + "'" + "%58200.%" + "'" + ' Or (tblConsumers.UPC) Like ' + "'" + "%81360.%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" +  ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Safeway%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" +  ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Albertson%" + "'" + ')) ORDER BY tblConsumers.ReportCode;';
        }        

        if(req.query.qname == 'ALDI') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.UPC) Like ' + "'" + "%41498.%" + "'" + ' Or (tblConsumers.UPC) Like ' + "'" + "%91000.%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%ALDI%" + "'" + ')) ORDER BY tblConsumers.ReportCode;';
        }
         
        if(req.query.qname == 'Boston Market') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ProductName) Like ' + "'" + "%Boston%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Boston%" + "'" + ')) ORDER BY tblConsumers.ReportCode;';
        }
            
        if(req.query.qname == 'Hormel') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.UPC) Like' + "'" + "%37600.%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Hormel%" + "'" + ')) ORDER BY tblConsumers.ReportCode;';
        }
            
       if(req.query.qname == 'Costco') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Costco%" + "'" + ')) ORDER BY tblConsumers.Time, tblConsumers.ReportCode;';
        }
                        
       if(req.query.qname == 'Sams Club') {    
        global.query = 'SELECT [CANumber] as CA#, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'"  + "%Sam%" + "'" + ')) ORDER BY tblConsumers.Time, tblConsumers.ReportCode;';
        }    
        
        if(req.query.qname == 'Sobeys') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.UPC) Like ' + "'" + "%68820.%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Sobey%" + "'" + ')) ORDER BY tblConsumers.ReportCode';
        }      
            
        if(req.query.qname == 'Subway') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Subway%" + "'" + ')) ORDER BY tblConsumers.ReportCode;';
        }
            
        if(req.query.qname == 'SYSCO') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.UPC) Like' + "'" + "%34730%" + "'" + '));';
        }
            
        if(req.query.qname == 'Walmart') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.[CustNum], tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.UPC) Like ' + "'" + "%81131%" + "'" + ' Or (tblConsumers.UPC) Like ' + "'" + "%78742%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Wal%" + "'" + ')) ORDER BY tblConsumers.Time, tblConsumers.ReportCode;';
        }
        
        if(req.query.qname == '7-11') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, tblConsumers.[CustNum], ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ProductName) Like ' + "'" + "%7-11%" + "'" + ' Or (tblConsumers.ProductName) Like ' + "'" + "%Select%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.UPC) Like ' + "'" + "%52548.%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%7-Eleven%" + "'" + ')) ORDER BY tblConsumers.Time, tblConsumers.ReportCode;';
        }       
            
        if(req.query.qname == 'MSB') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ProductName) Like ' + "'" + "%MSB%" + "'" + ' Or (tblConsumers.ProductName) Like ' + "'" + "%Main%" + "'" + ')) ORDER BY tblConsumers.ReportCode;';
        }
        
        if(req.query.qname == 'American Classics') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ProductName) Like ' + "'" + "%American%" + "'" + ' Or (tblConsumers.ProductName) Like ' + "'" + "%Classic%" + "'" + ' Or (tblConsumers.ProductName) Like ' + "'" + "%AC%" + "'" + ' Or (tblConsumers.ProductName) Like ' + "'" + "%AM%" + "'" + ')) ORDER BY tblConsumers.ReportCode;';
        }
                   
        if(req.query.qname == 'SMK') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ProductName) Like ' + "'" + "%SMK%" + "'" + ' Or (tblConsumers.ProductName) Like ' + "'" + "%Stonemill%" + "'" + ')) ORDER BY tblConsumers.ST;';
        }
                               
        if(req.query.qname == 'Snap Pack Only') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.UPC) = ' + "'" + "71117.02508" + "'" + ' Or (tblConsumers.UPC)= ' + "'" + "71117.02509" + "'" + '));';
        }
                                         
        if(req.query.qname == 'Co-Packed Reser Items') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers GROUP BY tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, tblConsumers.UPC, tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, tblConsumers.ReportCode, tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST HAVING (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.Plant) Like ' + "'" + "%00%" + "'" + '));';
        }
                                                     
        if(req.query.qname == 'Recall Report') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.DOI, tblConsumers.CAAgent, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.Salutation, tblConsumers.FirstName, tblConsumers.MiddleInitial, tblConsumers.LastName, tblConsumers.CompanyName, tblConsumers.Address1, tblConsumers.Address2, tblConsumers.City, tblConsumers.ST, tblConsumers.Zip, tblConsumers.Country, tblConsumers.PhoneNumber, tblConsumers.CellNumber, tblConsumers.EmailAddress, tblConsumers.[CustNum], tblConsumers.TypeofComp, tblConsumers.AmtofComp, tblConsumers.GCNumber FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ReportCode) Like ' + "'" + "%.99%" + "'" + '));';
        }
                                                                
        if(req.query.qname == 'Canada Complaints') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST, tblConsumers.Country FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.ST) In (' + "'" + "AB" + "'" + ", '" + "BC" + "', " + "'" + "MB" + "', " + "'" + "NB" + "', " + "'" + "NS" + "', " + "'" + "NT" + "', " + "'" + "NU" + "', " + "'" + "ON" + "', " + "'" + "PE" + "', " + "'" + "QC" + "', " + "'" + "SK" + "', " + "'" + "YT" + "'" + ')) AND ((tblConsumers.Country) Like ' + "'" + "%Canada%" + "'" + ')) OR (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Sobey%" + "'" + ' Or (tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Loblaw%" + "'" + ' Or (tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Westrow%" + "'" + ' Or (tblConsumers.[RetailerCustomer]) Like ' + "'" + "%Canada%" + "'" + '));';
        }
                                                                            
        if(req.query.qname == 'Template for Imports') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.DOI, tblConsumers.CAAgent, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.Salutation, tblConsumers.FirstName, tblConsumers.MiddleInitial, tblConsumers.LastName, tblConsumers.CompanyName, tblConsumers.Address1, tblConsumers.Address2, tblConsumers.City, tblConsumers.ST, tblConsumers.Zip, tblConsumers.Country, tblConsumers.PhoneNumber, tblConsumers.CellNumber, tblConsumers.EmailAddress, tblConsumers.ConsumerNotes, tblConsumers.StatusOfComplaint, tblConsumers.FoodService, tblConsumers.GovAgency, tblConsumers.[CustNum], tblConsumers.GCNumber, tblConsumers.TypeofComp, tblConsumers.AmtofComp, tblConsumers.Coupon55, tblConsumers.Coupon400, tblConsumers.Coupon800, tblConsumers.Coupon65CAN, tblConsumers.Coupon900CAN, tblConsumers.AmtofCoupon, tblConsumers.CAResponse FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ')) ORDER BY tblConsumers.ReceiveDate, tblConsumers.Time;';
        }
                                                                                        
        if(req.query.qname == 'Reports by UPC') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.UPC) like ' + "'" + "%" + req.query.rpt + "%" + "'" + ')) ORDER BY tblConsumers.ReportCode;';
        }
                                                                                                    
        if(req.query.qname == 'DSD') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.Plant) Like' + "'" + "DSD" +"'" + '));';
        }
        
        if(req.query.qname == 'Leaker Project 2014') {    
        global.query = 'SELECT tblConsumers.CustomerID, tblConsumers.ReceiveDate, tblConsumers.ProductName, ' + "('''' + tblConsumers.UPC + '''') as UPC" + ', tblConsumers.PackSize, tblConsumers.UBD, tblConsumers.Time, tblConsumers.Plant, tblConsumers.Complaint, ' + "('''' + tblConsumers.ReportCode+ '''') as ReportCode" + ', tblConsumers.[RetailerCustomer], tblConsumers.HowReceived, tblConsumers.ProductComments, tblConsumers.City, tblConsumers.ST FROM tblConsumers WHERE (((tblConsumers.ReceiveDate) Between ' + "'" + req.query.fromdate + "'" + ' And ' + "'" + req.query.todate + "'" + ') AND ((tblConsumers.Plant)=' + "'" + 21 + "'" + ' Or (tblConsumers.Plant)= ' + "'" + 23 + "'" + ' Or (tblConsumers.Plant)=' + "'" + 50 + "'" + ') AND ((tblConsumers.ReportCode) Like ' + "'" + "7.%" + "'" + '));';
        }
            
		request.query(query);
		var stringifier = csv.stringify({header: true});
		//Then simply call attachment and pipe it to the response
        let filename = req.query.qname + ' Query for dates ' + req.query.fromdate + ' to ' + req.query.todate + '.csv';
		res.attachment(filename);
		request.pipe(stringifier).pipe(res);
        }
	});
}) 
    
// failure page
app.use('/failure', function(req,res) {
    app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
  res.render('failure', {pageName: "Error", title: "Error"}); 
});

// success page
app.use('/success', function(req,res) {
    app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
  res.render('success', {}); 
});
    
// unauthorized page
app.use('/unauth', function(req,res) {
    app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
  res.render('unauth', {}); 
});
    
    
// define 404 handler
app.use(function(req,res) {
        app.locals.username = (req.connection.user).replace(/JVAPP\\/g, "").toLowerCase().replace(/resers\\/g, "") + "@resers.com";
  res.render('404', {pageName: "404", title: "404 - not found."} ); 
});

app.listen(app.get('port'), function() {
 console.log('Application has started');
});