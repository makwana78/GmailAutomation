// Webdriver.io
var webdriverio = require('webdriverio');
var options = { desiredCapabilities: { browserName: 'chrome' } };
var client = webdriverio.remote(options);
var msgsubject;

client
	// Login as First User and Send the Message
    .init()
	// Access GMail
    .url('https://mail.google.com/')
	//.windowHandleMaximize()
	// Read User Name from Excel and enter into the User Name field
    .setValue('#Email', readTextFile('B2'))
	//Clicks Next Button
    .click('#next')
	.pause('1000')
	// Read Password from Excel and enter into the Password field
	.setValue('#Passwd', readTextFile('C2'))
	//Clicks SignIn Button	
    .click('#signIn')
	.pause('10000')
        // clicking on compose button
	.click('div.T-I.J-J5-Ji.T-I-KE.L3')
	.pause('1000')
	// Read To-Address from Excel and enter into the To field
	.setValue('textarea[name=\'to\']', readTextFile('B3'))
	.keys('Enter')
	// Read Subject from Excel and enter into the Subject field
	.setValue('input[name=\'subjectbox\']', readTextFile('D2'))
	// Read Message Body from Excel and enter into the Message Body field
	.setValue('div[role=\'textbox\']', readTextFile('E2'))
	//Clicks Send Button
	.click('td.gU.Up')
	.pause('2000')
	//Clicks profile  Accounts Button 
	.click('span.gb_3a.gbii')
	.pause('1000')
	//Clicks Sign Out Button
	.click('#gb_71')
	.pause('5000')
	.close()	

//Function for read Excel	
function readTextFile(address_of_cell)
{
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('usersnew.xls');
var first_sheet_name = workbook.SheetNames[0];
var worksheet = workbook.Sheets[first_sheet_name];
var desired_cell = worksheet[address_of_cell];
var desired_value = desired_cell.v;
return desired_value;
}

	
    //.end();
