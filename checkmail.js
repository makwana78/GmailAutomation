var webdriverio = require('webdriverio');
var options = { desiredCapabilities: { browserName: 'chrome' } };
var client = webdriverio.remote(options);
var msgsubject;
var assert = require('assert');

client
	// Login as Second User and Check the Message
    .init()
    .url('https://mail.google.com/')
	.windowHandleMaximize()
	//Read User Name from Excel and enter into the User Name field
    .setValue('#Email', readTextFile('B3'))
	//Click Next Button
    .click('#next')
	.pause('1000')
	//Read Password from Excel and enter into the Password field
	.setValue('#Passwd', readTextFile('C3'))
	//Click Sign In Button
    .click('#signIn')
	.pause('10000')
	//Check for the Mail
	.getText('div.Cp').then(function(text) {
	CheckMail(text);
    
})
	
	.pause('2000')
	//Clicking Account Picture
	.click('span.gb_3a.gbii')
	.pause('1000')
	//Clicking Sign Out
	.click('#gb_71')
	.pause('5000')
	//Close the Browser
	.close()	

//Function for reading Excel Data
function readTextFile(address_of_cell)
{
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('users.xlsx');
var first_sheet_name = workbook.SheetNames[0];
var worksheet = workbook.Sheets[first_sheet_name];
var desired_cell = worksheet[address_of_cell];
var desired_value = desired_cell.v;
return desired_value;
}

//Function for Checking Mail 
function CheckMail(InBoxMails)
{
    var MailSubject = readTextFile('D2').trim();
	var AllMail = new String(InBoxMails);
	if(AllMail.includes(MailSubject))
	{
		console.log('Mail Received');
	}
	else
	{
		console.log('Mail Not Received');
	}
	
	
	
}

