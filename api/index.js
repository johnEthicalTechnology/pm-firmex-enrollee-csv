const fs = require('fs')
const nodemailer = require("nodemailer");

// async..await is not allowed in global scope, must use a wrapper
async function main(email) {
  // Generate test SMTP service account from ethereal.email
  // Only needed if you don't have a real mail account for testing
  // let testAccount = await nodemailer.createTestAccount();

  // create reusable transporter object using the default SMTP transport
  let transporter = nodemailer.createTransport({
    host: "smtp.zoho.com",
    port: 465,
    secure: true, // true for 465, false for other ports
    auth: {
      user: "brett.handley@prioritymanagement.com.au", // generated ethereal user
      pass: process.env.EMAIL_PW, // generated ethereal password
    },
  });

  // send mail with defined transport object
  let info = await transporter.sendMail({
    from: `"Fred Foo 👻" <${email}>`, // sender address
    to: "john.stewart@ethicaltechnology.co", // list of receivers
    subject: "Hello ✔", // Subject line
    text: "Hello world?", // plain text body
    html: "<b>Hello world?</b>", // html body
  });

  console.log("Message sent: %s", info.messageId);
  // Message sent: <b658f8ca-6296-ccf4-8306-87d57a0b4321@example.com>

  // Preview only available when sending through an Ethereal account
  // console.log("Preview URL: %s", nodemailer.getTestMessageUrl(info));
  // // Preview URL: https://ethereal.email/message/WaQKMgKddxQDoou...
}

function writeToCSVFile(users) {
  const filename = 'output.csv';
  fs.writeFile(filename, extractAsCSV(users), err => {
    if (err) {
      console.log('Error writing to csv file', err);
    } else {
      console.log(`saved as ${filename}`);
    }
  });
}

function extractAsCSV(users) {
  const header = ["Email Address, First Name, Last Name, Company, Office"];
  const rows = users.map(user =>
     `${user.emailAddress}, ${user.firstName}, ${user.lastName}, ${user.company}, ${user.office}`
  );
  return header.concat(rows).join("\n");
}




module.exports = async (req, res) => {

  // console.log('REQUEST body', req.body);
  // enrolleeInfoForFirmex
  // [
  //   {
  //     "Email Address":"john.stewart@ethicaltechnology.co",
  //     "First Name":"Second",
  //     "Last Name":"Attendee",
  //     "Company":"TEST COMPANY",
  //     "Office":"PM Sydney"
  //   },
  //   ...
  // ]
  console.log('testing before send');

  const test = await main('bretttest@zohomail.com');

  console.log('testing AFTER send brett test', test);
  // const enrolleeInfoForFirmex = JSON.parse(req.body);
  const test1 = await main('brett.handley@prioritymanagement.com.au');
  console.log('test1', test1);



  res.json({body: "Did it work!?"});
}