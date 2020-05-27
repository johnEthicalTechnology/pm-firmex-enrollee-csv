const nodemailer = require("nodemailer");
const ExcelJS = require('exceljs');

const testD = [
  {
    emailAddress: "john.stewart@ethicaltechnology.co",
    firstName: "Fist",
    lastName: "Attendee",
    company: "TEST COMPANY",
    office: "PM Sydney"
  },
  {
    emailAddress: "john.stewart@ethicaltechnology.co",
    firstName: "Second",
    lastName: "Attendee",
    company: "TEST COMPANY",
    office: "PM Sydney"
  },
  {
    emailAddress: "john.stewart@ethicaltechnology.co",
    firstName: "Third",
    lastName: "Attendee",
    company: "TEST COMPANY",
    office: "PM Sydney"
  }
]

// const usersAddToFirmexWorkbook = new ExcelJS.Workbook()
const usersAddToFirmexWorkbook = new ExcelJS.Workbook()
const importUsersWorksheet = usersAddToFirmexWorkbook.addWorksheet('Import Users')
importUsersWorksheet.columns = [
  {
    header: 'Email Address',
    key: 'emailAddress'
  },
  {
    header: 'First Name',
    key: 'firstName'
  },
  {
    header: 'Last Name',
    key: 'lastName'
  },
  {
    header: 'Company',
    key: 'company'
  },
  {
    header: 'Office',
    key: 'office'
  }
]

// async..await is not allowed in global scope, must use a wrapper
async function main(email) {
  testD.forEach(userDetails => {
    console.log('testing', userDetails);
    importUsersWorksheet.addRow(userDetails)
  })
  const buffer = await usersAddToFirmexWorkbook.xlsx.writeBuffer()
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
    attachments: [
      {
        filename: 'firmexSpreedsheet.xlsx',
        content: buffer
      }
    ]
  });

  console.log("Message sent: %s", info.messageId);
  // Message sent: <b658f8ca-6296-ccf4-8306-87d57a0b4321@example.com>

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
  // const enrolleeInfoForFirmex = JSON.parse(req.body);
  const test1 = await main('brett.handley@prioritymanagement.com.au');
  console.log('test1', test1);



  res.json({body: "Did it work!?"});
}