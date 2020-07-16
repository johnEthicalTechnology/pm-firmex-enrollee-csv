const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');

module.exports = async (req, res) => {
  //* 1) Parse JSON
  // enrolleeInfoForFirmex
  // [
  //   {
  //     'emailAddress':'john.stewart@ethicaltechnology.co',
  //     'firstName':'Second',
  //     'lastName':'Attendee',
  //     'company':'TEST COMPANY',
  //     'office':'PM Sydney'
  //   },
  //   ...
  // ]
  console.log('Recieved req.body.data', req.body.data);

  const enrolleeInfoForFirmex = JSON.parse(req.body.data)

  //* 2) Setup workbook, and worksheet (actual spreadsheet).
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
  //* 3) Add user details to worksheet
  enrolleeInfoForFirmex.forEach(userDetails => {
    console.log(`Adding ${userDetails.firstName[0]}. ${userDetails.lastName} to worksheet`)
    importUsersWorksheet.addRow(userDetails)
  })
  try {
    //* 4) Create buffer
    const buffer = await usersAddToFirmexWorkbook.xlsx.writeBuffer()
    //* 5) Create reusable transporter object using the default SMTP transport
    const transporter = nodemailer.createTransport({
      host: 'smtp.zoho.com',
      port: 465,
      secure: true, // true for 465, false for other ports
      auth: {
        user: 'brett.handley@prioritymanagement.com.au',
        pass: process.env.EMAIL_PW,
      },
    })

    //* 6) Send mail with defined transport object
    const companyName = enrolleeInfoForFirmex[0].company
    const emailRes = await transporter.sendMail({
      from: `'Priority Management Sydney' <brett.handley@prioritymanagement.com.au>`,
      to: 'sandir@prioritymanagement.com',
      cc: 'materials@prioritymanagement.com.au',
      subject: `Firmex spreadsheet for ${companyName}`,
      text: `Dear Sandi,/r This is the Firmex spreadsheet containing users from ${companyName}/r Regards,/rAdriana Parinetto`,
      html: `<p>Dear Sandi,</p><p>This is the Firmex spreadsheet containing users from ${companyName}</p><p>Regards,<br/>Adriana Parinetto</p>`,
      attachments: [
        {
          filename: 'firmexSpreedsheet.xlsx',
          content: buffer
        }
      ]
    })
    console.log('Message sent:', emailRes.messageId)
    res.json({body: `Message sent: ${emailRes.messageId}`})
  } catch (error) {
    console.error('Error:', error)
    res.json({body: `Error: ${error}`})
  }
}