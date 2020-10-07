const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const { join } = require('path')

module.exports = async (req, res) => {
  //* 1) Parse JSON
  // enrolleeInfoForFirmex
  // [
  //   {
  //     'emailAddress':'john.stewart@ethicaltechnology.co',
  //     'firstName':'Second',
  //     'lastName':'Attendee',
  //     'company':'TEST COMPANY',
  //     'office':'PM Sydney',
  //     'dateOfTraining: '17/09/20'
  //   },
  //   ...
  // ]
  console.log('Recieved req.body.data', req.body.data);

  const enrolleeInfoForFirmex = JSON.parse(req.body.data)

  //* 2) Setup workbook, and worksheet (actual spreadsheet).
  const firmexTemplateWb = new ExcelJS.Workbook()
  const firmexSpreedsheet = await firmexTemplateWb.xlsx.readFile(join(__dirname, '_files', 'firmexTemplate.xlsx'))
  const importUsersWorksheet = firmexSpreedsheet.getWorksheet('Import Users')

  //* 3) Add user details to worksheet
  enrolleeInfoForFirmex.forEach((userDetails, index) => {
    console.log(`Adding ${userDetails.firstName[0]}. ${userDetails.lastName} to worksheet`)
    const rowValue = index + 2
    importUsersWorksheet.getCell(`A${rowValue}`).value = userDetails.emailAddress
    importUsersWorksheet.getCell(`B${rowValue}`).value = userDetails.firstName
    importUsersWorksheet.getCell(`C${rowValue}`).value = userDetails.lastName
    importUsersWorksheet.getCell(`D${rowValue}`).value = userDetails.company
    importUsersWorksheet.getCell(`E${rowValue}`).value = userDetails.office
    importUsersWorksheet.getCell(`F${rowValue}`).value = 'adriana.parinetto@prioritymanagement.com.au'
    importUsersWorksheet.getCell(`G${rowValue}`).value = 'Yes'
  })
  try {
    //* 4) Create buffer
    const buffer = await firmexTemplateWb.xlsx.writeBuffer()
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
    console.log('Company Name ', companyName)
    const startDateOfTraining = enrolleeInfoForFirmex[0].dateOfTraining
    console.log('Date of training ', startDateOfTraining)
    const emailRes = await transporter.sendMail({
      from: `'Priority Management Sydney' <brett.handley@prioritymanagement.com.au>`,
      to: 'john.stewart@ethicaltechnology.co',
      cc: 'materials@prioritymanagement.com.au',
      subject: `Firmex spreadsheet for ${companyName}`,
      text: `Dear PM Admin,/r This is the Firmex spreadsheet containing users from ${companyName}/r Regards,/rZoho Automation`,
      html: `<p>Dear PM Admin,</p><p>This is the Firmex spreadsheet containing users from ${companyName}</p><p>Regards,<br/>Zoho Automation</p>`,
      attachments: [
        {
          filename: `${companyName} - ${startDateOfTraining}.xlsx`,
          content: buffer
        }
      ]
    })
    console.log('Message sent: ', emailRes.messageId)
    res.json({body: `Message sent: ${emailRes.messageId}`})
  } catch (error) {
    console.error('Error:', error)
    res.json({body: `Error: ${error}`})
  }
}