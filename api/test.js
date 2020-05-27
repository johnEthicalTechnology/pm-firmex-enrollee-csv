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

testD.forEach(userDetails => {
  console.log('testing', userDetails);
  importUsersWorksheet.addRow(userDetails)
})

usersAddToFirmexWorkbook.xlsx.writeFile('test.xlsx').catch((err) => console.error('err', err))