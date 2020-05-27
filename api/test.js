const fs = require('fs')

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


const foop = writeToCSVFile(testD);

console.log('foop', foop);
