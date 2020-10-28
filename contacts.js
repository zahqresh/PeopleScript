'use strict';
const path = require('path');
const fs = require('fs');
const {
  google
} = require('googleapis');
const {
  authenticate
} = require('@google-cloud/local-auth');
const readXlsxFile = require('read-excel-file/node');
const people = google.people('v1');
var XLSX = require('xlsx');
const {
  cloudidentity
} = require('googleapis/build/src/apis/cloudidentity');
var workbook = XLSX.readFile('test.xlsx');
var sheet_name_list = workbook.SheetNames;
var contacts = [];

//module to extract data from the xlsx file as json and save them to google contacts list
XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]).forEach(n => {
  // ,n.CustomerShippingAddress,n.CustomerPhoneNumber,n.shiptoAddress1, n.shiptoAddress2,n.City, n.State, n.Zip

  contacts.push({
    name: n.CustomerName,
    ShippingAddress: n.CustomerShippingAddress,
    CustomerPhoneNumber: n.CustomerPhoneNumber,
    address1: n.ShiptoAddress1,
    city: n.City,
    State: n.State,
    zip: n.Zip,
    po:n.PO
  })
});
//Creat contact in google contact list
async function runSample() {
  // Obtain user credentials to use for the request
  const auth = await authenticate({
    keyfilePath: path.join(__dirname, './credentials.json'),
    scopes: ['https://www.googleapis.com/auth/contacts'],
  });
  google.options({
    auth
  });

  // Create a new contact
  // https://developers.google.com/people/api/rest/v1/people/createContact


  //loop in array to create contact one by one
  contacts.forEach(n => {
    (async () => {
      //module to create contact
      const {
        data: newContact
      } = await people.people.createContact({
        requestBody: {
          names: [
            {
              displayName: `${n.name}`,
              givenName: `${n.name}`
            },
          ],
        phoneNumbers: [{
            "value": `${n.CustomerPhoneNumber}`
          }],
        addresses: [{
            "poBox": `${n.po}`,
            "streetAddress": `${n.ShippingAddress}`,
            "extendedAddress": `${n.address1}`,
            "city": `${n.city}`,
            "region": `${n.State}`,
            "postalCode": `${n.zip}`,
            "countryCode": 'unkown'
          }]
        },
      });
      console.log('Number of contacts created Successfully! '+ contacts.length);
    })();
  })
}

if (module === require.main) {
  runSample().catch(console.error);
  console.log('Error Occured! Fix fields!');
}
