const xlsx = require('node-xlsx').default;
const fs = require('fs');

// Parse a file
const workSheetsFromFile = xlsx.parse(`${__dirname}/files/data.xlsx`);
console.log(workSheetsFromFile[0].data[0]);
//create a for loop thats run to set states
let stateArray = [['', 'Nan']];
for (i = 0; i < workSheetsFromFile[0].data.length; i++) {
  if (workSheetsFromFile[0].data[i + 1] != undefined) {
    if (
      workSheetsFromFile[0].data[i + 1][1] > workSheetsFromFile[0].data[i][1]
    ) {
      console.log('greater');
      stateArray.push([workSheetsFromFile[0].data[i + 1][0], 'W1']);
    } else if (
      workSheetsFromFile[0].data[i + 1][1] < workSheetsFromFile[0].data[i][1]
    ) {
      console.log('less');
      stateArray.push([workSheetsFromFile[0].data[i + 1][0], 'W3']);
    } else if (
      workSheetsFromFile[0].data[i + 1][1] === workSheetsFromFile[0].data[i][1]
    ) {
      console.log('same');
      stateArray.push([workSheetsFromFile[0].data[i + 1][0], 'W2']);
    }
  }
}
console.log(stateArray);
console.log(stateArray.length);

var buffer = xlsx.build([{ name: 'mytest', data: stateArray }]);

try {
  fs.writeFileSync(`test.xlsx`, buffer);
  // file written successfully
  console.log('success');
} catch (err) {
  console.error(err);
}
