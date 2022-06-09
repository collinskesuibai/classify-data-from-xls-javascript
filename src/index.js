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
      stateArray.push([workSheetsFromFile[0].data[i + 1][0], 'W1']);
    } else if (
      workSheetsFromFile[0].data[i + 1][1] < workSheetsFromFile[0].data[i][1]
    ) {
      stateArray.push([workSheetsFromFile[0].data[i + 1][0], 'W3']);
    } else if (
      workSheetsFromFile[0].data[i + 1][1] === workSheetsFromFile[0].data[i][1]
    ) {
      stateArray.push([workSheetsFromFile[0].data[i + 1][0], 'W2']);
    }
  }
}
console.log(stateArray);
//counting number of state transitions
let W1_W1 = 0;
let W1_W2 = 0;
let W1_W3 = 0;
let W2_W1 = 0;
let W2_W2 = 0;
let W2_W3 = 0;
let W3_W1 = 0;
let W3_W2 = 0;
let W3_W3 = 0;

for (i = 0; i < stateArray.length; i++) {
  console.log('running');
  if (stateArray[i + 1] != undefined) {
    if (stateArray[i][1] == 'W1' && stateArray[i + 1][1] == 'W1') {
      W1_W1++;
    } else if (stateArray[i][1] == 'W1' && stateArray[i + 1][1] == 'W2') {
      W1_W2++;
    } else if (stateArray[i][1] == 'W1' && stateArray[i + 1][1] == 'W3') {
      W1_W3++;
    } else if (stateArray[i][1] == 'W2' && stateArray[i + 1][1] == 'W1') {
      W2_W1++;
    } else if (stateArray[i][1] == 'W2' && stateArray[i + 1][1] == 'W2') {
      W2_W2++;
    } else if (stateArray[i][1] == 'W2' && stateArray[i + 1][1] == 'W3') {
      W2_W3++;
    } else if (stateArray[i][1] == 'W3' && stateArray[i + 1][1] == 'W1') {
      W3_W1++;
    } else if (stateArray[i][1] == 'W3' && stateArray[i + 1][1] == 'W2') {
      W3_W2++;
    } else if (stateArray[i][1] == 'W3' && stateArray[i + 1][1] == 'W3') {
      W3_W3++;
    }
  }
}
console.log(stateArray.length);
console.log(W1_W1, W1_W2, W1_W3, W2_W1, W2_W2, W2_W3, W3_W1, W3_W2, W3_W3);

var buffer = xlsx.build([{ name: 'mytest', data: stateArray }]);

try {
  fs.writeFileSync(`test.xlsx`, buffer);
  // file written successfully
  console.log('success');
} catch (err) {
  console.error(err);
}
