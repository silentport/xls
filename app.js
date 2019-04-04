var xlsx = require ('node-xlsx');
var fs = require ('fs');
var obj = xlsx.parse (__dirname + '/input.xlsx');
var data = obj[0].data;
var len = data[1].length;
var arr = [data[0].slice (), data[1].slice ()];
let res = {};

let arrRes = [];
function concat (arr1, arr2) {
  let len = Math.max (arr1.length, arr2.length);
  let res = [];
  for (let i = 0; i < len; i++) {
    if (arr1[i] == arr2[i]) {
      res.push (arr1[i]);
    } else {
      if (arr1[i] == undefined && arr2[i] != undefined) {
        res.push (arr2[i]);
      } else if (arr2[i] == undefined && arr1[i] != undefined) {
        res.push (arr1[i]);
      } else {
        res.push (undefined);
      }
    }
  }
  return res;
}
data.forEach ((item, i) => {
  if (i > 1) {
    if (!res[item[1]]) {
      res[item[1]] = item;
    } else if (item[1] != undefined) {
      res[item[1]] = concat (res[item[1]], item);
    }
  }
});
const keys = Object.keys (res);
keys.forEach (k => {
  if (k) {
    arr.push (res[k]);
  }
});

let buffer = xlsx.build ([
  {
    name: 'sheet1',
    data: arr,
  },
]);

fs.writeFileSync (`res.xlsx`, buffer, {
  flag: 'w',
});
