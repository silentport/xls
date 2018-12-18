var xlsx = require('node-xlsx');
var fs = require('fs');
var obj = xlsx.parse(__dirname + '/input.xlsx'); 
var data = obj[0].data; 
var arr = [
    [],
    [],
    [],
    []
];
let res = [];
let arrRes = [];
let n = 0;
console.log(data[1][2])
data.forEach(item => {
    if (item.includes('产品部')) {
        arr[0].push(item)
    }
    if (item.includes('运营部')) {
        arr[1].push(item)
    }
    if (item.includes('研发部')) {
        arr[2].push(item)
    }
    if (item.length < 4 || item.includes('')) {
        arr[3].push(item)
    }
})

let find = (arr, n, m) => {
    let res = [];
    let c1 = c2 = 0;
    for (let i = 0; i < arr.length; i++) {
        if (arr[i].includes('男') && c1 < n) {
            res.push(arr[i]);
            arr.splice(i, 1);
            c1++;
        }
        if (arr[i].includes('女') && c2 < m) {
            res.push(arr[i]);
            arr.splice(i, 1);
            c2++;
        }
        if (c1 >= n && c2 >= m) {
            return res;
        }
    }
    return res;
}
for (let j = 0; j < 10; j++) {
    res[j] = [];
    let item = find(arr[1], 3, 3);
    item = item.concat(find(arr[2], 10, 3));
    item = item.concat(find(arr[3], 0, 1));
    res[j].push(item);
}
var c1 = 0
for (let i = 0; i < arr[2].length; i++) {
    if (arr[2][i].includes('男') && c1 < 3) {
        res[9][0].push(arr[2][i]);
        arr[2].splice(i, 1);
        c1++;
    }
}



const rest = arr[0].concat(arr[1], arr[2], arr[3]);


res.forEach((i) => {

    n++;
    i[0].push(rest.pop())
    if (n > 6) {
        i[0].push(rest.pop())
    }
    const arr = [`第${n}组名单--------------------- ${i[0].length}人`]
    i[0].unshift(arr);
    arrRes = arrRes.concat(i[0]);
    console.log(arrRes)

})

let buffer = xlsx.build([{
    name: 'sheet1',
    data: arrRes
}]);
console.log(rest)
fs.writeFileSync(`res.xlsx`, buffer, {
    'flag': 'w'
});