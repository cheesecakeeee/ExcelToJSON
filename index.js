let path = require('path');
//使用ejsexcel读取excel文件  npm install ejsexcel --save
let ejsExcel = require('ejsexcel');
let fs = require('fs');
//读取excel
let exBuf = fs.readFileSync(__dirname + '/Input/translatorProps.xlsx');
let _data = [];
// 中文json
let _Obj = {};
// 日文json
let _JObj = {};
// sheet名称翻译数组「从右往左顺序」
let sheetName = ['Color', 'Occasion', 'Style', 'Category']
ejsExcel.getExcelArr(exBuf).then(exlJson => {
    //获取excel数据
    let workBook = exlJson;
    //获取成功后
    // console.log(workBook)
    // 首字母大写
    function toUpCase(s) {
        return s.toLowerCase().replace(/\b([\w|']+)\b/g, function (word) {
            return word.replace(word.charAt(0), word.charAt(0).toUpperCase());
        });
    }

    // 首字母大写并移除空格
    function removeSpace(str) {
        return toUpCase(str).replace(/\s*/g, "");
    }

    // 首字母大写并移除空格去掉 `s
    function removedots(str) {
        return (removeSpace(str).replace("'s", ""))
    }

    // 剔除每个表所有sheet的空行
    for (let i = 0; i < sheetName.length; i++) {
        // 拿到每个sheet
        workSheets = workBook[i];
        // 去掉空行
        workSheets.forEach((item, index) => {
            if (item.length == 0) {
                workSheets.splice(index, 1)
            }
        });
        workSheets.forEach((item, index) => {
            //从第二行开始插入，避免连表头也插入_data里面
            if (index > 0) {
                // item的索引值为表的列数，从0开始
                // 1为英文，2为中文，3为日文
                item[1] = removedots(item[1])
                // // key为英文，value为中文
                _Obj[sheetName[i] + `.` + item[1]] = item[2]
                // key为英文，value为日文
                _JObj[sheetName[i] + `.` + item[1]] = item[3]
            }
        })
    }
    _data.push(_Obj)
    //导出中文js的路径
    let newfilepath = path.join(__dirname, "/Output/test.js");
    // 导出日文json
    let newfileJapanpath = path.join(__dirname, "/Output/Japantest.js");
    // //写入中文js文件
    fs.writeFileSync(newfilepath, 'let _data=' + JSON.stringify(_Obj) + ';export {_data}');
    // //写入日文js文件
    fs.writeFileSync(newfileJapanpath, 'let _data=' + JSON.stringify(_JObj) + ';export {_data}');
}).catch(error => {
    //打印获取失败信息
    console.log("读取错误!");
    console.log(error);
});
