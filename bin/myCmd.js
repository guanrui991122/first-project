#!/usr/bin/env node
"use strict";
function run(fileName) {
  const path = require('path')
  let filePath = path.join(process.cwd(),fileName)
  // 读取文件
  if (typeof require !== ' undefined ') XLSX = require('xlsx');
  // const workbook1 = XLSX.readFile('./Excel/新集.xlsx');
  // const workbook2 = XLSX.readFile('./Excel/陈集.xlsx');
  const workbook2 = XLSX.readFile(filePath)
  // const workbook2 = XLSX.readFile(args[2]);

  // 读取sheet1
  const first_sheet_name = workbook2.SheetNames[0];

  const first_worksheet = workbook2.Sheets[first_sheet_name];

  const list1 = XLSX.utils.sheet_to_json(first_worksheet);

  // 读取sheet2
  const second_sheet_name = workbook2.SheetNames[1];

  const second_worksheet = workbook2.Sheets[second_sheet_name];

  let list2 = XLSX.utils.sheet_to_json(second_worksheet);

  // new_list2 = list2.map()
  // 设置正则表达式
  const reg = /[0-9]{12,}[A-Z]{2}[0-9]{5}\s[\u4e00-\u9fa5]{2,}/g

  const list2New = list2
    // 过滤器检查是否处理成功
    // .filter((it) => it.DJH == "321081107212JC00071")
    .map((item, i) => {       //item:每行数据
      // 浅拷贝
      const itemCopy = Object.assign({}, item);
      // 将四个变量放在数组中进行遍历
      ["ZDSZB", "ZDSZD", "ZDSZN", "ZDSZX"].forEach((key) => {
        let str = item[key];
        // 正则匹配
        let group = str.match(reg);   //group:符合正则表达式匹配
        if (group != null) {
          group.forEach((match) => {
            const findItem = list1.find((it) => it["原"] == match);
            // ES6的解构
            // const findItem = list1.find(原 => 原 == match); 

            if (findItem != undefined) {
              // console.log(
              //   match,
              //   findItem["现"],
              //   itemCopy[key].replace(match, findItem["现"])
              // );
              itemCopy[key] = itemCopy[key].replace(match, findItem["现"]);
            }
          })
        }
      });
      return itemCopy;
    });

  // 将数据导入到新表中再导出到工作簿中
  // const sheetNew = XLSX.utils.json_to_sheet(list2New);
  // // console.log(sheetNew);
  // var ws_name = "SheetJS";
  // XLSX.utils.book_append_sheet(workbook2, sheetNew, ws_name);
  // XLSX.writeFile(workbook2, 'out.xlsx');

  // 直接更新sheet2
  workbook2.Sheets["Sheet2"] = XLSX.utils.json_to_sheet(list2New)
  let outPath = path.join(process.cwd(),'out1.xlsx')
  XLSX.writeFile(workbook2, outPath);
}
// 获取当前路径的不同情况
// console.log(process.execPath)
// console.log(__dirname)
// console.log(process.cwd())
run(process.argv[2]);