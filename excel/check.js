/*验证本地路径是否与表格路径一致*/
var xlsx = require('node-xlsx');
var fs = require('fs');
    stdin = process.stdin;
    stdout = process.stdout;
var path= require('path');
var stats = [];
var mulu = process.cwd();
var mulu2 = process.cwd();
var lock2 = false;
var locks =true;
var excel = [[]];
var excel2 = [];
var localTest = [];
var num = null;
// 导入excel文件
const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(`${__dirname}/20170419.xlsx`));
const workSheetsFromFile = xlsx.parse(`${__dirname}/20170419.xlsx`);

function showAll() {
    // 读取当前文件夹
    fs.readdir(mulu, function (err, files) {
        console.log('')
        if (!files.length) {
            return console.log(' \033[31m No files to show!\033[39m\n');
        }
        getExcel();
        //workSheetsFromFile[6].data[i][4]6代表左下角坐标，i代表列4代表行
        function getExcel() {
            //一个一个全文件夹搜索
            // for(var j in workSheetsFromFile[8].data[0]){
            //     if(workSheetsFromFile[8].data[0][j]=='素材文件名') {
            //         num = j;
            //         for (var i in workSheetsFromFile[8].data) {
            //             // if(workSheetsFromFile[0].data[i][2]=='打击乐'){
            //             // excel.push(workSheetsFromFile[8].data[i][num])
            //             // }
            //             if(i!=0&&workSheetsFromFile[8].data[i][num]){
            //                 // excel.push([workSheetsFromFile[8].data[i][num].split('/')])
            //                 excel.push(workSheetsFromFile[8].data[i][num])
            //             }
            //
            //         }
            //     }
            // }




            //全文件夹搜索 会卡
            for(var i in workSheetsFromFile){
              for(var j in workSheetsFromFile[i].data[0]){
                  if (workSheetsFromFile[i].data[0][j] == '素材文件名') {
                      num = j;
                      for(var z in workSheetsFromFile[i].data){
                          if(num!==null){
                              excel.push(workSheetsFromFile[i].data[z][num]);
                          }
                          if(z==workSheetsFromFile[i].data.length-1){
                              num = null;
                          }
                      }
                  }
              }
            }
            judge();
        }
        // 进行判断
        function judge() {
            var i = 0;
            while (i < excel.length) {
                if (excel[i]) {
                    var lockSrc = excel[i].toString();
                    var lockSrc2 = lockSrc.replace(/\//g, "\\");
                    if(!fs.existsSync(path.join(__dirname,lockSrc2))){
                        console.log(lockSrc2)
                    }
                    i++;
                }else{
                    i++;
                }
            }

        }

    });
}
showAll();











