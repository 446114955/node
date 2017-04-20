// 根据excel数据找到对应文件并改变文件名程为拼音搜字母缩写
var transliteration = require('transliteration');
var xlsx = require('node-xlsx');
var excelPort = require('excel-export');
var fs = require('fs');
var path= require('path');
var excel = [];
var excel2 = [];
var local = ''
var num3 = 1;
var num4 = 1;
var arryName1 = [];
var arryName2 = [];
var arryName3 = [];
var lock1 = false;
var lock2 = false;
var ss =  /.*\..*/;

var oneT = 8; //第一级目录位置
var twoT = 2; //第二级目录位置
var threeT = 3; //第三级目录位置
var fourT = 5;  //目录文件名
var allSrcT ='/sound/';
var newName1 = null;
/*transliteration.transliterate('你好，世界'); // Ni Hao ,Shi Jie
 transliteration.slugify('你好，世界'); // ni-hao-shi-jie
 transliteration.slugify('你好，世界', {lowercase: true, separator: '_'});*/

const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(`${__dirname}/20170419.xlsx`));
// Parse a file
const workSheetsFromFile = xlsx.parse(`${__dirname}/20170419.xlsx`);
var arry = [];
function textslug() {
    for (var j in excel) {
        //对字进行切割
        // if (j > 0) {
            var test = transliteration.slugify(excel[j], {lowercase: true, separator: '-'})
            var test2 = test.split('-');
            //取首字母
            for (var i in test2) {
                if (i == 0) {
                    if(parseFloat(test2[i])>0&&ss.test(parseFloat(test2[i]))){
                        var str = test2[i].substr(0, 1)
                    }
                    else{
                        var str = test2[i]
                    }
                    str += '_'
                } else {
                    str += test2[i].substr(0, 1)
                }
                if (i == test2.length - 1) {
                    // console.log(str);
                    arry.push(str);
                }
            }
        // }
    }
    lock1 = true
    // console.log(arry);


}

// 取excel表workSheetsFromFile[6].data[i][4] 6代表左下角坐标，i代表列4代表行
function getExcel() {

    for (var j in workSheetsFromFile[8].data[0]) {
        if (workSheetsFromFile[8].data[0][j] == '素材展示名') {
            num = j;
        }
    }

    for (var i in workSheetsFromFile[8].data) {
        if (i != 0&&workSheetsFromFile[8].data[i][num]) {
            var test = workSheetsFromFile[8].data[i][num].toString();
            excel.push(test)
        }

    }

    // console.log(excel)
    // textslug();
}
getExcel();

// //取路径one页数，two代表二级分类three代表三级分类
// function getDirName(one,two,three) {
//     for (var i in workSheetsFromFile[one].data) {
//         if (i > 0) {
//             if (i > 1) {
//                 if (workSheetsFromFile[one].data[i][two] != workSheetsFromFile[one].data[i - 1][two]) {
//                     num3++;
//                     num4 = 0;
//                 }
//                 if (workSheetsFromFile[one].data[i][three] != workSheetsFromFile[one].data[i - 1][three]) {
//                     num4++;
//                 }
//             }
//             dirslug(workSheetsFromFile[one].data[i][two]);
//             // if (workSheetsFromFile[8].data[i][3]) {
//             dirslug2(workSheetsFromFile[one].data[i][three]);
//             // }
//
//         }
//     }
// }
//
// function dirslug(str) {
//
//     // for (var j in str) {
//     //对字进行切割
//     var test = transliteration.slugify(str, {lowercase: true, separator: '-'})
//     var test2 = test.split('-');
//     //取首字母
//     for (var i in test2) {
//         if (i == 0) {
//             var str2 = num3 + '_'
//             str2 += test2[i].substr(0, 1)
//
//         } else {
//             str2 += test2[i].substr(0, 1)
//         }
//         if (i == test2.length - 1) {
//             str2 += '/';
//             // console.log(str2);
//             arryName1.push(str2);
//         }
//
//     }
//     // console.log(arryName1);
//     lock2 = true;
//     if (lock1 == true && lock2 == true) {
//         addLocal('/sound/');
//     }
// }
//
// function dirslug2(str) {
//
//     // for (var j in str) {
//     //对字进行切割
//     if(str){
//         var test = transliteration.slugify(str, {lowercase: true, separator: '-'})
//         var test2 = test.split('-');
//         //取首字母
//         for (var i in test2) {
//             if (i == 0) {
//                 var str2 = num4 + '_'
//                 str2 += test2[i].substr(0, 1)
//
//             } else {
//                 str2 += test2[i].substr(0, 1)
//             }
//             if (i == test2.length - 1) {
//                 str2 += '/';
//                 // console.log(str2);
//                 arryName2.push(str2);
//             }
//
//         }
//     }else{
//         arryName2.push('');
//     }
//
//     // console.log(arryName2);
// }
// //allSrc为总目录
// function addLocal(allSrc) {
//     for (var i in arry) {
//         if (arryName2[i]) {
//             arryName3[i] =allSrc+arryName1[i];
//             arryName3[i] += arryName2[i];
//             arryName3[i] += arry[i];
//         } else {
//             arryName3[i] =allSrc+arryName1[i];
//             arryName3[i] += arry[i];
//         }
//     }
//     // console.log(arryName3);
//     pushExcel();
// }
//
// getDirName(oneT,twoT,threeT);
//
// /*导出文件测试*/
// function pushExcel() {
//     var arry = arryName3
//     var arry2 = []
//     for (var i in arryName3) {
//         arry2.push([arryName3[i]])
//     }
//     var data = arry2;
//     var buffer = xlsx.build([{name: "mySheetName", data: data}]); // Returns a buffer
//     var conf = {};
//     // var filename = 'filename';  //只支持字母和数字命名
//     conf.cols = [
//         {caption: '路径', type: 'string', width: 80},
//     ];
//
//     conf.rows = [];
//     for (var i in arry2) {
//         conf.rows.push(arry2[i]);
//     }
//
//
//     var result = excelPort.execute(conf);
//     var random = Math.floor(Math.random() * 10000 + 0);
//     // var filePath = mulu3 + random + ".xlsx";
//     fs.writeFile("mySheetName.xlsx", result, 'binary', function (err) {
//         if (err) {
//             console.log(err);
//         } else {
//             console.log('导出成功')
//         }
//     });
//     // }
// }

// 根据excel进目录
function intoMulu() {
    // 遍历excel数据
    for (var i in workSheetsFromFile[oneT].data){
        if(i>0){
            if(i>1) {
                if (workSheetsFromFile[oneT].data[i][2] != workSheetsFromFile[oneT].data[i - 1][2]) {
                    num3++;
                    num4 = 0;
                }
                if (workSheetsFromFile[oneT].data[i][3] != workSheetsFromFile[oneT].data[i - 1][3]) {
                    num4++;
                }
            }
                if (workSheetsFromFile[oneT].data[i][3]) {
                    var filepath = process.cwd() + '\\' + workSheetsFromFile[oneT].data[i][1] + '\\' + num3+ '.' + workSheetsFromFile[oneT].data[i][twoT] + '\\' + num4+ '.' + workSheetsFromFile[oneT].data[i][threeT] + '\\' + workSheetsFromFile[oneT].data[i][fourT] + '.mp3';
                } else {
                    var filepath = process.cwd() + '\\' + workSheetsFromFile[oneT].data[i][1] + '\\' + num3+ '.' + workSheetsFromFile[oneT].data[i][twoT] + '\\' + workSheetsFromFile[oneT].data[i][fourT] + '.mp3';
                }

            console.log(filepath)
            var test = transliteration.slugify(workSheetsFromFile[oneT].data[i][5], {lowercase: true, separator: '-'})
            var test2 = test.split('-');
            //取首字母
            for (var i in test2) {
                if (i == 0) {
                    if(parseFloat(test2[i])>0&&ss.test(parseFloat(test2[i]))){
                        var str = test2[i].substr(0, 1)
                    }
                    else{
                        var str = test2[i]
                    }
                    str += '_'
                } else {
                    str += test2[i].substr(0, 1)
                }
                if (i == test2.length - 1) {
                    // console.log(str);
                    newName1 = str;
                }
            }
            try{
                var fileNow = fs.statSync(filepath);
                // 判断是否是文件
                if(fileNow.isFile()){
                    console.log("isFile,chaning filename...");
                    var filename = path.basename(filepath);
                    var parentDir = path.dirname(filepath);
                    var parentDirname = path.basename(path.dirname(filepath));
                    var thisFilename = path.basename(__filename);
                    //console.log(thisFilename);
                    var newName = newName1 + '.mp3'
                    newPath.replace(/\//g, "\\");
                    filepath.replace(/\//g, "\\");
                    // 文件改名
                    fs.rename(filepath, newPath, function (err) {
                        if (err) {
                            console.log(err)
                        }
                    });
                    console.log("going to rename from " + filepath + " to " + newPath);
                } else if (fileNow.isDirectory()) {
                    console.log("============[" + filepath + "] isDir===========");
                    // renameFilesInDir(filepath);
                } else {
                    console.log("unknow type of file");
                }
            }catch(err) {
                // console.log(err);
                // return;
            }

        }
    }
}
intoMulu()
