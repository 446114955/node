// 操作excel表格，在对应的路径下添加一个文件名
var transliteration = require('transliteration');
var xlsx = require('node-xlsx');
var excelPort = require('excel-export');
var fs = require('fs');
var excel = [];
var local = ''
var num3 = 1;
var num4 = 1;
var arryName1 = [];
var arryName2 = [];
var arryName3 = [];
var lock1 = false;
var lock2 = false;
var ss =  /.*\..*/;
/*transliteration.transliterate('你好，世界'); // Ni Hao ,Shi Jie
 transliteration.slugify('你好，世界'); // ni-hao-shi-jie
 transliteration.slugify('你好，世界', {lowercase: true, separator: '_'});*/

const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(`${__dirname}/20170417.xlsx`));
// Parse a file
const workSheetsFromFile = xlsx.parse(`${__dirname}/20170417.xlsx`);
var arry = [];
function textslug() {
    for (var j in excel) {
        //对字进行切割
        if (j > 0) {
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
        }
    }
    lock1 = true
}

// 取excel表workSheetsFromFile[6].data[i][4] 6代表左下角坐标，i代表列4代表行
function getExcel() {
    for(var j in workSheetsFromFile[4].data[0]){
        if(workSheetsFromFile[4].data[0][j]=='素材文件名') {
            num = j;
        }
    }

    for (var i in workSheetsFromFile[4].data) {
        // if(workSheetsFromFile[0].data[i][2]=='打击乐'){
        // excel.push(workSheetsFromFile[8].data[i][num])
        // }

        if(i!=0&&workSheetsFromFile[4].data[i][num]){
            // excel.push([workSheetsFromFile[8].data[i][num].split('/')])
            var ss = workSheetsFromFile[4].data[i][num].split('/');
            var sss = ''
            var name = ss[(ss.length-1)].split('.')
            for (var j in ss){
                if(j!=0){
                    sss += '/'+ss[j];
                }
                if(j==ss.length-2){
                    sss+='/'+name[0]
                }

            }
            var xx = sss.toString();
            excel.push(xx)
        }

    }
    // console.log(excel)
    pushExcel();
    // textslug();
}
getExcel();



/*导出文件测试*/
function pushExcel() {
    var arry2 = [];
    for (var i in excel) {
        arry2.push([excel[i]])
    }
    var data = arry2;

    var buffer = xlsx.build([{name: "mySheetName", data: data}]); // Returns a buffer
    var conf = {};
    // var filename = 'filename';  //只支持字母和数字命名
    // 表格的行
    conf.cols = [
        {caption: '路径',type: 'string', width: 500},
    ];
    // 表格的列
    conf.rows = [];
    for (var i in arry2) {
        conf.rows.push(arry2[i]);
    }

    var result = excelPort.execute(conf);
    var random = Math.floor(Math.random() * 10000 + 0);
    // 导出excel
    fs.writeFile("mySheetName.xlsx", result, 'binary', function (err) {
        if (err) {
            console.log(err);
        } else {

            console.log('导出成功')
        }
    });

    // }


}
