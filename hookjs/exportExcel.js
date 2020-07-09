<script type="text/javascript" src="/template/json2.js"></script>
<script type="text/javascript" src="/template/blob.js"></script>
<script type="text/javascript">
String.prototype.trim = function()    
{    
    return this.replace(/(^\s*)|(\s*$)/g, "");    
}

function downloadContent(content, fileName) {
    // 创建隐藏的可下载链接
    var data = new Blob([content]);
    // for IE
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {
        window.navigator.msSaveOrOpenBlob(data, fileName);
    }
    // for Non-IE (chrome, firefox etc.)
    else {
        var a = document.createElement('a');
        document.body.appendChild(a);
        a.style = 'display: none';
        var url = window.URL.createObjectURL(data);
        a.href = url;
        a.download = fileName;
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    }
}

function getLevelStr(score) {
    if(score>=90)
        return 5;
    else if(score >= 80)
        return 4;
    else if(score >= 70)
        return 3;
    else if(score >= 60)
        return 2;
    return 1;
}

function setLevelScore(scoreList) {
    for(var i=0;i<scoreList.length;i++) {
        var lastScore = scoreList[i];
        if(typeof(lastScore)==="object")
            lastScore = lastScore[lastScore.length-1];
        var level = getLevelStr(lastScore);
        var el = document.getElementById("zcj"+i);
        if(el != undefined)
            el.selectedIndex = level;
    }
}

function setScore(scoreList) {
    var successCount = 0;
    var failCount = 0;
    for (var i = 0; i <= scoreList.length; i++) {
        var examEl = document.getElementById("cj" + i + "|0");
        if (!examEl) {
            break;
        }

        var displayNo = document.getElementById("tr" + i).children[2].lastChild.data;
        var displayName = document.getElementById("tr" + i).children[3].firstChild.data;
        var realNo = scoreList[i][0];
        var realName = scoreList[i][1];
        if (realNo != displayNo || realName != displayName) {
            console.log("第" + (i + 1) + "条记录不一致,学号:" + realNo + "姓名:" + realName);
            failCount = failCount + 1;
            continue;
        }

        examEl.value = scoreList[i][2];// 考试成绩
        for (var j = 0; j < 4; j++) {
            document.getElementById("cjxm" + i + "|" + (1020 + j)).value = scoreList[i][3+j];//平时成绩
        }
        document.getElementById("cj" + i + "|1").value = scoreList[i][7];//平时成绩总分
        document.getElementById("cj" + i + "|2").value = scoreList[i][8];//考勤成绩
        document.getElementById("zcj" + i).value = scoreList[i][9];//最终成绩
        successCount = successCount + 1;
    }
    console.log("成功:" + successCount + " 失败:" + failCount); 
}

function getScore(text) {
    var score = parseFloat(text);
    return Math.round(score);
}

function textAreaChanged() {
    var text = document.getElementById("text_area").value;
    var numberList = text.split("\n");
    var scoreList = [];
    for(var i=1;i<numberList.length;i++) {
        var lineList = numberList[i].trim().split("\t");
        for(var j=0;j<lineList.length;j++) {
            if(j<2)
                lineList[j] = lineList[j].trim();
            else {
                var score = getScore(lineList[j]);
                lineList[j] = score;
            }
            lineList[j] = Math.round(lineList[j]);
        }
        scoreList.push(lineList);
    }
    var scoreRadio = document.getElementById("rbfsfs_0");
    if(scoreRadio.checked) {
        setScore(scoreList);
    } else {
        setLevelScore(scoreList);
    }
}

function createTextArea() {
    var textArea = document.createElement("textarea");
    textArea.id = "text_area";
    //textArea.style = "width:800px;height:100px;";
    document.body.appendChild(textArea);
    var setButton = document.createElement("button");
    setButton.onclick = textAreaChanged;
    setButton.innerText = "写入";
    document.body.appendChild(setButton);
}

function selectAll() {
    for(var i=0;i<100;i++){
        var s = document.getElementById('zcj'+i)
        if(s == null)
            break;
        s.children[2].selected = true;
    }
}

function downloadLink(link) {
    link = decodeURI(link);
    link = link.replace(/%2F/g, '/');
    var a = document.createElement('a');
    a.style.display = 'none';
    a.href = link;
    //a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

function createButton(value, clickFunc) {
    var btn = document.createElement("input");
    btn.className = "button";
    btn.value = value;
    btn.onclick = clickFunc;
    return btn;
}

function createButtons() {
    var exportButton = createButton("导出Excel", scoreExport);
    var uploadButton = createButton("上传Excel", scoreUpload);
    var score60Button = createButton("一键60分", setScore60);
    var score100Button = createButton("一键100分", setScore100);
    var scoreRandomButton = createButton("随机给分", setScoreRandom);
    var selectAllButton = createButton("所有及格", selectAll);
    var courseEl = document.getElementById("ddlkc");
    var examElement = document.getElementById("ddlksxz");
    var examType = courseEl.options[examElement.selectedIndex].text;
    var courseName = courseEl.options[courseEl.selectedIndex].text;
    if (courseName.length < 3) {
        //exportButton.disabled = "disabled";
        //uploadButton.disabled = "disabled";
        //score60Button.disabled = "disabled";
        //score100Button.disabled = "disabled";
    }

    var parent = document.getElementById("dafxd").parentNode;
    parent.appendChild(exportButton);
    //parent.appendChild(uploadButton);
    parent.appendChild(score60Button);
    parent.appendChild(score100Button);
    parent.appendChild(scoreRandomButton);
    parent.appendChild(selectAllButton);
}

function setScoreRandom() {
    var tbl = document.getElementById("mxh");
    for (var i = 2; i <= tbl.firstChild.childNodes.length; i++) {
        var index = i - 2;
        var examEl = document.getElementById("cj" + index + "|0");
        if (!examEl) {
            break;
        }

        //examEl.value = 60;
        for (var j = 0; j < 4; j++) {
            document.getElementById("cjxm" + index + "|" + (1020 + j)).value = Math.floor(Math.random()*40+60);
        }
        document.getElementById("cj" + index + "|1").value = 100;
        document.getElementById("cj" + index + "|2").value = Math.floor(Math.random()*40+60);
        document.getElementById("zcj" + index).value = Math.floor(Math.random()*40+60);
    }
}

function readWorkbookFromLocalFile(file, callback) {
    var reader = new FileReader();
    reader.onload = function(e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, { type: 'binary' });
        if (callback) callback(workbook);
    };
    reader.readAsBinaryString(file);
}

/**
 * 通用的打开下载对话框方法，没有测试过具体兼容性
 * @param url 下载地址，也可以是一个blob对象，必选
 * @param saveName 保存文件名，可选
 */
function openDownloadDialog(url, saveName) {
    if (typeof url == 'object' && url instanceof Blob) {
        url = URL.createObjectURL(url); // 创建blob地址
    }
    var aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
    var event;
    if (window.MouseEvent) event = new MouseEvent('click');
    else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
    }
    aLink.dispatchEvent(event);
}

// 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
function sheet2blob(sheet, sheetName) {
    sheetName = sheetName || 'sheet1';
    var workbook = {
        SheetNames: [sheetName],
        Sheets: {}
    };
    workbook.Sheets[sheetName] = sheet;
    // 生成excel的配置项
    var wopts = {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);
    var blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    // 字符串转ArrayBuffer
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    return blob;
}

function getFloat(value) {
    if(value == '')
        return 0.0;
    return parseFloat(value);
}

function getTotalScoreLevel(str) {
    if(str == '--')
        return '';
    var index = str.indexOf('[');
    if(index > 0)
        str = str.slice(0, index);
    return str;
}

function getScoreJsonData() {
    var courseEl = document.getElementById("ddlkc");
    var courseName = courseEl.options[courseEl.selectedIndex].text;
    if (courseName == "") {
        window.alert("课程名为空!");
        return;
    }
    var xlsData = {}
    var titleHeads = new Array("学号", "姓名", "考试成绩", "平时成绩1", "平时成绩2", "平时成绩3", "平时成绩4", "平时成绩总分", "考勤", "最终成绩");
    xlsData['heads'] = titleHeads;
    xlsData['data'] = []
    var tbl = document.getElementById("mxh");
    for (var i = 2; i <= tbl.firstChild.childNodes.length; i++) {
        var index = i - 2;
        var displayNo = document.getElementById("tr" + index).children[2].lastChild.data;
        var displayName = document.getElementById("tr" + index).children[3].firstChild.data;
        if (displayName == "")
            break;

        var record = [];
        var examEl = document.getElementById("cj" + index + "|0");
        record.push(displayNo);
        record.push(displayName);
        record.push(getFloat(examEl.value));
        for (var j = 0; j < 4; j++) {
            record.push(getFloat(document.getElementById("cjxm" + index + "|" + (1020 + j)).value));
        }
        record.push("=AVERAGE(D" + i + ":G" + i + ")");
        record.push(getFloat(document.getElementById("cj" + index + "|2").value));
        
        
        if(document.getElementById('rbfsfs_1').checked) {
            // 等级
            var totalScore = getTotalScoreLevel(document.getElementById("zcj" + index).value);
            record.push(totalScore);
        } else {
            record.push("=C" + i + "*0.6+H" + i + "*0.3+I" + i + "*0.1");
        }
        xlsData['data'].push(record);
    }

    courseName = courseName.replace("[", "(");
    courseName = courseName.replace("]", ")");
    xlsData['name'] = courseName + "成绩录入";
    return xlsData;
}

function postRequest(reqUrl, data) {
    var xmlhttp = null;
    if (window.XMLHttpRequest) {
        xmlhttp = new XMLHttpRequest();
    } else if (window.ActiveXObject) {
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    if (xmlhttp == null) {
        alert("您的浏览器不支持AJAX！");
        return;
    }
    xmlhttp.open("POST", reqUrl, false); 
    xmlhttp.setRequestHeader("Content-type", "multipart/form-data"); //post需要设置Content-type，防止乱码
    //第一个参数指明访问方式，第二次参数是目标url，第三个参数是“是否异步”，true表示异步，false表示同步
    xmlhttp.send(data);
    return xmlhttp.responseText
}

function getRequest(reqUrl) {
    var xmlhttp = null;
    if (window.XMLHttpRequest) {
        xmlhttp = new XMLHttpRequest();
    } else if (window.ActiveXObject) {
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    if (xmlhttp == null) {
        alert("您的浏览器不支持AJAX！");
        return;
    }
    xmlhttp.open("GET", reqUrl, false); 
    return xmlhttp.responseText
}

function scoreExport() {
    var jsonData = getScoreJsonData();
    jsonData = JSON.stringify(jsonData);
    var resData = postRequest("/theall/export", jsonData);
    if(resData == '') {
        alert("No response!");
        return;
    }
    resData = JSON.parse(resData);
    if(resData['msg'] == 'OK') {
        downloadLink(resData['url']);
    }
}

function createGuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        var r = Math.random()*16|0, v = c == 'x' ? r : (r&0x3|0x8);
        return v.toString(16);
    });
}

function scoreUpload() {
    var fileWrap = document.createElement("<div style=\"display:block\"></div>");
    var form = document.createElement("<form action=\"/theall/upload\" method=\"post\" enctype=\"multipart/form-data\"></form>");
    var excelFile = document.createElement("<input type=\"file\" name=\"file\"/>");
    var guidEl = document.createElement("<input type=\"hidden\" name=\"guid\">");
    var guid = createGuid();
    guidEl.value = guid;
    form.appendChild(excelFile);
    form.appendChild(guidEl);
    fileWrap.appendChild(form);
    document.body.appendChild(fileWrap);
    excelFile.click();
    if(excelFile.value === '') {
        window.alert("excel file is empty");
        return;
    }
    
    try {
        form.submit()
    } catch(err) {
        window.alert(err.message);
    }
    document.body.removeChild(fileWrap);
}

function updateData(data) {
    var lineCount = osheet.UsedRange.Cells.Rows.Count;
    console.log("总人数 " + (lineCount - 1))
    var successCount = 0;
    var failCount = 0;
    for (var i = 2; i <= lineCount; i++) {
        var index = i - 2;
        var examEl = document.getElementById("cj" + index + "|0");
        if (!examEl) {
            break;
        }

        var displayNo = document.getElementById("tr" + index).children[2].lastChild.data;
        var displayName = document.getElementById("tr" + index).children[3].firstChild.data;
        var realNo = osheet.cells(i, 1).value;
        var realName = osheet.cells(i, 2).value;
        if (realNo != displayNo || realName != displayName) {
            console.log("第" + (index + 1) + "条记录不一致,学号:" + realNo + "姓名:" + realName);
            failCount = failCount + 1;
            continue;
        }

        examEl.value = osheet.cells(i, 3).value;
        for (var j = 0; j < 4; j++) {
            document.getElementById("cjxm" + index + "|" + (1020 + j)).value = osheet.cells(i, 4 + j).value;
        }
        document.getElementById("cj" + index + "|1").value = osheet.cells(i, 8).value;
        document.getElementById("cj" + index + "|2").value = osheet.cells(i, 9).value;
        document.getElementById("zcj" + index).value = osheet.cells(i, 10).value.toFixed(1);
        successCount = successCount + 1;
    }
    document.body.removeChild(fileWrap);
    console.log("成功:" + successCount + " 失败:" + failCount);    
}

function getValidValue(value) {
    if(typeof(value) == "undefined")
        value = '';
    return value;
}

function scoreUploadLegacy() {
    var fileWrap = document.createElement("<div style=\"display:none\" </div>");
    var excelFile = document.createElement("<input type=\"file\"/>");
    fileWrap.appendChild(excelFile);
    document.body.appendChild(fileWrap);
    excelFile.click();

    var oxl = new ActiveXObject("Excel.application");
    var owb;
    //从Excel里读取数据到页面
    var path = excelFile.value;
    if(path === '')
        return;
    
    owb = oxl.workbooks.open(path);
    owb.worksheets(1).select();
    var osheet = owb.activesheet;
    var lineCount = osheet.UsedRange.Cells.Rows.Count;
    console.log("总人数 " + (lineCount - 1))
    var successCount = 0;
    var failCount = 0;
    for (var i = 2; i <= lineCount; i++) {
        var index = i - 2;
        var examEl = document.getElementById("cj" + index + "|0");
        if (!examEl) {
            break;
        }

        var displayNo = document.getElementById("tr" + index).children[2].lastChild.data;
        var displayName = document.getElementById("tr" + index).children[3].firstChild.data;
        var realNo = osheet.cells(i, 1).value;
        var realName = osheet.cells(i, 2).value;
        if (realNo != displayNo || realName != displayName) {
            console.log("第" + (index + 1) + "条记录不一致,学号:" + realNo + "姓名:" + realName);
            failCount = failCount + 1;
            continue;
        }

        examEl.value = getValidValue(osheet.cells(i, 3).value);
        for (var j = 0; j < 4; j++) {
            document.getElementById("cjxm" + index + "|" + (1020 + j)).value = getValidValue(osheet.cells(i, 4 + j).value);
        }
        document.getElementById("cj" + index + "|1").value = getValidValue(osheet.cells(i, 8).value);
        document.getElementById("cj" + index + "|2").value = getValidValue(osheet.cells(i, 9).value);
        
        var totalScore = getValidValue(osheet.cells(i, 10).value);
        if(typeof(totalScore) == 'number') {
            totalScore = totalScore.toFixed(1);
        }
        document.getElementById("zcj" + index).value = totalScore;
        successCount = successCount + 1;
    }
    document.body.removeChild(fileWrap);
    console.log("成功:" + successCount + " 失败:" + failCount);
}

function scoreVerify() {
    var excelFile = document.createElement("<input type=\"file\"/>");
    document.body.appendChild(excelFile);
    excelFile.click();

    var oxl = new ActiveXObject("Excel.application");
    var owb;
    //从Excel里读取数据到页面
    var path = excelFile.value;

    owb = oxl.workbooks.open(path);
    owb.worksheets(1).select();
    var osheet = owb.activesheet;
    var lineCount = osheet.UsedRange.Cells.Rows.Count;
    console.log("总人数 " + (lineCount - 1))
    var failCount = 0;
    for (var i = 2; i <= lineCount; i++) {
        var index = i - 2;
        var displayScore = document.getElementById("zcj" + index).value;
        if (!displayScore)
            break;
        var displayNo = document.getElementById("tr" + index).children[2].lastChild.data;
        var displayName = document.getElementById("tr" + index).children[3].firstChild.data;

        var realNo = osheet.cells(i, 1).value;
        var realName = osheet.cells(i, 2).value;
        var realScore = osheet.cells(i, 10).value;
        realScore = realScore.toFixed(1);
        if (realNo != displayNo || realName != displayName || displayScore != realScore) {
            console.log("学号" + osheet.cells(i, 1).value + " 姓名" + osheet.cells(i, 2).value + " 成绩不对!实际分数为" + realScore);
            failCount = failCount + 1;
        }
    }
    oxl.Quit();
    oxl = null;

    //结束excel进程，退出完成
    //window.setInterval("Cleanup();",1);
    idTmr = window.setInterval("Cleanup();", 1);
    // 下面代码用于解决IE call Excel的一个BUG, MSDN中提供的方法:
    //   setTimeout(CollectGarbage, 1);
    // 由于不能清除(或同步)网页的受信任状态, 所以将导致SaveAs()等方法在
    // 下次调用时无效.
    if (failCount == 0) {
        window.alert("校验成功!");
    } else {
        window.alert("校验失败!");
    }
}

function setScore60() {
    var tbl = document.getElementById("mxh");
    for (var i = 2; i <= tbl.firstChild.childNodes.length; i++) {
        var index = i - 2;
        var examEl = document.getElementById("cj" + index + "|0");
        if (!examEl) {
            break;
        }

        examEl.value = 60;
        for (var j = 0; j < 4; j++) {
            document.getElementById("cjxm" + index + "|" + (1020 + j)).value = 0;
        }
        document.getElementById("cj" + index + "|1").value = 0;
        document.getElementById("cj" + index + "|2").value = 0;
        document.getElementById("zcj" + index).value = 60;
    }
}

function setScore100() {
    var tbl = document.getElementById("mxh");
    for (var i = 2; i <= tbl.firstChild.childNodes.length; i++) {
        var index = i - 2;
        var examEl = document.getElementById("cj" + index + "|0");
        if (!examEl) {
            break;
        }

        //examEl.value = 60;
        for (var j = 0; j < 4; j++) {
            document.getElementById("cjxm" + index + "|" + (1020 + j)).value = 100;
        }
        document.getElementById("cj" + index + "|1").value = 100;
        document.getElementById("cj" + index + "|2").value = 100;
        //document.getElementById("zcj" + index).value = 60;
    }
}

function Cleanup() {
    window.clearInterval(idTmr);
    CollectGarbage();
}
createButtons();
createTextArea();
</script>