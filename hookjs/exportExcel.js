<script type="text/javascript" src="/template/json2.js"></script>
<script type="text/javascript">
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
};

function downloadLink(link) {
    var a = document.createElement('a');
    a.style.display = 'none';
    a.href = link;
    //a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

function createButtons() {
    var exportButton = document.createElement("<input type=\"button\" class=\"button\" value=\"导出Excel\" onclick=\"scoreExport();\">");
    var uploadButton = document.createElement("<input type=\"button\" class=\"button\" value=\"上传Excel\" onclick=\"scoreUpload();\">");
    var score60Button = document.createElement("<input type=\"button\" class=\"button\" value=\"一键60分\" onclick=\"setScore60();\">");
    var score100Button = document.createElement("<input type=\"button\" class=\"button\" value=\"一键100分\" onclick=\"setScore100();\">");
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
    parent.appendChild(uploadButton);
    parent.appendChild(score60Button);
    parent.appendChild(score100Button);
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
        record.push(examEl.value);
        for (var j = 0; j < 4; j++) {
            record.push(document.getElementById("cjxm" + index + "|" + (1020 + j)).value);
        }
        record.push("=AVERAGE(D" + i + ":G" + i + ")");
        record.push(document.getElementById("cj" + index + "|2").value);
        record.push("=C" + i + "*0.6+H" + i + "*0.3+I" + i + "*0.1");
        xlsData['data'].push(record);
    }

    courseName = courseName.replace("[", "(");
    courseName = courseName.replace("]", ")");
    xlsData['name'] = courseName + "成绩录入";
    return xlsData;
}

function sendRequest(reqUrl, data) {
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
    xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded"); //post需要设置Content-type，防止乱码
    //第一个参数指明访问方式，第二次参数是目标url，第三个参数是“是否异步”，true表示异步，false表示同步
    xmlhttp.send(JSON.stringify(data));
    return xmlhttp.responseText
}

function scoreExport() {
    var jsonData = getScoreJsonData();
    var resData = sendRequest("/theall/export", jsonData);
    resData = JSON.parse(resData);
    if(resData['msg'] == 'OK') {
        downloadLink(resData['url']);
    }
}

function scoreUpload() {
    var fileWrap = document.createElement("<div style=\"display:none\" </div>");
    var excelFile = document.createElement("<input type=\"file\"/>");
    fileWrap.appendChild(excelFile);
    document.body.appendChild(fileWrap);
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
    console.log("成功:" + successCount + " 失败:" + failCount);
    oxl.Quit();
    oxl = null;
    //结束excel进程，退出完成
    //window.setInterval("Cleanup();",1);
    idTmr = window.setInterval("Cleanup();", 1);
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
</script>