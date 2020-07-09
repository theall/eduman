<script type="text/javascript" src="/template/json2.js"></script>
<script type="text/javascript" src="/template/blob.js"></script>
<script type="text/javascript">
String.prototype.trim = function()    
{    
    return this.replace(/(^\s*)|(\s*$)/g, "");    
}

function downloadContent(content, fileName) {
    // �������صĿ���������
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
            console.log("��" + (i + 1) + "����¼��һ��,ѧ��:" + realNo + "����:" + realName);
            failCount = failCount + 1;
            continue;
        }

        examEl.value = scoreList[i][2];// ���Գɼ�
        for (var j = 0; j < 4; j++) {
            document.getElementById("cjxm" + i + "|" + (1020 + j)).value = scoreList[i][3+j];//ƽʱ�ɼ�
        }
        document.getElementById("cj" + i + "|1").value = scoreList[i][7];//ƽʱ�ɼ��ܷ�
        document.getElementById("cj" + i + "|2").value = scoreList[i][8];//���ڳɼ�
        document.getElementById("zcj" + i).value = scoreList[i][9];//���ճɼ�
        successCount = successCount + 1;
    }
    console.log("�ɹ�:" + successCount + " ʧ��:" + failCount); 
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
    setButton.innerText = "д��";
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
    var exportButton = createButton("����Excel", scoreExport);
    var uploadButton = createButton("�ϴ�Excel", scoreUpload);
    var score60Button = createButton("һ��60��", setScore60);
    var score100Button = createButton("һ��100��", setScore100);
    var scoreRandomButton = createButton("�������", setScoreRandom);
    var selectAllButton = createButton("���м���", selectAll);
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
 * ͨ�õĴ����ضԻ��򷽷���û�в��Թ����������
 * @param url ���ص�ַ��Ҳ������һ��blob���󣬱�ѡ
 * @param saveName �����ļ�������ѡ
 */
function openDownloadDialog(url, saveName) {
    if (typeof url == 'object' && url instanceof Blob) {
        url = URL.createObjectURL(url); // ����blob��ַ
    }
    var aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || ''; // HTML5���������ԣ�ָ�������ļ��������Բ�Ҫ��׺��ע�⣬file:///ģʽ�²�����Ч
    var event;
    if (window.MouseEvent) event = new MouseEvent('click');
    else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
    }
    aLink.dispatchEvent(event);
}

// ��һ��sheetת�����յ�excel�ļ���blob����Ȼ������URL.createObjectURL����
function sheet2blob(sheet, sheetName) {
    sheetName = sheetName || 'sheet1';
    var workbook = {
        SheetNames: [sheetName],
        Sheets: {}
    };
    workbook.Sheets[sheetName] = sheet;
    // ����excel��������
    var wopts = {
        bookType: 'xlsx', // Ҫ���ɵ��ļ�����
        bookSST: false, // �Ƿ�����Shared String Table���ٷ������ǣ�������������ٶȻ��½������ڵͰ汾IOS�豸���и��õļ�����
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);
    var blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    // �ַ���תArrayBuffer
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
        window.alert("�γ���Ϊ��!");
        return;
    }
    var xlsData = {}
    var titleHeads = new Array("ѧ��", "����", "���Գɼ�", "ƽʱ�ɼ�1", "ƽʱ�ɼ�2", "ƽʱ�ɼ�3", "ƽʱ�ɼ�4", "ƽʱ�ɼ��ܷ�", "����", "���ճɼ�");
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
            // �ȼ�
            var totalScore = getTotalScoreLevel(document.getElementById("zcj" + index).value);
            record.push(totalScore);
        } else {
            record.push("=C" + i + "*0.6+H" + i + "*0.3+I" + i + "*0.1");
        }
        xlsData['data'].push(record);
    }

    courseName = courseName.replace("[", "(");
    courseName = courseName.replace("]", ")");
    xlsData['name'] = courseName + "�ɼ�¼��";
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
        alert("�����������֧��AJAX��");
        return;
    }
    xmlhttp.open("POST", reqUrl, false); 
    xmlhttp.setRequestHeader("Content-type", "multipart/form-data"); //post��Ҫ����Content-type����ֹ����
    //��һ������ָ�����ʷ�ʽ���ڶ��β�����Ŀ��url�������������ǡ��Ƿ��첽����true��ʾ�첽��false��ʾͬ��
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
        alert("�����������֧��AJAX��");
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
    console.log("������ " + (lineCount - 1))
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
            console.log("��" + (index + 1) + "����¼��һ��,ѧ��:" + realNo + "����:" + realName);
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
    console.log("�ɹ�:" + successCount + " ʧ��:" + failCount);    
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
    //��Excel���ȡ���ݵ�ҳ��
    var path = excelFile.value;
    if(path === '')
        return;
    
    owb = oxl.workbooks.open(path);
    owb.worksheets(1).select();
    var osheet = owb.activesheet;
    var lineCount = osheet.UsedRange.Cells.Rows.Count;
    console.log("������ " + (lineCount - 1))
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
            console.log("��" + (index + 1) + "����¼��һ��,ѧ��:" + realNo + "����:" + realName);
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
    console.log("�ɹ�:" + successCount + " ʧ��:" + failCount);
}

function scoreVerify() {
    var excelFile = document.createElement("<input type=\"file\"/>");
    document.body.appendChild(excelFile);
    excelFile.click();

    var oxl = new ActiveXObject("Excel.application");
    var owb;
    //��Excel���ȡ���ݵ�ҳ��
    var path = excelFile.value;

    owb = oxl.workbooks.open(path);
    owb.worksheets(1).select();
    var osheet = owb.activesheet;
    var lineCount = osheet.UsedRange.Cells.Rows.Count;
    console.log("������ " + (lineCount - 1))
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
            console.log("ѧ��" + osheet.cells(i, 1).value + " ����" + osheet.cells(i, 2).value + " �ɼ�����!ʵ�ʷ���Ϊ" + realScore);
            failCount = failCount + 1;
        }
    }
    oxl.Quit();
    oxl = null;

    //����excel���̣��˳����
    //window.setInterval("Cleanup();",1);
    idTmr = window.setInterval("Cleanup();", 1);
    // ����������ڽ��IE call Excel��һ��BUG, MSDN���ṩ�ķ���:
    //   setTimeout(CollectGarbage, 1);
    // ���ڲ������(��ͬ��)��ҳ��������״̬, ���Խ�����SaveAs()�ȷ�����
    // �´ε���ʱ��Ч.
    if (failCount == 0) {
        window.alert("У��ɹ�!");
    } else {
        window.alert("У��ʧ��!");
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