<script type="text/javascript" src="/template/json2.js"></script>
<script type="text/javascript">
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
    var exportButton = document.createElement("<input type=\"button\" class=\"button\" value=\"����Excel\" onclick=\"scoreExport();\">");
    var uploadButton = document.createElement("<input type=\"button\" class=\"button\" value=\"�ϴ�Excel\" onclick=\"scoreUpload();\">");
    var score60Button = document.createElement("<input type=\"button\" class=\"button\" value=\"һ��60��\" onclick=\"setScore60();\">");
    var score100Button = document.createElement("<input type=\"button\" class=\"button\" value=\"һ��100��\" onclick=\"setScore100();\">");
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
    xlsData['name'] = courseName + "�ɼ�¼��";
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
        alert("�����������֧��AJAX��");
        return;
    }
    xmlhttp.open("POST", reqUrl, false); 
    xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded"); //post��Ҫ����Content-type����ֹ����
    //��һ������ָ�����ʷ�ʽ���ڶ��β�����Ŀ��url�������������ǡ��Ƿ��첽����true��ʾ�첽��false��ʾͬ��
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
    //��Excel���ȡ���ݵ�ҳ��
    var path = excelFile.value;

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

        examEl.value = osheet.cells(i, 3).value;
        for (var j = 0; j < 4; j++) {
            document.getElementById("cjxm" + index + "|" + (1020 + j)).value = osheet.cells(i, 4 + j).value;
        }
        document.getElementById("cj" + index + "|1").value = osheet.cells(i, 8).value;
        document.getElementById("cj" + index + "|2").value = osheet.cells(i, 9).value;
        document.getElementById("zcj" + index).value = osheet.cells(i, 10).value.toFixed(1);
        successCount = successCount + 1;
    }
    console.log("�ɹ�:" + successCount + " ʧ��:" + failCount);
    oxl.Quit();
    oxl = null;
    //����excel���̣��˳����
    //window.setInterval("Cleanup();",1);
    idTmr = window.setInterval("Cleanup();", 1);
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
</script>