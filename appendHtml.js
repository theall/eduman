<script type="text/javascript">
function createButtons() {
    var exportButton = document.createElement("<input type=\"button\" value=\"导出Excel\" onclick=\"scoreExport();\">");
    var uploadButton = document.createElement("<input type=\"button\" value=\"上传Excel\" onclick=\"scoreUpload();\">");
    var courseEl = document.getElementById("ddlkc");
    var courseName = courseEl.options[courseEl.selectedIndex].text;
    if(courseName.length<3) {
        exportButton.disabled = "disabled";
        uploadButton.disabled = "disabled";
    }
    var parent = document.getElementById("dafxd").parentNode;
    parent.appendChild(exportButton);
    parent.appendChild(uploadButton);
}
function scoreExport() {
    var courseEl = document.getElementById("ddlkc");
    var courseName = courseEl.options[courseEl.selectedIndex].text;
    if(courseName=="")
    {
        window.alert("课程名为空!");
        return;
    }
    var oxl = new ActiveXObject("Excel.application"); 
    var owb;
    owb = oxl.workbooks.Add();
    owb.worksheets(1).select();
    var osheet = owb.activesheet;
    var titleHeads=new Array("学号","姓名","考试成绩","平时成绩1","平时成绩2","平时成绩3","平时成绩4","平时成绩总分","考勤","最终成绩");
    for(var i=0;i<titleHeads.length;i++) {
        osheet.cells(1, i+1).value = titleHeads[i];
    }
    var tbl = document.getElementById("mxh");
    for(var i=2;i<=tbl.firstChild.childNodes.length;i++) {
        var index = i - 2;
        var displayNo = document.getElementById("tr"+index).children[2].lastChild.data;
        var displayName = document.getElementById("tr"+index).children[3].firstChild.data;
        if(displayName=="")
            break;
        
        var examEl = document.getElementById("cj"+index+"|0");
        osheet.cells(i, 1).value = displayNo;
        osheet.cells(i, 2).value = displayName;
        osheet.cells(i, 3).value = examEl.value;
        for(var j=0;j<4;j++) {
            osheet.cells(i, 4+j).value = document.getElementById("cjxm"+index+"|"+(1020+j)).value;
        }
        osheet.cells(i, 8).value = "=AVERAGE(D"+i+":G"+i+")";
        osheet.cells(i, 9).value = document.getElementById("cj"+index+"|2").value;
        osheet.cells(i, 10).value = "=C"+i+"*0.6+H"+i+"*0.3+I"+i+"*0.1";
    }
    // 设置格式
    for(var j=1;j<=10;j++) {
        osheet.cells(1, j).Font.Name = "黑体"; 
    }
    for(var i=1;i<=tbl.firstChild.childNodes.length;i++) {
        for(var j=1;j<=10;j++) {
            osheet.cells(i, j).Borders.Weight = 2;
            osheet.cells(i, j).HorizontalAlignment = 3;
            osheet.cells(i, j).VerticalAlignment = 2;
        }
    }
    try {
        courseName = courseName.replace("[", "(");
        courseName = courseName.replace("]", ")");
        var fname = oxl.Application.GetSaveAsFilename(courseName+"成绩录入", "Excel Spreadsheets (*.xlsx), *.xlsx");
    } catch (e) {
        print("Nested catch caught " + e);
    } finally {
        owb.SaveAs(fname);

        owb.Close(savechanges = false);
        //xls.visible = false;
        oxl.Quit();
        oxl = null;
        
        //结束excel进程，退出完成
        //window.setInterval("Cleanup();",1);
        idTmr = window.setInterval("Cleanup();", 1);
    }
}
function scoreUpload(){
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
    console.log("总人数 "+(lineCount-1))
    var failCount = 0;
    for(var i=2;i<=lineCount;i++) {
        var index = i - 2;
        var displayScore = document.getElementById("zcj"+index).value;
        if(!displayScore)
            break;
        var displayNo = document.getElementById("tr"+index).children[2].lastChild.data;
        var displayName = document.getElementById("tr"+index).children[3].firstChild.data;
        
        var realNo = osheet.cells(i, 1).value;
        var realName = osheet.cells(i, 2).value;
        var realScore = osheet.cells(i, 10).value;
        realScore = realScore.toFixed(1);
        if(realNo!=displayNo || realName!=displayName || displayScore!=realScore)
        {
            console.log("学号"+osheet.cells(i, 1).value+" 姓名"+osheet.cells(i, 2).value+" 成绩不对!实际分数为"+realScore);
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
    if(failCount==0) {
        window.alert("校验成功!");
    } else {
        window.alert("校验失败!");
    }
}
function Cleanup() {
    window.clearInterval(idTmr);
    CollectGarbage();
}
createButtons();
</script>