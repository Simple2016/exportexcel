var idTmr;

function getExplorer() {
    var explorer = window.navigator.userAgent;
    //ie
    if (explorer.indexOf("MSIE") >= 0) {
        return 'ie';
    }
    //firefox
    else if (explorer.indexOf("Firefox") >= 0) {
        return 'Firefox';
    }
    //Chrome
    else if (explorer.indexOf("Chrome") >= 0) {
        return 'Chrome';
    }
    //Opera
    else if (explorer.indexOf("Opera") >= 0) {
        return 'Opera';
    }
    //Safari
    else if (explorer.indexOf("Safari") >= 0) {
        return 'Safari';
    }
}

function exportExcel(tableid, fileName) {//整个表格拷贝到EXCEL中
    if (getExplorer() == 'ie') {
        var curTbl = document.getElementById(tableid);
        var oXL = new ActiveXObject("Excel.Application");

        //创建AX对象excel
        var oWB = oXL.Workbooks.Add();
        //获取workbook对象
        var xlsheet = oWB.Worksheets(1);
        //激活当前sheet
        var sel = document.body.createTextRange();
        sel.moveToElementText(curTbl);
        //把表格中的内容移到TextRange中
        sel.select;
        //全选TextRange中内容
        sel.execCommand("Copy");
        //复制TextRange中内容
        xlsheet.Paste();
        //粘贴到活动的EXCEL中
        oXL.Visible = true;
        //设置excel可见属性

        try {
            var fname = oXL.Application.GetSaveAsFilename(fileName + ".xls", "Excel Spreadsheets (*.xls), *.xls");
        } catch (e) {
            print("Nested catch caught " + e);
        } finally {
            oWB.SaveAs(fname);

            oWB.Close(savechanges = false);
            //xls.visible = false;
            oXL.Quit();
            oXL = null;
            //结束excel进程，退出完成
            //window.setInterval("Cleanup();",1);
            idTmr = window.setInterval("Cleanup();", 1);

        }

    } else {
        tableToExcel(tableid,'',fileName + ".xls")
    }
}

function Cleanup() {
    window.clearInterval(idTmr);
    CollectGarbage();
}

var tableToExcel = (function () {
    var uri = 'data:application/vnd.ms-excel;filename=exportData.doc;base64,',
        template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
            ' <style type="text/css">' +
            'table td ,table th{' +
            'border: 0.5px solid #000000;' +
            // 'width: 200px;' +
            // 'height: 30px;' +
            // ' text-align: center;' +
            // 'background-color: #4f891e;' +
            // 'color: #ffffff;' +
            ' }' +
            '</style>' +
            '</head><body><table>{table}</table></body></html>',
        base64 = function (s) {
            return window.btoa(unescape(encodeURIComponent(s)))
        },
        format = function (s, c) {
            return s.replace(/{(\w+)}/g,
                function (m, p) {
                    return c[p];
                })
        }
    return function (table, name,filename) {
        if (!table.nodeType) table = document.getElementById(table);
        var ctx = {worksheet: name || 'Worksheet', table: table.innerHTML};
        // window.location.href = uri + base64(format(template, ctx))
        // var atag=document.createElement("a");
        // atag.href=uri + base64(format(template, ctx));
        // atag.download = filename;
        // atag.click();
        //
        // document.getElementById('downExcelChangeBtn').href = uri + base64(format(template, ctx));
        // document.getElementById('downExcelChangeBtn').download =filename;

        download_csv_using_blob(filename, format(template, ctx));
    }
})()



var download_csv_using_blob = function (file_name, content) {
    var csvData = new Blob([content], { type: 'text/csv' });
    // for IE
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {
        window.navigator.msSaveOrOpenBlob(csvData, file_name);
    }
    // for Non-IE (chrome, firefox etc.)
    else {
        var a = document.createElement('a');
        document.body.appendChild(a);
        a.style = 'display: none';
        var url = window.URL.createObjectURL(csvData);
        a.href =  url;
        a.download = file_name;
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    }
};