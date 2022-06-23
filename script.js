function exportExcelFunc(tableId) {
    var timer = null;
    function  getExplorer() {
        var explorer = window.navigator.userAgent ;
        //ie
        if (explorer.indexOf("MSIE") >= 0) {
            return 'ie';
        }
        //firefox
        else if (explorer.indexOf("Firefox") >= 0) {
            return 'Firefox';
        }
        //Chrome
        else if(explorer.indexOf("Chrome") >= 0){
            return 'Chrome';
        }
        //Opera
        else if(explorer.indexOf("Opera") >= 0){
            return 'Opera';
        }
        //Safari
        else if(explorer.indexOf("Safari") >= 0){
            return 'Safari';
        }
    }

    function clearUp() {
        window.clearInterval(timer);
        CollectGarbage();
    }

    var tableToExcel = (function() {
        var uri = 'data:application/vnd.ms-excel;base64,';
        var template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv="Content-Type" charset="utf-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>';
        function base64(s) {
            return window.btoa(unescape(encodeURIComponent(s)));
        }
        function format(s, c) {
            return s.replace(/{(\w+)}/g, function(m, p) {
                return c[p];
            });
        }
        return function(table, name) {
            if (!table.nodeType) {
                table = document.getElementById(table);
            }
            var ctx = {
                worksheet: name || 'worksheet',
                table: table.innerHTML
            };
            window.location.href = uri + base64(format(template, ctx));
        }
    }) ()

    function getExcel(tableId) {
        if (getExplorer() === 'ie') {
            var currentTB = document.getElementById(tableId);
            var oXL = new ActiveXObject('Excel.Application');
            var oWB = oXL.Workbooks.Add();
            var xlsheet = oWB.Worksheets(1);
            var sel = document.body.createTextRange();
            sel.moveToElementText(currentTB);
            sel.select;
            sel.execCommand('Copy');
            xlsheet.Paste();
            oXL.Visible = true;

            try {
                var fname = oXL.Application.GetSaveAsFilename('haha.xls', 'Excel Spreadsheets (*.xls), *.xls');
            } catch (e) {
                print('Nested catch caught ' + e);
            } finally {
                oWB.SaveAs(fname);

                oWB.Close(savechanges = false);
                //xls.visible = false;
                oXL.Quit();
                oXL = null;
                //window.setInterval("Cleanup();",1);
                timer = window.setInterval('cleanup();', 1);

            }

        }
        else {
            tableToExcel(tableId);
        }

    }
    getExcel(tableId);
}

function saveTableFunc(){
	nomeText = document.getElementById("nomeAdd").value;
	contatosText = document.getElementById("contatosAdd").value;
	atividadesText = document.getElementById("atividadesAdd").value;

	let table = document.getElementById('mytable');
	let thead = document.createElement('tr');
	let nomeTable = document.createElement('td');
	let contatosTable = document.createElement('td');
	let atividadesTable = document.createElement('td');

	nomeTable.innerHTML = nomeText;
	contatosTable.innerHTML = contatosText;
	atividadesTable.innerHTML = atividadesText;

	table.appendChild(thead);
	thead.appendChild(nomeTable);
	thead.appendChild(contatosTable);
	thead.appendChild(atividadesTable);

	document.getElementById("form-container").reset();
	document.getElementById("nomeAdd").focus();
}
