// ==UserScript==
// @name         Tier 3 Escalation Tracker
// @namespace    https://teams.microsoft.com/
// @version      0.1
// @description  Track and log daily T3 escalations.
// @author       Zachary Kitcher
// @include      https://teams.microsoft.com/*
// @require      https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js
// @updateURL    https://github.com/ZKitcher/MS_Teams_Scripts/raw/master/Tier%203%20Escalation%20Tracker.user.js
// ==/UserScript==

var date = getDate();
var escalationJSON = [];

(function() {
    'use strict';
    var update = false;
    $('body').on('DOMSubtreeModified', function() {
        if(update == false){
            update = true;
            var pageHeader = $('h2[data-tid="messagesHeader-Title"]').attr('title');
            if(pageHeader == 'Tier 3 Escalations'){
                appendExportBtn();
                scanAvailableEntries();
            }else{
                $('#T3EscalationLog').remove();
            };
            setTimeout(function(){
                update = false;
            }, 1000);
        };
    });

    setInterval(function(){
        dailyDownload();
    }, 60000);
})();

function scanAvailableEntries(){
    var index = true;
    $('[data-scroll-pos]').each(function(){
        var timeStamp = $(this).find('[data-tid="messageTimeStamp"]').attr('title')
        if(timeStamp == undefined){
            return false;
        };
        if(index){
            $('#T3EscalationLog').find('[class="tab-display-name"]').text('Export')
        };
        if(timeStamp.includes(date[0])){
            if(index){
                $('#T3EscalationLog').find('[class="tab-display-name"]').text('Scroll up, top entry ' + timeStamp)
            };
            var id = $(this).attr('data-scroll-id');
            var agent = $(this).find('[data-tid="threadBodyDisplayName"]').first().text().replace(/\s  +/gm, '');
            var entry = $(this).find('[data-tid="messageBodyContent"]:eq(0)').text();
            var URL = [];
            $(this).find('[data-tid="messageBodyContent"]').find('a[href]').each(function(){
                URL.push($(this).attr('href'))
            });
            if(IDdupeCheck(id, escalationJSON)){
                escalationJSON.push({id, agent, timeStamp, entry, URL})
            };
        };
        index = false;
    });
    console.log(escalationJSON)
};

function dailyDownload(){
    var time = getTime();
    if(time == '23:59'){
        exportEscalation();
        escalationJSON = [];
        setTimeout(function(){
            date = getDate();
        }, 60000);
    };
};

function appendExportBtn(){
    if($('#T3EscalationLog').length == 0){
        var exportBtn = `
<li id="T3EscalationLog" dnd-draggable="tab" dnd-type="'tabType'" role="presentation" draggable="true">
  <div title="Export Escalation" role="presentation">
    <a class="btn btn-default" style="max-width: max-content;">
      <span class="tab-btn-container">
        <span class="tab-display-name">Export</span>
      </span>
    </a>
  </div>
</li>
`
        $('[data-tid="appTabs"]').append(exportBtn)

        $('#T3EscalationLog').click(function(){
            exportEscalation();
        });
    };
};

function exportEscalation(){
    buildExportTable()
    exportTableFormatting()
    fnExcelReport($('#exportTable'))
    $('#exportTable').remove();
};

function exportTableFormatting(){
    var URLs = [];
    $('[name="exportURL"]').each(function(){
        const regex = /https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.;~#?&//=]*)/g
        var text = $(this).html()
        if(text.match(regex)){
            var url = text.match(regex);
            buildHref($(this), url)
        };
    });
};

function buildHref($entry, urlArray){
    const newSet = new Set(urlArray);
    urlArray = [...newSet];
    urlArray.map(e => $entry.html($entry.html().replace(convertToRegex(e), '<a href='+e+' target="_blank">'+e+'</a>')))
};

function convertToRegex(query){
    var regex = query.replace(/\//g,'\\/');
    regex = regex.replace(/\?/g,'\\?');
    regex = new RegExp(regex, 'g')
    return regex;
};

function buildExportTable(){
    var exportTable = `
<table id="exportTable"
<tbody name="exportTbody">
  <tr>
    <td colspan="4" style="font-weight: bold;font-size: large;">Tier 3 Escalations `+date[1]+`</td>
  </tr>
  <tr>
    <td colspan="4" style="font-weight: bold;">Total Escalations: `+escalationJSON.length+`</td>
  </tr>
  <tr style="font-weight: bold;">
    <td>Agent</td>
    <td>Date</td>
    <td>Escalation</td>
    <td>URLs</td>
  </tr>
</tbody>
</table>
`
    $('body').append(exportTable)
    for(var i = 0; i < escalationJSON.length; i++){
        var tableEntry = `
<tr>
  <td name="exportAgent" style="vertical-align: middle;">`+escalationJSON[i].agent+`</td>
  <td name="exportTimesStamp" style="vertical-align: middle;">`+escalationJSON[i].timeStamp+`</td>
  <td name="exportEntry" style="width:700px;">`+escalationJSON[i].entry+`</td>
  <td name="exportURL" style="vertical-align: middle;width:400px;">`+escalationJSON[i].URL+`</td>
</tr>
`
        $('[name="exportTbody"]').append(tableEntry)
    };
};

function IDdupeCheck(query, JSON){
    var found = true;
    for(var i = 0; i < JSON.length; i++){
        if(query == JSON[i].id){
            found = false;
        }
    };
    return found;
};

function getDate(){
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec']
    var today = new Date();
    var dd = String(today.getDate())
    var mm = String(today.getMonth() + 1)
    var yyyy = today.getFullYear();

    today = [months[mm - 1] + ' ' + dd, mm + '-' + dd + '-' + yyyy]
    return today;
};

function fnExcelReport(table) {
    var tab_text = ['<html xmlns:x="urn:schemas-microsoft-com:office:excel">',
                    '<head><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>',
                    '<x:Name>',
                    'Tier 3 Escalations ' + date[1],
                    '</x:Name>',
                    '<x:WorksheetOptions><x:Panes></x:Panes></x:WorksheetOptions></x:ExcelWorksheet>',
                    '</x:ExcelWorksheets></x:ExcelWorkbook></xml></head><body>',
                    "<table border='1px'>",
                    $(table).html(),
                    '</table></body></html>']
    tab_text = tab_text.join('');
    exportExcel(tab_text)
};

function exportExcel(tab_text){
    var data_type = 'data:application/vnd.ms-excel';
    var link = document.createElement('a');
    link.download = 'Tier 3 Escalations ' + date[1] + '.xls';
    link.href = data_type + ', ' + encodeURIComponent(tab_text);
    link.click();
    link.remove();
};

function getTime(){
    var time = new Date();
    var AMPM = ['AM','PM']
    var h = time.getHours();
    var m = time.getMinutes()
    /*
    if(m < 10){
        m = '0' + m;
    };
    if(h > 11){
        if(h != 12){
            h = h - 12
        };
        return h + ':' + m + AMPM[1]
    }else{
        return h + ':' + m + AMPM[0]
    };*/
    return h + ':' + m;
};
