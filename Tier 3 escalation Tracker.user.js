// ==UserScript==
// @name         Tier 3 Escalation Tracker
// @namespace    https://teams.microsoft.com/
// @version      0.6
// @description  Track and log daily T3 escalations.
// @author       Zachary Kitcher
// @include      https://teams.microsoft.com/*
// @require      https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js
// @updateURL    https://github.com/ZKitcher/MS_Teams_Scripts/raw/master/Tier%203%20Escalation%20Tracker.user.js
// ==/UserScript==

var date = getDate();
var escalationJSON = [];
var dlTime;
loadDownloadTime();

(function() {
    'use strict';
    var update = false;
    $('body').on('DOMSubtreeModified', function() {
        if(update == false){
            update = true;
            var pageHeader = $('h2[data-tid="messagesHeader-Title"]').attr('title');
            if(pageHeader == 'Tier 3 Escalations'){
                appendExportBtn();

                runFunctions()

            }else{
                $('[name*="escalationLoghtml"]').remove();
            };
            setTimeout(function(){
                update = false;
            }, 500);
        };
    });

    setInterval(function(){
        dailyDownload();
    }, 60000);
})();

function runFunctions(){
    if($('#todaysEscalation').prop('checked') || $('#specificDate').prop('checked')){
        scanAvailableEntries();
    }else if($('#daysPrevious').prop('checked')){
        scanBackEntries();
    };
}

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
};

function scanBackEntries(){
    var monthsArray = ['Jan', 1, 'Feb', 2, 'Mar', 3, 'Apr', 4, 'May', 5, 'Jun', 6, 'Jul', 7, 'Aug', 8, 'Sept', 9, 'Oct', 10, 'Nov', 11, 'Dec', 12];
    var selectedMonth = date[0].replace(/\W\d+/g, '');
    selectedMonth = monthsArray[monthsArray.indexOf(selectedMonth) + 1];

    var selectedDay = parseInt(date[0].replace(/\D+/g, ''));

    var index = true;

    $('[data-scroll-pos]').each(function(){
        var timeStamp = $(this).find('[data-tid="messageTimeStamp"]').attr('title')
        if(timeStamp == undefined){
            return false;
        };
        if(index){
            $('#T3EscalationLog').find('[class="tab-display-name"]').text('Export')
        };
        var queryMonth = timeStamp.substring(0,3);
        queryMonth = monthsArray[monthsArray.indexOf(queryMonth) + 1];

        var queryDay = timeStamp.substring(0,timeStamp.indexOf(','));
        queryDay = parseInt(queryDay.replace(/\D+/g, ''));

        var push = false;

        if(selectedMonth < queryMonth){
            push = true;
        }else if(selectedMonth == queryMonth && selectedDay <= queryDay){
           push = true;
        };

        if(push){
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
                escalationJSON.push({id, agent, timeStamp, entry, URL, queryMonth, queryDay})
            };
        };
        index = false;
    });
};

function dailyDownload(){
    var time = getTime();
    if(time == dlTime){
        exportEscalation();
        escalationJSON = [];
        date = getDate();
    };
};

function appendExportBtn(){
    if($('#T3EscalationLog').length == 0){
        date = getDate()
        var exportBtn = `
<li id="T3EscalationLog" name="escalationLoghtml" dnd-draggable="tab" dnd-type="'tabType'" role="presentation" draggable="true">
  <div title="Export Escalation" role="presentation">
    <a class="btn btn-default" style="max-width: max-content;">
    <span class="tab-btn-container">
    <span class="tab-display-name">Export</span>
    </span>
    </a>
  </div>
</li>
<li name="escalationLoghtml escalationSettings" dnd-draggable="tab" dnd-type="'tabType'" role="presentation" draggable="true">
  <div title="Export Settings" role="presentation">
    <a class="btn btn-default" style="max-width: max-content;">
    <span class="tab-btn-container">
    <span id="escalationSettingsBtn" class="tab-display-name">?</span>
    </span>
    </a>
  </div>
</li>
<li name="escalationLoghtml searchSelect" dnd-draggable="tab" dnd-type="'tabType'" role="presentation" draggable="true" style="visibility: hidden;">
  <div title="Export Escalation" role="presentation">
    <a class="btn btn-default" style="max-width: max-content; ">
      <select id="searchMonth">
        <option></option>
        <option value="Jan">Jan</option>
        <option value="Feb">Feb</option>
        <option value="Mar">Mar</option>
        <option value="Apr">Apr</option>
        <option value="May">May</option>
        <option value="Jun">Jun</option>
        <option value="Jul">Jul</option>
        <option value="Aug">Aug</option>
        <option value="Sep">Sep</option>
        <option value="Oct">Oct</option>
        <option value="Nov">Nov</option>
        <option value="Dec">Dec</option>
      </select>
    </a>
  </div>
</li>
<li name="escalationLoghtml searchSelect" dnd-draggable="tab" dnd-type="'tabType'" role="presentation" draggable="true" style="visibility: hidden;">
  <div title="Export Escalation" role="presentation">
    <a class="btn btn-default" style="max-width: max-content;">
      <select id="searchDay">
        <option></option>
        <option value="1">01</option>
        <option value="2">02</option>
        <option value="3">03</option>
        <option value="4">04</option>
        <option value="5">05</option>
        <option value="6">06</option>
        <option value="7">07</option>
        <option value="8">08</option>
        <option value="9">09</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
        <option value="13">13</option>
        <option value="14">14</option>
        <option value="15">15</option>
        <option value="16">16</option>
        <option value="17">17</option>
        <option value="18">18</option>
        <option value="19">19</option>
        <option value="20">20</option>
        <option value="21">21</option>
        <option value="22">22</option>
        <option value="23">23</option>
        <option value="24">24</option>
        <option value="25">25</option>
        <option value="26">26</option>
        <option value="27">27</option>
        <option value="28">28</option>
        <option value="29">29</option>
        <option value="30">30</option>
        <option value="31">31</option>
      </select>
    </a>
  </div>
</li>

<div id="exportSettings" name="escalationLoghtml" class="popover tabs-overflow-dropdown minwidth-md maxwidth-common app-default-menu am-fade bottom" acc-role-dom="dialog" role="dialog" aria-modal="true" style="position: fixed; display: block; visibility: hidden;">
  <div class="scroll-container">
    <div class="ts-tabs-dropdown-list simple-scrollbar" simple-scrollbar="">
      <ul class="app-default-menu-ul" acc-role-dom="dialog menu" role="menu" aria-modal="true" kb-list="" kb-cyclic="">
        <li ng-repeat="tab in mh.tabs" ng-if="tab.isOverflow" dnd-type="'tabType'" acc-role-dom="menu-item focus-default" role="menuitem" draggable="true" class="kb-active" tabindex="0">
          <a class="ts-sym" draggable="true" tabindex="-1">
            <div class="button-content-container">
              <input id="todaysEscalation" type="radio" name="radSetting">
              <span class="ts-popover-label" name="radSettingSpan">Todays Escalations</span>
            </div>
          </a>
        </li>
        <li ng-repeat="tab in mh.tabs" ng-if="tab.isOverflow" dnd-type="'tabType'" acc-role-dom="menu-item focus-default" role="menuitem" draggable="true" class="kb-active" tabindex="0">
          <a class="ts-sym" draggable="true" tabindex="-1">
            <div class="button-content-container">
              <input id="specificDate" type="radio" name="radSetting">
              <span class="ts-popover-label" name="radSettingSpan">Specific Date</span>
            </div>
          </a>
        </li>
        <li ng-repeat="tab in mh.tabs" ng-if="tab.isOverflow" dnd-type="'tabType'" acc-role-dom="menu-item focus-default" role="menuitem" draggable="true" class="kb-active" tabindex="0">
          <a class="ts-sym" draggable="true" tabindex="-1">
            <div class="button-content-container">
              <input id="daysPrevious" type="radio" name="radSetting">
              <span class="ts-popover-label" name="radSettingSpan">Log Days Previous</span>
            </div>
          </a>
        </li>
        <li ng-repeat="tab in mh.tabs" ng-if="tab.isOverflow" dnd-type="'tabType'" acc-role-dom="menu-item focus-default" role="menuitem" class="kb-active" tabindex="0">
          <a class="ts-sym" tabindex="-1">
            <div class="button-content-container">
              <input id="downloadTime" style="width: 20%;">
              <span class="ts-popover-label" name="radSettingSpan" style="width: 60px;">Automatic Download Time</span>
            </div>
          </a>
        </li>
      </ul>
      <div class="tse-scrollbar simple-scrollbar">
        <div class="drag-handle handle-hidden"></div>
      </div>
    </div>
  </div>
</div>
`

        $('[data-tid="appTabs"]').append(exportBtn)
        $('#downloadTime').val(dlTime)
        $('#todaysEscalation').prop('checked', true)

        $('#T3EscalationLog').click(function(){
            exportEscalation();
        });

        $('[name*="escalationSettings"]').click(function(){
            var pos = $(this).offset();
            $('#exportSettings').css({'top':pos.top + 20, 'left':pos.left + 5})

            if($('#exportSettings').css('visibility') == 'hidden'){
                $('#exportSettings').css('visibility', 'visible');
            }else{
                $('#exportSettings').css('visibility', 'hidden');
            };
            $('[name*="searchSelect"]').css('visibility', 'visible');
        });

        $('[name="radSetting"]').click(function(){
            escalationJSON = [];
            runFunctions();
        });

        $('[name="radSettingSpan"]').click(function(){
            $(this).prev().prop('checked', true)
        });

        $('[name*="searchSelect"]').on('change', function(){
            var month = $('#searchMonth').val();
            var day = $('#searchDay').val();
            if(month != '' && day != ''){
                escalationJSON = [];
                var today = new Date();
                var yyyy = today.getFullYear();
                date = [month + ' ' + day, month + '-' + day + '-' + yyyy]
            }else{
                escalationJSON = [];
                date = getDate();
            };
        });
        $(document).click(function(e){
            if(!$(e.target).is('#exportSettings') && !$(e.target).is('#escalationSettingsBtn') && !$(e.target).is('[name="radSettingSpan"]') && !$(e.target).is('[name="radSetting"]') && !$(e.target).is('#downloadTime')){
                $('#exportSettings').css('visibility', 'hidden');
            };
        });
        $('#downloadTime').on('input', function(){
            confirmTimeInput();
        });
    };
};

function confirmTimeInput(){
    var valid = false;
    var input = $('#downloadTime').val();

    if(input.match(/:/g).length == 1){
        var hour = input.split(':')[0];
        var min = input.split(':')[1];
        if(!hour.match(/\D/g) && !min.match(/\D/g)){
            if(hour.length == 2 && min.length == 2){
                if((hour > 0 && hour < 25) && (min > -1 && min < 60)){
                    valid = true;
                    dlTime = input;
                };
            };
        };
    };

    if(valid){
        $('#downloadTime').css('background', 'green')
        saveDownloadTime();
    }else{
        $('#downloadTime').css('background', 'red')
    };
};

function saveDownloadTime(){
    localStorage.setItem('DownloadTime', dlTime);
};

function loadDownloadTime(){
    dlTime = localStorage.getItem('DownloadTime')
    if(dlTime == '' || dlTime == null){
        dlTime = '24:00';
    };
    console.log(dlTime);
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
    var timeFrame
    if($('#todaysEscalation').prop('checked') || $('#specificDate').prop('checked')){
        timeFrame = date[1];
    }else if($('#daysPrevious').prop('checked')){
        var today = getDate();
        timeFrame = date[1] + ' - ' + today[1];
    };

    var exportTable = `
<table id="exportTable"
<tbody name="exportTbody">
  <tr>
    <td id="exportHeading" timeframe="`+timeFrame+`" colspan="4" style="font-weight: bold;font-size: large;">Tier 3 Escalations `+timeFrame+`</td>
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
    $('body').append(exportTable);

    var tableEntry;
    if($('#todaysEscalation').prop('checked') || $('#specificDate').prop('checked')){
        for(var i = 0; i < escalationJSON.length; i++){
            tableEntry = `
<tr>
  <td name="exportAgent" style="vertical-align: middle;">`+escalationJSON[i].agent+`</td>
  <td name="exportTimesStamp" style="vertical-align: middle;">`+escalationJSON[i].timeStamp+`</td>
  <td name="exportEntry" style="width:700px;">`+escalationJSON[i].entry+`</td>
  <td name="exportURL" style="vertical-align: middle;width:400px;">`+escalationJSON[i].URL+`</td>
</tr>
`
            $('[name="exportTbody"]').append(tableEntry)
        };
    }else if($('#daysPrevious').prop('checked')){
        sortJSON()
        for(var j = 0; j < escalationJSON.length; j++){
            var dateHeader = escalationJSON[j].timeStamp
            dateHeader = dateHeader.substring(0,dateHeader.indexOf(','));
            dateHeader = dateHeader.replace(' ', '-');
            if($('[dateHeader="'+dateHeader+'"]').length == 0){
                var header = `
                <tr dateHeader="`+dateHeader+`">
                    <td colspan="4" style="font-weight: bold; font-size: large; text-align: left;">`+ dateHeader +`</td>
                </tr>`
                $('[name="exportTbody"]').append(header)
            };
            tableEntry = `
<tr>
  <td name="exportAgent" style="vertical-align: middle;">`+escalationJSON[j].agent+`</td>
  <td name="exportTimesStamp" style="vertical-align: middle;">`+escalationJSON[j].timeStamp+`</td>
  <td name="exportEntry" style="width:700px;">`+escalationJSON[j].entry+`</td>
  <td name="exportURL" style="vertical-align: middle;width:400px;">`+escalationJSON[j].URL+`</td>
</tr>
`
            $(tableEntry).insertAfter('[dateHeader="'+dateHeader+'"]')
        };
    };
};

function sortJSON(){
    console.log(escalationJSON)
    var masterArray = [];
    var tempArray = [];
    var lower = 50, upper = 0;
    var pointer;
    for(var i = 0; i < escalationJSON.length; i++){
        pointer = escalationJSON[i].queryMonth
        if(pointer > upper){
            upper = pointer
        };
        if(pointer < lower){
            lower = pointer
        };
    };

    while(lower <= upper){
        for(var j = 0; j < escalationJSON.length; j++){
            if(escalationJSON[j].queryMonth == lower){
                tempArray.push(escalationJSON[j]);
            };
        };
        masterArray.push(tempArray)
        tempArray = [];
        lower++
    };

    for(var k = 0; k < masterArray.length; k++){
        masterArray[k] = sortByKey(masterArray[k], 'queryDay')
    };

    console.log(masterArray)
    escalationJSON = masterArray.flat();
};

function sortByKey(array, key){
    return array.sort(function(a, b){
        var x = a[key]; var y = b[key];
        return ((x < y) ? -1 : ((x > y) ? 1 : 0));
    });
}

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

    today = [months[mm - 1] + ' ' + dd, months[mm - 1] + '-' + dd + '-' + yyyy]
    return today;
};

function fnExcelReport(table) {
    var tab_text = ['<html xmlns:x="urn:schemas-microsoft-com:office:excel">',
                    '<head><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>',
                    '<x:Name>',
                    'Tier 3 Escalations ' + $('#exportHeading').attr('timeframe'),
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
    link.download = 'Tier 3 Escalations ' + $('#exportHeading').attr('timeframe') + '.xls';
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
