// ==UserScript==
// @name         Teams Log Download
// @namespace    http://tampermonkey.net/
// @version      0.6
// @description  Download the current teams channel log.
// @author       Zachary Kitcher
// @include      https://teams.microsoft.com*
// @require      https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js
// @require      https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.js
// @updateURL    
// ==/UserScript==

(function() {
    'use strict';
    var url = "data:text/html;charset=UTF-8,";
    var channelTitle, log, entry, title;
    var loading = 3;
    var style = '<style type="text/css"> <!-- p { margin-left: 40px; } --> </style>';

    setInterval(function(){
        if ($('a[id="expandButton"]').length === 0){
            $('<a id="expandButton" style="cursor: pointer; font-size: 17px;">Expand All Entries</a>').insertAfter($('#settingsDropdown'));

            $('#expandButton').click(function() {
                $('<a id="scanningEntries" style="font-size: 17px;">Scroll Up to Scan Entries. Current entry: </a>').insertAfter($(this));
                $(this).hide();

                log = "";
                entry = "";

                var entryCount = 0;
                var headerCounter = 0;
                var topCount = 0;

                $('<a id="loadingIcon" style="font-size: 17px;">.</a>').insertAfter('#scanningEntries');
                var loadingInterval = setInterval(function() {
                    if(loading < 3){
                        loading++;
                    }else{
                        loading = 1;
                    };
                    if(loading == 3){
                        $('a[id="loadingIcon"]').text(entryCount+"...");
                    }else if(loading == 2){
                        $('a[id="loadingIcon"]').text(entryCount+".. ");
                    }else{
                        $('a[id="loadingIcon"]').text(entryCount+".  ");
                    };
                }, 500);

                var expantInterval = setInterval(function() {
					if(entryCount == 0){
						$('a[id="scanningEntries"]').text("Scroll down to the latest entry. # of entries scanned: ");
					}else{
						$('a[id="scanningEntries"]').text("Scroll Up to Scan Entries. # of entries scanned: ");
					};
					if ($('a[id="scanningEntries"]').length === 0){
						clearInterval(expantInterval);
						log = "";
						entry = "";
						console.log("Scan Interrupted");
					};

                    //Expand all rolled up threads
                    $('div[data-tid="threadViewToggle"]').each(function(){
                        if ($(this).attr("title").includes("Collapse all") == false){
                            $(this).click();
                        };
                    });

                    $('div[data-scroll-pos]').each(function(){
                        title = ("<h1>" + $('h2[class="ts-title team-name"]').text() + "</h1><br>");

                        if($(this).attr('data-scroll-pos') == entryCount){

                            var systemEntryCheck = $(this).find('system-message[data-tid]').attr('data-tid');
                            var deletedMessageCheck = $(this).find('div[id*="deleted-message"]').attr('id')
                            if(systemEntryCheck == null && deletedMessageCheck == null){
                                console.log(entryCount);
                                $(this).find('div[scroll-item-id]').each(function(){
                                    if($(this).attr("scroll-item-id") != null){
                                        var firstEntryCheck = $(this).attr("scroll-item-id") //First Message
                                        var agentName = $(this).find('div[data-tid="threadBodyDisplayName"]').html().replace(/\s\B/g, '') //Name of Agent
                                        var entryTimeSent = $(this).find('span[data-tid="messageTimeStamp"]').html().replace(/\s\B/g, '') //Time sent
                                        var entryBody = $(this).find('div[data-tid="messageBodyContent"]').html().replace(/#/g, '-') //Entry Body

                                        if (firstEntryCheck.includes("firstMessage") == true){
                                            log = "<br>"+ entry +"<br>"+ log;
                                            entry = "";
                                            entry = "<b>"+ agentName +"</b>"+ entryTimeSent+"<br>" +entryBody + entry;
                                        }else{
                                            entry = entry +"<p><b>"+ agentName +"</b>"+ entryTimeSent +"<br>"+ entryBody.replace(/<div>/g, '').replace(/<div/g, '<p')+ "</p>";
                                        };
                                    };
                                });
                                entryCount++;
                            }else{
                                console.log("System Message");
                                entryCount++;
                            };
                        };
                    });
                    var topOfPageCheck = $('div[class*="list-wrap list-wrap-v3 ts-message-list-container"]').position('top').top;
                    if(topOfPageCheck == 0 && entryCount != 0){
						$('a[id="scanningEntries"]').text("Formatting Download");
                        topCount++;
                        console.log("Top of Page count " + topCount);
                        if(topCount == 10){
                            clearInterval(expantInterval);
                            clearInterval(loadingInterval);
                            $('a[id="loadingIcon"]').hide();
                            log = style + title +entry + log;
                            var htmlTitle = $('h2[class="ts-title channel-name"]').text();
                            console.log(log);
                            $('<a id="downloadLog" style="font-size: 17px;">Download Channel Log</a>').insertAfter('a[id="scanningEntries"]');
                            $('a[id="scanningEntries"]').hide();
                            $('a[id="downloadLog"]').attr({
                                href:("data:text/html;charset=UTF-8,"+log),
                                download:(htmlTitle+".html")
                            });
                        };
                    }else{
                        topOfPageCheck = 0;
                    };
                }, 100);
            });
        };
    }, 1000);
})()
