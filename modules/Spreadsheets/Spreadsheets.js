/*+**********************************************************************************
 * The contents of this file are subject to the vtiger CRM Public License Version 1.0
 * ("License"); You may not use this file except in compliance with the License
 * The Original Code is:  vtiger CRM Open Source
 * The Initial Developer of the Original Code is vtiger.
 * Portions created by vtiger are Copyright (C) vtiger.
 * All Rights Reserved.
 ************************************************************************************/

function cbssmvcdowork(work, ssid) {
	fetch(
		'index.php?module=Spreadsheets&action=SpreadsheetsAjax&actionname=sactions&method='+work+'&sid='+ssid,
		{
			method: 'post',
			headers: {
				'Content-type': 'application/x-www-form-urlencoded; charset=UTF-8'
			},
			credentials: 'same-origin',
			body: '&'+csrfMagicName+'='+csrfMagicToken
		}
	).then(response => response.json()).then(response => {
		if (response.status == 'OK') {
			switch (response.msg) {
				case 'Open':
					var ssid = response.notify.substring(0, response.notify.lastIndexOf('?')).substring(response.notify.lastIndexOf('/')+1);
					window.open(response.notify, ssid);
					break;
				default:
					document.getElementById('appnotifydiv').outerHTML = response.notify;
					document.getElementById('appnotifydiv').style.display='block';
					break;
			}
		} else {
			document.getElementById('appnotifydiv').outerHTML = response.notify;
			document.getElementById('appnotifydiv').style.display='block';
		}
	});
}

function cbssmvcOpen(ssid) {
	cbssmvcdowork('openSheet', ssid);
}

function spreadsheetEdit(module, oButton) {
	var select_options = document.getElementById('allselectedboxes').value;
	var excludedRecords = document.getElementById('excludedRecords').value;
	var viewid = getviewId();
	var method_to_call = 'createFromModuleListView';

	//Added to remove the semi colen ';' at the end of the string.done to avoid error.
	var idstring = '';
	var count = 0;
	if (select_options == 'all') {
		document.getElementById('idlist').value=idstring;
		allids = select_options;
		count=numOfRows;
	} else {
		var x = select_options.split(';');
		count=x.length;
		select_options=select_options.slice(0, (select_options.length-1));
		if (count > 1) {
			idstring=select_options.replace(/;/g, ',');
			document.getElementById('idlist').value=idstring;
		} else {
			alert(alert_arr.SELECT);
			return false;
		}
		var allids = document.getElementById('idlist').value;
	}

	fetch(
		'index.php?module=Spreadsheets&action=SpreadsheetsAjax&actionname=sactions&method='+method_to_call+'&allids='+allids+'&sourcemodule='+module+'&viewid='+viewid,
		{
			method: 'post',
			headers: {
				'Content-type': 'application/x-www-form-urlencoded; charset=UTF-8'
			},
			credentials: 'same-origin',
			body: '&'+csrfMagicName+'='+csrfMagicToken
		}
	).then(response => response.json()).then(response => {
		if (response.status == 'OK') {
			switch (response.msg) {
				case 'Open':
					var ssid = response.notify.substring(0, response.notify.lastIndexOf('?')).substring(response.notify.lastIndexOf('/')+1);
					var spwindow = window.open(response.notify, ssid);
					setTimeout(function () {
						populateSpreadshetColumnsWithDropDown(response.sendcommandurl, response.command);
					}, 5 * 1000);
					break;
				default:
					document.getElementById('appnotifydiv').outerHTML = response.notify;
					document.getElementById('appnotifydiv').style.display='block';
					break;
			}
		} else {
			document.getElementById('appnotifydiv').outerHTML = response.notify;
			document.getElementById('appnotifydiv').style.display='block';
		}
	});
}

function populateSpreadshetColumnsWithDropDown(spreadsheeturl, command) {
	sendcommand = 'sendCommandToEtherCalc';
	command = encodeURIComponent(command);
	fetch(
		'index.php?module=Spreadsheets&action=SpreadsheetsAjax&actionname=sactions&method='+sendcommand+'&spurl='+spreadsheeturl+'&command='+command,
		{
			method: 'post',
			headers: {
				'Content-type': 'application/x-www-form-urlencoded; charset=UTF-8'
			},
			credentials: 'same-origin',
			body: '&'+csrfMagicName+'='+csrfMagicToken
		}
	).then(response => response.json()).then(response => {
		if (response.status == 'OK') {
		} else {
		}
	});
}
