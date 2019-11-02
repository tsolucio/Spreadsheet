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
		'index.php?module=Spreadsheets&action=SpreadsheetsAjax&actionname=sactions&method='+work+'&qid='+ssid,
		{
			method: 'post',
			headers: {
				'Content-type': 'application/x-www-form-urlencoded; charset=UTF-8'
			},
			credentials: 'same-origin',
			body: '&'+csrfMagicName+'='+csrfMagicToken
		}
	).then(response => response.json()).then(response => {
		document.getElementById('appnotifydiv').outerHTML = response.notify;
		document.getElementById('appnotifydiv').style.display='block';
	});
}

function cbssmvcOpen(ssid) {
	cbssmvcdowork('openSheet', ssid);
}
