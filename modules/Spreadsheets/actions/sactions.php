<?php
/*************************************************************************************************
 * Copyright 2019 JPL TSolucio, S.L. -- This file is a part of TSOLUCIO coreBOS Customizations.
 * Licensed under the vtiger CRM Public License Version 1.1 (the "License"); you may not use this
 * file except in compliance with the License. You can redistribute it and/or modify it
 * under the terms of the License. JPL TSolucio, S.L. reserves all rights not expressly
 * granted by the License. coreBOS distributed by JPL TSolucio S.L. is distributed in
 * the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. Unless required by
 * applicable law or agreed to in writing, software distributed under the License is
 * distributed on an "AS IS" BASIS, WITHOUT ANY WARRANTIES OR CONDITIONS OF ANY KIND,
 * either express or implied. See the License for the specific language governing
 * permissions and limitations under the License. You may obtain a copy of the License
 * at <http://corebos.org/documentation/doku.php?id=en:devel:vpl11>
 *************************************************************************************************/

class sactions_Action extends CoreBOS_ActionController {

	private function checkQIDParam() {
		$record = isset($_REQUEST['sid']) ? vtlib_purify($_REQUEST['sid']) : 0;
		if (empty($record)) {
			$rdo = array();
			$rdo['status'] = 'NOK';
			$rdo['msg'] = getTranslatedString('LBL_RECORD_NOT_FOUND');
			$smarty = new vtigerCRM_Smarty();
			$smarty->assign('ERROR_MESSAGE', $rdo['msg']);
			$rdo['notify'] = $smarty->fetch('applicationmessage.tpl');
			echo json_encode($rdo);
			die();
		}
		return $record;
	}

	private function checkEtherCalcURL() {
		$ecUrl = GlobalVariable::getVariable('EtherCalc_URL', '');
		$smarty = new vtigerCRM_Smarty();
		if (empty($ecUrl)) {
			$rdo = array();
			$rdo['status'] = 'NOK';
			$rdo['msg'] = getTranslatedString('NOECURL', 'Spreadsheets');
			$smarty->assign('ERROR_MESSAGE_CLASS', 'cb-alert-warning');
			$smarty->assign('ERROR_MESSAGE', $rdo['msg']);
			$rdo['notify'] = $smarty->fetch('applicationmessage.tpl');
			echo json_encode($rdo);
			die();
		}
		$ecUrl = trim($ecUrl, '/').'/';
		return $ecUrl;
	}

	private function createSpreadsheet($record, $ecUrl) {
		global $adb;
		$ch = curl_init();
		curl_setopt($ch, CURLOPT_URL, $ecUrl.'_');
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($ch, CURLOPT_HEADER, false);
		curl_setopt($ch, CURLOPT_POST, true);
		// get columns to update from map
		$cols = 'one,two,three';
		curl_setopt($ch, CURLOPT_POSTFIELDS, $cols);
		curl_setopt($ch, CURLOPT_HTTPHEADER, array('Content-Type: text/csv'));
		$ecname = curl_exec($ch);
		$adb->pquery('update vtiger_spreadsheets set ethercalcid=?', array($ecname));
		return $ecname;
	}

	public function openSheet() {
		global $adb, $site_URL, $current_user;
		$record = $this->checkQIDParam();
		$ecUrl = $this->checkEtherCalcURL();
		$rs = $adb->pquery('select ethercalcid from vtiger_spreadsheets where spreadsheetsid=?', array($record));
		$ecid = $adb->query_result($rs, 0, 0);
		if (empty($ecid)) { // we create it
			$ecid = $this->createSpreadsheet($record, $ecUrl);
		}
		$ecid = trim($ecid, '/');
		$rdo = array();
		$rdo['status'] = 'OK';
		$rdo['msg'] = 'Open';
		$rdo['notify'] = $ecUrl . $ecid . '?usr=' . $current_user->user_name . '&pwd=' . $current_user->accesskey
			. '&url=' . urlencode($site_URL) . '&mtd=updateFromEthercalc&sid='.$record;
		echo json_encode($rdo);
	}
}
?>