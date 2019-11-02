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
include_once 'include/Webservices/Create.php';
include_once 'include/Webservices/Update.php';

class sactions_Action extends CoreBOS_ActionController {

	private function checkQIDParam() {
		$record = isset($_REQUEST['qid']) ? vtlib_purify($_REQUEST['qid']) : 0;
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

	private function generaSpreadsheet($record, $ecUrl, $allstring) {
		global $adb;
		$ch = curl_init();
		curl_setopt($ch, CURLOPT_URL, $ecUrl.'_');
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($ch, CURLOPT_HEADER, false);
		curl_setopt($ch, CURLOPT_POST, true);
		curl_setopt($ch, CURLOPT_POSTFIELDS, $allstring);
		$response=curl_exec($ch);
		$adb->pquery('update vtiger_spreadsheets set ethercalcid=? where spreadsheetsid=?', array($response, $record));
	}

	public function openSheet() {
		global $adb;
		$record = $this->checkQIDParam();
		$ecUrl = $this->checkEtherCalcURL();
		$rs = $adb->pquery('select ethercalcid from vtiger_spreadsheets where spreadsheetsid=?', array($record));
		$ecid = $adb->query_result($rs, 0, 0);
		if (empty($ecid)) { // we create it
			$ecid = $this->createSpreadsheet($record, $ecUrl);
		}
		header('Location: ' . $ecUrl . $ecid);
	}

	public function updateSheet() {
		global $adb, $current_user;
		$record = $this->checkQIDParam();
		$ecUrl = $this->checkEtherCalcURL();
		$rs = $adb->pquery('select ethercalcid,spmodule from vtiger_spreadsheets where spreadsheetsid=?', array($record));
		$usemodule = $adb->query_result($rs, 0, 'spmodule');
		$ecid = $adb->query_result($rs, 0, 'ethercalcid');
		if (empty($ecid)) { // we don't have the sheet
			return false;
		}
		// get columns to update from map
		$colsarr=array();
		// get values from spreadsheet
		$ch = curl_init();
		curl_setopt($ch, CURLOPT_URL, $ecUrl.$ecid.'.csv.json');
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($ch, CURLOPT_HEADER, false);
		$response = curl_exec($ch);
		curl_close($ch);
		// update rows
		$respdecoded=json_decode($response);
		$allrows=array();
		$arrvals=array();
		$arrvalsall=array();
		$arrvalsjson=array();
		$alljson=array();
		$k=0;
		$h=0;
		foreach ($respdecoded as $key => $value) {
			if ($key==0) {
				$arrvals=array_values(array_slice($value, 3, $arr1dim));
				$arrvalsjson=array_values(array_slice($value, $arr2dim));
				$arrvalsall=array_values($value);
			} else {
				$z=0;
				$l=0;
				foreach (array_slice($value, 3, $arr1dim) as $key1 => $value1) {
					$checkfld=$arrvals[$key1];
					$fldtype=$adb->query("select * from vtiger_field where fieldname='$checkfld' and (uitype=5 || uitype=6)");
					$fldtime=$adb->query("select * from vtiger_field where fieldname='$checkfld' and (uitype=2 || uitype=14)");
					if ($adb->num_rows($fldtype)>0 && is_numeric($value1)) {
						$valuenew1=($value1 - 25569) * 86400;
						if (substr($valuenew1, 1)!='-') {
							$value1=gmdate('Y-m-d', $valuenew1);
						}
					} elseif ($adb->num_rows($fldtime)>0 && is_numeric($value1)) {
						$hourformatted=$value1;
						$value1=gmdate('H:i:s', floor($hourformatted * 86400));
					}
					$allrows[$k][$arrvals[$z]]=str_replace('#', ',', $value1);
					$z++;
				}
				foreach (array_slice($value, $arr2dim) as $value2) {
					$alljson[$h][$arrvalsjson[$l]]=str_replace('#', ',', $value2);
					$l++;
				}
				$k++;
				$h++;
			}
		}
		$allnewrows=implode(',', $arrvalsall)."\n";
		$rowact='';
		for ($k=0; $k<count($allrows); $k++) {
			$crthis=array_slice($allrows[$k], 0, count($colsarr));
			if (!is_null($alljson[$k]) && $alljson[$k]!='') {
				$crthis['description']=json_encode($alljson[$k]);
			}
			$crthisnew=json_encode($crthis);
			if ($respdecoded[$k+1][0]=='') {
				$createprojectquery = vtws_create($usemodule, $crthisnew);
				$crmidact=$createprojectquery['result']['id'];
				$createdtimeact=$createprojectquery['result']['createdtime'];
				if ($createdtimeact=='') {
					$createdtimeact=$createprojectquery['result']['CreatedTime'];
				}
				$modifiedtimeact=$createprojectquery['result']['modifiedtime'];
				if ($modifiedtimeact=='') {
					$modifiedtimeact=$createprojectquery['result']['ModifiedTime'];
				}
				$rowact=$rowact.$crmidact.','.$createdtimeact.','.$modifiedtimeact.','.implode(',', $allrows[$k])."\n";
			} else {
				$rowact=$rowact.$respdecoded[$k+1][0].','.$respdecoded[$k+1][1].','.$respdecoded[$k+1][2].','.implode(',', $allrows[$k])."\n";
				$crthis['id']=$respdecoded[$k+1][0];
				$crthisnew=json_encode($crthis);
				vtws_update($crthisnew);
			}
		}
		if ($rowact!='') {
			$generastring=$allnewrows.$rowact;
			$this->generaSpreadsheet($record, $ecUrl, $generastring);
		}
	}
}
?>