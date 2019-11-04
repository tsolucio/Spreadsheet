<?php
/*+***********************************************************************************
 * The contents of this file are subject to the vtiger CRM Public License Version 1.0
 * ("License"); You may not use this file except in compliance with the License
 * The Original Code is:  vtiger CRM Open Source
 * The Initial Developer of the Original Code is vtiger.
 * Portions created by vtiger are Copyright (C) vtiger.
 * All Rights Reserved.
 *************************************************************************************/
include_once 'include/Webservices/Create.php';
include_once 'include/Webservices/Update.php';

function cbws_updatefromethercalc($spreadsheetid, $user) {
	global $log, $adb;
	if (strpos($spreadsheetid, 'x')===false) {
		$spreadsheetid = vtws_getEntityId('Spreadsheets') . 'x' . $spreadsheetid;
	}
	$idList = vtws_getIdComponents($spreadsheetid);
	$webserviceObject = VtigerWebserviceObject::fromId($adb, $spreadsheetid);
	$handlerPath = $webserviceObject->getHandlerPath();
	$handlerClass = $webserviceObject->getHandlerClass();

	require_once $handlerPath;

	$handler = new $handlerClass($webserviceObject, $user, $adb, $log);
	$meta = $handler->getMeta();
	$entityName = $meta->getObjectEntityName($spreadsheetid);

	$types = vtws_listtypes(null, $user);
	if (!in_array($entityName, $types['types'])) {
		throw new WebServiceException(WebServiceErrorCode::$ACCESSDENIED, 'Permission to perform the operation is denied');
	}

	if ($entityName !== $webserviceObject->getEntityName()) {
		throw new WebServiceException(WebServiceErrorCode::$INVALIDID, 'Id specified is incorrect');
	}

	if (!$meta->hasPermission(EntityMeta::$UPDATE, $spreadsheetid)) {
		throw new WebServiceException(WebServiceErrorCode::$ACCESSDENIED, 'Permission to read given object is denied');
	}

	if (!$meta->exists($idList[1])) {
		throw new WebServiceException(WebServiceErrorCode::$RECORDNOTFOUND, 'Record you are trying to access is not found');
	}

	$entity = $handler->retrieve($spreadsheetid);
	$ecUrl = GlobalVariable::getVariable('EtherCalc_URL', '');
	if (empty($ecUrl) || empty($entity['ethercalcid'])) {
		throw new WebServiceException(WebServiceErrorCode::$INVALID_PARAMETER, 'Missing EtherCalc URL or spreadsheet name');
	}
	if (!in_array($entity['spmodule'], $types['types'])) {
		throw new WebServiceException(WebServiceErrorCode::$ACCESSDENIED, 'Permission to perform the operation on module is denied');
	}
	if (isPermitted($entity['spmodule'], 'EditView') != 'yes') {
		throw new WebServiceException(WebServiceErrorCode::$ACCESSDENIED, 'Permission to update module is denied');
	}

	$usemodule = $entity['spmodule'];
	$ecid = $entity['ethercalcid'];
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
		cbws_generaSpreadsheet($record, $ecUrl, $generastring);
	}
	VTWS_PreserveGlobal::flush();
	return '';
}


function cbws_generaSpreadsheet($record, $ecUrl, $allstring) {
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
?>
