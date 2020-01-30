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
include_once 'include/Webservices/DescribeObject.php';

function cbws_updatefromethercalc($spreadsheetid, $user) {
	global $log, $adb, $current_language, $default_language;
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
	$ecUrl = trim($ecUrl, '/').'/';
	if (!in_array($entity['spmodule'], $types['types'])) {
		throw new WebServiceException(WebServiceErrorCode::$ACCESSDENIED, 'Permission to perform the operation on module is denied');
	}
	if (isPermitted($entity['spmodule'], 'EditView') != 'yes') {
		throw new WebServiceException(WebServiceErrorCode::$ACCESSDENIED, 'Permission to update module is denied');
	}

	$user_current_language = $current_language;
	$current_language = $default_language;

	//$usemodule = $entity['spmodule'];
	$ecid = str_replace('/', '', $entity['ethercalcid']);
	// get columns to update from map
	$rs = $adb->pquery('select map, spmodule from vtiger_spreadsheets where ethercalcid=?', array($entity['ethercalcid']));
	$mapid = $adb->query_result($rs, 0, 'map');
	$sp_module = $adb->query_result($rs, 0, 'spmodule');

	$moduleinfo = vtws_describe($sp_module, $user);
	$module_fieldname_label_key_pairs = array();
	foreach ($moduleinfo['fields'] as $value) {
		$module_fieldname_label_key_pairs[$value['label']] = $value['name'];
		$module_fieldname_uitype[$value['name']] = isset($value['uitype']) ? $value['uitype'] : 0;
	}

	$mapped_field_array =  array();
	if (!empty($mapid) && !empty($sp_module)) {
		$cbMap = cbMap::getMapByID($mapid);
		$spreadsheet_mod_ins = CRMEntity::getInstance($sp_module);
		$mapped_field_array = $cbMap->Mapping($spreadsheet_mod_ins->column_fields, $mapped_field_array);
	}

	$fieldname_array = array();
	$vtlibmod = Vtiger_Module::getInstance($sp_module);
	foreach ($mapped_field_array as $fieldname => $value) {
		$fld = Vtiger_Field::getInstance($fieldname, $vtlibmod);
		if ($fld) {
			$fieldname_array[] = $fld->column;
		} else {
			$fieldname_array[] = $fieldname;
		}
		// if (empty($module_fieldname_label_key_pairs[$fieldname]) {
		// 	$trans_col_array[] = getTranslatedString($fieldname, $sp_module);
		// } else {
		// 	$translabel = getTranslatedString($module_fieldname_label_key_pairs[$fieldname], $sp_module);
		// 	if (empty($translabel)) {
		// 		$trans_col_array[] = getTranslatedString($fieldname, $sp_module);
		// 	} else {
		// 		$trans_col_array[] = $translabel;
		// 	}
		// }
	}

	$current_language = $user_current_language;
	//$colsarr = implode(',', $fieldname_array);
	// get values from spreadsheet
	$ch = curl_init();
	curl_setopt($ch, CURLOPT_URL, $ecUrl.$ecid.'.csv.json');
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
	curl_setopt($ch, CURLOPT_HEADER, false);
	$response = curl_exec($ch);
	curl_close($ch);
	// update rows
	$respdecoded = json_decode($response);
	$sp_all_rows = array();
	$sp_column_headers = array();
	$index_sp_records_with_header = 0;
	foreach ($respdecoded as $key => $value) {
		if ($key == 0) {
			foreach ($value as $index => $col_header) {
				$sp_column_headers[] = $module_fieldname_label_key_pairs[$col_header];
			}
		} else {
			$index_sp_records_with_no_header = 0;
			foreach ($value as $index => $field_value) {
				if ($module_fieldname_uitype[$sp_column_headers[$index]]==5 && is_numeric($field_value)) {
					$new_field_value = ($field_value - 25569) * 86400;
					if (substr($new_field_value, 1) != '-') {
						$field_value = gmdate('Y-m-d', $new_field_value);
					}
				} elseif ($module_fieldname_uitype[$sp_column_headers[$index]]==14 && is_numeric($field_value)) {
					$hourformatted = $field_value;
					$field_value = gmdate('H:i:s', floor($hourformatted * 86400));
				} elseif ($module_fieldname_uitype[$sp_column_headers[$index]]==50 && is_numeric($field_value)) {
					$new_field_value = ($field_value - 25569) * 86400;
					if (substr($new_field_value, 1) != '-') {
						$field_value = gmdate('Y-m-d H:i:s', $new_field_value);
					}
				}
				$sp_all_rows[$index_sp_records_with_header][$sp_column_headers[$index_sp_records_with_no_header]] = str_replace('#', ',', $field_value);
				$index_sp_records_with_no_header++;
			}
			$index_sp_records_with_header++;
		}
	}
	$message = '';
	for ($row=0; $row < count($sp_all_rows); $row++) {
		// Check if the Row on its first Column is Empty
		// If is Yes means the Record is New
		if ($respdecoded[$row + 1][0] == '') {
			foreach ($moduleinfo['fields'] as $value) {
				if (empty($value['uitype'])) { // ID field
					continue;
				}
				if ($value['uitype'] == 10) {
					$fldname = $value['name'];
					$fldmandatorystatus = $value['mandatory'];
					$fldvalue = $sp_all_rows[$row][$fldname];
					if (empty($fldmandatorystatus) || $fldmandatorystatus != 1) {
						if (empty($fldvalue) || $fldvalue == '0') {
							$sp_all_rows[$row][$fldname] = 0;
						} else {
							$newfldvalue =  explode('   ', $fldvalue);
							if (count($newfldvalue) == 2) {
								$sp_all_rows[$row][$fldname] = $newfldvalue[1];
							}
						}
					} else {
						if (empty($fldvalue) || $fldvalue == '0') {
							$message =  getTranslatedString($value['label'], $sp_module).' '.getTranslatedString('LBL_EMPTY_REQUIRED_FIELD');
							break;
						} else {
							$newfldvalue =  explode('   ', $fldvalue);
							if (count($newfldvalue) == 2) {
								$sp_all_rows[$row][$fldname] = $newfldvalue[1];
							}
						}
					}
				} elseif ($value['uitype'] == 53 || $value['uitype'] == 77 || $value['uitype'] == 101) { #for assigned user
					$fldname = $value['name'];
					$fldvalue = $sp_all_rows[$row][$fldname];
					$newfldvalue =  explode('   ', $fldvalue);
					if (count($newfldvalue) == 2) {
						$sp_all_rows[$row][$fldname] = $newfldvalue[1];
					} else {
						$sp_all_rows[$row][$fldname] = vtws_getEntityId('Users').'x'.$user->id;
					}
				}
			}
			unset($sp_all_rows[$row]['']);
			unset($sp_all_rows[$row]['id']);
			$createrecord = vtws_create($sp_module, $sp_all_rows[$row], $user);
			$rowindex = $row + 2;
			$updatecommand = '{"command":"set A'.$rowindex.' text t '.$createrecord['id'].'"}';
			$spreedsheeturl = GlobalVariable::getVariable('EtherCalc_URL', '').'_/'.$ecid;
			$command_response = updateCRMIDColumnEtherCalc($spreedsheeturl, $updatecommand);
			if ($command_response) {
				$message = getTranslatedString('SUCCESS_SAVE', 'Spreadsheets');
			} else {
				$message = getTranslatedString('FAIL_SAVE', 'Spreadsheets');
			}
		} else {
			foreach ($moduleinfo['fields'] as $value) {
				if (empty($value['uitype'])) { // ID field
					continue;
				}
				if ($value['uitype'] == 10) {
					$fldname = $value['name'];
					$fldmandatorystatus = $value['mandatory'];
					$fldvalue = $sp_all_rows[$row][$fldname];
					if (empty($fldmandatorystatus) || $fldmandatorystatus != 1) {
						if (empty($fldvalue) || $fldvalue == '0') {
							$sp_all_rows[$row][$fldname] = '';
						} else {
							$newfldvalue =  explode('   ', $fldvalue);
							if (count($newfldvalue) == 2) {
								$sp_all_rows[$row][$fldname] = $newfldvalue[1];
							}
						}
					} else {
						if (empty($fldvalue) || $fldvalue == '0') {
							$message =  getTranslatedString($value['label'], $sp_module).' '.getTranslatedString('LBL_EMPTY_REQUIRED_FIELD');
							break;
						} else {
							$newfldvalue =  explode('   ', $fldvalue);
							if (count($newfldvalue) == 2) {
								$sp_all_rows[$row][$fldname] = $newfldvalue[1];
							}
						}
					}
				} elseif ($value['uitype'] == 53 || $value['uitype'] == 77 || $value['uitype'] == 101) { #for assigned user
					$fldname = $value['name'];
					$fldvalue = $sp_all_rows[$row][$fldname];
					$newfldvalue =  explode('   ', $fldvalue);
					if (count($newfldvalue) == 2) {
						$sp_all_rows[$row][$fldname] = $newfldvalue[1];
					} else {
						$sp_all_rows[$row][$fldname] = vtws_getEntityId('Users').'x'.$user->id;
					}
				}
			}
			$sp_all_rows[$row]['id'] = $sp_all_rows[$row][$sp_column_headers[0]];
			unset($sp_all_rows[$row]['']);
			$record = vtws_update($sp_all_rows[$row], $user);
			if (!empty($record['id'])) {
				$message = getTranslatedString('SUCCESS_SAVE', 'Spreadsheets');
			} else {
				$message = getTranslatedString('FAIL_SAVE', 'Spreadsheets');
			}
		}
	}
	VTWS_PreserveGlobal::flush();
	return $message;
}

function updateCRMIDColumnEtherCalc($spreedsheeturl, $command) {
	$ch = curl_init();
	curl_setopt($ch, CURLOPT_URL, $spreedsheeturl);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
	curl_setopt($ch, CURLOPT_HEADER, false);
	curl_setopt($ch, CURLOPT_POST, true);
	curl_setopt($ch, CURLOPT_POSTFIELDS, $command);
	curl_setopt($ch, CURLOPT_HTTPHEADER, array("Content-Type: application/json"));
	$content = curl_exec($ch);
	$errors = curl_error($ch);
	$response = curl_getinfo($ch, CURLINFO_HTTP_CODE);
	return ($response == 202);
}
?>
