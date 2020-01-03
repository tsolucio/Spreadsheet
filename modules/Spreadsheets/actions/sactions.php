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
	public $commandtosend = '';
	public $fieldListData= array();
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

	private function createSpreadsheet($record, $ecUrl, $selected_record_ids_from_listview = '') {
		global $adb, $current_language, $default_language, $current_user;
		$nonSupportedFields = array('campaignrelstatus');
		$ch = curl_init();
		curl_setopt($ch, CURLOPT_URL, $ecUrl.'_');
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($ch, CURLOPT_HEADER, false);
		curl_setopt($ch, CURLOPT_POST, true);
		// Get columns to update from map
		$user_current_language = $current_language;
		$current_language = $default_language;
		$rs = $adb->pquery('select map, spmodule,question, filter from vtiger_spreadsheets where spreadsheetsid=?', array($record));
		$mapid = $adb->query_result($rs, 0, 'map');
		$sp_module = $adb->query_result($rs, 0, 'spmodule');
		$cbquestionid = $adb->query_result($rs, 0, 'question');
		$filter = $adb->query_result($rs, 0, 'filter');
		// Create Module Fieldname, label pair Array for Easy Translation
		include_once 'include/Webservices/DescribeObject.php';
		$moduleinfo = vtws_describe($sp_module, $current_user);
		$module_fieldname_label_key_pairs = array();
		foreach ($moduleinfo['fields'] as $value) {
			$module_fieldname_label_key_pairs[$value['name']] = $value['label'];
		}

		$col_array =  array();
		$cols = '';

		if (!empty($mapid) && !empty($sp_module)) {
			$cbMap = cbMap::getMapByID($mapid);
			$spreadsheet_mod_ins = CRMEntity::getInstance($sp_module);
			$col_array = $cbMap->Mapping($spreadsheet_mod_ins->column_fields, $col_array);
		}

		$trans_col_array = array();
		$untrans_col_array = array();
		$vtlibmod = Vtiger_Module::getInstance($sp_module);
		foreach ($col_array as $fieldname => $value) {
			$fld = Vtiger_Field::getInstance($fieldname, $vtlibmod);
			if ($fld) {
				$untrans_col_array[] = $fld->column;
			} else {
				$untrans_col_array[] = $fieldname;
			}
			if (empty($module_fieldname_label_key_pairs[$fieldname])) {
				$trans_col_array[] = getTranslatedString($fieldname, $sp_module);
			} else {
				$trans_col_array[] = getTranslatedString($module_fieldname_label_key_pairs[$fieldname], $sp_module);
			}
		}

		$cols = implode(',', $trans_col_array)."\n";
		$ethercalc_commands = array();
		// Filter the Record Set By Using Filters or cbQuestion or Selected Record from ListView
		if (!empty($selected_record_ids_from_listview)) {
			$qg = new QueryGenerator($sp_module, $current_user);
			$qg->setFields(array('*'));
			$qg->addCondition('id', explode(',', $selected_record_ids_from_listview), 'i');
			$result = $adb->query($qg->getQuery());
			$columnindex = 1;
			$rowindex = 1;
			if ($result) {
				while ($row = $adb->fetch_array($result)) {
					$columnindex = 1;
					$rowindex++;
					for ($field_index = 0; $field_index < count($untrans_col_array); $field_index++) {
						if (in_array($untrans_col_array[$field_index], $nonSupportedFields)) {
							continue;
						}
						if ($field_index == 0) {
							$crmid = vtws_getEntityId($sp_module)."x".$row[$untrans_col_array[$field_index]];
							$cols = $cols.$crmid.",";
							$wsid = $crmid;
						} else {
							$command_to_set_cell_value = $this->generateEtherCalcSheetCommand(
								$sp_module,
								$untrans_col_array[$field_index],
								$row[$untrans_col_array[$field_index]],
								$columnindex,
								$rowindex,
								$wsid
							);
							if (!empty($command_to_set_cell_value)) {
								$ethercalc_commands[] = $command_to_set_cell_value;
							}
							$cols = $cols.$row[$untrans_col_array[$field_index]].($field_index == count($untrans_col_array)? "" : ",");
						}
						$columnindex++;
					}
					$cols = $cols."\n";
				}
			}
		} elseif (!empty($filter)) {
			$queryGenerator = new QueryGenerator($sp_module, $current_user);
			$queryGenerator->initForCustomViewById($filter);
			$result = $adb->query($queryGenerator->getQuery());
			$columnindex = 1;
			$rowindex = 1;
			if ($result) {
				while ($row = $adb->fetch_array($result)) {
					$columnindex = 1;
					$rowindex++;
					for ($field_index = 0; $field_index < count($untrans_col_array); $field_index++) {
						if ($field_index == 0) {
							$crmid = vtws_getEntityId($sp_module)."x".$row[$untrans_col_array[$field_index]];
							$cols = $cols.$crmid.",";
							$wsid = $crmid;
						} else {
							$command_to_set_cell_value = $this->generateEtherCalcSheetCommand(
								$sp_module,
								$untrans_col_array[$field_index],
								$row[$untrans_col_array[$field_index]],
								$columnindex,
								$rowindex,
								$wsid
							);

							if (!empty($command_to_set_cell_value)) {
								$ethercalc_commands[] = $command_to_set_cell_value;
							}
							$cols = $cols.$row[$untrans_col_array[$field_index]].($field_index == count($untrans_col_array)? "" : ",");
						}
						$columnindex++;
					}
					$cols = $cols."\n";
				}
			}
		} elseif (!empty($cbquestionid)) {
			include_once 'modules/cbQuestion/cbQuestion.php';
			$query = cbQuestion::getSQL($cbquestionid);
			$result = $adb->query($query);
			$columnindex = 1;
			$rowindex = 1;
			if ($result) {
				while ($row = $adb->fetch_array($result)) {
					$columnindex = 1;
					$rowindex++;
					$wsid= '';
					for ($field_index = 0; $field_index < count($untrans_col_array); $field_index++) {
						if ($field_index == 0) {
							$crmid = vtws_getEntityId($sp_module)."x".$row[$untrans_col_array[$field_index]];
							$cols = $cols.$crmid.",";
							$wsid = $crmid;
						} else {
							$command_to_set_cell_value = $this->generateEtherCalcSheetCommand(
								$sp_module,
								$untrans_col_array[$field_index],
								$row[$untrans_col_array[$field_index]],
								$columnindex,
								$rowindex,
								$wsid
							);

							if (!empty($command_to_set_cell_value)) {
								$ethercalc_commands[] = $command_to_set_cell_value;
							}
							$cols = $cols.$row[$untrans_col_array[$field_index]].($field_index == count($untrans_col_array)? "" : ",");
						}
						$columnindex++;
					}
					$cols = $cols."\n";
				}
			}
		}
		curl_setopt($ch, CURLOPT_POSTFIELDS, $cols);
		curl_setopt($ch, CURLOPT_HTTPHEADER, array('Content-Type: text/csv'));
		$ecname = curl_exec($ch);
		curl_exec($ch);
		$adb->pquery('update vtiger_spreadsheets set ethercalcid=? where spreadsheetsid=?', array($ecname, $record));

		if (count($ethercalc_commands) > 0) {
			$this->commandtosend = '{"command":["'.implode('","', $ethercalc_commands).'"]}';
		}
		$current_language = $user_current_language;
		return $ecname;
	}


	private function getAutocompleteValue($fieldname, $module) {
		global $current_user, $log, $adb;
		if (array_key_exists($fieldname, $this->fieldListData)) {
			return $this->fieldListData[$fieldname];
		} else {
			//  Go to Database and Select Value(Module which was Related)
			include_once 'include/Webservices/DescribeObject.php';
			$moduleinfo = vtws_describe($module, $current_user);
			$linkListField = ""; // from rel module
			$relmodule = "";
			$relmoduletable = "";
			foreach ($moduleinfo['fields'] as $key => $value) {
				if ($value['name'] == $fieldname) {
					$relmodule = $value['type']['refersTo'][0];
					if (!empty($relmodule)) {
						$moduleInstance = CRMEntity::getInstance($relmodule);
						$id_field = $moduleInstance->customFieldTable[1];
						$relmoduletable = $moduleInstance->table_name;
						$linkListField = $moduleInstance->list_link_field;
						if (!empty($linkListField) && !empty($id_field) && !empty($relmoduletable)) {
							$query = "SELECT ".$relmoduletable.".".$linkListField.",".$relmoduletable.".".$id_field." FROM ".$relmoduletable." 
								INNER JOIN vtiger_crmentity ON ".$relmoduletable.".".$id_field." = vtiger_crmentity.crmid LEFT JOIN vtiger_users ON 
								vtiger_crmentity.smownerid = vtiger_users.id LEFT JOIN vtiger_groups 
								ON vtiger_crmentity.smownerid = vtiger_groups.groupid WHERE vtiger_crmentity.deleted=0 AND  ".$relmoduletable.".".$id_field." > 0";
							$result = $adb->query($query);
							if ($result) {
								$temp_arr = array();
								while ($row = $adb->fetch_array($result)) {
									$temp_arr[] = $row[$linkListField]."   ".vtws_getEntityId($relmodule)."x".$row[$id_field];
								}
								$this->fieldListData[$fieldname] = $temp_arr;
								return $temp_arr;
							}
						}
					}
				}
			}
		}
		return "";
	}

	public function sendCommandToEtherCalc() {
		global $log;
		$spurl = isset($_REQUEST['spurl']) ? vtlib_purify($_REQUEST['spurl']) : '';
		$command = isset($_REQUEST['command']) ? vtlib_purify($_REQUEST['command']) : '';
		$ch = curl_init();
		curl_setopt($ch, CURLOPT_URL, $spurl);
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($ch, CURLOPT_HEADER, false);
		curl_setopt($ch, CURLOPT_POST, true);
		curl_setopt($ch, CURLOPT_POSTFIELDS, $command);
		curl_setopt($ch, CURLOPT_HTTPHEADER, array("Content-Type: application/json"));
		$response = curl_exec($ch);
		echo json_encode($response);
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

	public function generateEtherCalcSheetCommand($module, $fieldname, $fieldvalue, $colindex, $rwindex, $wsid) {
		global $adb, $current_user;
		include_once 'include/Webservices/DescribeObject.php';
		require_once 'include/Webservices/Retrieve.php';
		$value = '';
		$moduleinfo = vtws_describe($module, $current_user);
		foreach ($moduleinfo['fields'] as $finfo) {
			if ($finfo['name'] == $fieldname) {
				if (!empty($finfo['uitype']) && (in_array($finfo['uitype'], array(53, 15, 77, 101)) || $fieldname == 'salutationtype')) {
					if ($fieldname == 'assigned_user_id') {
						$picklistValues = $finfo['type']['assignto']['users']['options'];
						$defaultValue = $current_user->user_name.'   '.vtws_getEntityId('Users').'x'.$current_user->id;
					} elseif ($fieldname == 'salutationtype') {
						require_once 'modules/PickList/PickListUtils.php';
						$roleid=$current_user->roleid;
						$picklistValues = array_values(getAssignedPicklistValues('salutationtype', $roleid, $adb));
						$defaultValue = $picklistValues[0];
					} else {
						$picklistValues = $finfo['type']['picklistValues'];
						$defaultValue = $finfo['type']['defaultValue'];
					}
					$picklist = array();
					if (!empty($picklistValues) && count($picklistValues) > 0) {
						foreach ($picklistValues as $key => $plvalue) {
							if ($fieldname == 'assigned_user_id') {
								$picklist[]= $plvalue['username'].'   '.$plvalue['userid'];
							} elseif ($fieldname == 'salutationtype') {
								$picklist[] = $plvalue;
							} else {
								$picklist[]= $plvalue['value'];
							}
						}
						if (count($picklist) > 0) {
							if (!empty($fieldvalue)) {
								$cell = $this->convertNumberToColumnHeaderLabel($colindex);
								$value = "set ".$cell.$rwindex." formula SELECT(\'".$fieldvalue."\',\'".implode(",", $picklist)."\')";
							} else {
								$cell = $this->convertNumberToColumnHeaderLabel($colindex);
								$value = "set ".$cell.$rwindex." formula SELECT(\'".$defaultValue."\',\'".implode(",", $picklist)."\')";
							}
						}
					}
				} elseif (!empty($finfo['uitype']) && $finfo['uitype']==56) {
						$chvalue = ($fieldvalue == 1) ? true : false;
						$cell = $this->convertNumberToColumnHeaderLabel($colindex);
						$value = "set ".$cell.$rwindex." formula CHECKBOX(\'".$chvalue."\')";
				} elseif (!empty($finfo['uitype']) && $finfo['uitype']==10) {
					$autocompletevalue = $this->getAutocompleteValue($fieldname, $module);
					$recordinfo = vtws_retrieve($wsid, $current_user);
					if (isset($recordinfo[$fieldname.'ename'])) {
						$fieldvalue = trim($recordinfo[$fieldname.'ename']['reference'].'   '.$recordinfo[$fieldname]);
					} else {
						$fieldvalue = '   '.$recordinfo[$fieldname];
					}
					if (count($autocompletevalue) > 0) {
						if (!empty($fieldvalue)) {
							$cell = $this->convertNumberToColumnHeaderLabel($colindex);
							$value = "set ".$cell.$rwindex." formula AUTOCOMPLETE(\'".$fieldvalue."\',\'".implode(",", $autocompletevalue)."\')";
						} else {
							$cell = $this->convertNumberToColumnHeaderLabel($colindex);
							$value = "set ".$cell.$rwindex." formula AUTOCOMPLETE(\'".array_keys($autocompletevalue)[0]."\',\'".implode(",", $autocompletevalue)."\')";
						}
					}
				}
				break;
			}
		}
		$value = stripcslashes($value);
		return $value;
	}

	public function convertNumberToColumnHeaderLabel($index) {
		$index -= 1;
		$cell = '';
		for ($index; $index >= 0; $index = intval($index / 26) - 1) {
			$cell = chr($index % 26 + 0x41) . $cell;
		}
		return $cell;
	}

	public function createFromModuleListView() {
		include_once 'include/Webservices/Create.php';
		global $log, $current_user,$adb, $site_URL, $currentModule;
		$module = isset($_REQUEST['sourcemodule']) ? vtlib_purify($_REQUEST['sourcemodule']) : '';
		$viewid = isset($_REQUEST['viewid']) ? vtlib_purify($_REQUEST['viewid']) : 0;
		$selected_record_ids_from_listview = isset($_REQUEST['allids']) ? vtlib_purify($_REQUEST['allids']) : '';
		// Creating Map
		$moduleInstance = CRMEntity::getInstance($module);
		$primary_key_field = $moduleInstance->customFieldTable[1];
		$column_fields_list = $moduleInstance->column_fields;
		$mapname = $module."2EtherCalc";
		// Before Create check if Map name aready Exit If yes Get the Map Id and Use it
		$mapid = cbMap::getMapIdByName($mapname);
		$mapcrmid = vtws_getEntityId('cbMap').'x'.cbMap::getMapIdByName($mapname);
		if (empty($mapid) && $mapid == 0) {
			$maptype = "Mapping";
			$fields_contents = '';
			$module_table = $moduleInstance->table_name;
			// Add the Primary Key Coulum First
			$fields_contents = $fields_contents."<field>
				<fieldname>".$primary_key_field."</fieldname>
				<Orgfields>
					<Orgfield>
						<OrgfieldName>".$primary_key_field."</OrgfieldName>
						<OrgfieldID>FIELD</OrgfieldID>
					</Orgfield>
				</Orgfields>
			</field>";

			foreach ($column_fields_list as $key => $value) {
				// Check if field is UItype 3, 4, 6, 8, 12, 25, 30, 31, 32, 52, 53, 69, 69m, 70
				// If its Skip it
				$result_moduletable = $adb->pquery(
					'select * from vtiger_field where uitype in (3,4,6,8,12,25,30,31,32,52,69,70) and tablename=? and fieldname=?',
					array($module_table, $key)
				);
				$result_crmentitytable = $adb->pquery(
					'select * from vtiger_field where uitype in (52,70) and tablename=? and fieldname=?',
					array('vtiger_crmentity', $key)
				);
				if ($result_moduletable && $adb->num_rows($result_moduletable) == 0 && $result_crmentitytable && $adb->num_rows($result_crmentitytable) == 0) {
					if ($key != $primary_key_field) {
						$fields_contents = $fields_contents."<field>
						<fieldname>".$key."</fieldname>
						<Orgfields>
							<Orgfield>
								<OrgfieldName>".$key."</OrgfieldName>
								<OrgfieldID>FIELD</OrgfieldID>
							</Orgfield>
						</Orgfields>
					</field>";
					}
					$fields_contents = $fields_contents."\n";
				}
			}
			$mapcontent = "<map>
			<originmodule>
				<originname>".$module."</originname>
			</originmodule>
			<targetmodule>
				<targetname>EtherCalc</targetname>
			</targetmodule>
			<fields>".
				$fields_contents
			."</fields>
		</map>";
			$maprecord = array(
			'mapname' => $mapname,
			'maptype' => $maptype,
			'targetname' => $module,
			'content' => $mapcontent,
			'assigned_user_id' => vtws_getEntityId('Users').'x'.$current_user->id
			);

			$created_map_object = vtws_create('cbMap', $maprecord, $current_user);
			$mapcrmid = $created_map_object['id'];
		}

		if (!empty($mapcrmid)) {
			// Create Spread Sheet
			$spreadsheet_record = array(
			'spreadsheetsname' => $module.'Spreadsheet',
			'spmodule' => $module,
			'map' => $mapcrmid,
			'filter' => $viewid,
			'assigned_user_id' => vtws_getEntityId('Users').'x'.$current_user->id
			);
			$created_spreadsheet_object = vtws_create('Spreadsheets', $spreadsheet_record, $current_user);
			if (!empty($created_spreadsheet_object['id'])) {
				$record = explode('x', $created_spreadsheet_object['id'])[1];
				$ecUrl = $this->checkEtherCalcURL();
				$ecid = $this->createSpreadsheet($record, $ecUrl, $selected_record_ids_from_listview);
				$ecid = trim($ecid, '/');
				$rdo = array();
				$rdo['status'] = 'OK';
				$rdo['msg'] = 'Open';
				$rdo['sendcommandurl'] = $ecUrl.'_/'.$ecid;
				$rdo['command'] = $this->commandtosend;
				$rdo['notify'] = $ecUrl . $ecid . '?usr=' . $current_user->user_name . '&pwd=' . $current_user->accesskey
					. '&url=' . urlencode($site_URL) . '&mtd=updateFromEthercalc&sid='.$record;
				echo json_encode($rdo);
			}
		}
	}
}
?>