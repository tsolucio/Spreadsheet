# coreBOS-EtherCalc Spreadsheet Integration

This module is part of the Spreadsheet editor integration between coreBOS and EtherCalc. It holds the information about which module and fields are being edited in each spreadsheet.

## Install and configure

First, you have to launch ethercalc using docker and our image which has a small modification to include the save to corebos button on the spreadsheet. You can find the docker-compose file in the ethercalc directory

Then install this module and set the **EtherCalc_URL** global variable to the URL of your EtherCalc install.

The module holds records that represent spreadsheets in ethercalc

Those spreadsheets are fields of a module and a certain set of records.

- The fields are represented by a [Field Set business map](https://corebos.com/documentation/doku.php?id=en:adminmanual:businessmappings:field_set)
- The set of records are defined by the set of conditions that are in the filter or question related to the sheet. If both are selected we will use the filter and ignore the question.

The idea is that when you click on the "open Sheet" link we get the set of columns from the Map, we get the conditions to filter the records from the Filter or Question, in other words, we get columns and rows to edit and we send that to ethercalc

Ethercalc returns a name which we save in the module (ethercalc field)

The next time you click on the Open Sheet link we see the name and use the same sheet

Once the sheet is open, the user can edit and change whatever they need and click on the corebos button. That button connects to coreBOS (login web service) and executes a method which gets the information from the spreadsheet and updates the coreBOS records.

You can add this Business Action to any module's List View to get a button that will create spreadsheets directly from the currently selected filter.

## EtherCalc changes

**NOTE:** These changes are already applied in the corebos/ethercalc docker image, they are only required if you start from another ethercalc image.

- Add the button:
  - edit static/ethercalc.js
  - search for `button_undo`
  - add in front `&nbsp;<img id="%id.button_corebos" src="%img.corebos.png" style="vertical-align:bottom;"> &nbsp;<img src="%img.divider1.png" style="vertical-align:bottom;">&nbsp;`
- Load corebos script and web service library
  - edit index.html and add at the end, inside the body `<script src="./static/WSClientp.js"></script><script src="./static/corebos.js"></script>`
  - copy these two scripts into the static directory
- copy the corebos image to the images directory with the name `sc_corebos.png`

## Information

- [EtherCalc API](https://ethercalc.docs.apiary.io)
- [Field Set Mapping](https://corebos.com/documentation/doku.php?id=en:adminmanual:businessmappings:field_set)

## Credits

Thanks Eri and Lorida for the first version of this integration and for laying out the way forward.
