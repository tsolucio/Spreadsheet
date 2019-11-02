# coreBOS-EtherCalc Spreadsheet Integration

This module is part of the Spreadsheet editor integration between coreBOS and EtherCalc. It holds the information about which module and fields are being edited in each spreadsheet.

## Install and configure

First you have to launch ethercalc using docker and our image which has small modification to include the save to corebos button on the spread sheet.

Then install this module and set the **EtherCalc_URL** global variable to the URL of your EtherCalc install.


## EtherCalc changes

- Add the button:
 - edit static/ethercalc.js
 - search for `button_undo`
 - add in front `&nbsp;<img id="%id.button_corebos" src="%img.corebos.png" style="vertical-align:bottom;"> `
- Load corebos script
 - edit index.html and add at the end, inside the body `<script src="./static/corebos.js"></script>`
- copy the corebos image to the images directory with the name `sc_corebos.png`
