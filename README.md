Developer Tab Functionality in Microsoft Excel 2016
====================================================

## Set Up

* Microsoft Excel 2016 is required.

* The Developer Tab must be activated. To do this, go to File, open options, click on "Customize Ribbon", and then check the box "Developer". 

* Note that macros must be enabled for this spreadsheet to work.

##Explanation

The Excel Spreadsheet has two tabs:

1. `VBA Code Basics` shows how VBA can be used to create Message Boxes and Input Boxes.

This sheet consists of two buttons with Sub Procedures assigned. We can assign Sub Procedures to buttons by right-clicking the button, and then selecting "Assign Macro". The Sub Procedure of interest can then be selected.

This sheet also has a function `=max_fraction()`. This can be used because we created a VBA Function.

To examine the VBA code behind this sheet either use hotkey `ALT + F11` or select "Visual Basic" in the Developer Tab.

2. `Without VBA` shows how the Developer Tab can be used without VBA code.

Here we apply a Shewart Chart approach to the monthly % change in FTSE 250 (sourced from [Yahoo Finance](https://uk.finance.yahoo.com/q/hp?s=%5EFTMC&b=31&a=11&c=1985&e=17&d=07&f=2016&g=m))

3. `Raw Data` is the data behind the visualisation in `Without VBA`


##Author

* [Andrew Burnie](https://github.com/Andrew47/)