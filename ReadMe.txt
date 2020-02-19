PoiDriver is a wrapper/driver for Apache POI that is geared towards testing needs.
The purpose of PoiDriver is to shorten coding needs and keep test files clean.
it is currently incomplete, and therefore still in progress.



---------------------------------------------------------------------------------------------------------------------------------------
PoiDriver.java
---------------------------------------------------------------------------------------------------------------------------------------
The main wrapper file, it holds all action code.
Although there is still plenty left to do, it already contains a lot of functionality.

Including (but not limited to):
	- Retrieve row data by selecting column header and a value (under selected column header) to search for.
	- Mark a cell as Passed/Failed/Skipped
		- Individually
		- Temporarily save results then apply them all at once later
	- Get/Set cell data by cell reference/address (ie: "A5", "B12", "C17", etc)
	- Attempts to automatically format cell based on value.
	- Setting cell style made easier (via CellStyleHelper class)
	- Add/Remove sheet by name or index
	- Add/Remove column

Adding soon:
	- Add/Remove row
	- Mark test case as enabled/disabled in excel then skip disabled tests in TestNG



---------------------------------------------------------------------------------------------------------------------------------------
CellStyleHelper.java
---------------------------------------------------------------------------------------------------------------------------------------
A wrapper class to create shortcuts to commonly used functionality.



---------------------------------------------------------------------------------------------------------------------------------------
To Test
---------------------------------------------------------------------------------------------------------------------------------------
1)  Run "PoiDriverTest.java" as a java application
		Open .xlsx file to view changes made, make sure to close the file afterwards
2)  Run "PredefinedColorsTest.java" as a java application
		Open .xlsx file, see the "Sample" sheet/tab, make sure to close the file afterwards
3)	Run "AddColumnTest.java" as a java application
		Open .xlsx file, see the "Demo" sheet/tab, make sure to close the file afterwards
4)	Run "RemoveColumnTest.java" as a java application
		Open .xlsx file, see the "Demo" sheet/tab, make sure to close the file afterwards
5)  Run "PoiDriverTest.java" again
		To see an example of why it is recommended to not select items by their index.



