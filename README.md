# excel-remove-workbook-protections

Removes all worksheet and workbook protections from Excel files without needing to know the passwort.
This does not work with **encrypted** Excel files, just protected ones.

Simply import the .bas file in Excel VBA using the Development Console found via **Alt + F11** and run this Sub.

It is prompting for a file to be selected. After the routine is done, there is a new file created in the original folder called *Original-Name_unprotected_YYYYMMDD.xl..*
