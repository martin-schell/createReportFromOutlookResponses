# Description

⚠️This script was written and testet in 2021 with Python 3.9 or later.

Python script for creating a report from Outlook responses with the numbers of positiv, negative and tentative responses.

## What is it about?

To inform stakeholders, it could be necessary to invite them to information events.
Depending on the event, a very large number of responses may arrive in the mail inbox.

In order to be able to tell the organizer of an event how many employees can be estimated, the received mails has to be evaluated.

This programm exports the following information from meeting responses in Microsoft Outlook into an Excel file:

* first and last name as well as id of employee
* name and date of an event
* kind of response (positive, negative or tentative)

New data will be appended to table within the Excel file.

## Features

* A specific post office (PO) box, which is connected to the current account, will be opened.
* A specific folder within the Opening PO will be opened.
* The above mentioned information will be extracted from meeting responses within the specific folder and passed into a dataframe.
* The dataframe will be exported to an Excel file. 
  * If no file exists, a new file will be created in the folder of the programm.
  * If a report already exists, new participiant data will be appended to the table with existing data.

## External libraries

* [openpyxl] as engine for creating / adapting workbooks
* [pandas] for manipulate data
* [pywin32] as interface to the object model of Outlook

Dependencies are not listed.

[pandas]: https://pypi.org/project/pandas/
[pywin32]: https://pypi.org/project/pywin32/
[openpyxl]: https://pypi.org/project/openpyxl/
