.. _bas:

Importing the EGAMMA VBA Code into an Excel Workbook
====================================================
This guide outlines the steps to integrate the EGAMMA VBA code into an individual Excel workbook, which allows you to utilize EGAMMA's capabilities within that specific file. If you wish to extend EGAMMA's functionalities to all workbooks on your system, consider installing the EGAMMA.xlam plugin instead (see :ref:`xlam` for installation guidance).

To begin using the EGAMMA VBA code within a single workbook, you will need to download the EGAMMA.bas file. This file is a VBA module containing the necessary code to be imported into the VBA editor of Excel. Doing so enables you to tailor and engage with the EGAMMA features directly in your chosen Excel workbook.

Please follow the detailed steps below to import the EGAMMA.bas file into your workbook:


Step 1: Download the EGAMMA.bas File
-------------------------------------
- Visit the [EGAMMA-VBA GitHub Releases page](https://github.com/FrodeDrevland/EGAMMA-VBA/releases/latest) to access the latest version of the EGAMMA.bas file.
- Click on the link for the EGAMMA.bas file to initiate the download.
- Save the file to a familiar directory on your computer for easy retrieval.

Step 2: Enable the Developer Tab in Excel
-----------------------------------------
- For Excel 2010 and newer:
  - Navigate to the ``File`` tab, select ``Options``, and then choose ``Customize Ribbon``.
  - In the customization area, ensure the ``Developer`` checkbox is selected and confirm with ``OK``.

- For Excel 2007:
  - Click the Microsoft Office Button, select ``Excel Options``, and within the ``Popular`` category, activate the ``Show Developer tab in the Ribbon`` option.

Step 3: Access the VBA Editor
-----------------------------
- From the Developer tab, initiate the ``Visual Basic`` command to open the VBA editor.

Step 4: Import the .bas File
-----------------------------
- Within the VBA editor, right-click on the ``VBAProject`` pertaining to your workbook.
- Choose ``Import File...`` and navigate to the directory where the EGAMMA.bas file is located.
- Select the file and opt for ``Open`` to import it into your project.

Step 5: Save Your Workbook with Macros
--------------------------------------
- Return to Excel and save your workbook by selecting ``Save``.
- If your workbook is in .xlsx format, you'll need to switch to the .xlsm format to preserve the VBA code.

Step 6: Enable Macros if Necessary
----------------------------------
- To execute the VBA code, ensure that macros are enabled by visiting the ``File`` tab, selecting ``Options``, followed by ``Trust Center``, and then ``Trust Center Settings...``.
- In the Macro Settings section, choose a suitable option to enable macros, keeping security in mind.

Troubleshooting Tips
--------------------
- If the Developer tab or VBA editor is not visible, they may need to be enabled in your Excel settings.
- Should you face any issues with the code execution, verify the compatibility of the code with your Excel version and check for required references.

Always back up your data before importing a .bas file, particularly if it includes extensive macros or code that could modify your data significantly.
