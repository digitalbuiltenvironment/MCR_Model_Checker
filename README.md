# MCR Model Checker

The MCR model checker is a Python code that would help IDD managers check for the compliance of Revit files against JTC's Model Content Requirement (MCR) requirements. This model checker serves to streamline and optimise the JTC MCR checking process. The code runs on the Revitpythonshell Add-In in Revit. 

## How does the code work

1. Iterate through a list of category codes (E.g. OST_StructuralColumns , OST_Doors)
2. Extract a list of RevitElementIDs for given category code
3. Filter out the required elements using Type Comments or Family and Type Name (E.g Pumps)
4. Extract all parameter names associated with the element
5. Extract the correct parameter names required to be checked from Plannerly excel export
6. Compare the two lists and return error messages
7. Repeat the step for all given category codes
8. Generate an error report for all elements and export report into an Excel File

## Functionalities

1. Check if parameters exist
2. Check if parameter names are correct
3. Check if parameters are of the correct type
4. Check if Family and Type name is correct
5. Derivation of MCR code from Family and Type name
6. Error report exported as an Excel File

## Limitations

1. Currently, the code can only check 6 categories (ProjectInformation, TitleBlocks, Doors, Pump, GeneratingSet, StructuralColumns)
2. Unable to check Revit element if both Type Comments (MCR code) and Family and Type name are wrong
3. Identification and correction of parameter names will not work with typo errors and partial inputs
   - E.g. “OerallHeight” will not be identified as “OverallHeight”
4. For inputs requirements which are ‘Any text’ and ‘Any number’
   - Check if parameter input is empty 
   - Check if parameter input is ‘0’, ‘N.A.’ , ‘-’
5. A list of category codes to check is still needed (Hardcoded)
6. File path finder function still needs to be manually edited when adding more categories to check OR when code is given to another person to test (hardcoded)

## Future Works
1. Extension of code to all categories of Revit elements
2. Functionality to skip checking files that have already been checked before
3. Online database of Plannerly export files or ability to access Plannerly data directly 
   - File paths do not need to be changed every time a new person wants to use the code
4. Designated Pyrevit button that can run the code with a single click 

## When to use the MCR Model Checker
1. When Revit files are too big to be uploaded to Plannerly (>300mb)
2. Confidential and secret projects that cannot be uploaded to Plannerly and offline checks are required  

## Onboarding

1. Ensure that you have a working Revit license. 
2. 

 
