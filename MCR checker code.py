#explanation:
#code is able to continue writing in the same file if the existing file exist
#code is able to create and write a new excel file (with headers) if the file does not exist
#can check project information
#automatically bold and set col width
#check input units
#added in non-hardcoded family and type name checker
#edited the name_filter() such that it will search for any 'AR' in the name
#added in a non_hardcoded mcr_from_ft()

#Things to add:
#1. expand the category list to more than 4 categories -> check_family_name, mcr_from_ft....
#2. add in non-hardcoded family and type name checker
#3. add in non-hardcoded mcr_from_ft 
#4. use online plannerly export instead of local files


#steps to use excel exporter:
#1. create and name new excel file
#2. copy file location and input into the code
#3.

#imports 
from Autodesk.Revit.DB import *
from System.Collections.Generic import List

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference("System.Windows.Forms")
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB import ElementId, FilteredElementCollector
from Autodesk.Revit.DB import ViewSheet, Viewport
from Autodesk.Revit.UI.Selection import ObjectType

import os.path
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
import System.Runtime.InteropServices

import re

import System
import System.Enum as Enum
from System.Windows.Forms import SendKeys


# Variables
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

#0. category code (key): mcr code(value) dictionary

category_mcr_dict = {
    'OST_Doors':['MCR03-01-01-01','MCR03-01-02-01', 'MCR03-02-01-01'],
    "OST_MechanicalEquipment":['MCR23-21-01-01'],
    "OST_ElectricalEquipment":["MCR23-14-01-01"],
    "OST_StructuralColumns":["MCR20-02-01-01", "MCR20-02-01-02"]
}


#1. extract value of type parameter given revitelementID and type parameter name (5ms)

def get_parameter_value(element_id, parameter_name):
    doc = __revit__.ActiveUIDocument.Document
    element = doc.GetElement(ElementId(element_id))
    
    if element:
        element_type_id = element.GetTypeId()
        
        if element_type_id == ElementId.InvalidElementId:
            return "Invalid Element Type ID"
        
        element_type = doc.GetElement(element_type_id)
        
        parameter = element_type.LookupParameter(parameter_name) #search in type parameters
        if not parameter:
            parameter = element.LookupParameter(parameter_name) #search in instance parameters

        if parameter:
            storage_type = parameter.StorageType
            if storage_type == StorageType.String:
                return parameter.AsString()
            elif storage_type == StorageType.Integer:
                return parameter.AsInteger()
            elif storage_type == StorageType.Double:
                return parameter.AsDouble()
            elif storage_type == StorageType.ElementId:
                return parameter.AsElementId()
            elif storage_type == StorageType.ElementIdArray:
                return parameter.AsElementIdArray()
            elif storage_type == StorageType.IntegerArray:
                return parameter.AsIntegerArray()
            elif storage_type == StorageType.StringArray:
                return parameter.AsStringArray()
            elif storage_type == StorageType.DoubleArray:
                return parameter.AsDoubleArray()
    
    return None

#2. noneType element filter
def noneType_filter(element_id):
    element = doc.GetElement(ElementId(element_id))
    
    element_type_id = element.GetTypeId()

    if element_type_id == ElementId.InvalidElementId:
        return "invalid element"
    else:
        return "valid element"

#3. get element ids 
#element family is string -> 'OST_MechanicalEquipment'
def get_element_ids(excel, file_path, sheet_name,element_family):
    doc = __revit__.ActiveUIDocument.Document
    
    # Get the built-in category for the specified element family
    category = Enum.Parse(BuiltInCategory, element_family)
    
    # Create a filtered element collector for the specified category
    collector = FilteredElementCollector(doc).OfCategory(category)
 
    # Create an empty list to store the element IDs
    element_ids = []
    
    # Use list comprehension to extract the element IDs
    for element in collector:
        element_id = int(element.Id.ToString())
        type_comments_input = get_parameter_value(element_id, 'Type Comments')
        if noneType_filter(element_id) == 'valid element':
            mcr_list = ['MCR03-01-01-01','MCR03-01-02-01', 'MCR03-02-01-01', 'MCR23-21-01-01' , 'MCR23-14-01-01' , "MCR20-02-01-01", "MCR20-02-01-02"]
            if type_comments_input in mcr_list or mcr_from_ft(excel, file_path, sheet_name, element_id) in mcr_list:
                element_ids.append(element_id)
        
    return element_ids

#4. Get all parameter names given RevitElementID (2ms)

def get_parameter_names(doc, element_id):
    element = doc.GetElement(ElementId(element_id))
    
    
    # Get the parameter names for the element
    element_parameter_names = [param.Definition.Name for param in element.Parameters]
    
    
    # Get the parameter names for the element type
    element_type = doc.GetElement(element.GetTypeId())
    
    type_parameter_names = [param.Definition.Name for param in element_type.Parameters] #problem
    
    
    # Get the built-in parameter names for the pump element
    built_in_parameter_names = [param.Definition.Name for param in element.GetOrderedParameters()]
    
    
    # Combine all parameter names
    parameter_names = element_parameter_names + type_parameter_names + built_in_parameter_names
    
    return parameter_names


# 5. get correct parameter names given the keyword (e.g MCR03-01-01) 

def find_row_and_column(excel,file_path, sheet_name, element_type):
    #excel = Excel.ApplicationClass()
    workbook = excel.Workbooks.Open(file_path)
    worksheet = workbook.Sheets[sheet_name]
    
    last_row = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
    
    if element_type == "Invalid Element Type ID":
        return 'MCR code does not exist'
    
    for row in range(1, last_row + 1):
        cell = worksheet.Cells[row, 5]  # Column E is represented by index 5
        
        if cell.Value2 == element_type:
            values = []
            current_row = row + 1
            
            while True:
                value_cell = worksheet.Cells[current_row, 24]  # Column X is represented by index 24
                
                if not value_cell.Value2:
                    break
                values.append(value_cell.Value2)
                current_row += 1
            #workbook.Close(False)
            #excel.Quit()
            return values  # Return values as a list
    return 'MCR code does not exist'
    
    
#6. compare two list of parameter string names to return list of missing parameters
def compare_parameters(first_list, second_list):
    missing_parameters = [parameter for parameter in second_list if parameter not in first_list]
    return missing_parameters

#7. get category code from element ID 
#element_id is a string & category code output is e.g. 'OST_MechanicalEquipment'
def get_category_code(element_id):
    doc = __revit__.ActiveUIDocument.Document

    # Convert the element ID from string to ElementId
    element_id = ElementId(int(element_id))

    # Retrieve the element from the document
    element = doc.GetElement(element_id)

    # Retrieve the category of the element
    category = element.Category

    # Retrieve the category code as a string
    category_code = 'OST_' + category.Name.replace(" ", "")

    return category_code
    
    
#8. Function to split the input by row and add them to a list
def split_input_to_list(input_text):
    lines = input_text.strip().split('\n')
    # Regular expression to remove the numbering at the start of each line
    return [re.sub(r'^\d+\.\s*', '', line.strip()) for line in lines]

#9. family name format finder
def family_name_format_finder(input_list):    
    for input in input_list:
        if 'Correct format' in input: # must be 'Correct format'
            correct_format_string = input
            input_list.remove(input)
    correct_format = correct_format_string.replace("Correct format:",'') # Main Category-Material-Subcategory:Differentiator 
    parts = correct_format.split(':')
    name_parts = parts[0].split('-')    #['Main Category','Material','Subcategory']
    name_parts.append(parts[-1])     #['Main Category','Material','Subcategory','Differentiator']
    
    component_list = []
    for components in name_parts:
        for input in input_list:
            if components in input: #find matching inputs (e.g. Main Category: DR) that contain keyword (e.g. Main Category)
                component_list.append(input)
    
    component_dict = {}
    for item in component_list:
        key, value = item.split(':') #e.g. Main Category: DR
        component_dict[key.strip()] = value.strip()       
    
    return correct_format , name_parts, component_dict


#7. family name checker (non-hard coded)
def check_family_name(element_id, element_type, category_mcr_dict): #can add in file_path and sheet_name in ()
    element = doc.GetElement(ElementId(element_id))
    family_name = element.Symbol.FamilyName
    type_name = element.Name
    combined_name = family_name + ":" + type_name
    ft_name_parts = combined_name.split(':')
    
    category_code = get_category_code(element_id)
    excel = Excel.ApplicationClass()
    file_path , sheet_name = file_path_finder(category_code)
    workbook = excel.Workbooks.Open(file_path)
    worksheet = workbook.Sheets[sheet_name]

    last_row = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
    
    for row in range(1, last_row + 1):
        cell = worksheet.Cells[row, 5]  # Column E is represented by index 5
        if cell.Value2 == element_type:
            description_cell = worksheet.Cells[row-1, 15]
            description_input = description_cell.Value2
            
    input_list = split_input_to_list(description_input)
       
    #note that Correct format cannot have space in between each component
    correct_format, name_parts, component_dict = family_name_format_finder(input_list) #e.g. name_parts = ['Main Category','Material','Subcategory','Differentiator']
    errors = []
    ft_name_format = ' Family and Type name should be in the format: ' + f'{correct_format}'
    
    if element_type in category_mcr_dict[category_code]:
        ft_name1 = ft_name_parts[0] #actual family name before ':' with no colon
        dash_ft_name1 = ft_name1.split('-')        
        ft_name2 = ft_name_parts[0] + ':' #actual family name before ':' with no colon
        dash_ft_name2 = ft_name2.split('-')
        
        if len(dash_ft_name1) >= len(name_parts)-1: #if number of actual name parts >= correct name parts
            for i in range(len(name_parts)-1):  # check all components except component after :
                name_component = name_parts[i]
                value  = component_dict[name_component]
                
                split_values = [value]
                if ',' in value:
                    split_values=[]
                    split_values = value.split(',')
                    
                if dash_ft_name1[i] not in split_values:
                    errors.append(f'{name_component}')
                
            if name_parts[-1] in component_dict: #check the component after :
                name_component = name_parts[-1]
                if component_dict[name_component] not in ft_name_parts[-1]:
                    errors.append(f'{name_component}')
            if ':' not in dash_ft_name2[-1]: #check for placement of :
                errors.append("Placement of ':'")
        else:
            errors.append("Family and Type name format")
    else:
        errors.append('MCR code')
        
    if errors == []:
        excel.Quit()
        return 'correct',combined_name,ft_name_format
    
    excel.Quit()
    return errors, combined_name, ft_name_format
        

#8. parameter corrector 
def parameter_corrector(missing_parameters, door_parameters):
    comments = []
    corrected_parameters = []
    correct_parameters = []
    
    missing_parameters_lower = [x.replace(" ", "").lower() for x in missing_parameters]
    
    for element in door_parameters:
        formatted_element = element.replace(" ", "").lower()

        if formatted_element in missing_parameters_lower:
            corresponding_element = missing_parameters[missing_parameters_lower.index(formatted_element)]

            if element != corresponding_element:
                if element not in corrected_parameters:
                    corrected_parameters.append(element)
                    correct_parameters.append(corresponding_element)
                    comments.append(f"'{element}' should be '{corresponding_element}' (case incorrect).")
    
    return comments,corrected_parameters, correct_parameters



#9. determine if input MCR code is correct and exist in the list of MCR codes in excel (1480ms)

def mcr_checker(excel,file_path, sheet_name, mcr_code):
    #excel = Excel.ApplicationClass()
    workbook = excel.Workbooks.Open(file_path)
    worksheet = workbook.Sheets[sheet_name]
    
    last_row = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row

    if mcr_code != None and "Invalid Element Type ID":    #check that type comments parameter input is valid and not none
        if mcr_code.startswith("MCR") and mcr_code.count("-") == 3: #check that MCR code is in the correct form            
            for row in range(1, last_row + 1):
                cell = worksheet.Cells[row, 5]  # Column E is represented by index 5               

                if cell.Value2 == mcr_code:
                    #workbook.Close(False)
                    #excel.Quit()
                    return 'correct'            
        #workbook.Close(False)
        #excel.Quit()
        return 'mcr code does not exist'
    else:
        #workbook.Close(False)
        #excel.Quit()
        return 'mcr code is invalid'


#10. obtain corresponding MCR code from Family and Type Name [need edit]

def mcr_from_ft(excel, file_path, sheet_name, element_id):
    workbook = excel.Workbooks.Open(file_path)
    worksheet = workbook.Sheets[sheet_name]    
    last_row = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
    
    element = doc.GetElement(ElementId(element_id))    
    family_name = element.Symbol.FamilyName
    type_name = element.Name
    combined_name = family_name + ":" + type_name
    name_parts = family_name.split(':')    
    dash_parts = name_parts[0].split('-')
    if len(dash_parts)>=3:
        keyword = dash_parts[1]
        for row in range(1, last_row + 1):
            cell = worksheet.Cells[row, 6] #column F is represented by index 6
            if ':' in str(cell.Value2):
                value = cell.Value2
                cell_name_parts=value.split(':')
                cell_dash_parts=cell_name_parts[0].split('-')
                cell_keyword = cell_dash_parts[1]
                if keyword == cell_keyword:
                    cell = worksheet.Cells[row, 5]
                    return str(cell.Value2)
            
    return 'mcr code is invalid'       


#11. noneType element filter
def noneType_filter(element_id):
    element = doc.GetElement(ElementId(element_id))
    
    element_type_id = element.GetTypeId()

    if element_type_id == ElementId.InvalidElementId:
        return "invalid element"
    else:
        return "valid element"
    

#12. actual parameter name finder
def actual_parameter_name_finder(parameter_name,list_to_check):
    comments = []
    corrected_parameters = []
    correct_parameters = []
    
    parameter_name_lower = parameter_name.replace(" ", "").lower()  
    
    for parameter in list_to_check:
        formatted_parameter = parameter.replace(" ", "").lower()

        if formatted_parameter ==  parameter_name_lower :
            actual_parameter_name = parameter            
            return actual_parameter_name
        
    return None

#13. is text NA checker  [add function to search in column AB for units]
def is_permutation_of_na(string):
    # Remove non-alphanumeric characters and convert to upper case
    string = re.sub(r'\W+', '', string).upper()
    
    # Check if the string is a permutation of 'NA'
    return sorted(string) == sorted('NA')

#14. input requirement finder
from datetime import datetime

def input_req(worksheet, parameter_name, column_z, parameter_input,current_row):
    invalid_entries = ['-','0','NIL']
    NA_check = is_permutation_of_na(parameter_input)

    if column_z == 'Any text':
        if NA_check == True or parameter_input in invalid_entries:
            return f'{parameter_name} input is invalid. Enter valid text'
        else:
            return 'correct input'
        
    elif column_z == 'Any value':
        if NA_check == True or parameter_input in invalid_entries:
            return f'{parameter_name} input is invalid. Enter valid value'
        else:
            return 'correct input'
    
    elif column_z == "Any number":
        if NA_check == True or parameter_input in invalid_entries:
            return f'{parameter_name} input is invalid. Enter valid number'
        else:
            try:
                float(parameter_input)
                return 'correct input'        
            except ValueError:
                column_AB_obj = worksheet.Cells[current_row, 28]
                column_AB = column_AB_obj.Value2
                if str(column_AB) in parameter_input:
                    return 'correct input'
                else:
                    return f'{parameter_name} input is not a number'
        
    elif column_z == "Boolean (True/False)":
        lower_case = parameter_input.lower()
        if NA_check == True or parameter_input in invalid_entries:
            return f'{parameter_name} input is invalid. Enter True/False'
        elif lower_case == 'true' or lower_case == 'false':
            return 'correct input'
        else:
            return f'{parameter_name} input is not True or False'
        
    elif column_z == 'Text contains':
        column_AA_obj = worksheet.Cells[current_row, 27]
        column_AA = column_AA_obj.Value2
        if NA_check or parameter_input in invalid_entries:
            return f'{parameter_name} input is invalid. Enter text containing {column_AA}'
        elif column_AA in parameter_input:
            return 'correct input'
        else:
            return f'{parameter_name} input does not contain {column_AA}'
    
    elif column_z == 'Any date':
        date_format = "%Y-%m-%d"
        if NA_check or parameter_input in invalid_entries:
            return f'{parameter_name} input is invalid. Enter valid date'
        else:
            try:
                datetime.strptime(parameter_input,date_format)
                return 'correct input'
            except ValueError:
                return f'{parameter_name}' + ' input date format incorrect'
    
    elif column_z == 'Value is one of (comma separated values)':
        column_AA_obj = worksheet.Cells[current_row, 27]
        column_AA = column_AA_obj.Value2
        values = column_AA.split(',')
        if NA_check or parameter_input in invalid_entries:
            return f'{parameter_name} input is invalid. Enter text containing {column_AA}'
        else: #parameter_input not in values
            return f'{parameter_name}' + ' input is not one of the given values'    
    else:
        return f'{parameter_name}'
    
    
#15. combined input checker
def input_checker(excel,file_path, sheet_name, element_id, door_type, parameters_to_check): #door_type = mcr code
    workbook = excel.Workbooks.Open(file_path)
    worksheet = workbook.Sheets[sheet_name]
    
    last_row = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
    
    if door_type == "Invalid Element Type ID":
        return 'MCR code does not exist'
    parameter_input_comments = []
    parameter_input_list = []
    try:
        for row in range(1, last_row + 1):
            cell = worksheet.Cells[row, 5]  # Column E is represented by index 5

            if cell.Value2 == door_type:
                current_row = row + 1

                found_empty_cell = False
                while True:
                    parameter_obj = worksheet.Cells[current_row, 24]  # Column X is represented by index 24
                    parameter_name = parameter_obj.Value2
                    if not parameter_obj.Value2 : #if cell is empty, end the loop
                        found_empty_cell = True
                        raise StopIteration

                    if found_empty_cell == False:
                        actual_parameter_name = actual_parameter_name_finder(parameter_name,parameters_to_check) 

                        if actual_parameter_name != None:  #if can find the actual_parameter_name
                            parameter_input = str(get_parameter_value(element_id, actual_parameter_name))

                            column_z_obj = worksheet.Cells[current_row, 26] # Column Z is represented by index 26, input req column
                            column_z = column_z_obj.Value2
                            #input checker
                            if parameter_input == 'None' :
                                comments = f'{parameter_name}' + ' input is empty'
                                parameter_input_comments.append(comments)
                                parameter_input_list.append(' ')
                            elif parameter_name == 'Family':
                                pass    
                            else:
                                comments = input_req(worksheet, parameter_name, column_z, parameter_input,current_row)
                                if comments != 'correct input':
                                    parameter_input_comments.append(comments)
                                    parameter_input_list.append(parameter_input)

                        current_row += 1
    
    except StopIteration:
        pass
    
    return parameter_input_list, parameter_input_comments

#16. retrieve excel file path
def file_path_finder(category_code):
    sheet_name = "As Built"
    if category_code == "OST_ProjectInformation" or category_code == "OST_TitleBlocks":
        file_path = r"C:\Users\geeko\Documents\Work\JTC (on com)\Code stuff\Scope Export Jul-21-2023 - Project Information.xls"
    elif category_code == "OST_Doors":
        file_path = r"C:\Users\geeko\Documents\Work\JTC (on com)\Code stuff\Scope Export Jul-27-2023_Doors_V2.xls"
    elif category_code == "OST_MechanicalEquipment":
        file_path = r"C:\Users\geeko\Documents\Work\JTC (on com)\Code stuff\Scope Export Jul-05-2023_Pump.xls"
    elif category_code == "OST_ElectricalEquipment":
        file_path = r"C:\Users\geeko\Documents\Work\JTC (on com)\Code stuff\Scope Export Jul-06-2023 - MCR23.xls"
    elif category_code == "OST_StructuralColumns":
        file_path = r"C:\Users\geeko\Documents\Work\JTC (on com)\Code stuff\Scope Export Jul-17-2023_Column.xls"
    
    return file_path , sheet_name


#17. general checker 
def element_checker(category_code,worksheet,current_row):
    excel = Excel.ApplicationClass()  #open excel instance
    file_path, sheet_name = file_path_finder(category_code)
    element_ids = get_element_ids(excel, file_path, sheet_name, category_code) #list of RevitElementIDs of pumps
        
    count=0
    row = current_row
    all_correct_elements = 0
    for element_id in element_ids:
        element_id = int(element_id)
        nt_filter = noneType_filter(element_id)
        
        if nt_filter == "valid element":
            element_type = get_parameter_value(element_id, "Type Comments") #return either MCR03-01-01 or MCR03-01-02 or Invalid Element Type ID       
            #return mcr code from family & type name
            ft_to_mcr = mcr_from_ft(excel, file_path, sheet_name, element_id)
            correct_element_parameters = find_row_and_column(excel,file_path, sheet_name, element_type) #list of correct parameter names
            #check family name
            if correct_element_parameters != 'MCR code does not exist' :
                output, combined_name,correct_format = check_family_name(element_id,element_type,category_mcr_dict) #check if family and type name is in the correct format 
                
                # wrong family and type name format
                if output != 'correct':

                    row+=1
                    element_category= worksheet.Cells(row, 1)
                    element_category.Value = category_code

                    number_header = worksheet.Cells(row, 2)
                    number_header.Value = count+1

                    revitelementID_header = worksheet.Cells(row, 3)
                    revitelementID_header.Value = element_id

                    element_type_header = worksheet.Cells(row, 4)
                    element_type_header.Value = element_type

                    error_type = worksheet.Cells(row, 5)
                    error_type.Value = 'Family and Type name format is incorrect'
                    error_comments7 = ''
                    error_comments7 += f'Family and Type name: {combined_name}' + '\n'

                    explanation = 'Incorrect '
                    for error in output:
                        if error != output[-1]:                   
                            explanation += f'{error}' + ', '
                        else:
                            explanation += f'{error}'

                    error_comments7 += f'  Explanation: {explanation}' + '\n'
                    error_comments7 += f' {correct_format}'

                    error_description = worksheet.Cells(row, 6)
                    error_description.Value = error_comments7

                    revitfilename_header = worksheet.Cells(row, 7)
                    revitfilename_header.Value = doc.Title
            
            elif correct_element_parameters == 'MCR code does not exist' and ft_to_mcr != 'mcr code is invalid':
                element_type1 = ft_to_mcr
                output, combined_name,correct_format = check_family_name(element_id,element_type1,category_mcr_dict) #check if family and type name is in the correct format 
                
                # wrong family and type name format
                if output != 'correct':

                    row+=1
                    element_category= worksheet.Cells(row, 1)
                    element_category.Value = category_code

                    number_header = worksheet.Cells(row, 2)
                    number_header.Value = count+1

                    revitelementID_header = worksheet.Cells(row, 3)
                    revitelementID_header.Value = element_id

                    element_type_header = worksheet.Cells(row, 4)
                    element_type_header.Value = element_type

                    error_type = worksheet.Cells(row, 5)
                    error_type.Value = 'Family and Type name format is incorrect'
                    error_comments7 = ''
                    error_comments7 += f'Family and Type name: {combined_name}' + '\n'

                    explanation = 'Incorrect '
                    for error in output:
                        if error != output[-1]:                   
                            explanation += f'{error}' + ', '
                        else:
                            explanation += f'{error}'

                    error_comments7 += f'  Explanation: {explanation}' + '\n'
                    error_comments7 += f' {correct_format}'

                    error_description = worksheet.Cells(row, 6)
                    error_description.Value = error_comments7

                    revitfilename_header = worksheet.Cells(row, 7)
                    revitfilename_header.Value = doc.Title
                
            if correct_element_parameters != 'MCR code does not exist' : #double up as mcr checker                
                                
                element_parameters = get_parameter_names(doc, element_id) #list of parameters in revit

                #check and return list of missing parameters
              
                missing_parameters = compare_parameters(element_parameters,correct_element_parameters) #list of missing parameters

                #provide corrected parameter names

                comments,corrected_parameters, correct_parameters = parameter_corrector(missing_parameters, element_parameters) #list of comments of corrected elements, list of correct elements

                #input checker 
                
                #list to check for parameter inputs -> list to check = correct_element_parameters - correct_parameters + corrected parameters
                parameters_to_check = correct_element_parameters + corrected_parameters                
                for parameters in correct_parameters:
                    parameters_to_check.remove(parameters)
                #comments for parameter with incorrect inputs
                parameter_input_list, parameter_input_comments = input_checker(excel,file_path, sheet_name, element_id, element_type, parameters_to_check) 
                #element has no problems 
                if missing_parameters == [] and parameter_input_comments == []:
                    all_correct_elements +=1
                
                else:
                    #GENERATE REPORT

                    count+=1                    
                    
                    # missing/wrong parameter names
                    if missing_parameters != []:
                        row+=1
                        element_category= worksheet.Cells(row, 1)
                        element_category.Value = category_code

                        number_header = worksheet.Cells(row, 2)
                        number_header.Value = count

                        revitelementID_header = worksheet.Cells(row, 3)
                        revitelementID_header.Value = element_id

                        element_type_header = worksheet.Cells(row, 4)
                        element_type_header.Value = element_type
                        
                        error_type = worksheet.Cells(row, 5)
                        error_type.Value = 'Wrong/Missing Parameter Name'
                        error_comments1 = ''
                        for i in range(len(comments)):
                            error_comments1 += f'{comments[i]}' + '\n'                                                

                        for element in correct_parameters:                    
                            if element in missing_parameters:
                                missing_parameters.remove(element)

                        if missing_parameters != []:
                            error_comments1 += f' Missing parameters: \n  {missing_parameters}' 
                        
                        error_description = worksheet.Cells(row, 6)
                        error_description.Value = error_comments1
                        
                        revitfilename_header = worksheet.Cells(row, 7)
                        revitfilename_header.Value = doc.Title

                    #incorrect parameter inputs -> can consider using dictionary

                    if parameter_input_comments != []:
                        row+=1
                        element_category= worksheet.Cells(row, 1)
                        element_category.Value = category_code

                        number_header = worksheet.Cells(row, 2)
                        number_header.Value = count

                        revitelementID_header = worksheet.Cells(row, 3)
                        revitelementID_header.Value = element_id

                        element_type_header = worksheet.Cells(row, 4)
                        element_type_header.Value = element_type
                        
                        error_type = worksheet.Cells(row, 5)
                        error_type.Value = 'Wrong Parameter Input'
                        error_comments2 = ''
                        for i in range(len(parameter_input_comments)):
                            number = i+1
                            error_comments2 += f'{number}. {parameter_input_comments[i]}' + '\n'
                            error_comments2 += f'   Parameter Input: {parameter_input_list[i]}' +'\n'
                        
                        error_description = worksheet.Cells(row, 6)
                        error_description.Value = error_comments2
                        
                        revitfilename_header = worksheet.Cells(row, 7)
                        revitfilename_header.Value = doc.Title                   
                                                
                parameter_input_comments.clear() #reset list
                parameter_input_list.clear() #reset list
                
            elif correct_element_parameters == 'MCR code does not exist' and ft_to_mcr != 'mcr code is invalid': # wrong mcr code but can identify parameters using f&t

                element_type1 = ft_to_mcr
                                
                element_parameters = get_parameter_names(doc, element_id) #list of parameters in revit

                #check and return list of missing parameters

                correct_element_parameters = find_row_and_column(excel,file_path, sheet_name, ft_to_mcr) #list of correct parameter names
                missing_parameters = compare_parameters(element_parameters,correct_element_parameters) #list of missing parameters

                #provide corrected parameter names

                comments,corrected_parameters, correct_parameters = parameter_corrector(missing_parameters, element_parameters) #list of comments of corrected elements, list of correct elements

                #input checker 
                #list to check for parameter inputs -> list to check = correct_element_parameters - correct_parameters + corrected parameters
                parameters_to_check = correct_element_parameters + corrected_parameters
                for parameters in correct_parameters:
                    parameters_to_check.remove(parameters)
                
                #comments for parameter with incorrect inputs
                parameter_input_list, parameter_input_comments = input_checker(excel,file_path, sheet_name, element_id, element_type1, parameters_to_check) 

                #GENERATE REPORT

                count+=1
                row+=1
                element_category= worksheet.Cells(row, 1)
                element_category.Value = category_code
                    
                number_header = worksheet.Cells(row, 2)
                number_header.Value = count

                revitelementID_header = worksheet.Cells(row, 3)
                revitelementID_header.Value = element_id

                element_type_header = worksheet.Cells(row, 4)
                element_type_header.Value = element_type
                                
                #wrong type comments parameter input
                error_type = worksheet.Cells(row, 5)
                error_type.Value = 'Type Comments Parameter is invalid'
                
                error_comments3 = f'Input of Type Comments parameter: {element_type}' +'\n' + f' Correct Type Comments input: {element_type1}'
                error_description = worksheet.Cells(row, 6)
                error_description.Value = error_comments3
                
                revitfilename_header = worksheet.Cells(row, 7)
                revitfilename_header.Value = doc.Title
                                               
                # missing/wrong parameter names
                if missing_parameters != []:
                    row+=1
                    element_category= worksheet.Cells(row, 1)
                    element_category.Value = category_code

                    number_header = worksheet.Cells(row, 2)
                    number_header.Value = count

                    revitelementID_header = worksheet.Cells(row, 3)
                    revitelementID_header.Value = element_id

                    element_type_header = worksheet.Cells(row, 4)
                    element_type_header.Value = element_type
                    
                    error_type = worksheet.Cells(row, 5)
                    error_type.Value = 'Wrong/Missing Parameter Name'
                    error_comments4 = ''
                    for i in range(len(comments)):
                        error_comments4 += f'{comments[i]}' + '\n'                                                

                    for element in correct_parameters:                    
                        if element in missing_parameters:
                            missing_parameters.remove(element)

                    if missing_parameters != []:
                        error_comments4 += f' Missing parameters: \n  {missing_parameters}' 

                    error_description = worksheet.Cells(row, 6)
                    error_description.Value = error_comments4
                    
                    revitfilename_header = worksheet.Cells(row, 7)
                    revitfilename_header.Value = doc.Title
                    

                #incorrect parameter inputs -> can consider using dictionary
                
                if parameter_input_comments != []:
                    row+=1
                    element_category= worksheet.Cells(row, 1)
                    element_category.Value = category_code

                    number_header = worksheet.Cells(row, 2)
                    number_header.Value = count

                    revitelementID_header = worksheet.Cells(row, 3)
                    revitelementID_header.Value = element_id

                    element_type_header = worksheet.Cells(row, 4)
                    element_type_header.Value = element_type
                    
                    error_type = worksheet.Cells(row, 5)
                    error_type.Value = 'Wrong Parameter Input'
                    error_comments5 = ''
                    for i in range(len(parameter_input_comments)):
                        number = i+1
                        error_comments5 += f'{number}. {parameter_input_comments[i]}' + '\n'
                        error_comments5 += f'   Parameter Input: {parameter_input_list[i]}' + '\n'

                    error_description = worksheet.Cells(row, 6)
                    error_description.Value = error_comments5
                    
                    revitfilename_header = worksheet.Cells(row, 7)
                    revitfilename_header.Value = doc.Title                                                   
               
                parameter_input_comments.clear() #reset list
                parameter_input_list.clear() #reset list

            else:    #both mcr and family name cannot be used to identify parameters       
                count+=1
                
                row+=1
                element_category= worksheet.Cells(row, 1)
                element_category.Value = category_code

                number_header = worksheet.Cells(row, 2)
                number_header.Value = count

                revitelementID_header = worksheet.Cells(row, 3)
                revitelementID_header.Value = element_id

                element_type_header = worksheet.Cells(row, 4)
                element_type_header.Value = element_type
                                
                error_type = worksheet.Cells(row, 5)
                error_type.Value = 'Type Comments Parameter is invalid'
                error_comments6 = f'Input of Type Comments parameter: {element_type}' 
                
                error_description = worksheet.Cells(row, 6)
                error_description.Value = error_comments6
                
                revitfilename_header = worksheet.Cells(row, 7)
                revitfilename_header.Value = doc.Title            
                                            
                                           
    
    if len(element_ids) == 0: 
        count+=1
        row+=1
        element_category= worksheet.Cells(row, 1)
        element_category.Value = category_code

        number_header = worksheet.Cells(row, 2)
        number_header.Value = count

        error_type = worksheet.Cells(row, 5)
        error_type.Value = 'No elements found'

        error_description = worksheet.Cells(row, 6)
        error_comments8 = 'Possible reasons: \n 1. Wrong file type \n 2. Elements were not modelled in the Revit File \n 3. Both Type Comments and Family and Type name are incorrect'
        error_description.Value = error_comments8

        revitfilename = worksheet.Cells(row, 7)
        revitfilename.Value = doc.Title
                 
    if all_correct_elements == len(element_ids) and len(element_ids) != 0:
        element_category= worksheet.Cells(row, 1)
        element_category.Value = category_code
        
        error_type = worksheet.Cells(row, 5)
        error_type.Value = 'All errors have been amended'
    count = 1
    excel.Quit()    #closes the excel       
    return row

#18. file name discipline filter
def name_filter():
    file_name = str(doc.Title)
    name_parts = file_name.split('-')
    name_parts_ = file_name.split('_')
    if'AR' in name_parts or 'AR' in name_parts_ :
        return 'AR'
    else:
        return 'skip'

#19. get PI parameters
def get_project_information_parameters():
    # Define the category for Project Information
    project_info_category = BuiltInCategory.OST_ProjectInformation

    # Get the elements of the specified category from the document
    collector = FilteredElementCollector(doc).OfCategory(project_info_category)

    # Retrieve the first Project Information element from the collection
    project_info_element = collector.FirstElement()

    parameter_names = []

    if project_info_element:
        # Get all parameters of the Project Information element
        parameters = project_info_element.Parameters

        # Loop through the parameters and add their names to the list
        for param in parameters:
            parameter_names.append(param.Definition.Name)

    return parameter_names
    
#20. get Title Block parameters
def get_title_block_parameter_names():
    # Define the category for title blocks
    title_block_category = BuiltInCategory.OST_TitleBlocks

    # Get the elements of the specified category from the document
    collector = FilteredElementCollector(doc).OfCategory(title_block_category)

    title_block_parameter_names = []

    # Loop through the title blocks and get their parameters
    for title_block in collector:
        # Get all parameters of the title block
        parameters = title_block.Parameters

        # Loop through the parameters and add their names to the list
        for param in parameters:
            title_block_parameter_names.append(param.Definition.Name)

    return title_block_parameter_names

#21. get PI parameter value #category code part is still not added
def get_PI_parameter_value(category_code,parameter_name):
    
    category = Enum.Parse(BuiltInCategory, category_code)

    # Get the elements of the specified category from the document
    collector = FilteredElementCollector(doc).OfCategory(category)

    # Retrieve the first Project Information element from the collection (there is only 1 element)
    element = collector.FirstElement()

    if element:
        # Get all parameters of the Project Information element
        parameters = element.Parameters

        # Loop through the parameters to find the one with the specified name
        for param in parameters:
            if param.Definition.Name == parameter_name:
                # Get the value of the parameter
                parameter_value = param.AsString()
                return parameter_value

    return None


#22. PI input checker
def PI_input_checker(excel,file_path, sheet_name, category_code, element_type, parameters_to_check): #element_type = mcr code
    workbook = excel.Workbooks.Open(file_path)
    worksheet = workbook.Sheets[sheet_name]
    
    last_row = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
    
    
    parameter_input_comments = []
    parameter_input_list = []
    try:
        for row in range(1, last_row + 1):
            cell = worksheet.Cells[row, 5]  # Column E is represented by index 5

            if cell.Value2 == element_type:
                current_row = row + 1

                found_empty_cell = False
                while True:
                    parameter_obj = worksheet.Cells[current_row, 24]  # Column X is represented by index 24
                    parameter_name = parameter_obj.Value2
                    if not parameter_obj.Value2 : #if cell is empty, end the loop
                        found_empty_cell = True
                        raise StopIteration

                    if found_empty_cell == False:
                        actual_parameter_name = actual_parameter_name_finder(parameter_name,parameters_to_check) 

                        if actual_parameter_name != None:  #if can find the actual_parameter_name
                            parameter_input = str(get_PI_parameter_value(category_code, actual_parameter_name))

                            column_z_obj = worksheet.Cells[current_row, 26] # Column Z is represented by index 26, input req column
                            column_z = column_z_obj.Value2
                            #input checker
                            if parameter_input == 'None' :
                                comments = f'{parameter_name}' + ' input is empty'
                                parameter_input_comments.append(comments)
                                parameter_input_list.append(' ')
                            elif parameter_name == 'Family':
                                pass    
                            else:
                                comments = input_req(worksheet, parameter_name, column_z, parameter_input,current_row)
                                if comments != 'correct input':
                                    parameter_input_comments.append(comments)
                                    parameter_input_list.append(parameter_input)

                        current_row += 1
    
    except StopIteration:
        pass
    
    return parameter_input_list, parameter_input_comments


#23. PI coordinating function
def PI_checker(category_code, worksheet,current_row):
    excel = Excel.ApplicationClass()  #open excel instance
    file_path, sheet_name = file_path_finder(category_code)    
       
    is_AR_file = name_filter()    
    
    row = current_row

    if is_AR_file == 'AR':        
        if category_code == "OST_ProjectInformation":
            element_type = "MCR00-01"
        elif category_code == "OST_TitleBlocks":
            element_type = "MCR00-02"
        
        retrieved_PI_parameters = get_project_information_parameters()
        correct_PI_parameters = find_row_and_column(excel,file_path, sheet_name, element_type)

        #check and return list of missing parameters
        missing_parameters = compare_parameters(retrieved_PI_parameters,correct_PI_parameters)

        #provide corrected parameter names
        comments,corrected_parameters, correct_parameters = parameter_corrector(missing_parameters, retrieved_PI_parameters)

        #list to check for parameter inputs
        parameters_to_check = correct_PI_parameters + corrected_parameters
        for parameters in correct_parameters:
                parameters_to_check.remove(parameters)
        #comments for parameter with incorrect inputs
        parameter_input_list, parameter_input_comments = PI_input_checker(excel,file_path, sheet_name, category_code,element_type, parameters_to_check)

        if missing_parameters != []:
            row+=1
            element_category= worksheet.Cells(row, 1)
            element_category.Value = category_code

            number_header = worksheet.Cells(row, 2)
            number_header.Value = '1'

            revitelementID_header = worksheet.Cells(row, 3)
            revitelementID_header.Value = ' '

            element_type_header = worksheet.Cells(row, 4)
            element_type_header.Value = element_type

            error_type = worksheet.Cells(row, 5)
            error_type.Value = 'Wrong/Missing Parameter Name'
            error_comments1 = ''
            for i in range(len(comments)):
                error_comments1 += f'{comments[i]}' + '\n'                                                

            for element in correct_parameters:                    
                if element in missing_parameters:
                    missing_parameters.remove(element)

            if missing_parameters != []:
                error_comments1 += f' Missing parameters: \n  {missing_parameters}'

            error_description = worksheet.Cells(row, 6)
            error_description.Value = error_comments1

            revitfilename_header = worksheet.Cells(row, 7)
            revitfilename_header.Value = doc.Title

        if parameter_input_comments != []:
            row+=1
            element_category= worksheet.Cells(row, 1)
            element_category.Value = category_code

            number_header = worksheet.Cells(row, 2)
            number_header.Value = '1'

            revitelementID_header = worksheet.Cells(row, 3)
            revitelementID_header.Value = ' '

            element_type_header = worksheet.Cells(row, 4)
            element_type_header.Value = element_type

            error_type = worksheet.Cells(row, 5)
            error_type.Value = 'Wrong Parameter Input'
            error_comments2 = ''
            for i in range(len(parameter_input_comments)):
                number = i+1
                error_comments2 += f'{number}. {parameter_input_comments[i]}' + '\n'
                error_comments2 += f'   Parameter Input: {parameter_input_list[i]}' +'\n'

            error_description = worksheet.Cells(row, 6)
            error_description.Value = error_comments2

            revitfilename_header = worksheet.Cells(row, 7)
            revitfilename_header.Value = doc.Title
        
        if missing_parameters == [] and parameter_input_comments == []:
            row+=1
            element_category= worksheet.Cells(row, 1)
            element_category.Value = category_code

            error_type = worksheet.Cells(row, 5)
            error_type.Value = 'All errors have been amended'        

        parameter_input_comments.clear() #reset list
        parameter_input_list.clear() #reset list            
        
    else:
        row+=1
        element_category= worksheet.Cells(row, 1)
        element_category.Value = category_code

        number_header = worksheet.Cells(row, 2)
        number_header.Value = '1'

        error_type = worksheet.Cells(row, 5)
        error_type.Value = 'No elements found'

        error_description = worksheet.Cells(row, 6)
        error_comments3 = 'Wrong file type'
        error_description.Value = error_comments3

        revitfilename = worksheet.Cells(row, 7)
        revitfilename.Value = doc.Title
    
    excel.Quit()    #closes the excel
    return row


#24. coordination function

def everything_checker(file_path):
    t = Transaction(doc, 'Write Excel.') 
    t.Start()

    # Check if the Excel file already exists
    outputFilePath = file_path
    is_file_exist = os.path.isfile(outputFilePath)

    # Create or open the Excel application instance
    xlAppType = System.Type.GetTypeFromProgID("Excel.Application")
    xlApp = System.Activator.CreateInstance(xlAppType)
    xlApp.Visible = False

    if is_file_exist:
        # Open the existing workbook
        workbook = xlApp.Workbooks.Open(outputFilePath)
    else:
        # Create a new workbook
        workbook = xlApp.Workbooks.Add()

    # Get the first worksheet in the workbook
    worksheet = workbook.Sheets[1]

    # Find the last used row in column A (Element Category)
    last_row = worksheet.Cells(worksheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
    # Start writing data from the next row
    row = last_row #dont need to plus 1 because the element checker will start by adding 1 row

    # Excel sheet headers (if not already present)
    if not is_file_exist or worksheet.UsedRange.Rows.Count <= 1:
        headers = [
            ('Element Category', 'A'),
            ('Number', 'B'),
            ('RevitElementID', 'C'),
            ('Element Type', 'D'),
            ('Error Type', 'E'),
            ('Error Description', 'F'),
            ('RevitFileName', 'G')
        ]

        column_widths = {
            'A': 24,  # Width for 'Element Category' column
            'B': 8,  # Width for 'Number' column
            'C': 14,  # Width for 'RevitElementID' column
            'D': 15,  # Width for 'Element Type' column
            'E': 39,  # Width for 'Error Type' column
            'F': 93,  # Width for 'Error Description' column
            'G': 39   # Width for 'RevitFileName' column
        }

        for header, col in headers:
            header_cell = worksheet.Cells(1, col)
            header_cell.Value = header
            header_cell.Font.Bold = True  # Bold the header text
            header_cell.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
            column_width = column_widths[col]
            worksheet.Columns(col).ColumnWidth = column_width

    category_list = ["OST_ProjectInformation", "OST_TitleBlocks", "OST_Doors", "OST_MechanicalEquipment", "OST_ElectricalEquipment", "OST_StructuralColumns"]

    # Start generating excel report body
    for category_code in category_list:
        if category_code in ["OST_ProjectInformation", "OST_TitleBlocks"]:
            row = PI_checker(category_code, worksheet, row)
        
        else:
            row = element_checker(category_code, worksheet, row)

    # Save the workbook to the existing/new file
    workbook.SaveAs(outputFilePath)

    # Close the workbook and quit Excel application
    workbook.Close(False)
    xlApp.Quit()

    t.Commit()

    __window__.Close()

    return 'Excel File has been generated'

file_path = r"C:\Users\geeko\Documents\Work\JTC (on com)\Test export.xlsx"
print(everything_checker(file_path))
