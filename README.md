# OutlookHierarchy
Powershell script to pull full hierarchy information from Outlook from a list of names for import into Visio.
Connects Employees to Manager based on alias (contains employeeID)
User will be queried on whether each name inputted matches the employee found 

## Usage:

**Main input-file output-file [-correct]**

**input-file** must be in format:

    LastName, FirstName    
    LastName, FirstName    
    LastName, FirstName
  
**output-file** follows format:

    Name, UID, Department, Role, Manager    
    Name, UID, Department, Role, Manager  
    Name, UID, Department, Role, Manager
  
**-correct** flag will make script assume all names are correct in input file
  
