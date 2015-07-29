emp_type_dict = {}
class_dict = {}
etp_unique_id_list = []

class EmpRecord:

    def __init__(self, name, ID, roster, date, className):
        self.name = name
        self.ID = ID
        self.roster = roster
        self.date = date
        self.className = className
        self.uniqueID = str(ID)+str(roster)

    def getName(self):
        return self.name

    def getID(self):
        return self.ID

    def getRoster(self):
        return self.roster

    def getDate(self):
        return self.date

    def getClassName(self):
        return self.className

    def getUniqueID(self):
        return self.uniqueID
    

class ColNames:
    # Class Info sheet columns
    name = 3
    emp_ID = 2
    roster = 13
    date = 11
    class_name = 14
    attended = 35
    etp_type = 34

    # Row where data starts on Class Info sheet
    cell_row = 7

    # ETP sheet columns
    etp_id = 4
    etp_roster = 7

    # Row where data starts on ETP sheet
    etp_cell_row = 3

    # HR sheet columns
    hr_id = 3
    hr_type = 7
    
    # Row where data starts on HR sheet
    hr_cell_row = 6
    

    # Destination cells
    header_row = 2
    
    new_name = 1
    new_emp_id = 2
    new_roster = 3
    new_date = 4
    new_class_name = 5
    new_unique_id = 6


def prepDoc():
   
    Cell(ColNames.header_row-1, ColNames.new_name).value = "Values not in ETP"
    
    Cell(ColNames.header_row, ColNames.new_name).value = "Name"
    Cell(ColNames.header_row, ColNames.new_emp_id).value = "Employee ID"
    Cell(ColNames.header_row, ColNames.new_roster).value = "Roster"
    Cell(ColNames.header_row, ColNames.new_date).value = "Date"
    Cell(ColNames.header_row, ColNames.new_class_name).value = "Class Name"
    Cell(ColNames.header_row, ColNames.new_unique_id).value = "Unique ID"


def create_emp_type_dict():
    print "Creating Employee Type Dictionary ..."
    current_row = ColNames.hr_cell_row
    id_col = ColNames.hr_id
    id_cell = Cell(current_row,id_col)
    type_col = ColNames.hr_type
    type_cell = Cell(current_row,type_col)
    while(not id_cell.is_empty()):
        if (type_cell.is_empty()):
            emp_type_dict[id_cell.value] = "EXTERNAL"
        else:
            emp_type_dict[id_cell.value] = type_cell.value
        current_row+=1
        id_cell = Cell(current_row,id_col)
        type_cell = Cell(current_row,type_col)


def create_class_dict():
    print "Creating DN Trainee Dictionary ..."
    current_row = ColNames.cell_row
    aCell = Cell(current_row, ColNames.name)
    while(not aCell.is_empty()):
        aName = Cell(current_row, ColNames.name).value
        aID = Cell(current_row, ColNames.emp_ID).value
        aRoster = Cell(current_row, ColNames.roster).value
        aDate = Cell(current_row, ColNames.date).value
        aClassName = Cell(current_row, ColNames.class_name).value
        aUniqueID = str(aID) + str(aRoster)
        aAttended = Cell(current_row, ColNames.attended).value
        a_etp_type = Cell(current_row, ColNames.etp_type).value

        #check trainee info page...
        if ( (not aID == "###") and (aAttended == "Yes") and (not (a_etp_type == "Ineligible" or a_etp_type == "#NOT IN CLASS LIST")) ):
            is_temp = (emp_type_dict[aID] == "EXTERNAL") or (emp_type_dict[aID] == "Temporary") or (emp_type_dict[aID] == "Intern")
            if(not is_temp):
                new_emp = EmpRecord(aName, aID, aRoster, aDate, aClassName)
                class_dict[aUniqueID] = new_emp
        current_row+=1
        aCell = Cell(current_row, ColNames.name)

        
def create_etp_unique_id_list():
    print "Creating ETP Unique ID List ..."
    current_row = ColNames.etp_cell_row
    aCell = Cell(current_row, ColNames.etp_id)
    while (not aCell.is_empty()):
        id_cell = Cell(current_row, ColNames.etp_id)
        roster_cell = Cell(current_row, ColNames.etp_roster)
        this_unique_id =  str(id_cell.value) + str(roster_cell.value)
        etp_unique_id_list.append(this_unique_id)
        current_row+=1
        aCell = Cell(current_row, ColNames.etp_id)

def find_discrepencies():
    print "Finding Discrepencies"
    # and print to screen
    current_row = ColNames.header_row+1
    for record in class_dict.items():
        key = record[0]
        value = record[1]
        if (key not in etp_unique_id_list):
            print_emp_record(value, current_row)
            current_row+=1

def print_emp_record(aRecord, aRow):
    Cell(aRow, ColNames.new_name).value = aRecord.getName()
    Cell(aRow, ColNames.new_emp_id).value = aRecord.getID()
    Cell(aRow, ColNames.new_roster).value = aRecord.getRoster()
    Cell(aRow, ColNames.new_date).value = aRecord.getDate()
    Cell(aRow, ColNames.new_class_name).value = aRecord.getClassName()
    Cell(aRow, ColNames.new_unique_id).value = aRecord.getUniqueID()
    


def main():


    this_wkbk = active_wkbk()
    sheets = all_sheets()

    # Create new sheet
    resultsSheet = "results"

    i = 0
    while (resultsSheet+"_"+str(i) in sheets):
        i+=1
    resultsSheet = resultsSheet + "_" + str(i)

    new_sheet(resultsSheet)
    active_sheet(resultsSheet)
    prepDoc()

    # Get Employee information
    trainee_sheet = raw_input("Enter name of sheet with HR Info: ")
    while (trainee_sheet not in sheets):
        trainee_sheet = raw_input("Not Found! Enter name of sheet with HR Info: ")
    active_sheet(trainee_sheet)
    create_emp_type_dict()

    # Get Class information
    class_info_sheet = raw_input("Enter name of sheet with Class Info: ")
    while (class_info_sheet not in sheets):
        class_info_sheet = raw_input("Not Found! Enter name of sheet with Class Info: ")
    active_sheet(class_info_sheet)
    create_class_dict()

    # Get ETP information
    etp_sheet = raw_input("Enter name of sheet with ETP Info: ")
    while (etp_sheet not in sheets):
        etp_sheet = raw_input("Not Found! Enter name of sheet with ETP Info: ")
    active_sheet(etp_sheet)
    create_etp_unique_id_list()

    active_sheet(resultsSheet);
    find_discrepencies()

    autofit(resultsSheet)

main()
