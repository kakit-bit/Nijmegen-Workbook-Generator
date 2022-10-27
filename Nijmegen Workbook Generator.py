# Script Takes User Input for Number of Samples to Run
# For Each Number of Sample, Accepts User Input for Demographics

# Script Generates Excel Documents from Template
# Script Inputs Patient Demographics into Generated Document
# Script Saves Generated Document as per SOP Requirements

# Class saves Patient Demographics into Profile Object
class Patient():
    def __init__(self, first_name, last_name, sample_id, factor, form):
        self.first_name = first_name
        self.last_name = last_name
        self.sample_id = sample_id
        self.factor = factor
        self.form = form

# User Input Error Check Functions
def is_number(value):
    try:
        int(value)
        return True
    except ValueError:
        return False

def is_float(value):
    try:
        float(value)
        return True
    except ValueError:
        return False

def in_range(value):
    if value not in range(1,3):
        return False
    else:
        return True

def is_limit(value):
    if value[0] not in (">","<"):
        return True
    elif (value[0] in (">","<")) and (len(value[1:]) > 0) and (is_float(value[1:]) == True):
        return True
    else:
        return False

# Function to Check How Many Patients to Generate
def get_workload():
    while True:
        workload_string = input("Number of Inhibitors to Run: ")
        # Check if Input is a Valid Number
        if not is_number(workload_string):
            print("Not a valid number")
            continue
        workload = int(workload_string)
        return workload

# Function to Obtain Factor Result: ")
def get_factor():
    while True:
        factor_string = input("Factor VIII Result: ")
        # Check if Limit Breaker
        if not is_limit(factor_string):
            # Check if Input is a Valid Number
            if not is_float(factor_string):
                print("Not a valid number")
                continue
            continue
        return factor_string

# Function to Obtain Nijmegen Form:
def get_form():
    while True:
        form_string = input("(1) Clotting or (2) Chromogenic: ")
        # Check if Input is Valid Number
        if not is_number(form_string):
            print("Not a valid number")
            continue
        form_input = int(form_string)
        if not in_range(form_input):
            print("Not a valid choice")
            continue

        if form_input == 1:
            form = "Clotting"
        elif form_input == 2:
            form = "Chromogenic"
        return form

# Function Allows User to Input Patient Demographics
def get_patient():
    last_name = str(input(("Patient Last Name: "))).lower().capitalize()
    first_name = str(input(("Patient First Name: "))).lower().capitalize()
    sample_id = str(input(("Sample ID: "))).lower().capitalize()
    factor = get_factor()
    form = get_form()
    print("")

    patient = Patient(first_name, last_name, sample_id, factor, form)
    return patient

# Function Opens Excel Template and Writes
def to_excel():
    import openpyxl
    srcfile = openpyxl.load_workbook('FVIII Nijmegen-Bethesda Assay - Inhibitor Worksheet v2.xlsx',read_only=False, keep_vba= False)
    sheetname = srcfile['Sheet1']
    sheetname['F5'] = str(patient.last_name + ', ' + patient.first_name)
    sheetname['F6'] = str(patient.sample_id)
    sheetname['C6'] = str(today_date)
    sheetname['D8'] = str(patient.factor)
    sheetname['D17'] = tech_initials

    if patient.form == "Clotting":
        sheetname['G18'] = "Y"
    else:
        sheetname['G17'] = "Y"

    filename = 'Nijmegen '+patient.form+'.' + patient.sample_id + '.' + today_date + '.xlsx'
    filename = str(filename)

    srcfile.save(filename)

if __name__ == '__main__':

    # Initialize Empty Patient List
    all_patients = []

    #Get Today's Date
    from time import gmtime, strftime
    today_date = strftime("%d.%m.%Y", gmtime())

    # Get User Initials
    tech_initials = str(input("Enter Your Initials: ").upper())

    # Get Workload and Patient Demographics
    num_samples = get_workload()
    for x in range(num_samples):
        patient = get_patient()
        all_patients.append(patient)

    # Code to Write Information to Template
    for patient in all_patients:
        to_excel()