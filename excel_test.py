import openpyxl, openpyxl_templates, openpyxl_utilities
from openpyxl.styles import Color, PatternFill


scenario_source = "C:\\Projects\\Internal testing\\riskinternaltesting\\RiskInternalTesting-WEB\\src\\test\\resources\\7005135.xlsx"
output_file = "C:\\Users\\EcBerry\\Desktop\\book1.xlsx"
myworkbook = openpyxl.load_workbook(scenario_source)


def match(object):

    value1 = object.value1
    value2 = object.value2
    workbook = object.workbook
    my_green = openpyxl.styles.colors.Color(rgb='00FF00')
    green_fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_green)
    my_red = openpyxl.styles.colors.Color(rgb='FF0000')
    red_fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_red)

    def matchmaker(value1, value2, workbook):
        workbook = workbook
        myworksheet = myworkbook.active

        if value1 == value2:
            myworksheet.cell(row=7, column=1).fill = green_fill
        else:
            myworksheet.cell(row=7, column=1).fill = red_fill

    return matchmaker(value1, value2, workbook)


class cameleon:

    value1 = None
    value2 = None
    workbook = None

    def __init__(self, value1, value2, workbook):
        self.value1 = value1
        self.value2 = value2
        self.workbook = workbook


camel_one = cameleon(10, 10, myworkbook)    # Cameleon object one
camel_two = cameleon(10, 11, myworkbook)    # Cameleon object two



match(camel_one)

@match
def camel_two(camel_two):
    #myworksheet = myworkbook.active
    #myworksheet.cell(row=7, column=2)
    #myworksheet.cell(row=7, column=2)
    #matchmaker2(camel_two)
    print("Hello")



myworkbook.save(output_file)
