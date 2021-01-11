import xlrd
from netmiko import ConnectHandler
from config import USER_NAME, PASSWORD



def excel2dict(file_pth='tracker.xlsx', sheet_name='Sheet2'):
    '''Reads xlxs file and convert it to a dictionary
    args: 
        file_pth: string -> path to the excel file
        sheet_name: string -> sheet to be read
        
    return: python dict -> a nested python dict
                outer dict -> all pops by name
                inner dict -> inidividual pop config details
    '''
    work_book = xlrd.open_workbook(file_pth)
    sheet = work_book.sheet_by_name(sheet_name) 
    pops = {}
    for row in range(1, sheet.nrows):
        pop_name = sheet.cell(row, 0).value
        pop_details = {}
        for col in range(1, sheet.ncols):
            config = sheet.cell(0, col).value
            pop_details[config] = sheet.cell(row, col).value
        pops[pop_name] = pop_details
    return pops


def connect(pop_name='ABUJA-1'):
    '''Çonnects to a specific pop
    args: 
        pop_name: string -> name of target pop     
    '''
    all_pops = pops = excel2dict()
    pop = all_pops[pop_name]
    device = ConnectHandler(device_type='cisco_ios', 
                            ip=pop['description5'].split(' ')[0], 
                            username=USER_NAME, 
                            password=PASSWORD
                           )
    return device


def show_version(pop_name='ABUJA-1'):
    '''Çonnects to a specific pop and executes the show version command
    args: 
        pop_name: string -> name of target pop     
    '''
    device = connect(pop_name=pop_name)
    output = device.send_command('show version')
    device.disconnect()
    print('Done!')
    
    
def configure_interface(pop_name='ABUJA-1'):
    '''Çonnects to a specific pop and executes interface configuration
    args: 
        pop_name: string -> name of target pop     
    '''
    all_pops = excel2dict()
    pop = all_pops[pop_name]
    device = connect(pop_name=pop_name)
    device.send_command(pop['Interface '])
    device.send_command('description Link to LAN 2')
    device.send_command(pop['description5'])
    device.disconnect('no shutdown')
    device.disconnect('exit')
    print('Done!')
    
    
