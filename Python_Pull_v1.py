from pprint import pprint
from docx import Document
from docx.shared import Inches
from shotgun_api3 import Shotgun

#Setting up API for Shotgun
SERVER_PATH = "https://medialab.shotgunstudio.com"
SCRIPT_NAME = 'Data_Pull'
SCRIPT_KEY = '6ea7f7d255e7face13aaba803e542e1cda4095068aaccfe6d07f20ff548e50c6'
sg = Shotgun(SERVER_PATH, SCRIPT_NAME, SCRIPT_KEY)

#Tools used to Query some things
'''
proj = sg.find_one("Project", [["name", "is", "Kohler Company"]])
print proj
query = sg.schema_read()['Asset'].keys()
print query
'''

#function to format the text that gets read by Note_Thread_Read
def formatted(x):
    #for key, value in x.iteritems():
        #print "$", key, "$", value
    x.pop("type", None)
    x.pop("id", None)
    Name = x.get('created_by', {}).get('name')
    Name2 = x.get('user', {}).get('name')
    global thread
    if 'content' in x:
        if Name == None:
            thread += "\n"
            thread += str(Name2) + ":"
            thread += "\n"
        else:
            thread += "\n"
            thread += str(Name) + ":"
            thread += "\n"
    else:
        pass
    x.pop("created_at", None)
    x.pop("user", None)
    return x

#Functions used to write out the documents to a directory in word.
def writeDoc(docTitle, feed, fileName):
    #Setting up word doc for file writing
    document = Document()
    document.add_heading(fileName, 0)

    #Writing it out
    try:
        document.add_paragraph(feed)
    except ValueError:
        pass
    try :
        document.save('I:\dev_JC\_Python\Data_Pull\Kohler\\' + fileName + '.docx')
    except IOError:
        document.save('I:\dev_JC\_Python\Data_Pull\Kohler\\' + "Bad_Name" + '.docx')

#Pulls a shotgun Thread based on the id
formatting ={}
filters = [ ['id', 'is', 74], ]  ##This is the controller, place whatever project ID you need docs for
project = sg.find_one('Project', filters, fields=['id'])
filters = [ ['project', 'is', {'type':'Project', 'id':project['id']}], ]
result = sg.find('Asset', filters, fields=['sg_rpm_number', 'notes', 'id', 'sg_rpm_number'])
for i in result:
    thread = ""
    dict = i['notes']
    excludeRPM = ['50428', '50825', '50625', '50213', '50048', '49721']
    excludeID = ['6072', '6352', '6196', '7701', '6176', '6171', '6158',
                 '6470', '7010', '7141', '7351', '7267', '7290', '7321']
    if not dict:
        continue
    if i['sg_rpm_number'] in excludeRPM:
        continue
    dictStr = (dict[0]['name'])
    dictStr = dictStr.splitlines()
    try:
        thread += str(dictStr[0])
    except IndexError:
        thread += 'null'
        continue
    try:
        fileName = 'Kohler - ' + i['sg_rpm_number']
    except TypeError:
        filename = 'Kohler - MISSING RPM NUMBER'
    print fileName
    currentID = str(dict[0]['id'])
    if currentID in excludeID:
        continue
    list3 = sg.note_thread_read(int(currentID), formatting)
    for key in list3:
        thread += "\n"
        note = formatted(key)
        if 'content' in note:
            thread += str(note['content'])
        else:
            pass
    thread += '\n**END OF THREAD**'
    writeDoc(dictStr, thread, fileName)
