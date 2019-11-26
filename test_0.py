def CreateEvent(start, subject, duration, location):

    import win32com.client #pywin32 API libs venum illana odaadhu
    #lib links directly to MAPI and outlookcal URI handler - but somehow gets things done without cross-threading/linking the open outlook window
    oOutlook = win32com.client.Dispatch("Outlook.Application")
    appointment = oOutlook.CreateItem(1)
    appointment.Start = start
    appointment.Subject = subject
    appointment.Duration = duration
    appointment.Location = location
    appointment.ReminderSet = True
    appointment.ReminderMinutesBeforeStart = 15
    appointment.ResponseRequested = True
    appointment.Display()



import sapcai #conda package la illa - use pip or nothing
from datetime import *
import dateutil.parser #python-dateutil venum idhuku

client = sapcai.Client("085edfd92327e593403a3859acb35471")
x = client.request.analyse_text('can we have a meeting next week at C401')
print(x.raw)

place = "Workplace"
object = "Meeting"
dt = str(datetime.now()+ timedelta (minutes = 15))
meeting_dur = 60

##this for we can suggest to write ?
for ent in x.entities:
    if ent.name == 'place':
        place = ent.value.title()

    elif ent.name == 'object':
        object = ent.value.title()

    elif ent.name == 'datetime':
        dtX = dateutil.parser.parse(ent.iso) + timedelta (hours = 5, minutes=30)
        dt = str(datetime.strftime(dtX, "%d %b %Y %H:%M"))
    elif ent.name == 'duration':
        dtX = (datetime.now() + timedelta (days = ent.days))
        dt = str(datetime.strftime(dtX, "%d %b %Y %H:%M"))

#this function shall be concrete, this is the final thing they'll call, no need to write this
CreateEvent(dt, object, meeting_dur, place)