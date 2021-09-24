
import config as config
import datetime as datetime

def Date_Recent():
        now = datetime.datetime.now()
        weekday = now.weekday()
        if weekday == 0:
            today = now - datetime.timedelta(days = 3)
            today =  today.strftime("%Y-%m-%d")
            if today in config.US_Trading_Holiday:
                today = now - datetime.timedelta(days = 4)
                today =  today.strftime("%Y-%m-%d")
        elif weekday == 1:
            today = now - datetime.timedelta(days = 1)
            today =  today.strftime("%Y-%m-%d")
            if today in config.US_Trading_Holiday:
                today = now - datetime.timedelta(days = 4)
                today =  today.strftime("%Y-%m-%d")
        else:
            today = now - datetime.timedelta(days = 1)
            today =  today.strftime("%Y-%m-%d")
            if today in config.US_Trading_Holiday:
                today = now - datetime.timedelta(days = 2)
                today = today.strftime("%Y-%m-%d")
        recent = str(today)
        print(recent)
        return recent


def Date_Previous():
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday == 0: 
        today = now - datetime.timedelta(days = 4)
        today =  today.strftime("%Y-%m-%d")
        holiday = now - datetime.timedelta(days = 1)
        holiday =  holiday.strftime("%Y-%m-%d")
        if holiday in config.US_Trading_Holiday:
            today = now - datetime.timedelta(days = 5)
            today =  today.strftime("%Y-%m-%d")
    elif weekday == 1:
        today = now - datetime.timedelta(days = 4)
        today =  today.strftime("%Y-%m-%d")
        holiday = now - datetime.timedelta(days = 1)
        holiday =  holiday.strftime("%Y-%m-%d")
        if holiday in config.US_Trading_Holiday:
            today = now - datetime.timedelta(days = 5)
            today =  today.strftime("%Y-%m-%d")
        print(today)
    else:
        today = now - datetime.timedelta(days = 2)
        today =  today.strftime("%Y-%m-%d")
        holiday = now - datetime.timedelta(days = 1)
        holiday =  holiday.strftime("%Y-%m-%d")
        if holiday in config.US_Trading_Holiday:
            today = now - datetime.timedelta(days = 3)
            today = today.strftime("%Y-%m-%d")
        holiday = now - datetime.timedelta(days = 2)
        holiday =  holiday.strftime("%Y-%m-%d")
        if holiday in config.US_Trading_Holiday:
            today = now - datetime.timedelta(days = 5)
            today =  today.strftime("%Y-%m-%d")
    previous = str(today)
    print(previous)
    return previous

def File_Path(File_Date,File_Path_Text):
    """Completes date adjust and adds in text string for a complete file path"""
    File_Path = File_Path_Text + File_Date + '.xlsx'
    return File_Path