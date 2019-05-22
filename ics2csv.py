#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: eva borger

Usage:
    Takes an calendar export in ics format and returns a csv file with calendar event names, dates, times and duration.
    Tested with Outlook and Google calendar ics exports.
    Provide two arguments: input filename (ics file) and output filename (csv file):
        > ptyhon ics2csv.py [inputfile] ['outputfilename.csv']
"""

import csv 
import sys
from datetime import datetime

#calendar = "OutlookCalendar.ics"

def parse_ics(calendar):
    '''takes an ics file and parses it into a list of dictionaries of calendar events'''
    data = []
    event = []
    with open(calendar, "r") as f:
        data_temp = []
        lines = f.readlines()
        for line in lines:
            if line.startswith((" ", "ATTENDEE", "ORGANIZER")):
                continue
            else:
                line = line.strip()
                event.append(line)
                if line == 'END:VEVENT':
                    data_temp.append(event)
                    event = []
        idx = data_temp[0].index('BEGIN:VEVENT')
        data_temp[0] = data_temp[0][idx:]

    data = make_dict(data_temp)
    
    return data


def make_dict(datalist):
    data_dict = []
    for d in datalist:
        event_dict = {}
        for i in d:
            kv = i.split(":")
            event_dict[kv[0]] = kv[1]
        data_dict.append(event_dict)
    
    return data_dict


def clean_data(data):
    clean_data = []
    for event in data:
        ev = {}
        ev['SUMMARY'] = event["SUMMARY"]
        
        try:
            start_datetime = event["DTSTART;TZID=GMT Standard Time"]
            start_datetime = datetime.strptime(start_datetime.replace("T", ""), "%Y%m%d%H%M%S")
            ev["start_date"] = datetime.strftime(start_datetime, "%Y-%m-%d")
            ev["start_time"] = datetime.strftime(start_datetime, "%H:%M")

            end_datetime = event["DTEND;TZID=GMT Standard Time"]
            end_datetime = datetime.strptime(end_datetime.replace("T", ""), "%Y%m%d%H%M%S")
            ev["end_date"] = datetime.strftime(end_datetime, "%Y-%m-%d")
            ev["end_time"] = datetime.strftime(end_datetime, "%H:%M")
        
            ev["duration [minutes]"] = (end_datetime - start_datetime).total_seconds()/60
            ev["duration [hours]"] = (end_datetime - start_datetime).total_seconds()/3600
        
        except:
            try:
                start_datetime = event["DTSTART"]
                start_datetime = datetime.strptime(start_datetime.replace("T", "").replace("Z", ""), "%Y%m%d%H%M%S")
                ev["start_date"] = datetime.strftime(start_datetime, "%Y-%m-%d")
                ev["start_time"] = datetime.strftime(start_datetime, "%H:%M")
    
                end_datetime = event["DTEND"]
                end_datetime = datetime.strptime(end_datetime.replace("T", "").replace("Z", ""), "%Y%m%d%H%M%S")
                ev["end_date"] = datetime.strftime(end_datetime, "%Y-%m-%d")
                ev["end_time"] = datetime.strftime(end_datetime, "%H:%M")
            
                ev["duration [minutes]"] = (end_datetime - start_datetime).total_seconds()/60
                ev["duration [hours]"] = (end_datetime - start_datetime).total_seconds()/3600
            
            except:
                ev["duration"] = "all day"
        
        clean_data.append(ev)
    
    return clean_data



def export_data(data, filename):
    fieldnames = ["SUMMARY", 
                  "start_date", "start_time", 
                  "end_date", "end_time", 
                  "duration [minutes]", "duration [hours]", "duration"]
    with open(filename, "w") as f:
        writer = csv.DictWriter(f, delimiter = ",", fieldnames = fieldnames, lineterminator="\n")
        writer.writeheader()
        for row in data:
            writer.writerow(row)
        print("CSV file", sys.argv[2], "has been saved in the working directory")
        


if __name__ == "__main__":    
    if (len(sys.argv) < 2):
        print(__doc__)
    else: 
        calendar = sys.argv[1]
    
    data = parse_ics(calendar)
    clean_data= clean_data(data)
    export_data(clean_data, sys.argv[2])
    print("Success")
