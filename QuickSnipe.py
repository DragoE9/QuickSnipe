import pandas
import datetime
import re

#Read everything into variables
do_not_hit = (input("Enter Embassies to Exclude: ")).strip().lower().split(",")
min_trig = int(input("Enter Minimum Trigger Length (in seccods): "))
update = input("Major or Minor: ").lower().strip()
sheet = pandas.read_excel("sheet.xlsx",usecols="A:J",dtype={'Embassies': str},na_filter=False)

with open("regions.txt","r") as regions_file:
    targets = [line.strip().lower() for line in regions_file]
    
#Make sure Pandas is interpreting the update column as DateTimes
sheet["Major Upd. (true)"] = pandas.to_datetime(sheet["Major Upd. (true)"], format="%H:%M:%S")
sheet["Minor Upd. (est)"] = pandas.to_datetime(sheet["Minor Upd. (est)"], format="%H:%M:%S")

#Make All The Regions Lowercase.
sheet["Regions"] = sheet["Regions"].apply(lambda x: x.lower())
    
trig_delta = datetime.timedelta(seconds=min_trig)
trig_sheet = []
#iterate throug every given target
for target in targets:
    #Find out where it is
    print("\nSetting Trigger For: {}".format(target))
    pattern = r"^" + target + r"(?:[*~^])?(?![a-zA-Z0-9\s])"
    try:
        index = sheet.loc[sheet["Regions"].str.contains(pattern)].index[0]
    except:
        print("ERROR - Couldn't Find {}".format(target))
        continue
    #Validate Embassies
    embassies = (sheet["Embassies"].iloc[index]).lower().split(",")
    overlap = set(embassies).intersection(set(do_not_hit))
    if len(overlap) > 0:
        print("WARN - {} has excluded embassies".format(target))
    #Find its update time
    if update == "major":
        update_time = sheet["Major Upd. (true)"].iloc[index]
        #Find a trigger for it
        for i in range(index,0,-1):
            difference = update_time - sheet["Major Upd. (true)"].iloc[i]
            if difference >= trig_delta:
                k = i
                if sheet["# Nations"].iloc[k] > 50:
                    k = i - 1
                trig_region = sheet["Regions"].iloc[k]
                trig_region_url = "https://www.nationstates.net/region=" + trig_region.lower().replace(" ","_")
                trig_length = difference.total_seconds()
                trig_sheet.append([target,update_time,trig_region,trig_length])
                break
    if update == "minor":
       update_time = sheet["Minor Upd. (est)"].iloc[index]
       #Find a trigger for it
       for i in range(index,0,-1):
            difference = update_time - sheet["Minor Upd. (est)"].iloc[i]
            if difference >= trig_delta:
                k = i
                if sheet["# Nations"].iloc[k] > 50:
                    k = i - 1
                trig_region = sheet["Regions"].iloc[k]
                trig_region_url = "https://www.nationstates.net/region=" + trig_region.lower().replace(" ","_")
                trig_length = difference.total_seconds()
                trig_sheet.append([target,update_time,trig_region,trig_length])
                break

#Report
the_now = datetime.datetime.now()
title = "trigsheets/" + str(the_now.year) + str(the_now.month) + str(the_now.day) + ".txt"
j = 1
trig_sheet = sorted(trig_sheet, key=lambda x: x[1])
with open(title,"w") as file:
    for entry in trig_sheet:
        file.write(str(j) + ". https://www.nationstates.net/region=" + entry[0].lower().replace(" ","_") + " [" + str(entry[1].hour) + ":" + str(entry[1].minute) + ":" + str(entry[1].second) + "]\n")
        file.write("https://www.nationstates.net/region=" + (re.sub(r'[^\w\s]+$', '', entry[2])).lower().replace(" ","_") + " (" + str(entry[3]) + "s)\n\n")
        j += 1