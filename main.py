# Created By Prishad Mitchell -- Replit(5tup1d) -- Github(BlackBatman1980)
# Read the ReadMe for more info
# For AngelBotics -- Team 1339

import os
import sys
from statistics import median
import random
import time
import json

if os.name == "nt":
  clear = "cls"
  rm = "del"
else:
  clear = "clear"
  rm = "rm"

try:
  import requests
  import openpyxl
  
except:
  print("[x] Could not find libraries needed to function...")
  print("[+] Installing Libraries: openpyxl, requests")
  print("THIS CAN TAKE A WHILE, PLEASE BE PATIENT")
  time.sleep(2)
  try:
    os.system("pip3 install openpyxl requests")
    import openpyxl
    import requests
    
  except Exception as Error:
    print("[x] Encountered an Error while installing")
    print(Error)
    exit()

os.system(clear)
time.sleep(2)
try:
  xlsxFile = sys.argv[1]
except:
  print("[!] Please provide a file name")
  exit()

ROW_INDEX = 1
ERRORS = 0
teamIndex = 0
abet = "abcdefghijklmnopqrstuvwxyz"
abet = abet.upper()
abetL = []

for bet in abet:
  abetL.append(bet)

def generate_filename():
  new_name = random.randint(100000, 1000000)
  new_name = str(new_name)
  return new_name

def saveJson(data):
  fileName = generate_filename() + "_to_json.json"
    
  with open(fileName, 'w') as jFile:
    jFile.write(data)
        
  return fileName

def Start(fileName):
  
  with open(fileName, "r") as json_file:
    json_data = json.load(json_file)
  
  teamIDS = []
  totalGames = []

  for i in range(100):
    try:
      DATA = (json_data["documents"][i])["_id"]
      teamIDS.append(str(DATA))
    except Exception:
      pass
    
  DATA = json_data["documents"]
  for index in range(len(DATA)):
    totalGames.append(DATA[index]["__v"])

  for i in range(len(teamIDS)):
    for id,V in zip(teamIDS, totalGames):
      obj = Find(id,V,json_data, teamIDS.index(teamIDS[i]))
      # obj = Find("1111",_v,json_data, 7)
      GET(obj)
    
  cleanUp()
  

def Find(ID,V,json_data,index):
  global teamIndex
  global ERRORS


  isPitScouted = []
  autoRoutineHeaders = []
  autoRoutineData = []
  pitScoutHeaders = []
  pitScoutData = []
  teamNotes = []
  shotLowMedian = []
  shotHighMedian = []
  scoredLowMedian = []
  scoredHighMedian = []
  climbMedian = []
  brokeDownMedian = []
  cargoLowMedian = []
  cargoHighMedian = []

  DATA = (json_data["documents"][index]).get("isPitScouted", {})
  isPitScouted.append(str(DATA))
  
  DATA = list((json_data["documents"][index]).get("pitScout", {}))
    
  
  for z in DATA:
    if z in pitScoutHeaders:
      pass
    elif z == "driveBaseType":
      pass
    elif z == "autoRoutines":
      pass
    else:
      pitScoutHeaders.append(z)

  
  try:
    DATA = (json_data["documents"][index]).get("pitScout", {})["autoRoutines"]
      
    for g in range(len(list(DATA))):
      for l in range(len(list(DATA[g]))):
        if list(DATA[g])[l] in autoRoutineHeaders or list(DATA[g])[l] == "_id":
          pass
        else:
          autoRoutineHeaders.append(str(list(DATA[g])[l]))
  
  except Exception as Error:
    ERRORS += 1

  try:
    for u in range(len(autoRoutineHeaders)):
      auto = autoRoutineHeaders[u]
    
      DATA = (json_data["documents"][index]).get('pitScout', {}).get("autoRoutines")
        
      for i in range(len(DATA)):
        autoRoutineData.append(str(DATA[i][auto]))
      
  except Exception as Error:
    ERRORS += 1

  
  try:
    for y in range(len(pitScoutHeaders)):
      pit = pitScoutHeaders[y]
        
      DATA = (json_data["documents"][index]).get('pitScout', {})[pit]
      if DATA in pitScoutData:
        pass
      
      else:
        pitScoutData.append(str(DATA))

        
  except Exception as Error:
    ERRORS += 1


  #--# Extracting Game Data #--#

  DATA = (json_data["documents"][index]).get("games", {})
  
  # Get Notes On Games
  for i in range(V):
    try:
      teamNotes.append(DATA[i]["notes"])
    except:
      ERRORS += 1
  
  #--# Averaging Data #--#
  # Extracting the data first
  
  totalCycles = []
    
  for i in range(V):
    try:
      cycle = DATA[i]["cycles"]
      totalCycles.append(len(cycle))
    except:
      ERRORS += 1
    
    try:
      m = (DATA[i]["cargoShotHigh"])
      shotHighMedian.append(m)
      
      m = (DATA[i]["cargoShotLow"])
      shotLowMedian.append(m)
    
      m = (DATA[i]["cargoScoredHigh"])
      scoredHighMedian.append(m)
      
      m = (DATA[i]["cargoScoredLow"])
      scoredLowMedian.append(m)
      
      m = (DATA[i]["brokeDown"])
      brokeDownMedian.append(m)
      
      m = (DATA[i]["climb"])
      climbMedian.append(m)
      
    except:
      ERRORS += 1
    
    try:
      cargoHighMedian = []
      cargoLowMedian = []
      l = DATA[i]["auto"]["cargoLow"]
      h = DATA[i]["auto"]["cargoHigh"]
      cargoHighMedian.append(h)
      cargoLowMedian.append(l)
      
    except:
      ERRORS += 1
    
    try:
      cycleTime = []
      cargoShot = []
      HighGoal = []
      cargoScored = []
      
      
      for y in range(len(totalCycles)):
        for z in range(totalCycles[y]):
          cycleTime.append(DATA[i]["cycles"][z]["cycleTime"])
          cargoShot.append(DATA[i]["cycles"][z]["cargoShot"])
          HighGoal.append(DATA[i]["cycles"][z]["HighGoal"])
          cargoScored.append(DATA[i]["cycles"][z]["cargoScored"])
      
    except:
      ERRORS += 1
  
  # test = str(median(cycleTime))
  # print(test)
  # exit()
  
  # TODO -- CYCLE TIME IS NOT WORKING -- statistics.StatisticsError: no median for empty data
  cycleTimeMedian = str(median(cycleTime))
  cargoShotMedian = str(median(cargoShot))
  HighGoalMedian = str(median(HighGoal))
  cargoScoredMedian = str(median(cargoScored))
  scoredHighMedian = str(median(scoredHighMedian))
  scoredLowMedian = str(median(scoredLowMedian))
  shotHighMedian = str(median(shotHighMedian))
  shotLowMedian = str(median(shotLowMedian))
  brokeDownMedian = str(median(brokeDownMedian))
  climbMedian = str(median(climbMedian))
  cargoHighMedian = str(median(cargoHighMedian))
  cargoLowMedian = str(median(cargoLowMedian))
  
  medians = [test, cargoShotMedian, HighGoalMedian, cargoScoredMedian, scoredHighMedian, scoredLowMedian, shotHighMedian, shotLowMedian, brokeDownMedian, climbMedian, cargoHighMedian, cargoLowMedian]
  
  # Make an instance of this team to work with
  team_obj = TEAM(ID, medians=medians, isPitScouted=isPitScouted , autoRoutineHeaders=autoRoutineHeaders , autoRoutineData=autoRoutineData , pitScoutHeaders=pitScoutHeaders , pitScoutData=pitScoutData)
  
  return team_obj

class TEAM:
  def __init__(self, teamID, medians, isPitScouted=None, autoRoutineHeaders=None, autoRoutineData=None, pitScoutHeaders=None, pitScoutData=None):
    
    self.ID = teamID
    self.isPitScouted = isPitScouted
    self.autoRoutineHeaders = autoRoutineHeaders
    self.autoRoutineData = autoRoutineData
    self.pitScoutHeaders = pitScoutHeaders
    self.pitScoutData = pitScoutData
    self.medians = medians


def PUT(worksheet, data, cell):
  worksheet[cell] = data
  

def GET(self):
  global ROW_INDEX
  global ERRORS
  
  abetIndex = 1

  wb.create_sheet(self.ID)
  worksheet = wb[self.ID]
  worksheet.title = "Team " + self.ID
  PUT(worksheet, "Is PitScouted", "A1")
  PUT(worksheet, self.isPitScouted[0], "A2")

  for data in self.pitScoutHeaders:
    cell = abetL[abetIndex] + str(ROW_INDEX)
    PUT(worksheet, str(data), cell)
    abetIndex += 1

  abetIndex = 1
  ROW_INDEX += 1
  
  for data in self.pitScoutData:
    cell = abetL[abetIndex] + str(ROW_INDEX)
    PUT(worksheet, str(data), cell)
    abetIndex += 1

  abetIndex = 0
  ROW_INDEX = 6
  
  for data in self.autoRoutineHeaders:
    cell = abetL[abetIndex] + str(ROW_INDEX)
    PUT(worksheet, str(data), cell)
    abetIndex += 1

  abetIndex = 0
  ROW_INDEX = 7
  counter = 0
  
  for data in self.autoRoutineData:
    if counter == 4:
      counter = 0
      ROW_INDEX = 7
      abetIndex += 1
      cell = abetL[abetIndex] + str(ROW_INDEX)
      PUT(worksheet, str(data), cell)

    else:
      cell = abetL[abetIndex] + str(ROW_INDEX)
      PUT(worksheet, str(data), cell)
      ROW_INDEX += 1
      counter += 1
      
  abetIndex = 0
  ROW_INDEX = 12
  counter = 0
  dex = 0
  
  
  # Set The Headers First
  cell = abetL[abetIndex] + str(ROW_INDEX)
  
  PUT(worksheet, "Average Cycle Time", cell)
  abetIndex += 1
  PUT(worksheet, "Average Cargo Shots", cell)
  abetIndex += 1
  PUT(worksheet, "Average High Goals", cell)
  abetIndex += 1
  PUT(worksheet, "Average High Scored", cell)
  abetIndex += 1
  PUT(worksheet, "Average Low Scored", cell)
  abetIndex += 1
  PUT(worksheet, "Average High Shots", cell)
  abetIndex += 1
  PUT(worksheet, "Average Low Shots", cell)
  abetIndex += 1
  PUT(worksheet, "Average Break Down", cell)
  abetIndex += 1
  PUT(worksheet, "Average Climbs", cell)
  abetIndex += 1
  PUT(worksheet, "Average High Cargo", cell)
  abetIndex += 1
  PUT(worksheet, "Average Low Cargo", cell)
  
  abetIndex = 0
  ROW_INDEX += 1
  
  # cycleTimeMedian, cargoShotMedian, HighGoalMedian, cargoScoredMedian, scoredHighMedian, scoredLowMedian, shotHighMedian, shotLowMedian, brokeDownMedian, climbMedian, cargoHighMedian, cargoLowMedian
  for data in self.medians:
    try:
      cell = abetL[abetIndex] + str(ROW_INDEX)
      data = data[dex]
      PUT(worksheet, str(data), cell)
      abetIndex += 1
      dex += 1
  
    except:
      ERRORS += 1
           
  ROW_INDEX = 1
  wb.save(xlsxFile)


def cleanUp():
  print("[*] Cleaning Up...")
  os.system(f"{rm} *.json")
  time.sleep(1)
  print("[+] Done!")
  print("[*] Encountered", ERRORS, "Errors")
  time.sleep(1)
  exit()
  
wb = openpyxl.load_workbook(xlsxFile)
spreadsheet = wb.active

url = "https://data.mongodb-api.com/app/data-abado/endpoint/data/beta/action/find"

payload = json.dumps({
    "collection": "teams",
    "database": "ScoutingCluster",
    "dataSource": "ScoutingCluster",
    "projection": {}
})

headers = {
    'Content-Type': 'application/json',
    'Access-Control-Request-Headers': '*',
    'api-key': 'arU5b7LKqMYun0lpPoUV3EXBxZ7YFcybYa2ZH9G4FSWnV9hnG6vXs38Al8tFVDqg'
}

response = requests.request("POST", url, headers=headers, data=payload)
if response.status_code == 200:
    print("[+] Connection Successful!")
    print("[*] Getting Data...")
else:
    print("[!] There was an error connecting to the database")
    print("Status code of request: {}".format(response.status_code))
    exit()
    
if len(response.content) == 0:
  print("[!] There is no Data availabe")
else:
  org_req = response.content
  decoded_req = org_req.decode()
  print("[*] Converting...")
  fileName = saveJson(decoded_req)
  Start(fileName)