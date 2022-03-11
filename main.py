# Created By Prishad Mitchell -- Replit(5tup1d) -- Github(BlackBatman1980)
# Read the README for more info
# For AngelBotics Team 1339

import os
import sys
from statistics import median
import random
import time
import json

FAILED = []

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
  print("[!] Please provide an xlsx file name")
  print("Usage: python3 main.py <file.xlsx>")
  exit()

ROW_INDEX = 1
ERRORS = 0
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

  for i in range(500):
    try:
      DATA = (json_data["documents"][i])["_id"]
      teamIDS.append(str(DATA))
    except Exception:
      pass
    
  DATA = json_data["documents"]
  totalTeams = len(DATA)
  print(f"[*] Extracting Data From A Total Of {totalTeams} Teams")

  teamCounter = 0
  for id in teamIDS:
    if teamCounter >= totalTeams:
      pass
      
    else:
      try:
        print("[*] Getting From Team", id)
        totalGames = len(DATA[teamCounter]["games"])
        teamObj = Find(id,totalGames,json_data, teamIDS.index(id))
        errno = GET(teamObj)
        teamCounter += 1
      except Exception as Error:
        # raise Error
        FAILED.append(id)
        logErrors(Error)
        # exit()
        # print("Failed to get team", id)
  cleanUp()

def logErrors(error):
  with open("main.log", "a") as log:
    log.write(str(error))
    log.write("\n")
    
def Find(ID,TG,json_data,index):
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
  cycleTime = []
  cargoShot = []
  HighGoal = []
  cargoScored = []
  cargoHighMedian = []
  cargoLowMedian = []

  
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
    logErrors(Error)

  try:
    for u in range(len(autoRoutineHeaders)):
      auto = autoRoutineHeaders[u]
    
      DATA = (json_data["documents"][index]).get('pitScout', {}).get("autoRoutines")
        
      for i in range(len(DATA)):
        autoRoutineData.append(str(DATA[i][auto]))
      
  except Exception as Error:
    ERRORS += 1
    logErrors(Error)

  
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
    logErrors(Error)


  #--# Extracting Game Data #--#

  DATA = (json_data["documents"][index]).get("games", {})
  
  # Get Notes On Games
  for i in range(TG):
    try:
      teamNotes.append(DATA[i]["notes"])
    except Exception as Error:
      ERRORS += 1
      logErrors(Error)
  
  #--# Averaging Data #--#
  # Extracting the data first
  
      
  team_obj = TEAM(ID)
  team_obj.isPitScouted=isPitScouted
  team_obj.pitScoutData=[]
  team_obj.pitScoutHeaders=[]
  team_obj.autoRoutineHeaders=[]
  team_obj.autoRoutineData=[]
  
  try:
    team_obj.pitScoutData=pitScoutData
    team_obj.pitScoutHeaders=pitScoutHeaders
    team_obj.autoRoutineHeaders=autoRoutineHeaders
  
    team_obj.autoRoutineData=autoRoutineData
    
  except Exception as Error:
    ERRORS += 1
    logErrors(Error)
    
  return team_obj
  
  # cycleTimeMedian
  # cargoShotMedian
  # HighGoalMedian
  # cargoScoredMedian
  # scoredHighMedian
  # scoredLowMedian
  # shotHighMedian
  # shotLowMedian
  # brokeDownMedian
  # climbMedian
  # cargoHighMedian
  # cargoLowMedian


class TEAM:
  def __init__(self, teamID, isPitScouted=None, autoRoutineHeaders=None, autoRoutineData=None, pitScoutHeaders=None, pitScoutData=None):
    
    self.ID = teamID
    self.isPitScouted = isPitScouted
    self.autoRoutineHeaders = autoRoutineHeaders
    self.autoRoutineData = autoRoutineData
    self.pitScoutHeaders = pitScoutHeaders
    self.pitScoutData = pitScoutData


def PUT(worksheet, data, cell):
  worksheet[cell] = str(data)
  

def GET(self):
  global ROW_INDEX
  global ERRORS
  
  abetIndex = 1


  # Initialize the worksheet
  try:
    wb.create_sheet(self.ID)
  except Exception:
    if Exception == AttributeError:
      print("[x] Unable to create team in spreadsheet...")
      print("[x] Deleting Team...")
      del self
      return False
      
  worksheet = wb[self.ID]
  worksheet.title = "Team " + self.ID
  PUT(worksheet, "Is PitScouted", "A1")
  PUT(worksheet, self.isPitScouted, "A2")

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
  maxLen = len(self.autoRoutineHeaders)
  
  
  for i in range(maxLen):
    data = self.autoRoutineData[i]
    if data == None:
      data = "Null"

    if counter == maxLen:
      abetIndex = 0
      ROW_INDEX += 1
      cell = abetL[abetIndex] + str(ROW_INDEX)
      PUT(worksheet, str(data), cell)
      counter = 0
    
    else:
      cell = abetL[abetIndex] + str(ROW_INDEX)
      PUT(worksheet, str(data), cell)
      counter += 1
      abetIndex += 1
      

  
  ROW_INDEX = 1
  wb.save(xlsxFile)

  print("[+] Done")


def cleanUp():
  os.system(clear)
  print("[*] Cleaning Up...")
  os.system(f"{rm} *.json")
  time.sleep(1)
  print("[+] Done!")
  print("[*] Encountered", ERRORS, "Errors")
  if len(FAILED) == 0:
    print("[+] No major errors encountered!")
  else:
    print("[x] The program failed to extract information from the following teams:", FAILED)
    print("[!] See 'main.log' for more info")
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
    'api-key': '5yFyYYw67A0qv1yWrH55V8JgltCOau314l8lq41wDxe5U8rLYzpVpCZIoZSWVqH3'
}

response = requests.request("POST", url, headers=headers, data=payload)
if response.status_code == 200:
    print("[+] Connection Successful!")
    print("[*] Getting Data...")
else:
    print("[!] There was an error connecting to the database")
    print("Status code of request: {}".format(response.status_code))
    print("[*] Check the api key and/or url")
    exit()
    
if len(response.content) == 0:
  print("[!] There is no Data availabe")
else:
  org_req = response.content
  decoded_req = org_req.decode()
  print("[*] Converting...")
  fileName = saveJson(decoded_req)
  Start(fileName)