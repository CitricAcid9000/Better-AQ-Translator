'''
    This was made By Macarkeys here https://github.com/Macarkeys/PFC_SpellReader/tree/main and was updated by Me

    install extensions:
    pip install python-docx
'''

# I have this because of how inconsistant the AQ spell format is
global DeferNumberCounter
DeferNumberCounter = 0 

global spellcount
spellcount = 0

# combines paragraphs from a list of paragraphs
def collatPara(paragraphList):
  strPara = ''
  for i in paragraphList:
    strPara += cleanText(i.text)
  return strPara

def cleanText(textInput: str):
  ## index to double check this works correctly
  if '\u2019' in textInput:
    textInput = textInput.replace('\u2019',"'")
  if '\u201c' in textInput:
    textInput = textInput.replace('\u201c',"'")
  if '\u201d' in textInput:
    textInput = textInput.replace('\u201d',"'")
  if '\u2026' in textInput:
    textInput = textInput.replace('\u2026','...')
  if '\u2264' in textInput:
    textInput = textInput.replace('\u2264','<=')
  if '\u2013' in textInput:
    textInput = textInput.replace('\u2013','-')
  if '\u2018' in textInput:
    textInput = textInput.replace('\u2018',"'")
  if '\u2020' in textInput:
    textInput = textInput.replace('\u2020','')
  if '\u2192' in textInput:
    textInput = textInput.replace('\u2192','->')
  if '\u00b0F' in textInput:
    textInput = textInput.replace('\u00b0F',' degrees F')
  if '\u00b0C' in textInput:
    textInput = textInput.replace('\u00b0C',' degrees C')
  

  textInput = textInput.strip()
  return textInput

# Cleans the lists in spell table and preps. Removes the little header row and changes the empty row at bottom of every spell to be the description, removing the third column.
def cleanSpellTable(SpellTable):
  # turns a spell table to an 2D array that has all the discriptors and their atributes
  cleanedSpells = [[],[]]
  for i in range(len(SpellTable[0])):
    if (SpellTable[0][i],SpellTable[1][i]) != ('','') and ('Description' not in SpellTable[2][i].strip()):
      # grabs names and their discriptors
      cleanedSpells[0].append(SpellTable[0][i])
      # grabs the atributes (except description)
      cleanedSpells[1].append(SpellTable[1][i])
    elif SpellTable[0][i].strip() == '' and SpellTable[0][i-1].strip() != '':
      cleanedSpells[0].append('Description')
      cleanedSpells[1].append(SpellTable[2][i-6])
  return cleanedSpells

# Gets the indexes of each spell in each spell table
def getSpellIndexes(SpellTable):
  spellIndex = []
  x = 0
  i = 0 
  for i in range(len(SpellTable[0])):
    if '\u2013' in SpellTable[0][i] or '\u002D' in SpellTable[0][i]:
      #print(SpellTable[0][i])                                                                      prints out every spell name
      if x != 0:
        spellIndex.append((i-x,x))
      x = 0
    x += 1
  spellIndex.append((len(SpellTable[0])-x, x))
  return spellIndex

# Extract level from spell name (number before \u2013 or -)
def extractLevelAndName(spellName):
  # I am just gonna be very lazy and do this here (potentially will update in the future for better code)
  weird_spells = ["Revocation (Divine)", "Defer" ,"Revocation (Elemental)"]
  for spell in weird_spells:
    if spell in spellName and ('\u2013' in spellName or '\u002D' in spellName or '-' in spellName):
      if(spell is "Defer"):
        global DeferNumberCounter
        if DeferNumberCounter == 0:
          DeferNumberCounter = 1
          return "1-16", "Defer (Elemental)"
        else:
          return "1-16", "Defer (Divine)"    
      else:   
        return "1-16", spell

  
  if '\u2013' in spellName:
    parts = spellName.split('\u2013', 1)
  elif '\u002D' in spellName or '-' in spellName:
    parts = spellName.split('-', 1)
  else:
    return None, spellName.strip()
  

  if len(parts) == 2:
    level_str = parts[0].strip()
    spell_name = parts[1].strip()
    try:
      level = int(level_str)
      return level, spell_name
    except ValueError:
      return None, spellName.strip()
  
  return None, spellName.strip()

# Clean up spell group names
def cleanSpellGroupName(spellGroupName):
  if spellGroupName == "Psionic Abilities":
    return "Psionic"
  elif spellGroupName == "Elemental magics":
    return "Elemental"
  elif spellGroupName == "Divine magics":
    return "Divine"
  elif "Abilities" in spellGroupName:
    # Remove the word "Abilities" (and any extra spaces)
    return spellGroupName.replace("Abilities", "").strip()
  return spellGroupName

# gets the spells for that spell table as a dictionary
def getSpellDicts(SpellTable, spellGroupName=None):
  spellListDict = {}
  for i in getSpellIndexes(SpellTable):
    keyL = SpellTable[0][i[0]+1:i[0]+i[1]]
    valL = SpellTable[1][i[0]+1:i[0]+i[1]]
    spellDict = {}
    for j in range(len(keyL)):
      spellDict[keyL[j]] = valL[j]
    
    # Extract level and clean name from the spell title
    full_spell_name = SpellTable[0][i[0]]
    level, clean_spell_name = extractLevelAndName(full_spell_name)
    global spellcount
    spellcount += 1
    
    # Add level to the spell dictionary if extracted successfully
    if level is not None:
      spellDict['Level'] = level
    
    # Add cleaned spell group name to the spell dictionary
    if spellGroupName:
      spellDict['Spell Group'] = cleanSpellGroupName(spellGroupName)
    
    # Use the clean spell name (without level prefix) as the key
    spell_key = clean_spell_name if clean_spell_name else full_spell_name
    spellListDict[spell_key] = spellDict
    
  return spellListDict

# This is the main function that is taking in a table and outputing a fully cleaned and organized spell table as a dictionary. Ex: SpellTable['Concern']
def organizeTable(docTable, spellGroupName=None):
  colLen = len(list(docTable.columns))
  rowLen = len(list(docTable.rows))
  docTableCells = [docTable.column_cells(i) for i in range(colLen)]
  SpellTable = []
  for i in range(colLen):
    SpellTable.append([docTableCells[i][j].paragraphs for j in range(rowLen)])
    SpellTable[i] = list(map(collatPara,SpellTable[i]))
  SpellTable = cleanSpellTable(SpellTable)
  SpellTable = getSpellDicts(SpellTable, spellGroupName)
  return SpellTable

def getDocsSpells(doc):
  # creating a single dictionary to store all spells
  allSpells = {}
  tableing = [False, None]
  docSpellName = None
  currentSpellGroup = None
  
  for element in doc.element.body:
    # Finds the headers of the spell groups and signifies when to start collecting tables
    if isinstance(element, CT_P):
      if element.style == 'Heading5':
        tableing[0] = True
        currentSpellGroup = element.text
      if (element.style == 'Heading2' or element.style == 'Title') and not docSpellName:
        docSpellName = element.text
    # checks if element is a table, then accesses it as a table (?!?!), then adds all spells directly to the main dictionary
    if isinstance(element, CT_Tbl) and tableing[0]:
      table = Table(element, doc)
      if len(table._cells) > 10:
        # Get spells from this table and add them directly to allSpells
        tableSpells = organizeTable(table, currentSpellGroup)
        allSpells.update(tableSpells)
  
  return allSpells, docSpellName

if __name__ == "__main__":
  from docx import Document
  from docx.table import Table
  from docx.oxml.text.paragraph import CT_P
  from docx.text.paragraph import Paragraph
  from docx.oxml.table import CT_Tbl
  import json
  import os
  #opening the document as a docx.Document object
  allSpells = {}
  for file in os.listdir('spell information\spellFolder'):
    docSpells, docName = getDocsSpells(Document("spell information\spellFolder\\"+file))
    # Add all spells from this document to the main dictionary
    # If there are duplicate spell names, they will be overwritten (you may want to handle this differently)
    allSpells.update(docSpells)
  
  # Now allSpells contains all spells with their spell group and level information embedded
  # Access pattern: allSpells['Concern']['Level'] and allSpells['Concern']['Spell Group']
  with open('spell information/pfc-elemental-divine-psionic-spells.json', 'w') as outfile:
    json.dump(allSpells, indent=4, fp=outfile)

    print(f"Total Spells Processed: {spellcount}")