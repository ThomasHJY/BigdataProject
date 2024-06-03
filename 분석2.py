# -*- coding: utf-8 -*-

import requests
import json
import pandas

url06 = 'https://api.neople.co.kr/df/servers/'
url06_2 = '/characters/'
url06_3 = '/equip/equipment?'
apikey = 'apikey=비공개(네오플 api 참조)'

excelDf = pandas.read_excel('.\\list3.xlsx')

initialUrl = url06 + excelDf['serverId'][0] + url06_2 + excelDf['characterId'][0] + url06_3 + apikey
initialJson = requests.get(initialUrl)
initialDict = json.loads(initialJson.text)
 
weapon = []
jacket = []
shoulder = []
pants = []
shoes = []
waist = []
amulet = []
wrist = []
ring = []
support = []
magic_ston = []
support_weapon = []
earing = []

for i in range(0, 14):
    if i == 1:
        continue
    
    isCustom = False
    if 'customOption' in initialDict['equipment'][i]:
        isCustom = True
    if i == 0:
        if(isCustom):
            weapon.append(initialDict['equipment'][i]['itemName'])
            weapon.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            weapon.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            weapon.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            weapon.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            weapon.append(initialDict['equipment'][i]['itemName'])
            weapon.append('x')
            weapon.append('x')
            weapon.append('x')
            weapon.append('x')
    elif i == 2:
        if(isCustom):
            jacket.append(initialDict['equipment'][i]['itemName'])
            jacket.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            jacket.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            jacket.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            jacket.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            jacket.append(initialDict['equipment'][i]['itemName'])
            jacket.append(initialDict['equipment'][i]['fixedOption']['explain'])
            jacket.append('x')
            jacket.append('x')
            jacket.append('x')
    elif i == 3:
        if(isCustom):
            shoulder.append(initialDict['equipment'][i]['itemName'])
            shoulder.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            shoulder.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            shoulder.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            shoulder.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            shoulder.append(initialDict['equipment'][i]['itemName'])
            shoulder.append(initialDict['equipment'][i]['fixedOption']['explain'])
            shoulder.append('x')
            shoulder.append('x')
            shoulder.append('x')
    elif i == 4:
        if(isCustom):
            pants.append(initialDict['equipment'][i]['itemName'])
            pants.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            pants.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            pants.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            pants.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            pants.append(initialDict['equipment'][i]['itemName'])
            pants.append(initialDict['equipment'][i]['fixedOption']['explain'])
            pants.append('x')
            pants.append('x')
            pants.append('x')
    elif i == 5:
        if(isCustom):
            shoes.append(initialDict['equipment'][i]['itemName'])
            shoes.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            shoes.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            shoes.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            shoes.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            shoes.append(initialDict['equipment'][i]['itemName'])
            shoes.append(initialDict['equipment'][i]['fixedOption']['explain'])
            shoes.append('x')
            shoes.append('x')
            shoes.append('x')
    elif i == 6:
        if(isCustom):
            waist.append(initialDict['equipment'][i]['itemName'])
            waist.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            waist.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            waist.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            waist.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            waist.append(initialDict['equipment'][i]['itemName'])
            waist.append(initialDict['equipment'][i]['fixedOption']['explain'])
            waist.append('x')
            waist.append('x')
            waist.append('x')
    elif i == 7:
        if(isCustom):
            amulet.append(initialDict['equipment'][i]['itemName'])
            amulet.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            amulet.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            amulet.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            amulet.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            amulet.append(initialDict['equipment'][i]['itemName'])
            amulet.append(initialDict['equipment'][i]['fixedOption']['explain'])
            amulet.append('x')
            amulet.append('x')
            amulet.append('x')
    elif i == 8:
        if(isCustom):
            wrist.append(initialDict['equipment'][i]['itemName'])
            wrist.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            wrist.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            wrist.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            wrist.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            wrist.append(initialDict['equipment'][i]['itemName'])
            wrist.append(initialDict['equipment'][i]['fixedOption']['explain'])
            wrist.append('x')
            wrist.append('x')
            wrist.append('x')
    elif i == 9:
        if(isCustom):
            ring.append(initialDict['equipment'][i]['itemName'])
            ring.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            ring.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            ring.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            ring.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            ring.append(initialDict['equipment'][i]['itemName'])
            ring.append(initialDict['equipment'][i]['fixedOption']['explain'])
            ring.append('x')
            ring.append('x')
            ring.append('x')
    elif i == 10:
        if(isCustom):
            support.append(initialDict['equipment'][i]['itemName'])
            support.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            support.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            support.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            support.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            support.append(initialDict['equipment'][i]['itemName'])
            support.append(initialDict['equipment'][i]['fixedOption']['explain'])
            support.append('x')
            support.append('x')
            support.append('x')
    elif i == 11:
        if(isCustom):
            magic_ston.append(initialDict['equipment'][i]['itemName'])
            magic_ston.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            magic_ston.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            magic_ston.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            magic_ston.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            magic_ston.append(initialDict['equipment'][i]['itemName'])
            magic_ston.append(initialDict['equipment'][i]['fixedOption']['explain'])
            magic_ston.append('x')
            magic_ston.append('x')
            magic_ston.append('x')
    elif i == 12:
        if initialDict['equipment'][i]['slotId'] == 'EARRING':
            support_weapon.append('x')
            support_weapon.append('x')
            support_weapon.append('x')
            support_weapon.append('x')
            support_weapon.append('x')
            if(isCustom):
                earing.append(initialDict['equipment'][i]['itemName'])
                earing.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
                earing.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
                earing.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
                earing.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
                break
            else:
                earing.append(initialDict['equipment'][i]['itemName'])
                earing.append(initialDict['equipment'][i]['fixedOption']['explain'])
                earing.append('x')
                earing.append('x')
                earing.append('x')
                break
        else:
            if(isCustom):
                support_weapon.append(initialDict['equipment'][i]['itemName'])
                support_weapon.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
                support_weapon.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
                support_weapon.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
                support_weapon.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
            else:
                support_weapon.append(initialDict['equipment'][i]['itemName'])
                support_weapon.append(initialDict['equipment'][i]['fixedOption']['explain'])
                support_weapon.append('x')
                support_weapon.append('x')
                support_weapon.append('x')
    else:
        if(isCustom):
            earing.append(initialDict['equipment'][i]['itemName'])
            earing.append(initialDict['equipment'][i]['customOption']['options'][0]['explain'])
            earing.append(initialDict['equipment'][i]['customOption']['options'][1]['explain'])
            earing.append(initialDict['equipment'][i]['customOption']['options'][2]['explain'])
            earing.append(initialDict['equipment'][i]['customOption']['options'][3]['explain'])
        else:
            earing.append(initialDict['equipment'][i]['itemName'])
            earing.append(initialDict['equipment'][i]['fixedOption']['explain'])
            earing.append('x')
            earing.append('x')
            earing.append('x')
         
data = [[
            initialDict['characterId'], initialDict['characterName'], initialDict['jobName'], initialDict['jobGrowName'],
            weapon[0], weapon[1], weapon[2], weapon[3], weapon[4],
            jacket[0], jacket[1], jacket[2], jacket[3], jacket[4],
            shoulder[0], shoulder[1], shoulder[2], shoulder[3], shoulder[4],
            pants[0], pants[1], pants[2], pants[3], pants[4],
            shoes[0], shoes[1], shoes[2], shoes[3], shoes[4],
            waist[0], waist[1], waist[2], waist[3], waist[4],
            amulet[0], amulet[1], amulet[2], amulet[3], amulet[4],
            wrist[0], wrist[1], wrist[2], wrist[3], wrist[4],
            ring[0], ring[1], ring[2], ring[3], ring[4],
            support[0], support[1], support[2], support[3], support[4],
            magic_ston[0], magic_ston[1], magic_ston[2], magic_ston[3], magic_ston[4],
            support_weapon[0], support_weapon[1], support_weapon[2], support_weapon[3], support_weapon[4], 
            earing[0], earing[1], earing[2], earing[3], earing[4]
        ]]
df = pandas.DataFrame(data, columns = [
                                       'characterId', 'characterName', 'jobName', 'jobGrowName',
                                       'weapon', 'weaponOption1', 'weaponOption2', 'weaponOption3', 'weaponOption4',
                                       'jacket', 'jacketOption1', 'jacketOption2', 'jacketOption3', 'jacketOption4',
                                       'shoulder', 'shoulderOption1', 'shoulderOption2', 'shoulderOption3', 'shoulderOption4',
                                       'pants', 'pantsOption1', 'pantsOption2', 'pantsOption3', 'pantsOption4',
                                       'shoes', 'shoesOption1', 'shoesOption2', 'shoesOption3', 'shoesOption4',
                                       'waist', 'waistOption1', 'waistOption2', 'waistOption3', 'waistOption4',
                                       'amulet', 'amuletOption1', 'amuletOption2', 'amuletOption3', 'amuletOption4',
                                       'wrist', 'wristOption1', 'wristOption2', 'wristOption3', 'wristOption4',
                                       'ring', 'ringOption1', 'ringOption2', 'ringOption3', 'ringOption4',
                                       'support', 'supportOption1', 'supportOption2', 'supportOption3', 'supportOption4',
                                       'magic_ston', 'magic_stonOption1', 'magic_stonOption2', 'magic_stonOption3', 'magic_stonOption4',
                                       'support_weapon', 'support_weaponOption1', 'support_weaponOption2', 'support_weaponOption3', 'support_weaponOption4', 
                                       'earing', 'earingOption1', 'earingOption2', 'earingOption3', 'earingOption4'
                                       ])
countNum = 1
# countNum = 0

for i in excelDf.index:
    if(i == 0):   
        continue
    if(i % 90 != 1):
        continue
    tempUrl = url06 + excelDf['serverId'][i] + url06_2 + excelDf['characterId'][i] + url06_3 + apikey
    tempJson = requests.get(tempUrl)
    tempDict = json.loads(tempJson.text)
    weapon = []
    jacket = []
    shoulder = []
    pants = []
    shoes = []
    waist = []
    amulet = []
    wrist = []
    ring = []
    support = []
    magic_ston = []
    support_weapon = []
    earing = []
    noItem = False
    equipNum = len(tempDict['equipment']) - 1
    if(equipNum < 12):
        noItem = True
        
    for j in range(0, 14):
        if(noItem):
            break
        if j == 1:
            continue
        isCustom = False
        if 'customOption' in tempDict['equipment'][j]:
            isCustom = True

        if j == 0:
            if(isCustom):
                weapon.append(tempDict['equipment'][j]['itemName'])
                weapon.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                weapon.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                weapon.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                weapon.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'WEAPON':
                weapon.append(tempDict['equipment'][j]['itemName'])
                weapon.append('x')
                weapon.append('x')
                weapon.append('x')
                weapon.append('x')
            else:
                noItem = True
                break
        elif j == 2:
            if(isCustom):
                jacket.append(tempDict['equipment'][j]['itemName'])
                jacket.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                jacket.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                jacket.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                jacket.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'JACKET':
                if 'fixedOption' in tempDict['equipment'][j]:
                    jacket.append(tempDict['equipment'][j]['itemName'])
                    jacket.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    jacket.append('x')
                    jacket.append('x')
                    jacket.append('x')
                else:
                    jacket.append(tempDict['equipment'][j]['itemName'])
                    jacket.append('x')
                    jacket.append('x')
                    jacket.append('x')
                    jacket.append('x')
            else:
                noItem = True
                break
        elif j == 3:
            if(isCustom):
                shoulder.append(tempDict['equipment'][j]['itemName'])
                shoulder.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                shoulder.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                shoulder.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                shoulder.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'SHOULDER':
                if 'fixedOption' in tempDict['equipment'][j]:
                    shoulder.append(tempDict['equipment'][j]['itemName'])
                    shoulder.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    shoulder.append('x')
                    shoulder.append('x')
                    shoulder.append('x')
                else:
                    shoulder.append(tempDict['equipment'][j]['itemName'])
                    shoulder.append('x')
                    shoulder.append('x')
                    shoulder.append('x')
                    shoulder.append('x') 
            else:
                noItem = True
                break
        elif j == 4:
            if(isCustom):
                pants.append(tempDict['equipment'][j]['itemName'])
                pants.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                pants.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                pants.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                pants.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'PANTS':
                if 'fixedOption' in tempDict['equipment'][j]:
                    pants.append(tempDict['equipment'][j]['itemName'])
                    pants.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    pants.append('x')
                    pants.append('x')
                    pants.append('x')
                else:
                    pants.append(tempDict['equipment'][j]['itemName'])
                    pants.append('x')
                    pants.append('x')
                    pants.append('x')
                    pants.append('x')
            else:
                noItem = True
                break
        elif j == 5:
            if(isCustom):
                shoes.append(tempDict['equipment'][j]['itemName'])
                shoes.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                shoes.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                shoes.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                shoes.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'SHOES':
                if 'fixedOption' in tempDict['equipment'][j]:
                    shoes.append(tempDict['equipment'][j]['itemName'])
                    shoes.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    shoes.append('x')
                    shoes.append('x')
                    shoes.append('x')
                else:
                    shoes.append(tempDict['equipment'][j]['itemName'])
                    shoes.append('x')
                    shoes.append('x')
                    shoes.append('x')
                    shoes.append('x')
            else:
                noItem = True
                break
        elif j == 6:
            if(isCustom):
                waist.append(tempDict['equipment'][j]['itemName'])
                waist.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                waist.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                waist.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                waist.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'WAIST':
                if 'fixedOption' in tempDict['equipment'][j]:
                    waist.append(tempDict['equipment'][j]['itemName'])
                    waist.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    waist.append('x')
                    waist.append('x')
                    waist.append('x')
                else:
                    waist.append(tempDict['equipment'][j]['itemName'])
                    waist.append('x')
                    waist.append('x')
                    waist.append('x')
                    waist.append('x')
            else:
                noItem = True
                break
        elif j == 7:
            if(isCustom):
                amulet.append(tempDict['equipment'][j]['itemName'])
                amulet.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                amulet.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                amulet.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                amulet.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'AMULET':
                if 'fixedOption' in tempDict['equipment'][j]:
                    amulet.append(tempDict['equipment'][j]['itemName'])
                    amulet.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    amulet.append('x')
                    amulet.append('x')
                    amulet.append('x')
                else:
                    amulet.append(tempDict['equipment'][j]['itemName'])
                    amulet.append('x')
                    amulet.append('x')
                    amulet.append('x')
                    amulet.append('x')
            else:
                noItem = True
                break
        elif j == 8:
            if(isCustom):
                wrist.append(tempDict['equipment'][j]['itemName'])
                wrist.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                wrist.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                wrist.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                wrist.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'WRIST':
                if 'fixedOption' in tempDict['equipment'][j]:
                    wrist.append(tempDict['equipment'][j]['itemName'])
                    wrist.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    wrist.append('x')
                    wrist.append('x')
                    wrist.append('x')
                else:
                    wrist.append(tempDict['equipment'][j]['itemName'])
                    wrist.append('x')
                    wrist.append('x')
                    wrist.append('x')
                    wrist.append('x')
            else:
                noItem = True
                break
        elif j == 9:
            if(isCustom):
                ring.append(tempDict['equipment'][j]['itemName'])
                ring.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                ring.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                ring.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                ring.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'RING':
                if 'fixedOption' in tempDict['equipment'][j]:
                    ring.append(tempDict['equipment'][j]['itemName'])
                    ring.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    ring.append('x')
                    ring.append('x')
                    ring.append('x')
                else:
                    ring.append(tempDict['equipment'][j]['itemName'])
                    ring.append('x')
                    ring.append('x')
                    ring.append('x')
                    ring.append('x')
            else:
                noItem = True
                break
        elif j == 10:
            if(isCustom):
                support.append(tempDict['equipment'][j]['itemName'])
                support.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                support.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                support.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                support.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'SUPPORT':
                if 'fixedOption' in tempDict['equipment'][j]:
                    support.append(tempDict['equipment'][j]['itemName'])
                    support.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    support.append('x')
                    support.append('x')
                    support.append('x')
                else:
                    support.append(tempDict['equipment'][j]['itemName'])
                    support.append('x')
                    support.append('x')
                    support.append('x')
                    support.append('x')
            else:
                noItem = True
                break
        elif j == 11:
            if(isCustom):
                magic_ston.append(tempDict['equipment'][j]['itemName'])
                magic_ston.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                magic_ston.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                magic_ston.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                magic_ston.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'MAGIC_STON':
                if 'fixedOption' in tempDict['equipment'][j]:
                    magic_ston.append(tempDict['equipment'][j]['itemName'])
                    magic_ston.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    magic_ston.append('x')
                    magic_ston.append('x')
                    magic_ston.append('x')
                else:
                    magic_ston.append(tempDict['equipment'][j]['itemName'])
                    magic_ston.append(tempDict['equipment'][j]['fixedOption']['explain'])
                    magic_ston.append('x')
                    magic_ston.append('x')
                    magic_ston.append('x')
            else:
                noItem = True
                break
        elif j == 12:
            if tempDict['equipment'][j]['slotId'] == 'EARRING':
                support_weapon.append('x')
                support_weapon.append('x')
                support_weapon.append('x')
                support_weapon.append('x')
                support_weapon.append('x')
                if(isCustom):
                    earing.append(tempDict['equipment'][j]['itemName'])
                    earing.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                    earing.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                    earing.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                    earing.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
                    break
                elif 'itemName' in tempDict['equipment'][j]:
                    if 'fixedOption' in tempDict['equipment'][j]:
                        earing.append(tempDict['equipment'][j]['itemName'])
                        earing.append(tempDict['equipment'][j]['fixedOption']['explain'])
                        earing.append('x')
                        earing.append('x')
                        earing.append('x')
                    else:
                        earing.append(tempDict['equipment'][j]['itemName'])
                        earing.append('x')
                        earing.append('x')
                        earing.append('x')
                        earing.append('x')
                    break
                else:
                    noItem = True
                    break
            else:
                if(isCustom):
                    support_weapon.append(tempDict['equipment'][j]['itemName'])
                    support_weapon.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                    support_weapon.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                    support_weapon.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                    support_weapon.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
                elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'SUPPORT_WEAPON':
                    if 'fixedOption' in tempDict['equipment'][j]:
                        support_weapon.append(tempDict['equipment'][j]['itemName'])
                        support_weapon.append(tempDict['equipment'][j]['fixedOption']['explain'])
                        support_weapon.append('x')
                        support_weapon.append('x')
                        support_weapon.append('x')
                    else:
                        support_weapon.append(tempDict['equipment'][j]['itemName'])
                        support_weapon.append('x')
                        support_weapon.append('x')
                        support_weapon.append('x')
                        support_weapon.append('x')
                else:
                    noItem = True
                    break
        else:
            if(isCustom):
                earing.append(tempDict['equipment'][j]['itemName'])
                earing.append(tempDict['equipment'][j]['customOption']['options'][0]['explain'])
                earing.append(tempDict['equipment'][j]['customOption']['options'][1]['explain'])
                earing.append(tempDict['equipment'][j]['customOption']['options'][2]['explain'])
                earing.append(tempDict['equipment'][j]['customOption']['options'][3]['explain'])
            elif 'itemName' in tempDict['equipment'][j] and tempDict['equipment'][j]['slotId'] == 'EARRING':
                    if 'fixedOption' in tempDict['equipment'][j]:
                        earing.append(tempDict['equipment'][j]['itemName'])
                        earing.append(tempDict['equipment'][j]['fixedOption']['explain'])
                        earing.append('x')
                        earing.append('x')
                        earing.append('x')
                    else:
                        earing.append(tempDict['equipment'][j]['itemName'])
                        earing.append('x')
                        earing.append('x')
                        earing.append('x')
                        earing.append('x')
            else:
                noItem = True
                break
    tempdata = []
    if(noItem):
        print('check ' + str(i))
        tempdata = [[
                        tempDict['characterId'], tempDict['characterName'], tempDict['jobName'], tempDict['jobGrowName'],
                        'x', 'x', 'x', 'x', 'x', # weapon
                        'x', 'x', 'x', 'x', 'x', # jacket
                        'x', 'x', 'x', 'x', 'x', # shoulder
                        'x', 'x', 'x', 'x', 'x', # pants
                        'x', 'x', 'x', 'x', 'x', # shoes
                        'x', 'x', 'x', 'x', 'x', # waist
                        'x', 'x', 'x', 'x', 'x', # amulet
                        'x', 'x', 'x', 'x', 'x', # wrist
                        'x', 'x', 'x', 'x', 'x', # ring
                        'x', 'x', 'x', 'x', 'x', # support
                        'x', 'x', 'x', 'x', 'x', # magic_ston
                        'x', 'x', 'x', 'x', 'x', # support_weapon
                        'x', 'x', 'x', 'x', 'x' # earring
                    ]]
    else:
        print(i)
        tempdata = [[
                        tempDict['characterId'], tempDict['characterName'], tempDict['jobName'], tempDict['jobGrowName'],
                        weapon[0], weapon[1], weapon[2], weapon[3], weapon[4],
                        jacket[0], jacket[1], jacket[2], jacket[3], jacket[4],
                        shoulder[0], shoulder[1], shoulder[2], shoulder[3], shoulder[4],
                        pants[0], pants[1], pants[2], pants[3], pants[4],
                        shoes[0], shoes[1], shoes[2], shoes[3], shoes[4],
                        waist[0], waist[1], waist[2], waist[3], waist[4],
                        amulet[0], amulet[1], amulet[2], amulet[3], amulet[4],
                        wrist[0], wrist[1], wrist[2], wrist[3], wrist[4],
                        ring[0], ring[1], ring[2], ring[3], ring[4],
                        support[0], support[1], support[2], support[3], support[4],
                        magic_ston[0], magic_ston[1], magic_ston[2], magic_ston[3], magic_ston[4],
                        support_weapon[0], support_weapon[1], support_weapon[2], support_weapon[3], support_weapon[4], 
                        earing[0], earing[1], earing[2], earing[3], earing[4]
                    ]]
    tempdf = pandas.DataFrame(tempdata, columns = [
                                                    'characterId', 'characterName', 'jobName', 'jobGrowName',
                                                    'weapon', 'weaponOption1', 'weaponOption2', 'weaponOption3', 'weaponOption4',
                                                    'jacket', 'jacketOption1', 'jacketOption2', 'jacketOption3', 'jacketOption4',
                                                    'shoulder', 'shoulderOption1', 'shoulderOption2', 'shoulderOption3', 'shoulderOption4',
                                                    'pants', 'pantsOption1', 'pantsOption2', 'pantsOption3', 'pantsOption4',
                                                    'shoes', 'shoesOption1', 'shoesOption2', 'shoesOption3', 'shoesOption4',
                                                    'waist', 'waistOption1', 'waistOption2', 'waistOption3', 'waistOption4',
                                                    'amulet', 'amuletOption1', 'amuletOption2', 'amuletOption3', 'amuletOption4',
                                                    'wrist', 'wristOption1', 'wristOption2', 'wristOption3', 'wristOption4',
                                                    'ring', 'ringOption1', 'ringOption2', 'ringOption3', 'ringOption4',
                                                    'support', 'supportOption1', 'supportOption2', 'supportOption3', 'supportOption4',
                                                    'magic_ston', 'magic_stonOption1', 'magic_stonOption2', 'magic_stonOption3', 'magic_stonOption4',
                                                    'support_weapon', 'support_weaponOption1', 'support_weaponOption2', 'support_weaponOption3', 'support_weaponOption4',
                                                    'earing', 'earingOption1', 'earingOption2', 'earingOption3', 'earingOption4'
                                                  ])
    result = pandas.concat([df, tempdf])
    df = result
    countNum += 1
    if(countNum == 200):
        print('countNum: ' + str(i))
        break
#    
df['weaponOption1'] = df['weaponOption1'].str.replace('\n', '&')
df['weaponOption2'] = df['weaponOption2'].str.replace('\n', '&')
df['weaponOption3'] = df['weaponOption3'].str.replace('\n', '&')
df['weaponOption4'] = df['weaponOption4'].str.replace('\n', '&')
#    
df['jacketOption1'] = df['jacketOption1'].str.replace('\n', '&')
df['jacketOption2'] = df['jacketOption2'].str.replace('\n', '&')
df['jacketOption3'] = df['jacketOption3'].str.replace('\n', '&')
df['jacketOption4'] = df['jacketOption4'].str.replace('\n', '&')
# Ӹ    
df['shoulderOption1'] = df['shoulderOption1'].str.replace('\n', '&')
df['shoulderOption2'] = df['shoulderOption2'].str.replace('\n', '&')
df['shoulderOption3'] = df['shoulderOption3'].str.replace('\n', '&')
df['shoulderOption4'] = df['shoulderOption4'].str.replace('\n', '&')
#    
df['pantsOption1'] = df['pantsOption1'].str.replace('\n', '&')
df['pantsOption2'] = df['pantsOption2'].str.replace('\n', '&')
df['pantsOption3'] = df['pantsOption3'].str.replace('\n', '&')
df['pantsOption4'] = df['pantsOption4'].str.replace('\n', '&')
# Ź 
df['shoesOption1'] = df['shoesOption1'].str.replace('\n', '&')
df['shoesOption2'] = df['shoesOption2'].str.replace('\n', '&')
df['shoesOption3'] = df['shoesOption3'].str.replace('\n', '&')
df['shoesOption4'] = df['shoesOption4'].str.replace('\n', '&')
#  Ʈ
df['waistOption1'] = df['waistOption1'].str.replace('\n', '&')
df['waistOption2'] = df['waistOption2'].str.replace('\n', '&')
df['waistOption3'] = df['waistOption3'].str.replace('\n', '&')
df['waistOption4'] = df['waistOption4'].str.replace('\n', '&')
#     
df['amuletOption1'] = df['amuletOption1'].str.replace('\n', '&')
df['amuletOption2'] = df['amuletOption2'].str.replace('\n', '&')
df['amuletOption3'] = df['amuletOption3'].str.replace('\n', '&')
df['amuletOption4'] = df['amuletOption4'].str.replace('\n', '&')
#    
df['wristOption1'] = df['wristOption1'].str.replace('\n', '&')
df['wristOption2'] = df['wristOption2'].str.replace('\n', '&')
df['wristOption3'] = df['wristOption3'].str.replace('\n', '&')
df['wristOption4'] = df['wristOption4'].str.replace('\n', '&')
#    
df['ringOption1'] = df['ringOption1'].str.replace('\n', '&')
df['ringOption2'] = df['ringOption2'].str.replace('\n', '&')
df['ringOption3'] = df['ringOption3'].str.replace('\n', '&')
df['ringOption4'] = df['ringOption4'].str.replace('\n', '&')
#       
df['supportOption1'] = df['supportOption1'].str.replace('\n', '&')
df['supportOption2'] = df['supportOption2'].str.replace('\n', '&')
df['supportOption3'] = df['supportOption3'].str.replace('\n', '&')
df['supportOption4'] = df['supportOption4'].str.replace('\n', '&')
#      
df['magic_stonOption1'] = df['magic_stonOption1'].str.replace('\n', '&')
df['magic_stonOption2'] = df['magic_stonOption2'].str.replace('\n', '&')
df['magic_stonOption3'] = df['magic_stonOption3'].str.replace('\n', '&')
df['magic_stonOption4'] = df['magic_stonOption4'].str.replace('\n', '&')
#        
df['support_weaponOption1'] = df['support_weaponOption1'].str.replace('\n', '&')
df['support_weaponOption2'] = df['support_weaponOption2'].str.replace('\n', '&')
df['support_weaponOption3'] = df['support_weaponOption3'].str.replace('\n', '&')
df['support_weaponOption4'] = df['support_weaponOption4'].str.replace('\n', '&')
# Ͱ   
df['earingOption1'] = df['earingOption1'].str.replace('\n', '&')
df['earingOption2'] = df['earingOption2'].str.replace('\n', '&')
df['earingOption3'] = df['earingOption3'].str.replace('\n', '&')
df['earingOption4'] = df['earingOption4'].str.replace('\n', '&')

df.to_excel('equip.xlsx')
# df.to_excel('equip2.xlsx')
# df.to_excel('equip3.xlsx')
# df.to_excel('equip4.xlsx')
# df.to_excel('equip5.xlsx')