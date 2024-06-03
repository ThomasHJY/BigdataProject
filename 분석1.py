# -*- coding: utf-8 -*-
# °°Àº ¸í¼ºÀº ÃÖ´ë 200¸í±îÁö¸¸ °Ë»ö °¡´É

import requests
import json
import pandas

url16 = 'https://api.neople.co.kr/df/servers/all/characters-fame?'
apikey = 'apikey=ºñ°ø°³(³×¿ÀÇÃ api ÂüÁ¶)'

# maxFameUrl = url16 + 'limit=1&' + apikey
# maxFameUrl = url16 + 'maxFame=65000&limit=1&' + apikey
# maxFameUrl = url16 + 'maxFame=63000&limit=1&' + apikey
# maxFameUrl = url16 + 'maxFame=61000&limit=1&' + apikey
maxFameUrl = url16 + 'maxFame=59000&limit=1&' + apikey
maxFameJson = requests.get(maxFameUrl)
maxFameDict = json.loads(maxFameJson.text)

maxFame = maxFameDict['rows'][0]['fame']
# minFame = 65000 # ÄöÅÒ Ä«Áö³ë(2024-05-07) 57501
# minFame = 63000 # ÄöÅÒ Ä«Áö³ë(2024-05-07) 57501
# minFame = 61000 # ÄöÅÒ Ä«Áö³ë(2024-05-07) 57501
# minFame = 59000 # ÄöÅÒ Ä«Áö³ë(2024-05-07) 57501
minFame = 57501 # ÄöÅÒ Ä«Áö³ë(2024-05-07) 57501
overlapId = maxFameDict['rows'][0]['characterId']
over200 = False

data = [[maxFame, maxFameDict['rows'][0]['serverId'], overlapId]]
df = pandas.DataFrame(data, columns = ['fame', 'serverId', 'characterId'])
num = 1
while maxFame >= minFame:
    tempUrl = url16 + 'maxFame=' + str(maxFame) + '&limit=200&' + apikey
    tempJson = requests.get(tempUrl)
    tempDict = json.loads(tempJson.text)
    isOverLap = True
    start = False
    lastIndex = len(tempDict['rows']) - 1
    
    for player in tempDict['rows']:
        if((player['characterId'] == overlapId) or over200):
            isOverLap = False
            over200 = False
            start = True
            continue
        if(not(start)):
            if(isOverLap):
                continue
            if(player['fame'] < minFame):
                break
        templist = [[player['fame'], player['serverId'], player['characterId']]]
        tempdf = pandas.DataFrame(templist, columns = ['fame', 'serverId', 'characterId'])
        result = pandas.concat([df, tempdf])
        df = result
        num += 1
        print(num)
    maxFame = tempDict['rows'][lastIndex]['fame']
    overlapId = tempDict['rows'][lastIndex]['characterId']
    if(maxFame == tempDict['rows'][0]['fame']):
        maxFame -= 1
        overLapId = ''
        over200 = True
# df.to_excel('list.xlsx')
# df.to_excel('list2.xlsx')
# df.to_excel('list3.xlsx')
# df.to_excel('list4.xlsx')  
df.to_excel('list5.xlsx')      