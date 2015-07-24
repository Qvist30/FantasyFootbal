'''
Copyright (C) 2014 Thomas Kennedy

    This program is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License along
    with this program; if not, write to the Free Software Foundation, Inc.,
    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'''


import numpy as np
import csv
import glob
import xlwt
import os
import operator
import re

    
def createPostionalRanking(position, starters):
    filename = 'outputfiles/' + position + '_output.csv'
    with open(filename, 'w') as output:
        writer = csv.writer(output, delimiter = ',', lineterminator='\n')
        writer.writerow(['Player', 'Team', 'Position', 'FFtoolbox', 'ESPN', 'Pts. Per Game', 'Total Pts.', 'Bye Week', 'Points over Bench', 'Points over Waiver'])
        rows_list = []
        with open('inputfiles/' + position + '_FFToolbox.csv') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=',')
            for row in reader:
                found = False
                ptsPerGame = float(row["PTS"])/16
                player = row["PLAYER"]
                team = row["NFL"]
                player = player.replace('Contract Year Player', '').replace('Recently Updated Outlook', '').replace('Injury', '').replace('News','').replace('Suprise','').replace('Surprise','').replace('INSIDER','').replace('EXPERT','').replace('Breakout','').replace('COMEBACK','').replace('?','').replace('STASH','')
                player = re.sub('\W+',' ', player)
                
                with open('inputfiles/' + position + '_ESPN.csv') as espnfile:
                    espnReader = csv.DictReader(espnfile, delimiter=',')
                    for espnRow in espnReader:
                        espnPtsPerGame = float(espnRow["PTS"])/16
                        espnPlayer = espnRow["PLAYER"]
                        espnPlayer = espnPlayer.replace('Contract Year Player', '').replace('Recently Updated Outlook', '').replace('Injury', '').replace('News','')
                        espnPlayer = re.sub('\W+',' ', espnPlayer)
                        if(espnPlayer.replace(' ', '').lower().strip().startswith(player.replace(' ', '').lower().strip()) | player.replace(' ', '').lower().strip().startswith(espnPlayer.replace(' ', '').lower().strip())):
                            rows_list.append([espnPlayer, row["NFL"], position, ptsPerGame*16, espnPtsPerGame *16,np.mean([espnPtsPerGame, ptsPerGame]), np.mean([espnPtsPerGame, ptsPerGame]) * 16, row["BYE"]])
                            found = True
                            break
                        
                if found == False:
                    print player, " NOT FOUND"
        numpArray = np.array(sorted(rows_list, key=operator.itemgetter(5), reverse=True))
        pptsPerGameBench = float(numpArray[starters, 5])
        pptsPerGameReplacement = float(numpArray[starters*2, 5])
        for x in numpArray:
            writer.writerow([x[0], x[1], x[2], x[3], x[4], x[5],x[6],x[7], float(x[5]) - pptsPerGameBench, float(x[5]) - pptsPerGameReplacement])
    output.close()
    return filename

if __name__ == '__main__':
    qbFile = createPostionalRanking('QB', 12)
    rbFile = createPostionalRanking('RB', 30)
    wrFile = createPostionalRanking('WR', 30)
    teFile = createPostionalRanking('TE', 12)
    defFile = createPostionalRanking('DEF', 12)
    kFile = createPostionalRanking('K', 12)
    qbArray = np.genfromtxt(qbFile, delimiter=',', dtype=[('player', 'S30'), ('team', 'S30'), ('position', 'S30'), ('FFtoolbox', float), ('ESPN', float), ('PPG', float), ('Pts', float), ('Bye', int), ('PAB', float), ('PAR', float)], skip_header=1)
    rbArray = np.genfromtxt(rbFile, delimiter=',', dtype=[('player', 'S30'), ('team', 'S30'), ('position', 'S30'), ('FFtoolbox', float), ('ESPN', float), ('PPG', float), ('Pts', float), ('Bye', int), ('PAB', float), ('PAR', float)], skip_header=1)
    wrArray = np.genfromtxt(wrFile, delimiter=',', dtype=[('player', 'S30'), ('team', 'S30'), ('position', 'S30'), ('FFtoolbox', float), ('ESPN', float), ('PPG', float), ('Pts', float), ('Bye', int), ('PAB', float), ('PAR', float)], skip_header=1)
    teArray = np.genfromtxt(teFile, delimiter=',', dtype=[('player', 'S30'), ('team', 'S30'), ('position', 'S30'), ('FFtoolbox', float), ('ESPN', float), ('PPG', float), ('Pts', float), ('Bye', int), ('PAB', float), ('PAR', float)], skip_header=1)
    defArray = np.genfromtxt(defFile, delimiter=',', dtype=[('player', 'S30'), ('team', 'S30'), ('position', 'S30'), ('FFtoolbox', float), ('ESPN', float), ('PPG', float), ('Pts', float), ('Bye', int), ('PAB', float), ('PAR', float)], skip_header=1)
    kArray = np.genfromtxt(kFile, delimiter=',', dtype=[('player', 'S30'), ('team', 'S30'), ('position', 'S30'), ('FFtoolbox', float), ('ESPN', float), ('PPG', float), ('Pts', float), ('Bye', int), ('PAB', float), ('PAR', float)], skip_header=1)

    totalArray = np.concatenate((qbArray, rbArray, wrArray, teArray, defArray, kArray))
    totalArray.sort(order='PAB')
    numpArray= totalArray[::-1]
    overallFile = 'outputfiles/overall_output.csv'
    with open(overallFile, 'w') as output:
        writer = csv.writer(output, delimiter = ',', lineterminator='\n')
        writer.writerow(['Player', 'Team', 'Position', 'FFtoolbox', 'ESPN', 'Pts. Per Game', 'Total Pts.', 'Bye Week', 'Points over Bench', 'Points over Waiver'])

        for x in numpArray:
            writer.writerow(x)

    wb = xlwt.Workbook()
    filenames= [qbFile, rbFile, wrFile, teFile, defFile, kFile, overallFile]
    for filename in filenames:
        (f_path, f_name) = os.path.split(filename)
        (f_short_name, f_extension) = os.path.splitext(f_name)
        ws = wb.add_sheet(str(f_short_name))
        spamReader = csv.reader(open(filename, 'rb'), delimiter=',',quotechar='"')
        row_count = 0
        for row in spamReader:
            for col in range(len(row)):
                ws.write(row_count,col,row[col])
            row_count +=1

    wb.save("outputfiles/rankings.xls")
 
