#Created by: KLSantos, NL-SRSA
#For weekly NL Open Signal DL Speed initial results/analysis.

import numpy as np
import pandas as pd
import os
from datetime import datetime
from matplotlib import pyplot as plt
import csv
import seaborn as sns
sns.set()

#plt.style.use('ggplot')
dtnow = datetime.now().strftime('%d%m%Y%H')
count = 0

print('*** Open Signal Stats Processor***')
xl = input('Enter Excel Raw File: ')
print('-' * 60)
df1 = pd.read_excel(xl)
xdate = df1['Day of Report End Date'].drop_duplicates().dt.strftime('%b-%d')
latestdate = xdate.iloc[-1]
#print(len(xdate))

filename = f'Open Signal Results ao {latestdate}'

with open(f'{filename}.csv','w') as output:
    fieldnames = ['Province','G-Lead', '3-Weeks Degraded','Abrupt Decrease (>5mbps)']   #create field names/header of the table
    openwrite = csv.DictWriter(output, fieldnames = fieldnames, delimiter = ',')    #write csv file with new delimiter
    openwrite.writeheader()

    while count < 22:
        for province in df1['Location']:
            flag = 0
            lead = ''
            deg = ''
            abr = ''

            provdf = df1.loc[df1['Location'] == province,['Day of Report End Date','Network Name Mapped','Download Mean']]

            sdf = provdf.loc[provdf['Network Name Mapped'] == 'Smart',['Day of Report End Date','Download Mean']]
            ySmart = sdf.groupby('Day of Report End Date').mean()
            slatest1 = ySmart['Download Mean'].iloc[-3]
            slatest2 = ySmart['Download Mean'].iloc[-2]
            slatest3 = ySmart['Download Mean'].iloc[-1]
            difflast = slatest3 - slatest2


            gdf = provdf.loc[provdf['Network Name Mapped'] == 'Globe',['Day of Report End Date','Download Mean']]
            yGlobe = gdf.groupby('Day of Report End Date').mean()
            glatest = yGlobe['Download Mean'].iloc[-1]

            if glatest > slatest3:
                print(f'{province} is G-Lead as of Week {len(xdate)}')
                lead = 'Yes'
                flag += 1
            #else:
                #print(f'{province} is S-Lead as of Week {len(xdate)}')

            if slatest3 < slatest2 and slatest2 < slatest1:
                print(f'{province} is 3 weeks degraded as of Week {len(xdate)}')
                deg = 'Yes'
                flag += 1

            if difflast < -5:
                print(f'{province} has an abrupt decrease from previous week')
                abr = 'Yes'
                flag += 1

            openwrite.writerow({'Province':province,'G-Lead':lead, '3-Weeks Degraded':deg,'Abrupt Decrease (>5mbps)':abr})

            plt.title(f'NL Open Signal DL Speed ({province})')
            plt.xlabel('Date')
            plt.ylabel('Download Speed')
            plt.plot(xdate, ySmart, color = 'g',marker = 'o', linewidth = 3, label = 'Smart')
            plt.plot(xdate, yGlobe, color = 'b', marker = 'o',linewidth = 3, label = 'Globe')
            plt.xticks(ticks =xdate, labels = xdate)
            plt.legend() #list order is plot return order
            plt.grid(True)  #enable gridlines
            #plt.tight_layout()  #improve padding?

            if flag > 0:
                plt.savefig(f'{province}_opensignal_{latestdate}')
                #plt.show()
            plt.clf()

            count += 1

            if count == 22:
                break

out = pd.read_csv(f'{filename}.csv')
out.to_excel(f'{filename}.xlsx', index = False)
