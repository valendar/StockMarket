import arrow


def lss(csepstring):
    return csepstring.rstrip(',\r\n').split(',')


realbiglist = []
biglist = []
newlist = []
today_date = arrow.now().format('DD/MM/YYYY') #  Topshare likes this format
namedate = arrow.now().format('DD-MM-YYYY')   #  No / in filenames
date = today_date
print(today_date)
with open('C:/Users/Val/Downloads/Watchlist.csv', newline='') as readobject:
    sl = [line for line in readobject]
    del sl[0]
    del sl[-1]
    # print(sl, sep='\n')
    # print(*sl, sep='\n')
    # exit(0)
    for item in sl:
        var = lss(item)
        # print('Item', item)  # , exit(0)
        # print('Var', var)
        # print('Items', today_date, var[1], var[4])
        biglist = [today_date, var[1], var[4]]
        realbiglist.append(biglist)
print('Closing prices')
print(*realbiglist, sep='\n')

with open('C:/Users/Val/Downloads/EODPricesTopShare' + namedate + '.csv', 'w') as writefile:
    for item in realbiglist:
        make_all_strings = [str(entry) for entry in item]
        # print(item[0], testforpos)
        r = ','.join(make_all_strings)
        writefile.write(r + '\n')

