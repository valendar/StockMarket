import os
import csv
import newstockdefs as ns
import xlwings as xw
import time

# HGEPS10 = 5
# HGEPS5 = 5
# STAEPS10 = 75
# STAEPS5 = 75


DEFAULT = 3  # Most recent period to average ROE, HGEPS10 and HGEPS5
HGEPS10 = 3  # % default HGROWTH 10 for earnings per share
HGEPS5 = 3  # % default HGROWTH 5 for earnings per share

STAEPS10 = 70  # % default STAEGER 10 for earnings per share
STAEPS5 = 85  # % default STAEGER 5 for earnings per share
PER = 30  # Maximum PE ratio
DE = 50  # 50% default debt to equity ratio
IC = 2.0  # default interest cover
MCAP = 50  # default market capitalisation (millions)
ROE = 20  # % default return on equity

YEARSOFDATA = 10
HEADER1 = ['STOCK', 'HGEPS10', 'HGEPS5', 'STAEPS10', 'STAEPS5']
HEADER2 = ['MCAP', 'P/E ratio', 'ROE', 'Debt/Equity', 'Interest Cover']

LongHeader = 'STOCK, HGEPS10, HGEPS5, STAEPS10, STAEPS5, ROESTAG5, ROESTAG10, EQU/SHARE, MCAP, \
 P/E ratio, ROE, Debt/Equity, Interest Cover, PO Ratio, EPS, TARG, STRET, PRICE,PEG, RMVal, RM/Price\n'

# STOCK,HGEPS10,HGEPS5,STAEPS10,STAEPS5,ROESTAG10,ROESTAG5,EQU/SHARE,MCAP,PER,ROE,D/E,ICover,POR,EPS,TARG,STRET,PRICE,PEG,RMVal,RM/Price,,,,
# LongHeaderOLD = 'STOCK, HGEPS10, HGEPS5, STAEPS10, STAEPS5, ROESTAG5, ROESTAG10, EQU/SHARE, MCAP, \
#  P/E ratio, ROE, Debt/Equity, Interest Cover, PO Ratio, EPS, Price,TARG, STRET, MOS, IVALUE, IV/Price, ROC, RMIVal\n'


row_name_filter1 = ['Market Cap (m)', 'Avge Annual PE Ratio(%)', 'Return on Equity (%)', 'Debt/Equity (%)',
                    'Net Interest Cover', 'Payout Ratio (%)', 'Earnings (cents)',
                    'Shareholders Equity (m)', 'Shares Outstanding (m)', 'Last price']  # Earnings per share

# number of entries to average in key values in myDict
myDict2 = {9: 3, 8: 1, 11: 3, 12: 3, 10: 3, 13: 3, 14: 3}
# eg Market cap (5) is the last (current) Market cap.

stocknames = []  # list to store stock names in the ASX file data
stocklist = []  # list to store the data from each stock name in the screened file
screened_list = []
testlist1 = []

path1 = 'C:/Users/Val/Desktop/SharesStuff/INTERIM/Mar2018ASXData/'  # Main folder
# for initial sweep
path = 'C:/Users/Val/Desktop/SharesStuff/INTERIM/Mar2018ASXData/ASX-USE/'
# filtered list of stocknames for EXCEL to do its VBA
path2 = 'C:/Users/Val/Desktop/SharesStuff/INTERIM/Mar2018ASXData/ASX-USE-NEW/'

#  stocknames = next(os.walk(path))[2]  # directory of stocks in ASX-USE folder
passedlist1 = []
failedlist1 = []
# files that end up in ASX-USE-NEW folder and are further filtered into FirstFilter.csv
passedlist2 = []
failedlist2 = []


def lss(csepstring):
    return csepstring.rstrip(',\r\n').split(',')


def setUpData():
    cplist = ns.getCPList()  # returns list of closing prices from Commsec  Rename ASXClosingPrices-20180314.csv
    # Root, Directories, Files.  [2] = stocknames as list in ASX-USE directory
    stocknames = next(os.walk(path))[2]
    for stockname in stocknames:
        # open each stock one at a time in the ASX-USE file
        with open(path + stockname, newline='') as datafile:
            # print('Stock', stockname)
            datafile.readline()  # skip Company Historicals line
            # returns iterable object
            readCSV = csv.reader(datafile, delimiter=',')
            # put data rows from one stock into a list of lists.  Each data row is a csep list of strings
            zlist = [aline for aline in readCSV]
            # where closing price is put to calculate STRET and others
            zlist.append(['Last price', 'lp'])
            zlist.insert(0, [stockname])  # add the stockname.csv to the list
            # print(zlist[39][0], zlist[39][1])
        '''
        Following code uses sets to test if all items in the row_name_filter are present in the list of row titles.
        The sets intersection returns strings common to both sets and == tests if this result is identical to the row filter set
        '''

        # set turns row headings into a set for manipulation
        setstuff = set(row_name_filter1)
        # zlist [1:] is zlist minus the first (zeroeth) element.  [x] is the xth list in zlist{1:] and [0]
        # is the first item in the xth list.  That is, the name of each row of stock data. Eg, "Franking".
        # find all row titles for this stock
        yal = [zlist[1:][x][0] for x, y in enumerate(zlist[1:])]
        # print(yal)
        # sleep(.1)
        setrows = set(yal)
        # intersection returns what is common between the two sets
        result = setrows.intersection(setstuff)
        # print (zlist); exit(0)
        if result == setstuff:  # is True  meaning a stock has all the row names in the row name filter
            passedlist1.append(stockname)
        # print(stockname)
        else:  # is False
            failedlist1.append(stockname)
        # print('Passing ', stockname.rstrip('.csv'))
        # returns True and closing price
        retval = ns.hasLastPrice(stockname.rstrip('.csv'), cplist)
        print(zlist[39])
        # exit(0)
        # Has 10 YEARSOFDATA and has price data
        if len(zlist[1][1:]) == 10 and retval[0] is True:
            zlist[39][1] = retval[1]
            ns.writeListToFile(stockname, zlist[1:])
            # print(zlist[39])
            passedlist2.append(stockname)
        # add price to fundamental data
        else:
            failedlist2.append(stockname)
            # for stock in failedlist2:
            #     os.remove(path + stock)
        cpdict = dict(cplist)
        ns.insertCP(stockname, cpdict, zlist)
        # print(stockname)
    print(len(passedlist1), 'xxxxx', len(failedlist1))
    print('Passed List2', len(passedlist2))  # , passedlist2)
    print('Failed List2', len(failedlist2))  # , failedlist2)
    print(set(passedlist2).intersection(set(failedlist2)))
    print(len(passedlist2) + len(failedlist2))

    #  exit(0)
    '''SO FAR HAVE A FOLDER WITH ASX FILES WITH STOCKS 10 YEARS OF DATA AND A CLOSING PRICE
    RETURN TO WINDOWS AND RUN RUN 1VALVBA.xlsm ON ASX-USE-NEW.csv  TO GET THE JOHN PRICE DATA TO FURTHER PROCESS
    SAVE THIS FILE AS FirstFilter.csv
    '''


# BELOW COPIED FROM StockScreen.py

# NEXT CODE REMOVES RUBBISH DATA '#VALUE!' and -99990'from FirstFilter.csv and saves in SecondFilter.csv
#  Stocks with #DIV/0! and #NUM! errors in SecondFilter.csv should be analysed using the original csv file for that
#  stock.  It happens because they don't contain the data rows used in most other stocks.  For example, in banks or
# insurance companies.  Make a copy of SecondFilter.csv as SecondFilter.xlsx to facilitate searching.
def removeRubbish():
    # Repository for next stage in filtering
    with open(path1 + 'FirstFilter.csv', 'r') as firstreadobj:
        firstreadobj.readline()  # skip columns headers
        stock_data = [item for item in (firstreadobj.readlines())]
        stock_names = [lss(astring)[0] for astring in stock_data]
        # print(stock_names)
        for stock in stock_names:
            # print(stock)
            with open(path1 + 'SecondFilter.csv', 'w') as writeobj:
                # write one line which puts SecondFilter into appending mode
                writeobj.write(LongHeader)
                for line in stock_data:  #
                    linedata = ns.lss(line)
                    # print(linedata)
                    # exit(0)
                    # list expression  these are the data
                    alex = [items for items in linedata[1:]]
                    # print('alex',alex)
                    if '#VALUE!' in alex:
                        continue
                    elif '-99990' in alex:
                        continue
                    elif linedata[0] in row_name_filter1:
                        if not ns.all_are_floats(linedata[1:]):
                            continue
                    # print(linedata[0], linedata[1:])
                    writeobj.write(line)


def main():
    setUpData()
    #removeRubbish()
    print('Finished')
    print('{} {:.1f} seconds'.format('Program finished in', time.time() - start_time))
    print('HOORAY!')


'''MAIN'''
start_time = time.time()
if __name__ == '__main__':
    main()

#  NOTE BENE:  IN THIS LATEST VERSION MARCH 21ST 2018 I STOPPED HERE BECAUSE MUSTERING IS EASILY AND MORE
# COMPREHENSIVELY PERFORMED USING EXCEL FUNCTIONS IN THE SecondFilter.csv file

# def dataDefaults():
#     # NEXT REMOVE STOCKS FAILING THE DATA DEFAULTS BELOW
#     lextra = []  # list to hold collected five fields
#     SFstocknames = []  # those remaining in SecondFilter for iteration
#     SFstockfile = []  # store names of stocks passing these defaults
#     SFstockdata = []  # temporary storage for (x,y) tuple for each field
#
#     # get stocks remaining in SecondFilter
#     # get list of Second Filter stocks to append the extra data from next block of code
#     f = open(path1 + 'SecondFilter.csv', 'r')
#     f.readline()  # skip header
#     lines = f.readlines()
#     f.close()
#     SFdatalist = [ns.lss(line)
#                   for line in lines[1:]]  # rows titles removed
#     # first field in each line is the stock name
#     SFstocknames = [x[0] for x in SFdatalist]
#     print('UFL', SFstocknames)
#     print(len(SFstocknames), SFdatalist)
#
#     for stock in SFstocknames:
#         # print('Current stock', stock)
#         SFstockdata = []  # zero temporary list for next stock
#         count = 0
#         fileIsValid = True
#         print('Stock', stock)
#         # open the original ASX COMMSEC csv data file for this stock
#         fh = open(path2 + stock, 'r')
#         # a list containing all the rows/lines in the csv file
#         readCSV = [x for x in csv.reader(fh)]
#         fh.close()  # less confusion by immediately closing the file
#         # replace Company Historicals with the stockname
#         readCSV[0] = [stock]
#         # print(len(readCSV))  #for interest
#         for item in readCSV[1:]:
#             # item is one of the typically 39 lines/rows from the csv file.
#             # Each a list of strings, for example, ['Sales ($)', '0.17', '0.29',...]
#             # sleep(0.1)
#             # print('I', item[0], '\n')
#             if not item[0] in row_name_filter1:  # item[0] is the title of the row
#                 continue  # to the next row
#             r, x, y = ns.lineisvalid(item, stock)  # ','.join(item))
#             if not r:
#                 # print('T', item, ','.join(item))
#                 # if row title in filter1 then check to see if the data are valid#separator.join(sequence)
#                 print('Failed', stock)
#                 fileIsValid = False
#                 break  # break to for stock in SFstocknames
#             if not fileIsValid:
#                 continue  # to the next row
#             # print('Passed', stock)
#             # print
#             # these are the valid lines
#             # x, y = lineisvalid(','.join(item))
#             # print([x, y])
#             SFstockdata.append(
#                 (str(x), str(y)))  # building the seven extra data fields to be appended in the SecondFilter file
#             count += 1
#             if count == 7:
#                 lextra.append([stock] + SFstockdata)
#                 # list of stocks that passed muster
#                 SFstockfile.append(stock)
#                 break
#     # print(len(lextra))
#     # print(SFstockfile)
#     # print('Total passed', len(SFstocknames))
#     print('F', lextra)
#
#     # INSTALL THE SEVEN FIELDS AT THE APPROPRIATE POSITIONS IN THE SecondFilter file.
#     grandDictList = []
#     # contains the stock name and tuples of (index in header, data)
#     for item in lextra:
#         # change tuples after the stock name to a list of tuples
#         x = item[1:]
#         # print(item[0]) #the stock name
#         print('The tuples', x)  # the tuples
#         z = [(float(item[0]), float(item[1])) for item in x]
#         s = (sorted(z, key=(lambda item: item[0])))
#         # key in sorted is a function that determines which value to sort on
#         # item[0] is the column number of the csv header
#         print('Sorted', s)
#         z = [item[0]] + [value[1] for value in s]
#         # extract the values from the keys into their own list with stock name the first entry
#         b = {z[0]: z[1:] for i in range(0, len(z))}
#         # make a dictionary to look up the stock name. z[0] is the key, tuples values now a list of floats
#         # add each stock + data to the grandDictList (grandDictList is a LIST of dictionaries)
#         grandDictList.append(b)
#         # print(grandDictList)
#         # for entry in grandDictList:
#         #     print('B', entry, '\n')
#
# # INSTALL THEM IN THE ThirdFilterFile
# newsflist = []  # holds JPrice data with appended 5 fields
# sf = open(path1 + 'SecondFilter.csv', 'r')
# sf.readline()  # skip header
# lines = sf.readlines()
# # print('Line', lines)
# sf.close()
# sflist = [ns.lss(line) for line in lines[1:]]  # rows titles removed
# print('Second Filter list', sflist)
# for entry in sflist:
#     key = entry[0]
#     for dickt in grandDictList:
#         if key in dickt:
#             entry = entry + dickt.get(key)
#             newsflist.append(entry)
# # exit(0)
#
# # install extra fields into Third Filter
# tf = open(path1 + 'ThirdFilter.csv', 'w')
# tf.write(LongHeader)
# for item in newsflist:
#     # print('I', item)
#     testforpos = [float(item[x + 1]) >= 0 for x, y in enumerate(item[1:])]
#     # print(testforpos)
#     if all(testforpos):
#         makeAllStrings = [str(entry) for entry in item]
#         # print(item[0], testforpos)
#         r = ','.join(makeAllStrings)
#         tf.write(r + '\n')
# tf.close()
# # exit()
#
# # PASS MUSTER TESTING IS DONE HERE
# inv = open(path1 + 'ThirdFilter.csv', 'r')  # inv = investigate
# inv.readline()
# searchlist = [ns.lss(line) for line in inv.readlines()]
# finalList = []
# for item in searchlist:
#     # print("Nearly there", item)
#     testlist = [float(item[1]) >= HGEPS10, float(item[2]) >= HGEPS5, float(item[3]) >= STAEPS10,
#                 float(item[4]) >= STAEPS5,
#                 float(item[8]) >= MCAP, float(item[9]) <= PER, float(
#             item[10]) >= ROE, float(item[11]) <= DE,
#                 float(item[12]) >= IC]
#     # print(testlist)
#     if all(testlist):  # if all conditions are true
#         finalList.append(item[0])
# print("Passed muster = ", len(finalList))

# print(finalList)
# # exit()
#

# MAIN IS DONE IN TWO STAGES.  setUpData() followed by removeRubbish().  In between run 1ValVBA.xlsm and save Sheet2
#  as FirstFilter.csv



'''
STOCK    HGEPS10    HGEPS5    STAEPS10    STAEPS5    ROESTAG5    ROESTAG10    EQU/SHARE  MCAP  PER   ROE   DE  IC PO Ratio, EPS, Price, TARG, STRET, MOS, IVALUE
  0        1           2         3           4           5           6            7       8     9     10   11  12     13     14    15     16     17   18    19

Check above for current values
MCAP 100 [8]
PER 30   [9]
ROE 20   [10]
DE  50  [11]
IC  2.0  [12]

HGEPS10 = 5
HGEPS5  = 5
STAEPS10 = 75
STAEPS5 =75
'''

# wb = xw.Book(path1 + "1VALVBA.xlsm")
# wb.app.api.RegisterXLL("C:/Windows/System32/vsoft15.xll")
# wb.macro("ScreenASX1")()
# sht = wb.sheets[0]
# wb.save(sht)
# wb.close()
# exit(0)

# EXCEL Forumulae
# RMValue
# ThisWorkbook.Sheets("Sheet2").Cells(i, 25) = "=ROUND(Sheet1!K33/Sheet1!K10*Sheet1!K35*Sheet1!K36/1000 + Sheet1!K33/Sheet1!K10*(-15.975*POWER(Sheet1!K35/100,6) + 54.57*POWER(Sheet1!K35/100,5) - 59.153*POWER(Sheet1!K35/100,4) + 17.107*POWER(Sheet1!K35/100,3) + 61.923*POWER(Sheet1!K35/100,2) + 5.9754*(Sheet1!K35/100) -0.2458) *(100-Sheet1!K36)/100,2)"
# RMValue/Price
# ThisWorkbook.Sheets("Sheet2").Cells(i, 25) = "=ROUND(Sheet1!K33/Sheet1!K10*Sheet1!K35*Sheet1!K36/1000 + Sheet1!K33/Sheet1!K10*(-15.975*POWER(Sheet1!K35/100,6) + 54.57*POWER(Sheet1!K35/100,5) - 59.153*POWER(Sheet1!K35/100,4) + 17.107*POWER(Sheet1!K35/100,3) + 61.923*POWER(Sheet1!K35/100,2) + 5.9754*(Sheet1!K35/100) -0.2458) *(100-Sheet1!K36)/100/Sheet1!B40,2)"
