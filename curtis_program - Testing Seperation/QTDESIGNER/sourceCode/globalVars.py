from PyQt5.QtCore import QObject, QThread, pyqtSignal,pyqtSlot,QTimer
#PAGE1
page1e1 = "0"
page1e2 = "0"
page1e3 = "0"
page1Qual = "0"
page1Block = "0"
page1StartDate = '0'
page1EndDate = '0'




#PAGE2 Var
Page2fileName = "0"



#Page4
page4ReadFromFile = '0'
page4DesFile = '0'


#EXTRAS
qual1block1mirs=["3 MIRS / 1:6 / 3 HRS","0 MIRS","1 MIR / 1:12 / 3 HRS","0 MIRS"]
qual1block2mirs=["0 MIRS","0 MIRS","0 MIRS","1 MIR / 1:12 / 7 HRS","1 MIR / 1:12 / 4 HRS","1 MIR / 1:12 / 3 HRS","1 MIR / 1:12 / 2 HRS","1 MIR / 1:12 / 3.5 HRS","0 MIRS","1 MIR / 1:12 / 1.5HRS","1 MIR / 1:12 / 7.5HRS","1 MIR / 1:12 / 7.5HRS","1 MIR / 1:12 / 7.5HRS","0 MIRS","1 MIR / 1:12 / 7.5HRS","1 MIR / 1:12 / 7.5HRS","0 MIRS","0 MIRS"]
qual1block3mirs = [" 2 MIR / 1:8 / 4 HRS"," 2 MIR / 1:8 / 4.5 HRS"," 2 MIR / 1:8 / 5.5 HRS"," 2 MIR / 1:8 / 2.5 HRS"," 2 MIR / 1:8 / 4.5 HRS","0 MIRS"]

qual1block4mirs = ['2 MIRS / 1:8 / 6 HRS','2 MIRS/ 1:8 / 7 HRS','2 MIRS / 1:8 / 6 HRS','2 MIRS / 1:8 / 6.25 HRS','2 MIRS / 1:8 / 2 HRS','0 MIRS','2 MIRS / 1:8 / 5.5 HRS']

qual1block5mirs = ['0 MIRS','2 MIRS/ 1:8 / 3 HRS','2 MIRS/ 1:8 / 5.5 HRS','2 MIRS/ 1:8 / 5 HRS','2 MIRS/ 1:8 / 3.25 HRS','0 MIRS']

qual1block6mirs = ['1 MIRS / 1:12 / 3.5 HRS','1 MIRS / 1:12 / 5.75 HRS','1 MIRS / 1:12 / 4 HRS','1 MIRS / 1:12 / 6 HRS','1 MIRS / 1:12 / 5.25 HRS','1 MIRS / 1:12 / 4 HRS','1 MIRS / 1:12 / 6 HRS']
#QUAL2

qual2block1mirs =['1 MIR / 1:12 / 4 HRS','1 MIR / 1:12 / 5 HRS','3 MIRS / 1:6 / 8 HRS','3 MIRS / 1:6 / 8 HRS','3 MIRS / 1:6 / 8 HRS','3 MIRS / 1:6 / 8 HRS','3 MIRS / 1:6 / 8 HRS','2 MIRS / 1:8 / 8 HRS']

qual2block2mirs =['1 MIR / 1:12 / 3 HRS','1 MIR / 1:12 / 1 HR','3 MIRS / 1:6 / 3 HRS','3 MIRS / 1:6 / 8 HRS','3 MIRS / 1:6 / 8 HRS','3 MIR / 1:6 / 3 HRS','1 MIR / 1:12 / 2 HRS','2 MIRS / 1:8 / 6.5 HRS']
#QUAL3

qual3block1mirs = ['0 MIRS','1 MIRS / 1:12 / 4 HRS','1 MIRS / 1:12 / 8 HRS','1 MIRS / 1:12 / 3 HRS','1 MIRS / 1:12 / 5 HRS','2 MIRS / 1:8 / 8 HRS','2 MIRS / 1:8 / 8 HRS','1 MIRS / 1:12 / 8 HRS','1 MIRS / 1:12 / 8 HRS','1 MIRS / 1:12 / 6 HRS','1 MIRS / 1:12 / 6 HRS','1 MIRS / 1:12 / 1 HRS','0 MIRS']

qual3block2mirs =['1 MIRS / 1:12 / 3 HRS','1 MIRS / 1:12 / 8 HRS','1 MIRS / 1:12 / 8 HRS','1 MIRS / 1:12 / 8 HRS','1 MIRS / 1:12 / 8 HRS']

qual3block3mirs = ['1 MIRS / 1:12 / 1 HRS','1 MIRS / 1:12 / 8 HRS','1 MIRS / 1:12 / 5 HRS','1 MIRS / 1:12 / 8 HRS','1 MIRS / 1:12 / 4 HRS','2 MIRS / 1:8 / 6 HRS','2 MIRS / 1:8 / 8 HRS','2 MIRS / 1:8 / 4 HRS','0 MIRS']
fileerrorLabel = None
qualerrorLabel = None
blockerrorLabel = None
starterrorLabel = None
enderrorLabel = None
monthCheck = False

janEndDay = 29
febEndDay = 26
marEndDay = 31
aprEndDay = 30
mayEndDay = 31
junEndDay = 30
julEndDay = 31
augEndDay = 31
sepEndDay = 30
octEndDay = 31
novEndDay = 30
decEndDay = 31
monthEndDays = [janEndDay,febEndDay,marEndDay,aprEndDay,mayEndDay,junEndDay,julEndDay,augEndDay,sepEndDay,octEndDay,novEndDay,decEndDay]
monthEndDaysIndex = 0

qual1Block = [0,1,2,3,4,5,6]
qual1BlockIndex =1

blockTotalDayCount=[0,3,17,5,6,5,6]
qual2BlockTotalDayCount = [0,7,7]
qual3BlockTotalDayCount = [0,13,5,8]
blockTotalDayCountIndex=1
lastMonthDay = 0

startDate = 0
endDate = 0
currentblockTotalDayCount = 0
weekendCount=0

instructor = "None"

qualName = 0
blockName = 0
currentBlock = 0
sheetIndex = 0

totalDayAmount = 0

isChecked = False

holidayCheck = False

familyDays = 0


#QUAL1

# global qual1block1daytotal
qual1block1daytotal = 3

# global qual1block2daytotal
qual1block2daytotal = 17

# global qual1block3daytotal
qual1block3daytotal = 5

# global qual1block4daytotal
qual1block4daytotal = 6

# global qual1block5daytotal
qual1block5daytotal = 5

# global qual1block6daytotal
qual1block6daytotal = 6


#QUAL2

qual2block1daytotal = 7


qual2block2daytotal = 7

#Qual 3

qual3block1daytotal = 13


qual3block2daytotal = 5

qual3block3daytotal = 8





progressbarValue = 0
