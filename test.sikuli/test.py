## co-operate with MS Office 356 Traditional Chinese version
import time
import os

# all path
skulilScriptParentPath = os.path.dirname(getBundlePath())
xlsxPath = skulilScriptParentPath  + '\\char.xlsx' 
saveDir = skulilScriptParentPath    # Do not contains chinese characters in your saveDir !!
excelPath = '"C:\\Program Files\\Microsoft Office\\root\\Office16\\excel.exe"'


now = time.strftime('%Y%m%d%H%M%S')
print ('NOW: %s' % now )

# Kill all running excel process before running our export operation
print('Kill all excel processes')
vcCMD= 'taskkill /f /im excel.exe'
run(vcCMD)

#type('D', Key.WIN) #show desktop

# open the XLSX file
vcCMD = excelPath+' '+xlsxPath 
App.open(vcCMD)


# once the XLSX bring out, locate the chart , right click on it 
#  click the "Save as pic", wait the "Save as Picture" diaglog pop out
wait("1618566418019.png")
chartLoc = find("1618798848260.png").offset(5,20)
rightClick(chartLoc)
click("2021-04-19 10_28_58-.png")
wait("1618799391266.png")

# Before we type our save path, we need switch the Input method to EN
type(Key.SPACE, Key.CTRL)  #Switch ime to En mode
pngFullPath = saveDir+'\\'+now+".png"
type(pngFullPath)
type(Key.ENTER)
print('file saved to: '+pngFullPath)
sleep(1)

# export operation is done, let us close the document
type(Key.F4, Key.ALT)
wait("1618799450287.png")
type('n')

