import pandas as pd

print("enter the path to the txt file")
givenFilePath = input().strip('"')
givenFile=open(givenFilePath,"r")
rawData=givenFile.read()
splitData=rawData.split("\n")
basicInfo=[]
wavelengthStart=0
wavelengthEnd=0
peaks=[]
absorbanceValues=[]

def readSampleLineForWavelengths():
    global wavelengthStart, wavelengthEnd
    splitLine=line.split()    
    wavelengthEnd=int(splitLine[len(splitLine)-1])
    for j in range(0,len(splitLine)):
        if("Wavelength" in splitLine[j]):
            wavelengthStart=int(splitLine[j+1])
            break


i=0
for line in splitData:
    basicInfo.append(line)
    if("Start Wavelength" in line):
        wavelengthStart=int(line.split()[2])
    elif("End Wavelength" in line):
        wavelengthEnd=int(line.split()[2])    
    elif("Sample" in line):
        if(wavelengthStart == 0 and wavelengthEnd == 0):
            readSampleLineForWavelengths()
        #This points i at the line with peaks and absorbance
        i+=1
        break
    i+=1


#Peakmode is only needed for single wavelengths, which have extra information on the same line as peaks
def dealWithDataBlock():
    absorbanceMode=False
    peakMode=False
    skipNext=True
    for entry in splitData[absorbance].split():
        if (skipNext):
            skipNext=False    
        elif (absorbanceMode):
            absorbanceValues[absorbance-i-1].append(entry)
        #detection state, if it's not in absorbance and it hits and int, start looking for peaks    
        #Only needed for single wavelengths
        elif (entry.isdigit()):
            peakMode=True  
        elif (peakMode):
            peaks[absorbance-i-1].append(entry)
            peakMode=False    
        elif(not "Absorbance" in entry and "Wavescan" in basicInfo[0]):
            peaks[absorbance-i-1].append(entry)
        #detection state    
        elif ("Absorbance" in entry):
            absorbanceMode=True
            skipNext=True


for absorbance in range(i,len(splitData)):
    try:
        peaks.append([])
        absorbanceValues.append([])
        dealWithDataBlock()
    except Exception:
        print("error reading file")


givenFile.close()

try:
    absorbanceValues = [array for array in absorbanceValues if array]
    peaks = [array for array in peaks if array]
    columnNames=["Wavelength"]
    for i in range(0,len(absorbanceValues)):
        columnNames.append("Absorbance "+ str(i+1))
        columnNames.append("Peaks "+ str(i+1))


    WavelengthNums=[]
    for x in range(wavelengthStart, wavelengthEnd+1):
        WavelengthNums.append(x)

    columns=[WavelengthNums]
    for i in range(0, len(absorbanceValues)):
        columns.append(absorbanceValues[i])
        columns.append(peaks[i])


    data = pd.DataFrame(columns).T
    data.columns=columnNames
    fileName=givenFilePath.split('/')
    data.to_excel(fileName[-1].split('.')[0]+".xlsx", sheet_name="sheet1", index=False)
    print('New file created: ',fileName[-1].split('.')[0]+".xlsx")

except Exception:
    print("error creating excel file")

print("press any key to close the window")
input()