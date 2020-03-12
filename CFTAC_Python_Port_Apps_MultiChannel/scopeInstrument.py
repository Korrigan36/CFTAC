# -*- coding: utf-8 -*-
"""
Created on Tue Jan  8 08:28:12 2019

@author: v-coteal
"""

import copy
import time
import visa

class ScopeChannel():
    
    def __init__(self, ID):
        self.ID = ID
        
    def getChannelID(self):
        return str(copy.copy(self.ID))
    
              
           
class Scope():
    
    def __init__(self, scopeObject):
        self.channelCapacity = 0
        self.channelList = []
        self.scopeObject = scopeObject
        self.measurementList = []
    
    def sendCommand(self, command):
        self.scopeObject.write(str(command))

    def sendQuery(self, queryCommand):
        return self.scopeObject.query(queryCommand)
        
    def addChannel(self, scopeChannel):
        self.channelList.append(scopeChannel)
    
    def getNumMeasurements(self):
        return len(self.measurementList)
    
    def getChannelCapacity(self):
        return copy.copy(self.channelCapacity)
    
    def populateChannelList(self):
        for i in range (0, self.channelCapacity):
            channelName = "CH" + str(i + 1)
            self.channelList.append(ScopeChannel(channelName))
        
        
class MSO(Scope):
    
    def __init__(self, scopeObject):
        Scope.__init__(self, scopeObject)
        self.channelCapacity = 0
        self.populateChannelList()
        self.measurementList = self.getMeasurementList()
        
    def setHorizontalPosition(self, position):
        scopeCommand = "HORizontal:POSition " + str(position)
        self.sendCommand(scopeCommand)
    
    def getScopeMaxBW(self):
        scopeCommand = "CONFIGuration:ANALOg:BANDWidth?"
        return self.sendQuery(scopeCommand)
    
    def addMeasurement(self, channel, measType):
        channelID = self.channelList[channel - 1].getChannelID()
       
        scopeCommand = "MEASU:ADDMEAS " + measType
        self.sendCommand(scopeCommand)
        self.measurementList = self.getMeasurementList()
        
        numMeas = self.getNumMeasurements()
        scopeCommand = "MEASU:MEAS" + str(numMeas) + ":SOURCE " + channelID
        self.sendCommand(scopeCommand)
        
        scopeCommand = "MEASUREMENT:STATISTICS:CYCLEMODE 1"
        self.sendCommand(scopeCommand)
        
    def setMeasurementSource2(self, measNum, source2):
        channelID = self.channelList[source2 - 1].getChannelID()
        scopeCommand = "MEASUrement:MEAS" + str(measNum) + ":SOURCE2" + channelID
        self.sendCommand(scopeCommand)

        
#    def getMeasurementResult(self, measNum):
#        
#    
    def getMeasurementList(self):
        measListString = str(self.sendQuery("MEASU:LIST?")).encode('utf8')
        measList = measListString.split(",")
        
        for meas in measList:
            meas.rstrip()
        
        return measList
    
    def setLabel(self, channel, labelText):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = str(channelID) + ":LABel:NAMe \"" + str(labelText) + "\""
        self.sendCommand(scopeCommand)
        
    def setLabelYPos(self, channel, yPos):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = str(channelID) + ":LABel:YPOS " + str(yPos)
        self.sendCommand(scopeCommand)
        
    def setLabelXPos(self, channel, xPos):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = str(channelID) + ":LABel:XPOS " + str(xPos)
        self.sendCommand(scopeCommand)
        
    def setLabelBold(self, channel, onOff):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = str(channelID) + ":LABel:FONT:BOLD " + onOff
        self.sendCommand(scopeCommand)
        
    def setLabelFontSize(self, channel, fontSize):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = str(channelID) + ":LABel:FONT:SIZE " + str(fontSize)
        self.sendCommand(scopeCommand)
        
    def setLabelItalic(self, channel, onOff):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = str(channelID) + ":LABel:FONT:ITALic " + onOff
        self.sendCommand(scopeCommand)
        
    def setOffset(self, channel, offset):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = str(channelID) + ":OFFSet " + str(offset)
        self.sendCommand(scopeCommand)
        
    def setPosition(self, channel, position):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = str(channelID) + ":POSition " + str(position)
        self.sendCommand(scopeCommand)
        
    '''
    Sets vertical scale (default units Volts)
    '''
    def setVerticalScale(self, channel, vertScale):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = str(channelID) + ":SCALE " + str(vertScale)
        self.sendCommand(scopeCommand)
        
    def setEdgeTrigger(self, channel, level, slope):
        self.setEdgeTriggerSource(channel)
        self.setEdgeTriggerSlope(slope)
        self.setTriggerLevel(channel, level)
        
        
    def setEdgeTriggerSlope(self, slope):
        scopeCommand = ""
        if slope.lower() == "rise":
            scopeCommand = "TRIG:A:EDGE:SLO RISE"
        elif slope.lower() == "fall":
            scopeCommand = "TRIG:A:EDGE:SLO FALL"
        elif slope.lower() == "either":
            scopeCommand = "TRIG:A:EDGE:SLO EITher"
            
        self.sendCommand(scopeCommand)
        
        
    def setEdgeTriggerSource(self, source):
        channelID = self.channelList[source - 1].getChannelID()
        scopeCommand = "TRIG:A:EDGE:SOU " + channelID
        
        self.sendCommand(scopeCommand)
        
    def setTriggerLevel(self, channel, level):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = "TRIG:A:LEV:" + channelID + " " + str(level)
        
        self.sendCommand(scopeCommand)
        
    def setBandwidth(self, channel, bandwidth):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = channelID + ":BANdwidth " + str(bandwidth)
        
        self.sendCommand(scopeCommand)
        
    def setState(self, channel, state):
        channelID = self.channelList[channel - 1].getChannelID()
        scopeCommand = "DIS:GLO:" + channelID + ":STATE " + str(state)
        
        self.sendCommand(scopeCommand)
        
    def addSingleMeasurement(self, channel, measurement):
        self.addMeasurement(channel, measurement)

    def addMajorMeasurements(self, channel):
        self.addMeasurement(channel, "MEAN")
        self.addMeasurement(channel, "MAXIMUM")
        self.addMeasurement(channel, "MINIMUM")
        self.addMeasurement(channel, "PK2PK")
        
    def addTopBaseMeasurement(self, channel):
        self.addMeasurement(channel, "TOP")
        self.addMeasurement(channel, "BASE")
    

    def deleteMeasurement(self, measureNum):
        indexToRemove = self.measurementList.index(measureNum.upper())
        measurementToRemove = self.measurementList.pop(indexToRemove).rstrip()       
        scopeCommand = "MEASU:DEL \"" + str(measurementToRemove) + "\""
        
#        print "deleting measurement: " + str(measurementToRemove)
        self.sendCommand(scopeCommand)
        
    def deleteAllMeasurements(self):
        while len(self.measurementList) > 0:
            for measurement in self.measurementList:
                self.deleteMeasurement(str(measurement))
            
    def setHorizontalScale(self, scale):
        scopeCommand = "HOR:SCA " + str(scale)
        self.sendCommand(scopeCommand)
        
    def getHorizontalDivisions(self):
        scopeCommand = "HOR:DIV?"
        divisions = self.sendQuery(scopeCommand)
        return divisions
    
    '''
    Sets view style ie stacked or overlay
    
    String input {OVErlay or STAcked}
    '''
    def setViewStyle(self, viewStyle):
        scopeCommand = "DISplay:WAVEView1:VIEWStyle " + str(viewStyle.upper())
        self.sendCommand(scopeCommand)
    
    def setSampleRate(self, sampleRate):
        scopeCommand = "ORizontal:MODE:SAMPLERate " + sampleRate
        self.sendCommand(scopeCommand)
    
    def setPersistence(self, persist):
        scopeCommand = "DISplay:PERSistence " + persist
        self.sendCommand(scopeCommand)
    
    def setMode(self, mode):
        scopeCommand = "ACQUIRE:MODE " + mode
        self.sendCommand(scopeCommand)
    
    def returnSpecificMeasAllAqq(self, measNum, measType):
        scopeCommand = "MEASUrement:" + str(measNum.rstrip().upper()) + ":RESUlts:ALLAcqs:" + str(measType.upper()) + "?"
        result = self.sendQuery(scopeCommand).encode('utf8')
        
        return result
    
    def returnSpecificMeasCurrAqq(self, measNum, measType):
        scopeCommand = "MEASUrement:" + str(measNum.rstrip().upper()) + ":RESUlts:CURRentacq:" + str(measType.upper()) + "?"
        result = self.sendQuery(scopeCommand).encode('utf8')
        
        return result
    
    def returnAllMeasAllAqq(self, measType):
        listOfResults = []
        for measurement in self.measurementList:
            listOfResults.append(self.returnSpecificMeasAllAqq(measurement, measType).rstrip())

        return listOfResults
    
    def clear(self):
        scopeCommand = "CLEAR"
        self.sendCommand(scopeCommand)
        
    def scopeScreenCaptureCopyToPC(self, fileName):
        tempString = "*CLS"
        self.sendCommand(tempString)
        tempString = "SAVe:ASSIgn:TYPe IMAGe"
        self.sendCommand(tempString)
        tempString = "SAVe:IMAGe:FILEFormat PNG"
        self.sendCommand(tempString)
        tempString = "HARDCOPY:PORT ETHERnet"
        self.sendCommand(tempString)
        tempString = "HARDCOPY START"
        self.sendCommand(tempString)
        time.sleep(2)
    
#        print "before read raw"
        #Save image on scope harddrive
        self.sendCommand('SAVE:IMAGE \'c:/TEMP.PNG\'')
        self.sendQuery("*OPC?")  #Make sure the image has been saved before trying to read the file
        #Read file data over
        self.sendCommand('FILESYSTEM:READFILE \'c:/TEMP.PNG\'')
        raw_data = self.scopeObject.read_raw()
#        raw_data = visa.visalib.read(scopeObject, 1024)
#        print "after read raw"
        

        fid = open(str(fileName), 'wb')
        fid.write(raw_data)
        fid.close()    
        
    def stopAcq(self):
        scopeCommand = "ACQuire:STATE STOP"
        self.sendCommand(scopeCommand)
        
    def runAcq(self):
        scopeCommand = "ACQuire:STATE RUN"
        self.sendCommand(scopeCommand)
        
class MSO54(MSO):
    
    def __init__(self, scopeObject):
        MSO.__init__(self, scopeObject)
        self.channelCapacity = 4
        self.populateChannelList()
        self.measurementList = self.getMeasurementList()
        
class MSO58(MSO):
    
    def __init__(self, scopeObject):
        MSO.__init__(self, scopeObject)
        self.channelCapacity = 8
        self.populateChannelList()
        self.measurementList = self.getMeasurementList()