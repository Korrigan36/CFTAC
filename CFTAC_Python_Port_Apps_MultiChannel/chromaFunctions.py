# -*- coding: utf-8 -*-
"""
Created on Wed Feb 13 10:05:40 2019

@author: v-stpur
"""
#GIT Version

class ChromaFunctions(object):
    """A customer of ABC Bank with a checking account. Customers have the
    following properties:

    Attributes:
        name: A string representing the customer's name.
        balance: A float tracking the current balance of the customer's account.
    """

    def __init__(self, loadObject):
        """Return a Customer object whose name is *name* and starting
        balance is *balance*."""
        self.loadObject = loadObject

    def fetchCurrent(self):
        tempString = "FETC:CURR?"
        return self.loadObject.query(tempString)
    
    def setLoad(self, loadNumber, loadValue):
        tempString = "CHAN:LOAD " + str(loadNumber)
        self.loadObject.write(tempString)
        self.loadObject.write("CURR:STAT:L1 0\n")
        self.loadObject.write("CONFigure:AUTO:ON 0\n")
        self.loadObject.write("CHANnel:ACTive 1\n")
        tempString = "CURR:STAT:L1 " + str(loadValue) 
        self.loadObject.write(tempString)
        
    def setLoadOn(self, channel):
        tempString = "CHAN:LOAD " + str(channel)
        self.loadObject.write(tempString)
        tempString = "LOAD:STATe ON\n" 
        self.loadObject.write(tempString)
            
    def setLoadOff(self, channel):
        tempString = "CHAN:LOAD " + str(channel)
        self.loadObject.write(tempString)
        tempString = "LOAD:STATe OFF\n" 
        self.loadObject.write(tempString)
      
    def getLoad(self, channel):
        tempString = "CHAN:LOAD " + str(channel)
        self.loadObject.write(tempString)
        tempString = "CURR:STAT:L1?"  
        return self.loadObject.query(tempString)
      
    def load_DigSampleTime(self, number):
        tempString = "DIG:SAMP:TIME " + str(number) 
        self.loadObject.write(tempString)
    
    def load_DigSamplePoints(self, number):
        tempString = "DIG:SAMP:POIN " + str(number) 
        self.loadObject.write(tempString)
    
    def load_DigTriggerPoint(self, number):
        tempString = "DIG:TRIG:POIN " + str(number) 
        self.loadObject.write(tempString)
    
    def load_DigTriggerSource(self, text):
        tempString = "DIG:TRIG:SOUR " + text 
        self.loadObject.write(tempString)
    
    def load_SelectChannel(self, channel):
        tempString = "CHAN " + str(channel) 
        self.loadObject.write(tempString)
    
    def load_SelectSingleChannel(self, channel):
        tempString = "CHAN:LOAD " + str(channel) 
        self.loadObject.write(tempString)
    
    def load_EnableSingleChannel(self, text):
        tempString = "CHAN:ACT " + text 
        self.loadObject.write(tempString)
    
    def load_TurnOnOffLoad(self, text):
        tempString = "LOAD:STAT " + text 
        self.loadObject.write(tempString)
    
    def load_TurnOnOffLoadShort(self, text):
        tempString = "LOAD:SHOR " + text 
        self.loadObject.write(tempString)
    
    def load_AllRunOnOff(self, text):
        tempString = "CONF:ALLR " + text 
        self.loadObject.write(tempString)
    
    def load_AutoOnOff(self, text):
        tempString = "CONF:AUTO:ON " + text 
        self.loadObject.write(tempString)
    
    def load_ConfigWindow(self, wtime):
        tempString = "CONF:WIND " + str(wtime) 
        self.loadObject.write(tempString)
    
    def load_ConfigParallel(self, text):
        tempString = "CONF:PARA:MODE " + text 
        self.loadObject.write(tempString)
    
    def load_ConfigSync(self, text):
        tempString = "CONF:SYNC:MODE " + text 
        self.loadObject.write(tempString)
    
    def load_VoltLatchOnOff(self, text):
        tempString = "CONF:VOLT:LATC " + text 
        self.loadObject.write(tempString)
    
    def load_VoltOff(self, volts):
        tempString = "CONF:VOLT:OFF " + str(volts) 
        self.loadObject.write(tempString)
    
    def load_VoltOn(self, volts):
        tempString = "CONF:VOLT:ON " + str(volts) 
        self.loadObject.write(tempString)
    
    def load_VoltSign(self, text):
        tempString = "CONF:VOLT:SIGN " + text 
        self.loadObject.write(tempString)
    
    def load_CurrentMode(self, text):
        tempString = "MODE " + text 
        self.loadObject.write(tempString)
    
    def load_VoltRange(self, text):
        tempString = "CONF:VOLT:RANG " + text 
        self.loadObject.write(tempString)
    
    def load_RiseTime(self, text):
        tempString = "CURR:STAT:RISE " + text 
        self.loadObject.write(tempString)
    
    def load_FallTime(self, text):
        tempString = "CURR:STAT:FALL " + text 
        self.loadObject.write(tempString)
    
    def load_DynamicRiseTime(self, text):
        tempString = "CURR:DYN:RISE " + text 
        self.loadObject.write(tempString)
    
    def load_DynamicFallTime(self, text):
        tempString = "CURR:DYN:FALL " + text 
        self.loadObject.write(tempString)
    
    def load_DynamicLoad1(self, text):
        tempString = "CURR:DYN:L1 " + text 
        self.loadObject.write(tempString)
    
    def load_DynamicLoad2(self, text):
        tempString = "CURR:DYN:L2 " + text 
        self.loadObject.write(tempString)
    
    def load_DynamicTime1(self, text):
        tempString = "CURR:DYN:T1 " + text 
        self.loadObject.write(tempString)
    
    def load_DynamicTime2(self, text):
        tempString = "CURR:DYN:T2 " + text 
        self.loadObject.write(tempString)
    
    def load_DynamicPulseCount(self, text):
        tempString = "CURR:DYN:REP " + text 
        self.loadObject.write(tempString)
    
    def load_SetCurrent(self, current):
        tempString = "CURR:STAT:L1 " + str(current) 
        self.loadObject.write(tempString)
    
    def load_SetActiveOff(self, text):
        tempString = "CHAN:LOAD " + text 
        self.loadObject.write(tempString)
        tempString = "CHAN:ACT OFF"
        self.loadObject.write(tempString)
        


