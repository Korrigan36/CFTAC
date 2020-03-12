# -*- coding: utf-8 -*-
"""
Created on Wed Feb 13 08:51:30 2019

@author: v-stpur
"""
import subprocess
import time
import timerThread
#GIT Version
#sys.path.append("C:\\Users\\v-stpur\\Documents\\tools\\")
path = "C:\\Users\\v-stpur\\Documents\\tools\\"

class DsmcdbgFunctions(object):

#    def __init__(self, vregUnderTest):
    def __init__(self):
        
        self.deviceSerialNum = 0
     
    def testDsmcdbgInterface(self):           
        p = subprocess.Popen(path + "dsmcdbg license", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        if returnString.find("licensing") == -1:
            print "dsmcdbg interface broken"
            return False
        else:
            print "dsmcdbg interface Ok"
            return True

#    def getFTDIserialNumber(self):  
#        d = ftd.open(0)  
#        deviceList = d.getDeviceInfo()
#        self.deviceSerialNum = deviceList['serial']
#        print "FTDI Serial Number is: " + str(self.deviceSerialNum)
#        d.close()
#        return str(self.deviceSerialNum)
#            
    def setUserConfigRegisters(self):   

        self.hex1 = 0x32A9        
        self.hex2 = 0x005B   
        self.hex3 = 0x0054   
        
        print "Setting user config register"
        
        self.writeRegister(0x20, 0x01, 0x50, 0x8563, False)
        self.writeRegister(0x20, 0x01, 0x50, 0x85, True)
           
        #Disable Write Protect
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x01 stop wr 0x20 0x50 0x8563", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x01 stop wr 0x23 0x50 0x8563", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())

        #Increase OCP
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x00 stop wr 0x20 0x46 0x9B8F", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x01 stop wr 0x20 0x46 0xD0C0", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
#        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x02 stop wr 0x23 0x29 0x0020", stdout=subprocess.PIPE, shell=True)
#        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x02 stop wr 0x23 0x46 0xF9E4", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x02 stop wr 0x23 0x49 0x0088", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x00 stop wr 0x23 0x46 0xB8A0", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
#        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x02 stop wr 0x23 0x6C 0x05B3", stdout=subprocess.PIPE, shell=True)
#        returnString = str(p.communicate())

        #Enable Rails
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x00 stop wr 0x20 0x01 0x80", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x01 stop wr 0x20 0x01 0x80", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x00 stop wr 0x23 0x01 0x80", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x01 stop wr 0x23 0x01 0x80", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x02 stop wr 0x23 0x01 0x80", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())

        #Enable PMBus Mode
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x00 stop wr 0x20 0xc1 0x1cd4", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x00 stop wr 0x23 0xc1 0x18BD", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())

        #Disable UVL Fault
#        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x00 stop wr 0x20 0x32 0x722E", stdout=subprocess.PIPE, shell=True)
#        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x00 stop wr 0x23 0x32 0x0A38", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())

        #Configure User Trim Registers
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop wr 0x20 0x3F 0xCDA9", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x02 stop wr 0x23 0x3C " + '0x%0*X' % (4,self.hex1), stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x02 stop wr 0x23 0x3D " + '0x%0*X' % (4,self.hex2), stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x67 " + '0x%0*X' % (4,self.hex3), stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        #Configure RIMON
#        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x6E 0x33FD", stdout=subprocess.PIPE, shell=True)
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x6E 0x3400", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        #Set to 7 Phase
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x70 0x1327", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())

#        time.sleep(1)
        
        
    def setUserLowRange(self):
        #Set fine offset to zero
        self.setOffsetRegister(int("0x0000", 16))
        #Configure User Trim Registers
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop wr 0x20 0x3F 0xCDAB", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        #Set up gain coefficient and RIMON
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x6E 0x4400", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        #Set to 3 Phase
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x70 0x1323", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        #Increase Current Limit
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x00 stop wr 0x20 0x46 0xFFFF", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())

    def setUserHiRange(self):
        #Configure User Trim Registers
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop wr 0x20 0x3F 0xCDA9", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        #Set up gain coefficient and RIMON
#        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x6E 0x3400", stdout=subprocess.PIPE, shell=True)
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x6E 0x33FD", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        #Set to 7 Phase
        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x70 0x1327", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        #Increase Current Limit
        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x00 stop wr 0x20 0x46 0x9B8F", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())

#        time.sleep(1)

    def setGainRegister(self, gainHex, hiLow):   

        self.gainHex = gainHex        
            
        print "Setting gain to: " + '0x%0*X' % (4,self.gainHex)
        if self.vregUnderTest == "GFX":
            if hiLow:
                p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop wr 0x20 0xCA " + '0x%0*X' % (4,self.gainHex), stdout=subprocess.PIPE, shell=True)
            else:
                p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop wr 0x20 0xCC " + '0x%0*X' % (4,self.gainHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "CPU":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop wr 0x20 0xCB " + '0x%0*X' % (4,self.gainHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMIO":
#            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x00 stop wr 0x23 0x0B " + '0x%0*X' % (4,self.gainHex), stdout=subprocess.PIPE, shell=True)
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0xCB " + '0x%0*X' % (4,self.gainHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
#            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x00 stop wr 0x23 0x0A " + '0x%0*X' % (4,self.gainHex), stdout=subprocess.PIPE, shell=True)
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0xCA " + '0x%0*X' % (4,self.gainHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
#            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0xC5 " + '0x%0*X' % (4,self.gainHex), stdout=subprocess.PIPE, shell=True)
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0xCC " + '0x%0*X' % (4,self.gainHex), stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
#        print returnString
#        time.sleep(1)
            
    def readIntelliTempRegister(self):           
            
         if self.vregUnderTest == "GFX" or self.vregUnderTest == "CPU":
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop rdword 0x20 0x8D", stdout=subprocess.PIPE, shell=True)
         elif self.vregUnderTest == "MEMIO":
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0x8D", stdout=subprocess.PIPE, shell=True)
         elif self.vregUnderTest == "MEMPHY":
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x00 stop rdword 0x23 0x8D", stdout=subprocess.PIPE, shell=True)
         elif self.vregUnderTest == "SOC":
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x02 stop rdword 0x23 0x8D", stdout=subprocess.PIPE, shell=True)
         tempString = str(p.communicate())
         lowTempByte = tempString.split(" ")[2]
         lowTempByte = lowTempByte.replace('0x', "")
         highTempByte = tempString.split(" ")[3]
         highTempByte = highTempByte[:-6]
         temp = highTempByte + lowTempByte
#         print "Gain setting is now: " + str(gain)
         return int(temp, 16)

    def readControllerTempRegister(self):           
            
         if self.vregUnderTest == "GFX" or self.vregUnderTest == "CPU":
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop rdword 0x20 0xD5", stdout=subprocess.PIPE, shell=True)
         elif self.vregUnderTest == "MEMIO" or self.vregUnderTest == "MEMPHY" or self.vregUnderTest == "SOC":
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0xD5", stdout=subprocess.PIPE, shell=True)
         tempString = str(p.communicate())
         lowTempByte = tempString.split(" ")[2]
         lowTempByte = lowTempByte.replace('0x', "")
         highTempByte = tempString.split(" ")[3]
         highTempByte = highTempByte[:-6]
         temp = highTempByte + lowTempByte
         tempHex = (int(temp, 16)) & 0b0000001111111111
         tempInt = int(str(tempHex), 16)
         temperature = (tempInt*1.5625 - 275) / 3
#         if self.vregUnderTest == "GFX" or self.vregUnderTest == "CPU":
#             tempInt = (int(temp, 16)) / 3.0 
#         elif self.vregUnderTest == "MEMIO" or self.vregUnderTest == "MEMPHY" or self.vregUnderTest == "SOC":
#             tempInt = (int(temp, 16) - 300.0) / 3.0 
#         print "Gain setting is now: " + str(gain)
         return temperature

    def readGainRegister(self, hiLow):           
            
         if self.vregUnderTest == "GFX":
             if hiLow:
                 p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop rdword 0x20 0xCA", stdout=subprocess.PIPE, shell=True)
             else:
                 p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop rdword 0x20 0xCC", stdout=subprocess.PIPE, shell=True)
         elif self.vregUnderTest == "CPU":
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop rdword 0x20 0xCB", stdout=subprocess.PIPE, shell=True)
         elif self.vregUnderTest == "MEMIO":
#             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x00 stop rdword 0x23 0x0B", stdout=subprocess.PIPE, shell=True)
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0xCB", stdout=subprocess.PIPE, shell=True)
         elif self.vregUnderTest == "MEMPHY":
#             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x00 stop rdword 0x23 0x0A", stdout=subprocess.PIPE, shell=True)
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0xCA", stdout=subprocess.PIPE, shell=True)
         elif self.vregUnderTest == "SOC":
#             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0xC5", stdout=subprocess.PIPE, shell=True)
             p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0xCC", stdout=subprocess.PIPE, shell=True)
         gainString = str(p.communicate())
         lowTempByte = gainString.split(" ")[2]
         lowTempByte = lowTempByte.replace('0x', "")
         highTempByte = gainString.split(" ")[3]
         highTempByte = highTempByte[:-6]
         gain = highTempByte + lowTempByte
#         print "Gain setting is now: " + str(gain)
         return gain

    def setOffsetRegister(self, fineHex):           
            
        print "Setting fine offset to: " + '0x%0*X' % (4,fineHex)
        if self.vregUnderTest == "GFX":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x40 " + '0x%0*X' % (4,fineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "CPU":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop wr 0x20 0x40 " + '0x%0*X' % (4,fineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMIO":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x07 " + '0x%0*X' % (4,fineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x00 stop wr 0x23 0x1F " + '0x%0*X' % (4,fineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x10 " + '0x%0*X' % (4,fineHex), stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
#        time.sleep(1)
            
    def readOffsetRegister(self):           
            
        if self.vregUnderTest == "GFX":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop rdword 0x20 0x40", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "CPU":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop rdword 0x20 0x40", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMIO":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0x07", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x00 stop rdword 0x23 0x1F", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0x10", stdout=subprocess.PIPE, shell=True)
        gainString = str(p.communicate())
        lowTempByte = gainString.split(" ")[2]
        lowTempByte = lowTempByte.replace('0x', "")
        highTempByte = gainString.split(" ")[3]
        highTempByte = highTempByte[:-6]
        gain = highTempByte + lowTempByte
#        print "Fine Offset is Now: " + str(gain)
        return gain
     
     
    def setOutputVoltage(self, hexValue):
        
        if self.vregUnderTest == "GFX":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x21 " + '0x%0*X' % (4,hexValue), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "CPU":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x01 stop wr 0x20 0x21 " + '0x%0*X' % (4,hexValue), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMIO":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x21 " + '0x%0*X' % (4,hexValue), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x00 stop wr 0x23 0x21 " + '0x%0*X' % (4,hexValue), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x02 stop wr 0x23 0x21 " + '0x%0*X' % (4,hexValue), stdout=subprocess.PIPE, shell=True)
        gainString = str(p.communicate())
#        time.sleep(1)
 
    def writeNvrm(self):
        
        #Disable Rails
#        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x00 stop wr 0x20 0x01 0x00", stdout=subprocess.PIPE, shell=True)
#        returnString = str(p.communicate())
#        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x20 0x00 0x01 stop wr 0x20 0x01 0x00", stdout=subprocess.PIPE, shell=True)
#        returnString = str(p.communicate())
#        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x00 stop wr 0x23 0x01 0x00", stdout=subprocess.PIPE, shell=True)
#        returnString = str(p.communicate())
#        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x01 stop wr 0x23 0x01 0x00", stdout=subprocess.PIPE, shell=True)
#        returnString = str(p.communicate())
#        p = subprocess.Popen("dsmcdbg.exe sjm i2c wr 0x23 0x00 0x02 stop wr 0x23 0x01 0x00", stdout=subprocess.PIPE, shell=True)
#        returnString = str(p.communicate())
#
        if self.vregUnderTest == "GFX" or self.vregUnderTest == "CPU":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x20 0x00 0x00 stop wr 0x20 0x15", stdout=subprocess.PIPE, shell=True)
            returnString = str(p.communicate())
            time.sleep(1)
        elif self.vregUnderTest == "MEMIO" or self.vregUnderTest == "MEMPHY" or self.vregUnderTest == "SOC" :
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x00 stop wr 0x23 0x15", stdout=subprocess.PIPE, shell=True)
            returnString = str(p.communicate())
#            time.sleep(1)
 
    def setUpUltraFine(self, ultraFineHex):
        self.ultraFineHex = ultraFineHex
        
        if self.vregUnderTest == "GFX":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "CPU":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "MEMIO":
            self.ultraFineHex = self.ultraFineHex | 0b0000001000000000  
            #Set up 07 as ultrafine register"
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
#            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 0x0200", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
            self.ultraFineHex = self.ultraFineHex | 0b0000000010000000  
            #Set up 1F as ultrafine register"
            #They appear to be on page 1?
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
#            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 0x0080", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
            self.ultraFineHex = self.ultraFineHex | 0b0000100000000000  
            #Set up 10 as ultrafine register"
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
#            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 0x0800", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
#        time.sleep(1)
            
    def resetUltraFine(self):
        if self.vregUnderTest == "GFX":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "CPU":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "MEMIO":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 0x0000", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 0x0000", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 0x0000", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
#        time.sleep(1)
        
    def clearUltraFine(self, ultraFineHex):
        self.ultraFineHex = ultraFineHex
        
        if self.vregUnderTest == "GFX":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "CPU":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "MEMIO":
            self.ultraFineHex = self.ultraFineHex & 0b1111110011111111  
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
            #Data sheet says page 0 but need page 1 to make it work
            self.ultraFineHex = self.ultraFineHex & 0b1111111100111111  
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
            self.ultraFineHex = self.ultraFineHex & 0b1111001111111111  
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
#        time.sleep(1)
        
    def setUltraFineHighBit(self, ultraFineHex):
        self.ultraFineHex = ultraFineHex
        
        if self.vregUnderTest == "GFX":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "CPU":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "MEMIO":
            self.ultraFineHex = self.ultraFineHex | 0b0000000100000000  
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
            #This is on different paget then configure bit 7??
            self.ultraFineHex = self.ultraFineHex | 0b0000000001000000  
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
            self.ultraFineHex = self.ultraFineHex | 0b0000010000000000  
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
#        time.sleep(1)
        
    def clearUltraFineHighBit(self, ultraFineHex):
        self.ultraFineHex = ultraFineHex
        
        if self.vregUnderTest == "GFX":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "CPU":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "MEMIO":
            self.ultraFineHex = self.ultraFineHex & 0b1111111011111111  
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
            #Data sheet says page 0 but need page 1 to make it work
            self.ultraFineHex = self.ultraFineHex & 0b1111111110111111  
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
            self.ultraFineHex = self.ultraFineHex & 0b1111101111111111  
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop wr 0x23 0x13 " + '0x%0*X' % (4,self.ultraFineHex), stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
#        time.sleep(1)

    def readUltraFine(self):
     
        if self.vregUnderTest == "GFX":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "CPU":
            print "This Vreg Doesn't Do Ultra-Fine"
            return False
        elif self.vregUnderTest == "MEMIO":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0x13", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "MEMPHY":
            #This is on different paget then configure bit 7??
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0x13", stdout=subprocess.PIPE, shell=True)
        elif self.vregUnderTest == "SOC":
            p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0x13", stdout=subprocess.PIPE, shell=True)
        gainString = str(p.communicate())
        lowTempByte = gainString.split(" ")[2]
        lowTempByte = lowTempByte.replace('0x', "")
        highTempByte = gainString.split(" ")[3]
        highTempByte = highTempByte[:-6]
        gain = highTempByte + lowTempByte
#        print "Fine Offset is Now: " + str(gain)
        return  int(gain, 16)
#        return  gain
   
    def writeRegister(self, address, page, register, hexValue, byte):
        
        if byte:
            writeString = "dsmcdbg sjm i2c system wr " + '0x%0*X' % (2,address) + " 0x00 " + '0x%0*X' % (2,page) + " stop wr " + '0x%0*X' % (2,address) + " " + '0x%0*X' % (2,register) + " " + '0x%0*X' % (2,hexValue)
            print writeString
        else:
            writeString = "dsmcdbg sjm i2c system wr " + '0x%0*X' % (2,address) + " 0x00 " + '0x%0*X' % (2,page) + " stop wr " + '0x%0*X' % (2,address) + " " + '0x%0*X' % (2,register) + " " + '0x%0*X' % (4,hexValue)
            print writeString
        
#        p = subprocess.Popen("dsmcdbg sjm i2c system wr 0x23 0x00 0x01 stop rdword 0x23 0x13", stdout=subprocess.PIPE, shell=True)

    def resetBoard(self):
        p = subprocess.Popen("dsmcdbg reset", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        p = subprocess.Popen("dsmcdbg go", stdout=subprocess.PIPE, shell=True)
        returnString = str(p.communicate())
        return returnString
