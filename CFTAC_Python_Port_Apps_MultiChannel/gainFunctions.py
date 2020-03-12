# -*- coding: utf-8 -*-
"""
Created on Wed Feb 13 13:39:37 2019

@author: v-stpur
"""
import time
import timerThread
#import xlsxwriter

#from getVoltage import GetVoltage
#from chromaInstrument import ChromaMainframe
#from dsmcdbgFunctions import DsmcdbgFunctions
#from miscellaneousFunctions import MiscellaneousFunctions

version_Info = "Python PyQt5 IMON Automated Calibration V0.7"
#GIT Version


class GainFunctions(object):

#    def __init__(self, load, vregUnderTest, DMM, adInstr, runNotes, configID, pcbaSN, productSN, logFile, scsiCB, workBook, SummarySheet, PlotSheet):
#    def __init__(self, load, DMM, adInstr, runNotes, configID, pcbaSN, productSN, logFile, scsiCB, workBook, SummarySheet, PlotSheet):
    def __init__(self):
        
        #Create local versions of globals so code below doesn't change
        self.vregUnderTest      = timerThread.VregUnderTest
        self.dmm                = timerThread.DMM
        self.adInstr            = timerThread.AdInstr
        self.runNotes           = timerThread.RunNotes
        self.Load               = timerThread.Load
        self.configID           = timerThread.ConfigID
        self.pcbaSN             = timerThread.PcbaSN
        self.productSN          = timerThread.ProductSN
        self.logFilePointer     = timerThread.LogFilePointer
        self.scsiCB             = timerThread.ScsiCB
        self.offsetCount        = 0
        
        self.dsmcdbg            = timerThread.Dsmcdbg
        self.voltMeter          = timerThread.VoltMeter
        self.misc               = timerThread.Misc
        self.rowIndex           = 0 
        self.workBook           = timerThread.WorkBook
        self.summarySheet       = timerThread.SummarySheet
        self.plotSheet          = timerThread.PlotSheet

        tempString = self.runNotes
        tempString = tempString.replace(" ", "_")
        tempString = tempString.replace(":", "_")
        tempString = tempString.replace("-", "_")
        
#        self.excelFileName = "Chroma_IMON_Compare_" + tempString + ".xlsx"
#        self.workBook = xlsxwriter.Workbook(self.excelFileName)
#
        self.headerCellFormat = self.workBook.add_format()
        self.headerCellFormat.set_font_size(16)
        self.headerCellFormat.set_bold()
        
    def testInterface(self):
        return self.dsmcdbg.testDsmcdbgInterface()
        
    def gainLoopMPS(self, gainHex, hiLow):

        self.logFilePointer.write("Doing MPS Gain Method for " + self.vregUnderTest + "\n\n")
        print "Doing MPS Gain Method"
        
        self.gainHex = gainHex
        if self.vregUnderTest == "GFX":
#            minError = 0.2
            minPercent = 0.005
            GIMON = 1.0 #/16x
            Gcs   = 8.7 #uA/A
            if hiLow:
                RIMON = 5.0 #kOhms
                ideal_VIMON_gain = Gcs / 16.0 * GIMON * RIMON
            else:
#                RIMON = 40.0 #kOhms
                RIMON = 10.0 #kOhms
                ideal_VIMON_gain = Gcs / 8.0 * GIMON * RIMON
            
            print "Ideal VIMON Gain: " + str(ideal_VIMON_gain)
            VIMON_0A = 10.0 * ideal_VIMON_gain
            print "Ideal VIMON at 0A: " + str(VIMON_0A)
        elif self.vregUnderTest == "CPU":
#            minError = 0.2
            minPercent = 0.005
            RIMON = 10.0 #kOhms
            GIMON = 2.0 #/8x
            Gcs   = 8.7 #uA/A
            
            ideal_VIMON_gain = Gcs / 16.0 * GIMON * RIMON
            print "Ideal VIMON Gain: " + str(ideal_VIMON_gain)
            VIMON_0A = 10.0 * ideal_VIMON_gain
            print "Ideal VIMON at 0A: " + str(VIMON_0A)
        elif self.vregUnderTest == "SOC":
#            minError = 0.1
            minPercent = 0.005
            RIMON = 40.0 #kOhms
            GIMON = 1.0 #/16x
            Gcs   = 8.7 #uA/A
            
            ideal_VIMON_gain = Gcs / 16.0 * GIMON * RIMON
            print "Ideal VIMON Gain: " + str(ideal_VIMON_gain)
            VIMON_0A = 5.0 * ideal_VIMON_gain
            print "Ideal VIMON at 0A: " + str(VIMON_0A)
        
        elif (self.vregUnderTest == "MEMIO") or (self.vregUnderTest == "MEMPHY"):
#            minError = 0.1
            minPercent = 0.005
            RIMON = 40.0 #kOhms
            GIMON = 1.0 #/16x
            Gcs   = 10.0 #uA/A
            
            ideal_VIMON_gain = Gcs / 16.0 * GIMON * RIMON
            print "Ideal VIMON Gain: " + str(ideal_VIMON_gain)
            VIMON_0A = 5.0 * ideal_VIMON_gain
            print "Ideal VIMON at 0A: " + str(VIMON_0A)
        
        loopIndex = 0
        continueLoop = True
        firstTime = True
        countSuccess = 0
        while continueLoop:
            
            self.misc.setChromaCurrent(0.0)
            self.misc.loadOn()
            time.sleep(2)

            self.logFilePointer.write("Setting gain to: " + '0x%0*X' % (4,self.gainHex) + "\n")
            self.dsmcdbg.setGainRegister(self.gainHex, hiLow)
            self.dsmcdbg.readGainRegister(hiLow)

            vImon = self.voltMeter.getVoltage(1)
    
            zeroLoadIMON = vImon
            #Convert to mV
            zeroLoadIMON = zeroLoadIMON * 1000
            print "Zero Load IMON: " + str(zeroLoadIMON)

            #Set Load to TDC
            if self.vregUnderTest == "GFX":
                if hiLow:
                    TDC = 260.0
                else:
                    TDC = 63.0
            elif self.vregUnderTest == "CPU":
                TDC = 98.0
            elif (self.vregUnderTest == "MEMIO") or  (self.vregUnderTest == "MEMPHY"):
                TDC = 35.0
            elif self.vregUnderTest == "SOC":
                TDC = 50.0
            print "Setting chroma to TDC: " + str(TDC)
            self.misc.setChromaCurrent(TDC)
            time.sleep(2)

            vImon = self.voltMeter.getVoltage(1)
            
            readBackLoad = self.misc.readBackLoad()
            self.misc.loadOff()
    
            #Convert to mV
            tdcIMON = vImon * 1000
            print "TDC IMON: " + str(tdcIMON)
            #Run hardcoded for now will use slider later
            slope = (tdcIMON - zeroLoadIMON)/(readBackLoad)
            self.logFilePointer.write("Gain is: " + str(slope) + "\n")
            print "Gain is: " + str(slope)
            gainError = slope / ideal_VIMON_gain - 1
#            gainError = slope / ideal_VIMON_gain
            self.logFilePointer.write("Gain Error is: " + str(gainError) + "\n")
            print "Gain Error is: " + str(gainError)
            
            if abs(gainError) < minPercent:
                #Go again a few times just to be sure
                self.logFilePointer.write("No further gain correction needed\n")
                print "No further gain correction needed"
                countSuccess = countSuccess + 1
                if countSuccess > 3:
                    continueLoop = False
            elif (gainError > 0) and firstTime:
                #Or in the sign bit
                self.gainHex = self.gainHex | 0b0000000000100000
                firstTime = False
            elif abs(gainError) > minPercent: 
                self.gainHex = self.gainHex + 1

            loopIndex = loopIndex + 1
            
            #Heating causes drift so lets cool down a while
            if self.vregUnderTest == "GFX":
                if hiLow:
                    time.sleep(10)
                else:
                    time.sleep(5)
            elif self.vregUnderTest == "CPU":
                time.sleep(5)

        return self.gainHex

    def offsetLoopMPS(self, gainHex, hiLow, skipOffset):
        
        self.offsetCount += 1
        
        self.logFilePointer.write("Doing MPS Offset Loop for " + self.vregUnderTest + "\n\n")
        print "Doing MPS Offset Loop"
        self.gainHex = gainHex
        self.fineHex = 0x0000
        self.ultraFineHex = 0x0000
        target = 0
        minLoad = 0
        ultraFine = True

        if self.vregUnderTest == "GFX":
#            minError = 0.1
            minError = 0.17
#            GIMON = 1.0 #/16x
#            Gcs   = 8.7 #uA/A
#            if hiLow:
#                RIMON = 5.0 #kOhms
#                ideal_VIMON_gain = Gcs / 16.0 * GIMON * RIMON
#            else:
##                RIMON = 40.0 #kOhms
#                RIMON = 10.0 #kOhms
#                ideal_VIMON_gain = Gcs / 8.0 * GIMON * RIMON
#            print "RIMON is: " + str(RIMON)
            target = 10.0
            minLoad = 3.0
            
        elif self.vregUnderTest == "CPU":
            minError = 0.23
#            RIMON = 10.0 #kOhms
#            GIMON = 2.0 #/8x
#            Gcs   = 8.7 #uA/A
            target = 10.0
            minLoad = 1.0
#            ideal_VIMON_gain = Gcs / 16.0 * GIMON * RIMON
            
        elif self.vregUnderTest == "SOC":
#            minError = 0.12
            minError = 0.23
#            RIMON = 40.0 #kOhms
#            GIMON = 1.0 #/16x
#            Gcs   = 8.7 #uA/A
            target = 5.0
            minLoad = 1.0
#            ideal_VIMON_gain = Gcs / 16.0 * GIMON * RIMON
            
            self.ultraFineHex = self.dsmcdbg.readUltraFine()
#            print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
            if not skipOffset:
                self.dsmcdbg.setUpUltraFine(self.ultraFineHex)
            self.ultraFineHex = self.dsmcdbg.readUltraFine()
            print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)

        elif (self.vregUnderTest == "MEMIO") or (self.vregUnderTest == "MEMPHY"):
            minError = 0.15
#            RIMON = 40.0 #kOhms
#            GIMON = 1.0 #/16x
#            Gcs   = 10.0 #uA/A
            target = 5.0
            minLoad = 1.0
#            ideal_VIMON_gain = Gcs / 16.0 * GIMON * RIMON
           
            self.ultraFineHex = self.dsmcdbg.readUltraFine()
#            print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
            if not skipOffset:
                self.dsmcdbg.setUpUltraFine(self.ultraFineHex)
            self.ultraFineHex = self.dsmcdbg.readUltraFine()
            print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
        

        self.misc.setChromaCurrent(minLoad)
        self.misc.loadOn()
        time.sleep(2)
        
        readBackLoad = self.misc.readBackLoad()

#        print "Ideal VIMON Gain: " + str(ideal_VIMON_gain)
#        self.logFilePointer.write("Ideal VIMON Gain: " + str(ideal_VIMON_gain) + "\n")
#        print "readbackload is: " + str(readBackLoad)
#        VIMON_MIN = (target + readBackLoad) * ideal_VIMON_gain
#        print "Ideal VIMON at MIN Load: " + str(VIMON_MIN)
#        self.logFilePointer.write("Ideal VIMON at MIN Load: " + str(VIMON_MIN) + "\n")

        offsetLoop = True
        countSuccess = 0
        errorSum = 0
        while offsetLoop:
            
#            vImon = self.voltMeter.getVoltage(1)
#            #Change to mV 
#            vImon = vImon * 1000.0    
#            print "VIMON is: " + str(vImon)
#            readBackLoad = self.misc.readBackLoad()
#            VIMON_MIN = (target + readBackLoad) * ideal_VIMON_gain
#            offset_Error = (vImon - VIMON_MIN) / RIMON
#            self.logFilePointer.write("Offset Error at MIN: " + str(offset_Error) + "\n")
#            print "Offset Error Method 1: " + str(offset_Error)
#
#            vImon = self.voltMeter.getVoltage(1)
#            print "VIMON is: " + str(vImon)
#            imon = vImon * (1.0/((Gcs*10E-6)*(GIMON/16.0)*RIMON*1000))
#            print "IMON is: " + str(imon)
#            readBackLoad = self.misc.readBackLoad()
#            offset_Error2 = (imon - (target + readBackLoad))
#            print "Offset Error Method 2: " + str(offset_Error2)
#            
            vImon = self.voltMeter.getVoltage(1)
            print "VIMON is: " + str(vImon)
            imon = self.misc.calculateCurrent(vImon, hiLow)
            print "IMON is: " + str(imon)
            readBackLoad = self.misc.readBackLoad()
            offset_Error = (imon - (target + readBackLoad))
            print "Offset Error Method 3: " + str(offset_Error)
            self.logFilePointer.write("Offset Error at MIN: " + str(offset_Error) + "\n")
            
            #No need for sign bit because we always start with large offset
            if offset_Error > minError: 
                countSuccess = 0
                errorSum = 0
                self.logFilePointer.write("Adjust Offset Down 1 LSB at a time\n")
                print "Adjust Offset Down 1 LSB at a time"
                        
                #Set Offset to one less than last run
                self.gainHex = self.gainHex - 0x0040
                self.dsmcdbg.setGainRegister(self.gainHex, hiLow)
                self.logFilePointer.write("Setting gain to: " + '0x%0*X' % (4,self.gainHex) + "\n")
                
            elif offset_Error < -minError:
                countSuccess = 0
                errorSum = 0
                self.logFilePointer.write("Adjust Offset Up 1 LSB at a time\n")
                print "Adjust Offset Up 1 LSB at a time"
                
                #Set Offset to one more than last run
                self.gainHex = self.gainHex + 0x0040
                self.dsmcdbg.setGainRegister(self.gainHex, hiLow)
                self.logFilePointer.write("Setting gain to: " + '0x%0*X' % (4,self.gainHex) + "\n")
                
            else:
                self.logFilePointer.write("No Coarse Offset Adjustment Needed\n")
                print "No Coarse Offset Adjustment Needed"
                countSuccess = countSuccess + 1
                errorSum += offset_Error
                time.sleep(3)
                if countSuccess > 4:
                    offsetLoop = False
            
            offset = self.dsmcdbg.readGainRegister(hiLow)
            self.logFilePointer.write("Gain setting is now: " + str(offset) + "\n")
            print "Gain setting is now: " + str(offset)
            
        if ((self.vregUnderTest == "MEMIO" or self.vregUnderTest == "MEMPHY" or self.vregUnderTest == "SOC" or self.vregUnderTest == "CPU") or ((self.vregUnderTest == "GFX") and hiLow)) and (self.offsetCount == 2) and (not skipOffset):
            if (self.vregUnderTest == "MEMIO" or self.vregUnderTest == "MEMPHY" or self.vregUnderTest == "SOC") and offset_Error < 0.0:
                #We can use ultra-fine but have to start above target since ultra-fine is only negative correction
                self.logFilePointer.write("Offset error is now: " + str(offset_Error) + "\n")
                print "Offset error is now: " + str(offset_Error)
                self.logFilePointer.write("Bumping back up one because correction is negative\n")
                print "Bumping back up one because correction is negative"
                self.gainHex = self.gainHex + 0x0040
                self.dsmcdbg.setGainRegister(self.gainHex, hiLow)
                self.logFilePointer.write("Setting gain to: " + '0x%0*X' % (4,self.gainHex) + "\n")
                offset = self.dsmcdbg.readGainRegister(hiLow)
                print ""
                self.logFilePointer.write("Gain setting is now: " + str(offset) + "\n")
                print "Gain setting is now: " + str(offset)
                print ""
            elif (self.vregUnderTest == "GFX" or self.vregUnderTest == "CPU") and offset_Error > 0.0:
                #Fine adjust for these rails is only positive
                self.logFilePointer.write("Offset error is now: " + str(offset_Error) + "\n")
                print "Offset error is now: " + str(offset_Error)
                self.logFilePointer.write("Bumping down one because correction is positive\n")
                print "Bumping down one because correction is positive"
                self.gainHex = self.gainHex - 0x0040
                self.dsmcdbg.setGainRegister(self.gainHex, hiLow)
                self.logFilePointer.write("Setting gain to: " + '0x%0*X' % (4,self.gainHex) + "\n")
                offset = self.dsmcdbg.readGainRegister(hiLow)
                print ""
                self.logFilePointer.write("Gain setting is now: " + str(offset) + "\n")
                print "Gain setting is now: " + str(offset)
                print ""
                    
#            vImon = self.voltMeter.getVoltage(1)
#            #Change to mV 
#            vImon = vImon * 1000    
#            print "VIMON is: " + str(vImon)
#            readBackLoad = self.misc.readBackLoad()
#            VIMON_MIN = (target + readBackLoad) * ideal_VIMON_gain
#            offset_Error = (vImon - VIMON_MIN) / RIMON
#            print ""
#            self.logFilePointer.write("Offset Error at MIN: " + str(offset_Error) + "\n")
#            print "Offset Error at MIN: " + str(offset_Error)
#            print ""
#
            vImon = self.voltMeter.getVoltage(1)
            print "VIMON is: " + str(vImon)
            imon = self.misc.calculateCurrent(vImon, hiLow)
            print "IMON is: " + str(imon)
            readBackLoad = self.misc.readBackLoad()
            offset_Error = (imon - (target + readBackLoad))
            print "Offset Error Method 3: " + str(offset_Error)
                    
            if self.vregUnderTest == "MEMIO" or self.vregUnderTest == "MEMPHY" or self.vregUnderTest == "SOC":
                minError = 0.014
#                if (offset_Error > 0.220) and ultraFine:
#                    self.logFilePointer.write("Error too large for ultra-fine register so clear ultra-fine bits\n")
#                    print "Error too large for ultra-fine register so clear ultra-fine bits"
#                    self.ultraFineHex = self.dsmcdbg.readUltraFine()
#                    self.dsmcdbg.clearUltraFine(self.ultraFineHex)
#                    self.ultraFineHex = self.dsmcdbg.readUltraFine()
#                    print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
#                    ultraFine = False
#                    minError = 0.017
#                elif ultraFine:
#                    minError = 0.004
#                    self.ultraFineHex = self.dsmcdbg.readUltraFine()
#                    print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
#                    self.dsmcdbg.setUpUltraFine(self.ultraFineHex)
#                    self.ultraFineHex = self.dsmcdbg.readUltraFine()
#                    print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
#                    self.dsmcdbg.clearUltraFineHighBit(self.ultraFineHex)
#                    self.ultraFineHex = self.dsmcdbg.readUltraFine()
#                    print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
#                else:
#                    minError = 0.017
#                    
            elif self.vregUnderTest == "CPU":
                minError = 0.017
            else:
                minError = 0.017
#            minError = 0.017
                
                    
                
            self.logFilePointer.write("Doing fine MPS offset \n")
            print "Doing fine MPS offset"
            offsetLoop = True
            firstTime = True
            wasNegative = False
            self.finHex = 0x0000
            countSuccess = 0
            errorSum = 0
            while offsetLoop:
    
#                vImon = self.voltMeter.getVoltage(4)
#                #Change to mV 
#                vImon = vImon * 1000    
#                print "VIMON is: " + str(vImon)
#                readBackLoad = self.misc.readBackLoad()
#                VIMON_MIN = (target + readBackLoad) * ideal_VIMON_gain
#                offset_Error = (vImon - VIMON_MIN) / RIMON
#                self.logFilePointer.write("Offset Error at MIN: " + str(offset_Error) + "\n")
#                print "Offset Error at MIN: " + str(offset_Error)
                
                vImon = self.voltMeter.getVoltage(1)
                print "VIMON is: " + str(vImon)
                imon = self.misc.calculateCurrent(vImon, hiLow)
                print "IMON is: " + str(imon)
                readBackLoad = self.misc.readBackLoad()
                offset_Error = (imon - (target + readBackLoad))
                print "Offset Error Method 3: " + str(offset_Error)
                self.logFilePointer.write("Offset Error at MIN: " + str(offset_Error) + "\n")
                    
                #No need for sign bit because we always start with large offset
                if offset_Error > minError: 
                    countSuccess = 0
                    errorSum = 0
                    self.logFilePointer.write("Adjust Offset Down 1 LSB at a time\n")
                    print "Adjust Offset Down 1 LSB at a time"
                    if self.vregUnderTest == "MEMIO" or self.vregUnderTest == "MEMPHY" or self.vregUnderTest == "SOC":
                        #Set offset to one more than last run
                        #For ultra-fine registers we should never have the sign bit set
#                        if self.fineHex >=  0x8040:
#                            #Should never do this Ios_C not available
#                            #Sign bit set and we have positive bits set so subtract a positive bit
#                            #This code should never get called for ultra-fine registers
#                            self.fineHex = self.fineHex - 0x0040
#                        elif self.fineHex == 0x8000:
#                            #This code should never get called for ultra-fine registers
#                            #Clear the sign bit and add a negative bit
#                            self.fineHex = self.fineHex & 0b0111111111111111  
#                            self.fineHex = self.fineHex + 0x0040
#                        else:
#                            #Sign bit already cleared so just add a negative bit
                        self.fineHex = self.fineHex + 0x0040
                
                        if self.fineHex == 0x0400 and ultraFine:
                            if firstTime:
                                print ""
                                self.logFilePointer.write("Setting high bit in 13h\n")
                                print "Setting high bit in 13h"
                                print ""
                                #Zero Low Bits
                                self.fineHex = 0x0000
                                #Set Upper Bit
                                self.ultraFineHex = self.dsmcdbg.readUltraFine()
                                self.dsmcdbg.setUltraFineHighBit(self.ultraFineHex)
                                self.ultraFineHex = self.dsmcdbg.readUltraFine()
                                print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
                                firstTime = False
                            else:
                                self.fineHex = self.fineHex - 0x0040
#                                if ultraFine == False:
#                                    offsetLoop = False
#                                    break
                                self.logFilePointer.write("We've set all the bits but unable to reach target. Quit\n")
                                print "We've set all the bits but unable to reach target. Quit"
#                                self.ultraFineHex = self.dsmcdbg.readUltraFine()
#                                self.dsmcdbg.clearUltraFine(self.ultraFineHex)
#                                self.ultraFineHex = self.dsmcdbg.readUltraFine()
#                                print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
                                ultraFine = False
                                offsetLoop = False
                                break
    #                            minError = 0.017
            
            #                    print "Setting Fine Offset to: " + '0x%0*X' % (4, self.fineHex)
                    elif self.vregUnderTest == "CPU" or self.vregUnderTest == "GFX":
                        self.logFilePointer.write("Crossed Zero: Vreg Fine Correction Only Positive\n")
                        print "Crossed Zero: Vreg Fine Correction Only Positive" 
                        if wasNegative and (self.fineHex >= 0x0004):
                            self.fineHex = self.fineHex - 0x0004
                        else:
                            self.logFilePointer.write("Should Not be here. This rail needs to start at negative\n")
                            print "Should Not be here. This rail needs to start at negative"
                            #Bump the gain register down one to fix this
                            self.gainHex = self.gainHex - 0x0040
                            self.dsmcdbg.setGainRegister(self.gainHex, hiLow)
                            self.logFilePointer.write("Setting gain to: " + '0x%0*X' % (4,self.gainHex) + "\n")
    
    
    
                    self.dsmcdbg.setOffsetRegister(self.fineHex)
                    self.logFilePointer.write("Setting fine offset to: " + '0x%0*X' % (4,self.fineHex) + "\n")
                            
                elif offset_Error < -minError:
                    countSuccess = 0
                    errorSum = 0
                    wasNegative = True
                    self.logFilePointer.write("Adjust Offset Up 1 LSB at a time\n")
                    print "Adjust Offset Up 1 LSB at a time"
                    if self.vregUnderTest == "MEMIO" or self.vregUnderTest == "MEMPHY" or self.vregUnderTest == "SOC":
                        #Set offset to one less than last run to lower imonCurrent
                        #For ultra-fine registers we should never have the sign bit set
#                        if self.fineHex >= 0x0040 and self.fineHex < 0x8000:
#                            #Sign bit not set but we have negative bits set to subtract a bit
#                            self.fineHex = self.fineHex - 0x0040
#                        elif self.fineHex == 0x0000:
#                            #Set the sign bit and add
#                            #This code should never get called for ultra-fine registers
#                            self.fineHex = self.fineHex | 0b1000000000000000   
#                            self.fineHex = self.fineHex + 0x0040
#                        else:
#                            #Sign bit already set to just add a bit
#                            #This code should never get called for ultra-fine registers
                        self.fineHex = self.fineHex + 0x0040
                            
                        #We should never run this code
                        if self.fineHex == 0x0400 and ultraFine:
                            if firstTime:
                                print ""
                                self.logFilePointer.write("Setting high bit in 13h from place we shouldn't be\n")
                                print "Setting high bit in 13h from place we shouldn't be"
                                print ""
                                #Zero Low Bits
                                self.fineHex = 0x0000
                                #Set Upper Bit
                                self.ultraFineHex = self.dsmcdbg.readUltraFine()
                                self.dsmcdbg.setUltraFineHighBit(self.ultraFineHex)
                                self.ultraFineHex = self.dsmcdbg.readUltraFine()
                                print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
                                firstTime = False
                            else:
                                self.fineHex = self.fineHex - 0x0040
#                                if ultraFine == False:
#                                    offsetLoop = False
#                                    break
                                self.logFilePointer.write("We've set all the bits but unable to reach target. Quitting\n")
                                print "We've set all the bits but unable to reach target. Quit"
#                                self.ultraFineHex = self.dsmcdbg.readUltraFine()
#                                self.dsmcdbg.clearUltraFine(self.ultraFineHex)
#                                self.ultraFineHex = self.dsmcdbg.readUltraFine()
#                                print "Ultra Fine Register contains: " + '0x%0*X' % (4,self.ultraFineHex)
                                ultraFine = False
                                offsetLoop = False
                                break
    #                            minError = 0.017
                        
                    elif self.vregUnderTest == "CPU" or self.vregUnderTest == "GFX":
                        if (self.fineHex <=  0x00EC) and (offset_Error <= -0.05):
                            self.fineHex = self.fineHex + 0x0010
                        elif (self.fineHex <=  0x00F8) and (offset_Error > -0.05): 
                            self.fineHex = self.fineHex + 0x0004
                        elif self.fineHex ==  0x00FC:
                            print "All bits set quitting"
                            self.logFilePointer.write("We've set all the bits but unable to reach target. Quitting\n")
                            offsetLoop = False
                            break
                            
                    
                    self.dsmcdbg.setOffsetRegister(self.fineHex)
                    self.logFilePointer.write("Setting fine offset to: " + '0x%0*X' % (4,self.fineHex) + "\n")
                
                else:
                    self.logFilePointer.write("No Fine Offset Adjustment Needed\n")
                    print "No Fine Offset Adjustment Needed"
                    #Go again a few times just to be sure
                    countSuccess = countSuccess + 1
                    errorSum += offset_Error
                    time.sleep(3)
                    if countSuccess > 4:
                        offsetLoop = False
            
        self.misc.setChromaCurrent(minLoad)
        self.misc.loadOn()
        time.sleep(2)

#        readBackLoad = self.misc.readBackLoad()
#
#        VIMON_MIN = (target + readBackLoad) * ideal_VIMON_gain
#        print "Ideal VIMON at Min Load: " + str(VIMON_MIN)
##        vImon = self.voltMeter.getVoltage(20)
#        vImon = self.voltMeter.getVoltage(4)
#        #Change to mV 
#        vImon = vImon * 1000    
#        print "VIMON is: " + str(vImon)
#        offset_Error = (vImon - VIMON_MIN) / RIMON

        vImon = self.voltMeter.getVoltage(1)
        print "VIMON is: " + str(vImon)
        imon = self.misc.calculateCurrent(vImon, hiLow)
        print "IMON is: " + str(imon)
        readBackLoad = self.misc.readBackLoad()
        offset_Error = (imon - (target + readBackLoad))
        print "Offset Error Method 3: " + str(offset_Error)
                    
        errorSum += offset_Error

        if ultraFine == True:
            self.summarySheet.write(timerThread.TweakCell, errorSum / 6.0)
                
            print ""
            self.logFilePointer.write("Averaged Offset Error at Min Load: " + str(errorSum / 6.0) + "\n")
            print "Averaged Offset Error at Min Load: " + str(errorSum / 6.0)
            print ""
        else:
            self.summarySheet.write(timerThread.TweakCell, offset_Error)
                
            print ""
            self.logFilePointer.write("Offset Error at Min Load: " + str(offset_Error) + "\n")
            print "Offset Error at Min Load: " + str(offset_Error)
            print ""
            
        offset = self.dsmcdbg.readOffsetRegister()
        self.logFilePointer.write("Fine offset setting is now: " + str(offset) + "\n")
        print "Fine offset setting is now: " + str(offset)
        print ""
            
#        for i in range (0, 10):
#            readBackLoad = self.misc.readBackLoad()
#            VIMON_MIN = (target + readBackLoad) * ideal_VIMON_gain
#            vImon = self.voltMeter.getVoltage(1)
#            #Change to mV 
#            vImon = vImon * 1000    
#            print "VIMON is: " + str(vImon)
#            offset_Error = (vImon - VIMON_MIN) / RIMON
#            print "Offset Error at MIN Load: " + str(offset_Error)

        self.misc.loadOff()

        if ultraFine == True:
            return self.gainHex, errorSum / 6.0
        else:
            return self.gainHex, offset_Error
    
    def imonCheck(self, hiLow, sheetCount):
        
        print "Doing IMON Check"
        self.logFilePointer.write("Doing Load Sweep for " + self.vregUnderTest + "\n\n")
        self.rowIndex = 0
#        self.summaryRowIndex = 0
        voutHex = 0x00A0

        if hiLow and (self.vregUnderTest == "GFX"):
            dmmSheet = self.workBook.add_worksheet(str(sheetCount) + "IMON " + self.vregUnderTest + " " + "High")
            dmmSheetName =  "'" + str(sheetCount) + "IMON " + self.vregUnderTest + " " + "High'"
            plotName = "GFX High Range Error Delta"
            percentName = "GFX High Range Percent Error"
            
            if sheetCount == 1:
                timerThread.ImonChartGFXHi = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.ImonChartGFXHi.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.ImonChartGFXHi.set_y_axis({'name': 'Error (Amps)'})
                timerThread.ImonChartGFXHi.set_title({'name': plotName})  
                timerThread.ImonChartGFXHi.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
                timerThread.PercentChartGFXHi = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.PercentChartGFXHi.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.PercentChartGFXHi.set_y_axis({'name': 'Percent Error'})
                timerThread.PercentChartGFXHi.set_title({'name': percentName})  
                timerThread.PercentChartGFXHi.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
        elif self.vregUnderTest == "GFX":
            dmmSheet = self.workBook.add_worksheet(str(sheetCount) + "IMON " + self.vregUnderTest + " " + "Low")
            dmmSheetName =  "'" + str(sheetCount) + "IMON " + self.vregUnderTest + " " + "Low'"
            plotName = "GFX Low Range Error Delta"
            percentName = "GFX Low Range Percent Error"
            
            if sheetCount == 1:
                timerThread.ImonChartGFXLow = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.ImonChartGFXLow.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.ImonChartGFXLow.set_y_axis({'name': 'Error (Amps)'})
                timerThread.ImonChartGFXLow.set_title({'name': plotName})  
                timerThread.ImonChartGFXLow.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
                timerThread.PercentChartGFXLow = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.PercentChartGFXLow.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.PercentChartGFXLow.set_y_axis({'name': 'Percent Error'})
                timerThread.PercentChartGFXLow.set_title({'name': percentName})  
                timerThread.PercentChartGFXLow.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
        elif self.vregUnderTest == "CPU":
            dmmSheet = self.workBook.add_worksheet(str(sheetCount) + "IMON " + self.vregUnderTest)
            dmmSheetName =  "'" + str(sheetCount) + "IMON " + self.vregUnderTest + "'"
            plotName = self.vregUnderTest + " Error Delta"
            percentName = self.vregUnderTest + " Percent Error"
            
            if sheetCount == 1:
                timerThread.ImonChartCPU = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.ImonChartCPU.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.ImonChartCPU.set_y_axis({'name': 'Error (Amps)'})
                timerThread.ImonChartCPU.set_title({'name': plotName})  
                timerThread.ImonChartCPU.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
                timerThread.PercentChartCPU = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.PercentChartCPU.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.PercentChartCPU.set_y_axis({'name': 'Percent Error'})
                timerThread.PercentChartCPU.set_title({'name': percentName})  
                timerThread.PercentChartCPU.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
        elif self.vregUnderTest == "MEMPHY":
            dmmSheet = self.workBook.add_worksheet(str(sheetCount) + "IMON " + self.vregUnderTest)
            dmmSheetName =  "'" + str(sheetCount) + "IMON " + self.vregUnderTest + "'"
            plotName = self.vregUnderTest + " Error Delta"
            percentName = self.vregUnderTest + " Percent Error"
            
            if sheetCount == 1:
                timerThread.ImonChartMEMPHY = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.ImonChartMEMPHY.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.ImonChartMEMPHY.set_y_axis({'name': 'Error (Amps)'})
                timerThread.ImonChartMEMPHY.set_title({'name': plotName})  
                timerThread.ImonChartMEMPHY.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
                timerThread.PercentChartMEMPHY = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.PercentChartMEMPHY.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.PercentChartMEMPHY.set_y_axis({'name': 'Percent Error'})
                timerThread.PercentChartMEMPHY.set_title({'name': percentName})  
                timerThread.PercentChartMEMPHY.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
        elif self.vregUnderTest == "SOC":
            dmmSheet = self.workBook.add_worksheet(str(sheetCount) + "IMON " + self.vregUnderTest)
            dmmSheetName =  "'" + str(sheetCount) + "IMON " + self.vregUnderTest + "'"
            plotName = self.vregUnderTest + " Error Delta"
            percentName = self.vregUnderTest + " Percent Error"
            
            if sheetCount == 1:
                timerThread.ImonChartSOC = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.ImonChartSOC.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.ImonChartSOC.set_y_axis({'name': 'Error (Amps)'})
                timerThread.ImonChartSOC.set_title({'name': plotName})  
                timerThread.ImonChartSOC.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
                timerThread.PercentChartSOC = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                timerThread.PercentChartSOC.set_x_axis({'name': 'Chroma Load (Amps)'})
                timerThread.PercentChartSOC.set_y_axis({'name': 'Percent Error'})
                timerThread.PercentChartSOC.set_title({'name': percentName})  
                timerThread.PercentChartSOC.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
                
        dmmSheet.set_column(8, 8, 33)

            
        headerCellFormat = self.workBook.add_format()
        headerCellFormat.set_font_size(18)
        headerCellFormat.set_bold()
        headerCellFormat.set_bg_color('#2c749e')

#        if sheetCount == 1:
#            self.imonChart = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
#            self.imonChart.set_x_axis({'name': 'Chroma Load (Amps)'})
#            self.imonChart.set_y_axis({'name': 'Error (Amps)'})
#            self.imonChart.set_title({'name': plotName})  
#            self.imonChart.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
#            
#            self.percentChart = self.workBook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
#            self.percentChart.set_x_axis({'name': 'Chroma Load (Amps)'})
#            self.percentChart.set_y_axis({'name': 'Percent Error'})
#            self.percentChart.set_title({'name': percentName})  
#            self.percentChart.set_size({'x_scale': 1.5, 'y_scale': 1.5}) 
#        
        for i in range(0,40):
            self.plotSheet.write(0, i, "", headerCellFormat)
            self.plotSheet.write(46, i, "", headerCellFormat)
            self.plotSheet.write(92, i, "", headerCellFormat)
            self.plotSheet.write(138, i, "", headerCellFormat)
            self.plotSheet.write(185, i, "", headerCellFormat)
        self.plotSheet.write(0, 2, "GFX Plots High Current Calibration", headerCellFormat)
        self.plotSheet.write(46, 2, "GFX Plots Low Current Calibration", headerCellFormat)
        self.plotSheet.write(92, 2, "CPU Plots", headerCellFormat)
        self.plotSheet.write(138, 2, "MEMPHY Plots", headerCellFormat)
        self.plotSheet.write(185, 2, "SOC Plots", headerCellFormat)

        
        self.write_Measurement_Header(dmmSheet)
        self.rowIndex = self.rowIndex + 1

        chromaLoad = 0
        
        self.misc.setChromaCurrent(0.0)
        self.misc.loadOn()
        time.sleep(2)
        
        self.dsmcdbg.setOutputVoltage(voutHex)
        
        for voltIndex in range(0,2):
                        
            if voltIndex == 1:
                self.dsmcdbg.setOutputVoltage(0x0070)
                self.misc.setChromaCurrent(0.0)
                time.sleep(2)

            loopIndex = 0
            loopCount = 25
            for loopIndex in range(0, loopCount):
                
                self.misc.loadOn()
                time.sleep(2)

                if loopIndex <= 5:
                    imonVoltage = self.voltMeter.getVoltage(4)
                elif (loopIndex > 5) and (loopIndex > 10):
                    imonVoltage = self.voltMeter.getVoltage(2)
                else:
                    imonVoltage = self.voltMeter.getVoltage(1)
    
                readBackLoad = self.misc.readBackLoad()
                    
                gain = self.dsmcdbg.readGainRegister(hiLow)
                intelliphaseTemp = self.dsmcdbg.readIntelliTempRegister()
                
                if (self.vregUnderTest) == "GFX" or (self.vregUnderTest == "CPU"):
                    print "no ultra-fine"
                else:
                    ultraFine = self.dsmcdbg.readUltraFine()
                
                fineOffset = self.dsmcdbg.readOffsetRegister()
            
                vout = self.voltMeter.getVout(1)
                dmmSheet.write(self.rowIndex, 0, vout)
                dmmSheet.write(self.rowIndex, 1, float(readBackLoad))
                dmmSheet.write(self.rowIndex, 2, imonVoltage)
                dmmSheet.write(self.rowIndex, 3, "=C" + str(self.rowIndex+1) + " * Summary!" + timerThread.ScalingFactorCell)
                            
                dmmSheet.write(self.rowIndex, 4, "=D" + str(self.rowIndex+1) + " - " + "B" + str(self.rowIndex+1))
                dmmSheet.write(self.rowIndex, 5, "=E" + str(self.rowIndex+1) +  " - Summary!" + timerThread.NoLoadOffsetCell)
                dmmSheet.write(self.rowIndex, 6, "=F" + str(self.rowIndex+1) +  "/" + "B" + str(self.rowIndex+1) +  " * 100")
    
                #Tweaked Correction
                if voltIndex == 0: 
                    dmmSheet.write(self.rowIndex, 7, "=F" + str(self.rowIndex+1) + " - Summary!" + timerThread.TweakCell)
                else:
                    dmmSheet.write(self.rowIndex, 7, "=F" + str(self.rowIndex+1) + " - Summary!" + timerThread.TweakCell + "- " + "Summary!" + timerThread.OffsetCell + "- "  + "Summary!" + timerThread.CorrectionCell  + "* A" + str(self.rowIndex+1))
                dmmSheet.write(self.rowIndex, 8, "=H" + str(self.rowIndex+1) + "/ B" + str(self.rowIndex+1) + "* 100")

                dmmSheet.write(self.rowIndex, 9, gain)
                dmmSheet.write(self.rowIndex, 10, fineOffset)
                if (self.vregUnderTest) == "GFX" or (self.vregUnderTest == "CPU"):
                    dmmSheet.write(self.rowIndex, 11, "n/a")
                else:
                    dmmSheet.write(self.rowIndex, 11, '0x%0*X' % (4,ultraFine))
                dmmSheet.write(self.rowIndex, 12, intelliphaseTemp)
#                dmmSheet.write(self.rowIndex, 16, "TBD")
                         
                self.rowIndex = self.rowIndex + 1
                
                if (self.vregUnderTest == "GFX"):
                    if hiLow:
                        if loopIndex < 10:
                            chromaLoad = chromaLoad + 1.0
                        else:
                            chromaLoad = chromaLoad + 14.0
                            if loopIndex > 22:
                                self.misc.loadOff()
                                time.sleep(5)
                    else: 
                        if (loopIndex < 10):
                            chromaLoad = chromaLoad + 1.0
                        else:
                            chromaLoad = chromaLoad + 3.5
                elif (self.vregUnderTest == "CPU"):
                    if loopIndex < 10:
                        chromaLoad = chromaLoad + 0.5
                    else:
                        chromaLoad = chromaLoad + 6.5
                elif ((self.vregUnderTest == "MEMIO") or (self.vregUnderTest == "MEMPHY")):
                    if loopIndex < 10:
                        chromaLoad = chromaLoad + 0.5
                    else:
                        chromaLoad = chromaLoad + 2.0
                elif (self.vregUnderTest == "SOC"):
                    if loopIndex < 10:
                        chromaLoad = chromaLoad + 0.5
                    else:
                        chromaLoad = chromaLoad + 3.5
                    
                print "setting current to: " + str(chromaLoad)
                self.misc.setChromaCurrent(chromaLoad)
    
                print "loopIndex is: " + str(loopIndex)
                print "chroma load is: " + str(chromaLoad)
                
            if self.vregUnderTest == "GFX":
                if hiLow:
                    if voltIndex == 0: 
                        timerThread.ImonChartGFXHi.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$H$3:$H$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                        timerThread.PercentChartGFXHi.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$I$3:$I$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                    else:
                        timerThread.ImonChartGFXHi.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$H$30:$H$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                        timerThread.PercentChartGFXHi.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$I$30:$I$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                else:
                    if voltIndex == 0: 
                        timerThread.ImonChartGFXLow.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$H$3:$H$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                        timerThread.PercentChartGFXLow.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$I$3:$I$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                    else:
                        timerThread.ImonChartGFXLow.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$H$30:$H$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                        timerThread.PercentChartGFXLow.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$I$30:$I$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
            elif self.vregUnderTest == "CPU":
                if voltIndex == 0: 
                    timerThread.ImonChartCPU.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$H$3:$H$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                    timerThread.PercentChartCPU.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$I$3:$I$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                else:
                    timerThread.ImonChartCPU.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$H$30:$H$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                    timerThread.PercentChartCPU.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$I$30:$I$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
            elif self.vregUnderTest == "MEMPHY":
                if voltIndex == 0: 
                    timerThread.ImonChartMEMPHY.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$H$3:$H$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                    timerThread.PercentChartMEMPHY.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$I$3:$I$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                else:
                    timerThread.ImonChartMEMPHY.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$H$30:$H$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                    timerThread.PercentChartMEMPHY.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$I$30:$I$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
            elif self.vregUnderTest == "SOC":
                if voltIndex == 0: 
                    timerThread.ImonChartSOC.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$H$3:$H$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                    timerThread.PercentChartSOC.add_series({'categories': '=' + dmmSheetName + '!$B$3:$B$26','values':'=' + dmmSheetName + '!$I$3:$I$26', 'name': self.vregUnderTest + " 1V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                else:
                    timerThread.ImonChartSOC.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$H$30:$H$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})
                    timerThread.PercentChartSOC.add_series({'categories': '=' + dmmSheetName + '!$B$30:$B$53','values':'=' + dmmSheetName + '!$I$30:$I$53', 'name': self.vregUnderTest + " 0.7V" + "Pass " + str(sheetCount), 'marker': {'type': 'diamond', 'size': 5}})

#                if loopIndex >= loopCount:
#                    continueLoop = False
    
            self.misc.loadOff()
            if voltIndex == 0:
                time.sleep(180)
            self.rowIndex = self.rowIndex + 2
            chromaLoad = 0
            
#        #Insert Plots Here
#        self.plotSheet.insert_chart(timerThread.PlotCell, imonChart)    
#        self.plotSheet.insert_chart(timerThread.PercentCell, percentChart)    

        self.misc.loadOff()
        self.dsmcdbg.setOutputVoltage(0x00A0)

    def defaultWorkBookConfig(self, dmmSheet):

        dmmSheet.write('A1', self.excelFileName, self.headerCellFormat)
        dmmSheet.write('A2', version_Info, self.headerCellFormat) 
        dmmSheet.write('A3', "Config ID", self.headerCellFormat)
        dmmSheet.write('B3', self.configID)
        dmmSheet.write('A4', "PCBA S.N.", self.headerCellFormat)
        dmmSheet.write('B4', self.pcbaSN)
        dmmSheet.write('A5', "Product S.N.", self.headerCellFormat)
        dmmSheet.write('B5', self.productSN)
        dmmSheet.write('A8', "Date:", self.headerCellFormat)
        dmmSheet.write('A9', "Notes:", self.headerCellFormat)
        dmmSheet.write('B9', self.runNotes)
        dmmSheet.write('A10', "RLL", self.headerCellFormat)
        
        self.rowIndex = 11

    def write_Measurement_Header(self, dmmSheet):
        
        headerArray1 = ["Vout ","Chroma (Amps)", "VIMON (Volts)","IMON (Amps)",  "IMON - Chroma (Amps) ", "Error (Amps)", "% Error",  "No Load Adjusted Error (Amps)", "No Load Adjusted % Error", "Gain", "Fine Offset", "Ultra-Fine", "Intelliphase Temp", "Controller Temp", "", "", "", "", "", ""]
        headerArray2 = ["GFX_H Iout Offset", "GFX_H V-Slope Coeff.", "GFX_H V-Slope Offset.", "GFX_L Iout Offset", "GFX _L V-Slope Coeff.", "GFX _L V-Slope Offset.", "CPU Iout Offset", "CPU V-Slope Coeff.", "CPU V-Slope Offset.", "SOC Iout Offset", "SOC V-Slope Coeff.", "SOC V-Slope Offset.", "MEMP Iout Offset", "MEMP V-Slope Coeff.", "MEMP V-Slope Offset.", "", "", "", "", "", ""]
        
        for columnIndex in range (0, 18):
            dmmSheet.write(0, columnIndex, headerArray1[columnIndex], self.headerCellFormat)
            self.summarySheet.write(16, columnIndex, headerArray2[columnIndex], self.headerCellFormat)

    def recordOffsetOnly(self):
        
        target = 10.0
        hiLow = True
        
        vImon = self.voltMeter.getVoltage(1)
        print "VIMON is: " + str(vImon)
        imon = self.misc.calculateCurrent(vImon, hiLow)
        print "IMON is: " + str(imon)
        readBackLoad = self.misc.readBackLoad()
        offset_Error = (imon - (target + readBackLoad))
        print "Offset Error Method 3: " + str(offset_Error)
                    
        self.summarySheet.write(timerThread.TweakCell, offset_Error)
                
        print ""
        self.logFilePointer.write("Offset Error at Min Load: " + str(offset_Error) + "\n")
        print "Offset Error at Min Load: " + str(offset_Error)
        print ""
            
        offset = self.dsmcdbg.readOffsetRegister()
        self.logFilePointer.write("Fine offset setting is now: " + str(offset) + "\n")
        print "Fine offset setting is now: " + str(offset)
        print ""
            
        return 

    def quitThread(self):
        
#        self.workBook.close()
        self.voltMeter.close()
        
#        if self.logFile:
#            self.logFilePointer.close()
        
        self.misc.loadOff()
                
