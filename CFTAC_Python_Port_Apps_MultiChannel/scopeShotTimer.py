import sys
from PyQt5.QtWidgets import (QLabel, QLineEdit, QPushButton, QDialog, QGridLayout, QLCDNumber)
from PyQt5.QtGui import QIcon, QPainter, QPen, QFont
from PyQt5 import QtCore , QtGui
 
import datetime
from time import strftime
 #GIT Version

class ScopeShotTimer(QDialog):
    def __init__(self):
        super(ScopeShotTimer, self).__init__()

        self.initUI()

    def initUI(self):
        
        self.count = 0
        
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.Time)
        self.timer.start(1000)
 
        self.lcd = QLCDNumber(self)
        # get the palette
        palette = self.lcd.palette()
        
        # foreground color
        palette.setColor(palette.WindowText, QtGui.QColor(85, 85, 255))
        # background color
        palette.setColor(palette.Background, QtGui.QColor(0, 170, 255))
        # "light" border
        palette.setColor(palette.Light, QtGui.QColor(255, 0, 0))
        # "dark" border
        palette.setColor(palette.Dark, QtGui.QColor(0, 255, 0))
        
        # set the palette
        self.lcd.setPalette(palette)    
        self.lcd.setMinimumHeight(100)
        self.lcd.setMinimumWidth(148)
        

        self.lcd.display(self.count)
 
        diologGrid      = QGridLayout()
        diologGrid.addWidget(self.lcd, 0, 0)

        self.okButton = QPushButton('Capture Scope Shot', self)
        self.okButton.clicked[bool].connect(self.close)
        diologGrid.addWidget(self.okButton, 0, 1)

        self.setLayout(diologGrid)
 
         
        self.setGeometry(200, 200, 150, 100)
        self.setWindowTitle('Scope Capture')
        self.setWindowIcon(QIcon('xbox_icon.ico')) 
 
    def Time(self):
#        self.lcd.display(strftime("%H"+":"+"%M"+":"+"%S"))
        self.count += 1
        self.lcd.display(self.count)
        if self.count == 120:
            print '\a'
         