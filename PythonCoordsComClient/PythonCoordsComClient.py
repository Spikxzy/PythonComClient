import datetime
import pythoncom
import win32com.client
import unittest
import logging
import pywintypes
from comtypes.client import CreateObject
from datetime import datetime


def Init():
    ComObject = CreateObject('PythonComObject', clsctx=None, machine=None, interface=None, dynamic=False, pServerInfo=None)
    coords = ComObject.Coordinates()

    for coordinate in coords:
        print("X: " + coordinate.getX() + " - Y: " + coordinate.getY())
    

if __name__ == "__main__":
    Init()