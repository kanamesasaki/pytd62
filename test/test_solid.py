# -*- coding: utf-8 -*-
import sys, os, clr
import unittest
import math
import numpy as np
import pandas as pd
fileobj = open("path.txt", "r", encoding="utf_8")
sys.path.append(fileobj.readline().strip())
clr.AddReference("OpenTDv62")
import OpenTDv62
sys.path.append(os.path.abspath('../'))
from pytd62 import solid

td = OpenTDv62.ThermalDesktop()
td.Connect()

class TestSurface(unittest.TestCase):
    def test_create_solids(self):
        file_name = 'import_solid.csv'
        df = pd.read_csv(file_name)
        solid.create_solids(td, file_name)
        # test solidbrick
        solidbricks = td.GetSolidBricks()
        actual = [solidbricks[0].XMax.GetValueSI(), solidbricks[0].YMax.GetValueSI(), solidbricks[0].ZMax.GetValueSI()]
        expected = [df['Xmax [mm]'][0]*0.001, df['Ymax [mm]'][0]*0.001, df['Zmax [mm]'][0]*0.001]
        np.testing.assert_allclose(expected, actual)
        # test solidcylinder
        solidcylinders = td.GetSolidCylinders()
        actual = [solidcylinders[0].Rmax.GetValueSI(), solidcylinders[0].Rmin.GetValueSI(), solidcylinders[0].Height.GetValueSI(), solidcylinders[0].StartAngle.GetValueSI(), solidcylinders[0].EndAngle.GetValueSI()]
        expected = [df['Rmax [mm]'][1]*0.001, df['Rmin [mm]'][1]*0.001, df['Height [mm]'][1]*0.001, df['Start angle [deg]'][1], df['End angle [deg]'][1]]
        # test solidsphere
        solidsphere = td.GetSolidSpheres()
        actual = [solidsphere[0].Rmax.GetValueSI(), solidsphere[0].Rmin.GetValueSI(), solidsphere[0].StartAngle.GetValueSI(), solidsphere[0].EndAngle.GetValueSI(), solidsphere[0].Bmin.GetValueSI(), solidsphere[0].Bmax.GetValueSI()]
        expected = [df['Rmax [mm]'][2]*0.001, df['Rmin [mm]'][2]*0.001, df['Start angle [deg]'][2], df['End angle [deg]'][2], df['Bmin [deg]'][2], df['Bmax [deg]'][2]]
        # test solidcone
        solidcone = td.GetSolidCones()
        actual = [solidcone[0].BaseRmax.GetValueSI(), solidcone[0].BaseRmin.GetValueSI(), solidcone[0].TopRmax.GetValueSI(), solidcone[0].TopRmin.GetValueSI(), solidcone[0].Height.GetValueSI(), solidcone[0].StartAngle.GetValueSI(), solidcone[0].EndAngle.GetValueSI()]
        expected = [df['Base Rmax [mm]'][3]*0.001, df['Base Rmin [mm]'][3]*0.001, df['Top Rmax [mm]'][3]*0.001, df['Top Rmin [mm]'][3]*0.001, df['Height [mm]'][3]*0.001, df['Start angle [deg]'][3], df['End angle [deg]'][3]]
        td.ZoomExtents()
        
if __name__ == "__main__":
    unittest.main()
