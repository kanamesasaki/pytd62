# -*- coding: utf-8 -*-
import unittest
import imghdr
import math
import numpy as np
import pandas as pd
import sys, os, clr
sys.path.append("C:/Windows/Microsoft.NET/assembly/GAC_MSIL/OpenTDv62/v4.0_6.2.0.3__65e6d95ed5c2e178")
clr.AddReference("OpenTDv62")
import OpenTDv62
sys.path.append(os.path.abspath('../'))
from pytd62 import common

class TestCommon(unittest.TestCase):
    def test_rotx(self):
        theta = 30.0*math.pi/180
        expected = np.array([[1.0, 0.0, 0.0],
                             [0.0, math.sqrt(3.0)/2, -0.5],
                             [0.0, 0.5, math.sqrt(3.0)/2]])
        actual = common.rotx(theta)
        np.testing.assert_allclose(expected, actual)
        
    def test_roty(self):
        theta = 30.0*math.pi/180
        expected = np.array([[math.sqrt(3.0)/2, 0.0, 0.5],
                             [0.0, 1.0, 0.0],
                             [-0.5, 0.0, math.sqrt(3.0)/2]])
        actual = common.roty(theta)
        np.testing.assert_allclose(expected, actual)
        
    def test_rotz(self):
        theta = 30.0*math.pi/180
        expected = np.array([[math.sqrt(3.0)/2, -0.5, 0.0],
                             [0.5, math.sqrt(3.0)/2, 0.0],
                             [0.0, 0.0, 1.0]])
        actual = common.rotz(theta)
        np.testing.assert_allclose(expected, actual)
        
    def test_rotate_base_euler_ZYX(self):
        expected = np.array([[0.0, 0.0, 1.0],
                             [0.0, 1.0, 0.0],
                             [-1.0, 0.0, 0.0]])
        actual = common.rotate_base_euler(math.pi/2, math.pi/2, math.pi/2)
        np.testing.assert_allclose(expected, actual, atol=1.0E-10)
        
    def test_rotate_base_euler_ZXZ(self):
        expected = np.array([[0.0, 0.0, 1.0],
                             [0.0, -1.0, 0.0],
                             [1.0, 0.0, 0.0]])
        actual = common.rotate_base_euler(math.pi/2, math.pi/2, math.pi/2, order='ZXZ')
        np.testing.assert_allclose(expected, actual, atol=1.0E-10)
        
    def test_import_thermo_properties(self):
        file_name = 'import_thermo.csv'
        df = pd.read_csv('import_thermo.csv')
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        common.import_thermo_properties(td, file_name)
        expected = [int(df['Anisotropic'][0]), df['Conductivity [W/mK]'][0], df['Density [kg/m3]'][0], df['Specific heat [J/kgK]'][0]]
        thermo = td.GetThermoProps(df['Name'][0])
        actual = [thermo.Anisotropic, thermo.Conductivity.GetValueSI(), thermo.Density.GetValueSI(), thermo.SpecificHeat.GetValueSI()]
        np.testing.assert_allclose(expected, actual)
        expected = [int(df['Anisotropic'][1]), df['Conductivity [W/mK]'][1], df['ConductivityY [W/mK]'][1], df['ConductivityZ [W/mK]'][1], df['Density [kg/m3]'][1], df['Specific heat [J/kgK]'][1]]
        thermo = td.GetThermoProps(df['Name'][1])
        actual = [thermo.Anisotropic, thermo.Conductivity.GetValueSI(), thermo.ConductivityY.GetValueSI(), thermo.ConductivityZ.GetValueSI(), thermo.Density.GetValueSI(), thermo.SpecificHeat.GetValueSI()]
        np.testing.assert_allclose(expected, actual)
        td.Quit()
        
    def test_import_optical_properties(self):
        file_name = 'import_optical.csv'
        df = pd.read_csv('import_optical.csv')
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        common.import_optical_properties(td, file_name)
        expected = [df['Absorptivity'][0], df['Emissivity'][0]]
        optical = td.GetOpticalProps(df['Name'][0])
        actual = [optical.Alph, optical.Emis]
        np.testing.assert_allclose(expected, actual)
        td.Quit()
        
    def test_screenshot(self):
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        cylinder = td.CreateCylinder()
        cylinder.TopStartId = 1
        cylinder.Radius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](10.0E-3)
        cylinder.Height = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](30.0E-3)
        cylinder.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](0.0)
        cylinder.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](270.0)
        cylinder.BreakdownU.Num = int(3)
        cylinder.BreakdownV.Num = int(3)
        cylinder.Update()
        td.RestoreIsoView(0)
        td.ZoomExtents()
        
        # export all types of imageformat
        name = 'format_'
        format = 'png'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        expected = format
        actual = imghdr.what(filename)
        self.assertEqual(expected, actual)
        format = 'jpeg'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        expected = format
        actual = imghdr.what(filename)
        self.assertEqual(expected, actual)
        format = 'bmp'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        expected = format
        actual = imghdr.what(filename)
        self.assertEqual(expected, actual)
        format = 'tiff'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        expected = format
        actual = imghdr.what(filename)
        self.assertEqual(expected, actual)
        format = 'gif'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        expected = format
        actual = imghdr.what(filename)
        self.assertEqual(expected, actual)
        format = 'emf'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        format = 'wmf'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        format = 'exif'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        format = 'exif'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        format = 'icon'
        filename = name + format + '.' + format
        common.screenshot(td, name+format, imageformat=format)
        
        # export all types of visual style
        name = 'style_'
        modelstyle = 'WIRE_2D'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'WIRE'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'CONCEPTUAL'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'REALISTIC'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'SHADED'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'SHADED_W_EDGES'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'SHADES_OF_GREY'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'SKETCHY'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'THERMAL'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'THERMAL_PP'
        common.screenshot(td, name+modelstyle, style=modelstyle)
        modelstyle = 'XRAY'
        common.screenshot(td, name+modelstyle)
                
        # export all types of view
        name = 'view_'
        viewdirc = 'SW'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        viewdirc = 'SE'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        viewdirc = 'NE'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        viewdirc = 'NW'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        viewdirc = 'TOP'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        viewdirc = 'BOTTOM'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        viewdirc = 'FRONT'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        viewdirc = 'BACK'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        viewdirc = 'LEFT'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        viewdirc = 'RIGHT'
        common.screenshot(td, name+viewdirc, view=viewdirc)
        td.Quit()
        
if __name__ == "__main__":
    unittest.main()
