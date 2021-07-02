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
from pytd62 import surface

class TestSurface(unittest.TestCase):
    def test_area_disk(self):
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        disk = td.CreateDisk()
        disk.MaxRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](20.0e-3)
        disk.MinRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](10.0e-3)
        disk.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](-30.0)
        disk.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](120.0)
        disk.Update()
        disks = td.GetDisks()
        actual = surface.area_disk(disks[0])
        expected = math.pi * (20.0e-3**2 - 10.0e-3**2) * 150/360
        self.assertAlmostEqual(expected, actual)
        td.Quit()
        
    def test_area_rectangle(self):
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        rectangle = td.CreateRectangle()
        rectangle.XMax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](10.0e-3)
        rectangle.YMax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](20.0e-3)
        rectangle.Update()
        rectangles = td.GetRectangles()
        actual = surface.area_rectangle(rectangles[0])
        expected = 20.0e-3 * 10.0e-3
        self.assertAlmostEqual(expected, actual)
        td.Quit()
        
    def test_area_cone(self):
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        cone = td.CreateCone()
        cone.BaseRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](2.0)
        cone.TopRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](1.0)
        cone.Height = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](math.sqrt(32.0)/2)
        cone.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](10.0)
        cone.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](150.0)
        cone.Update()
        cones = td.GetCones()
        actual = surface.area_cone(cones[0])
        expected = (math.pi*2.0*6.0 - math.pi*1.0*3.0) * 140.0/360.0
        self.assertAlmostEqual(expected, actual)
        td.Quit()
        
    def test_area_torus(self):
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        torus = td.CreateTorus()
        torus.LargeRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](5.0)
        torus.SmallRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](1.0)
        torus.StartAngleLargeRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](0.0)
        torus.EndAngleLargeRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](150.0)
        torus.StartAngleSmallRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](0.0)
        torus.EndAngleSmallRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](360.0)
        torus.Update()
        toruses = td.GetToruses()
        actual = surface.area_torus(toruses[0])
        expected = 4*math.pi**2*5.0*1.0 * 150.0/360.0
        self.assertAlmostEqual(expected, actual)
        td.Quit()
        
    def test_area_scarfedcylinder(self):
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        scarfedcylinder = td.CreateScarfedCylinder()
        scarfedcylinder.Radius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](1.0)
        scarfedcylinder.Height = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](5.0)
        scarfedcylinder.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](0.0)
        scarfedcylinder.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](180.0)
        scarfedcylinder.ScarfAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](30.0)
        scarfedcylinder.Update()
        scarfedcylinders = td.GetScarfedCylinders()
        actual = surface.area_scarfedcylinder(scarfedcylinders[0])
        expected = 2*math.pi*1.0*5.0/2
        self.assertAlmostEqual(expected, actual)
        td.Quit()

    def test_area_sphere(self):
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        sphere = td.CreateSphere()
        sphere.Radius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](5.0)
        sphere.MinHeight = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](3.0)
        sphere.MaxHeight = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](5.0)
        sphere.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](0.0)
        sphere.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](360.0)
        sphere.Update()
        spheres = td.GetSpheres()
        actual = surface.area_sphere(spheres[0])
        expected = math.pi*(4.0**2+2.0**2)
        self.assertAlmostEqual(expected, actual)
        td.Quit()

    def test_create_surfaces(self):
        td = OpenTDv62.ThermalDesktop()
        td.Connect()
        file_name = 'import_surface.csv'
        df = pd.read_csv(file_name)
        surface.create_surfaces(td, file_name)
        # test rectangle
        rectangles = td.GetRectangles()
        actual = [rectangles[0].XMax.GetValueSI(), rectangles[0].YMax.GetValueSI()]
        expected = [df['XMax [mm]'][0]*0.001, df['YMax [mm]'][0]*0.001]
        np.testing.assert_allclose(expected, actual)
        # test cylinder
        cylinders = td.GetCylinders()
        actual = [cylinders[0].Radius.GetValueSI(), cylinders[0].Height.GetValueSI(), cylinders[0].StartAngle.GetValueSI(), cylinders[0].EndAngle.GetValueSI()]
        expected = [df['Radius [mm]'][1]*0.001, df['Height [mm]'][1]*0.001, df['Start angle [deg]'][1], df['End angle [deg]'][1]]
        np.testing.assert_allclose(expected, actual)
        # test disk
        disks = td.GetDisks()
        actual = [disks[0].MaxRadius.GetValueSI(), disks[0].MinRadius.GetValueSI(), disks[0].StartAngle.GetValueSI(), disks[0].EndAngle.GetValueSI()]
        expected = [df['Max radius [mm]'][2]*0.001, df['Min radius [mm]'][2]*0.001, df['Start angle [deg]'][2], df['End angle [deg]'][2]]
        np.testing.assert_allclose(expected, actual)
        # test cone
        cones = td.GetCones()
        actual = [cones[0].BaseRadius.GetValueSI(), cones[0].TopRadius.GetValueSI(), cones[0].Height.GetValueSI(), cones[0].StartAngle.GetValueSI(), cones[0].EndAngle.GetValueSI()]
        expected = [df['Base radius [mm]'][3]*0.001, df['Top radius [mm]'][3]*0.001, df['Height [mm]'][3]*0.001, df['Start angle [deg]'][3], df['End angle [deg]'][3]]
        np.testing.assert_allclose(expected, actual)
        # test torus
        toruses = td.GetToruses()
        actual = [toruses[0].LargeRadius.GetValueSI(), toruses[0].SmallRadius.GetValueSI(), toruses[0].StartAngleLargeRadius.GetValueSI(), toruses[0].EndAngleLargeRadius.GetValueSI(), toruses[0].StartAngleSmallRadius.GetValueSI(), toruses[0].EndAngleSmallRadius.GetValueSI()]
        expected = [df['Large radius [mm]'][4]*0.001, df['Small radius [mm]'][4]*0.001, df['Start angle large radius [deg]'][4], df['End angle large radius [deg]'][4], df['Start angle small radius [deg]'][4], df['End angle small radius [deg]'][4]]
        np.testing.assert_allclose(expected, actual)
        # test scarfedcylinder
        scarfedcylinders = td.GetScarfedCylinders()
        actual = [scarfedcylinders[0].Radius.GetValueSI(), scarfedcylinders[0].Height.GetValueSI(), scarfedcylinders[0].StartAngle.GetValueSI(), scarfedcylinders[0].EndAngle.GetValueSI(), scarfedcylinders[0].ScarfAngle.GetValueSI()]
        expected = [df['Radius [mm]'][5]*0.001, df['Height [mm]'][5]*0.001, df['Start angle [deg]'][5], df['End angle [deg]'][5], df['Scarf angle [deg]'][5]]
        np.testing.assert_allclose(expected, actual)
        # test sphere
        spheres = td.GetSpheres()
        actual = [spheres[0].Radius.GetValueSI(), spheres[0].MinHeight.GetValueSI(), spheres[0].MaxHeight.GetValueSI(), spheres[0].StartAngle.GetValueSI(), spheres[0].EndAngle.GetValueSI()]
        expected = [df['Radius [mm]'][6]*0.001, df['Min height [mm]'][6]*0.001, df['Max height [mm]'][6]*0.001, df['Start angle [deg]'][6], df['End angle [deg]'][6]]
        np.testing.assert_allclose(expected, actual)
        td.ZoomExtents()
        
if __name__ == "__main__":
    unittest.main()
