import math
import pandas as pd
import numpy as np
import sys, clr
fileobj = open("path.txt", "r", encoding="utf_8")
sys.path.append(fileobj.readline().strip())
clr.AddReference("OpenTDv62")
import OpenTDv62
clr.AddReference('System')
from System.Collections.Generic import Dictionary
from System.Collections.Generic import List
from pytd62 import common

def surface_analysisgroup(top: str, bottom: str):
    analgroup = Dictionary[str, OpenTDv62.RadCAD.AnalysisGroupSurfaceInfo]()
    if top==bottom:
        analgroup.Add(top, OpenTDv62.RadCAD.AnalysisGroupSurfaceInfo(top, OpenTDv62.RadCAD.RcEntityData.Active.BOTH))
    else:
        analgroup.Add(top, OpenTDv62.RadCAD.AnalysisGroupSurfaceInfo(top, OpenTDv62.RadCAD.RcEntityData.Active.TOP))
        analgroup.Add(bottom, OpenTDv62.RadCAD.AnalysisGroupSurfaceInfo(bottom, OpenTDv62.RadCAD.RcEntityData.Active.BOTTOM))
    return analgroup

def create_surfaces(td: OpenTDv62.ThermalDesktop, file_name: str):
    df = pd.read_csv(file_name)
    for index, data in df.iterrows():
        if data['Element type'] == 'Rectangle':
            create_rectangle(td, data)
        elif data['Element type'] == 'Cylinder':
            create_cylinder(td, data)
        elif data['Element type'] == 'Disk':
            create_disk(td, data)
        elif data['Element type'] == 'Cone':
            create_cone(td, data)
        elif data['Element type'] == 'Torus':
            create_torus(td, data)
        elif data['Element type'] == 'ScarfedCylinder':
            create_scarfedcylinder(td, data)
        elif data['Element type'] == 'Sphere':
            create_sphere(td, data)
        elif data['Element type'] == 'Polygon':
            create_polygon(td, data)
        else:
            raise ValueError('unexpected element type')

def create_rectangle(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    rectangle = td.CreateRectangle()
    rectangle.TopStartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    rectangle.TopStartId = data['Start ID']
    rectangle.XMax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['XMax [mm]']*0.001)
    rectangle.YMax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['YMax [mm]']*0.001)
    rectangle.BreakdownU.Num = data['Breakdown U']
    rectangle.BreakdownV.Num = data['Breakdown V']
    if data['Node positions'] == 'EDGE':
        rectangle.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.EDGE
    else:
        rectangle.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.CENTERED
    rectangle.TopOpticalProp = data['Top optical property']
    rectangle.BotOpticalProp = data['Bottom optical property']
    rectangle.TopMaterial = data['Material']
    rectangle.TopThickness = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Thickness [mm]']*0.001)
    rectangle.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    rectangle.BaseTrans.entry[0][0] = new_base[0,0]
    rectangle.BaseTrans.entry[1][0] = new_base[1,0]
    rectangle.BaseTrans.entry[2][0] = new_base[2,0]
    rectangle.BaseTrans.entry[0][1] = new_base[0,1]
    rectangle.BaseTrans.entry[1][1] = new_base[1,1]
    rectangle.BaseTrans.entry[2][1] = new_base[2,1]
    rectangle.BaseTrans.entry[0][2] = new_base[0,2]
    rectangle.BaseTrans.entry[1][2] = new_base[1,2]
    rectangle.BaseTrans.entry[2][2] = new_base[2,2]
    rectangle.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    rectangle.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    rectangle.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    rectangle.AnalysisGroups = surface_analysisgroup(data['Top analysis group'], data['Bottom analysis group'])
    rectangle.ColorIndex = data['Color']
    rectangle.Comment = str(data['Comment'])
    rectangle.Update()

def create_cylinder(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    cylinder = td.CreateCylinder()
    cylinder.TopStartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    cylinder.TopStartId = data['Start ID']
    cylinder.Radius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Radius [mm]']*0.001)
    cylinder.Height = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Height [mm]']*0.001)
    cylinder.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle [deg]']*1.0)
    cylinder.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle [deg]']*1.0)
    cylinder.BreakdownU.Num = data['Breakdown U']
    cylinder.BreakdownV.Num = data['Breakdown V']
    if data['Node positions'] == 'EDGE':
        cylinder.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.EDGE
    else:
        cylinder.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.CENTERED
    cylinder.TopOpticalProp = data['Top optical property']
    cylinder.BotOpticalProp = data['Bottom optical property']
    cylinder.TopMaterial = data['Material']
    cylinder.TopThickness = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Thickness [mm]']*0.001)
    cylinder.BaseTrans.entry[0][0] = new_base[0,0]
    cylinder.BaseTrans.entry[1][0] = new_base[1,0]
    cylinder.BaseTrans.entry[2][0] = new_base[2,0]
    cylinder.BaseTrans.entry[0][1] = new_base[0,1]
    cylinder.BaseTrans.entry[1][1] = new_base[1,1]
    cylinder.BaseTrans.entry[2][1] = new_base[2,1]
    cylinder.BaseTrans.entry[0][2] = new_base[0,2]
    cylinder.BaseTrans.entry[1][2] = new_base[1,2]
    cylinder.BaseTrans.entry[2][2] = new_base[2,2]
    cylinder.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    cylinder.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    cylinder.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    cylinder.AnalysisGroups = surface_analysisgroup(data['Top analysis group'], data['Bottom analysis group'])
    cylinder.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    cylinder.ColorIndex = data['Color']
    cylinder.Comment = str(data['Comment'])
    cylinder.Update()
    
def create_disk(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    disk = td.CreateDisk()
    disk.TopStartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    disk.TopStartId = data['Start ID']
    disk.MaxRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Max radius [mm]']*0.001)
    disk.MinRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Min radius [mm]']*0.001)
    disk.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle [deg]']*1.0)
    disk.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle [deg]']*1.0)
    disk.BreakdownU.Num = data['Breakdown U']
    disk.BreakdownV.Num = data['Breakdown V']
    if data['Node positions'] == 'EDGE':
        disk.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.EDGE
    else:
        disk.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.CENTERED
    disk.TopOpticalProp = data['Top optical property']
    disk.BotOpticalProp = data['Bottom optical property']
    disk.TopMaterial = data['Material']
    disk.TopThickness = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Thickness [mm]']*0.001)
    disk.BaseTrans.entry[0][0] = new_base[0,0]
    disk.BaseTrans.entry[1][0] = new_base[1,0]
    disk.BaseTrans.entry[2][0] = new_base[2,0]
    disk.BaseTrans.entry[0][1] = new_base[0,1]
    disk.BaseTrans.entry[1][1] = new_base[1,1]
    disk.BaseTrans.entry[2][1] = new_base[2,1]
    disk.BaseTrans.entry[0][2] = new_base[0,2]
    disk.BaseTrans.entry[1][2] = new_base[1,2]
    disk.BaseTrans.entry[2][2] = new_base[2,2]
    disk.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    disk.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    disk.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    disk.AnalysisGroups = surface_analysisgroup(data['Top analysis group'], data['Bottom analysis group'])
    disk.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    disk.ColorIndex = data['Color']
    disk.Comment = str(data['Comment'])
    disk.Update()
    
def create_cone(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    cone = td.CreateCone()
    cone.TopStartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    cone.TopStartId = data['Start ID']
    cone.BaseRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Base radius [mm]']*0.001)
    cone.TopRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Top radius [mm]']*0.001)
    cone.Height = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Height [mm]']*0.001)
    cone.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle [deg]']*1.0)
    cone.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle [deg]']*1.0)
    cone.BreakdownU.Num = data['Breakdown U']
    cone.BreakdownV.Num = data['Breakdown V']
    if data['Node positions'] == 'EDGE':
        cone.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.EDGE
    else:
        cone.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.CENTERED
    cone.TopOpticalProp = data['Top optical property']
    cone.BotOpticalProp = data['Bottom optical property']
    cone.TopMaterial = data['Material']
    cone.TopThickness = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Thickness [mm]']*0.001)
    cone.BaseTrans.entry[0][0] = new_base[0,0]
    cone.BaseTrans.entry[1][0] = new_base[1,0]
    cone.BaseTrans.entry[2][0] = new_base[2,0]
    cone.BaseTrans.entry[0][1] = new_base[0,1]
    cone.BaseTrans.entry[1][1] = new_base[1,1]
    cone.BaseTrans.entry[2][1] = new_base[2,1]
    cone.BaseTrans.entry[0][2] = new_base[0,2]
    cone.BaseTrans.entry[1][2] = new_base[1,2]
    cone.BaseTrans.entry[2][2] = new_base[2,2]
    cone.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    cone.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    cone.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    cone.AnalysisGroups = surface_analysisgroup(data['Top analysis group'], data['Bottom analysis group'])
    cone.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    cone.ColorIndex = data['Color']
    cone.Comment = str(data['Comment'])
    cone.Update()
    
def create_torus(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    torus = td.CreateTorus()
    torus.TopStartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    torus.TopStartId = data['Start ID']
    torus.LargeRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Large radius [mm]']*0.001)
    torus.SmallRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Small radius [mm]']*0.001)
    torus.StartAngleLargeRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle large radius [deg]']*1.0)
    torus.EndAngleLargeRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle large radius [deg]']*1.0)
    torus.StartAngleSmallRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle small radius [deg]']*1.0)
    torus.EndAngleSmallRadius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle small radius [deg]']*1.0)
    torus.BreakdownU.Num = data['Breakdown U']
    torus.BreakdownV.Num = data['Breakdown V']
    if data['Node positions'] == 'EDGE':
        torus.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.EDGE
    else:
        torus.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.CENTERED
    torus.TopOpticalProp = data['Top optical property']
    torus.BotOpticalProp = data['Bottom optical property']
    torus.TopMaterial = data['Material']
    torus.TopThickness = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Thickness [mm]']*0.001)
    torus.BaseTrans.entry[0][0] = new_base[0,0]
    torus.BaseTrans.entry[1][0] = new_base[1,0]
    torus.BaseTrans.entry[2][0] = new_base[2,0]
    torus.BaseTrans.entry[0][1] = new_base[0,1]
    torus.BaseTrans.entry[1][1] = new_base[1,1]
    torus.BaseTrans.entry[2][1] = new_base[2,1]
    torus.BaseTrans.entry[0][2] = new_base[0,2]
    torus.BaseTrans.entry[1][2] = new_base[1,2]
    torus.BaseTrans.entry[2][2] = new_base[2,2]
    torus.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    torus.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    torus.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    torus.AnalysisGroups = surface_analysisgroup(data['Top analysis group'], data['Bottom analysis group'])
    torus.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    torus.ColorIndex = data['Color']
    torus.Comment = str(data['Comment'])
    torus.Update()
    
def create_scarfedcylinder(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    scarfedcylinder = td.CreateScarfedCylinder()
    scarfedcylinder.TopStartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    scarfedcylinder.TopStartId = data['Start ID']
    scarfedcylinder.Radius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Radius [mm]']*0.001)
    scarfedcylinder.Height = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Height [mm]']*0.001)
    scarfedcylinder.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle [deg]']*1.0)
    scarfedcylinder.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle [deg]']*1.0)
    scarfedcylinder.ScarfAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Scarf angle [deg]']*1.0)
    scarfedcylinder.BreakdownU.Num = data['Breakdown U']
    scarfedcylinder.BreakdownV.Num = data['Breakdown V']
    if data['Node positions'] == 'EDGE':
        scarfedcylinder.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.EDGE
    else:
        scarfedcylinder.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.CENTERED
    scarfedcylinder.TopOpticalProp = data['Top optical property']
    scarfedcylinder.BotOpticalProp = data['Bottom optical property']
    scarfedcylinder.TopMaterial = data['Material']
    scarfedcylinder.TopThickness = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Thickness [mm]']*0.001)
    scarfedcylinder.BaseTrans.entry[0][0] = new_base[0,0]
    scarfedcylinder.BaseTrans.entry[1][0] = new_base[1,0]
    scarfedcylinder.BaseTrans.entry[2][0] = new_base[2,0]
    scarfedcylinder.BaseTrans.entry[0][1] = new_base[0,1]
    scarfedcylinder.BaseTrans.entry[1][1] = new_base[1,1]
    scarfedcylinder.BaseTrans.entry[2][1] = new_base[2,1]
    scarfedcylinder.BaseTrans.entry[0][2] = new_base[0,2]
    scarfedcylinder.BaseTrans.entry[1][2] = new_base[1,2]
    scarfedcylinder.BaseTrans.entry[2][2] = new_base[2,2]
    scarfedcylinder.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    scarfedcylinder.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    scarfedcylinder.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    scarfedcylinder.AnalysisGroups = surface_analysisgroup(data['Top analysis group'], data['Bottom analysis group'])
    scarfedcylinder.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    scarfedcylinder.ColorIndex = data['Color']
    scarfedcylinder.Comment = str(data['Comment'])
    scarfedcylinder.Update()
    
def create_sphere(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    sphere = td.CreateSphere()
    sphere.TopStartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    sphere.TopStartId = data['Start ID']
    sphere.Radius = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Radius [mm]']*0.001)
    sphere.MinHeight = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Min height [mm]']*0.001)
    sphere.MaxHeight = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Max height [mm]']*0.001)
    sphere.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle [deg]']*1.0)
    sphere.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle [deg]']*1.0)
    sphere.BreakdownU.Num = data['Breakdown U']
    sphere.BreakdownV.Num = data['Breakdown V']
    if data['Node positions'] == 'EDGE':
        sphere.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.EDGE
    else:
        sphere.NodePositions = OpenTDv62.RadCAD.RcEntityData.NodePositionsType.CENTERED
    sphere.TopOpticalProp = data['Top optical property']
    sphere.BotOpticalProp = data['Bottom optical property']
    sphere.TopMaterial = data['Material']
    sphere.TopThickness = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Thickness [mm]']*0.001)
    sphere.BaseTrans.entry[0][0] = new_base[0,0]
    sphere.BaseTrans.entry[1][0] = new_base[1,0]
    sphere.BaseTrans.entry[2][0] = new_base[2,0]
    sphere.BaseTrans.entry[0][1] = new_base[0,1]
    sphere.BaseTrans.entry[1][1] = new_base[1,1]
    sphere.BaseTrans.entry[2][1] = new_base[2,1]
    sphere.BaseTrans.entry[0][2] = new_base[0,2]
    sphere.BaseTrans.entry[1][2] = new_base[1,2]
    sphere.BaseTrans.entry[2][2] = new_base[2,2]
    sphere.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    sphere.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    sphere.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    sphere.AnalysisGroups = surface_analysisgroup(data['Top analysis group'], data['Bottom analysis group'])
    sphere.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    sphere.ColorIndex = data['Color']
    sphere.Comment = str(data['Comment'])
    sphere.Update()

def create_polygon(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    edges = List[OpenTDv62.Point3d]()
    edges.Add(OpenTDv62.Point3d(data['X1 [mm]']*0.001, data['Y1 [mm]']*0.001, data['Z1 [mm]']*0.001))
    edges.Add(OpenTDv62.Point3d(data['X2 [mm]']*0.001, data['Y2 [mm]']*0.001, data['Z2 [mm]']*0.001))
    edges.Add(OpenTDv62.Point3d(data['X3 [mm]']*0.001, data['Y3 [mm]']*0.001, data['Z3 [mm]']*0.001))
    if not (math.isnan(data['X4 [mm]']) or math.isnan(data['Y4 [mm]']) or math.isnan(data['Z4 [mm]'])):
        edges.Add(OpenTDv62.Point3d(data['X4 [mm]']*0.001, data['Y4 [mm]']*0.001, data['Z4 [mm]']*0.001))
    polygon = td.CreatePolygon(edges)
    polygon.TopStartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    polygon.TopStartId = data['Start ID']
    polygon.BreakdownU.Num = 1
    polygon.BreakdownV.Num = 1
    polygon.TopOpticalProp = data['Top optical property']
    polygon.BotOpticalProp = data['Bottom optical property']
    polygon.TopMaterial = data['Material']
    polygon.TopThickness = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Thickness [mm]']*0.001)
    polygon.AnalysisGroups = surface_analysisgroup(data['Top analysis group'], data['Bottom analysis group'])
    polygon.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    polygon.ColorIndex = data['Color']
    polygon.Comment = str(data['Comment'])
    polygon.Update()
    
def area_surface(surface):
    if isinstance(surface, OpenTDv62.RadCAD.Rectangle):
        area = area_rectangle(surface)
    elif isinstance(surface, OpenTDv62.RadCAD.Cylinder):
        area = area_cylinder(surface)
    elif isinstance(surface, OpenTDv62.RadCAD.Disk):
        area = area_disk(surface)
    elif isinstance(surface, OpenTDv62.RadCAD.Cone):
        area = area_cone(surface)
    elif isinstance(surface, OpenTDv62.RadCAD.Torus):
        area = area_torus(surface)
    elif isinstance(surface, OpenTDv62.RadCAD.ScarfedCylinder):
        area = area_scarfedcylinder(surface)
    elif isinstance(surface, OpenTDv62.RadCAD.Sphere):
        area = area_sphere(surface)
    elif isinstance(surface, OpenTDv62.RadCAD.Polygon):
        area = area_polygon(surface)
    else:
        raise ValueError('unexpected element type')
    return area

def heat_capacity_surface(td, surface):
    area = area_surface(surface)
    volume = area * surface.TopThickness.GetValueSI()
    material = td.GetThermoProps(surface.TopMaterial)
    rho = material.Density.GetValueSI()
    scale = material.DensityMult 
    cp = material.SpecificHeat.GetValueSI()
    heat_capacity = volume * rho * scale * cp
    return heat_capacity

def area_disk(disk: OpenTDv62.RadCAD.Disk):
    angle = abs(disk.EndAngle.GetValueSI() - disk.StartAngle.GetValueSI())
    area = math.pi * (disk.MaxRadius.GetValueSI()**2 - disk.MinRadius.GetValueSI()**2) * angle/360.0
    return area

def area_rectangle(rectangle: OpenTDv62.RadCAD.Rectangle):
    area = rectangle.XMax.GetValueSI() * rectangle.YMax.GetValueSI()
    return area

def area_cylinder(cylinder: OpenTDv62.RadCAD.Cylinder):
    angle = abs(cylinder.EndAngle.GetValueSI() - cylinder.StartAngle.GetValueSI())
    area = 2 * math.pi * cylinder.Radius.GetValueSI() * angle/360 * cylinder.Height.GetValueSI()
    return area

def area_cone(cone: OpenTDv62.RadCAD.Cone):
    angle = abs(cone.EndAngle.GetValueSI() - cone.StartAngle.GetValueSI())
    rtop = cone.TopRadius.GetValueSI()
    rbase = cone.BaseRadius.GetValueSI()
    height = cone.Height.GetValueSI()
    area = math.pi * math.sqrt((rtop-rbase)**2+height**2) * (rbase+rtop) * angle/360
    return area

def area_torus(torus: OpenTDv62.RadCAD.Torus):
    r_large = torus.LargeRadius.GetValueSI()
    r_small = torus.SmallRadius.GetValueSI()
    start_angle_small = torus.StartAngleSmallRadius.GetValueSI() * math.pi/180
    end_angle_small = torus.EndAngleSmallRadius.GetValueSI() * math.pi/180
    if start_angle_small > end_angle_small:
        raise ValueError('StartAngleSmallRadius shall be smaller than EndAngleSmallRadius')
    angle_large = abs(torus.EndAngleLargeRadius.GetValueSI() - torus.StartAngleLargeRadius.GetValueSI())
    area = 2*math.pi*r_large*r_small*(end_angle_small-start_angle_small) + 2*math.pi*r_small**2*(math.sin(end_angle_small)-math.sin(start_angle_small))
    return area * angle_large/360

def area_lineartri(td: OpenTDv62.ThermalDesktop, lineartri: OpenTDv62.RadCAD.FEM.LinearTri):
    edge1 = td.GetNode(lineartri.AttachedNodeHandles[0])
    edge2 = td.GetNode(lineartri.AttachedNodeHandles[1])
    edge3 = td.GetNode(lineartri.AttachedNodeHandles[2])
    xyz1 = np.array([edge1.Origin.X.GetValueSI(), edge1.Origin.Y.GetValueSI(), edge1.Origin.Z.GetValueSI()])
    xyz2 = np.array([edge2.Origin.X.GetValueSI(), edge2.Origin.Y.GetValueSI(), edge2.Origin.Z.GetValueSI()])
    xyz3 = np.array([edge3.Origin.X.GetValueSI(), edge3.Origin.Y.GetValueSI(), edge3.Origin.Z.GetValueSI()])
    vec12 = xyz2 - xyz1
    vec13 = xyz3 - xyz1
    area = np.linalg.norm(np.cross(vec12, vec13), ord=2) * 0.5
    return area

def area_scarfedcylinder(scarfedcylinder: OpenTDv62.RadCAD.ScarfedCylinder):
    start = scarfedcylinder.StartAngle.GetValueSI()*math.pi/180
    end = scarfedcylinder.EndAngle.GetValueSI()*math.pi/180
    scarf = scarfedcylinder.ScarfAngle.GetValueSI()*math.pi/180
    radius = scarfedcylinder.Radius.GetValueSI()
    area = radius*scarfedcylinder.Height.GetValueSI()*(end-start) - radius**2*math.tan(scarf)*(math.sin(end)-math.sin(start))
    return area

def area_polygon(polygon: OpenTDv62.RadCAD.Polygon):
    points = []
    vectors = []
    num_vertices = len(polygon.Vertices)//2
    for i in range(num_vertices):
        poly = polygon.Vertices[i]
        point = np.array([poly.X.GetValueSI(),poly.Y.GetValueSI(),poly.Z.GetValueSI()])
        points.append(point)
        if i > 0:
            vectors.append(points[i]-points[0])
    normal = np.cross(vectors[0], vectors[1])
    area = 0.0
    for i in range(num_vertices-2):
        if not math.isclose(np.inner(normal, vectors[i+1]), 0.0, abs_tol=1.0e-6):
            raise ValueError('The vertex is not on the surface')
        area += np.linalg.norm(np.cross(vectors[i], vectors[i+1]), ord=2) * 0.5
    return area

def area_sphere(sphere: OpenTDv62.RadCAD.Sphere):
    angle = abs(sphere.EndAngle.GetValueSI() - sphere.StartAngle.GetValueSI())*math.pi/180
    height = abs(sphere.MaxHeight.GetValueSI() - sphere.MinHeight.GetValueSI())
    radius = sphere.Radius.GetValueSI()
    area = angle * height * radius
    return area
