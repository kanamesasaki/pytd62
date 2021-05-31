import math
import pandas as pd
import sys, clr
sys.path.append("C:/Windows/Microsoft.NET/assembly/GAC_MSIL/OpenTDv62/ReplaceMe")
clr.AddReference("OpenTDv62")
import OpenTDv62
clr.AddReference('System')
from System.Collections.Generic import Dictionary
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
    rectangle.Comment = data['Comment']
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
    cylinder.Comment = data['Comment']
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
    disk.Comment = data['Comment']
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
    cone.Comment = data['Comment']
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
    torus.Comment = data['Comment']
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
    scarfedcylinder.Comment = data['Comment']
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
    sphere.Comment = data['Comment']
    sphere.Update()
