import math
import pandas as pd
import sys, clr
sys.path.append("C:/Windows/Microsoft.NET/assembly/GAC_MSIL/OpenTDv62/ReplaceMe")
clr.AddReference("OpenTDv62")
import OpenTDv62
clr.AddReference('System')
from System.Collections.Generic import Dictionary
from System.Collections.Generic import List
from pytd62 import common

def solid_analysisgroup(s0: str, s1: str, s2: str, s3: str, s4: str, s5: str):
    analgroup = Dictionary[str, OpenTDv62.RadCAD.FdSolid.AnalysisGroupSolidInfo]()
    faces = [s0, s1, s2, s3, s4, s5]
    groups = set(faces)
    for group in groups:
        active = List[int]()
        active.Add(1 if s0==group else 0)
        active.Add(1 if s1==group else 0)
        active.Add(1 if s2==group else 0)
        active.Add(1 if s3==group else 0)
        active.Add(1 if s4==group else 0)
        active.Add(1 if s5==group else 0)
        analgroup.Add(group, OpenTDv62.RadCAD.FdSolid.AnalysisGroupSolidInfo(group, OpenTDv62.RadCAD.FdSolid.RcFdSolidData.Active.OUTSIDE, active))
    return analgroup

def create_solids(td: OpenTDv62.ThermalDesktop, file_name: str):
    df = pd.read_csv(file_name)
    for index, data in df.iterrows():
        if data['Element type'] == 'SolidBrick':
            create_solidbrick(td, data)
        elif data['Element type'] == 'SolidCylinder':
            create_solidcylinder(td, data)
        elif data['Element type'] == 'SolidSphere':
            create_solidsphere(td, data)
        elif data['Element type'] == 'SolidCone':
            create_solidcone(td, data)
        else:
            raise ValueError('unexpected element type')

def create_solidbrick(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    solidbrick = td.CreateSolidBrick()
    solidbrick.StartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    solidbrick.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    solidbrick.StartId = 1
    solidbrick.XMax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Xmax [mm]']*0.001)
    solidbrick.YMax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Ymax [mm]']*0.001)
    solidbrick.ZMax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Zmax [mm]']*0.001)
    solidbrick.BreakdownU.Num = data['Breakdown U']
    solidbrick.BreakdownV.Num = data['Breakdown V']
    solidbrick.BreakdownW.Num = data['Breakdown W']
    if data['Node positions'] == 'EDGE':
        solidbrick.NodeType = OpenTDv62.RadCAD.FdSolid.RcFdSolidData.NodePositionsType.EDGE
    else:
        solidbrick.NodeType = OpenTDv62.RadCAD.FdSolid.RcFdSolidData.NodePositionsType.CENTERED
    opt_properties = List[str]()
    opt_properties.Add(data['XMIN optical']) # GMIN
    opt_properties.Add(data['XMAX optical']) # GMAX
    opt_properties.Add(data['YMIN optical']) # RMIN
    opt_properties.Add(data['YMAX optical']) # RMAX
    opt_properties.Add(data['ZMIN optical']) # HMIN
    opt_properties.Add(data['ZMAX optical']) # HMAX
    solidbrick.InsideOpticalProperties = opt_properties
    solidbrick.OutsideOpticalProperties = opt_properties
    solidbrick.ThermoMaterial = data['Material']
    solidbrick.BaseTrans.entry[0][0] = new_base[0,0]
    solidbrick.BaseTrans.entry[1][0] = new_base[1,0]
    solidbrick.BaseTrans.entry[2][0] = new_base[2,0]
    solidbrick.BaseTrans.entry[0][1] = new_base[0,1]
    solidbrick.BaseTrans.entry[1][1] = new_base[1,1]
    solidbrick.BaseTrans.entry[2][1] = new_base[2,1]
    solidbrick.BaseTrans.entry[0][2] = new_base[0,2]
    solidbrick.BaseTrans.entry[1][2] = new_base[1,2]
    solidbrick.BaseTrans.entry[2][2] = new_base[2,2]
    solidbrick.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    solidbrick.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    solidbrick.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    solidbrick.AnalysisGroups = solid_analysisgroup(data['XMIN group'], data['XMAX group'], data['YMIN group'], data['YMAX group'], data['ZMIN group'], data['ZMAX group'])
    solidbrick.ColorIndex = data['Color']
    solidbrick.Comment = str(data['Comment'])
    solidbrick.Update()

def create_solidcylinder(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    solidcylinder = td.CreateSolidCylinder()
    solidcylinder.StartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    solidcylinder.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    solidcylinder.StartId = 1
    solidcylinder.Rmax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Rmax [mm]']*0.001)
    solidcylinder.Rmin = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Rmin [mm]']*0.001)
    solidcylinder.Height = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Height [mm]']*0.001)
    solidcylinder.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle [deg]']*1.0)
    solidcylinder.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle [deg]']*1.0)
    solidcylinder.BreakdownU.Num = data['Breakdown U']
    solidcylinder.BreakdownV.Num = data['Breakdown V']
    solidcylinder.BreakdownW.Num = data['Breakdown W']
    if data['Node positions'] == 'EDGE':
        solidcylinder.NodeType = OpenTDv62.RadCAD.FdSolid.RcFdSolidData.NodePositionsType.EDGE
    else:
        solidcylinder.NodeType = OpenTDv62.RadCAD.FdSolid.RcFdSolidData.NodePositionsType.CENTERED
    opt_properties = List[str]()
    opt_properties.Add(data['GMIN optical']) # GMIN
    opt_properties.Add(data['GMAX optical']) # GMAX
    opt_properties.Add(data['RMIN optical']) # RMIN
    opt_properties.Add(data['RMAX optical']) # RMAX
    opt_properties.Add(data['HMIN optical']) # HMIN
    opt_properties.Add(data['HMAX optical']) # HMAX
    solidcylinder.InsideOpticalProperties = opt_properties
    solidcylinder.OutsideOpticalProperties = opt_properties
    solidcylinder.ThermoMaterial = data['Material']
    solidcylinder.BaseTrans.entry[0][0] = new_base[0,0]
    solidcylinder.BaseTrans.entry[1][0] = new_base[1,0]
    solidcylinder.BaseTrans.entry[2][0] = new_base[2,0]
    solidcylinder.BaseTrans.entry[0][1] = new_base[0,1]
    solidcylinder.BaseTrans.entry[1][1] = new_base[1,1]
    solidcylinder.BaseTrans.entry[2][1] = new_base[2,1]
    solidcylinder.BaseTrans.entry[0][2] = new_base[0,2]
    solidcylinder.BaseTrans.entry[1][2] = new_base[1,2]
    solidcylinder.BaseTrans.entry[2][2] = new_base[2,2]
    solidcylinder.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    solidcylinder.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    solidcylinder.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    solidcylinder.AnalysisGroups = solid_analysisgroup(data['GMIN group'], data['GMAX group'], data['RMIN group'], data['RMAX group'], data['HMIN group'], data['HMAX group'])
    solidcylinder.ColorIndex = data['Color']
    solidcylinder.Comment = str(data['Comment'])
    solidcylinder.Update()
    
def create_solidsphere(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    solidsphere = td.CreateSolidSphere()
    solidsphere.StartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    solidsphere.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    solidsphere.StartId = 1
    solidsphere.Rmax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Rmax [mm]']*0.001)
    solidsphere.Rmin = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Rmin [mm]']*0.001)
    solidsphere.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle [deg]']*1.0)
    solidsphere.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle [deg]']*1.0)
    solidsphere.Bmin = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Bmin [deg]']*1.0)
    solidsphere.Bmax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Bmax [deg]']*1.0)
    solidsphere.BreakdownU.Num = data['Breakdown U']
    solidsphere.BreakdownV.Num = data['Breakdown V']
    solidsphere.BreakdownW.Num = data['Breakdown W']
    if data['Node positions'] == 'EDGE':
        solidsphere.NodeType = OpenTDv62.RadCAD.FdSolid.RcFdSolidData.NodePositionsType.EDGE
    else:
        solidsphere.NodeType = OpenTDv62.RadCAD.FdSolid.RcFdSolidData.NodePositionsType.CENTERED
    opt_properties = List[str]()
    opt_properties.Add(data['GMIN optical']) # GMIN
    opt_properties.Add(data['GMAX optical']) # GMAX
    opt_properties.Add(data['RMIN optical']) # RMIN
    opt_properties.Add(data['RMAX optical']) # RMAX
    opt_properties.Add(data['BMIN optical']) # HMIN
    opt_properties.Add(data['BMAX optical']) # HMAX
    solidsphere.InsideOpticalProperties = opt_properties
    solidsphere.OutsideOpticalProperties = opt_properties
    solidsphere.ThermoMaterial = data['Material']
    solidsphere.BaseTrans.entry[0][0] = new_base[0,0]
    solidsphere.BaseTrans.entry[1][0] = new_base[1,0]
    solidsphere.BaseTrans.entry[2][0] = new_base[2,0]
    solidsphere.BaseTrans.entry[0][1] = new_base[0,1]
    solidsphere.BaseTrans.entry[1][1] = new_base[1,1]
    solidsphere.BaseTrans.entry[2][1] = new_base[2,1]
    solidsphere.BaseTrans.entry[0][2] = new_base[0,2]
    solidsphere.BaseTrans.entry[1][2] = new_base[1,2]
    solidsphere.BaseTrans.entry[2][2] = new_base[2,2]
    solidsphere.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    solidsphere.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    solidsphere.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    solidsphere.AnalysisGroups = solid_analysisgroup(data['GMIN group'], data['GMAX group'], data['RMIN group'], data['RMAX group'], data['BMIN group'], data['BMAX group'])
    solidsphere.ColorIndex = data['Color']
    solidsphere.Comment = str(data['Comment'])
    solidsphere.Update()
    
def create_solidcone(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    new_base = common.rotate_base_euler(data['rot Z [deg]']*math.pi/180, data['rot Y [deg]']*math.pi/180, data['rot X [deg]']*math.pi/180)
    solidcone = td.CreateSolidCone()
    solidcone.StartSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    solidcone.CondSubmodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    solidcone.StartId = 1
    solidcone.BaseRmax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Base Rmax [mm]']*0.001)
    solidcone.BaseRmin = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Base Rmin [mm]']*0.001)
    solidcone.TopRmax = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Top Rmax [mm]']*0.001)
    solidcone.TopRmin = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Top Rmin [mm]']*0.001)
    solidcone.Height = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](data['Height [mm]']*0.001)
    solidcone.StartAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['Start angle [deg]']*1.0)
    solidcone.EndAngle = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Angle](data['End angle [deg]']*1.0)
    solidcone.BreakdownU.Num = data['Breakdown U']
    solidcone.BreakdownV.Num = data['Breakdown V']
    solidcone.BreakdownW.Num = data['Breakdown W']
    if data['Node positions'] == 'EDGE':
        solidcone.NodeType = OpenTDv62.RadCAD.FdSolid.RcFdSolidData.NodePositionsType.EDGE
    else:
        solidcone.NodeType = OpenTDv62.RadCAD.FdSolid.RcFdSolidData.NodePositionsType.CENTERED
    opt_properties = List[str]()
    opt_properties.Add(data['GMIN optical']) # GMIN
    opt_properties.Add(data['GMAX optical']) # GMAX
    opt_properties.Add(data['RMIN optical']) # RMIN
    opt_properties.Add(data['RMAX optical']) # RMAX
    opt_properties.Add(data['HMIN optical']) # HMIN
    opt_properties.Add(data['HMAX optical']) # HMAX
    solidcone.InsideOpticalProperties = opt_properties
    solidcone.OutsideOpticalProperties = opt_properties
    solidcone.ThermoMaterial = data['Material']
    solidcone.BaseTrans.entry[0][0] = new_base[0,0]
    solidcone.BaseTrans.entry[1][0] = new_base[1,0]
    solidcone.BaseTrans.entry[2][0] = new_base[2,0]
    solidcone.BaseTrans.entry[0][1] = new_base[0,1]
    solidcone.BaseTrans.entry[1][1] = new_base[1,1]
    solidcone.BaseTrans.entry[2][1] = new_base[2,1]
    solidcone.BaseTrans.entry[0][2] = new_base[0,2]
    solidcone.BaseTrans.entry[1][2] = new_base[1,2]
    solidcone.BaseTrans.entry[2][2] = new_base[2,2]
    solidcone.BaseTrans.entry[0][3] = data['X [mm]']*0.001
    solidcone.BaseTrans.entry[1][3] = data['Y [mm]']*0.001
    solidcone.BaseTrans.entry[2][3] = data['Z [mm]']*0.001
    solidcone.AnalysisGroups = solid_analysisgroup(data['GMIN group'], data['GMAX group'], data['RMIN group'], data['RMAX group'], data['HMIN group'], data['HMAX group'])
    solidcone.ColorIndex = data['Color']
    solidcone.Comment = str(data['Comment'])
    solidcone.Update()

def volume_solid(solid):
    if isinstance(surface, OpenTDv62.RadCAD.FdSolid.SolidBrick):
        area = volume_solidbrick(solid)
    elif isinstance(surface, OpenTDv62.RadCAD.FdSolid.SolidCylinder):
        area = volume_solidcylinder(solid)
    elif isinstance(surface, OpenTDv62.RadCAD.FdSolid.SolidCone):
        area = volume_solidcone(solid)
    elif isinstance(surface, OpenTDv62.RadCAD.FdSolid.SolidSphere):
        area = volume_solidsphere(solid)
    else:
        raise ValueError('unexpected element type')
    return area

def heat_capacity_solid(td, solid):
    volume = volume_solid(solid)
    material = td.GetThermoProps(solid.ThermoMaterial)
    rho = material.Density.GetValueSI()
    scale = material.DensityMult 
    cp = material.SpecificHeat.GetValueSI()
    heat_capacity = volume * rho * scale * cp
    return heat_capacity

def volume_solidbrick(solidbrick: OpenTDv62.RadCAD.FdSolid.SolidBrick):
    volume = solidbrick.XMax.GetValueSI() * solidbrick.YMax.GetValueSI() * solidbrick.ZMax.GetValueSI()
    return volume

def volume_solidcylinder(solidcylinder: OpenTDv62.RadCAD.FdSolid.SolidCylinder):
    angle = abs(solidcylinder.EndAngle.GetValueSI() - solidcylinder.StartAngle.GetValueSI())
    volume = math.pi * (solidcylinder.Rmax.GetValueSI()**2 - solidcylinder.Rmin.GetValueSI()**2) * angle/360 * solidcylinder.Height.GetValueSI()
    return volume

def volume_solidcone(solidcone: OpenTDv62.RadCAD.FdSolid.SolidCone):
    rmax_diff = solidcone.TopRmax.GetValueSI() - solidcone.BaseRmax.GetValueSI()
    rmin_diff = solidcone.TopRmin.GetValueSI() - solidcone.BaseRmin.GetValueSI()
    volume_max = math.pi * solidcone.Height.GetValueSI() * (rmax_diff**2/3 + solidcone.BaseRmax.GetValueSI()*rmax_diff + solidcone.BaseRmax.GetValueSI()**2)
    volume_min = math.pi * solidcone.Height.GetValueSI() * (rmin_diff**2/3 + solidcone.BaseRmax.GetValueSI()*rmin_diff + solidcone.BaseRmin.GetValueSI()**2)
    angle = abs(solidcone.EndAngle.GetValueSI() - solidcone.StartAngle.GetValueSI())
    volume = (volume_max - volume_min) * angle/360
    return volume

def volume_solidsphere(solidbrick: OpenTDv62.RadCAD.FdSolid.SolidBrick):
    angle = abs(solidcone.EndAngle.GetValueSI() - solidcone.StartAngle.GetValueSI())
    bmax_rad = sphere.Bmax*math.pi/180
    bmin_rad = sphere.Bmin*math.pi/180
    volume = 2/3*math.pi*(sphere.Rmax**3-sphere.Rmin**3)*(-math.cos(bmax_rad)+math.cos(bmin_rad))
    return volume