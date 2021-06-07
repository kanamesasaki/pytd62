import math
import numpy as np
import pandas as pd
import sys, clr
sys.path.append("C:/Windows/Microsoft.NET/assembly/GAC_MSIL/OpenTDv62/ReplaceMe")
clr.AddReference("OpenTDv62")
import OpenTDv62
clr.AddReference('System')
from System.Drawing.Imaging import ImageFormat
from System.Drawing import RectangleF
from System.Collections.Generic import List

def rotx(theta: float) -> np.ndarray:
    mat = np.array([[1.0, 0.0, 0.0],
                    [0.0, math.cos(theta), -math.sin(theta)],
                    [0.0, math.sin(theta), math.cos(theta)]])
    return mat

def roty(theta: float) -> np.ndarray:
    mat = np.array([[math.cos(theta), 0.0, math.sin(theta)],
                    [0.0, 1.0, 0.0],
                    [-math.sin(theta), 0.0, math.cos(theta)]])
    return mat

def rotz(theta: float) -> np.ndarray:
    mat = np.array([[math.cos(theta), -math.sin(theta), 0.0],
                    [math.sin(theta), math.cos(theta), 0.0],
                    [0.0, 0.0, 1.0]])
    return mat

def rotate_base_euler(ang1, ang2, ang3, order='ZYX', base=np.eye(3)):
    orders = ['XYX', 'XYZ', 'XZX', 'XZY',
              'YXY', 'YXZ', 'YZX', 'YZY',
              'ZXY', 'ZXZ', 'ZYX', 'ZYZ']
    
    if not order in orders: 
        raise ValueError('unexpected rotation order')
    
    angles = [ang1, ang2, ang3]
    base_trans = np.transpose(base)
    for i, axis in enumerate(order):
        if axis == 'X':
            base_trans = np.transpose(rotx(angles[i])) @ base_trans
        elif axis == 'Y':
            base_trans = np.transpose(roty(angles[i])) @ base_trans
        else:
            base_trans = np.transpose(rotz(angles[i])) @ base_trans
    
    return np.transpose(base_trans)

def get_number_of_nodes(td: OpenTDv62.ThermalDesktop, submodel_name: str='') -> int:
    nodes = td.GetNodes()
    if submodel_name == '':
        nodes = [inode for inode in nodes if submodel_name == inode.Submodel.Name]
    return len(nodes)

def get_node_handle(td, submodel_name, node_id, nodes: list=[]):
    if nodes == []:
        nodes = td.GetNodes()
    node_extracted = [inode for inode in nodes if submodel_name == inode.Submodel.Name and node_id == inode.Id]
    if len(node_extracted) == 1:
        return node_extracted[0].Handle
    elif len(node_extracted) == 0:
        raise ValueError('The requested node does not exist')
    elif len(node_extracted) > 1:
        raise ValueError('The requested node spec is assigned to multiple nodes')
    
def get_submodel_nodes(td: OpenTDv62.ThermalDesktop, submodel_name: str) -> list:
    nodes = td.GetNodes()
    submodel_nodes = [inode for inode in nodes if submodel_name == inode.Submodel.Name]
    # sorted_nodes = sorted(submodel_nodes, key=lambda node: node.id)
    return submodel_nodes
    
def renumber_submodel_nodes(td: OpenTDv62.ThermalDesktop, submodel_name: str):
    submodel_nodes = get_submodel_nodes(td, submodel_name)
    for i, inode in enumerate(submodel_nodes):
        inode.Id = i+1
        inode.Update()
        
def merge_submodel_nodes(td: OpenTDv62.ThermalDesktop, submodel_name: str, tolerance: float=0.1e-3):
    merge = OpenTDv62.MergeNodesOptionsData()
    merge.KeepMethod = OpenTDv62.MergeNodesOptionsData.KeepMethods.SMALLEST_NODE_ID
    nodes = List[str]()
    submodel_nodes = get_submodel_nodes(td, submodel_name)
    for node in submodel_nodes:
        nodes.Add(node.Handle)
    merge.NodeHandles = nodes
    merge.Tolerance = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength](tolerance)
    td.MergeNodes(merge)
    
def create_node(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    node = td.CreateNode()
    node.Submodel = OpenTDv62.SubmodelNameData(data['Submodel name'])
    node.Id = data['ID']
    node.InitialTemp = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Temp](data['Initial temperature [K]']*1.0)
    node.Origin = OpenTDv62.Point3d(data['X [mm]']*0.001, data['Y [mm]']*0.001, data['Z [mm]']*0.001)
    node.MassVol = data['Mass volume [J/K]']*1.0
    if data['Node type'] == 'DIFFUSION':
        node.NodeType = OpenTDv62.RcNodeData.NodeTypes.DIFFUSION
    elif data['Node type'] == 'ARITHMETIC':
        node.NodeType = OpenTDv62.RcNodeData.NodeTypes.ARITHMETIC
    elif data['Node type'] == 'BOUNDARY':
        node.NodeType = OpenTDv62.RcNodeData.NodeTypes.BOUNDARY
    elif data['Node type'] == 'CLONE':
        node.NodeType = OpenTDv62.RcNodeData.NodeTypes.CLONE
    else:
        raise ValueError('Unexpected node type')
    node.Comment = str(data['Comment'])
    node.Update()
    
def create_nodes(td: OpenTDv62.ThermalDesktop, file_name: str):
    df = pd.read_csv(file_name)
    for index, data in df.iterrows():
        create_node(td, data)
    
def delete_contactors(td: OpenTDv62.ThermalDesktop):
    contactors = td.GetContactors()
    for contactor in contactors:
        entity = OpenTDv62.TdDbEntityData(contactor.Handle)
        td.DeleteEntity(entity)
        
def delete_conductors(td: OpenTDv62.ThermalDesktop):
    contactors = td.GetConductors()
    for contactor in contactors:
        entity = OpenTDv62.TdDbEntityData(contactor.Handle)
        td.DeleteEntity(entity)
        
def delete_heaters(td: OpenTDv62.ThermalDesktop):
    heaters = td.GetHeaters()
    for heater in heaters:
        entity = OpenTDv62.TdDbEntityData(heater.Handle)
        td.DeleteEntity(entity)
        
def delete_heatloads(td: OpenTDv62.ThermalDesktop):
    heatloads = td.GetHeatLoads()
    for heatload in heatloads:
        entity = OpenTDv62.TdDbEntityData(heatload.Handle)
        td.DeleteEntity(entity)

def import_thermo_properties(td: OpenTDv62.ThermalDesktop, file_name: str):
    df = pd.read_csv(file_name)
    for index, data in df.iterrows():
        add_thermo_property(td, data)

def add_thermo_property(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    try:
        thermo = td.CreateThermoProps(data['Name'])
    except OpenTDv62.OpenTDException:
        thermo = td.GetThermoProps(data['Name'])
    thermo.Anisotropic = int(data['Anisotropic'])
    thermo.Comment = str(data['Comment'])
    thermo.Density = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Density](float(data['Density [kg/m3]']))  # [kg/m3]
    thermo.Conductivity = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.CondPerLength](float(data['Conductivity [W/mK]'])) # [W/mK]
    thermo.ConductivityY = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.CondPerLength](float(data['ConductivityY [W/mK]'])) # [W/mK]
    thermo.ConductivityZ = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.CondPerLength](float(data['ConductivityZ [W/mK]'])) # [W/mK]
    thermo.SpecificHeat = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.SpecificHeat](float(data['Specific heat [J/kgK]'])) # [J/kgK]
    thermo.Update()
    OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.ModelLength]
    
def import_optical_properties(td: OpenTDv62.ThermalDesktop, file_name: str):
    df = pd.read_csv(file_name)
    for index, data in df.iterrows():
        add_optical_property(td, data)
    
def add_optical_property(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    try:
        optical = td.CreateOpticalProps(data['Name'])
    except OpenTDv62.OpenTDException:
        optical = td.GetOpticalProps(data['Name'])
    optical.Comment = str(data['Comment'])
    optical.Alph = float(data['Absorptivity'])
    optical.Emis = float(data['Emissivity'])
    optical.Update()
        
def screenshot(td: OpenTDv62.ThermalDesktop, name: str, view: str='', style: str='XRAY', imageformat: str='png', width: int=0, height: int=0):
    td.SendCommand('CLEANSCREENON ')
    
    #IsoViews Enummeration
    iso_dict = {'SW':0, 'SE':1, 'NE':2, 'NW':3}
    #OrthoViews Enummeration
    ortho_dict = {'TOP':0, 'BOTTOM':1, 
                  'FRONT':2, 'BACK':3, 
                  'LEFT':4, 'RIGHT':5}
    if view in iso_dict:
        td.RestoreIsoView(iso_dict[view])
        td.ZoomExtents()
    elif view in ortho_dict:
        td.RestoreOrthoView(ortho_dict[view])
        td.ZoomExtents()
    
    # VisualStyles Enumeration
    style_dict = {'WIRE_2D':0, 'WIRE':1, 'CONCEPTUAL':2, 
                  'HIDDEN':3, 'REALISTIC':4, 'SHADED':5, 
                  'SHADED_W_EDGES':6, 'SHADES_OF_GREY':7, 'SKETCHY':8, 
                  'THERMAL':9, 'THERMAL_PP':10, 'XRAY':11}
    if style in style_dict:
        td.SetVisualStyle(style_dict[style])
    else:
        td.SetVisualStyle(OpenTDv62.VisualStyles.XRAY)

    # screenshot
    bmp = td.CaptureGraphicsArea()
    
    # trimming
    if  width <= 0 and height <= 0:
        # do nothing
        pass
    elif width <= 0 and height > 0:
        if bmp.Height > height:
            # trim height
            rect = RectangleF((0.0, (bmp.Height-height)/2, bmp.Width*1.0, height*1.0))
            bmp = bmp.Clone(rect, bmp.PixelFormat)
        else:
            # do nothing
            pass
    elif width > 0 and height <= 0:
        if bmp.Width > width:
            # trim width
            rect = RectangleF(((bmp.Width-width)/2, 0.0, width*1.0, bmp.Height*1.0))
            bmp = bmp.Clone(rect, bmp.PixelFormat)
        else:
            # do nothing
            pass
    else:
        # trim width and hight
        rect = RectangleF((bmp.Width-width)/2, (bmp.Height-height)/2, width*1.0, height*1.0)
        bmp = bmp.Clone(rect, bmp.PixelFormat)
    
    # ImageFormat Class
    # Bmp, Emf, Exif, Gif, Icon, Jpeg, Png, Tiff, Wmf
    if imageformat=='png':
        file_name = name + '.png'
        bmp.Save(file_name, ImageFormat.Png)
    elif imageformat=='jpeg' or imageformat=='jpg':
        file_name = name + '.jpeg'
        bmp.Save(file_name, ImageFormat.Jpeg)
    elif imageformat=='bmp':
        file_name = name + '.bmp'
        bmp.Save(file_name, ImageFormat.Bmp)
    elif imageformat=='emf':
        file_name = name + '.emf'
        bmp.Save(file_name, ImageFormat.Emf)
    elif imageformat=='wmf':
        file_name = name + '.wmf'
        bmp.Save(file_name, ImageFormat.Wmf)
    elif imageformat=='tiff':
        file_name = name + '.tiff'
        bmp.Save(file_name, ImageFormat.Tiff)
    elif imageformat=='gif':
        file_name = name + '.gif'
        bmp.Save(file_name, ImageFormat.Gif)
    elif imageformat=='exif':
        file_name = name + '.exif'
        bmp.Save(file_name, ImageFormat.Exif)
    elif imageformat=='icon':
        file_name = name + '.icon'
        bmp.Save(file_name, ImageFormat.Icon)
    else:
        raise ValueError('unexpected imageformat')
    td.SendCommand('CLEANSCREENOFF ')