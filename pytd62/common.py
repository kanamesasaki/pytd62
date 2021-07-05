import math
import numpy as np
import pandas as pd
import sys, clr
fileobj = open("path.txt", "r", encoding="utf_8")
sys.path.append(fileobj.readline().strip())
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
        error_message = submodel_name + '.' + str(node_id) + ' does not exist.'
        raise ValueError(error_message)
    elif len(node_extracted) > 1:
        error_message = submodel_name + '.' + str(node_id) + ' is assigned to multiple nodes.'
        raise ValueError(error_message)
        
def get_node(td: OpenTDv62.ThermalDesktop, submodel_name: str, node_id: int, nodes: list=[]):
    if nodes == []:
        nodes = td.GetNodes()
    node_extracted = [inode for inode in nodes if submodel_name == inode.Submodel.Name and node_id == inode.Id]
    if len(node_extracted) == 1:
        return node_extracted[0]
    elif len(node_extracted) == 0:
        error_message = submodel_name + '.' + str(node_id) + ' does not exist.'
        raise ValueError(error_message)
    else:
        error_message = submodel_name + '.' + str(node_id) + ' is assigned to multiple nodes.'
        raise ValueError(error_message)
    
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

def create_conductor(td: OpenTDv62.ThermalDesktop, data: pd.Series, nodes: list=[]):
    if nodes == []:
        nodes = td.GetNodes()
    fr_handle = OpenTDv62.Connection(get_node_handle(td, data['From submodel'], data['From ID'], nodes))
    to_handle = OpenTDv62.Connection(get_node_handle(td, data['To submodel'], data['To ID'], nodes))
    cond = td.CreateConductor(fr_handle, to_handle)
    cond.Submodel = OpenTDv62.SubmodelNameData(data['Submodel'])
    cond.Value = data['Conductance [W/K]']*1.0
    cond.Comment = str(data['Comment'])
    cond.Update()

def create_conductors(td: OpenTDv62.ThermalDesktop, file_name):
    df = pd.read_csv(file_name)
    nodes = td.GetNodes()
    for index, data in df.iterrows():
        create_conductor(td, data, nodes)
    
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

def rotate_all(td: OpenTDv62.ThermalDesktop, rotate: np.ndarray):
    elements = get_elements(td)
    for key in elements:
        for elem in elements[key]:
            if key == 'Polygon':
                rotate_polygon(elem)
            else:
                rotate_element(elem)

def get_elements(td: OpenTDv62.ThermalDesktop):
    # create empty dictionary object
    elements = {}
    elements['Disk'] = td.GetDisks()
    elements['Cylinder'] = td.GetCylinders()
    elements['Rectangle'] = td.GetRectangles()
    elements['Cone'] = td.GetCones()
    elements['Sphere'] = td.GetSpheres()
    elements['Torus'] = td.GetToruses()
    elements['ScarfedCylinder'] = td.GetScarfedCylinders()
    elements['Polygon'] = td.GetPolygons()
    elements['SolidCylinder'] = td.GetSolidCylinders()
    elements['SolidBrick'] = td.GetSolidBricks()
    elements['SolidSphere'] = td.GetSolidSpheres()
    return elements

def rotate_element(element, rotate: np.ndarray):
    base = np.array([[element.BaseTrans.entry[0][0], element.BaseTrans.entry[0][1], element.BaseTrans.entry[0][2]],
                     [element.BaseTrans.entry[1][0], element.BaseTrans.entry[1][1], element.BaseTrans.entry[1][2]],
                     [element.BaseTrans.entry[2][0], element.BaseTrans.entry[2][1], element.BaseTrans.entry[2][2]]])
    origin = np.array([element.BaseTrans.entry[0][3], element.BaseTrans.entry[1][3], element.BaseTrans.entry[2][3]])
    new_origin = rotate @ origin
    new_base = rotate @ base
    element.BaseTrans.entry[0][0] = new_base[0,0]
    element.BaseTrans.entry[1][0] = new_base[1,0]
    element.BaseTrans.entry[2][0] = new_base[2,0]
    element.BaseTrans.entry[0][1] = new_base[0,1]
    element.BaseTrans.entry[1][1] = new_base[1,1]
    element.BaseTrans.entry[2][1] = new_base[2,1]
    element.BaseTrans.entry[0][2] = new_base[0,2]
    element.BaseTrans.entry[1][2] = new_base[1,2]
    element.BaseTrans.entry[2][2] = new_base[2,2]
    element.BaseTrans.entry[0][3] = new_origin[0]
    element.BaseTrans.entry[1][3] = new_origin[1]
    element.BaseTrans.entry[2][3] = new_origin[2]
    element.Update()

def rotate_polygon(td: OpenTDv62.ThermalDesktop, polygon: OpenTDv62.RadCAD.Polygon, rotate: np.ndarray):
    """
    The polygon property "Vertices" is not settable. Therefore, it is not possible to move the existing polygon.
    With this function, a new polygon is created at a new position, and the original polygon is deleted.  
    """
    new_edges = List[OpenTDv62.Point3d]()
    for i in range(len(polygon.Vertices)//2):
        poly = polygon.Vertices[i]
        point = np.array([poly.X.GetValueSI(),poly.Y.GetValueSI(),poly.Z.GetValueSI()])
        new_point = rotate @ point
        new_edges.Add(OpenTDv62.Point3d(new_point[0], new_point[1], new_point[2]))
    new_polygon = td.CreatePolygon(new_edges)
    new_polygon.TopStartSubmodel = polygon.TopStartSubmodel
    new_polygon.TopStartId = polygon.TopStartId
    new_polygon.BreakdownU.Num = polygon.BreakdownU.Num
    new_polygon.BreakdownV.Num = polygon.BreakdownV.Num
    new_polygon.TopOpticalProp = polygon.TopOpticalProp
    new_polygon.BotOpticalProp = polygon.BotOpticalProp
    new_polygon.TopMaterial = polygon.TopMaterial
    new_polygon.TopThickness = polygon.TopThickness
    new_polygon.AnalysisGroups = polygon.AnalysisGroups
    new_polygon.CondSubmodel = polygon.CondSubmodel
    new_polygon.ColorIndex = polygon.ColorIndex
    new_polygon.Comment = polygon.Comment
    new_polygon.Update()
    td.DeleteEntity(OpenTDv62.TdDbEntityData(polygon.Handle))

def issurface(element):
    if isinstance(element, OpenTDv62.RadCAD.Disk):
        flag = True
    elif isinstance(element, OpenTDv62.RadCAD.Cylinder):
        flag = True
    elif isinstance(element, OpenTDv62.RadCAD.Rectangle):
        flag = True
    elif isinstance(element, OpenTDv62.RadCAD.Cone):
        flag = True
    elif isinstance(element, OpenTDv62.RadCAD.Sphere):
        flag = True
    elif isinstance(element, OpenTDv62.RadCAD.Torus):
        flag = True
    elif isinstance(element, OpenTDv62.RadCAD.ScarfedCylinder):
        flag = True
    elif isinstance(element, OpenTDv62.RadCAD.Polygon):
        flag = True
    else:
        flag = False
    return flag

def issolid(element):
    if isinstance(element, OpenTDv62.RadCAD.FdSolid.SolidCylinder):
        flag = True
    elif isinstance(element, OpenTDv62.RadCAD.FdSolid.SolidBrick):
        flag = True
    elif isinstance(element, OpenTDv62.RadCAD.FdSolid.SolidSphere):
        flag = True
    else:
        flag = False
    return flag

def get_element(td: OpenTDv62.ThermalDesktop, submodel_name: str, element_type: str, comment: str, elements: dict={}):
    if elements == {}:
        elements = get_elements(td)
    if element_type in ['Disk', 'Cylinder', 'Rectangle', 'Cone', 'Sphere', 'Torus', 'ScarfedCylinder', 'Polygon']:
        element = [ielement for ielement in elements[element_type] if submodel_name == ielement.TopStartSubmodel.Name and comment == ielement.Comment]
    elif element_type in ['SolidCylinder', 'SolidBrick', 'SolidSphere']:
        element = [ielement for ielement in elements[element_type] if submodel_name == ielement.StartSubmodel.Name and comment == ielement.Comment]
    else:
        raise ValueError('Unexpected element type')
    num = len(element)
    if num == 1:
        return element[0]
    elif num == 0:
        raise ValueError('the requested element does not exist')
    else:
        raise ValueError('multiple elements are found with the same designation')

def create_contactors(td: OpenTDv62.ThermalDesktop, file_name: str):
    df = pd.read_csv(file_name)
    for index, data in df.iterrows():
        create_contactor(td, data)

def create_contactor(td: OpenTDv62.ThermalDesktop, data: pd.Series):
    elements = get_elements(td)
    contfrom = List[OpenTDv62.Connection]()
    contfrom.Add(OpenTDv62.Connection(get_element(td, data['FromSubmodel'], data['FromElement'], str(data['FromComment']), elements), int(data['Marker'])))
    i = int(1)
    while True:
        try:
            submodel = str(data[f'FromSubmodel.{i}'])
            element = str(data[f'FromElement.{i}'])
            comment = str(data[f'FromComment.{i}'])
            marker = int(data[f'Marker.{i}'])
            if submodel=='nan' or element=='nan' or comment=='nan': raise ValueError
        except:
            break
        else:
            contfrom.Add(OpenTDv62.Connection(get_element(td, submodel, element, comment, elements), marker))
            i = i + 1
    contto = List[OpenTDv62.Connection]()
    contto.Add(OpenTDv62.Connection(get_element(td, data['ToSubmodel'], data['ToElement'], str(data['ToComment']), elements)))
    j = int(1)
    while True:
        try:
            submodel = str(data[f'ToSubmodel.{j}'])
            element = str(data[f'ToElement.{j}'])
            comment = str(data[f'ToComment.{j}'])
            if submodel=='nan' or element=='nan' or comment=='nan': raise ValueError
        except:
            break
        else:
            contto.Add(OpenTDv62.Connection(get_element(td, submodel, element, comment, elements)))
            j = j + 1
    cont = td.CreateContactor(contfrom, contto)
    cont.ContactCond = data['Conductance [W/K]'] # [W/K]
    if data['UseFace'] == 'Face':
        cont.UseFace = 1 # 1:Face 
    elif data['UseFace'] == 'Edges':
        cont.UseFace = 0 # 0:Edges
    cont.CondSubmodel = OpenTDv62.SubmodelNameData(data['FromSubmodel'])
    if data['InputValueType'] == 'PER_AREA_OR_LENGTH':
        cont.InputValueType = OpenTDv62.RcConnData.ContactorInputValueTypes.PER_AREA_OR_LENGTH
    elif data['InputValueType'] == 'ABSOLUTE_COND_REDUCED_BY_UNCONNECTED':
        cont.InputValueType = OpenTDv62.RcConnData.ContactorInputValueTypes.ABSOLUTE_COND_REDUCED_BY_UNCONNECTED
    elif data['InputValueType'] == 'ABSOLUTE_ADJUST_FOR_UNCONNECTED':
        cont.InputValueType = OpenTDv62.RcConnData.ContactorInputValueTypes.ABSOLUTE_ADJUST_FOR_UNCONNECTED
    cont.Name = str(data['Name'])
    cont.Update()

def set_variable_nodetemp(filename: str, td: OpenTDv62.ThermalDesktop, submodel: str, nodeid: int, columnname1: str='Time [s]', columnname2: str='Data [K]'):
    bc = pd.read_csv(filename)
    nodes = td.GetNodes()
    node_list = [inode for inode in nodes if inode.Submodel.Name==submodel and inode.Id==nodeid]
    if len(node_list) == 1:
        node = node_list[0]
    elif len(node_list) == 0:
        raise ValueError('The requested node does not exist')
    else:
        raise ValueError('The given node_id is assigned to multiple nodes')

    addlist_time = List[float]()
    addlist_temp = List[float]()
    for itime in bc[columnname1]:
        addlist_time.Add(itime)
    for itemp in bc[columnname2]:
        addlist_temp.Add(itemp)
        
    timelist = OpenTDv62.Dimension.DimensionalList[OpenTDv62.Dimension.Time]()
    timelist.AddRange(addlist_time)
    node.TimeArray = timelist
    templist = OpenTDv62.Dimension.DimensionalList[OpenTDv62.Dimension.Temp]()
    templist.AddRange(addlist_temp)
    node.ValueArray = templist
    node.SteadyStateBoundaryType = OpenTDv62.RcNodeData.SteadyStateBoundaryTypes.INITIAL_TEMP
    node.InitialTemp = OpenTDv62.Dimension.Dimensional[OpenTDv62.Dimension.Temp](bc[columnname2][0])
    node.UseVersusTime = 1
    node.Update()

def set_variable_heatload(filename: str, td: OpenTDv62.ThermalDesktop, submodel: str, comment: str, columnname1: str='Time [s]', columnname2: str='Data [W]'):
    bc = pd.read_csv(filename)
    heats = td.GetHeatLoads()
    heat_list = [iheat for iheat in heats if iheat.Submodel.Name==submodel and iheat.Name==comment]
    if len(heat_list) == 1:
        heat = heat_list[0]
    elif len(heat_list) == 0:
        raise ValueError('The requested heatload does not exist')
    else:
        raise ValueError('The given name is assigned to multiple heatloads')
        
    heat.TempVaryType = OpenTDv62.RcHeatLoadData.HeatLoadTypes.LOAD
    heat.AppliedType = OpenTDv62.RcHeatLoadData.AppliedTypeBoundaryConds.NODE
    heat.HeatLoadTransientType = OpenTDv62.RcHeatLoadData.HeatLoadTransientTypes.TIME_VARY_HEAT_LOAD
    heat.TimeDependentSteadyStateType = OpenTDv62.RcHeatLoadData.TimeDependentSteadyStateTypes.TIME_INTERP
    addlist_time = List[float]()
    addlist_heat = OpenTDv62.ListSI()
    for itime in bc[columnname1]:
        addlist_time.Add(itime)
    for iheat in bc[columnname2]:
        addlist_heat.Add(iheat)
    timelist = OpenTDv62.Dimension.DimensionalList[OpenTDv62.Dimension.Time]()
    timelist.AddRange(addlist_time)
    heat.TimeArray = timelist
    heat.ValueArraySI = addlist_heat
    heat.Update()

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