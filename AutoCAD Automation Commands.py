# import pandas as pd                             # For Data framework
import math
# import numpy as np
import comtypes
import comtypes.client
import win32com.client                          # For Application connecting
import pythoncom
import pywintypes
pi = math.pi
acad = win32com.client.Dispatch("AutoCAD.Application", True)

if not acad.Visible:
    print('AutoCAD is not Running.!!!')
    print('Please open AutoCAD Drawing then try again.')
    quit(-1)
doc = acad.ActiveDocument
#print(type(doc.SelectionSets))
# Verify AutoCAD connection
try:
    print('File {} connected.'.format(doc.Name))
except AttributeError:
    print('AutoCAD is in use.!!!')
    print('Press Esc on AutoCAD window then try again.')
    quit(1)
#doc.Utility.Prompt("Execute from python\n")
ms = doc.ModelSpace
# Declare code dictionary
codedict = {0:['Text', 'Line', 'Circle'], 1:['TextString', None, None], 8:['Layer', 'Layer', 'Layer'],
            40:['Height', None, 'Radius'], 50:['Rotation', None, None], 62:['Color', 'Color', 'Color']}


# Class of AutoCAD Selection set
class AcSelectionSets:
    num_ss = 0
    def __init__(self, ss_name):
        self.ss_name = ss_name
        self._initslset()
        AcSelectionSets.num_ss += 1
    def _initslset(self):
        # Add the name "ss_name" selection set
        try:
            doc.SelectionSets.Item(self.ss_name).Delete()
        except:
            print("Delete selection failed")
        self.slset = doc.SelectionSets.Add(self.ss_name)
    def ssCond(self, ftyp, ftdt):
        # Set format of filter
        filterType = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, ftyp)
        filterData = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ftdt)
        self.slset.Select(5, 0, 0, filterType, filterData)            # Select all with filtering
        print('{} entities have been selected.'.format(self.slset.Count))
    def chProp(self, code, val):
        i = 0
        if code==8 and not layerexist(val):                         # If code=8 & layer not exist add it
            doc.Layers.Add(val)
        for e in self.slset:
            #print(dir(e))
            #etype = e.EntityName.replace('AcDb', '')
            etype = e.EntityName[4:]                                # EntityName = 'AcDbCircle'
            #print(etype)
            try:
                setattr(e, codedict[code][codedict[0].index(etype)], val)                 # Assign val to e.variable
                i += 1
            except:
                #print('{} can not be changed'.format(codedict[code][codedict[0].index(etype)]))
                print('{}: can not be changed by code = {}'.format(etype, code))
        pycad_prompt('Number of entities {}/{} have been changed.'.format(i, self.slset.count))


#Testing class AcSelectionSets
ass = AcSelectionSets('SS8')
# ass9 = AcSelectionSets('SS9')
ass.ssCond([0, 8], ['Circle', '0'])
# ass9.ssCond([0, 8], ['Text', lay1])
# print('{} entities selected by Class'.format(ass.slset.count))
#ass.chProp(50, 0.45)
#ass.chProp(62, 3)
#print(dir(ass.slset))
# ass.slset.Clear()                   # Clear entities in selection set
                                       # if not clear, previous selection shall be added
ass.ssCond([0, 8], ['Line', '*'])
"""
Example for doc.SelectionSets.AddItems()
nObj = doc.Utility.GetEntity()[0]
nObj2 = doc.Utility.GetEntity()[0]
nObj3 = doc.Utility.GetEntity()[0]
ass.slset.AddItems(vtobj([nObj, nObj2]))                        # Add item for 2 entities
ass.slset.AddItems(vtobj([nObj3]))                              # Add item for 1 entity
print(ass.slset.count)
"""
ass.slset.AddItems(vtobj(ass9.slset))
print('{} Texts'.format(ass9.slset.Count))
print(ass.slset.Count)



""" Passing and Obtaining Python Objects from COM>>> The functions below must be used before passing to ACTIVE X Command.
            ie Active X only understands the values below only"""
def vtpt(x, y, z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))
    """Converts to COM recognized type Float"""
def vtobj(obj):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)
    """Converts to COM recognized type Dispatch Object(Object of Application)"""
def vtFloat(lis):
    """ list converted to floating points"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lis)
def vtint(val):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, val)
    """Converts to COM recognized type Integer"""
def vtvariant(var):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, var)
    """Converts to COM recognized type Variant(that lies under application object(Like:Circle, Line...))"""

# Convert pt to vtpt
def pt_vtpt(pt):
    if len(pt)==2:
        return vtpt(pt[0], pt[1])
    else:
        return vtpt(pt[0], pt[1], pt[2])


# Polar function by giving point, angle, distance & Return 3D point(z=0)
def polar(p, a, d):
    x = p[0] + d * math.cos(a)
    y = p[1] + d * math.sin(a)
    return [x, y, 0.0]


# Distance function by giving 2D point1, point2
def distance(p, q):
    dx = p[0] - q[0]
    dy = p[1] - q[1]
    return math.sqrt(math.pow(dx, 2) + math.pow(dy, 2))


# Angle function by giving 2D point1, point2
def angle(p, q):
    dx = q[0] - p[0]
    dy = q[1] - p[1]
    return math.atan2(dy, dx)


# Compute boundary of giving Line entity & buffer
def line_bounds(e, b):
    al = e.Angle + pi * 0.5
    ar = e.Angle - pi * 0.5
    p1 = e.StartPoint
    p2 = e.EndPoint
    p11 = polar(p1, al, b)
    p12 = polar(p1, ar, b)
    p21 = polar(p2, al, b)
    p22 = polar(p2, ar, b)
    return [p11, p12, p22, p21, p11]


# Checking layer exist or not
def layerexist(lay):
    layers = doc.Layers
    print(type(layers))
    layers_nums = layers.Count
    layers_names = [layers.Item(i).Name for i in range(layers_nums)]  # List of ACAD layers
    if lay in layers_names:
        return True
    else:
        return False


# Create point
def make_point(pt, lay):
    pointObj = ms.AddPoint(pt_vtpt(pt))
    if not layerexist(lay):
        doc.Layers.Add(lay)
    pointObj.Layer = lay
    return pointObj


# Create line
def make_line(p1, p2, lay):
    lineObj = ms.AddLine(pt_vtpt(p1), pt_vtpt(p2))
    if not layerexist(lay):
        doc.Layers.Add(lay)
    lineObj.Layer = lay
    return lineObj


# Create polyline
def make_pline(pts, lay):
    plineObj = ms.AddPolyline(vtFloat(pts2pnts(pts)))
    if not layerexist(lay):
        doc.Layers.Add(lay)
    plineObj.Layer = lay
    return plineObj


# Create circle
def make_circle(pt, r, lay):
    circleObj = ms.AddCircle(pt_vtpt(pt), r)
    if not layerexist(lay):
        doc.Layers.Add(lay)
    circleObj.Layer = lay


# Create Text
def ptxt(txt, pt, ht, lay):
    textObj = ms.AddText(txt, pt_vtpt(pt), ht)
    if not layerexist(lay):
        doc.Layers.Add(lay)
    textObj.Layer = lay
    return textObj


# Prompt on Python & AutoCAD window
def pycad_prompt(msg):
    print(msg)
    doc.Utility.Prompt(msg + '\n')


# Get pick points from CAD window
def getpts(msg):
    pts = []
    pt = []
    pt0 = []
    print(dir(doc.Utility))
    i = 0
    while pt != None:
        try:
            doc.Utility.Prompt(msg + ' No.{} <enter to end> : '.format(i + 1))
            pt = doc.Utility.GetPoint()
            if i > 0:
                cmd = '(grdraw ' + '\'' + str(pt0).replace(",", "") + '\'' + str(pt).replace(",", "") + ' 1) '
                doc.SendCommand(cmd)
            pts.append(pt)
            pt0 = pt
            i += 1
        except:
            pt = None
    return pts


# Convert points array to list
def pts2pnts(pts):
    pnts = ()
    for p in pts:
        pnts = pnts + p
    return pnts


# Testing
p1 = [1000.5, 2000.5, 100.0]
p2 = [2100.0, 3100.0, 120.0]
p3 = [2700.0, 2100.0]
r2 = 40
r3 = 30
txt = 'Python start'
txt2 = 'Make text from Python'
lay1 = 'layer_test'
# pycad_prompt(txt)
make_line(p1, p2, lay1)
make_line(p1, p3, lay1)
ptxt(txt2, p1, 25, lay1)
make_line(p2, p3, lay1)
make_circle(p2, r2, '0')
make_circle(p3, r3, '0')
#
# pts = getpts('Pick polyline point')
# make_pline(pts, lay1)
#
#
# #Testing class AcSelectionSets
# ass = AcSelectionSets('SS8')
# ass.ssCond([0, 8], ['Circle', '0'])
#
# ass9 = AcSelectionSets('SS9')
# ass9.ssCond([0, 8], ['Text', lay1])
# # print('{} entities selected by Class'.format(ass.slset.count))
# #ass.chProp(50, 0.45)
# #ass.chProp(62, 3)
# #print(dir(ass.slset))
# # ass.slset.Clear()                   # Clear entities in selection set
#                                        # if not clear, previous selection shall be added
# ass.ssCond([0, 8], ['Line', '*'])
# """
# Example for doc.SelectionSets.AddItems()
# nObj = doc.Utility.GetEntity()[0]
# nObj2 = doc.Utility.GetEntity()[0]
# nObj3 = doc.Utility.GetEntity()[0]
# ass.slset.AddItems(vtobj([nObj, nObj2]))                        # Add item for 2 entities
# ass.slset.AddItems(vtobj([nObj3]))                              # Add item for 1 entity
# print(ass.slset.count)
# """
# ass.slset.AddItems(vtobj(ass9.slset))
# print('{} Texts'.format(ass9.slset.Count))
# print(ass.slset.Count)