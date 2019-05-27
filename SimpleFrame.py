import os
import win32com.client

#create Sap2000 object
SapObject = win32com.client.Dispatch("Sap2000v15.SapObject")

#start Sap2000 application
SapObject.ApplicationStart()

#create SapModel object
SapModel = SapObject.SapModel

#initialize model
SapModel.InitializeNewModel()

#create new blank model
ret = SapModel.File.NewBlank()

#define material property
MATERIAL_BALSA = 2
ret = SapModel.PropMaterial.SetMaterial('BALSA', MATERIAL_BALSA)

#assign isotropic mechanical properties to material (arbituary imput for now)
#(Name, modulus of elasticity, poisson's ratio, thermal coefficient)
ret = SapModel.PropMaterial.SetMPIsotropic('BALSA', 3600, 0.2, 0.0000055)

#define rectangular frame section property
#(name, material property, section depth, section width)
ret = SapModel.PropFrame.SetRectangle('R1', 'BALSA', 10, 10)

#define frame section property modifiers
#[cross section area,shear area in local 2 direction, '' in local 3 directionm torsional constant, 
# moment of inertia about local 2 axis, '' local 3 axis, mass, weight]
ModValue = [1000, 0, 0, 1, 1, 1, 1, 1]
ret = SapModel.PropFrame.SetModifiers('R1', ModValue)

#switch to k-ft units
kip_ft_F = 4
ret = SapModel.SetPresentUnits(kip_ft_F)

#add frame object by coordinates
FrameName1 = ' '
FrameName2 = ' '
FrameName3 = ' '
FrameName4 = ' '
FrameName5 = ' '
FrameName6 = ' '
FrameName7 = ' '
[ret, FrameName1] = SapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 1, FrameName1, 'R1', '1', 'Global')
[ret, FrameName2] = SapModel.FrameObj.AddByCoord(0, 0, 1, 1, 0, 1, FrameName2, 'R1', '2', 'Global')
[ret, FrameName3] = SapModel.FrameObj.AddByCoord(1, 0, 1, 1, 0, 0, FrameName3, 'R1', '3', 'Global')
[ret, FrameName4] = SapModel.FrameObj.AddByCoord(0, 0, 0, 0.5, 0, 0.5, FrameName4, 'R1', '4', 'Global')
[ret, FrameName5] = SapModel.FrameObj.AddByCoord(0.5, 0, 0.5, 1, 0, 1, FrameName5, 'R1', '5', 'Global')
[ret, FrameName6] = SapModel.FrameObj.AddByCoord(1, 0, 0, 0.5, 0, 0.5, FrameName6, 'R1', '6', 'Global')
[ret, FrameName7] = SapModel.FrameObj.AddByCoord(0.5, 0, 0.5, 0, 0, 1, FrameName7, 'R1', '7', 'Global')

#assign point object restraint at left base node
PointName1 = ' '
PointName2 = ' '
Restraint = [True, True, True, True, True, True]
[ret, PointName1, PointName2] = SapModel.FrameObj.GetPoints(FrameName1, PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName1, Restraint)

#assign point object restraint at right base node
Restraint = [True, True, False, False, False, False]
[ret, PointName1, PointName2] = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName2, Restraint)
#refresh view, update (initialize) zoom
ret = SapModel.View.RefreshView(0, False)

#add load patterns
LTYPE_OTHER = 8
ret = SapModel.LoadPatterns.Add('1', LTYPE_OTHER)

#assign loading for load pattern 
[ret, PointName1, PointName2] = SapModel.FrameObj.GetPoints(FrameName4, PointName1, PointName2)
PointLoadValue = [0,0,-20,0,0,0]
ret = SapModel.PointObj.SetLoadForce(PointName2, '1', PointLoadValue)

#Define time history function
time_history = r'C:\Users\shirl\OneDrive - University of Toronto\Desktop\Seismic\GM1.txt'
N_m_C = 10
SapModel.SetPresentUnits(N_m_C)
SapModel.Func.FuncTH.SetFromFile('GM', time_history, 1, 0, 1, 2, True)
#Set the time history load case
N_m_C = 10
SapModel.SetPresentUnits(N_m_C)
SapModel.LoadCases.ModHistLinear.SetCase('GM')
SapModel.LoadCases.ModHistLinear.SetMotionType('GM', 1)
SapModel.LoadCases.ModHistLinear.SetLoads('GM', 1, ['Accel'], ['U1'], ['GM'], [1], [1], [0], ['Global'], [0])
SapModel.LoadCases.ModHistLinear.SetTimeStep('GM', 250, 0.1)

#switch to k-in units
kip_in_F = 3;
ret = SapModel.SetPresentUnits(kip_in_F)

#save model
APIPath = 'C:\API'
if not os.path.exists(APIPath):
 try:
     os.makedirs(APIPath)
 except OSError:
     pass
ret = SapModel.File.Save(APIPath + os.sep + 'API_1-002.sdb')

#run model (this will create the analysis model)
ret = SapModel.Analyze.RunAnalysis()

#close Sap2000
ret = SapObject.ApplicationExit(False)
SapModel = 0;
SapObject = 0;
