#---------How to run
#python ANSYS-dat_to_vtm.py mesh.dat Table.xlsx
import sys
Dir=''						# Work directory
MeshFileName=sys.argv[1]	# ANSYS dat file
BCsFileName=sys.argv[2]		# Excel file with thermal boundary conditions
#---------Example---------------------
#Dir='D:\\Ansaldo\\GTI\\Annotations\\'
#MeshFileName='mesh.dat'
#BCsFileName='Table.xlsx'
#--Settings--------------------------------
print('*Start')
Hex=[[0,1,2,3],[0,1,5,4],[1,2,6,5],[2,3,7,6],[0,4,7,3],[4,5,6,7]]
Tet=[[0,1,2],[0,1,3],[1,2,3],[0,3,2]]
#path to Excel reading module
#sys.path.append('C:\\Users\\alexey.makrushin\\AppData\\Local\\Continuum\\anaconda2\\Lib\\site-packages')
#path to vtk module
sys.path.append('D:\\Programs\\ParaView-5.7.0-RC1-Windows-msvc2015-64bit\\bin\\Lib\\site-packages')
import vtk
#------------------------------------------
from openpyxl import load_workbook
wb=load_workbook(Dir+BCsFileName, read_only=True)
ws=wb['BC Zones']
TableRows={}
for i in range(4,ws.max_row+1): TableRows[str(ws.cell(i,15).value)]=i
#----------------------------------------------------------
#------------DAT-file reading------------------------------
print('*Mesh reading')
ElemNodesNum=[0,0,0,0,0,0,0,0]
PointsArray=vtk.vtkPoints()
PointsArray.InsertNextPoint(0.0,0.0,0.0)
Faces={}	#Keys:[Min][Max][Third]
CMBLOCK={}
BlocksList=[]
BlocksList.append('')
BlocksInfo={}
BlocksInfo['']=[0,0,vtk.vtkUnstructuredGrid(),''] 	# Index in the List, Number of Cells, vtkUnstructuredGrid
mf=open(Dir+MeshFileName,'r')
txt=mf.readline()
while txt[0:5]!='/solu':
	if txt[0:25]=='/com,*********** Create "':
		cmnt=txt.split('"')[1]
		NamedSelection=''
		for TableKey in TableRows:
			if (TableKey in cmnt)and(not TableKey in BlocksInfo):
				NamedSelection=TableKey
				BlocksInfo[NamedSelection]=[len(BlocksList),0,vtk.vtkUnstructuredGrid()]
				BlocksList.append(NamedSelection)
				print(NamedSelection+' has been found')
	elif txt[0:3]=='et,': ElemType=int(txt.split(',')[2])
	elif txt[0:9]=='nblock,3,':			#-----Nodes block
		txt=mf.readline()
		Values=txt.split(',')
		NumLen=int(Values[0].split('i')[1])
		CoordLen=int(Values[1].split('.')[0].split('e')[1])
		txt=mf.readline()
		while txt[0:2]!='-1':
			PointsArray.InsertNextPoint(float(txt[NumLen:NumLen+CoordLen]),float(txt[NumLen+CoordLen:NumLen+2*CoordLen]),float(txt[NumLen+2*CoordLen:NumLen+3*CoordLen]))
			txt=mf.readline()
	elif txt[0:7]=='eblock,':
		NumLen=int(mf.readline()[1:-2].split('i')[1])
		txt=mf.readline()
		if ElemType==87:		#-----QUADRATIC TETRA
			while txt[0:2]!='-1':
				for i in range(0,4):ElemNodesNum[i]=int(txt[(i+11)*NumLen:(i+12)*NumLen])
				for Face in Tet:
					FaceNodesNum=[ElemNodesNum[Face[0]],ElemNodesNum[Face[1]],ElemNodesNum[Face[2]]]
					MinNum=FaceNodesNum[0]
					MaxNum1=FaceNodesNum[0]
					if MinNum>FaceNodesNum[1]:
						MaxNum1=MinNum
						MinNum=FaceNodesNum[1]
					else:MaxNum1=FaceNodesNum[1]
					if MinNum>FaceNodesNum[2]:
						MidNum1=MinNum
						MinNum=FaceNodesNum[2]
					elif MaxNum1<FaceNodesNum[2]:
						MidNum1=MaxNum1
						MaxNum1=FaceNodesNum[2]
					else:MidNum1=FaceNodesNum[2]
					if not MinNum in Faces: Faces[MinNum]={}
					if not MaxNum1 in Faces[MinNum]: Faces[MinNum][MaxNum1]={}
					if not MidNum1 in Faces[MinNum][MaxNum1]: Faces[MinNum][MaxNum1][MidNum1]=0
					else:
						Faces[MinNum][MaxNum1].__delitem__(MidNum1)
						if len(Faces[MinNum][MaxNum1])==0:
							Faces[MinNum].__delitem__(MaxNum1)
							if len(Faces[MinNum])==0:Faces.__delitem__(MinNum)
				txt=mf.readline()
				txt=mf.readline()
		elif ElemType==90:		#-----QUADRATIC HEX
			while txt[0:2]!='-1':
				for i in range(0,8):ElemNodesNum[i]=int(txt[(i+11)*NumLen:(i+12)*NumLen])
				for Face in Hex:
					FaceNodesNum=[ElemNodesNum[Face[0]],ElemNodesNum[Face[1]],ElemNodesNum[Face[2]],ElemNodesNum[Face[3]]]
					MinNum=int(min(FaceNodesNum))
					indx=FaceNodesNum.index(MinNum)
					if indx==3:
						MaxNum1=FaceNodesNum[0]
						MidNum1=FaceNodesNum[0]
					else:
						MaxNum1=FaceNodesNum[indx+1]
						MidNum1=FaceNodesNum[indx+1]
					if indx==0:
						MaxNum2=FaceNodesNum[3]
						MidNum2=FaceNodesNum[3]
					else:
						MaxNum2=FaceNodesNum[indx-1]
						MidNum2=FaceNodesNum[indx-1]
					#----------
					if indx>1:
						if MaxNum1<FaceNodesNum[indx-2]:MaxNum1=FaceNodesNum[indx-2]
						else:MidNum1=FaceNodesNum[indx-2]
						if MaxNum2<FaceNodesNum[indx-2]:MaxNum2=FaceNodesNum[indx-2]
						else:MidNum2=FaceNodesNum[indx-2]					
					else:
						if MaxNum1<FaceNodesNum[indx+2]:MaxNum1=FaceNodesNum[indx+2]
						else:MidNum1=FaceNodesNum[indx+2]
						if MaxNum2<FaceNodesNum[indx+2]:MaxNum2=FaceNodesNum[indx+2]
						else:MidNum2=FaceNodesNum[indx+2]
					if not MinNum in Faces: Faces[MinNum]={}
					if not MaxNum1 in Faces[MinNum]: Faces[MinNum][MaxNum1]={}
					if not MaxNum2 in Faces[MinNum]: Faces[MinNum][MaxNum2]={}
					if not MidNum1 in Faces[MinNum][MaxNum1]: Faces[MinNum][MaxNum1][MidNum1]=0
					else:
						Faces[MinNum][MaxNum1].__delitem__(MidNum1)
						if len(Faces[MinNum][MaxNum1])==0:
							Faces[MinNum].__delitem__(MaxNum1)
							if len(Faces[MinNum])==0:Faces.__delitem__(MinNum)
					if not MidNum2 in Faces[MinNum][MaxNum2]: Faces[MinNum][MaxNum2][MidNum2]=0
					else:
						Faces[MinNum][MaxNum2].__delitem__(MidNum2)
						if len(Faces[MinNum][MaxNum2])==0:
							Faces[MinNum].__delitem__(MaxNum2)
							if len(Faces[MinNum])==0:Faces.__delitem__(MinNum)
				txt=mf.readline()
				txt=mf.readline()
		elif ElemType==152 and NamedSelection!='':		#-----SURFACE ELEMENTS
			while txt[0:2]!='-1':
				for i in range(0,4):ElemNodesNum[i]=int(txt[(i+5)*NumLen:(i+6)*NumLen])
				if ElemNodesNum[2]==ElemNodesNum[3]:
					FaceNodesNum=[ElemNodesNum[0],ElemNodesNum[1],ElemNodesNum[2]]
					MinNum=int(min(FaceNodesNum))
					FaceNodesNum.remove(MinNum)
					MaxNum1=int(max(FaceNodesNum))
					FaceNodesNum.remove(MaxNum1)
					if MinNum in Faces:
						if MaxNum1 in Faces[MinNum]:
							if FaceNodesNum[0] in Faces[MinNum][MaxNum1]:Faces[MinNum][MaxNum1][FaceNodesNum[0]]=BlocksInfo[NamedSelection][0]
				else:
					FaceNodesNum=[ElemNodesNum[0],ElemNodesNum[1],ElemNodesNum[2],ElemNodesNum[3]]
					MinNum=int(min(FaceNodesNum))
					indx=FaceNodesNum.index(MinNum)
					if indx==3:
						MaxNum1=FaceNodesNum[0]
						MidNum1=FaceNodesNum[0]
					else:
						MaxNum1=FaceNodesNum[indx+1]
						MidNum1=FaceNodesNum[indx+1]
					if indx==0:
						MaxNum2=FaceNodesNum[3]
						MidNum2=FaceNodesNum[3]
					else:
						MaxNum2=FaceNodesNum[indx-1]
						MidNum2=FaceNodesNum[indx-1]
					#----------
					if indx>1:
						if MaxNum1<FaceNodesNum[indx-2]:MaxNum1=FaceNodesNum[indx-2]
						else:MidNum1=FaceNodesNum[indx-2]
						if MaxNum2<FaceNodesNum[indx-2]:MaxNum2=FaceNodesNum[indx-2]
						else:MidNum2=FaceNodesNum[indx-2]					
					else:
						if MaxNum1<FaceNodesNum[indx+2]:MaxNum1=FaceNodesNum[indx+2]
						else:MidNum1=FaceNodesNum[indx+2]
						if MaxNum2<FaceNodesNum[indx+2]:MaxNum2=FaceNodesNum[indx+2]
						else:MidNum2=FaceNodesNum[indx+2]
					if MinNum in Faces:
						if MaxNum1 in Faces[MinNum]:
							if MidNum1 in Faces[MinNum][MaxNum1]: Faces[MinNum][MaxNum1][MidNum1]=BlocksInfo[NamedSelection][0]						
						if MaxNum2 in Faces[MinNum]:
							if MidNum2 in Faces[MinNum][MaxNum2]: Faces[MinNum][MaxNum2][MidNum2]=BlocksInfo[NamedSelection][0]
				txt=mf.readline()
	elif txt[0:8]=='CMBLOCK,':
		Values=txt.split(',')
		NamedSelection=Values[1].replace(' ','')
		NodesNum=int(Values[3])
		if NamedSelection in TableRows:
			CMBLOCK[NamedSelection]=[]
			BlocksInfo[NamedSelection]=[len(BlocksList),0,vtk.vtkUnstructuredGrid()]
			BlocksList.append(NamedSelection)
			print(NamedSelection+' has been found')
			NumLen=int(mf.readline()[1:-2].split('i')[1])
			j=0
			while j<NodesNum:
				txt=mf.readline()
				k0=len(txt)-1
				k=0					
				while k*NumLen<k0:
					CMBLOCK[NamedSelection].append(int(txt[k*NumLen:(k+1)*NumLen]))
					k+=1
				j+=k
	txt=mf.readline()			
mf.close()
#----------------------------------------------------------
#------------DATA treatment--------------------------------
print('**Data treatment')
#-- Node blocks and cells counter
for MinNum in Faces:
	for MaxNum1 in Faces[MinNum]:
		for MidNum1 in Faces[MinNum][MaxNum1]:
			for NamedSelection in CMBLOCK:
				if (MinNum in CMBLOCK[NamedSelection])and(MaxNum1 in CMBLOCK[NamedSelection])and(MidNum1 in CMBLOCK[NamedSelection]):
					Faces[MinNum][MaxNum1][MidNum1]=BlocksInfo[NamedSelection][0]
			BlocksInfo[BlocksList[Faces[MinNum][MaxNum1][MidNum1]]][1]+=1
#-- Blocks preaparation
for NamedSelection in BlocksList:
	BlocksInfo[NamedSelection][2].SetPoints(PointsArray)
	BlocksInfo[NamedSelection][2].Allocate(BlocksInfo[NamedSelection][1])
for MinNum in Faces:
	for MaxNum1 in Faces[MinNum]:
		for MidNum1 in Faces[MinNum][MaxNum1]:
			NamedSelection=BlocksList[Faces[MinNum][MaxNum1][MidNum1]]
			BlocksInfo[NamedSelection][2].InsertNextCell(vtk.VTK_TRIANGLE,3,[MinNum,MidNum1,MaxNum1]) #triangle Other: VTK_QUAD (2D),VTK_TETRA (3D), VTK_HEXAHEDRON (3D)
#----------------------------------------------------------
#------------OUTPUT----------------------------------------
output=vtk.vtkMultiBlockDataSet()
output.SetNumberOfBlocks(len(BlocksList))
for NamedSelection in BlocksList:
	if NamedSelection=='':AnnotationText=''
	elif float(ws.cell(TableRows[NamedSelection],12).value)!=-1e-30:
		AnnotationText='Temp:'+str(ws.cell(TableRows[NamedSelection],12).value)+' '+str(ws.cell(3,12).value)+'\n'+\
			'Press:'+str(ws.cell(TableRows[NamedSelection],9).value)+' '+str(ws.cell(3,9).value)
	elif float(ws.cell(TableRows[NamedSelection],13).value)!=-1e-30:
		AnnotationText='Q:'+str(ws.cell(TableRows[NamedSelection],13).value)+' '+str(ws.cell(3,13).value)+'\n'+\
			'Press:'+str(ws.cell(TableRows[NamedSelection],9).value)+' '+str(ws.cell(3,9).value)
	else:
		AnnotationText='Temp:'+str(ws.cell(TableRows[NamedSelection],7).value)+' '+str(ws.cell(3,7).value)+'\n'+\
			'HTC:'+str(ws.cell(TableRows[NamedSelection],8).value)+' '+str(ws.cell(3,8).value)+'\n'+\
			'Press:'+str(ws.cell(TableRows[NamedSelection],9).value)+' '+str(ws.cell(3,9).value)	
	output.GetMetaData(BlocksInfo[NamedSelection][0]).Set(vtk.vtkCompositeDataSet.NAME(), NamedSelection)
	output.GetMetaData(BlocksInfo[NamedSelection][0]).Set(output.FIELD_NAME(), AnnotationText)
	output.SetBlock(BlocksInfo[NamedSelection][0], BlocksInfo[NamedSelection][2])
wb.close()
print('***Data output')
MBDSwriter=vtk.vtkXMLMultiBlockDataWriter()
MBDSwriter.SetFileName(Dir+MeshFileName[0:-3]+'vtm')
MBDSwriter.SetInputData(output)
MBDSwriter.Write()
