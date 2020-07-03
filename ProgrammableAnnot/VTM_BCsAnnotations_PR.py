Dir='D:\\Annotations\\'
FileName='mesh.vtm'
HorizAxis=0 #0-X, 1-Y, 2-Z
VertAxis=2 #Second axis of the annotation plane: 0-X, 1-Y, 2-Z
ViewDirect=True #True - Horizontal Axis is from left to right, False - Horizontal Axis is from right to left
Projection=-0.001 #Shift from the plane
FontScale = 0.002
AnnDist=0.004*FontScale*15 #vertical distance between annotations. 15-Height of vtkTextSource
RowNum=3 #number of rows for the nearest annotations
TxtColors=[[1.0, 0.0, 0.0],[1.0,1.0,0.0],[0.0, 1.0, 0.0],[0.0,1.0,1.0],[0.0,0.0,1.0]]
#------------------------------------------
#---------key function for sorting---------
Rotation=[0,0,0]
if VertAxis==0:
	Rotation[0]=(HorizAxis-1)*90-ViewDirect*180
	Rotation[2]=-90.0
elif VertAxis==1: Rotation[1]=-HorizAxis*45-ViewDirect*180+180
else:
	Rotation[0]=90.0
	Rotation[1]=HorizAxis*90-ViewDirect*180+180
DataTable=[]
BlockNums=[]
def HorizAxisCoord(Num):
	return DataTable[Num]['Coord'][HorizAxis]
#------------------------------------------
BlocksList=[]
BlockName=''
from paraview.simple import *
#----------------------------------
# create a new 'CGNS Series Reader'
meshvtm = XMLMultiBlockDataReader(FileName=[Dir+FileName])
# get active view
renderView1 = GetActiveViewOrCreate('RenderView')
# show geometry in view
meshvtmDisplay = Show(meshvtm, renderView1)
ColorBy(meshvtmDisplay, ('FIELD', 'vtkBlockColors'))
vtkBlockColorsLUT = GetColorTransferFunction('vtkBlockColors')
#----------------------------------
# data
BlockArr = servermanager.Fetch(meshvtm)				#to get full access to object data
BlocksNum=BlockArr.GetNumberOfBlocks()-1
#color by default
LegendAnnot=[]
LegendColor=[]
j=len(vtkBlockColorsLUT.IndexedColors)
if (j/3-1)<BlocksNum:
	for i in range(0,BlocksNum+1):
		LegendAnnot+=[str(i),str(i)]
		LegendColor+=[0.75,0.75,0.75]
	vtkBlockColorsLUT.Annotations=LegendAnnot
	vtkBlockColorsLUT.IndexedColors=LegendColor
else:	
	for i in range(0,j): vtkBlockColorsLUT.IndexedColors[i]=0.75			
meshvtmDisplay.BlockColorsDistinctValues = BlocksNum
#colors for annotations
colorLUT=[]
for i in range(0,len(TxtColors)):
	colorLUT.append(GetColorTransferFunction('Color'+str(i)))
	colorLUT[i].RGBPoints = [0.0, 1.0, 1.0, 1.0, 1.0, TxtColors[i][0], TxtColors[i][1], TxtColors[i][2]]
#----------------------------------
for i in range(1,BlocksNum+1):
	SelBlock=BlockArr.GetBlock(i) 						#the last blocks - vtkUnstructuredGrid
#-----Name&Annotations---------
	Info=BlockArr.GetMetaData(i)
	BlockName=Info.Get(BlockArr.NAME())
	BlockValues=SelBlock.GetPointData()
	FirstValueName=BlockValues.GetArray(0).GetName()
	if FirstValueName=='Temp':
		AnnotText="['"+BlockName+"',"+\
			"'Temp: "+format(BlockValues.GetArray(0).GetTuple(0)[0],'0.1f')+"',"+\
			"'Press: "+format(BlockValues.GetArray(1).GetTuple(0)[0],'0.1f')+"']"
	elif FirstValueName=='Heat Flux':
		AnnotText="['"+BlockName+"',"+\
			"'Heat Flux: "+format(BlockValues.GetArray(0).GetTuple(0)[0],'0.1f')+"',"+\
			"'Press: "+format(BlockValues.GetArray(1).GetTuple(0)[0],'0.1f')+"']"
	else:
		AnnotText="['"+BlockName+"',"+\
			"'Temp: "+format(BlockValues.GetArray(0).GetTuple(0)[0],'0.1f')+"',"+\
			"'HTC: "+format(BlockValues.GetArray(1).GetTuple(0)[0],'0.1f') +"',"+\
			"'Press: "+format(BlockValues.GetArray(2).GetTuple(0)[0],'0.1f')+"']"
#-----Cells--------
	BlockCells=SelBlock.GetCells()
	CellsNum=BlockCells.GetNumberOfCells()
	CellDat=BlockCells.GetData()
# Search of Horizontal Coordinate for an annotation
	Cellj=0
	j=0
	while Cellj<CellsNum:
		CellsNodes=CellDat.GetValue(j)
		for k in range(0,CellsNodes):
			NodeNum=CellDat.GetValue(j+k+1)
			Coord=SelBlock.GetPoints().GetPoint(NodeNum)
			if (j==0 and k==0)or MinHorCoord>Coord[HorizAxis]: MinHorCoord=Coord[HorizAxis]
			if (j==0 and k==0)or MaxHorCoord<Coord[HorizAxis]: MaxHorCoord=Coord[HorizAxis]
		Cellj+=1
		j+=CellsNodes+1
	MdlHorCoord=(MinHorCoord+MaxHorCoord)/2
# Search of rest coordinates for an annotation
	Cellj=0
	j=0
	Flag=True
	while Cellj<CellsNum:
		CellsNodes=CellDat.GetValue(j)
		CellCoord=[0.0,0.0,0.0]
		for k in range(0,CellsNodes):
			NodeNum=CellDat.GetValue(j+k+1)
			Coord=SelBlock.GetPoints().GetPoint(NodeNum)
			CellCoord[0]+=Coord[0]/CellsNodes
			CellCoord[1]+=Coord[1]/CellsNodes
			CellCoord[2]+=Coord[2]/CellsNodes			
			if (k==0)or MinHorCoord>Coord[HorizAxis]: MinHorCoord=Coord[HorizAxis]
			if (k==0)or MaxHorCoord<Coord[HorizAxis]: MaxHorCoord=Coord[HorizAxis]
		if (MinHorCoord-MdlHorCoord)*(MaxHorCoord-MdlHorCoord)<=0:
			if Flag:
				AnnCoord=[CellCoord[0],CellCoord[1],CellCoord[2]]
				Flag=False
			elif AnnCoord[VertAxis]<CellCoord[VertAxis]:
				AnnCoord=[CellCoord[0],CellCoord[1],CellCoord[2]]
		Cellj+=1
		j+=CellsNodes+1
#Table
	if Projection!=0:
		CellCoord=[Projection,Projection,Projection]
		CellCoord[HorizAxis]=AnnCoord[HorizAxis]
		CellCoord[VertAxis]=AnnCoord[VertAxis]
		DataTable.append({'Coord':[CellCoord[0],CellCoord[1],CellCoord[2]],'Name':BlockName,'Num':i-1,'ArrowLength':AnnDist/2,'Text':AnnotText})
	else:
		DataTable.append({'Coord':[AnnCoord[0],AnnCoord[1],AnnCoord[2]],'Name':BlockName,'Num':i-1,'ArrowLength':AnnDist/2,'Text':AnnotText})
	BlockNums.append(i-1)
BlockNums.sort(key=HorizAxisCoord)
#----------------------------------
i=0
# show annotations
for j in range(0,BlocksNum):
	Num=BlockNums[j]
	j0=j-RowNum
	if j0<0:j0=0
	for k in BlockNums[j0:j]:
		if abs((DataTable[Num]['Coord'][VertAxis]+DataTable[Num]['ArrowLength'])-(DataTable[k]['Coord'][VertAxis]+DataTable[k]['ArrowLength']))<AnnDist:
			DataTable[Num]['ArrowLength']=DataTable[k]['Coord'][VertAxis]+DataTable[k]['ArrowLength']-DataTable[Num]['Coord'][VertAxis]+AnnDist
# create a new 'Text'
	PrSource = ProgrammableSource()
#----------------------------------
	PrSource.Script = "AnnotText="+DataTable[Num]['Text']+"\n"+\
					"Rotation="+str(Rotation)+"\n"+\
					"Coordinates="+str(DataTable[Num]['Coord'])+"\n"+\
					"LineLength="+str(DataTable[Num]['ArrowLength'])+"\n"+\
					"FontScale="+str(FontScale)+"\n"+\
					"import DrawAnnot\n"+\
					"DrawAnnot.DrawAnnot(output,AnnotText,Rotation,Coordinates,LineLength,FontScale)"
#----------------------------------
# rename source object
	RenameSource(DataTable[Num]['Name'], PrSource)
	PrSourceDisplay = Show(PrSource, renderView1)
	PrSourceDisplay.ColorArrayName = ['CELLS', 'Color']
	PrSourceDisplay.LookupTable = colorLUT[i]    
# show data in view
	vtkBlockColorsLUT.IndexedColors[3*(DataTable[Num]['Num']+1)]=TxtColors[i][0]/1.5
	vtkBlockColorsLUT.IndexedColors[3*(DataTable[Num]['Num']+1)+1]=TxtColors[i][1]/1.5
	vtkBlockColorsLUT.IndexedColors[3*(DataTable[Num]['Num']+1)+2]=TxtColors[i][2]/1.5
	i+=1
	if i==len(TxtColors):i=0
#----------------------------------
renderView1.Update()
renderView1.CameraParallelProjection = 1
