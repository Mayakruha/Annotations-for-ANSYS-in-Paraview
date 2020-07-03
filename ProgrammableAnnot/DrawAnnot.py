#----- Code in the Programmable source: --------
#AnnotText=['Zone1','Temp: 100 C','HTC: 350 W/m2K','Press: 2.5 bar']
#Rotation=[90,0,0]
#Coordinates=[0.1,0.1,0.2]
#LineLength=0.5
#FontScale = 0.1
#import DrawAnnot
#DrawAnnot.DrawAnnot(output,AnnotText,Rotation,Coordinates,LineLength,FontScale)
#---------------------------------------
import vtk
def DrawAnnot(output,AnnotText,Rotation,Coordinates,LineLength,FontScale):
    Projection=0.0001 #Shift from the plane to a viewer
    RowNum=len(AnnotText)
    CellsNum=2
    PointsNum=0
    Width=0
    FontPoints=vtk.vtkPoints()
    TextSource=vtk.vtkTextSource()
    TextSource.BackingOff()
    #--common number of Cells
    for i in range(0,RowNum,1):
        if Width<len(AnnotText[i]):Width=len(AnnotText[i])
        TextSource.SetText(AnnotText[i])
        TextSource.Update()
        CellsNum+=TextSource.GetOutput().GetNumberOfCells()
        PointsNum+=TextSource.GetOutput().GetPoints().GetNumberOfPoints()
    #--vtkPolyData for Fonts
    outputFont=vtk.vtkPolyData()
    outputFont.Allocate(CellsNum)
    ColorArray=vtk.vtkFloatArray()
    ColorArray.SetName('Color')
    j=k=CellsNum
    PointsNum=0
    for i in range(0,RowNum,1):
        TextSource.SetText(AnnotText[i])
        TextSource.Update()
        for j in range(0,TextSource.GetOutput().GetPoints().GetNumberOfPoints()):
            PointCoord=TextSource.GetOutput().GetPoints().GetPoint(j)
            FontPoints.InsertNextPoint(FontScale*(PointCoord[0]-Width*4.5),FontScale*(PointCoord[1]+15*(RowNum-1-i))+LineLength,Projection) #--vtkSource: Hight - 15, Width -9
        for j in range(0,TextSource.GetOutput().GetNumberOfCells()):
            CellType=TextSource.GetOutput().GetCellType(j)
            CellPoints=vtk.vtkIdList()
            for k in range(0,TextSource.GetOutput().GetCell(j).GetPointIds().GetNumberOfIds()):
                Num=TextSource.GetOutput().GetCell(j).GetPointIds().GetId(k)
                CellPoints.InsertNextId(Num+PointsNum)
            outputFont.InsertNextCell(CellType,CellPoints)
            ColorArray.InsertNextValue(1.0)
        PointsNum+=TextSource.GetOutput().GetPoints().GetNumberOfPoints()
    #-------BackGround
    FontPoints.InsertNextPoint(-FontScale*Width*4.5,LineLength,0)
    FontPoints.InsertNextPoint(FontScale*Width*4.5,LineLength,0)
    FontPoints.InsertNextPoint(FontScale*Width*4.5,FontScale*RowNum*15+LineLength,0)
    FontPoints.InsertNextPoint(-FontScale*Width*4.5,FontScale*RowNum*15+LineLength,0)
    CellPoints=vtk.vtkIdList()
    CellPoints.InsertNextId(PointsNum)
    CellPoints.InsertNextId(PointsNum+1)
    CellPoints.InsertNextId(PointsNum+2)
    CellPoints.InsertNextId(PointsNum+3)
    outputFont.InsertNextCell(vtk.VTK_QUAD,CellPoints)
    ColorArray.InsertNextValue(0.0)
    #-------Line
    FontPoints.InsertNextPoint(-FontScale*0.5,0,Projection)
    FontPoints.InsertNextPoint(FontScale*0.5,0,Projection)
    FontPoints.InsertNextPoint(FontScale*0.5,LineLength,Projection)
    FontPoints.InsertNextPoint(-FontScale*0.5,LineLength,Projection)
    CellPoints=vtk.vtkIdList()
    CellPoints.InsertNextId(PointsNum+4)
    CellPoints.InsertNextId(PointsNum+5)
    CellPoints.InsertNextId(PointsNum+6)
    CellPoints.InsertNextId(PointsNum+7)
    outputFont.InsertNextCell(vtk.VTK_QUAD,CellPoints)
    ColorArray.InsertNextValue(1.0)
    outputFont.SetPoints(FontPoints)
    #-------Transformation
    Transformer=vtk.vtkTransform()
    Transformer.RotateX(Rotation[0])
    Transformer.RotateY(Rotation[1])
    Transformer.RotateZ(Rotation[2])
    Transformer.Translate(Coordinates)
    TransformFilter = vtk.vtkTransformPolyDataFilter() 
    TransformFilter.SetInputData(outputFont)
    TransformFilter.SetTransform(Transformer)
    TransformFilter.Update()
    #----------Output
    output.CopyStructure(TransformFilter.GetOutput())
    output.GetCellData().SetScalars(ColorArray)