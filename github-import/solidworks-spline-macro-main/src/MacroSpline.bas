Attribute VB_Name = "MacroSpline"
'**********************
'Copyright(C) 2025 Xarial Pty Limited
'Reference: https://www.codestack.net/solidworks-api/document/sketch/csv-create-spline/
'License: https://www.codestack.net/license/
'
' Modified by Víctor Liotti (2025) to include tangent angle (theta) control
' - Added AddTangencyControl to define tangent radial direction at each point
' - User must specify the file path and name in FILE_PATH
' - Any file extension (txt, csv, etc.) can be used as long as values are comma-separated and use a period as the decimal separator
'**********************

Const FILE_PATH As String = "C:\Users\victo\Python\Plug Nozzle\Files\data_13c.txt"

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSkMgr As SldWorks.SketchManager
Dim vPts As Variant
Dim dSplinePts() As Double
Dim theta() As Double
Dim i As Integer
Dim nPoints As Integer
Dim vPt As Variant
Dim x As Double
Dim y As Double
Dim z As Double
''Dim swSkSegment As SldWorks.SketchSegment
Dim swSkSpline As SldWorks.SketchSpline
Dim swSplineHandle As SldWorks.SplineHandle
Dim Newx As Double
Dim Newy As Double
Dim Newz As Double

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swSkMgr = swModel.SketchManager
    
    If Not swSkMgr.ActiveSketch Is Nothing Then
        
        vPts = ReadFile(FILE_PATH, False)
        ' Reads a CSV or TXT file containing comma-separated numerical values.
        ' If 'firstRowHeader' is True, the first row is skipped (use this if the first row contains column headers).
        ' If 'firstRowHeader' is False, all rows are treated as data.
        
        DrawSpline swSkMgr, vPts 'function
        
    Else
        Err.Raise vbError, "", "Please activate sketch"
    End If
    
End Sub

Sub DrawSpline(skMgr As SldWorks.SketchManager, vPoints As Variant)

    ' Desabilita atualizações de banco de dados para melhorar o desempenho
    skMgr.AddToDB = True
        
    ' Existem (UBound(vPoints) + 1) pontos
    ' Cada ponto tem 3 coordenadas (x,y,z)
    ' O primeiro elemento de um array é indicado zero, então subtrai-se 1
    ReDim dSplinePts((UBound(vPoints) + 1) * 3 - 1)
    ReDim theta(UBound(vPoints))


    ' For loop em cada ponto
    For i = 0 To UBound(vPoints)
    
        vPt = vPoints(i)
        x = vPt(0)
        y = vPt(1)
        z = vPt(2)
     
        ' Point i coordinates
        ' dSplinePts = (x1, y1, z1, x2, y2, z2, ..., xn, yn, zn)
        dSplinePts(i * 3) = x
        dSplinePts(i * 3 + 1) = y
        dSplinePts(i * 3 + 2) = z
        
        If UBound(vPt) >= 3 Then
            theta(i) = vPt(3)
        Else
            theta(i) = 0
        End If

    Next i ' Fim do loop
    
    
    ' Cria a spline sem tangentes
    Set swSkSpline = skMgr.CreateSpline(dSplinePts)
    If swSkSpline Is Nothing Then
        Err.Raise vbError, "", "Failed to create spline"
    End If
    
        
    For i = 0 To UBound(vPoints): ' For loop em cada ponto
        
        x = dSplinePts(i * 3)
        y = dSplinePts(i * 3 + 1)
        z = dSplinePts(i * 3 + 2)
        
        ' Aplica um controle de tangência no ponto i
        Set swSplineHandle = swSkSpline.AddTangencyControl(x, y, z)
        
        If swSplineHandle Is Nothing Then
            Debug.Print "Controle de tangência NÃO adicionado para o ponto " & i
        Else
            swSplineHandle.TangentRadialDirection = theta(i)
            'swSplineHandle.TangentDriving = True 'importante??
        End If
       
    Next i ' fim do loop
    
    ' Reativa as atualizações do esboço
    skMgr.AddToDB = False
    
End Sub

Function ReadFile(filePath As String, firstRowHeader As Boolean) As Variant
    
    'rows x columns
    Dim vTable() As Variant
    
    Dim fileName As String
    Dim tableRow As String
    Dim fileNo As Integer

    fileNo = FreeFile
    
    Open filePath For Input As #fileNo
    
    Dim isFirstRow As Boolean
        
    isFirstRow = True
    isTableInit = False
    
    Do While Not EOF(fileNo)
        
        Line Input #fileNo, tableRow
            
        If Not isFirstRow Or Not firstRowHeader Then
            
            Dim vCells As Variant
            vCells = Split(tableRow, ",")
            
            Dim i As Integer
            
            Dim dCells() As Double
            ReDim dCells(UBound(vCells))
            
            For i = 0 To UBound(vCells)
                vCells(i) = Replace(vCells(i), ".", ",") ' Troca ponto por vírgula
                dCells(i) = CDbl(vCells(i))
            Next
                    
            If (Not vTable) = -1 Then
                ReDim vTable(0)
            Else
                ReDim Preserve vTable(UBound(vTable) + 1)
            End If
                    
            vTable(UBound(vTable)) = dCells
            
        End If
        
        If isFirstRow Then
            isFirstRow = False
        End If
    
    Loop
    
    Close #fileNo
    
    ReadFile = vTable
    
End Function

