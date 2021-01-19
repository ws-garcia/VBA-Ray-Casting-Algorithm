VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PolygonShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Acumulator As Double
Private ArrayOfProperties() As Variant
Private i As Double
Private j As Double
Private k As Double
Private KeysArray As Variant 'Holds all the keys for the PopCollection collection
Private L As Double
Private m As Double
Private n As Double
Private PropertiesArray() As Variant
Private PropertiesCalcFlag As Boolean
Private SenseToMatch As VerticesSense
Private tmpVarArray() As Variant
Private VPCheck As Single 'Valid Points check for a 2d PointsArray

'Properties
Private P_ErrorDescription As String
Private P_ErrorNumber As Variant
Private P_ErrorSource As String
Private P_LastOperationDescription As String
Private P_LastOperationSuccess  As Boolean
Private P_OperationsAmount As Double
Private P_OuterBoundary As Variant
Private P_OuterBoundaryFlag As Boolean 'TRUE when the shape's Outer Boundary is defined
Private P_Properties As Variant 'Holds all the section properties
Private P_PropertiesColl As Collection 'Holds all the section properties [for provide one to one access]
'Enums
Public Enum VerticesSense
    ClockWise = 0
    CounterClockwise = 1
End Enum

Public Property Get ErrorDescription() As String
    ErrorDescription = P_ErrorDescription
End Property

Public Property Get ErrorNumber() As String
    ErrorNumber = CStr(P_ErrorNumber)
End Property

Public Property Get ErrorSource() As String
    ErrorSource = P_ErrorSource
End Property

Public Property Get LastOperationDescription() As String
    LastOperationDescription = P_LastOperationDescription
End Property

Public Property Get LastOperationSuccess() As Boolean
    LastOperationSuccess = P_LastOperationSuccess
End Property

Public Property Get OperationsAmount() As Double
    OperationsAmount = P_OperationsAmount
End Property

Public Property Get OuterBoundary() As Variant
    OuterBoundary = P_OuterBoundary
End Property

Public Property Let OuterBoundary(ByRef PointsArray As Variant)
    On Error Resume Next
    P_OuterBoundaryFlag = False
    If IsValid2DArray(PointsArray, CounterClockwise) Then 'Check the input
        P_OuterBoundaryFlag = True
        P_OuterBoundary = PointsArray
    End If
    'Begin error Handler section
    If Err.Number = 0 Then
        P_LastOperationSuccess = True
    Else
        P_LastOperationSuccess = False
        With Err
            P_ErrorNumber = .Number
            P_ErrorDescription = .Description
            P_ErrorSource = .Source
        End With
    End If
    Err.Clear
    P_LastOperationDescription = "Define external boundary"
    P_OperationsAmount = 1
End Property

Public Property Get OuterBoundaryFlag() As Boolean
    OuterBoundaryFlag = P_OuterBoundaryFlag
End Property

Public Property Get Properties() As Variant
    Properties = P_Properties
End Property

Public Property Get PropertiesColl() As Collection
    Set PropertiesColl = P_PropertiesColl
End Property

Private Function AreaAndCentroidOfPolygon(ByRef Boundary As Variant) As Variant()
    'This function Return the area and the centroid of a polygon given its vertices array.
    'Is mandatory to input two or more vertices, otherwise you will get a runtime error.
    'The results will be saved in a vector array with the following structure:
    '|AREA : CGX : CGy|
    ReDim ArrayOfProperties(2)
    L = UBound(Boundary, 1)
    For k = LBound(Boundary, 1) To L
        If k < L Then
            'Cumulate the Area
            ArrayOfProperties(0) = ArrayOfProperties(0) + _
                                                (Boundary(k, 0) * Boundary(k + 1, 1) - _
                                                Boundary(k + 1, 0) * Boundary(k, 1))
            'Cumulate Center Of Gravity (CGx)
            ArrayOfProperties(1) = ArrayOfProperties(1) + ((Boundary(k, 0) + Boundary(k + 1, 0)) * _
                                                ((Boundary(k, 0) * Boundary(k + 1, 1)) - (Boundary(k + 1, 0) * _
                                                Boundary(k, 1))))
            'Cumulate Center Of Gravity (CGy)
            ArrayOfProperties(2) = ArrayOfProperties(2) + ((Boundary(k, 1) + Boundary(k + 1, 1)) * _
                                                ((Boundary(k, 0) * Boundary(k + 1, 1)) - (Boundary(k + 1, 0) * _
                                                Boundary(k, 1))))
        Else
            'AREA
            ArrayOfProperties(0) = 0.5 * (ArrayOfProperties(0) + _
                                                (Boundary(k, 0) * Boundary(LBound(Boundary, 1), 1) - _
                                                Boundary(LBound(Boundary, 1), 0) * Boundary(k, 1)))
            'CENTER OF GRAVITY (CGx)
            ArrayOfProperties(1) = (ArrayOfProperties(1) + ((Boundary(k, 0) + _
                                                Boundary(LBound(Boundary, 1), 0)) * ((Boundary(k, 0) * _
                                                Boundary(LBound(Boundary, 1), 1)) - (Boundary(LBound(Boundary, 1), 0) * _
                                                Boundary(k, 1))))) / (6 * ArrayOfProperties(0))
            'CENTER OF GRAVITY (CGy)
            ArrayOfProperties(2) = (ArrayOfProperties(2) + ((Boundary(k, 1) + _
                                                Boundary(LBound(Boundary, 1), 1)) * ((Boundary(k, 0) * _
                                                Boundary(LBound(Boundary, 1), 1)) - (Boundary(LBound(Boundary, 1), 0) * _
                                                Boundary(k, 1))))) / (6 * ArrayOfProperties(0))
        End If
    Next k
    AreaAndCentroidOfPolygon = ArrayOfProperties
End Function

Public Function BoundaryPointsCage(ByRef Boundary As Variant) As Double()
    'This function returns an array with the coordinates of the vertices that
    'define the quadrilateral inside which all the vertices of a given polygon
    'can be contained. The Boundary argument needs to be a bidimensional array
    'with the [x] and [y] coordinates. The returned onedimentional array have
    'the estructure |Right[x] Top[y] Left[x] Bottom[y]|.
    
    Dim tmpCoordinates(0 To 3) As Double
    On Error Resume Next
    j = LBound(Boundary, 2)
    For i = LBound(Boundary, 1) To UBound(Boundary, 1)
        If i = LBound(Boundary, 1) Then 'Asign the first value to the variables
            tmpCoordinates(0) = CDbl(Boundary(i, j)) 'Right[x] Value
            tmpCoordinates(1) = CDbl(Boundary(i, j + 1)) 'Top[y] Value
            tmpCoordinates(2) = CDbl(Boundary(i, j)) 'Left[x] Value
            tmpCoordinates(3) = CDbl(Boundary(i, j + 1)) 'Bottom[y] Value
        Else
            If tmpCoordinates(0) < CDbl(Boundary(i, j)) Then
                tmpCoordinates(0) = CDbl(Boundary(i, j)) 'Choose the easternmost coordinate
            End If
            If tmpCoordinates(1) < CDbl(Boundary(i, j + 1)) Then
                tmpCoordinates(1) = CDbl(Boundary(i, j + 1)) 'Choose the northernmost ordinate
            End If
            If tmpCoordinates(2) > CDbl(Boundary(i, j)) Then
                tmpCoordinates(2) = CDbl(Boundary(i, j)) 'Choose the westernmost coordinate
            End If
            If tmpCoordinates(3) > CDbl(Boundary(i, j + 1)) Then
                tmpCoordinates(3) = CDbl(Boundary(i, j + 1)) 'Choose the southernmost ordinate
            End If
        End If
    Next i
    BoundaryPointsCage = tmpCoordinates
    'Begin error Handler section
    If Err.Number = 0 Then
        P_LastOperationSuccess = True
    Else
        P_LastOperationSuccess = False
        With Err
            P_ErrorNumber = .Number
            P_ErrorDescription = .Description
            P_ErrorSource = .Source
        End With
    End If
    Err.Clear
    P_LastOperationDescription = "Boundary point cage selection"
    P_OperationsAmount = 1
End Function

Public Sub ComputeProperties()
    'The returned array has the form:
    '|0    1   2  |
    '|AREA Cgx Cgy|
    Err.Clear
    On Error Resume Next
    PropertiesCalcFlag = False
    Erase PropertiesArray
    KeysArray = Array("Area", "Cgx", "Cgy")
    If P_OuterBoundaryFlag Then
        'Resize the VerticesSectionProperties array
        n = 2
        m = 3
        'Compute Outer Boundary properties
        PropertiesArray = AreaAndCentroidOfPolygon(P_OuterBoundary)
        'Fullfill the properties collection
        Set P_PropertiesColl = New Collection
        j = 0
        For i = LBound(KeysArray) To UBound(KeysArray)
            P_PropertiesColl.Add Item:=PropertiesArray(j), Key:=KeysArray(i)
            j = j + 1
        Next i
    Else
        'The procedure will not cumpute anything
    End If
    P_Properties = PropertiesArray
    'Begin error Handler section
    If Err.Number = 0 Then
        P_LastOperationSuccess = True
        PropertiesCalcFlag = True
    Else
        P_LastOperationSuccess = False
        With Err
            P_ErrorNumber = .Number
            P_ErrorDescription = .Description
            P_ErrorSource = .Source
        End With
    End If
    Err.Clear
    P_LastOperationDescription = "Computation of geometric properties of the section"
    P_OperationsAmount = 1
End Sub

Private Function GetIntersectionsNumber(LinesData() As Double, _
                                        RayData() As Double, FromPoint As Double) As Long
    Dim Delta1 As Double, Delta2 As Double
    Dim iCounter
    Dim Intersection As Variant
    Dim IntersectionCumulator As Long
    Dim LineMagnitude As Double
    Dim obLB As Long, obUB As Long
    Dim SideIntersection As Boolean
    Dim TestLine(0 To 1, 0 To 1) As Double
    
    obLB = LBound(P_OuterBoundary)
    obUB = UBound(P_OuterBoundary)
    For iCounter = obLB To obUB
            If LineInForward(LinesData(iCounter, 0), _
                                    LinesData(iCounter, 2), FromPoint) Then
                TestLine(0, 0) = LinesData(iCounter, 0)
                TestLine(0, 1) = LinesData(iCounter, 1)
                TestLine(1, 0) = LinesData(iCounter, 2)
                TestLine(1, 1) = LinesData(iCounter, 3)
                Debug.Print "Line check at:"; "(" & TestLine(0, 0) & ", " & TestLine(0, 1) & ")" _
                                                & "|"; "(" & TestLine(1, 0) & ", " & TestLine(1, 1) & ")"
                Intersection = Me.LinesIntersection(RayData, TestLine) 'Compute intersection
                If Not IsNull(Intersection) Then
                    LineMagnitude = Sqr((TestLine(1, 0) - TestLine(0, 0)) ^ 2 + _
                                            (TestLine(1, 1) - TestLine(0, 1)) ^ 2)
                    Delta1 = Sqr((TestLine(0, 0) - Intersection(0, 0)) ^ 2 + _
                                        (TestLine(0, 1) - Intersection(0, 1)) ^ 2)
                    Delta2 = Sqr((TestLine(1, 0) - Intersection(0, 0)) ^ 2 + _
                                        (TestLine(1, 1) - Intersection(0, 1)) ^ 2)
                    SideIntersection = (Round(Delta1 + Delta2, 2) = Round(LineMagnitude, 2))
                    If SideIntersection Then
                        If Intersection(0, 0) - FromPoint > 0 Then 'Only intersection in forward sense
                            IntersectionCumulator = IntersectionCumulator + 1
                            Debug.Print "Intersection found in:"; "(" & TestLine(0, 0) & ", " & TestLine(0, 1) & ")" _
                                                                & "|"; "(" & TestLine(1, 0) & ", " & TestLine(1, 1) & ")"
                        End If
                    End If
                End If
            End If
        Next iCounter
        GetIntersectionsNumber = IntersectionCumulator
End Function

Private Function GetPolygonSides(PolygonVertices As Variant) As Double()
    Dim iCounter
    Dim obLB As Long, obUB As Long
    Dim obLB2 As Long
    Dim tmpResult() As Double
    
    obLB = LBound(PolygonVertices)
    obUB = UBound(PolygonVertices)
    obLB2 = LBound(PolygonVertices, 2)
    ReDim tmpResult(obLB To obUB, 0 To 3)
    For iCounter = obLB To obUB
        tmpResult(iCounter, 0) = CDbl(PolygonVertices(iCounter, obLB2))
        tmpResult(iCounter, 1) = CDbl(PolygonVertices(iCounter, obLB2 + 1))
        If iCounter < obUB Then
            tmpResult(iCounter, 2) = CDbl(PolygonVertices(iCounter + 1, obLB2))
            tmpResult(iCounter, 3) = CDbl(PolygonVertices(iCounter + 1, obLB2 + 1))
        Else
            tmpResult(iCounter, 2) = CDbl(PolygonVertices(obLB, obLB2))
            tmpResult(iCounter, 3) = CDbl(PolygonVertices(obLB, obLB2 + 1))
        End If
    Next iCounter
    GetPolygonSides = tmpResult
End Function

Private Function IsSimpleDataType(V As Variant) As Boolean
    On Error Resume Next
    If IsArray(V) = True Then
        IsSimpleDataType = False
        Exit Function
    End If
    If IsObject(V) = True Then
        IsSimpleDataType = False
        Exit Function
    End If
    Select Case VarType(V)
        Case vbArray, vbDataObject, vbObject, vbUserDefinedType
            IsSimpleDataType = False
        Case Else
            IsSimpleDataType = True
    End Select

End Function

Private Function IsSimpleNumericType(V As Variant) As Boolean
    If IsSimpleDataType(V) = True Then
        Select Case VarType(V)
            Case vbBoolean, _
                vbByte, _
                vbCurrency, _
                vbDate, _
                vbDecimal, _
                vbDouble, _
                vbInteger, _
                vbLong, _
                vbSingle
            IsSimpleNumericType = True
            Case vbVariant
                If IsNumeric(V) = True Then
                    IsSimpleNumericType = True
                Else
                    IsSimpleNumericType = False
                End If
            Case Else
                IsSimpleNumericType = False
        End Select
    Else
        IsSimpleNumericType = False
    End If
End Function

Private Function IsValid2DArray(Array2D As Variant, Optional ExpectedVerticesSense As VerticesSense = 1, _
                                Optional MinimumRowLimit As Integer = 3) As Boolean
    IsValid2DArray = False
    n = 0
    If Not IsSimpleNumericType(Array2D) And NumberOfArrayDimensions(Array2D) = 2 Then 'Check for a valid 2d array
        For i = LBound(Array2D, 1) To UBound(Array2D, 1)
            For j = LBound(Array2D, 2) To UBound(Array2D, 2)
                If IsNumeric(Array2D(i, j)) And Not IsEmpty(Array2D(i, j)) Then
                    n = n + 1
                End If
            Next j
        Next i
        VPCheck = (n / 2) - Fix(n / 2) 'Check for valid data pairs
        If VPCheck = 0 And n / 2 >= MinimumRowLimit Then 'To require MinimumRowLimit or more valid points
            'Sort array in reverse if its vertices sense is not the expected one
            If Not MatchVerticesSense(Array2D, ExpectedVerticesSense) Then Array2D = SortRowsInReverse(Array2D)
            IsValid2DArray = True
        End If
    End If
End Function

Private Function LineInForward(X1 As Double, X2 As Double, _
                                xref As Double) As Boolean
    LineInForward = ((X1 > xref) Or (X2 > xref))
End Function

Public Function LinesIntersection(ByRef PointsInLine1 As Variant, PointsInLine2 As Variant, Optional k As Integer = 1) As Variant
    'This function return the intersection of two lines given two ordered pairs of points for each line.
    'The PointsInLine1 and PointsInLine2 arguments must have the array form |[x] [y]| with the same size (2 rows and 2 columns).
    'The returned value is an array |[x] [y]| if the lines meet and a Null value for parallel lines.
    
    Dim i As Double, j As Double, m As Double, n As Double, O As Double, P As Double
    Dim HomogeneousPointsL1(0 To 1, 0 To 2) As Variant, HomogeneousPointsL2(0 To 1, 0 To 2) As Variant
    Dim HomogeneousEcuations(0 To 1, 0 To 2) As Variant, EcuationsProduct(0 To 0, 0 To 2) As Variant
    Dim Intersection(0 To 0, 0 To 1) As Variant
    
    If k <= 0 Then k = 1 'Ensure positive scalar value
    m = LBound(PointsInLine1, 1)
    n = LBound(PointsInLine1, 2)
    'Represent the points with homogeneous coordinates
    i = 0
    For O = m To m + 1
        j = 0
        For P = n To n + 2
            If P < n + 2 Then
                HomogeneousPointsL1(i, j) = PointsInLine1(O, P)
                HomogeneousPointsL2(i, j) = PointsInLine2(O, P)
            Else
                HomogeneousPointsL1(i, j) = k
                HomogeneousPointsL2(i, j) = k
            End If
            j = j + 1
        Next P
        i = i + 1
    Next O
    'Compute the lines homogeneous ecuations using Vectorial Matrix Products
    HomogeneousEcuations(0, 0) = (HomogeneousPointsL1(0, 1) * HomogeneousPointsL1(1, 2)) - _
                                (HomogeneousPointsL1(1, 1) * HomogeneousPointsL1(0, 2)) '[a01*a12 - a11*a02]
    HomogeneousEcuations(0, 1) = -1 * ((HomogeneousPointsL1(0, 0) * HomogeneousPointsL1(1, 2)) - _
                                (HomogeneousPointsL1(1, 0) * HomogeneousPointsL1(0, 2))) '-[a00*a12 - a10*a02]
    HomogeneousEcuations(0, 2) = (HomogeneousPointsL1(0, 0) * HomogeneousPointsL1(1, 1)) - _
                                (HomogeneousPointsL1(1, 0) * HomogeneousPointsL1(0, 1)) '[a00*a11 - a10*a01]
    HomogeneousEcuations(1, 0) = (HomogeneousPointsL2(0, 1) * HomogeneousPointsL2(1, 2)) - _
                                (HomogeneousPointsL2(1, 1) * HomogeneousPointsL2(0, 2))
    HomogeneousEcuations(1, 1) = -1 * ((HomogeneousPointsL2(0, 0) * HomogeneousPointsL2(1, 2)) - _
                                (HomogeneousPointsL2(1, 0) * HomogeneousPointsL2(0, 2)))
    HomogeneousEcuations(1, 2) = (HomogeneousPointsL2(0, 0) * HomogeneousPointsL2(1, 1)) - _
                                (HomogeneousPointsL2(1, 0) * HomogeneousPointsL2(0, 1))
    'Compute the lines meet using Vectorial Matrix Products
    EcuationsProduct(0, 0) = (HomogeneousEcuations(0, 1) * HomogeneousEcuations(1, 2)) - _
                                (HomogeneousEcuations(1, 1) * HomogeneousEcuations(0, 2)) '[a01*a12 - a11*a02]
    EcuationsProduct(0, 1) = -1 * ((HomogeneousEcuations(0, 0) * HomogeneousEcuations(1, 2)) - _
                                (HomogeneousEcuations(1, 0) * HomogeneousEcuations(0, 2))) '-[a00*a12 - a10*a02]
    EcuationsProduct(0, 2) = (HomogeneousEcuations(0, 0) * HomogeneousEcuations(1, 1) - _
                                (HomogeneousEcuations(1, 0) * HomogeneousEcuations(0, 1))) '[a00*a11 - a10*a01]
    If EcuationsProduct(0, 2) = 0 Then
        LinesIntersection = Null 'Return a Null value. The lines meet at the infinity.
    Else
        Intersection(0, 0) = k * EcuationsProduct(0, 0) / EcuationsProduct(0, 2)
        Intersection(0, 1) = k * EcuationsProduct(0, 1) / EcuationsProduct(0, 2)
        LinesIntersection = Intersection
    End If
End Function

Public Function MatchVerticesSense(PointsArray As Variant, ExpectedSense As VerticesSense) As Boolean
    'This function return TRUE if the array's vertices sense match with the given one in ExpectedSense
    Acumulator = 0
    L = UBound(PointsArray, 1)
    For k = LBound(PointsArray, 1) To L
        If k < L Then
            Acumulator = Acumulator + ((PointsArray(k + 1, 0) - PointsArray(k, 0)) * _
                            (PointsArray(k + 1, 1) + PointsArray(k, 1)))
        Else
            Acumulator = Acumulator + ((PointsArray(LBound(PointsArray, 1), 0) - PointsArray(k, 0)) * _
                            (PointsArray(LBound(PointsArray, 1), 1) + PointsArray(k, 1)))
        End If
    Next k
    If Acumulator < 0 Then
        SenseToMatch = CounterClockwise
    Else
        SenseToMatch = ClockWise
    End If
    If ExpectedSense <> SenseToMatch Then
        MatchVerticesSense = False
    Else
        MatchVerticesSense = True
    End If
End Function

Private Function NumberOfArrayDimensions(Arr As Variant) As Integer
    Dim Ndx As Integer
    Dim Res As Integer
    On Error Resume Next
    Do
        Ndx = Ndx + 1
        Res = UBound(Arr, Ndx)
    Loop Until Err.Number <> 0

    NumberOfArrayDimensions = Ndx - 1

End Function

Public Function PointInPolygon(xCoordinate As Double, _
                                        yCoordinate As Double) As Boolean
    'This function returns True if the given coordinates lie on the area
    'of the polygon defined in the current instance
    Dim Cage() As Double
    Dim Epsilon As Double
    Dim IntersectionCounter As Long
    Dim LinesPoints() As Double
    Dim Ray(0 To 1, 0 To 1) As Double
    Dim RefPoint As Double
    Dim XC As Double, YC As Double
    
    On Error Resume Next
    If PropertiesCalcFlag Then
        'Centroid
        XC = P_PropertiesColl("Cgx")
        YC = P_PropertiesColl("Cgy")
        Debug.Print "Starting check over point:"; "(" & xCoordinate & ", " & yCoordinate & ")..."
        Debug.Print "Polygon AREA:"; P_PropertiesColl("Area")
        Debug.Print "Polygon BARYCENTER:"; "(" & XC & ", " & YC & ")"
        Cage() = BoundaryPointsCage(P_OuterBoundary) 'Cages all vertex
        'Represent the ray as a horizontal line
        Epsilon = ((YC - Cage(3)) / 100)  'Break symmetry
        If YC < yCoordinate Then 'point over the centroid
            Epsilon = -1 * Epsilon
        End If
        Ray(0, 0) = xCoordinate: Ray(0, 1) = yCoordinate + Epsilon
        Ray(1, 0) = Cage(2): Ray(1, 1) = yCoordinate + Epsilon
        'Lines data
        LinesPoints() = GetPolygonSides(P_OuterBoundary)
        'Lines intersection
        RefPoint = Ray(0, 0)
        IntersectionCounter = GetIntersectionsNumber(LinesPoints, Ray, RefPoint)
        PointInPolygon = ((IntersectionCounter Mod 2) = 1)
    End If
End Function

Public Function SidesOverlaps(ByVal PolyShape As PolygonShape) As Boolean
    On Error Resume Next
    If PropertiesCalcFlag Then
        If Not PolyShape Is Nothing Then
            Dim VerticesA As Variant, VerticesB As Variant
            Dim SidesA() As Double, SidesB() As Double
            Dim xCounter As Long, zCounter As Long
            Dim TestLineA(0 To 1, 0 To 1) As Double
            Dim TestLineB(0 To 1, 0 To 1) As Double
            Dim AtmpResult As Boolean, BtmpResult As Boolean
            Dim Intersection As Variant
            Dim ALineMagnitude As Double, BLineMagnitude As Double
            Dim Delta1 As Double, Delta2 As Double
            Dim Delta3 As Double, Delta4 As Double
            
            VerticesA = Me.OuterBoundary
            VerticesB = PolyShape.OuterBoundary
            SidesA() = GetPolygonSides(VerticesA)
            SidesB() = GetPolygonSides(VerticesB)
            For zCounter = LBound(SidesA) To UBound(SidesA)
                TestLineA(0, 0) = SidesA(zCounter, 0)
                TestLineA(0, 1) = SidesA(zCounter, 1)
                TestLineA(1, 0) = SidesA(zCounter, 2)
                TestLineA(1, 1) = SidesA(zCounter, 3)
                For xCounter = LBound(SidesB) To UBound(SidesB)
                    TestLineB(0, 0) = SidesB(xCounter, 0)
                    TestLineB(0, 1) = SidesB(xCounter, 1)
                    TestLineB(1, 0) = SidesB(xCounter, 2)
                    TestLineB(1, 1) = SidesB(xCounter, 3)
                    Debug.Print "Check sides intersection between sides:"; "{(" & TestLineA(0, 0) & ", " & TestLineA(0, 1) & ")" _
                                                & "|"; "(" & TestLineA(1, 0) & ", " & TestLineA(1, 1) & ")} and " _
                                                ; "{(" & TestLineB(0, 0) & ", " & TestLineB(0, 1) & ")" _
                                                & "|"; "(" & TestLineB(1, 0) & ", " & TestLineB(1, 1) & ")}"
                    Intersection = Me.LinesIntersection(TestLineA, TestLineB) 'Compute intersection
                    If Not IsNull(Intersection) Then
                        ALineMagnitude = Sqr((TestLineA(1, 0) - TestLineA(0, 0)) ^ 2 + _
                                                (TestLineA(1, 1) - TestLineA(0, 1)) ^ 2)
                        Delta1 = Sqr((TestLineA(0, 0) - Intersection(0, 0)) ^ 2 + _
                                            (TestLineA(0, 1) - Intersection(0, 1)) ^ 2)
                        Delta2 = Sqr((TestLineA(1, 0) - Intersection(0, 0)) ^ 2 + _
                                        (TestLineA(1, 1) - Intersection(0, 1)) ^ 2)
                        AtmpResult = (Round(Delta1 + Delta2, 2) = Round(ALineMagnitude, 2))
                        
                        BLineMagnitude = Sqr((TestLineB(1, 0) - TestLineB(0, 0)) ^ 2 + _
                                                (TestLineB(1, 1) - TestLineB(0, 1)) ^ 2)
                        Delta3 = Sqr((TestLineB(0, 0) - Intersection(0, 0)) ^ 2 + _
                                            (TestLineB(0, 1) - Intersection(0, 1)) ^ 2)
                        Delta4 = Sqr((TestLineB(1, 0) - Intersection(0, 0)) ^ 2 + _
                                        (TestLineB(1, 1) - Intersection(0, 1)) ^ 2)
                        BtmpResult = (Round(Delta3 + Delta4, 2) = Round(BLineMagnitude, 2))
                        If AtmpResult And BtmpResult Then
                            Debug.Print "Intersection found between sides:"; "{(" & TestLineA(0, 0) & ", " & TestLineA(0, 1) & ")" _
                                                & "|"; "(" & TestLineA(1, 0) & ", " & TestLineA(1, 1) & ")} and " _
                                                ; "{(" & TestLineB(0, 0) & ", " & TestLineB(0, 1) & ")" _
                                                & "|"; "(" & TestLineB(1, 0) & ", " & TestLineB(1, 1) & ")}"
                            GoTo EndTask
                        End If
                    End If
                Next xCounter
            Next zCounter
EndTask:
            SidesOverlaps = (AtmpResult And BtmpResult)
        End If
    End If
End Function
Public Function SingleProperty(PropertyName As String) As Variant
    On Error Resume Next
    SingleProperty = P_PropertiesColl.Item(PropertyName)
    'Begin error Handler section
    If Err.Number = 0 Then
        P_LastOperationSuccess = True
    Else
        P_LastOperationSuccess = False
        With Err
            P_ErrorNumber = .Number
            P_ErrorDescription = .Description
            P_ErrorSource = .Source
        End With
    End If
    Err.Clear
    P_LastOperationDescription = "Get polygon property"
End Function

Public Function SortRowsInReverse(Arr As Variant) As Variant
    ReDim tmpVarArray(LBound(Arr, 1) To UBound(Arr, 1), LBound(Arr, 2) To UBound(Arr, 2))
    L = UBound(Arr, 1)
    k = UBound(Arr, 2)
    For i = LBound(Arr, 1) To L
        For j = LBound(Arr, 2) To k
            tmpVarArray(i, j) = Arr(L - i, j)
        Next j
    Next i
    SortRowsInReverse = tmpVarArray
End Function

Public Sub TranslateAxisToCentroid()
    If PropertiesCalcFlag Then
        Dim i As Integer
        For i = LBound(P_OuterBoundary, 1) To UBound(P_OuterBoundary, 1)
            P_OuterBoundary(i, LBound(P_OuterBoundary, 2)) = P_OuterBoundary(i, LBound(P_OuterBoundary, 2)) - _
                                                            PropertiesArray(1)
            P_OuterBoundary(i, LBound(P_OuterBoundary, 2) + 1) = P_OuterBoundary(i, LBound(P_OuterBoundary, 2) + 1) - _
                                                            PropertiesArray(2)
        Next i
    End If
End Sub

Public Sub TranslateAxisToOrigin()
    If PropertiesCalcFlag Then
        Dim i As Integer
        For i = LBound(P_OuterBoundary, 1) To UBound(P_OuterBoundary, 1)
            P_OuterBoundary(i, LBound(P_OuterBoundary, 2)) = P_OuterBoundary(i, LBound(P_OuterBoundary, 2)) + _
                                                            PropertiesArray(1)
            P_OuterBoundary(i, LBound(P_OuterBoundary, 2) + 1) = P_OuterBoundary(i, LBound(P_OuterBoundary, 2) + 1) + _
                                                            PropertiesArray(2)
        Next i
    End If
End Sub

