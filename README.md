# VBA Ray Casting Algorithm

Implementation of the ray casting algorithm in VBA, with some extra features like polygons sides intersection detection. 
 
## Description

The ray casting algorithm is used to check if a given point lies inside a polygon. The majority of ray casting algorithm implementations use traditional methods to compute lines intersections. These approaches need logic to handle all the specials cases. 

The solution presented here introduces the particularity that the intersections are calculated using [homogeneous coordinates](https://en.wikipedia.org/wiki/Homogeneous_coordinates). This feature allows the ray to be defined horizontally, vertically or in any direction on a given reference system, with slight modifications to the code of the proposed solution.

## Using examples

Download the Class Module and import it into your VBA project. Insert a new "normal" module and paste the following code:

```vb
Option Explicit
Sub testIrregularPolygon()
    Dim A() As Variant, b() As Variant, c As Variant, i As Integer, j As Integer
    Dim k As Integer
    Dim PointsToCheck() As Double
    Dim Polygon As PolygonShape
    
    ReDim A(0 To 13, 0 To 1)
    A(0, 0) = 2: A(0, 1) = 6
    A(1, 0) = -2: A(1, 1) = 2
    A(2, 0) = 0: A(2, 1) = -2
    A(3, 0) = 2: A(3, 1) = 0
    A(4, 0) = 6: A(4, 1) = 2
    A(5, 0) = 8: A(5, 1) = -2
    A(6, 0) = 4: A(6, 1) = -4
    A(7, 0) = 8: A(7, 1) = -6
    A(8, 0) = 12: A(8, 1) = -6
    A(9, 0) = 16: A(9, 1) = -2
    A(10, 0) = 12: A(10, 1) = 0
    A(11, 0) = 18: A(11, 1) = 0
    A(12, 0) = 16: A(12, 1) = 6
    A(13, 0) = 10: A(13, 1) = 4
    Set Polygon = New PolygonShape
    Polygon.OuterBoundary = A
    Polygon.ComputeProperties
    ReDim PointsToCheck(0 To 6, 0 To 1)
    PointsToCheck(0, 0) = 15.75: PointsToCheck(0, 1) = 5.5
    PointsToCheck(1, 0) = 5.75: PointsToCheck(1, 1) = 1.5
    PointsToCheck(2, 0) = 10: PointsToCheck(2, 1) = -5
    PointsToCheck(3, 0) = -1: PointsToCheck(3, 1) = 0.75
    PointsToCheck(4, 0) = 13.5: PointsToCheck(4, 1) = -0.5
    PointsToCheck(5, 0) = 7: PointsToCheck(5, 1) = 5
    PointsToCheck(6, 0) = -3: PointsToCheck(6, 1) = 2
    For k = LBound(PointsToCheck) To UBound(PointsToCheck)
        Debug.Print "Point In Polygon:"; Polygon.PointInPolygon(PointsToCheck(k, 0), PointsToCheck(k, 1))
        Debug.Print "*****************************************************************************************"
    Next k
    Set Polygon = Nothing
End Sub
Sub testRegularPolygon()
    Dim A() As Variant, b() As Variant, c As Variant, i As Integer, j As Integer
    Dim k As Integer
    Dim PointsToCheck() As Double
    Dim Polygon As PolygonShape
    
    ReDim A(0 To 9, 0 To 1)
    A(0, 0) = 6: A(0, 1) = 1
    A(1, 0) = 11: A(1, 1) = 1
    A(2, 0) = 15.05: A(2, 1) = 3.94
    A(3, 0) = 16.59: A(3, 1) = 8.69
    A(4, 0) = 15.05: A(4, 1) = 13.45
    A(5, 0) = 11: A(5, 1) = 16.39
    A(6, 0) = 6: A(6, 1) = 16.39
    A(7, 0) = 1.95: A(7, 1) = 13.45
    A(8, 0) = 0.41: A(8, 1) = 8.69
    A(9, 0) = 1.95: A(9, 1) = 3.94
    Set Polygon = New PolygonShape
    Polygon.OuterBoundary = A
    Polygon.ComputeProperties
    ReDim PointsToCheck(0 To 2, 0 To 1)
    PointsToCheck(0, 0) = -2: PointsToCheck(0, 1) = 8.69
    PointsToCheck(1, 0) = 4: PointsToCheck(1, 1) = 14.5
    PointsToCheck(2, 0) = 15.5: PointsToCheck(2, 1) = 3.75
    For k = LBound(PointsToCheck) To UBound(PointsToCheck)
        Debug.Print "Point In Polygon:"; Polygon.PointInPolygon(PointsToCheck(k, 0), PointsToCheck(k, 1))
        Debug.Print "*****************************************************************************************"
    Next k
    Set Polygon = Nothing
End Sub
```

This is the output returned after run the `testIrregularPolygon` procedure:

```vb
Starting check over point:(15.75, 5.5)...
Polygon AREA: 134 
Polygon BARYCENTER:(8.72636815920398, 0.606965174129353)
Line check at:(12, -6)|(16, -2)
Line check at:(16, -2)|(12, 0)
Line check at:(12, 0)|(18, 0)
Line check at:(18, 0)|(16, 6)
Intersection found in:(18, 0)|(16, 6)
Line check at:(16, 6)|(10, 4)
Point In Polygon:True
*****************************************************************************************
Starting check over point:(5.75, 1.5)...
Polygon AREA: 134 
Polygon BARYCENTER:(8.72636815920398, 0.606965174129353)
Line check at:(2, 0)|(6, 2)
Line check at:(6, 2)|(8, -2)
Intersection found in:(6, 2)|(8, -2)
Line check at:(8, -2)|(4, -4)
Line check at:(4, -4)|(8, -6)
Line check at:(8, -6)|(12, -6)
Line check at:(12, -6)|(16, -2)
Line check at:(16, -2)|(12, 0)
Line check at:(12, 0)|(18, 0)
Line check at:(18, 0)|(16, 6)
Intersection found in:(18, 0)|(16, 6)
Line check at:(16, 6)|(10, 4)
Line check at:(10, 4)|(2, 6)
Point In Polygon:False
*****************************************************************************************
Starting check over point:(10, -5)...
Polygon AREA: 134 
Polygon BARYCENTER:(8.72636815920398, 0.606965174129353)
Line check at:(8, -6)|(12, -6)
Line check at:(12, -6)|(16, -2)
Intersection found in:(12, -6)|(16, -2)
Line check at:(16, -2)|(12, 0)
Line check at:(12, 0)|(18, 0)
Line check at:(18, 0)|(16, 6)
Line check at:(16, 6)|(10, 4)
Line check at:(10, 4)|(2, 6)
Point In Polygon:True
*****************************************************************************************
Starting check over point:(-1, 0.75)...
Polygon AREA: 134 
Polygon BARYCENTER:(8.72636815920398, 0.606965174129353)
Line check at:(2, 6)|(-2, 2)
Line check at:(-2, 2)|(0, -2)
Line check at:(0, -2)|(2, 0)
Line check at:(2, 0)|(6, 2)
Intersection found in:(2, 0)|(6, 2)
Line check at:(6, 2)|(8, -2)
Intersection found in:(6, 2)|(8, -2)
Line check at:(8, -2)|(4, -4)
Line check at:(4, -4)|(8, -6)
Line check at:(8, -6)|(12, -6)
Line check at:(12, -6)|(16, -2)
Line check at:(16, -2)|(12, 0)
Line check at:(12, 0)|(18, 0)
Line check at:(18, 0)|(16, 6)
Intersection found in:(18, 0)|(16, 6)
Line check at:(16, 6)|(10, 4)
Line check at:(10, 4)|(2, 6)
Point In Polygon:True
*****************************************************************************************
Starting check over point:(13.5, -0.5)...
Polygon AREA: 134 
Polygon BARYCENTER:(8.72636815920398, 0.606965174129353)
Line check at:(12, -6)|(16, -2)
Line check at:(16, -2)|(12, 0)
Line check at:(12, 0)|(18, 0)
Line check at:(18, 0)|(16, 6)
Line check at:(16, 6)|(10, 4)
Point In Polygon:False
*****************************************************************************************
Starting check over point:(7, 5)...
Polygon AREA: 134 
Polygon BARYCENTER:(8.72636815920398, 0.606965174129353)
Line check at:(6, 2)|(8, -2)
Line check at:(8, -2)|(4, -4)
Line check at:(4, -4)|(8, -6)
Line check at:(8, -6)|(12, -6)
Line check at:(12, -6)|(16, -2)
Line check at:(16, -2)|(12, 0)
Line check at:(12, 0)|(18, 0)
Line check at:(18, 0)|(16, 6)
Intersection found in:(18, 0)|(16, 6)
Line check at:(16, 6)|(10, 4)
Intersection found in:(16, 6)|(10, 4)
Line check at:(10, 4)|(2, 6)
Point In Polygon:False
*****************************************************************************************
Starting check over point:(-3, 2)...
Polygon AREA: 134 
Polygon BARYCENTER:(8.72636815920398, 0.606965174129353)
Line check at:(2, 6)|(-2, 2)
Line check at:(-2, 2)|(0, -2)
Intersection found in:(-2, 2)|(0, -2)
Line check at:(0, -2)|(2, 0)
Line check at:(2, 0)|(6, 2)
Intersection found in:(2, 0)|(6, 2)
Line check at:(6, 2)|(8, -2)
Intersection found in:(6, 2)|(8, -2)
Line check at:(8, -2)|(4, -4)
Line check at:(4, -4)|(8, -6)
Line check at:(8, -6)|(12, -6)
Line check at:(12, -6)|(16, -2)
Line check at:(16, -2)|(12, 0)
Line check at:(12, 0)|(18, 0)
Line check at:(18, 0)|(16, 6)
Intersection found in:(18, 0)|(16, 6)
Line check at:(16, 6)|(10, 4)
Line check at:(10, 4)|(2, 6)
Point In Polygon:False
*****************************************************************************************
```

This is the output returned after run the `testRegularPolygon` procedure:

```vb
Starting check over point:(-2, 8.69)...
Polygon AREA: 192.4404 
Polygon BARYCENTER:(8.5, 8.69487316072924)
Line check at:(6, 1)|(11, 1)
Line check at:(11, 1)|(15.05, 3.94)
Line check at:(15.05, 3.94)|(16.59, 8.69)
Line check at:(16.59, 8.69)|(15.05, 13.45)
Intersection found in:(16.59, 8.69)|(15.05, 13.45)
Line check at:(15.05, 13.45)|(11, 16.39)
Line check at:(11, 16.39)|(6, 16.39)
Line check at:(6, 16.39)|(1.95, 13.45)
Line check at:(1.95, 13.45)|(0.41, 8.69)
Intersection found in:(1.95, 13.45)|(0.41, 8.69)
Line check at:(0.41, 8.69)|(1.95, 3.94)
Line check at:(1.95, 3.94)|(6, 1)
Point In Polygon:False
*****************************************************************************************
Starting check over point:(4, 14.5)...
Polygon AREA: 192.4404 
Polygon BARYCENTER:(8.5, 8.69487316072924)
Line check at:(6, 1)|(11, 1)
Line check at:(11, 1)|(15.05, 3.94)
Line check at:(15.05, 3.94)|(16.59, 8.69)
Line check at:(16.59, 8.69)|(15.05, 13.45)
Line check at:(15.05, 13.45)|(11, 16.39)
Intersection found in:(15.05, 13.45)|(11, 16.39)
Line check at:(11, 16.39)|(6, 16.39)
Line check at:(6, 16.39)|(1.95, 13.45)
Line check at:(1.95, 3.94)|(6, 1)
Point In Polygon:True
*****************************************************************************************
Starting check over point:(15.5, 3.75)...
Polygon AREA: 192.4404 
Polygon BARYCENTER:(8.5, 8.69487316072924)
Line check at:(15.05, 3.94)|(16.59, 8.69)
Line check at:(16.59, 8.69)|(15.05, 13.45)
Point In Polygon:False
*****************************************************************************************
```
