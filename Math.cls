VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Math"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Original code made by Elroy Sullivan, PhD (User in VbForums)
' Was a module and now is a Class, by George Karras
' NOT FINISHED YET
' We can't pass UDT from M2000 script but we can pass a pointer to memory
' which we make using Buffer statement, and using Structure we can create offsets too.
Private mcheck As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
  ByVal lpvDest As Long, ByVal lpvSource As Long, ByVal cbCopy As Long)
Private Const PI = 3.14159265358979
Private Const pihalf = 1.5707963267949
Private Const pidivby180 = 1.74532925199433E-02
Private Const xyz As Double = 1               '       X+ to X-        Z- to Z+
Private Const xzy As Double = -1              '       X+ to X-        Y+ to Y-
Private Const yzx As Double = 2               '       Y+ to y-        X+ to X-
Private Const yxz As Double = -2              '       Y+ to Y-        Z- to Z+
Private Const zxy As Double = 3               '       Z+ to Z-        Y- to Y+
Private Const zyx As Double = -3           '       Z+ to Z-        X+ to X-

Public Type VecType
    ' This is used for both 3D spatial vectors, and Euler angles as well.
    X As Double
    y As Double
    Z As Double
End Type

Public Type SegType
    Origin As VecType
    AxisF  As VecType   ' Forward, West,  AxisX
    AxisL  As VecType   ' Left,    North, AxisY
    AxisU  As VecType   ' Up,      Up,    AxisZ
End Type
'
Public Type LineType
    point1 As VecType
    point2 As VecType
End Type
'
Public Type QuatType                      ' To get axis and theta back out (JPL convention):
    W As Double         ' cos(theta/2)              theta = ACos(q.w) * 2
    X As Double         ' v.x * sin(theta/2)        v.x = q.x / sin(theta/2)    solve for theta first, and then plug in here.
    y As Double         ' v.y * sin(theta/2)        v.y = q.y / sin(theta/2)
    Z As Double         ' v.z * sin(theta/2)        v.z = q.z / sin(theta/2)
End Type
Public Sub Vector(a As Long, X As Double, y As Double, Z As Double)
Dim V1 As VecType
V1.X = X
V1.y = y
V1.Z = Z
CopyMemory a, VarPtr(V1.X), 24
End Sub
Private Function GetVector(a As Long) As VecType
    CopyMemory VarPtr(GetVector), a, 24
End Function
Private Sub PutVector(v As VecType, a As Long)
    CopyMemory a, VarPtr(v), 24
End Sub
Private Function GetRealType(a As Long) As Double
    CopyMemory VarPtr(GetRealType), a, 8
End Function
Private Function GetQuatType(a As Long) As QuatType
    CopyMemory VarPtr(GetQuatType), a, 32
End Function
Private Sub PutQuatType(q As QuatType, a As Long)
    CopyMemory a, VarPtr(q), 32
End Sub
Private Function GetLine(a As Long) As LineType
    CopyMemory VarPtr(GetLine), a, 48
End Function
Private Sub PutLine(q As LineType, a As Long)
    CopyMemory a, VarPtr(q), 48
End Sub
Public Function DotProduct(a As Long, b As Long) As Double
    Dim V1 As VecType, V2 As VecType
    CopyMemory VarPtr(V1), a, 24
    CopyMemory VarPtr(V2), b, 24
    DotProduct = V1.X * V2.X + V1.y * V2.y + V1.Z * V2.Z
End Function
Public Sub XProduct(a As Long, b As Long, c As Long)
    Dim V1 As VecType, V2 As VecType, V3 As VecType
    CopyMemory VarPtr(V1), a, 24
    CopyMemory VarPtr(V2), b, 24
    V3.X = V1.y * V2.Z - V1.Z * V2.y
    V3.y = V1.Z * V2.X - V1.X * V2.Z
    V3.Z = V1.X * V2.y - V1.y * V2.X
    PutVector V3, c
End Sub
Public Sub RotVect(a As Long, b As Long, c As Long, Angle As Double, Optional bDegrees As Boolean = True)  '(vec As VecType, axis As VecType, angle As Double, Optional bDegrees As Boolean = True) As VecType
    Dim vec As VecType, axis As VecType, RotVect As VecType
    CopyMemory VarPtr(vec), a, 24
    CopyMemory VarPtr(axis), b, 24
    RotVect = RotVecByQuat(vec, AxisAngle2Quat(axis, Angle, bDegrees))
    CopyMemory c, VarPtr(RotVect), 24
End Sub
Public Sub RotVectMult(n As Long, a As Long, b As Long, c As Long, Angle As Double, Optional bDegrees As Boolean = True)  '(vec As VecType, axis As VecType, angle As Double, Optional bDegrees As Boolean = True) As VecType
    Dim vec As VecType, axis As VecType, RotVect As VecType, m As Long
    CopyMemory VarPtr(axis), b, 24
    For m = 1 To n
    CopyMemory VarPtr(vec), a, 24
    RotVect = RotVecByQuat(vec, AxisAngle2Quat(axis, Angle, bDegrees))
    CopyMemory c, VarPtr(RotVect), 24
    a = a + 24
    c = c + 24
    Next m
End Sub

'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'
' Make Segment (ordered basis) functions.
'
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************

Public Function SegFrom3Pts(Origin As VecType, point1 As VecType, point2 As VecType, Optional Order As Double = xyz) As SegType
    ' BodyBuilderSegment = [Origin, Point1 - Origin, Point2 - Origin, Order]
    SegFrom3Pts = SegFrom5Pts(Origin, point1, Origin, point2, Origin, Order)
End Function

Public Function SegFrom5Pts(Origin As VecType, Point1Line1 As VecType, Point2Line1 As VecType, Point1Line2 As VecType, Point2Line2 As VecType, Optional Order As Double = xyz) As SegType
    ' BodyBuilderSegment = [Origin, Point2Line1 - Point1Line2, Point2Line2 - Point2Line2, Order]
    '
    SegFrom5Pts = SegFromLines(Origin, MakeLine(Point1Line1, Point2Line1), MakeLine(Point1Line2, Point2Line2), Order)
End Function

Public Function SegFromLines(Origin As VecType, Line1 As LineType, Line2 As LineType, Optional Order As Double = xyz) As SegType
    ' This returns an ordered basis (following the right-hand-rule) with all vectors being a unit-in-length (orthonormal).
    '
    ' This function is designed to replicate the line in BodyBuilder that allows you to create a segment from an origin and two lines (but just for one frame):
    '       Segment = [Origin, Line1.Point1 - Line1.Point2, Line2.Point1 - Line2.Point2, Order]
    '
    ' Note that the output axes are set at origin <0,0,0>.  Only the origin reflects an offset translation.
    ' The return axes will always be orthonormal.
    '
    '                                                                        (Up)
    '                                                                          Z
    ' The figure at right attempts to illustrate one of these segments:        |    Y (Left, North)
    ' This is the output of this function.                                     |   /
    '                                                                          |  /
    '                                                                          | /
    '                                                                          |/
    '                                                                 (Origin) *------- X (Forward, East)
    '
    ' The input is an origin and two lines.  The origin really stands alone and often isn't terribly important,
    ' especially if we're just concerned with rotations.
    '
    ' The actual axes of the segment are constructed from the two lines, and things are constructed differently depending on the specified order.
    ' The first line (Line1) defines the first letter of the order of the segment, with the positive direction defined by Line1.Point1
    ' (and the line's origin defined by Line1.Point2).  The following may help:
    '
    '       For Xyz(1) & Xzy(-1)    AxisX is defined by Line1 with positive heading to Line1.Point1.
    '       For Yzx(2) & Yxz(-2)    AxisY is defined by Line1 with positive heading to Line1.Point1.
    '       For Zxy(3) & Zyx(-3)    AxisZ is defined by Line1 with positive heading to Line1.Point1.
    '
    ' The second line (Line2) approximates the last (third) letter of the order, with the alphabetic order specifying whether Line2.Point1 or Line2.Point2 is positive.
    ' With respect to Line2, the following may provide as a cheat-sheet:
    '
    '       For xyZ(1)      AxisZ is approximated by Line2 with positive heading to Line2.Point1.
    '       For xzY(-1)     AxisY is approximated by Line2 with positive heading to Line2.Point2.
    '       For yzX(2)      AxisX is approximated by Line2 with positive heading to Line2.Point1.
    '       For yxZ(-2)     AxisZ is approximated by Line2 with positive heading to Line2.Point2.
    '       For zxY(3)      AxisY is approximated by Line2 with positive heading to Line2.Point1.
    '       For zyX(-3)     AxisX is approximated by Line2 with positive heading to Line2.Point2.
    '
    ' Now, we've still got another axis of our segment to address.  For instance, for XYZ, we've yet to talk about Y.
    ' This "middle" axis is defined by the cross product of the other two axes.
    ' Also when discussing the last axis, I used the word "approximate".  I did this because, once the middle axis is determined,
    ' this "last" axis is recalculated as the cross product of the first two.  This is the way an orthogonal segment is achieved.
    ' The following are the actual calculations (with × denoting a cross product):
    '
    '       LineA = (Line1.Point1 - Line1.Point2)          ; Unitized.
    '       LineB = (Line2.Point1 - Line2.Point2)          ; Unitized.
    '
    '       If constant is XYZ (1):         AxisX = LineA
    '                                       AxisY = LineB × LineA
    '                                       AxisZ = LineA × AxisY
    '
    '       If constant is XZY (-1):        AxisX = LineA
    '                                       AxisZ = LineB × LineA
    '                                       AxisY = AxisZ × LineA
    '
    '       If constant is YZX (2):         AxisY = LineA
    '                                       AxisZ = LineB × LineA
    '                                       AxisX = LineA × AxisZ
    '
    '       If constant is YXZ (-2):        AxisY = LineA
    '                                       AxisX = LineB × LineA
    '                                       AxisZ = AxisX × LineA
    '
    '       If constant is ZXY (3):         AxisZ = LineA
    '                                       AxisX = LineB × LineA
    '                                       AxisY = LineA × AxisX
    '
    '       If constant is ZYX (-3):        AxisZ = LineA
    '                                       AxisY = LineB × LineA
    '                                       AxisX = AxisY × LineA
    '
    ' And finally, here's another cheat-sheet that may help:
    '
    '       Segment = [Point0, (L1.p1 - L1.p2), (L2.p1 - L2.p2), xyz]       (1)
    '                  Origin     + AxisX -        + AxisZ -
    '
    '       Segment = [Point0, (L1.p1 - L1.p2), (L2.p1 - L2.p2), xzy]       (-1)
    '                  Origin     + AxisX -        - AxisY +
    '
    '       Segment = [Point0, (L1.p1 - L1.p2), (L2.p1 - L2.p2), yzx]       (2)
    '                  Origin     + AxisY -        + AxisX -
    '
    '       Segment = [Point0, (L1.p1 - L1.p2), (L2.p1 - L2.p2), yxz]       (-2)
    '                  Origin     + AxisY -        - AxisZ +
    '
    '       Segment = [Point0, (L1.p1 - L1.p2), (L2.p1 - L2.p2), zxy]       (3)
    '                  Origin     + AxisZ -        + AxisY -
    '
    '       Segment = [Point0, (L1.p1 - L1.p2), (L2.p1 - L2.p2), zyx]       (-3)
    '                  Origin     + AxisZ -        - AxisX +
    '
    Dim V1 As VecType
    Dim V2 As VecType
    '
    SegFromLines.Origin = Origin
    '
    V1 = VecDif(Line1.point1, Line1.point2)
    V2 = VecDif(Line2.point1, Line2.point2)
    V1 = UnitVec(V1)
    V2 = UnitVec(V2)
    '
    Select Case Order
    Case xyz    ' 1
        SegFromLines.AxisF = V1                                                         ' X
        SegFromLines.AxisL = UnitVec(XProd(V2, SegFromLines.AxisF))                      ' Y
        SegFromLines.AxisU = UnitVec(XProd(SegFromLines.AxisF, SegFromLines.AxisL))     ' Z
    Case xzy    ' -1
        SegFromLines.AxisF = V1                                                         ' X
        SegFromLines.AxisU = UnitVec(XProd(V2, SegFromLines.AxisF))                     ' Z
        SegFromLines.AxisL = UnitVec(XProd(SegFromLines.AxisU, SegFromLines.AxisF))     ' Y
    Case yzx    ' 2
        SegFromLines.AxisL = V1                                                         ' Y
        SegFromLines.AxisU = UnitVec(XProd(V2, SegFromLines.AxisL))                     ' Z
        SegFromLines.AxisF = UnitVec(XProd(SegFromLines.AxisL, SegFromLines.AxisU))     ' X
    Case yxz    ' -2
        SegFromLines.AxisL = V1                                                         ' Y
        SegFromLines.AxisF = UnitVec(XProd(V2, SegFromLines.AxisL))                     ' X
        SegFromLines.AxisU = UnitVec(XProd(SegFromLines.AxisF, SegFromLines.AxisL))     ' Z
    Case zxy    ' 3
        SegFromLines.AxisU = V1                                                         ' Z
        SegFromLines.AxisF = UnitVec(XProd(V2, SegFromLines.AxisU))                     ' X
        SegFromLines.AxisL = UnitVec(XProd(SegFromLines.AxisU, SegFromLines.AxisF))     ' Y
    Case zyx    ' -3
        SegFromLines.AxisU = V1                                                         ' Z
        SegFromLines.AxisL = UnitVec(XProd(V2, SegFromLines.AxisU))                     ' Y
        SegFromLines.AxisF = UnitVec(XProd(SegFromLines.AxisL, SegFromLines.AxisU))     ' X
    End Select
End Function

'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'
' Conversion functions.
'
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Public Function Euler2Quat(euler As VecType, Optional Order As Double = xyz, Optional bDegrees As Boolean = True) As QuatType
    ' To get the same return from Quat2Euler, the following Euler angle constraints should be observed:
    '       pi   >  euler.x  >  -pi            ' In theory, 180 (or -180) should be possible, but not here.
    '       pi/2 >  euler.y  >   pi/2          ' Avoid pi/2 to avoid gimbal lock.
    '       pi   >  euler.z  >  -pi            ' In theory, 180 (or -180) should be possible, but not here.
    '
    Dim sX As Double, sY As Double, SZ As Double
    Dim cx As Double, cy As Double, cZ As Double
    Dim EulerTemp As VecType
    '
    EulerTemp = euler
    If bDegrees Then EulerTemp = VecDeg2Rad(EulerTemp)
    '
    ' Speed-up variables, so that Sin and Cos aren't repeatedly called.
    sX = Sin(EulerTemp.X / 2#): sY = Sin(EulerTemp.y / 2#): SZ = Sin(EulerTemp.Z / 2#)
    cx = Cos(EulerTemp.X / 2#): cy = Cos(EulerTemp.y / 2#): cZ = Cos(EulerTemp.Z / 2#)
    '
    Select Case Order
    Case xyz
        Euler2Quat.W = cx * cy * cZ - sX * sY * SZ
        Euler2Quat.X = sX * cy * cZ + cx * sY * SZ
        Euler2Quat.y = cx * sY * cZ - sX * cy * SZ
        Euler2Quat.Z = cx * cy * SZ + sX * sY * cZ
    Case xzy
        Euler2Quat.W = cx * cy * cZ + sX * sY * SZ
        Euler2Quat.X = sX * cy * cZ - cx * sY * SZ
        Euler2Quat.y = cx * cy * SZ - sX * sY * cZ
        Euler2Quat.Z = cx * sY * cZ + sX * cy * SZ
    Case yzx
        Euler2Quat.W = cx * cy * cZ - sX * sY * SZ
        Euler2Quat.X = cx * cy * SZ + sX * sY * cZ
        Euler2Quat.y = sX * cy * cZ + cx * sY * SZ
        Euler2Quat.Z = cx * sY * cZ - sX * cy * SZ
    Case yxz
        Euler2Quat.W = cx * cy * cZ + sX * sY * SZ
        Euler2Quat.X = cx * sY * cZ + sX * cy * SZ
        Euler2Quat.y = sX * cy * cZ - cx * sY * SZ
        Euler2Quat.Z = cx * cy * SZ - sX * sY * cZ
    Case zxy
        Euler2Quat.W = cx * cy * cZ - sX * sY * SZ
        Euler2Quat.X = cx * sY * cZ - sX * cy * SZ
        Euler2Quat.y = cx * cy * SZ + sX * sY * cZ
        Euler2Quat.Z = sX * cy * cZ + cx * sY * SZ
    Case zyx
        Euler2Quat.W = cx * cy * cZ + sX * sY * SZ
        Euler2Quat.X = cx * cy * SZ - sX * sY * cZ
        Euler2Quat.y = cx * sY * cZ + sX * cy * SZ
        Euler2Quat.Z = sX * cy * cZ - cx * sY * SZ
    End Select
End Function

Public Function Quat2Euler(q As QuatType, Optional Order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' The order (xyzEnum) may be reversed to some mathematicians.
    ' However, it's done this way so that the same xyzEnum constant can be used for Euler2Quat and Quat2Euler, and return to the same Euler angles.
    ' Also, this is the way BodyBuilder does it, and these procedures are meant to reproduce that functionality.
    '
    Select Case Order
    Case xyz: Quat2Euler = Quat2EulerHelper(-2# * (q.y * q.Z - q.W * q.X), q.W * q.W - q.X * q.X - q.y * q.y + q.Z * q.Z, 2# * (q.X * q.Z + q.W * q.y), -2# * (q.X * q.y - q.W * q.Z), q.W * q.W + q.X * q.X - q.y * q.y - q.Z * q.Z)
    Case xzy: Quat2Euler = Quat2EulerHelper(2# * (q.y * q.Z + q.W * q.X), q.W * q.W - q.X * q.X + q.y * q.y - q.Z * q.Z, -2# * (q.X * q.y - q.W * q.Z), 2# * (q.X * q.Z + q.W * q.y), q.W * q.W + q.X * q.X - q.y * q.y - q.Z * q.Z)
    Case yzx: Quat2Euler = Quat2EulerHelper(-2# * (q.X * q.Z - q.W * q.y), q.W * q.W + q.X * q.X - q.y * q.y - q.Z * q.Z, 2# * (q.X * q.y + q.W * q.Z), -2# * (q.y * q.Z - q.W * q.X), q.W * q.W - q.X * q.X + q.y * q.y - q.Z * q.Z)
    Case yxz: Quat2Euler = Quat2EulerHelper(2# * (q.X * q.Z + q.W * q.y), q.W * q.W - q.X * q.X - q.y * q.y + q.Z * q.Z, -2# * (q.y * q.Z - q.W * q.X), 2# * (q.X * q.y + q.W * q.Z), q.W * q.W - q.X * q.X + q.y * q.y - q.Z * q.Z)
    Case zxy: Quat2Euler = Quat2EulerHelper(-2# * (q.X * q.y - q.W * q.Z), q.W * q.W - q.X * q.X + q.y * q.y - q.Z * q.Z, 2# * (q.y * q.Z + q.W * q.X), -2# * (q.X * q.Z - q.W * q.y), q.W * q.W - q.X * q.X - q.y * q.y + q.Z * q.Z)
    Case zyx: Quat2Euler = Quat2EulerHelper(2# * (q.X * q.y + q.W * q.Z), q.W * q.W + q.X * q.X - q.y * q.y - q.Z * q.Z, -2# * (q.X * q.Z - q.W * q.y), 2# * (q.y * q.Z + q.W * q.X), q.W * q.W - q.X * q.X - q.y * q.y + q.Z * q.Z)
    End Select
    '
    If bDegrees Then Quat2Euler = VecRad2Deg(Quat2Euler)
End Function

Private Function Quat2EulerHelper(r11 As Double, r12 As Double, r21 As Double, r31 As Double, r32 As Double) As VecType
    Quat2EulerHelper.X = Atan2(r11, r12)
    Quat2EulerHelper.y = Asin(r21)
    Quat2EulerHelper.Z = Atan2(r31, r32)
End Function

Public Function Seg2Quat(seg As SegType) As QuatType
    ' Ignores origin.
    '
    Seg2Quat = Axes2Quat(seg.AxisF, seg.AxisL, seg.AxisU)
End Function

Public Function Quat2Seg(q As QuatType) As SegType
    ' Leaves the segment origin as zero.
    '
    Quat2Seg.AxisF = Quat2Fwd(q)
    Quat2Seg.AxisL = Quat2Left(q)
    Quat2Seg.AxisU = Quat2Up(q)
End Function

Public Function Euler2Seg(euler As VecType, Optional Order As Double = xyz, Optional bDegrees As Boolean = True) As SegType
    ' Leaves the segment origin as zero.
    Euler2Seg = Quat2Seg(Euler2Quat(euler, Order, bDegrees))
End Function

Public Function Seg2Euler(seg As SegType, Optional Order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' Ignores origin.
    Seg2Euler = Quat2Euler(Seg2Quat(seg), Order, bDegrees)
End Function

Public Function Axes2Quat(AxisF As VecType, AxisL As VecType, AxisU As VecType) As QuatType
    ' Must be three orthogonal unit vectors.
    ' Basically, the vectors that define a BodyBuilder segment.
    '
    Dim s As Double
    Dim tr As Double
    Dim Max As Double
    Dim X As Double, y As Double, Z As Double
    '
    tr = (AxisF.X + AxisL.y + AxisU.Z + 1#)
    '
    If (tr >= 1#) Then
        s = (0.5 / Sqr(tr))
        Axes2Quat.W = 0.25 / s
        Axes2Quat.X = (AxisL.Z - AxisU.y) * s
        Axes2Quat.y = (AxisU.X - AxisF.Z) * s
        Axes2Quat.Z = (AxisF.y - AxisL.X) * s
    Else
        If (AxisL.y > AxisU.Z) Then Max = AxisL.y Else Max = AxisU.Z
        '
        If (Max < AxisF.X) Then
            s = (Sqr(AxisF.X - (AxisL.y + AxisU.Z) + 1#))
            X = (s * 0.5)
            s = (0.5 / s)
            Axes2Quat.W = (AxisL.Z - AxisU.y) * s
            Axes2Quat.X = X
            Axes2Quat.y = (AxisF.y + AxisL.X) * s
            Axes2Quat.Z = (AxisU.X + AxisF.Z) * s
        ElseIf (Max = AxisL.y) Then
            s = (Sqr(AxisL.y - (AxisU.Z + AxisF.X) + 1#))
            y = (s * 0.5)
            s = (0.5 / s)
            Axes2Quat.W = (AxisU.X - AxisF.Z) * s
            Axes2Quat.X = (AxisF.y + AxisL.X) * s
            Axes2Quat.y = y
            Axes2Quat.Z = (AxisL.Z + AxisU.y) * s
        Else
            s = (Sqr(AxisU.Z - (AxisF.X + AxisL.y) + 1#))
            Z = (s * 0.5)
            s = (0.5 / s)
            Axes2Quat.W = (AxisF.y - AxisL.X) * s
            Axes2Quat.X = (AxisU.X + AxisF.Z) * s
            Axes2Quat.y = (AxisL.Z + AxisU.y) * s
            Axes2Quat.Z = Z
        End If
    End If
End Function

Public Function Quat2Fwd(q As QuatType) As VecType
    ' Make sure quat is normed.
    Quat2Fwd = RotVecByQuat(MakeVec(1, 0, 0), q)
End Function

Public Function Quat2Left(q As QuatType) As VecType
    ' Make sure quat is normed.
    Quat2Left = RotVecByQuat(MakeVec(0, 1, 0), q)
End Function

Public Function Quat2Up(q As QuatType) As VecType
    ' Make sure quat is normed.
    Quat2Up = RotVecByQuat(MakeVec(0, 0, 1), q)
End Function

Public Function AxisAngle2Quat(axis As VecType, ByVal Angle As Double, Optional bDegrees As Boolean = True) As QuatType
    ' You should be sure the vector is a UNIT vector before calling this, else the quaternion will not be a unit quaternion.
    ' Note that the Axis(vector) is the rotation vector, and has nothing to do with basis vectors.
    '
    Dim SinThetaDivTwo As Double
    '
    If bDegrees Then Angle = Deg2Rad(Angle)
    SinThetaDivTwo = Sin(Angle / 2)
    '
    AxisAngle2Quat.W = Cos(Angle / 2)
    AxisAngle2Quat.X = axis.X * SinThetaDivTwo
    AxisAngle2Quat.y = axis.y * SinThetaDivTwo
    AxisAngle2Quat.Z = axis.Z * SinThetaDivTwo
End Function

Public Function QuatAxis(quat As QuatType) As VecType
    ' This is the rotation axis of the quaternion.
    ' It has nothing to do with the quaternion's basis axes.
    Dim SinThetaByTwo As Double
    '
    SinThetaByTwo = Sin(QuatAngle(quat, False) / 2)
    '
    On Error Resume Next
        QuatAxis.X = quat.X / SinThetaByTwo
        QuatAxis.y = quat.y / SinThetaByTwo
        QuatAxis.Z = quat.Z / SinThetaByTwo
    On Error GoTo 0
End Function

Public Function QuatAngle(quat As QuatType, Optional bDegrees As Boolean = True) As Double
    ' Just the angle from the quaternion.
    QuatAngle = Acos(quat.W) * 2
    If bDegrees Then QuatAngle = Rad2Deg(QuatAngle)
End Function

'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'
' Rotation functions.
'
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************

Public Function RotSeg(seg As SegType, axis As VecType, Angle As Double, Optional bDegrees As Boolean = True) As SegType
    ' This is exactly like a rotation: (Seg2 = Rot(Seg1, Axis, Angle)).
    '
    RotSeg = RotSegByQuat(seg, AxisAngle2Quat(axis, Angle, bDegrees))
End Function

Public Function RotSegByQuat(seg As SegType, quat As QuatType) As SegType
    RotSegByQuat.AxisF = RotVecByQuat(seg.AxisF, quat)
    RotSegByQuat.AxisL = RotVecByQuat(seg.AxisL, quat)
    RotSegByQuat.AxisU = RotVecByQuat(seg.AxisU, quat)
    RotSegByQuat.Origin = seg.Origin
End Function

Public Function RotQuat(q1 As QuatType, q2 As QuatType) As QuatType
    ' q1 is being rotated by q2 amount.
    '
    RotQuat.W = q1.W * q2.W - q1.X * q2.X - q1.y * q2.y - q1.Z * q2.Z
    RotQuat.X = q1.W * q2.X + q1.X * q2.W - q1.y * q2.Z + q1.Z * q2.y
    RotQuat.y = q1.W * q2.y + q1.X * q2.Z + q1.y * q2.W - q1.Z * q2.X
    RotQuat.Z = q1.W * q2.Z - q1.X * q2.y + q1.y * q2.X + q1.Z * q2.W
End Function
Public Function MulQuat(q1 As QuatType, q2 As QuatType) As QuatType
    ' q1 is being rotated by q2 amount.
    '
    MulQuat.W = q1.W * q2.W - q1.X * q2.X - q1.y * q2.y - q1.Z * q2.Z
    MulQuat.X = q1.W * q2.X + q1.X * q2.W + q1.y * q2.Z - q1.Z * q2.y
    MulQuat.y = q1.W * q2.y - q1.X * q2.Z + q1.y * q2.W + q1.Z * q2.X
    MulQuat.Z = q1.W * q2.Z + q1.X * q2.y - q1.y * q2.X + q1.Z * q2.W
End Function
Public Function UnRotQuat(q1 As QuatType, q2 As QuatType) As QuatType
    ' q1 is being UNrotated by q2 amount.
    '
    UnRotQuat = RotQuat(q1, NegQuat(q2))
End Function

Public Function RotVec(vec As VecType, axis As VecType, Angle As Double, Optional bDegrees As Boolean = True) As VecType
    RotVec = RotVecByQuat(vec, AxisAngle2Quat(axis, Angle, bDegrees))
End Function

Public Function RotVecByQuat(vec As VecType, quat As QuatType) As VecType
    ' This is essentially a vector rotation.
    ' To emulate BodyBuilder: NewVec = RotVecByQuat(Vec, AxisAng2Quat(Axis, Deg2Rad(Angle)))
    '
    Dim num1 As Double, num2 As Double, num3 As Double
    Dim num4 As Double, num5 As Double, num6 As Double
    Dim num7 As Double, num8 As Double, num9 As Double
    Dim numa As Double, numb As Double, numc As Double
    '
    num1 = quat.X * 2#:       num2 = quat.y * 2#:       num3 = quat.Z * 2#
    num4 = quat.X * num1:     num5 = quat.y * num2:     num6 = quat.Z * num3
    num7 = quat.X * num2:     num8 = quat.X * num3:     num9 = quat.y * num3
    numa = quat.W * num1:     numb = quat.W * num2:     numc = quat.W * num3
    '
    ' RotVecByQuat.x = (1# - (num5 + num6)) * vec.x +       (num7 - numc)  * vec.y +       (num8 + numb)  * vec.z
    ' RotVecByQuat.y =       (num7 + numc)  * vec.x + (1# - (num4 + num6)) * vec.y +       (num9 - numa)  * vec.z
    ' RotVecByQuat.z =       (num8 - numb)  * vec.x +       (num9 + numa)  * vec.y + (1# - (num4 + num5)) * vec.z
    '
    RotVecByQuat.X = (1# - (num5 + num6)) * vec.X + (num7 - numc) * vec.y + (num8 + numb) * vec.Z
    RotVecByQuat.y = (num7 + numc) * vec.X + (1# - (num4 + num6)) * vec.y + (num9 - numa) * vec.Z
    RotVecByQuat.Z = (num8 - numb) * vec.X + (num9 + numa) * vec.y + (1# - (num4 + num5)) * vec.Z
End Function

'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'
' Angles Between functions.
'
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************

Public Function FixedBetweenSegs(seg1 As SegType, seg2 As SegType, Optional Order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' In BodyBuilder terms, this is:  FixedBetweenSegs = <seg1, seg2, order>
    ' This takes the "Fixed" angle approach rather than the "Euler" angle approach.
    ' Conceptually, seg1 & seg2 are reversed with respect to what's rotated against what for this "Fixed" approach,
    ' but it's left this way to correspond to BodyBuilder code.
    '
    FixedBetweenSegs = NegVec(EulerBetweenSegs(seg1, seg2, Order, bDegrees))
End Function

Public Function EulerBetweenSegs(seg1 As SegType, seg2 As SegType, Optional Order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' In BodyBuilder terms, this is:  EulerBetweenSegs = -<seg1, seg2, order>
    ' We do NOT need to negate (as BodyBuilder does) when using this function.
    '
    EulerBetweenSegs = EulerBetweenQuats(Seg2Quat(seg1), Seg2Quat(seg2), Order, bDegrees)
End Function

Public Function EulerBetweenQuats(q1 As QuatType, q2 As QuatType, Optional Order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' Put q2's rotation basis at <0,0,0>, and then convert to Euler angles.
    ' In other words, what Euler angles will rotate q1 to q2.
    '
    EulerBetweenQuats = Quat2Euler(UnRotQuat(q2, q1), Order, bDegrees)
End Function

'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'
' Simple Linear Algebra functions.
'
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Public Function VecAvg(V1 As VecType, V2 As VecType) As VecType
    VecAvg = VecDivNum(VecSum(V1, V2), 2)
End Function
Public Sub VecAver(a As Long, b As Long, c As Long)
    Dim V1 As VecType
    V1 = VecDivNum(VecSumm1(a, b), 2)
    CopyMemory c, VarPtr(V1), 24
End Sub
Public Sub VecAverMult(n As Long, a As Long, b As Long, c As Long)
    Dim V1 As VecType, m As Long
    For m = 1 To n
    V1 = VecDivNum(VecSumm1(a, b), 2)
    CopyMemory c, VarPtr(V1), 24
    a = a + 24
    b = b + 24
    c = c + 24
    Next m
End Sub
Private Function VecSumm1(a As Long, b As Long) As VecType
    Dim V1 As VecType, V2 As VecType, V3 As VecType
    CopyMemory VarPtr(V1), a, 24
    CopyMemory VarPtr(V2), b, 24
    VecSumm1.X = V1.X + V2.X
    VecSumm1.y = V1.y + V2.y
    VecSumm1.Z = V1.Z + V2.Z
    
End Function

Public Function VecSum(V1 As VecType, V2 As VecType) As VecType
    VecSum.X = V1.X + V2.X
    VecSum.y = V1.y + V2.y
    VecSum.Z = V1.Z + V2.Z
End Function
Public Sub VecSumm(a As Long, b As Long, c As Long)
    Dim V1 As VecType, V2 As VecType, V3 As VecType
    CopyMemory VarPtr(V1), a, 24
    CopyMemory VarPtr(V2), b, 24
    V3.X = V1.X + V2.X
    V3.y = V1.y + V2.y
    V3.Z = V1.Z + V2.Z
    CopyMemory c, VarPtr(V3), 24
End Sub
Public Sub VecSummMult(n As Long, a As Long, b As Long, c As Long)
    Dim V1 As VecType, V2 As VecType, V3 As VecType, m As Long
    CopyMemory VarPtr(V2), b, 24
    For m = 1 To n
    CopyMemory VarPtr(V1), a, 24
    V3.X = V1.X + V2.X
    V3.y = V1.y + V2.y
    V3.Z = V1.Z + V2.Z
    CopyMemory c, VarPtr(V3), 24
    a = a + 24
    c = c + 24
    Next m
End Sub
Public Function VecDif(V1 As VecType, V2 As VecType) As VecType
    VecDif.X = V1.X - V2.X
    VecDif.y = V1.y - V2.y
    VecDif.Z = V1.Z - V2.Z
End Function
Public Sub VecDiff(a As Long, b As Long, c As Long)
    Dim V1 As VecType, V2 As VecType, V3 As VecType
    CopyMemory VarPtr(V1), a, 24
    CopyMemory VarPtr(V2), b, 24
    V3.X = V1.X - V2.X
    V3.y = V1.y - V2.y
    V3.Z = V1.Z - V2.Z
    CopyMemory c, VarPtr(V3), 24
End Sub
Public Sub VecDiffMult(n As Long, a As Long, b As Long, c As Long)
    Dim V1 As VecType, V2 As VecType, V3 As VecType, m As Long
    CopyMemory VarPtr(V2), b, 24
    For m = 1 To n
    CopyMemory VarPtr(V1), a, 24
    V3.X = V1.X - V2.X
    V3.y = V1.y - V2.y
    V3.Z = V1.Z - V2.Z
    CopyMemory c, VarPtr(V3), 24
    a = a + 24
    c = c + 24
    Next m
End Sub


Sub VecDiffNum(a As Long, D As Double, b As Long)
    Dim v As VecType, V1 As VecType
    CopyMemory VarPtr(V1), a, 24
    v.X = V1.X - D
    v.y = V1.y - D
    v.Z = V1.Z - D
    CopyMemory b, VarPtr(v), 24
End Sub
Sub VecDiffNumMult(n As Long, a As Long, D As Double, b As Long)
    Dim v As VecType, V1 As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(V1), a, 24
    v.X = V1.X - D
    v.y = V1.y - D
    v.Z = V1.Z - D
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub

Public Function VecAddNum(v As VecType, D As Double) As VecType
    VecAddNum.X = v.X + D
    VecAddNum.y = v.y + D
    VecAddNum.Z = v.Z + D
End Function
Sub VecAddiNum(a As Long, D As Double, b As Long)
    Dim v As VecType, V1 As VecType
    CopyMemory VarPtr(V1), a, 24
    v.X = V1.X + D
    v.y = V1.y + D
    v.Z = V1.Z + D
    CopyMemory b, VarPtr(v), 24
End Sub
Sub VecAddiNumMult(n As Long, a As Long, D As Double, b As Long)
    Dim v As VecType, V1 As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(V1), a, 24
    v.X = V1.X + D
    v.y = V1.y + D
    v.Z = V1.Z + D
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub

Sub VecDiviNum(a As Long, D As Double, b As Long)
    Dim v As VecType, V1 As VecType
    CopyMemory VarPtr(V1), a, 24
    v.X = V1.X / D
    v.y = V1.y / D
    v.Z = V1.Z / D
    CopyMemory b, VarPtr(v), 24
End Sub
Sub VecDiviNumMult(n As Long, a As Long, D As Double, b As Long)
    Dim v As VecType, V1 As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(V1), a, 24
    v.X = V1.X / D
    v.y = V1.y / D
    v.Z = V1.Z / D
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub
Public Function VecDivNum(V1 As VecType, D As Double) As VecType
    VecDivNum.X = V1.X / D
    VecDivNum.y = V1.y / D
    VecDivNum.Z = V1.Z / D
End Function
Sub VecMulNum(a As Long, D As Double, b As Long)
    Dim v As VecType, V1 As VecType
    CopyMemory VarPtr(V1), a, 24
    v.X = V1.X * D
    v.y = V1.y * D
    v.Z = V1.Z * D
    CopyMemory b, VarPtr(v), 24
End Sub
Sub VecMulNumMult(n As Long, a As Long, D As Double, b As Long)
    Dim v As VecType, V1 As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(V1), a, 24
    v.X = V1.X * D
    v.y = V1.y * D
    v.Z = V1.Z * D
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub
Public Function VecMultNum(V1 As VecType, D As Double) As VecType
    VecMultNum.X = V1.X * D
    VecMultNum.y = V1.y * D
    VecMultNum.Z = V1.Z * D
End Function

Public Function VecMagnitude(a As Long) As Double
    Dim v As VecType
    CopyMemory VarPtr(v), a, 24
    VecMagnitude = Sqr(v.X * v.X + v.y * v.y + v.Z * v.Z)
End Function
Public Function VecMag(v As VecType) As Double
    VecMag = Sqr(v.X * v.X + v.y * v.y + v.Z * v.Z)
End Function

Public Function NegVec(v As VecType) As VecType
    NegVec.X = -v.X
    NegVec.y = -v.y
    NegVec.Z = -v.Z
End Function
Public Sub UnitVect(a As Long)
    ' Returns ZERO vector if vector has no magnitude.
    Dim v As VecType, d1 As VecType
    Dim D As Double
    CopyMemory VarPtr(v), a, 24
    '
    D = VecMag(v)
    If D > 0 Then
        d1.X = v.X / D
        d1.y = v.y / D
        d1.Z = v.Z / D
    End If
    CopyMemory a, VarPtr(d1), 24
End Sub
Public Sub UnitVectMult(n As Long, a As Long)
    ' Returns ZERO vector if vector has no magnitude.
    Dim v As VecType, d1 As VecType, Zero As VecType, m As Long
    Dim D As Double
    For m = 1 To n
    CopyMemory VarPtr(v), a, 24
    '
    D = VecMag(v)
    If D > 0 Then
        d1.X = v.X / D
        d1.y = v.y / D
        d1.Z = v.Z / D
    Else
        d1 = Zero
    End If
    CopyMemory a, VarPtr(d1), 24
    a = a + 24
    Next m
End Sub
Public Function UnitVec(v As VecType) As VecType
    ' Returns ZERO vector if vector has no magnitude.
    Dim D As Double
    '
    D = VecMag(v)
    If D > 0 Then
        UnitVec.X = v.X / D
        UnitVec.y = v.y / D
        UnitVec.Z = v.Z / D
    End If
End Function
' XProduct
Public Function XProd(V1 As VecType, V2 As VecType) As VecType
    XProd.X = V1.y * V2.Z - V1.Z * V2.y
    XProd.y = V1.Z * V2.X - V1.X * V2.Z
    XProd.Z = V1.X * V2.y - V1.y * V2.X
End Function
' XDotProduct
Public Function DotProd(V1 As VecType, V2 As VecType) As Double
    DotProd = V1.X * V2.X + V1.y * V2.y + V1.Z * V2.Z
End Function
Sub NegateQuat(a As Long, b As Long)
    Dim quat As QuatType
    CopyMemory VarPtr(quat), a, 48
'    quat.w = quat.w
    quat.X = -quat.X
    quat.y = -quat.y
    quat.Z = -quat.Z
    CopyMemory b, VarPtr(quat), 48
End Sub
Sub NegateQuatMult(n As Long, a As Long, b As Long)
    Dim quat As QuatType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(quat), a, 48
'    quat.w = quat.w
    quat.X = -quat.X
    quat.y = -quat.y
    quat.Z = -quat.Z
    CopyMemory b, VarPtr(quat), 48
    a = a + 48
    b = b + 48
    Next m
End Sub
Public Function NegQuat(quat As QuatType) As QuatType
    ' There are actually two ways to do this, negate the angle or negate the axis.
    ' It turns out that negating the axis works a bit better with the RotQuat function
    NegQuat.W = quat.W
    NegQuat.X = -quat.X
    NegQuat.y = -quat.y
    NegQuat.Z = -quat.Z
End Function
Public Function NegQuatAll(quat As QuatType) As QuatType
    NegQuatAll.W = -quat.W
    NegQuatAll.X = -quat.X
    NegQuatAll.y = -quat.y
    NegQuatAll.Z = -quat.Z
End Function
Public Sub MakeLineFrom2Vec(a As Long, b As Long, c As Long) ' Pointers VecType, VecType, LineType
    CopyMemory c, a, 24
    CopyMemory c + 24, a, 24
End Sub
Public Function MakeLine(point1 As VecType, point2 As VecType) As LineType
    MakeLine.point1 = point1
    MakeLine.point2 = point2
End Function

Public Function MakeVec(X As Double, y As Double, Z As Double) As VecType
    MakeVec.X = X
    MakeVec.y = y
    MakeVec.Z = Z
End Function
Public Sub Quaternion(a As Long, W As Double, X As Double, y As Double, Z As Double)
    CopyMemory a, VarPtr(W), 8
    CopyMemory a + 8, VarPtr(X), 8
    CopyMemory a + 16, VarPtr(y), 8
    CopyMemory a + 24, VarPtr(Z), 8
End Sub
Public Function MakeQuat(W As Double, X As Double, y As Double, Z As Double) As QuatType
    MakeQuat.W = W
    MakeQuat.X = X
    MakeQuat.y = y
    MakeQuat.Z = Z
End Function
' quad functions
Public Sub QuatNeg(a As Long, Optional b As Long = 0)
    Dim q As QuatType
    q = NegQuatAll(GetQuatType(a))
    If b = 0 Then CopyMemory a, VarPtr(q), 32: Exit Sub
    CopyMemory b, VarPtr(q), 32
End Sub
Public Sub QuatNegMul(n As Long, a As Long, Optional b As Long = 0)
    Dim q As QuatType, m As Long
    For m = 1 To n
    q = NegQuatAll(GetQuatType(a))
    If b = 0 Then CopyMemory a, VarPtr(q), 32 Else CopyMemory b, VarPtr(q), 32: b = b + 32
    a = a + 32
    Next m
End Sub
Public Sub QuatMultQuat(a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, q2 As QuatType
    q = GetQuatType(a)
    q2 = GetQuatType(b)
    q = MulQuat(q, q2)
    If c = 0 Then CopyMemory a, VarPtr(q), 32: Exit Sub
    CopyMemory c, VarPtr(q), 32
End Sub
Public Sub QuatMultQuatMul(n As Long, a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, q2 As QuatType, m As Long
    If c = 0 Then
    For m = 1 To n
        q = GetQuatType(a)
        q2 = GetQuatType(b)
        q = MulQuat(q, q2)
        CopyMemory a, VarPtr(q), 32
        a = a + 32
        b = b + 32
    Next m
    Else
    For m = 1 To n
        q = GetQuatType(a)
        q2 = GetQuatType(b)
        q = MulQuat(q, q2)
        CopyMemory c, VarPtr(q), 32
        c = c + 32
        a = a + 32
        b = b + 32
    Next m
    End If
End Sub
Public Sub QuatMultByReal(a As Long, R As Double, Optional c As Long = 0)
    Dim q As QuatType
    q = GetQuatType(a)
    q.W = q.W * R
    q.X = q.X * R
    q.y = q.y * R
    q.Z = q.Z * R
    If c = 0 Then CopyMemory a, VarPtr(q), 32: Exit Sub
    CopyMemory c, VarPtr(q), 32
End Sub
Public Sub QuatMultReal(a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, R As Double
    q = GetQuatType(a)
    R = GetRealType(b)
    q.W = q.W * R
    q.X = q.X * R
    q.y = q.y * R
    q.Z = q.Z * R
    If c = 0 Then CopyMemory a, VarPtr(q), 32: Exit Sub
    CopyMemory c, VarPtr(q), 32
End Sub
Public Sub QuatMultRealMul(n As Long, a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, R As Double, m As Long
    For m = 1 To n
    q = GetQuatType(a)
    R = GetRealType(b)
    q.W = q.W * R
    q.X = q.X * R
    q.y = q.y * R
    q.Z = q.Z * R
    If c = 0 Then CopyMemory a, VarPtr(q), 32 Else CopyMemory c, VarPtr(q), 32: c = c + 32
    a = a + 32
    b = b + 8
    Next m
End Sub
Public Sub QuatMultRealMulOne(n As Long, a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, R As Double, m As Long
    R = GetRealType(b)
    For m = 1 To n
    q = GetQuatType(a)
    q.W = q.W * R
    q.X = q.X * R
    q.y = q.y * R
    q.Z = q.Z * R
    If c = 0 Then CopyMemory a, VarPtr(q), 32 Else CopyMemory c, VarPtr(q), 32: c = c + 32
    a = a + 32
    Next m
End Sub
Public Sub QuatAddByReal(a As Long, R As Double, Optional c As Long = 0)
    Dim q As QuatType
    If c = 0 Then
        q = GetQuatType(a)
        q.W = q.W + R
        CopyMemory a, VarPtr(q), 8
    Else
        q = GetQuatType(a)
        q.W = q.W + R
        CopyMemory c, VarPtr(q), 32
    End If
End Sub
Public Sub QuatAddReal(a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType
    If c = 0 Then
        q.W = GetRealType(a) + GetRealType(b)
        CopyMemory a, VarPtr(q), 8
    Else
        q = GetQuatType(a)
        q.W = q.W + GetRealType(b)
        CopyMemory c, VarPtr(q), 32
    End If
End Sub
Public Sub QuatAddRealMulConst(n As Long, a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, m As Long
    If c = 0 Then
        For m = 1 To n
            q.W = GetRealType(a) + GetRealType(b)
            CopyMemory a, VarPtr(q), 8
            a = a + 32
        Next m
    Else
        For m = 1 To n
            q = GetQuatType(a)
            q.W = q.W + GetRealType(b)
            CopyMemory c, VarPtr(q), 32
            c = c + 32
            a = a + 32
        Next m
    End If
End Sub
Public Sub QuatAddRealMulMul(n As Long, a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, m As Long
    If c = 0 Then
        For m = 1 To n
            q.W = GetRealType(a) + GetRealType(b)
            CopyMemory a, VarPtr(q), 8
            a = a + 32
            b = b + 8
        Next m
    Else
        For m = 1 To n
            q = GetQuatType(a)
            q.W = q.W + GetRealType(b)
            CopyMemory c, VarPtr(q), 32
            c = c + 32
            a = a + 32
            b = b + 8
        Next m
    End If
End Sub
Public Sub QuatAddQuat(a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, q2 As QuatType
    q = GetQuatType(a)
    q2 = GetQuatType(b)
    With q
        .W = .W + q2.W
        .X = .X + q2.X
        .y = .y + q2.y
        .Z = .Z + q2.Z
    End With
    If c = 0 Then CopyMemory a, VarPtr(q), 32: Exit Sub
    CopyMemory c, VarPtr(q), 32
End Sub
Public Sub QuatAddQuatMul(n As Long, a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, q2 As QuatType, m As Long
    If c = 0 Then
    For m = 1 To n
        q = GetQuatType(a)
        q2 = GetQuatType(b)
        With q
            .W = .W + q2.W
            .X = .X + q2.X
            .y = .y + q2.y
            .Z = .Z + q2.Z
        End With
        CopyMemory a, VarPtr(q), 32
        a = a + 32
        b = b + 32
    Next m
    Else
    For m = 1 To n
        q = GetQuatType(a)
        q2 = GetQuatType(b)
        With q
            .W = .W + q2.W
            .X = .X + q2.X
            .y = .y + q2.y
            .Z = .Z + q2.Z
        End With
        CopyMemory c, VarPtr(q), 32
        a = a + 32
        b = b + 32
        c = c + 32
    Next m
    End If
End Sub
Public Sub QuatConj(a As Long, Optional b As Long = 0)
    Dim q As QuatType
    q = NegQuat(GetQuatType(a))
    If b = 0 Then CopyMemory a, VarPtr(q), 32: Exit Sub
    CopyMemory b, VarPtr(q), 32
End Sub
Public Sub QuatConjMul(n As Long, a As Long, Optional b As Long = 0)
    Dim q As QuatType, m As Long
    For m = 1 To n
    q = NegQuat(GetQuatType(a))
    If b = 0 Then CopyMemory a, VarPtr(q), 32 Else CopyMemory b, VarPtr(q), 32: b = b + 32
    a = a + 32
    Next m
End Sub
Public Function QuatNormRetDouble(a As Long) As Double
    QuatNormRetDouble = NormQuat(GetQuatType(a))
End Function
Public Sub QuatNorm(a As Long, b As Long)
    Dim Z As Double
    Z = NormQuat(GetQuatType(a))
    CopyMemory b, VarPtr(Z), 8
End Sub
Public Sub QuatNormMul(n As Long, a As Long, b As Long)
    Dim Z As Double, m As Long
    For m = 1 To n
    Z = NormQuat(GetQuatType(a))
    CopyMemory b, VarPtr(Z), 8
    a = a + 32
    b = b + 8
    Next m
End Sub
Public Function NormQuat(v As QuatType) As Double
    NormQuat = Sqr(v.W * v.W + v.X * v.X + v.y * v.y + v.Z * v.Z)
End Function
Public Function QuatStringQuat(q As QuatType, Optional frm$ = "#0.00000") As String
    QuatStringQuat = MyFrm(q.W, frm$) & ", " & MyFrm(q.X, frm$) & ", " & MyFrm(q.y, frm$) & ", " & MyFrm(q.Z, frm$)
End Function
Public Function QuatString(a As Long, Optional frm$ = "#0.00000") As String
    Dim q As QuatType
    q = GetQuatType(a)
    QuatString = MyFrm(q.W, frm$) & ", " & MyFrm(q.X, frm$) & ", " & MyFrm(q.y, frm$) & ", " & MyFrm(q.Z, frm$)
End Function
Public Function VecStringVec(v As VecType, Optional frm$ = "#0.00000") As String
    VecStringVec = MyFrm(v.X, frm$) & ", " & MyFrm(v.y, frm$) & ", " & MyFrm(v.Z, frm$)
End Function
Public Function VecString(a As Long, Optional frm$ = "#0.00000") As String
    Dim v As VecType
    v = GetVector(a)
    VecString = MyFrm(v.X, frm$) & ", " & MyFrm(v.y, frm$) & ", " & MyFrm(v.Z, frm$)
End Function
Public Function SegString(seg As SegType) As String
    SegString = "Origin: " & VecStringVec(seg.Origin) & vbCrLf & _
                "AxisF : " & VecStringVec(seg.AxisF) & vbCrLf & _
                "AxisL : " & VecStringVec(seg.AxisL) & vbCrLf & _
                "AxisU : " & VecStringVec(seg.AxisU)
End Function
' trig functions
Public Sub VectorRad2Deg(a As Long, Optional b As Long = 0)
    Dim v As VecType
    v = GetVector(a)
    v.X = Rad2Deg(v.X)
    v.y = Rad2Deg(v.y)
    v.Z = Rad2Deg(v.Z)
    If b = 0 Then CopyMemory a, VarPtr(v), 24: Exit Sub
    CopyMemory b, VarPtr(v), 24
End Sub
Public Sub VectorRad2DegMult(n As Long, a As Long, b As Long)
    Dim v As VecType, m As Long
    For m = 1 To n
    v = GetVector(a)
    v.X = Rad2Deg(v.X)
    v.y = Rad2Deg(v.y)
    v.Z = Rad2Deg(v.Z)
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub

Public Sub VectorDeg2Rad(a As Long, Optional b As Long = 0)
    Dim v As VecType
    CopyMemory VarPtr(v), a, 24
    v.X = Deg2Rad(v.X)
    v.y = Deg2Rad(v.y)
    v.Z = Deg2Rad(v.Z)
    If b = 0 Then CopyMemory a, VarPtr(v), 24: Exit Sub
    CopyMemory b, VarPtr(v), 24
End Sub
Public Sub VectorDeg2RadMult(n As Long, a As Long, b As Long)
    Dim v As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(v), a, 24
    v.X = Deg2Rad(v.X)
    v.y = Deg2Rad(v.y)
    v.Z = Deg2Rad(v.Z)
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub

Public Function VecRad2Deg(v As VecType) As VecType
    VecRad2Deg.X = Rad2Deg(v.X)
    VecRad2Deg.y = Rad2Deg(v.y)
    VecRad2Deg.Z = Rad2Deg(v.Z)
End Function

Public Function VecDeg2Rad(v As VecType) As VecType
    VecDeg2Rad.X = Deg2Rad(v.X)
    VecDeg2Rad.y = Deg2Rad(v.y)
    VecDeg2Rad.Z = Deg2Rad(v.Z)
End Function

Public Function Rad2Deg(RadAngle As Double) As Double
    Rad2Deg = RadAngle / pidivby180
End Function

Public Function Deg2Rad(DegAngle As Double) As Double
    Deg2Rad = DegAngle * pidivby180
End Function

Public Function Acos(ByVal D As Double) As Double
    ' Radians.
    ' d is forced between 1 and -1.  This prevents IEEE rounding from causing any problems.
    Dim dRoot As Double
    Dim d2 As Double
    '
    Select Case True
    Case D > 1:     d2 = 1
    Case D < -1:    d2 = -1
    Case Else:      d2 = D
    End Select
    '
    dRoot = Sqr(1 - d2 * d2)
    If Abs(dRoot) <> 0 Then
        Acos = pihalf - Atn(d2 / dRoot)
    Else
        Acos = pihalf - pihalf * Sgn(d2)
    End If
End Function

Public Function Asin(D As Double) As Double
    ' Radians.
    ' d is forced between 1 and -1.  This prevents IEEE rounding from causing any problems.
    Dim dRoot As Double
    Dim d2 As Double
    '
    Select Case True
    Case D > 1:     d2 = 1
    Case D < -1:    d2 = -1
    Case Else:      d2 = D
    End Select
    '
    dRoot = Sqr(1 - d2 * d2)
    If Abs(dRoot) <> 0 Then
        Asin = Atn(d2 / dRoot)
    Else
        Asin = pihalf * Sgn(d2)
    End If
End Function

Public Function Atan2(y As Double, X As Double) As Double
    Const epsilon As Double = 1.4E-45
    ' Radians.
    '
    Select Case X
    Case Is > epsilon:  Atan2 = Atn(y / X)
    Case Is < -epsilon: Atan2 = Atn(y / X) + PI * Sgn(y): If y = 0# Then Atan2 = Atan2 + PI
    Case Else:          Atan2 = pihalf * Sgn(y)
    End Select
End Function
Public Property Get CheckThis() As Boolean
CheckThis = mcheck
End Property
Public Property Let CheckThis(RHS As Boolean)
mcheck = RHS
End Property
Private Function MyFrm(W As Double, frm$) As String
If frm$ = vbNullString Then
    MyFrm = LTrim$(str$(W))
            If Left$(MyFrm, 1) = "." Then
                    MyFrm = "0" + MyFrm
            ElseIf Left$(MyFrm, 2) = "-." Then
                    MyFrm = "-0" + Mid$(MyFrm, 2)
            End If
Else

        MyFrm = Format$(W, frm$)
        If Not NoUseDec Then
            If OverideDec Then
                MyFrm = Replace$(MyFrm, GetDeflocaleString(LOCALE_SDECIMAL), Chr(2))
                MyFrm = Replace$(MyFrm, GetDeflocaleString(LOCALE_STHOUSAND), Chr(3))
                MyFrm = Replace$(MyFrm, Chr(2), NowDec$)
                MyFrm = Replace$(MyFrm, Chr(3), NowThou$)
                If Left$(MyFrm, 1) = NowDec$ Then
                    MyFrm = "0" + MyFrm
                ElseIf Left$(MyFrm, 2) = "-" + NowDec$ Then
                    MyFrm = "-0" + Mid$(MyFrm, 2)
                End If
            ElseIf InStr(MyFrm, NowDec$) > 0 And InStr(frm$, ".") > 0 Then
                ElseIf InStr(MyFrm, NowDec$) > 0 Then
                MyFrm = Replace$(MyFrm, NowDec$, Chr(2))
                MyFrm = Replace$(MyFrm, NowThou$, Chr(3))
                MyFrm = Replace$(MyFrm, Chr(2), ".")
                MyFrm = Replace$(MyFrm, Chr(3), ",")
                If Left$(MyFrm, 1) = "." Then
                    MyFrm = "0" + MyFrm
                ElseIf Left$(MyFrm, 2) = "-." Then
                    MyFrm = "-0" + Mid$(MyFrm, 2)
                End If
            End If
        End If
    End If
End Function

Property Get Version() As Variant
Version = 1.3
End Property
Public Function ArcCos(X As Variant) As Variant

    Select Case X
        Case -1
            ArcCos = 4 * Atn(1)
             
        Case 0:
            ArcCos = 2 * Atn(1)
             
        Case 1:
            ArcCos = 0#
             
        Case Else:
            ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    End Select
End Function

Public Function ArcSin(X As Variant) As Variant

    Select Case X
        Case -1
            ArcSin = 6 * Atn(1)
             
        Case 0:
            ArcSin = 0#
             
        Case 1:
            ArcSin = 2 * Atn(1)
             
        Case Else:
            ArcSin = Atn(X / Sqr(-X * X + 1))
    End Select
End Function
Public Function Cos(X As Double) As Double
        Cos = VBA.Cos(X)
End Function
Public Function Sin(X As Double) As Double
        Sin = VBA.Sin(X)
End Function
Public Function Atn(X As Double) As Double
        Atn = VBA.Atn(X)
End Function
Public Function EXP(X As Double) As Double
        EXP = VBA.EXP(X)
End Function
' arc cotangent
' error if NUMBER is zero

Public Function ACot(Value As Double) As Double
    ACot = Atn(1 / Value)
End Function

' arc secant
' error if value is inside the range [-1,1]

Public Function Asec(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ASec = ACos(1 / value)
    If Abs(Value) <> 1 Then
        Asec = 1.5707963267949 - Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        Asec = 3.14159265358979 * Sgn(Value)
    End If
End Function

' arc cosecant
' error if value is inside the range [-1,1]

Public Function Acsc(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
    If Abs(Value) <> 1 Then
        Acsc = Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        Acsc = 1.5707963267949 * Sgn(Value)
    End If
End Function

' hyperbolic sine

Public Function Sinh(Value As Double) As Double
    Dim Temp As Double
    Temp = EXP(Value)
    Sinh = (Temp - 1 / Temp) / 2
End Function

' hyperbolic cosine

Public Function Cosh(Value As Double) As Double
    Dim Temp As Double
    Temp = EXP(Value)
    Cosh = (Temp + 1 / Temp) / 2
End Function

' hyperbolic tangent

Public Function Tanh(Value As Double) As Double
    Dim Temp As Double
    Temp = EXP(Value)
    Tanh = (Temp - 1 / Temp) / (Temp + 1 / Temp)
End Function

' hyperbolic cotangent

Public Function CotH(Value As Double) As Double
    Dim Temp As Double
    Temp = EXP(Value)
    CotH = (Temp + 1 / Temp) / (Temp - 1 / Temp)
End Function

' hyperbolic secant

Public Function Sech(Value As Double) As Double
    Dim Temp As Double
    Temp = EXP(Value)
    Sech = 2 / (Temp + 1 / Temp)
End Function

' hyperbolic cosecant

Public Function Csch(Value As Double) As Double
    Dim Temp As Double
    Temp = EXP(Value)
    Csch = 2 / (Temp - 1 / Temp)
End Function


' hyperbolic arc sine

Public Function Asinh(Value As Double) As Double
    Asinh = Log(Value + Sqr(Value * Value + 1))
End Function

' hyperbolic arc cosine
' error if NUMBER is inside the range [-1,1]

Public Function Acosh(Value As Double) As Double
    Acosh = Log(Value + Sqr(Value * Value - 1))
End Function

' hyperbolic arc tangent
' error if value is zero

Public Function Atanh(Value As Double) As Double
    Atanh = Log((1 / Value + 1) / (1 / Value - 1)) / 2
End Function

' hyperbolic arc cotangent

Public Function ACotH(Value As Double) As Double
    ACotH = Log((Value + 1) / (Value - 1)) / 2
End Function

' hyperbolic arc secant
' error if value is outside the range [-1,1]

Public Function Asech(Value As Double) As Double
    Asech = Log((Sqr(1 - Value * Value) + 1) / Value)
End Function

' hyperbolic arc cosecant

Public Function Acsch(Value As Double) As Double
    Acsch = Log((Sgn(Value) * Sqr(1 + Value * Value) + 1) / Value)
End Function
Public Function LogX(ByVal Number As Variant, ByVal X As Variant) As Double
Dim num As Double
Dim nBase As Double

num = CDbl(Number)
nBase = CDbl(X)
LogX = Log(num) / Log(nBase)

End Function
