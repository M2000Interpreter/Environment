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
Private Const Pi = 3.14159265358979
Private Const pihalf = 1.5707963267949
Private Const pidivby180 = 1.74532925199433E-02
Private Const xyz As Double = 1               '       X+ to X-        Z- to Z+
Private Const xzy As Double = -1              '       X+ to X-        Y+ to Y-
Private Const yzx As Double = 2               '       Y+ to y-        X+ to X-
Private Const yxz As Double = -2              '       Y+ to Y-        Z- to Z+
Private Const zxy As Double = 3               '       Z+ to Z-        Y- to Y+
Private Const zyx As Double = -3           '       Z+ to Z-        X+ to X-

Private Type VecType
    ' This is used for both 3D spatial vectors, and Euler angles as well.
    x As Double
    y As Double
    z As Double
End Type
Private Type SegType
    Origin As VecType
    AxisF  As VecType   ' Forward, West,  AxisX
    AxisL  As VecType   ' Left,    North, AxisY
    AxisU  As VecType   ' Up,      Up,    AxisZ
End Type
'
Private Type LineType
    point1 As VecType
    point2 As VecType
End Type
'
Private Type QuatType                      ' To get axis and theta back out (JPL convention):
    w As Double         ' cos(theta/2)              theta = ACos(q.w) * 2
    x As Double         ' v.x * sin(theta/2)        v.x = q.x / sin(theta/2)    solve for theta first, and then plug in here.
    y As Double         ' v.y * sin(theta/2)        v.y = q.y / sin(theta/2)
    z As Double         ' v.z * sin(theta/2)        v.z = q.z / sin(theta/2)
End Type
Public Sub Vector(a As Long, x As Double, y As Double, z As Double)
Dim v1 As VecType
v1.x = x
v1.y = y
v1.z = z
CopyMemory a, VarPtr(v1.x), 24
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
    Dim v1 As VecType, v2 As VecType
    CopyMemory VarPtr(v1), a, 24
    CopyMemory VarPtr(v2), b, 24
    DotProduct = v1.x * v2.x + v1.y * v2.y + v1.z * v2.z
End Function
Public Sub XProduct(a As Long, b As Long, c As Long)
    Dim v1 As VecType, v2 As VecType, v3 As VecType
    CopyMemory VarPtr(v1), a, 24
    CopyMemory VarPtr(v2), b, 24
    v3.x = v1.y * v2.z - v1.z * v2.y
    v3.y = v1.z * v2.x - v1.x * v2.z
    v3.z = v1.x * v2.y - v1.y * v2.x
    PutVector v3, c
End Sub
Public Sub RotVect(a As Long, b As Long, c As Long, angle As Double, Optional bDegrees As Boolean = True)  '(vec As VecType, axis As VecType, angle As Double, Optional bDegrees As Boolean = True) As VecType
    Dim vec As VecType, axis As VecType, RotVect As VecType
    CopyMemory VarPtr(vec), a, 24
    CopyMemory VarPtr(axis), b, 24
    RotVect = RotVecByQuat(vec, AxisAngle2Quat(axis, angle, bDegrees))
    CopyMemory c, VarPtr(RotVect), 24
End Sub
Public Sub RotVectMult(n As Long, a As Long, b As Long, c As Long, angle As Double, Optional bDegrees As Boolean = True)  '(vec As VecType, axis As VecType, angle As Double, Optional bDegrees As Boolean = True) As VecType
    Dim vec As VecType, axis As VecType, RotVect As VecType, m As Long
    CopyMemory VarPtr(axis), b, 24
    For m = 1 To n
    CopyMemory VarPtr(vec), a, 24
    RotVect = RotVecByQuat(vec, AxisAngle2Quat(axis, angle, bDegrees))
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

Private Function SegFrom3Pts(Origin As VecType, point1 As VecType, point2 As VecType, Optional order As Double = xyz) As SegType
    ' BodyBuilderSegment = [Origin, Point1 - Origin, Point2 - Origin, Order]
    SegFrom3Pts = SegFrom5Pts(Origin, point1, Origin, point2, Origin, order)
End Function

Private Function SegFrom5Pts(Origin As VecType, Point1Line1 As VecType, Point2Line1 As VecType, Point1Line2 As VecType, Point2Line2 As VecType, Optional order As Double = xyz) As SegType
    ' BodyBuilderSegment = [Origin, Point2Line1 - Point1Line2, Point2Line2 - Point2Line2, Order]
    '
    SegFrom5Pts = SegFromLines(Origin, MakeLine(Point1Line1, Point2Line1), MakeLine(Point1Line2, Point2Line2), order)
End Function

Private Function SegFromLines(Origin As VecType, Line1 As LineType, Line2 As LineType, Optional order As Double = xyz) As SegType
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
    Dim v1 As VecType
    Dim v2 As VecType
    '
    SegFromLines.Origin = Origin
    '
    v1 = VecDif(Line1.point1, Line1.point2)
    v2 = VecDif(Line2.point1, Line2.point2)
    v1 = UnitVec(v1)
    v2 = UnitVec(v2)
    '
    Select Case order
    Case xyz    ' 1
        SegFromLines.AxisF = v1                                                         ' X
        SegFromLines.AxisL = UnitVec(XProd(v2, SegFromLines.AxisF))                     ' Y
        SegFromLines.AxisU = UnitVec(XProd(SegFromLines.AxisF, SegFromLines.AxisL))     ' Z
    Case xzy    ' -1
        SegFromLines.AxisF = v1                                                         ' X
        SegFromLines.AxisU = UnitVec(XProd(v2, SegFromLines.AxisF))                     ' Z
        SegFromLines.AxisL = UnitVec(XProd(SegFromLines.AxisU, SegFromLines.AxisF))     ' Y
    Case yzx    ' 2
        SegFromLines.AxisL = v1                                                         ' Y
        SegFromLines.AxisU = UnitVec(XProd(v2, SegFromLines.AxisL))                     ' Z
        SegFromLines.AxisF = UnitVec(XProd(SegFromLines.AxisL, SegFromLines.AxisU))     ' X
    Case yxz    ' -2
        SegFromLines.AxisL = v1                                                         ' Y
        SegFromLines.AxisF = UnitVec(XProd(v2, SegFromLines.AxisL))                     ' X
        SegFromLines.AxisU = UnitVec(XProd(SegFromLines.AxisF, SegFromLines.AxisL))     ' Z
    Case zxy    ' 3
        SegFromLines.AxisU = v1                                                         ' Z
        SegFromLines.AxisF = UnitVec(XProd(v2, SegFromLines.AxisU))                     ' X
        SegFromLines.AxisL = UnitVec(XProd(SegFromLines.AxisU, SegFromLines.AxisF))     ' Y
    Case zyx    ' -3
        SegFromLines.AxisU = v1                                                         ' Z
        SegFromLines.AxisL = UnitVec(XProd(v2, SegFromLines.AxisU))                     ' Y
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
Private Function Euler2Quat(euler As VecType, Optional order As Double = xyz, Optional bDegrees As Boolean = True) As QuatType
    ' To get the same return from Quat2Euler, the following Euler angle constraints should be observed:
    '       pi   >  euler.x  >  -pi            ' In theory, 180 (or -180) should be possible, but not here.
    '       pi/2 >  euler.y  >   pi/2          ' Avoid pi/2 to avoid gimbal lock.
    '       pi   >  euler.z  >  -pi            ' In theory, 180 (or -180) should be possible, but not here.
    '
    Dim sX As Double, sY As Double, SZ As Double
    Dim cX As Double, cY As Double, cZ As Double
    Dim EulerTemp As VecType
    '
    EulerTemp = euler
    If bDegrees Then EulerTemp = VecDeg2Rad(EulerTemp)
    '
    ' Speed-up variables, so that Sin and Cos aren't repeatedly called.
    sX = Sin(EulerTemp.x / 2#): sY = Sin(EulerTemp.y / 2#): SZ = Sin(EulerTemp.z / 2#)
    cX = Cos(EulerTemp.x / 2#): cY = Cos(EulerTemp.y / 2#): cZ = Cos(EulerTemp.z / 2#)
    '
    Select Case order
    Case xyz
        Euler2Quat.w = cX * cY * cZ - sX * sY * SZ
        Euler2Quat.x = sX * cY * cZ + cX * sY * SZ
        Euler2Quat.y = cX * sY * cZ - sX * cY * SZ
        Euler2Quat.z = cX * cY * SZ + sX * sY * cZ
    Case xzy
        Euler2Quat.w = cX * cY * cZ + sX * sY * SZ
        Euler2Quat.x = sX * cY * cZ - cX * sY * SZ
        Euler2Quat.y = cX * cY * SZ - sX * sY * cZ
        Euler2Quat.z = cX * sY * cZ + sX * cY * SZ
    Case yzx
        Euler2Quat.w = cX * cY * cZ - sX * sY * SZ
        Euler2Quat.x = cX * cY * SZ + sX * sY * cZ
        Euler2Quat.y = sX * cY * cZ + cX * sY * SZ
        Euler2Quat.z = cX * sY * cZ - sX * cY * SZ
    Case yxz
        Euler2Quat.w = cX * cY * cZ + sX * sY * SZ
        Euler2Quat.x = cX * sY * cZ + sX * cY * SZ
        Euler2Quat.y = sX * cY * cZ - cX * sY * SZ
        Euler2Quat.z = cX * cY * SZ - sX * sY * cZ
    Case zxy
        Euler2Quat.w = cX * cY * cZ - sX * sY * SZ
        Euler2Quat.x = cX * sY * cZ - sX * cY * SZ
        Euler2Quat.y = cX * cY * SZ + sX * sY * cZ
        Euler2Quat.z = sX * cY * cZ + cX * sY * SZ
    Case zyx
        Euler2Quat.w = cX * cY * cZ + sX * sY * SZ
        Euler2Quat.x = cX * cY * SZ - sX * sY * cZ
        Euler2Quat.y = cX * sY * cZ + sX * cY * SZ
        Euler2Quat.z = sX * cY * cZ - cX * sY * SZ
    End Select
End Function

Private Function Quat2Euler(q As QuatType, Optional order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' The order (xyzEnum) may be reversed to some mathematicians.
    ' However, it's done this way so that the same xyzEnum constant can be used for Euler2Quat and Quat2Euler, and return to the same Euler angles.
    ' Also, this is the way BodyBuilder does it, and these procedures are meant to reproduce that functionality.
    '
    Select Case order
    Case xyz: Quat2Euler = Quat2EulerHelper(-2# * (q.y * q.z - q.w * q.x), q.w * q.w - q.x * q.x - q.y * q.y + q.z * q.z, 2# * (q.x * q.z + q.w * q.y), -2# * (q.x * q.y - q.w * q.z), q.w * q.w + q.x * q.x - q.y * q.y - q.z * q.z)
    Case xzy: Quat2Euler = Quat2EulerHelper(2# * (q.y * q.z + q.w * q.x), q.w * q.w - q.x * q.x + q.y * q.y - q.z * q.z, -2# * (q.x * q.y - q.w * q.z), 2# * (q.x * q.z + q.w * q.y), q.w * q.w + q.x * q.x - q.y * q.y - q.z * q.z)
    Case yzx: Quat2Euler = Quat2EulerHelper(-2# * (q.x * q.z - q.w * q.y), q.w * q.w + q.x * q.x - q.y * q.y - q.z * q.z, 2# * (q.x * q.y + q.w * q.z), -2# * (q.y * q.z - q.w * q.x), q.w * q.w - q.x * q.x + q.y * q.y - q.z * q.z)
    Case yxz: Quat2Euler = Quat2EulerHelper(2# * (q.x * q.z + q.w * q.y), q.w * q.w - q.x * q.x - q.y * q.y + q.z * q.z, -2# * (q.y * q.z - q.w * q.x), 2# * (q.x * q.y + q.w * q.z), q.w * q.w - q.x * q.x + q.y * q.y - q.z * q.z)
    Case zxy: Quat2Euler = Quat2EulerHelper(-2# * (q.x * q.y - q.w * q.z), q.w * q.w - q.x * q.x + q.y * q.y - q.z * q.z, 2# * (q.y * q.z + q.w * q.x), -2# * (q.x * q.z - q.w * q.y), q.w * q.w - q.x * q.x - q.y * q.y + q.z * q.z)
    Case zyx: Quat2Euler = Quat2EulerHelper(2# * (q.x * q.y + q.w * q.z), q.w * q.w + q.x * q.x - q.y * q.y - q.z * q.z, -2# * (q.x * q.z - q.w * q.y), 2# * (q.y * q.z + q.w * q.x), q.w * q.w - q.x * q.x - q.y * q.y + q.z * q.z)
    End Select
    '
    If bDegrees Then Quat2Euler = VecRad2Deg(Quat2Euler)
End Function

Private Function Quat2EulerHelper(r11 As Double, r12 As Double, r21 As Double, r31 As Double, r32 As Double) As VecType
    Quat2EulerHelper.x = ATan2(r11, r12)
    Quat2EulerHelper.y = ASin(r21)
    Quat2EulerHelper.z = ATan2(r31, r32)
End Function

Private Function Seg2Quat(seg As SegType) As QuatType
    ' Ignores origin.
    '
    Seg2Quat = Axes2Quat(seg.AxisF, seg.AxisL, seg.AxisU)
End Function

Private Function Quat2Seg(q As QuatType) As SegType
    ' Leaves the segment origin as zero.
    '
    Quat2Seg.AxisF = Quat2Fwd(q)
    Quat2Seg.AxisL = Quat2Left(q)
    Quat2Seg.AxisU = Quat2Up(q)
End Function

Private Function Euler2Seg(euler As VecType, Optional order As Double = xyz, Optional bDegrees As Boolean = True) As SegType
    ' Leaves the segment origin as zero.
    Euler2Seg = Quat2Seg(Euler2Quat(euler, order, bDegrees))
End Function

Private Function Seg2Euler(seg As SegType, Optional order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' Ignores origin.
    Seg2Euler = Quat2Euler(Seg2Quat(seg), order, bDegrees)
End Function

Private Function Axes2Quat(AxisF As VecType, AxisL As VecType, AxisU As VecType) As QuatType
    ' Must be three orthogonal unit vectors.
    ' Basically, the vectors that define a BodyBuilder segment.
    '
    Dim s As Double
    Dim tr As Double
    Dim Max As Double
    Dim x As Double, y As Double, z As Double
    '
    tr = (AxisF.x + AxisL.y + AxisU.z + 1#)
    '
    If (tr >= 1#) Then
        s = (0.5 / Sqr(tr))
        Axes2Quat.w = 0.25 / s
        Axes2Quat.x = (AxisL.z - AxisU.y) * s
        Axes2Quat.y = (AxisU.x - AxisF.z) * s
        Axes2Quat.z = (AxisF.y - AxisL.x) * s
    Else
        If (AxisL.y > AxisU.z) Then Max = AxisL.y Else Max = AxisU.z
        '
        If (Max < AxisF.x) Then
            s = (Sqr(AxisF.x - (AxisL.y + AxisU.z) + 1#))
            x = (s * 0.5)
            s = (0.5 / s)
            Axes2Quat.w = (AxisL.z - AxisU.y) * s
            Axes2Quat.x = x
            Axes2Quat.y = (AxisF.y + AxisL.x) * s
            Axes2Quat.z = (AxisU.x + AxisF.z) * s
        ElseIf (Max = AxisL.y) Then
            s = (Sqr(AxisL.y - (AxisU.z + AxisF.x) + 1#))
            y = (s * 0.5)
            s = (0.5 / s)
            Axes2Quat.w = (AxisU.x - AxisF.z) * s
            Axes2Quat.x = (AxisF.y + AxisL.x) * s
            Axes2Quat.y = y
            Axes2Quat.z = (AxisL.z + AxisU.y) * s
        Else
            s = (Sqr(AxisU.z - (AxisF.x + AxisL.y) + 1#))
            z = (s * 0.5)
            s = (0.5 / s)
            Axes2Quat.w = (AxisF.y - AxisL.x) * s
            Axes2Quat.x = (AxisU.x + AxisF.z) * s
            Axes2Quat.y = (AxisL.z + AxisU.y) * s
            Axes2Quat.z = z
        End If
    End If
End Function

Private Function Quat2Fwd(q As QuatType) As VecType
    ' Make sure quat is normed.
    Quat2Fwd = RotVecByQuat(MakeVec(1, 0, 0), q)
End Function

Private Function Quat2Left(q As QuatType) As VecType
    ' Make sure quat is normed.
    Quat2Left = RotVecByQuat(MakeVec(0, 1, 0), q)
End Function

Private Function Quat2Up(q As QuatType) As VecType
    ' Make sure quat is normed.
    Quat2Up = RotVecByQuat(MakeVec(0, 0, 1), q)
End Function

Private Function AxisAngle2Quat(axis As VecType, ByVal angle As Double, Optional bDegrees As Boolean = True) As QuatType
    ' You should be sure the vector is a UNIT vector before calling this, else the quaternion will not be a unit quaternion.
    ' Note that the Axis(vector) is the rotation vector, and has nothing to do with basis vectors.
    '
    Dim SinThetaDivTwo As Double
    '
    If bDegrees Then angle = Deg2Rad(angle)
    SinThetaDivTwo = Sin(angle / 2)
    '
    AxisAngle2Quat.w = Cos(angle / 2)
    AxisAngle2Quat.x = axis.x * SinThetaDivTwo
    AxisAngle2Quat.y = axis.y * SinThetaDivTwo
    AxisAngle2Quat.z = axis.z * SinThetaDivTwo
End Function

Private Function QuatAxis(quat As QuatType) As VecType
    ' This is the rotation axis of the quaternion.
    ' It has nothing to do with the quaternion's basis axes.
    Dim SinThetaByTwo As Double
    '
    SinThetaByTwo = Sin(QuatAngle(quat, False) / 2)
    '
    On Error Resume Next
        QuatAxis.x = quat.x / SinThetaByTwo
        QuatAxis.y = quat.y / SinThetaByTwo
        QuatAxis.z = quat.z / SinThetaByTwo
    On Error GoTo 0
End Function

Private Function QuatAngle(quat As QuatType, Optional bDegrees As Boolean = True) As Double
    ' Just the angle from the quaternion.
    QuatAngle = ACos(quat.w) * 2
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

Private Function RotSeg(seg As SegType, axis As VecType, angle As Double, Optional bDegrees As Boolean = True) As SegType
    ' This is exactly like a rotation: (Seg2 = Rot(Seg1, Axis, Angle)).
    '
    RotSeg = RotSegByQuat(seg, AxisAngle2Quat(axis, angle, bDegrees))
End Function

Private Function RotSegByQuat(seg As SegType, quat As QuatType) As SegType
    RotSegByQuat.AxisF = RotVecByQuat(seg.AxisF, quat)
    RotSegByQuat.AxisL = RotVecByQuat(seg.AxisL, quat)
    RotSegByQuat.AxisU = RotVecByQuat(seg.AxisU, quat)
    RotSegByQuat.Origin = seg.Origin
End Function

Private Function RotQuat(q1 As QuatType, q2 As QuatType) As QuatType
    ' q1 is being rotated by q2 amount.
    '
    RotQuat.w = q1.w * q2.w - q1.x * q2.x - q1.y * q2.y - q1.z * q2.z
    RotQuat.x = q1.w * q2.x + q1.x * q2.w - q1.y * q2.z + q1.z * q2.y
    RotQuat.y = q1.w * q2.y + q1.x * q2.z + q1.y * q2.w - q1.z * q2.x
    RotQuat.z = q1.w * q2.z - q1.x * q2.y + q1.y * q2.x + q1.z * q2.w
End Function
Private Function MulQuat(q1 As QuatType, q2 As QuatType) As QuatType
    ' q1 is being rotated by q2 amount.
    '
    MulQuat.w = q1.w * q2.w - q1.x * q2.x - q1.y * q2.y - q1.z * q2.z
    MulQuat.x = q1.w * q2.x + q1.x * q2.w + q1.y * q2.z - q1.z * q2.y
    MulQuat.y = q1.w * q2.y - q1.x * q2.z + q1.y * q2.w + q1.z * q2.x
    MulQuat.z = q1.w * q2.z + q1.x * q2.y - q1.y * q2.x + q1.z * q2.w
End Function
Private Function UnRotQuat(q1 As QuatType, q2 As QuatType) As QuatType
    ' q1 is being UNrotated by q2 amount.
    '
    UnRotQuat = RotQuat(q1, NegQuat(q2))
End Function

Private Function RotVec(vec As VecType, axis As VecType, angle As Double, Optional bDegrees As Boolean = True) As VecType
    RotVec = RotVecByQuat(vec, AxisAngle2Quat(axis, angle, bDegrees))
End Function

Private Function RotVecByQuat(vec As VecType, quat As QuatType) As VecType
    ' This is essentially a vector rotation.
    ' To emulate BodyBuilder: NewVec = RotVecByQuat(Vec, AxisAng2Quat(Axis, Deg2Rad(Angle)))
    '
    Dim num1 As Double, num2 As Double, num3 As Double
    Dim num4 As Double, num5 As Double, num6 As Double
    Dim num7 As Double, num8 As Double, num9 As Double
    Dim numa As Double, numb As Double, numc As Double
    '
    num1 = quat.x * 2#:       num2 = quat.y * 2#:       num3 = quat.z * 2#
    num4 = quat.x * num1:     num5 = quat.y * num2:     num6 = quat.z * num3
    num7 = quat.x * num2:     num8 = quat.x * num3:     num9 = quat.y * num3
    numa = quat.w * num1:     numb = quat.w * num2:     numc = quat.w * num3
    '
    ' RotVecByQuat.x = (1# - (num5 + num6)) * vec.x +       (num7 - numc)  * vec.y +       (num8 + numb)  * vec.z
    ' RotVecByQuat.y =       (num7 + numc)  * vec.x + (1# - (num4 + num6)) * vec.y +       (num9 - numa)  * vec.z
    ' RotVecByQuat.z =       (num8 - numb)  * vec.x +       (num9 + numa)  * vec.y + (1# - (num4 + num5)) * vec.z
    '
    RotVecByQuat.x = (1# - (num5 + num6)) * vec.x + (num7 - numc) * vec.y + (num8 + numb) * vec.z
    RotVecByQuat.y = (num7 + numc) * vec.x + (1# - (num4 + num6)) * vec.y + (num9 - numa) * vec.z
    RotVecByQuat.z = (num8 - numb) * vec.x + (num9 + numa) * vec.y + (1# - (num4 + num5)) * vec.z
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

Private Function FixedBetweenSegs(seg1 As SegType, seg2 As SegType, Optional order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' In BodyBuilder terms, this is:  FixedBetweenSegs = <seg1, seg2, order>
    ' This takes the "Fixed" angle approach rather than the "Euler" angle approach.
    ' Conceptually, seg1 & seg2 are reversed with respect to what's rotated against what for this "Fixed" approach,
    ' but it's left this way to correspond to BodyBuilder code.
    '
    FixedBetweenSegs = NegVec(EulerBetweenSegs(seg1, seg2, order, bDegrees))
End Function

Private Function EulerBetweenSegs(seg1 As SegType, seg2 As SegType, Optional order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' In BodyBuilder terms, this is:  EulerBetweenSegs = -<seg1, seg2, order>
    ' We do NOT need to negate (as BodyBuilder does) when using this function.
    '
    EulerBetweenSegs = EulerBetweenQuats(Seg2Quat(seg1), Seg2Quat(seg2), order, bDegrees)
End Function

Private Function EulerBetweenQuats(q1 As QuatType, q2 As QuatType, Optional order As Double = xyz, Optional bDegrees As Boolean = True) As VecType
    ' Put q2's rotation basis at <0,0,0>, and then convert to Euler angles.
    ' In other words, what Euler angles will rotate q1 to q2.
    '
    EulerBetweenQuats = Quat2Euler(UnRotQuat(q2, q1), order, bDegrees)
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
Private Function VecAvg(v1 As VecType, v2 As VecType) As VecType
    VecAvg = VecDivNum(VecSum(v1, v2), 2)
End Function
Public Sub VecAver(a As Long, b As Long, c As Long)
    Dim v1 As VecType
    v1 = VecDivNum(VecSumm1(a, b), 2)
    CopyMemory c, VarPtr(v1), 24
End Sub
Public Sub VecAverMult(n As Long, a As Long, b As Long, c As Long)
    Dim v1 As VecType, m As Long
    For m = 1 To n
    v1 = VecDivNum(VecSumm1(a, b), 2)
    CopyMemory c, VarPtr(v1), 24
    a = a + 24
    b = b + 24
    c = c + 24
    Next m
End Sub
Private Function VecSumm1(a As Long, b As Long) As VecType
    Dim v1 As VecType, v2 As VecType, v3 As VecType
    CopyMemory VarPtr(v1), a, 24
    CopyMemory VarPtr(v2), b, 24
    VecSumm1.x = v1.x + v2.x
    VecSumm1.y = v1.y + v2.y
    VecSumm1.z = v1.z + v2.z
    
End Function

Private Function VecSum(v1 As VecType, v2 As VecType) As VecType
    VecSum.x = v1.x + v2.x
    VecSum.y = v1.y + v2.y
    VecSum.z = v1.z + v2.z
End Function
Public Sub VecSumm(a As Long, b As Long, c As Long)
    Dim v1 As VecType, v2 As VecType, v3 As VecType
    CopyMemory VarPtr(v1), a, 24
    CopyMemory VarPtr(v2), b, 24
    v3.x = v1.x + v2.x
    v3.y = v1.y + v2.y
    v3.z = v1.z + v2.z
    CopyMemory c, VarPtr(v3), 24
End Sub
Public Sub VecSummMult(n As Long, a As Long, b As Long, c As Long)
    Dim v1 As VecType, v2 As VecType, v3 As VecType, m As Long
    CopyMemory VarPtr(v2), b, 24
    For m = 1 To n
    CopyMemory VarPtr(v1), a, 24
    v3.x = v1.x + v2.x
    v3.y = v1.y + v2.y
    v3.z = v1.z + v2.z
    CopyMemory c, VarPtr(v3), 24
    a = a + 24
    c = c + 24
    Next m
End Sub
Private Function VecDif(v1 As VecType, v2 As VecType) As VecType
    VecDif.x = v1.x - v2.x
    VecDif.y = v1.y - v2.y
    VecDif.z = v1.z - v2.z
End Function
Public Sub VecDiff(a As Long, b As Long, c As Long)
    Dim v1 As VecType, v2 As VecType, v3 As VecType
    CopyMemory VarPtr(v1), a, 24
    CopyMemory VarPtr(v2), b, 24
    v3.x = v1.x - v2.x
    v3.y = v1.y - v2.y
    v3.z = v1.z - v2.z
    CopyMemory c, VarPtr(v3), 24
End Sub
Public Sub VecDiffMult(n As Long, a As Long, b As Long, c As Long)
    Dim v1 As VecType, v2 As VecType, v3 As VecType, m As Long
    CopyMemory VarPtr(v2), b, 24
    For m = 1 To n
    CopyMemory VarPtr(v1), a, 24
    v3.x = v1.x - v2.x
    v3.y = v1.y - v2.y
    v3.z = v1.z - v2.z
    CopyMemory c, VarPtr(v3), 24
    a = a + 24
    c = c + 24
    Next m
End Sub


Sub VecDiffNum(a As Long, D As Double, b As Long)
    Dim v As VecType, v1 As VecType
    CopyMemory VarPtr(v1), a, 24
    v.x = v1.x - D
    v.y = v1.y - D
    v.z = v1.z - D
    CopyMemory b, VarPtr(v), 24
End Sub
Sub VecDiffNumMult(n As Long, a As Long, D As Double, b As Long)
    Dim v As VecType, v1 As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(v1), a, 24
    v.x = v1.x - D
    v.y = v1.y - D
    v.z = v1.z - D
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub

Private Function VecAddNum(v As VecType, D As Double) As VecType
    VecAddNum.x = v.x + D
    VecAddNum.y = v.y + D
    VecAddNum.z = v.z + D
End Function
Sub VecAddiNum(a As Long, D As Double, b As Long)
    Dim v As VecType, v1 As VecType
    CopyMemory VarPtr(v1), a, 24
    v.x = v1.x + D
    v.y = v1.y + D
    v.z = v1.z + D
    CopyMemory b, VarPtr(v), 24
End Sub
Sub VecAddiNumMult(n As Long, a As Long, D As Double, b As Long)
    Dim v As VecType, v1 As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(v1), a, 24
    v.x = v1.x + D
    v.y = v1.y + D
    v.z = v1.z + D
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub

Sub VecDiviNum(a As Long, D As Double, b As Long)
    Dim v As VecType, v1 As VecType
    CopyMemory VarPtr(v1), a, 24
    v.x = v1.x / D
    v.y = v1.y / D
    v.z = v1.z / D
    CopyMemory b, VarPtr(v), 24
End Sub
Sub VecDiviNumMult(n As Long, a As Long, D As Double, b As Long)
    Dim v As VecType, v1 As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(v1), a, 24
    v.x = v1.x / D
    v.y = v1.y / D
    v.z = v1.z / D
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub
Private Function VecDivNum(v1 As VecType, D As Double) As VecType
    VecDivNum.x = v1.x / D
    VecDivNum.y = v1.y / D
    VecDivNum.z = v1.z / D
End Function
Sub VecMulNum(a As Long, D As Double, b As Long)
    Dim v As VecType, v1 As VecType
    CopyMemory VarPtr(v1), a, 24
    v.x = v1.x * D
    v.y = v1.y * D
    v.z = v1.z * D
    CopyMemory b, VarPtr(v), 24
End Sub
Sub VecMulNumMult(n As Long, a As Long, D As Double, b As Long)
    Dim v As VecType, v1 As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(v1), a, 24
    v.x = v1.x * D
    v.y = v1.y * D
    v.z = v1.z * D
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub
Private Function VecMultNum(v1 As VecType, D As Double) As VecType
    VecMultNum.x = v1.x * D
    VecMultNum.y = v1.y * D
    VecMultNum.z = v1.z * D
End Function

Public Function VecMagnitude(a As Long) As Double
    Dim v As VecType
    CopyMemory VarPtr(v), a, 24
    VecMagnitude = Sqr(v.x * v.x + v.y * v.y + v.z * v.z)
End Function
Private Function VecMag(v As VecType) As Double
    VecMag = Sqr(v.x * v.x + v.y * v.y + v.z * v.z)
End Function

Private Function NegVec(v As VecType) As VecType
    NegVec.x = -v.x
    NegVec.y = -v.y
    NegVec.z = -v.z
End Function
Public Sub UnitVect(a As Long)
    ' Returns ZERO vector if vector has no magnitude.
    Dim v As VecType, d1 As VecType
    Dim D As Double
    CopyMemory VarPtr(v), a, 24
    '
    D = VecMag(v)
    If D > 0 Then
        d1.x = v.x / D
        d1.y = v.y / D
        d1.z = v.z / D
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
        d1.x = v.x / D
        d1.y = v.y / D
        d1.z = v.z / D
    Else
        d1 = Zero
    End If
    CopyMemory a, VarPtr(d1), 24
    a = a + 24
    Next m
End Sub
Private Function UnitVec(v As VecType) As VecType
    ' Returns ZERO vector if vector has no magnitude.
    Dim D As Double
    '
    D = VecMag(v)
    If D > 0 Then
        UnitVec.x = v.x / D
        UnitVec.y = v.y / D
        UnitVec.z = v.z / D
    End If
End Function
' XProduct
Private Function XProd(v1 As VecType, v2 As VecType) As VecType
    XProd.x = v1.y * v2.z - v1.z * v2.y
    XProd.y = v1.z * v2.x - v1.x * v2.z
    XProd.z = v1.x * v2.y - v1.y * v2.x
End Function
' XDotProduct
Private Function DotProd(v1 As VecType, v2 As VecType) As Double
    DotProd = v1.x * v2.x + v1.y * v2.y + v1.z * v2.z
End Function
Sub NegateQuat(a As Long, b As Long)
    Dim quat As QuatType
    CopyMemory VarPtr(quat), a, 48
'    quat.w = quat.w
    quat.x = -quat.x
    quat.y = -quat.y
    quat.z = -quat.z
    CopyMemory b, VarPtr(quat), 48
End Sub
Sub NegateQuatMult(n As Long, a As Long, b As Long)
    Dim quat As QuatType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(quat), a, 48
'    quat.w = quat.w
    quat.x = -quat.x
    quat.y = -quat.y
    quat.z = -quat.z
    CopyMemory b, VarPtr(quat), 48
    a = a + 48
    b = b + 48
    Next m
End Sub
Private Function NegQuat(quat As QuatType) As QuatType
    ' There are actually two ways to do this, negate the angle or negate the axis.
    ' It turns out that negating the axis works a bit better with the RotQuat function
    NegQuat.w = quat.w
    NegQuat.x = -quat.x
    NegQuat.y = -quat.y
    NegQuat.z = -quat.z
End Function
Private Function NegQuatAll(quat As QuatType) As QuatType
    NegQuatAll.w = -quat.w
    NegQuatAll.x = -quat.x
    NegQuatAll.y = -quat.y
    NegQuatAll.z = -quat.z
End Function
Public Sub MakeLineFrom2Vec(a As Long, b As Long, c As Long) ' Pointers VecType, VecType, LineType
    CopyMemory c, a, 24
    CopyMemory c + 24, a, 24
End Sub
Private Function MakeLine(point1 As VecType, point2 As VecType) As LineType
    MakeLine.point1 = point1
    MakeLine.point2 = point2
End Function

Private Function MakeVec(x As Double, y As Double, z As Double) As VecType
    MakeVec.x = x
    MakeVec.y = y
    MakeVec.z = z
End Function
Public Sub Quaternion(a As Long, w As Double, x As Double, y As Double, z As Double)
    CopyMemory a, VarPtr(w), 8
    CopyMemory a + 8, VarPtr(x), 8
    CopyMemory a + 16, VarPtr(y), 8
    CopyMemory a + 24, VarPtr(z), 8
End Sub
Private Function MakeQuat(w As Double, x As Double, y As Double, z As Double) As QuatType
    MakeQuat.w = w
    MakeQuat.x = x
    MakeQuat.y = y
    MakeQuat.z = z
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
    q.w = q.w * R
    q.x = q.x * R
    q.y = q.y * R
    q.z = q.z * R
    If c = 0 Then CopyMemory a, VarPtr(q), 32: Exit Sub
    CopyMemory c, VarPtr(q), 32
End Sub
Public Sub QuatMultReal(a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, R As Double
    q = GetQuatType(a)
    R = GetRealType(b)
    q.w = q.w * R
    q.x = q.x * R
    q.y = q.y * R
    q.z = q.z * R
    If c = 0 Then CopyMemory a, VarPtr(q), 32: Exit Sub
    CopyMemory c, VarPtr(q), 32
End Sub
Public Sub QuatMultRealMul(n As Long, a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, R As Double, m As Long
    For m = 1 To n
    q = GetQuatType(a)
    R = GetRealType(b)
    q.w = q.w * R
    q.x = q.x * R
    q.y = q.y * R
    q.z = q.z * R
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
    q.w = q.w * R
    q.x = q.x * R
    q.y = q.y * R
    q.z = q.z * R
    If c = 0 Then CopyMemory a, VarPtr(q), 32 Else CopyMemory c, VarPtr(q), 32: c = c + 32
    a = a + 32
    Next m
End Sub
Public Sub QuatAddByReal(a As Long, R As Double, Optional c As Long = 0)
    Dim q As QuatType
    If c = 0 Then
        q = GetQuatType(a)
        q.w = q.w + R
        CopyMemory a, VarPtr(q), 8
    Else
        q = GetQuatType(a)
        q.w = q.w + R
        CopyMemory c, VarPtr(q), 32
    End If
End Sub
Public Sub QuatAddReal(a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType
    If c = 0 Then
        q.w = GetRealType(a) + GetRealType(b)
        CopyMemory a, VarPtr(q), 8
    Else
        q = GetQuatType(a)
        q.w = q.w + GetRealType(b)
        CopyMemory c, VarPtr(q), 32
    End If
End Sub
Public Sub QuatAddRealMulConst(n As Long, a As Long, b As Long, Optional c As Long = 0)
    Dim q As QuatType, m As Long
    If c = 0 Then
        For m = 1 To n
            q.w = GetRealType(a) + GetRealType(b)
            CopyMemory a, VarPtr(q), 8
            a = a + 32
        Next m
    Else
        For m = 1 To n
            q = GetQuatType(a)
            q.w = q.w + GetRealType(b)
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
            q.w = GetRealType(a) + GetRealType(b)
            CopyMemory a, VarPtr(q), 8
            a = a + 32
            b = b + 8
        Next m
    Else
        For m = 1 To n
            q = GetQuatType(a)
            q.w = q.w + GetRealType(b)
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
        .w = .w + q2.w
        .x = .x + q2.x
        .y = .y + q2.y
        .z = .z + q2.z
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
            .w = .w + q2.w
            .x = .x + q2.x
            .y = .y + q2.y
            .z = .z + q2.z
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
            .w = .w + q2.w
            .x = .x + q2.x
            .y = .y + q2.y
            .z = .z + q2.z
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
    Dim z As Double
    z = NormQuat(GetQuatType(a))
    CopyMemory b, VarPtr(z), 8
End Sub
Public Sub QuatNormMul(n As Long, a As Long, b As Long)
    Dim z As Double, m As Long
    For m = 1 To n
    z = NormQuat(GetQuatType(a))
    CopyMemory b, VarPtr(z), 8
    a = a + 32
    b = b + 8
    Next m
End Sub
Private Function NormQuat(v As QuatType) As Double
    NormQuat = Sqr(v.w * v.w + v.x * v.x + v.y * v.y + v.z * v.z)
End Function

Public Function QuatString(a As Long, Optional frm$ = "#0.00000") As String
    Dim q As QuatType
    q = GetQuatType(a)
    QuatString = MyFrm(q.w, frm$) & ", " & MyFrm(q.x, frm$) & ", " & MyFrm(q.y, frm$) & ", " & MyFrm(q.z, frm$)
End Function
Public Function VecString(a As Long, Optional frm$ = "#0.00000") As String
    Dim v As VecType
    v = GetVector(a)
    VecString = MyFrm(v.x, frm$) & ", " & MyFrm(v.y, frm$) & ", " & MyFrm(v.z, frm$)
End Function
' trig functions
Public Sub VectorRad2Deg(a As Long, Optional b As Long = 0)
    Dim v As VecType
    v = GetVector(a)
    v.x = Rad2Deg(v.x)
    v.y = Rad2Deg(v.y)
    v.z = Rad2Deg(v.z)
    If b = 0 Then CopyMemory a, VarPtr(v), 24: Exit Sub
    CopyMemory b, VarPtr(v), 24
End Sub
Public Sub VectorRad2DegMult(n As Long, a As Long, b As Long)
    Dim v As VecType, m As Long
    For m = 1 To n
    v = GetVector(a)
    v.x = Rad2Deg(v.x)
    v.y = Rad2Deg(v.y)
    v.z = Rad2Deg(v.z)
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub

Public Sub VectorDeg2Rad(a As Long, Optional b As Long = 0)
    Dim v As VecType
    CopyMemory VarPtr(v), a, 24
    v.x = Deg2Rad(v.x)
    v.y = Deg2Rad(v.y)
    v.z = Deg2Rad(v.z)
    If b = 0 Then CopyMemory a, VarPtr(v), 24: Exit Sub
    CopyMemory b, VarPtr(v), 24
End Sub
Public Sub VectorDeg2RadMult(n As Long, a As Long, b As Long)
    Dim v As VecType, m As Long
    For m = 1 To n
    CopyMemory VarPtr(v), a, 24
    v.x = Deg2Rad(v.x)
    v.y = Deg2Rad(v.y)
    v.z = Deg2Rad(v.z)
    CopyMemory b, VarPtr(v), 24
    a = a + 24
    b = b + 24
    Next m
End Sub

Private Function VecRad2Deg(v As VecType) As VecType
    VecRad2Deg.x = Rad2Deg(v.x)
    VecRad2Deg.y = Rad2Deg(v.y)
    VecRad2Deg.z = Rad2Deg(v.z)
End Function

Private Function VecDeg2Rad(v As VecType) As VecType
    VecDeg2Rad.x = Deg2Rad(v.x)
    VecDeg2Rad.y = Deg2Rad(v.y)
    VecDeg2Rad.z = Deg2Rad(v.z)
End Function

Public Function Rad2Deg(RadAngle As Double) As Double
    Rad2Deg = RadAngle / pidivby180
End Function

Public Function Deg2Rad(DegAngle As Double) As Double
    Deg2Rad = DegAngle * pidivby180
End Function

Public Function ACos(ByVal D As Double) As Double
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
        ACos = pihalf - Atn(d2 / dRoot)
    Else
        ACos = pihalf - pihalf * Sgn(d2)
    End If
End Function

Public Function ASin(D As Double) As Double
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
        ASin = Atn(d2 / dRoot)
    Else
        ASin = pihalf * Sgn(d2)
    End If
End Function

Public Function ATan2(y As Double, x As Double) As Double
    Const epsilon As Double = 1.4E-45
    ' Radians.
    '
    Select Case x
    Case Is > epsilon:  ATan2 = Atn(y / x)
    Case Is < -epsilon: ATan2 = Atn(y / x) + Pi * Sgn(y): If y = 0# Then ATan2 = ATan2 + Pi
    Case Else:          ATan2 = pihalf * Sgn(y)
    End Select
End Function
Public Property Get CheckThis() As Boolean
CheckThis = mcheck
End Property
Public Property Let CheckThis(RHS As Boolean)
mcheck = RHS
End Property
Private Function MyFrm(w As Double, frm$) As String
If frm$ = vbNullString Then
    MyFrm = LTrim$(Str$(w))
            If Left$(MyFrm, 1) = "." Then
                    MyFrm = "0" + MyFrm
            ElseIf Left$(MyFrm, 2) = "-." Then
                    MyFrm = "-0" + Mid$(MyFrm, 2)
            End If
Else

        MyFrm = Format$(w, frm$)
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

Property Get version() As Variant
version = 1.2
End Property
Public Function ArcCos(x As Variant) As Variant

    Select Case x
        Case -1
            ArcCos = 4 * Atn(1)
             
        Case 0:
            ArcCos = 2 * Atn(1)
             
        Case 1:
            ArcCos = 0#
             
        Case Else:
            ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End Select
End Function

Public Function ArcSin(x As Variant) As Variant

    Select Case x
        Case -1
            ArcSin = 6 * Atn(1)
             
        Case 0:
            ArcSin = 0#
             
        Case 1:
            ArcSin = 2 * Atn(1)
             
        Case Else:
            ArcSin = Atn(x / Sqr(-x * x + 1))
    End Select
End Function
Public Function Cos(x As Double) As Double
        Cos = VBA.Cos(x)
End Function
Public Function Sin(x As Double) As Double
        Sin = VBA.Sin(x)
End Function
Public Function Atn(x As Double) As Double
        Atn = VBA.Atn(x)
End Function
Public Function Exp(x As Double) As Double
        Exp = VBA.Exp(x)
End Function
' arc cotangent
' error if NUMBER is zero

Public Function ACot(Value As Double) As Double
    ACot = Atn(1 / Value)
End Function

' arc secant
' error if value is inside the range [-1,1]

Public Function ASec(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ASec = ACos(1 / value)
    If Abs(Value) <> 1 Then
        ASec = 1.5707963267949 - Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ASec = 3.14159265358979 * Sgn(Value)
    End If
End Function

' arc cosecant
' error if value is inside the range [-1,1]

Public Function ACsc(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
    If Abs(Value) <> 1 Then
        ACsc = Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ACsc = 1.5707963267949 * Sgn(Value)
    End If
End Function

' hyperbolic sine

Public Function SinH(Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    SinH = (Temp - 1 / Temp) / 2
End Function

' hyperbolic cosine

Public Function CosH(Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    CosH = (Temp + 1 / Temp) / 2
End Function

' hyperbolic tangent

Public Function TanH(Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    TanH = (Temp - 1 / Temp) / (Temp + 1 / Temp)
End Function

' hyperbolic cotangent

Public Function CotH(Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    CotH = (Temp + 1 / Temp) / (Temp - 1 / Temp)
End Function

' hyperbolic secant

Public Function SecH(Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    SecH = 2 / (Temp + 1 / Temp)
End Function

' hyperbolic cosecant

Public Function CscH(Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    CscH = 2 / (Temp - 1 / Temp)
End Function


' hyperbolic arc sine

Public Function ASinH(Value As Double) As Double
    ASinH = Log(Value + Sqr(Value * Value + 1))
End Function

' hyperbolic arc cosine
' error if NUMBER is inside the range [-1,1]

Public Function ACosH(Value As Double) As Double
    ACosH = Log(Value + Sqr(Value * Value - 1))
End Function

' hyperbolic arc tangent
' error if value is zero

Public Function ATanH(Value As Double) As Double
    ATanH = Log((1 / Value + 1) / (1 / Value - 1)) / 2
End Function

' hyperbolic arc cotangent

Public Function ACotH(Value As Double) As Double
    ACotH = Log((Value + 1) / (Value - 1)) / 2
End Function

' hyperbolic arc secant
' error if value is outside the range [-1,1]

Public Function ASecH(Value As Double) As Double
    ASecH = Log((Sqr(1 - Value * Value) + 1) / Value)
End Function

' hyperbolic arc cosecant

Public Function ACscH(Value As Double) As Double
    ACscH = Log((Sgn(Value) * Sqr(1 + Value * Value) + 1) / Value)
End Function
Public Function LogX(ByVal Number As Variant, ByVal x As Variant) As Double
Dim num As Double
Dim nBase As Double

num = CDbl(Number)
nBase = CDbl(x)
LogX = Log(num) / Log(nBase)

End Function
Sub TestStr(a$)
a$ = a$ + "ok"
End Sub
Sub testStrArr(a() As Byte)
Dim b$
b$ = a()
Debug.Print b$
a(0) = 65
End Sub
Function RetStrArr() As Byte()
Dim b() As Byte
b() = "ok"

RetStrArr = b()
End Function
Sub TestValue(a)
Dim b As Long
b = a
b = b + 1
a = b
End Sub
Sub TestArr(a() As Variant)
On Error GoTo there
If UBound(a()) - LBound(a()) < 0 Then Exit Sub
a(LBound(a)) = 100
ReDim Preserve a(20 - 1)
there:
End Sub
Sub TestObj(a)
Dim z As mArray, m As mArray
If TypeOf a Is mArray Then
Set m = a
Set z = New mArray
m.CopyArray z
z.Compute 1, "++"
Set a = z
Else
Set m = a.objref
Set z = New mArray
m.CopyArray z
z.Compute 1, "++"
Set a.objref = z
End If
End Sub
Public Function TestOpt(Optional a As Long = 100, Optional b As Long = 1) As Long
    TestOpt = a * b
End Function
