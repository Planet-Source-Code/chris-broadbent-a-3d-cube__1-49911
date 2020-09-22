VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Cube by Chris Broadbent"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   214
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   503
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPoints 
      Caption         =   "8"
      Height          =   225
      Index           =   7
      Left            =   6930
      TabIndex        =   15
      Top             =   2730
      Width           =   435
   End
   Begin VB.CheckBox chkPoints 
      Caption         =   "7"
      Height          =   225
      Index           =   6
      Left            =   6405
      TabIndex        =   14
      Top             =   2730
      Width           =   435
   End
   Begin VB.CheckBox chkPoints 
      Caption         =   "6"
      Height          =   225
      Index           =   5
      Left            =   5880
      TabIndex        =   13
      Top             =   2730
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkPoints 
      Caption         =   "5"
      Height          =   225
      Index           =   4
      Left            =   5355
      TabIndex        =   12
      Top             =   2730
      Width           =   435
   End
   Begin VB.CheckBox chkPoints 
      Caption         =   "4"
      Height          =   225
      Index           =   3
      Left            =   4830
      TabIndex        =   11
      Top             =   2730
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkPoints 
      Caption         =   "3"
      Height          =   225
      Index           =   2
      Left            =   4305
      TabIndex        =   10
      Top             =   2730
      Width           =   435
   End
   Begin VB.CheckBox chkPoints 
      Caption         =   "2"
      Height          =   225
      Index           =   1
      Left            =   3780
      TabIndex        =   9
      Top             =   2730
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   750
      Left            =   3150
      TabIndex        =   20
      Top             =   2415
      Width           =   4320
      Begin VB.CheckBox chkPoints 
         Caption         =   "1"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   315
         Value           =   1  'Checked
         Width           =   435
      End
      Begin VB.CheckBox chkDot 
         Caption         =   "Hightlight Points"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   1485
      End
   End
   Begin VB.HScrollBar ZScroll 
      Height          =   330
      LargeChange     =   50
      Left            =   3255
      Max             =   360
      TabIndex        =   5
      Top             =   1890
      Width           =   3165
   End
   Begin VB.HScrollBar XScroll 
      Height          =   330
      LargeChange     =   50
      Left            =   3255
      Max             =   360
      TabIndex        =   3
      Top             =   1260
      Width           =   3165
   End
   Begin VB.HScrollBar YScroll 
      Height          =   330
      LargeChange     =   50
      Left            =   3255
      Max             =   360
      TabIndex        =   1
      Top             =   630
      Width           =   3165
   End
   Begin VB.PictureBox picView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   3000
      Left            =   105
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   105
      Width           =   3000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rotation Controls"
      Height          =   2220
      Left            =   3150
      TabIndex        =   16
      Top             =   105
      Width           =   4320
      Begin VB.CommandButton cmdZSPIN 
         Caption         =   "Spin"
         Height          =   330
         Left            =   3360
         TabIndex        =   6
         Top             =   1785
         Width           =   855
      End
      Begin VB.CommandButton cmdXSPIN 
         Caption         =   "Spin"
         Height          =   330
         Left            =   3360
         TabIndex        =   4
         Top             =   1155
         Width           =   855
      End
      Begin VB.CommandButton cmdYSPIN 
         Caption         =   "Spin"
         Height          =   330
         Left            =   3360
         TabIndex        =   2
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Z Rotation"
         Height          =   225
         Left            =   105
         TabIndex        =   19
         Top             =   1575
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "X Rotation"
         Height          =   225
         Left            =   105
         TabIndex        =   18
         Top             =   945
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Y Rotation"
         Height          =   330
         Left            =   105
         TabIndex        =   17
         Top             =   315
         Width           =   3060
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/================================\
'| A 3D Cube, my first ever 3D    |
'| Program, by Chris Broadbent    |
'|                                |
'| Use portions of this code, but |
'| PLEASE refer to the website    |
'| below , where i learnt 3D      |
'|                                |
'|http://www.geocities.com/       |
'|SiliconValley/2151/math3d.html  |
'\================================/


Private Type point3D    'Specifies a point in 3d space
    x As Long
    y As Long
    z As Long
End Type

Private Type Point2D    'Specifies a point in 2d space
    x As Long
    y As Long
End Type

Private Const Pi = 3.14159265 ' the value of pi acording to my calculator

Dim CubePoints(7) As point3D    'the original cube points, used to transfer
Dim NewCubePoints(7) As point3D 'the moved cube points

Private MousePoint As Point2D   'mouse point used for draging

Private dothings  As Boolean    'should physics and draw do stuff?
Private Running As Boolean      'used for the main loop below

'values for deciding which axis should spin
Private XSPIN As Boolean, YSPIN As Boolean, ZSPIN As Boolean

'AUTO SPIN TRIGGERS

Private Sub cmdYSPIN_Click()
    YSPIN = Not YSPIN   'Trigger the YSPIN
End Sub
Private Sub cmdXSPIN_Click()
    XSPIN = Not XSPIN   'Trigger xspin
End Sub
Private Sub cmdZSPIN_Click()
    ZSPIN = Not ZSPIN   'trigger zspin
End Sub

Private Sub Form_Load()

    'sets the 8 cube points
    
    CubePoints(0).x = 10
    CubePoints(0).y = 10
    CubePoints(0).z = 10
    CubePoints(1).x = -10
    CubePoints(1).y = 10
    CubePoints(1).z = 10
    CubePoints(2).x = 10
    CubePoints(2).y = -10
    CubePoints(2).z = 10
    CubePoints(3).x = 10
    CubePoints(3).y = 10
    CubePoints(3).z = -10
    CubePoints(4).x = -10
    CubePoints(4).y = -10
    CubePoints(4).z = 10
    CubePoints(5).x = -10
    CubePoints(5).y = 10
    CubePoints(5).z = -10
    CubePoints(6).x = 10
    CubePoints(6).y = -10
    CubePoints(6).z = -10
    CubePoints(7).x = -10
    CubePoints(7).y = -10
    CubePoints(7).z = -10
    
    'set various variables to true
    dothings = True
    Running = True
    
    'shows the form
    Me.Show
    
    'just a simple loop, see the relevent subs to find out
    'whats happening
    Do While Running
        DoSpin
        Wait
    Loop
    
    End
End Sub

Private Sub Form_Paint()

    'If the form is told to paint, make the picture box draw

    Physics
    Draw
End Sub



Private Sub Physics()

Dim YCameraAngle As Single      'Camera angles in radians
Dim XCameraAngle As Single
Dim ZCameraAngle As Single

    If Not dothings Then Exit Sub

    'changes the camera angles into radians ny the formula
    'Radains = (Degrees * Pi) / 180
    
    'y camera angle is inverted to fix up thee scrollbar values, making it look more realistic
    YCameraAngle = -(YScroll.Value * Pi / 180)
    XCameraAngle = XScroll.Value * Pi / 180
    ZCameraAngle = ZScroll.Value * Pi / 180
    
    
    Dim i As Integer
    
    
    For i = 0 To 7
        'THe following lines rotate the points around axe
        'using the below functions
        NewCubePoints(i) = RotateX(CubePoints(i), XCameraAngle)
        NewCubePoints(i) = RotateY(NewCubePoints(i), YCameraAngle)
        NewCubePoints(i) = RotateZ(NewCubePoints(i), ZCameraAngle)
    Next i
    
    On Error Resume Next
    
    'SEts the zoom and perspective of all the points
    For i = 0 To 7
        'Translating the point back 50 gives a good view
        'and stop -Z's, which cause errors (try it without this line)
        NewCubePoints(i) = TranslatePoint(NewCubePoints(i), , , 50)
        With NewCubePoints(i)
            'set the zoom (256) and perspective, 100 specifies the centre of the screen
            '.x/.y causes perspective, making things further back move slower(as in real life)
            .x = (.x / .z) * 256 + 100
            .y = (.y / .z) * 256 + 100
        End With
    Next i

End Sub

Private Sub Draw()

    If Not dothings Then Exit Sub
    
    Dim i As Integer, l As Integer
    picView.Cls
    
    'Draws all the lines on the cube, took a while to
    'figure out all the point-point's
    
    'i is point number 1
    'l is point number 2
    
    'line from point1 to 2
    i = 0
    l = 1
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 0
    l = 2
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 0
    l = 3
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 1
    l = 5
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 1
    l = 4
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 5
    l = 7
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 5
    l = 3
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 3
    l = 6
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 0
    l = 1
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 2
    l = 4
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 2
    l = 6
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 7
    l = 4
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    i = 7
    l = 6
    picView.Line (NewCubePoints(i).x, NewCubePoints(i).y)-(NewCubePoints(l).x, NewCubePoints(l).y)
    
    'cycles through all points and draws a circle if wanted
    If chkDot.Value = 1 Then
        For i = 0 To 7
            If chkPoints(i).Value = 1 Then picView.Circle (NewCubePoints(i).x, NewCubePoints(i).y), 5, vbRed
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'stops the rendering loop
    Running = False
End Sub


'===============SECTION SCROLL==================='
'this section is just the user input on the scroll bars
Private Sub YScroll_Change()
    Physics
    Draw
End Sub

Private Sub YScroll_Scroll()
    Physics
    Draw
End Sub
Private Sub ZScroll_Change()
    Physics
    Draw
End Sub

Private Sub ZScroll_Scroll()
    Physics
    Draw
End Sub
Private Sub XScroll_Change()
    Physics
    Draw
End Sub

Private Sub XScroll_Scroll()
    Physics
    Draw
End Sub

'==========END SECTION=========

Private Sub DoSpin()

    'Basically Spins all the hscrollbars +1, then draws

    dothings = False

    If YSPIN Then
        If YScroll.Value >= 359 Then YScroll.Value = 0
        YScroll.Value = YScroll.Value + 1
    End If
    
    If XSPIN Then
        If XScroll.Value >= 359 Then XScroll.Value = 0
        XScroll.Value = XScroll.Value + 1
    End If
    
    If ZSPIN Then
        If ZScroll.Value >= 359 Then ZScroll.Value = 0
        ZScroll.Value = ZScroll.Value + 1
    End If
    
    dothings = True

    Physics
    Draw
End Sub

'============== TRIG SECTION ===========

'rotates around Z (depth axis)
'causes roll
Private Function RotateZ(Point As point3D, Angle As Single) As point3D
    RotateZ.x = (Cos(Angle) * Point.x) - (Sin(Angle) * Point.y)
    RotateZ.y = (Sin(Angle) * Point.x) + (Cos(Angle) * Point.y)
    RotateZ.z = Point.z
End Function

'X axis causes pitch
Private Function RotateX(Point As point3D, Angle As Single) As point3D
    RotateX.x = Point.x
    RotateX.y = (Cos(Angle) * Point.y) - (Sin(Angle) * Point.z)
    RotateX.z = (Sin(Angle) * Point.y) + (Cos(Angle) * Point.z)

End Function

'Y axis causes yaw
Private Function RotateY(Point As point3D, Angle As Single) As point3D
    RotateY.x = (Cos(Angle) * Point.x) + (Sin(Angle) * Point.z)
    RotateY.y = Point.y
    RotateY.z = -(Sin(Angle) * Point.x) + (Cos(Angle) * Point.z)

End Function

'============END SECTION=============

'moves a point in 3d space

'if you don't understand, don't do 3d!
Private Function TranslatePoint(Point As point3D, Optional x As Long, Optional y As Long, Optional z As Long) As point3D
    TranslatePoint.x = Point.x + x
    TranslatePoint.y = Point.y + y
    TranslatePoint.z = Point.z + z
End Function

'causes the code to freez
'this is used to give the computer time to draw the lines

Private Sub Wait()
    Dim l As Long
    
    For l = 0 To 500
        DoEvents
    Next l
End Sub

Private Sub picView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This set the offset for click and draging on the image
    MousePoint.x = x - YScroll.Value
    MousePoint.y = y - XScroll.Value
End Sub

Private Sub picView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Rotates the cube by mouse movement

'i don't really understand how this works (i just used
'the old method of guess and check
Dim temp

    If Button = vbLeftButton Then
        'move x axis
        temp = x - MousePoint.x
        Do Until temp >= 0 And temp <= 360
            If temp >= 360 Then temp = temp - 360
            If temp < 0 Then temp = temp + 360
        Loop
        YScroll.Value = temp
        
        'move y axis
        temp = y - MousePoint.y
        Do Until temp >= 0 And temp <= 360
            If temp >= 360 Then temp = temp - 360
            If temp < 0 Then temp = temp + 360
        Loop
        XScroll.Value = temp
    End If
    

End Sub
