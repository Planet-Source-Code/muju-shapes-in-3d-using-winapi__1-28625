Attribute VB_Name = "MyModule"
Option Explicit

'-- Constants for drawing functions
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const BLACKNESS = &H42
Public Const WHITENESS = &HFF0062

'-- Constants for Pens and Brush functions
Public Const PS_SOLID = 0
Public Const PS_DASH = 1
Public Const PS_DOT = 2
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6

'-- Other Contants
Public Const Pi = 3.14159265358979 / 180

'-- Types for drawing functions
Public Type POINTAPI
        X As Long
        Y As Long
End Type

'-- My own Type for the 3D objects information
Public Type My3DPosXYZType
    X As Long
    Y As Long
    Z As Long
End Type

Public Type My3DInfoType
    PosX As Long
    PosY As Long
    PosZ As Long
    TurnLR As Long
    TurnUD As Long
    TurnTU As Long
    MyPoints() As POINTAPI              '-- For storage after 2D convertion
    My3DPoints() As My3DPosXYZType      '-- Original 3D space Coordinates
    My3DCoordinates() As My3DPosXYZType '-- 3D space Coordinates after Rotations and Transformations
    DrawOrder() As Long                 '-- Order of which the 3D Panes will be drawn
End Type

'-- Device Context functions
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'-- Drawing functions
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

'-- Pens and Brushes
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'-- Other functions
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'-- Variables used for Device Context and Drawing functions
Public ScreenDC As Long        '--To store the handle to the Screen Device Context
Public BackBuffer As Long      '--To store the HDC (Handle to Device Context) to a Memory Device Context to be used as a BackBuffer
Public BackBitmap As Long      '--To store the Drawing Space Object Handle for the HDC that is used as the BackBuffer

'-- Variables to store Pen and Brush handles
Public MyPens() As Long
Public MyBrushes() As Long

'-- Variables For 3D Objects
Public ShapeA As My3DInfoType   '--3D Object = Cube
Public ShapeB As My3DInfoType   '--3D Object = Pyramid
Public ShapeC As My3DInfoType   '--3D Object = Cylinder
Public ShapeD As My3DInfoType   '--3D Object = Cone
Public ShapeE As My3DInfoType   '--3D Object = Sphere -- kinda useless but have to use in order not to violate the 3D Engine and Atomosphere

'-- Temporary Variables for miscellaneous use
Public Aa As Long
Public Ab As Long
Public Ac As Long
Public Ad As Long
Public Ax As Long
Public Ay As Long
Public TempDrawOrder() As Long
Public TempArray() As Long
Public TempPoints As My3DPosXYZType
Public Delay

Sub Main()
    
    InitializeDeviceContext
    'Testing
    Initialize3DObjects
    CreatePensBrushes
    
    Do
        'For Delay = 0 To 100000: Next Delay
        
        Ad = Ad + 1
        Ad = Ad Mod 360
        
        ShapeA.PosX = Sin((Ad) * Pi) * 400
        ShapeA.PosY = 0
        ShapeA.PosZ = Cos((Ad) * Pi) * 2000 + 3000
        ShapeA.TurnUD = ShapeA.TurnUD + 1 'Int(Rnd * 11) - 5
        ShapeA.TurnLR = ShapeA.TurnLR + 2 'Int(Rnd * 11) - 5
        ShapeA.TurnTU = ShapeA.TurnTU + 0 'Int(Rnd * 11) - 5
        'DrawShapeA
        
        ShapeB.PosX = Sin((Ad + 72) * Pi) * 400
        ShapeB.PosY = 0
        ShapeB.PosZ = Cos((Ad + 72) * Pi) * 2000 + 3000
        ShapeB.TurnUD = ShapeB.TurnUD + 0 'Int(Rnd * 11) - 5
        ShapeB.TurnLR = ShapeB.TurnLR + 2 'Int(Rnd * 11) - 5
        ShapeB.TurnTU = ShapeB.TurnTU + 1 'Int(Rnd * 11) - 5
        'DrawShapeB
        
        ShapeC.PosX = Sin((Ad + 144) * Pi) * 400
        ShapeC.PosY = 0
        ShapeC.PosZ = Cos((Ad + 144) * Pi) * 2000 + 3000
        ShapeC.TurnUD = ShapeC.TurnUD + 3
        ShapeC.TurnLR = ShapeC.TurnLR + 2
        ShapeC.TurnTU = ShapeC.TurnTU + 1
        'DrawShapeC
        
        ShapeD.PosX = Sin((Ad + 216) * Pi) * 400
        ShapeD.PosY = 0
        ShapeD.PosZ = Cos((Ad + 216) * Pi) * 2000 + 3000
        ShapeD.TurnUD = ShapeD.TurnUD + 1
        ShapeD.TurnLR = ShapeD.TurnLR + 2
        ShapeD.TurnTU = ShapeD.TurnTU + 1
        'DrawShapeD
        
        ShapeE.PosX = Sin((Ad + 288) * Pi) * 400
        ShapeE.PosY = 0
        ShapeE.PosZ = Cos((Ad + 288) * Pi) * 2000 + 3000
        'DrawShapeE
        
        DrawAllShapes
        
        Ax = ((Screen.Width / Screen.TwipsPerPixelX) - 320) / 2
        Ay = ((Screen.Height / Screen.TwipsPerPixelY) - 240) / 2
        BitBlt ScreenDC, Ax, Ay, 320, 240, BackBuffer, 0, 0, SRCCOPY
        BitBlt BackBuffer, 0, 0, 320, 240, BackBuffer, 0, 0, WHITENESS
        DoEvents
    Loop Until GetAsyncKeyState(27) < -1
    
    DeletePensBrushes
    ReleaseDeviceContext
End Sub

'-- This Sub Initializes all the Memory Device Context
Sub InitializeDeviceContext()
'----Retrieving the handle to Screen Device Context
    ScreenDC = GetDC(0)

'----Creating Memory BackBuffer compatible to the current Screen mode
    BackBuffer = CreateCompatibleDC(ScreenDC)
    BackBitmap = CreateCompatibleBitmap(ScreenDC, 320, 240)
    DeleteObject SelectObject(BackBuffer, BackBitmap)

End Sub

'-- This Sub Unloads all the Memory Device Context
Sub ReleaseDeviceContext()
'----Releasing the handle to Screen Device Context
    ReleaseDC 0, ScreenDC

'----Flush the Memory BackBuffer
    DeleteDC BackBuffer
    DeleteObject BackBitmap

End Sub

'-- This Sub Sets all the neccesary variables for the 3D animation
Sub Initialize3DObjects()
'----Setting All the Variables
'----ShapeA - Cube
    ShapeA.PosX = 100
    ShapeA.PosY = 100
    ShapeA.PosZ = 100
    ShapeA.TurnLR = 0
    ShapeA.TurnUD = 0
    ShapeA.TurnTU = 0
    ReDim ShapeA.MyPoints(23)
    ReDim ShapeA.My3DPoints(23)
    ReDim ShapeA.My3DCoordinates(23)
    ReDim ShapeA.DrawOrder(5)
    
    ShapeA.My3DPoints(0).X = -100
    ShapeA.My3DPoints(0).Y = -100
    ShapeA.My3DPoints(0).Z = -100
    ShapeA.My3DPoints(1).X = 100
    ShapeA.My3DPoints(1).Y = -100
    ShapeA.My3DPoints(1).Z = -100
    ShapeA.My3DPoints(2).X = 100
    ShapeA.My3DPoints(2).Y = 100
    ShapeA.My3DPoints(2).Z = -100
    ShapeA.My3DPoints(3).X = -100
    ShapeA.My3DPoints(3).Y = 100
    ShapeA.My3DPoints(3).Z = -100
    
    ShapeA.My3DPoints(4).X = -100
    ShapeA.My3DPoints(4).Y = -100
    ShapeA.My3DPoints(4).Z = 100
    ShapeA.My3DPoints(5).X = 100
    ShapeA.My3DPoints(5).Y = -100
    ShapeA.My3DPoints(5).Z = 100
    ShapeA.My3DPoints(6).X = 100
    ShapeA.My3DPoints(6).Y = 100
    ShapeA.My3DPoints(6).Z = 100
    ShapeA.My3DPoints(7).X = -100
    ShapeA.My3DPoints(7).Y = 100
    ShapeA.My3DPoints(7).Z = 100
    
    ShapeA.My3DPoints(8).X = -100
    ShapeA.My3DPoints(8).Y = -100
    ShapeA.My3DPoints(8).Z = -100
    ShapeA.My3DPoints(9).X = -100
    ShapeA.My3DPoints(9).Y = 100
    ShapeA.My3DPoints(9).Z = -100
    ShapeA.My3DPoints(10).X = -100
    ShapeA.My3DPoints(10).Y = 100
    ShapeA.My3DPoints(10).Z = 100
    ShapeA.My3DPoints(11).X = -100
    ShapeA.My3DPoints(11).Y = -100
    ShapeA.My3DPoints(11).Z = 100
    
    ShapeA.My3DPoints(12).X = 100
    ShapeA.My3DPoints(12).Y = -100
    ShapeA.My3DPoints(12).Z = -100
    ShapeA.My3DPoints(13).X = 100
    ShapeA.My3DPoints(13).Y = 100
    ShapeA.My3DPoints(13).Z = -100
    ShapeA.My3DPoints(14).X = 100
    ShapeA.My3DPoints(14).Y = 100
    ShapeA.My3DPoints(14).Z = 100
    ShapeA.My3DPoints(15).X = 100
    ShapeA.My3DPoints(15).Y = -100
    ShapeA.My3DPoints(15).Z = 100
    
    ShapeA.My3DPoints(16).X = -100
    ShapeA.My3DPoints(16).Y = -100
    ShapeA.My3DPoints(16).Z = -100
    ShapeA.My3DPoints(17).X = 100
    ShapeA.My3DPoints(17).Y = -100
    ShapeA.My3DPoints(17).Z = -100
    ShapeA.My3DPoints(18).X = 100
    ShapeA.My3DPoints(18).Y = -100
    ShapeA.My3DPoints(18).Z = 100
    ShapeA.My3DPoints(19).X = -100
    ShapeA.My3DPoints(19).Y = -100
    ShapeA.My3DPoints(19).Z = 100
    
    ShapeA.My3DPoints(20).X = -100
    ShapeA.My3DPoints(20).Y = 100
    ShapeA.My3DPoints(20).Z = -100
    ShapeA.My3DPoints(21).X = 100
    ShapeA.My3DPoints(21).Y = 100
    ShapeA.My3DPoints(21).Z = -100
    ShapeA.My3DPoints(22).X = 100
    ShapeA.My3DPoints(22).Y = 100
    ShapeA.My3DPoints(22).Z = 100
    ShapeA.My3DPoints(23).X = -100
    ShapeA.My3DPoints(23).Y = 100
    ShapeA.My3DPoints(23).Z = 100
    
'----ShapeB - Pyramid
    ShapeB.PosX = 100
    ShapeB.PosY = 100
    ShapeB.PosZ = 100
    ShapeB.TurnLR = 0
    ShapeB.TurnUD = 0
    ShapeB.TurnTU = 0
    ReDim ShapeB.MyPoints(15)
    ReDim ShapeB.My3DPoints(15)
    ReDim ShapeB.My3DCoordinates(15)
    ReDim ShapeB.DrawOrder(4)
    
    ShapeB.My3DPoints(0).X = 100
    ShapeB.My3DPoints(0).Y = -100
    ShapeB.My3DPoints(0).Z = -100
    ShapeB.My3DPoints(1).X = 100
    ShapeB.My3DPoints(1).Y = -100
    ShapeB.My3DPoints(1).Z = 100
    ShapeB.My3DPoints(2).X = 0
    ShapeB.My3DPoints(2).Y = 200
    ShapeB.My3DPoints(2).Z = 0
    
    ShapeB.My3DPoints(3).X = -100
    ShapeB.My3DPoints(3).Y = -100
    ShapeB.My3DPoints(3).Z = -100
    ShapeB.My3DPoints(4).X = -100
    ShapeB.My3DPoints(4).Y = -100
    ShapeB.My3DPoints(4).Z = 100
    ShapeB.My3DPoints(5).X = 0
    ShapeB.My3DPoints(5).Y = 200
    ShapeB.My3DPoints(5).Z = 0
    
    ShapeB.My3DPoints(6).X = 100
    ShapeB.My3DPoints(6).Y = -100
    ShapeB.My3DPoints(6).Z = 100
    ShapeB.My3DPoints(7).X = -100
    ShapeB.My3DPoints(7).Y = -100
    ShapeB.My3DPoints(7).Z = 100
    ShapeB.My3DPoints(8).X = 0
    ShapeB.My3DPoints(8).Y = 200
    ShapeB.My3DPoints(8).Z = 0
    
    ShapeB.My3DPoints(9).X = -100
    ShapeB.My3DPoints(9).Y = -100
    ShapeB.My3DPoints(9).Z = -100
    ShapeB.My3DPoints(10).X = 100
    ShapeB.My3DPoints(10).Y = -100
    ShapeB.My3DPoints(10).Z = -100
    ShapeB.My3DPoints(11).X = 0
    ShapeB.My3DPoints(11).Y = 200
    ShapeB.My3DPoints(11).Z = 0
    
    ShapeB.My3DPoints(12).X = -100
    ShapeB.My3DPoints(12).Y = -100
    ShapeB.My3DPoints(12).Z = -100
    ShapeB.My3DPoints(13).X = 100
    ShapeB.My3DPoints(13).Y = -100
    ShapeB.My3DPoints(13).Z = -100
    ShapeB.My3DPoints(14).X = 100
    ShapeB.My3DPoints(14).Y = -100
    ShapeB.My3DPoints(14).Z = 100
    ShapeB.My3DPoints(15).X = -100
    ShapeB.My3DPoints(15).Y = -100
    ShapeB.My3DPoints(15).Z = 100
    
'----ShapeC - Cylinder
    ShapeC.PosX = 100
    ShapeC.PosY = 100
    ShapeC.PosZ = 100
    ShapeC.TurnLR = 0
    ShapeC.TurnUD = 0
    ShapeC.TurnTU = 0
    ReDim ShapeC.MyPoints(107)
    ReDim ShapeC.My3DPoints(107)
    ReDim ShapeC.My3DCoordinates(107)
    ReDim ShapeC.DrawOrder(19)
    For Aa = 0 To 17
        ShapeC.My3DPoints((Aa * 4)).X = Sin(Aa / 18 * 360 * Pi) * 100
        ShapeC.My3DPoints((Aa * 4)).Y = -150
        ShapeC.My3DPoints((Aa * 4)).Z = Cos(Aa / 18 * 360 * Pi) * 100
        ShapeC.My3DPoints((Aa * 4) + 1).X = Sin(Aa / 18 * 360 * Pi) * 100
        ShapeC.My3DPoints((Aa * 4) + 1).Y = 150
        ShapeC.My3DPoints((Aa * 4) + 1).Z = Cos(Aa / 18 * 360 * Pi) * 100
        ShapeC.My3DPoints((Aa * 4) + 2).X = Sin(((Aa + 1) Mod 18) / 18 * 360 * Pi) * 100
        ShapeC.My3DPoints((Aa * 4) + 2).Y = 150
        ShapeC.My3DPoints((Aa * 4) + 2).Z = Cos(((Aa + 1) Mod 18) / 18 * 360 * Pi) * 100
        ShapeC.My3DPoints((Aa * 4) + 3).X = Sin(((Aa + 1) Mod 18) / 18 * 360 * Pi) * 100
        ShapeC.My3DPoints((Aa * 4) + 3).Y = -150
        ShapeC.My3DPoints((Aa * 4) + 3).Z = Cos(((Aa + 1) Mod 18) / 18 * 360 * Pi) * 100
    Next Aa
    For Aa = 0 To 17
        ShapeC.My3DPoints(Aa + 72).X = Sin(Aa / 18 * 360 * Pi) * 100
        ShapeC.My3DPoints(Aa + 72).Y = 150
        ShapeC.My3DPoints(Aa + 72).Z = Cos(Aa / 18 * 360 * Pi) * 100
    Next Aa
    For Aa = 0 To 17
        ShapeC.My3DPoints(Aa + 90).X = Sin(Aa / 18 * 360 * Pi) * 100
        ShapeC.My3DPoints(Aa + 90).Y = -150
        ShapeC.My3DPoints(Aa + 90).Z = Cos(Aa / 18 * 360 * Pi) * 100
    Next Aa

'----ShapeD - Cone
    ShapeD.PosX = 100
    ShapeD.PosY = 100
    ShapeD.PosZ = 100
    ShapeD.TurnLR = 0
    ShapeD.TurnUD = 0
    ShapeD.TurnTU = 0
    ReDim ShapeD.MyPoints(107)
    ReDim ShapeD.My3DPoints(107)
    ReDim ShapeD.My3DCoordinates(107)
    ReDim ShapeD.DrawOrder(18)
    For Aa = 0 To 17
        ShapeD.My3DPoints((Aa * 3)).X = Sin(Aa / 18 * 360 * Pi) * 100
        ShapeD.My3DPoints((Aa * 3)).Y = 150
        ShapeD.My3DPoints((Aa * 3)).Z = Cos(Aa / 18 * 360 * Pi) * 100
        ShapeD.My3DPoints((Aa * 3) + 1).X = Sin(((Aa + 1) Mod 18) / 18 * 360 * Pi) * 100
        ShapeD.My3DPoints((Aa * 3) + 1).Y = 150
        ShapeD.My3DPoints((Aa * 3) + 1).Z = Cos(((Aa + 1) Mod 18) / 18 * 360 * Pi) * 100
        ShapeD.My3DPoints((Aa * 3) + 2).X = 0
        ShapeD.My3DPoints((Aa * 3) + 2).Y = -150
        ShapeD.My3DPoints((Aa * 3) + 2).Z = 0
    Next Aa
    For Aa = 0 To 17
        ShapeD.My3DPoints(Aa + 54).X = Sin(Aa / 18 * 360 * Pi) * 100
        ShapeD.My3DPoints(Aa + 54).Y = 150
        ShapeD.My3DPoints(Aa + 54).Z = Cos(Aa / 18 * 360 * Pi) * 100
    Next Aa

'----ShapeE - Sphere
    ShapeD.PosX = 100
    ShapeD.PosY = 100
    ShapeD.PosZ = 100

End Sub

'-- Call this Sub to Draw the ShapeA - Cube
Sub DrawShapeA()
    On Error Resume Next
       
        '--Adjusting the rotation variables
        ShapeA.TurnUD = ShapeA.TurnUD Mod 360
        ShapeA.TurnLR = ShapeA.TurnLR Mod 360
        ShapeA.TurnTU = ShapeA.TurnTU Mod 360
        
        '-- Calculation Of 3D Coordinates
        For Aa = 0 To 23
            '-- Set values to temporary variables for adjustment before drawing
            ShapeA.My3DCoordinates(Aa).X = ShapeA.My3DPoints(Aa).X
            ShapeA.My3DCoordinates(Aa).Y = ShapeA.My3DPoints(Aa).Y
            ShapeA.My3DCoordinates(Aa).Z = ShapeA.My3DPoints(Aa).Z

            '--Rotation
            TempPoints.X = (Cos(ShapeA.TurnTU * Pi) * ShapeA.My3DCoordinates(Aa).X) + (-Sin(ShapeA.TurnTU * Pi) * ShapeA.My3DCoordinates(Aa).Y)
            ShapeA.My3DCoordinates(Aa).Y = (Sin(ShapeA.TurnTU * Pi) * ShapeA.My3DCoordinates(Aa).X) + (Cos(ShapeA.TurnTU * Pi) * ShapeA.My3DCoordinates(Aa).Y)
            ShapeA.My3DCoordinates(Aa).X = TempPoints.X
        
            TempPoints.X = (Cos(ShapeA.TurnUD * Pi) * ShapeA.My3DCoordinates(Aa).X) + (-Sin(ShapeA.TurnUD * Pi) * ShapeA.My3DCoordinates(Aa).Z)
            ShapeA.My3DCoordinates(Aa).Z = (Sin(ShapeA.TurnUD * Pi) * ShapeA.My3DCoordinates(Aa).X) + (Cos(ShapeA.TurnUD * Pi) * ShapeA.My3DCoordinates(Aa).Z)
            ShapeA.My3DCoordinates(Aa).X = TempPoints.X
        
            TempPoints.Y = (Cos(ShapeA.TurnLR * Pi) * ShapeA.My3DCoordinates(Aa).Y) + (-Sin(ShapeA.TurnLR * Pi) * ShapeA.My3DCoordinates(Aa).Z)
            ShapeA.My3DCoordinates(Aa).Z = (Sin(ShapeA.TurnLR * Pi) * ShapeA.My3DCoordinates(Aa).Y) + (Cos(ShapeA.TurnLR * Pi) * ShapeA.My3DCoordinates(Aa).Z)
            ShapeA.My3DCoordinates(Aa).Y = TempPoints.Y
        
            '--Z Vertices - Calculate depth
            ShapeA.MyPoints(Aa).X = ((ShapeA.My3DCoordinates(Aa).X - ShapeA.PosX) / (ShapeA.My3DCoordinates(Aa).Z - ShapeA.PosZ) * 600) + 160
            ShapeA.MyPoints(Aa).Y = ((ShapeA.My3DCoordinates(Aa).Y - ShapeA.PosY) / (ShapeA.My3DCoordinates(Aa).Z - ShapeA.PosZ) * 600) + 120
        Next Aa
        
        '-- Calculation Drawing Order
        ReDim TempDrawOrder(5)
        For Aa = 0 To 5
            TempDrawOrder(Aa) = (ShapeA.My3DCoordinates((Aa * 4)).Z + ShapeA.My3DCoordinates((Aa * 4) + 1).Z + ShapeA.My3DCoordinates((Aa * 4) + 2).Z + ShapeA.My3DCoordinates((Aa * 4) + 3).Z) / 4
            ShapeA.DrawOrder(Aa) = Aa '-- Reset this variable
        Next Aa
        For Aa = 0 To 4
            If TempDrawOrder(Aa) > TempDrawOrder(Aa + 1) Then
                '--Swaping Variables manually since there is no such function that I know of in VB
                Ab = ShapeA.DrawOrder(Aa)
                ShapeA.DrawOrder(Aa) = ShapeA.DrawOrder(Aa + 1)
                ShapeA.DrawOrder(Aa + 1) = Ab
                
                Ab = TempDrawOrder(Aa)
                TempDrawOrder(Aa) = TempDrawOrder(Aa + 1)
                TempDrawOrder(Aa + 1) = Ab
                Aa = Aa - 2
                If Aa < -1 Then Aa = -1
                'If Aa < -10 Then Debug.Print Error
            End If
        Next Aa
        'Debug.Print "<<<<-------->>>>"
        'For Aa = 0 To 5
        '    Debug.Print ShapeA.DrawOrder(Aa), TempDrawOrder(Aa)
        'Next Aa
        
        '-- Drawing
        'SelectObject BackBuffer, MyPens(0)
        'SelectObject BackBuffer, MyBrushes(0)
        For Aa = 0 To 5
            SelectObject BackBuffer, MyPens(ShapeA.DrawOrder(Aa) + 1)
            SelectObject BackBuffer, MyBrushes(ShapeA.DrawOrder(Aa) + 1)
            Polygon BackBuffer, ShapeA.MyPoints(ShapeA.DrawOrder(Aa) * 4), 4
        Next Aa
        
End Sub

'-- Call this Sub to Draw the ShapeB - Pyramid
Sub DrawShapeB()
    On Error Resume Next
       
        '--Adjusting the rotation variables
        ShapeB.TurnUD = ShapeB.TurnUD Mod 360
        ShapeB.TurnLR = ShapeB.TurnLR Mod 360
        ShapeB.TurnTU = ShapeB.TurnTU Mod 360
        
        '-- Calculation Of 3D Coordinates
        For Aa = 0 To 15
            '-- Set values to temporary variables for adjustment before drawing
            ShapeB.My3DCoordinates(Aa).X = ShapeB.My3DPoints(Aa).X
            ShapeB.My3DCoordinates(Aa).Y = ShapeB.My3DPoints(Aa).Y
            ShapeB.My3DCoordinates(Aa).Z = ShapeB.My3DPoints(Aa).Z

            '--Rotation
            TempPoints.X = (Cos(ShapeB.TurnTU * Pi) * ShapeB.My3DCoordinates(Aa).X) + (-Sin(ShapeB.TurnTU * Pi) * ShapeB.My3DCoordinates(Aa).Y)
            ShapeB.My3DCoordinates(Aa).Y = (Sin(ShapeB.TurnTU * Pi) * ShapeB.My3DCoordinates(Aa).X) + (Cos(ShapeB.TurnTU * Pi) * ShapeB.My3DCoordinates(Aa).Y)
            ShapeB.My3DCoordinates(Aa).X = TempPoints.X
        
            TempPoints.X = (Cos(ShapeB.TurnUD * Pi) * ShapeB.My3DCoordinates(Aa).X) + (-Sin(ShapeB.TurnUD * Pi) * ShapeB.My3DCoordinates(Aa).Z)
            ShapeB.My3DCoordinates(Aa).Z = (Sin(ShapeB.TurnUD * Pi) * ShapeB.My3DCoordinates(Aa).X) + (Cos(ShapeB.TurnUD * Pi) * ShapeB.My3DCoordinates(Aa).Z)
            ShapeB.My3DCoordinates(Aa).X = TempPoints.X
        
            TempPoints.Y = (Cos(ShapeB.TurnLR * Pi) * ShapeB.My3DCoordinates(Aa).Y) + (-Sin(ShapeB.TurnLR * Pi) * ShapeB.My3DCoordinates(Aa).Z)
            ShapeB.My3DCoordinates(Aa).Z = (Sin(ShapeB.TurnLR * Pi) * ShapeB.My3DCoordinates(Aa).Y) + (Cos(ShapeB.TurnLR * Pi) * ShapeB.My3DCoordinates(Aa).Z)
            ShapeB.My3DCoordinates(Aa).Y = TempPoints.Y
        
            '--Z Vertices - Calculate depth
            ShapeB.MyPoints(Aa).X = ((ShapeB.My3DCoordinates(Aa).X - ShapeB.PosX) / (ShapeB.My3DCoordinates(Aa).Z - ShapeB.PosZ) * 600) + 160
            ShapeB.MyPoints(Aa).Y = ((ShapeB.My3DCoordinates(Aa).Y - ShapeB.PosY) / (ShapeB.My3DCoordinates(Aa).Z - ShapeB.PosZ) * 600) + 120
        Next Aa
        
        '-- Calculation Drawing Order
        ReDim TempDrawOrder(4)
        For Aa = 0 To 3
            TempDrawOrder(Aa) = (ShapeB.My3DCoordinates((Aa * 3)).Z + ShapeB.My3DCoordinates((Aa * 3) + 1).Z + ShapeB.My3DCoordinates((Aa * 3) + 2).Z) / 3
            ShapeB.DrawOrder(Aa) = Aa '-- Reset this variable
        Next Aa
        Aa = 4
        TempDrawOrder(Aa) = (ShapeB.My3DCoordinates((Aa * 3)).Z + ShapeB.My3DCoordinates((Aa * 3) + 1).Z + ShapeB.My3DCoordinates((Aa * 3) + 2).Z + ShapeB.My3DCoordinates((Aa * 3) + 3).Z) / 4
        ShapeB.DrawOrder(Aa) = Aa '-- Reset this variable
        For Aa = 0 To 3
            If TempDrawOrder(Aa) > TempDrawOrder(Aa + 1) Then
                '--Swaping Variables manually since there is no such function that I know of in VB
                Ab = ShapeB.DrawOrder(Aa)
                ShapeB.DrawOrder(Aa) = ShapeB.DrawOrder(Aa + 1)
                ShapeB.DrawOrder(Aa + 1) = Ab
                
                Ab = TempDrawOrder(Aa)
                TempDrawOrder(Aa) = TempDrawOrder(Aa + 1)
                TempDrawOrder(Aa + 1) = Ab
                Aa = Aa - 2
                If Aa < -1 Then Aa = -1
                'If Aa < -10 Then Debug.Print Error
            End If
        Next Aa
        'Debug.Print "<<<<-------->>>>"
        'For Aa = 0 To 4
        '    Debug.Print ShapeB.DrawOrder(Aa), TempDrawOrder(Aa)
        'Next Aa
        
        '-- Drawing
        'SelectObject BackBuffer, MyPens(0)
        'SelectObject BackBuffer, MyBrushes(0)
        For Aa = 0 To 4
            'SelectObject BackBuffer, MyPens(ShapeB.DrawOrder(Aa) + 1)
            SelectObject BackBuffer, MyPens(0)
            SelectObject BackBuffer, MyBrushes(ShapeB.DrawOrder(Aa) + 1)
            If ShapeB.DrawOrder(Aa) < 4 Then Polygon BackBuffer, ShapeB.MyPoints(ShapeB.DrawOrder(Aa) * 3), 3
            If ShapeB.DrawOrder(Aa) = 4 Then Polygon BackBuffer, ShapeB.MyPoints(ShapeB.DrawOrder(Aa) * 3), 4
        Next Aa
        
End Sub

'-- Call this Sub to Draw the ShapeC - Cylinder
Sub DrawShapeC()
    On Error Resume Next
       
        '--Adjusting the rotation variables
        ShapeC.TurnUD = ShapeC.TurnUD Mod 360
        ShapeC.TurnLR = ShapeC.TurnLR Mod 360
        ShapeC.TurnTU = ShapeC.TurnTU Mod 360
        
        '-- Calculation Of 3D Coordinates
        For Aa = 0 To 107
            '-- Set values to temporary variables for adjustment before drawing
            ShapeC.My3DCoordinates(Aa).X = ShapeC.My3DPoints(Aa).X
            ShapeC.My3DCoordinates(Aa).Y = ShapeC.My3DPoints(Aa).Y
            ShapeC.My3DCoordinates(Aa).Z = ShapeC.My3DPoints(Aa).Z

            '--Rotation
            TempPoints.X = (Cos(ShapeC.TurnTU * Pi) * ShapeC.My3DCoordinates(Aa).X) + (-Sin(ShapeC.TurnTU * Pi) * ShapeC.My3DCoordinates(Aa).Y)
            ShapeC.My3DCoordinates(Aa).Y = (Sin(ShapeC.TurnTU * Pi) * ShapeC.My3DCoordinates(Aa).X) + (Cos(ShapeC.TurnTU * Pi) * ShapeC.My3DCoordinates(Aa).Y)
            ShapeC.My3DCoordinates(Aa).X = TempPoints.X
        
            TempPoints.X = (Cos(ShapeC.TurnUD * Pi) * ShapeC.My3DCoordinates(Aa).X) + (-Sin(ShapeC.TurnUD * Pi) * ShapeC.My3DCoordinates(Aa).Z)
            ShapeC.My3DCoordinates(Aa).Z = (Sin(ShapeC.TurnUD * Pi) * ShapeC.My3DCoordinates(Aa).X) + (Cos(ShapeC.TurnUD * Pi) * ShapeC.My3DCoordinates(Aa).Z)
            ShapeC.My3DCoordinates(Aa).X = TempPoints.X
        
            TempPoints.Y = (Cos(ShapeC.TurnLR * Pi) * ShapeC.My3DCoordinates(Aa).Y) + (-Sin(ShapeC.TurnLR * Pi) * ShapeC.My3DCoordinates(Aa).Z)
            ShapeC.My3DCoordinates(Aa).Z = (Sin(ShapeC.TurnLR * Pi) * ShapeC.My3DCoordinates(Aa).Y) + (Cos(ShapeC.TurnLR * Pi) * ShapeC.My3DCoordinates(Aa).Z)
            ShapeC.My3DCoordinates(Aa).Y = TempPoints.Y
        
            '--Z Vertices - Calculate depth
            ShapeC.MyPoints(Aa).X = ((ShapeC.My3DCoordinates(Aa).X - ShapeC.PosX) / (ShapeC.My3DCoordinates(Aa).Z - ShapeC.PosZ) * 600) + 160
            ShapeC.MyPoints(Aa).Y = ((ShapeC.My3DCoordinates(Aa).Y - ShapeC.PosY) / (ShapeC.My3DCoordinates(Aa).Z - ShapeC.PosZ) * 600) + 120
        Next Aa
        
        '-- Calculation Drawing Order
        ReDim TempDrawOrder(19)
        For Aa = 0 To 17
            TempDrawOrder(Aa) = (ShapeC.My3DCoordinates((Aa * 4)).Z + ShapeC.My3DCoordinates((Aa * 4) + 1).Z + ShapeC.My3DCoordinates((Aa * 4) + 2).Z + ShapeC.My3DCoordinates((Aa * 4) + 3).Z) / 4
            ShapeC.DrawOrder(Aa) = Aa '-- Reset this variable
        Next Aa
        TempDrawOrder(18) = 0
        For Aa = 0 To 17
            TempDrawOrder(18) = TempDrawOrder(18) + ShapeC.My3DCoordinates(Aa + 72).Z
        Next Aa
        TempDrawOrder(18) = TempDrawOrder(18) / 18
        TempDrawOrder(19) = 0
        For Aa = 0 To 17
            TempDrawOrder(19) = TempDrawOrder(19) + ShapeC.My3DCoordinates(Aa + 90).Z
        Next Aa
        TempDrawOrder(19) = TempDrawOrder(19) / 18
        ShapeC.DrawOrder(18) = 18 '-- Reset this variable
        ShapeC.DrawOrder(19) = 19 '-- Reset this variable
        
        For Aa = 0 To 18
            If TempDrawOrder(Aa) > TempDrawOrder(Aa + 1) Then
                '--Swaping Variables manually since there is no such function that I know of in VB
                Ab = ShapeC.DrawOrder(Aa)
                ShapeC.DrawOrder(Aa) = ShapeC.DrawOrder(Aa + 1)
                ShapeC.DrawOrder(Aa + 1) = Ab
        
                Ab = TempDrawOrder(Aa)
                TempDrawOrder(Aa) = TempDrawOrder(Aa + 1)
                TempDrawOrder(Aa + 1) = Ab
                Aa = Aa - 2
                If Aa < -1 Then Aa = -1
                'If Aa < -10 Then Debug.Print Error
            End If
        Next Aa
        'Debug.Print "<<<<-------->>>>"
        'For Aa = 0 To 5
        '    Debug.Print SHAPEc.DrawOrder(Aa), TempDrawOrder(Aa)
        'Next Aa
        
        '-- Drawing
        'SelectObject BackBuffer, MyPens(0)
        'SelectObject BackBuffer, MyBrushes(0)
        For Aa = 0 To 19
            'SelectObject BackBuffer, MyPens(ShapeC.DrawOrder(Aa) + 1)
            'SelectObject BackBuffer, MyBrushes(ShapeC.DrawOrder(Aa) + 1)
            If ShapeC.DrawOrder(Aa) <= 17 Then
                SelectObject BackBuffer, MyPens(0)
                SelectObject BackBuffer, MyBrushes(ShapeC.DrawOrder(Aa) Mod 2 + 3)
                Polygon BackBuffer, ShapeC.MyPoints(ShapeC.DrawOrder(Aa) * 4), 4
            ElseIf ShapeC.DrawOrder(Aa) = 18 Then
                SelectObject BackBuffer, MyPens(0)
                SelectObject BackBuffer, MyBrushes(3)
                Polygon BackBuffer, ShapeC.MyPoints(72), 18
            ElseIf ShapeC.DrawOrder(Aa) = 19 Then
                SelectObject BackBuffer, MyPens(0)
                SelectObject BackBuffer, MyBrushes(4)
                Polygon BackBuffer, ShapeC.MyPoints(90), 18
            End If
        Next Aa
        
End Sub

'-- Call this Sub to Draw the ShapeD - Cone
Sub DrawShapeD()
    On Error Resume Next
       
        '--Adjusting the rotation variables
        ShapeD.TurnUD = ShapeD.TurnUD Mod 360
        ShapeD.TurnLR = ShapeD.TurnLR Mod 360
        ShapeD.TurnTU = ShapeD.TurnTU Mod 360
        
        '-- Calculation Of 3D Coordinates
        For Aa = 0 To 71
            '-- Set values to temporary variables for adjustment before drawing
            ShapeD.My3DCoordinates(Aa).X = ShapeD.My3DPoints(Aa).X
            ShapeD.My3DCoordinates(Aa).Y = ShapeD.My3DPoints(Aa).Y
            ShapeD.My3DCoordinates(Aa).Z = ShapeD.My3DPoints(Aa).Z

            '--Rotation
            TempPoints.X = (Cos(ShapeD.TurnTU * Pi) * ShapeD.My3DCoordinates(Aa).X) + (-Sin(ShapeD.TurnTU * Pi) * ShapeD.My3DCoordinates(Aa).Y)
            ShapeD.My3DCoordinates(Aa).Y = (Sin(ShapeD.TurnTU * Pi) * ShapeD.My3DCoordinates(Aa).X) + (Cos(ShapeD.TurnTU * Pi) * ShapeD.My3DCoordinates(Aa).Y)
            ShapeD.My3DCoordinates(Aa).X = TempPoints.X
        
            TempPoints.X = (Cos(ShapeD.TurnUD * Pi) * ShapeD.My3DCoordinates(Aa).X) + (-Sin(ShapeD.TurnUD * Pi) * ShapeD.My3DCoordinates(Aa).Z)
            ShapeD.My3DCoordinates(Aa).Z = (Sin(ShapeD.TurnUD * Pi) * ShapeD.My3DCoordinates(Aa).X) + (Cos(ShapeD.TurnUD * Pi) * ShapeD.My3DCoordinates(Aa).Z)
            ShapeD.My3DCoordinates(Aa).X = TempPoints.X
        
            TempPoints.Y = (Cos(ShapeD.TurnLR * Pi) * ShapeD.My3DCoordinates(Aa).Y) + (-Sin(ShapeD.TurnLR * Pi) * ShapeD.My3DCoordinates(Aa).Z)
            ShapeD.My3DCoordinates(Aa).Z = (Sin(ShapeD.TurnLR * Pi) * ShapeD.My3DCoordinates(Aa).Y) + (Cos(ShapeD.TurnLR * Pi) * ShapeD.My3DCoordinates(Aa).Z)
            ShapeD.My3DCoordinates(Aa).Y = TempPoints.Y
        
            '--Z Vertices - Calculate depth
            ShapeD.MyPoints(Aa).X = ((ShapeD.My3DCoordinates(Aa).X - ShapeD.PosX) / (ShapeD.My3DCoordinates(Aa).Z - ShapeD.PosZ) * 600) + 160
            ShapeD.MyPoints(Aa).Y = ((ShapeD.My3DCoordinates(Aa).Y - ShapeD.PosY) / (ShapeD.My3DCoordinates(Aa).Z - ShapeD.PosZ) * 600) + 120
        Next Aa
        
        '-- Calculation Drawing Order
        ReDim TempDrawOrder(18)
        For Aa = 0 To 17
            TempDrawOrder(Aa) = (ShapeD.My3DCoordinates((Aa * 3)).Z + ShapeD.My3DCoordinates((Aa * 3) + 1).Z + ShapeD.My3DCoordinates((Aa * 3) + 2).Z) / 3
            ShapeD.DrawOrder(Aa) = Aa '-- Reset this variable
        Next Aa
        TempDrawOrder(18) = 0
        For Aa = 0 To 17
            TempDrawOrder(18) = TempDrawOrder(18) + ShapeD.My3DCoordinates(Aa + 54).Z
        Next Aa
        TempDrawOrder(18) = TempDrawOrder(18) / 18
        ShapeD.DrawOrder(18) = 18 '-- Reset this variable
        
        For Aa = 0 To 17
            If TempDrawOrder(Aa) > TempDrawOrder(Aa + 1) Then
                '--Swaping Variables manually since there is no such function that I know of in VB
                Ab = ShapeD.DrawOrder(Aa)
                ShapeD.DrawOrder(Aa) = ShapeD.DrawOrder(Aa + 1)
                ShapeD.DrawOrder(Aa + 1) = Ab
        
                Ab = TempDrawOrder(Aa)
                TempDrawOrder(Aa) = TempDrawOrder(Aa + 1)
                TempDrawOrder(Aa + 1) = Ab
                Aa = Aa - 2
                If Aa < -1 Then Aa = -1
                'If Aa < -10 Then Debug.Print Error
            End If
        Next Aa
        'Debug.Print "<<<<-------->>>>"
        'For Aa = 0 To 5
        '    Debug.Print SHAPEd.DrawOrder(Aa), TempDrawOrder(Aa)
        'Next Aa
        
        '-- Drawing
        'SelectObject BackBuffer, MyPens(0)
        'SelectObject BackBuffer, MyBrushes(0)
        For Aa = 0 To 18
            'SelectObject BackBuffer, MyPens(SHAPEd.DrawOrder(Aa) + 1)
            'SelectObject BackBuffer, MyBrushes(SHAPEd.DrawOrder(Aa) + 1)
            If ShapeD.DrawOrder(Aa) <= 17 Then
                SelectObject BackBuffer, MyPens(0)
                SelectObject BackBuffer, MyBrushes(ShapeD.DrawOrder(Aa) Mod 2 + 1)
                Polygon BackBuffer, ShapeD.MyPoints(ShapeD.DrawOrder(Aa) * 3), 3
            ElseIf ShapeD.DrawOrder(Aa) = 18 Then
                SelectObject BackBuffer, MyPens(0)
                SelectObject BackBuffer, MyBrushes(3)
                Polygon BackBuffer, ShapeD.MyPoints(54), 18
            End If
        Next Aa
        
End Sub

'-- Call this Sub to Draw the ShapeE - Sphere
Sub DrawShapeE()
    '----Draw the Sphere
    SelectObject BackBuffer, MyPens(0)
    SelectObject BackBuffer, MyBrushes(3)
    Aa = (1000 / ShapeE.PosZ * 80) '--Sphere size
            
    'Ellipse BackBuffer, ShapeE.PosX + 160 - Aa, ShapeE.PosY + 120 - Aa, ShapeE.PosX + 160 + Aa, ShapeE.PosY + 120 + Aa
    Ellipse BackBuffer, (ShapeE.PosX / ShapeE.PosZ * 600) + 160 - Aa, (ShapeE.PosY / ShapeE.PosZ * 600) + 120 - Aa, (ShapeE.PosX / ShapeE.PosZ * 600) + 160 + Aa, (ShapeE.PosY / ShapeE.PosZ * 600) + 120 + Aa
    
End Sub

'-- Call this sub to create Pens and Brushes to draw with
Sub CreatePensBrushes()
    '----Creating Pens
    ReDim MyPens(6)
    MyPens(0) = CreatePen(PS_SOLID, 1, RGB(0, 0, 0))
    MyPens(1) = CreatePen(PS_SOLID, 1, RGB(255, 0, 0))
    MyPens(2) = CreatePen(PS_SOLID, 1, RGB(0, 255, 0))
    MyPens(3) = CreatePen(PS_SOLID, 1, RGB(0, 0, 255))
    MyPens(4) = CreatePen(PS_SOLID, 1, RGB(255, 255, 0))
    MyPens(5) = CreatePen(PS_SOLID, 1, RGB(0, 255, 255))
    MyPens(6) = CreatePen(PS_SOLID, 1, RGB(255, 0, 255))
    
    '----Creating Brushes
    ReDim MyBrushes(6)
    MyBrushes(0) = CreateSolidBrush(RGB(0, 0, 0))
    MyBrushes(1) = CreateSolidBrush(RGB(255, 0, 0))
    MyBrushes(2) = CreateSolidBrush(RGB(0, 255, 0))
    MyBrushes(3) = CreateSolidBrush(RGB(0, 0, 255))
    MyBrushes(4) = CreateSolidBrush(RGB(255, 255, 0))
    MyBrushes(5) = CreateSolidBrush(RGB(0, 255, 255))
    MyBrushes(6) = CreateSolidBrush(RGB(255, 0, 255))

End Sub

'-- Call this sub to remove the Pens and Brushes that was created by 'CreatePensBrushes'
Sub DeletePensBrushes()
    '----Deleting Pens
    For Aa = 0 To 6
        DeleteObject MyPens(Aa)
    Next Aa

    '----Deleting Brushes
    For Aa = 0 To 6
        DeleteObject MyBrushes(Aa)
    Next Aa

End Sub

'-- Call this sub to arrange the order in which to draw the Shapes and draw them automatically in that order
Sub DrawAllShapes()
    
        '-- Calculation Drawing Order
        ReDim TempDrawOrder(4)
        ReDim TempArray(4)
        TempDrawOrder(0) = ShapeA.PosZ
        TempDrawOrder(1) = ShapeB.PosZ
        TempDrawOrder(2) = ShapeC.PosZ
        TempDrawOrder(3) = ShapeD.PosZ
        TempDrawOrder(4) = ShapeE.PosZ
        For Aa = 0 To 4
            TempArray(Aa) = Aa '-- reset this variable with incrementing numbers
        Next Aa
        For Aa = 0 To 3
            If TempDrawOrder(Aa) > TempDrawOrder(Aa + 1) Then
                '--Swaping Variables manually since there is no such function that I know of in VB
                Ab = TempArray(Aa)
                TempArray(Aa) = TempArray(Aa + 1)
                TempArray(Aa + 1) = Ab
                
                Ab = TempDrawOrder(Aa)
                TempDrawOrder(Aa) = TempDrawOrder(Aa + 1)
                TempDrawOrder(Aa + 1) = Ab
                Aa = Aa - 2
                If Aa < -1 Then Aa = -1
                'If Aa < -10 Then Debug.Print Error
            End If
        Next Aa
        'Debug.Print "<<<<-------->>>>"
        'For Aa = 0 To 5
        '    Debug.Print ShapeA.DrawOrder(Aa), TempDrawOrder(Aa)
        'Next Aa
        
        '---- Calling the Shapes drawing Subs to draw in the appropriate order
        For Ac = 0 To 4 '-- ref. don't use the variable 'Ac' in the Shapes drawing Subs
            Select Case TempArray(4 - Ac)
                Case 0 '--Draw Cube
                    DrawShapeA
                    
                Case 1 '--Draw Pyramid
                    DrawShapeB
                    
                Case 2 '--Draw Cylinder
                    DrawShapeC
                    
                Case 3 '--Draw Cone
                    DrawShapeD
                    
                Case 4 '--Draw Sphere
                    DrawShapeE
                
            End Select
            
        Next Ac
        
        
End Sub
