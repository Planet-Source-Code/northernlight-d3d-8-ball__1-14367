VERSION 5.00
Begin VB.Form frmVB3d 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "8 Ball NLX "
   ClientHeight    =   5910
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      Visible         =   0   'False
      X1              =   630
      X2              =   1350
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Menu mnuRender 
      Caption         =   "Render"
      Begin VB.Menu mnuWireframe 
         Caption         =   "Wireframe"
      End
      Begin VB.Menu mnuFlat 
         Caption         =   "Flat"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmVB3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDx7 As DirectX7
Private mDrw As DirectDraw7
Private mDrm As Direct3DRM3
Private mFrS As Direct3DRMFrame3
Private mFrC As Direct3DRMFrame3
Private mFrO As Direct3DRMFrame3
Private mFrL As Direct3DRMFrame3
Private mDev As Direct3DRMDevice3
Private mVpt As Direct3DRMViewport2

Private mDownX As Single
Private mDownY As Single
Private mStopFlag As Boolean
Private mMouseDown As Boolean

Private Type dxPTM
    dX As Single
    dY As Single
    Distance As Single
End Type

Private Sub LoadMesh()
Dim DxMeshB As Direct3DRMMeshBuilder3

    mDrm.SetSearchPath App.Path
    Set DxMeshB = mDrm.CreateMeshBuilder()
    With DxMeshB
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture mDrm.LoadTexture("8ball.bmp")
    End With
    
    mFrO.AddVisual DxMeshB
    
    Me.Show: DoEvents
    
End Sub

Private Sub Rotate(X As Single, Y As Single, Optional Button As Integer)
Dim PTM As dxPTM
Dim Theta As Single
    
    PointToMouse PTM, X, Y
    
    With PTM
        Theta = .Distance / 1000
        mFrO.SetRotation Nothing, .dY, .dX, 0, Theta
    End With
    
End Sub
Private Sub RefreshLoop()

    Do While mStopFlag = False
        mFrS.Move 1
        With mVpt
            .Clear D3DRMCLEAR_ALL
            .Render mFrS
        End With
        mDev.Update
        DoEvents
    Loop
    
End Sub

Private Sub PointToMouse(PTM As dxPTM, X As Single, Y As Single)
Dim sX As Single, sY As Single
    
    With PTM
        .dX = mDownX - X
        .dY = mDownY - Y
        sX = (.dX * .dX)
        sY = (.dY * .dY)
        .Distance = Sqr(sX + sY)
    End With
    
    With Line1
        .X1 = mDownX
        .Y1 = mDownY
        .X2 = X
        .Y2 = Y
    End With
    
End Sub


Private Sub Initialise()
    Set mDx7 = New DirectX7
    Set mDrm = mDx7.Direct3DRMCreate
    Set mDrw = mDx7.DirectDrawCreate("")
End Sub

Private Sub CreateSceneGraph()
Dim DxL1 As Direct3DRMLight
Dim DxL2 As Direct3DRMLight

    With mDrm
        Set mFrS = .CreateFrame(Nothing)
        Set mFrC = .CreateFrame(mFrS)
        Set mFrO = .CreateFrame(mFrS)
        Set mFrL = .CreateFrame(mFrS)
        Set DxL1 = .CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 0.8, 0.8, 0.8)
        Set DxL2 = .CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5)
    End With
    mFrL.AddLight DxL1
    mFrL.AddLight DxL2
    mFrC.SetPosition Nothing, 0, 0, -3
End Sub

Private Sub CreateDisplay()
Dim DxClipper As DirectDrawClipper

    Set mVpt = Nothing
    Set mDev = Nothing
    Set DxClipper = mDrw.CreateClipper(0)
    
    ScaleMode = vbPixels
    DxClipper.SetHWnd hWnd
    Set mDev = mDrm.CreateDeviceFromClipper(DxClipper, "", ScaleWidth, ScaleHeight)
    Set mVpt = mDrm.CreateViewport(mDev, mFrC, 0, 0, ScaleWidth, ScaleHeight)

End Sub

Private Sub Form_Load()
    
    Initialise
    CreateSceneGraph
    CreateDisplay
    LoadMesh
    RefreshLoop
    Cleanup
    End

End Sub

Private Sub HitTest(X As Single, Y As Single)
Dim PickArray As Direct3DRMPickArray
Dim Desc As D3DRMPICKDESC

    Set PickArray = mVpt.Pick(CLng(X), CLng(Y))
    If PickArray.GetSize() = 0 Then
        Caption = "Drag the ball"
    Else
        Caption = "Drag !"
    End If
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMouseDown = True
    mDownX = X
    mDownY = Y
    HitTest X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mMouseDown = True Then Exit Sub
    Rotate X, Y
    Line1.Visible = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMouseDown = False
    Line1.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mStopFlag = True
End Sub

Private Sub Form_Resize()
    CreateDisplay
End Sub

Private Sub mnuExit_Click()
    mStopFlag = True
End Sub

Public Sub Cleanup()
    Set mVpt = Nothing
    Set mDev = Nothing
    Set mFrL = Nothing
    Set mFrO = Nothing
    Set mFrC = Nothing
    Set mFrS = Nothing
    Set mDrm = Nothing
    Set mDx7 = Nothing
End Sub

Private Sub SetQuality(Quality As CONST_D3DRMRENDERQUALITY)

    mDev.SetQuality Quality
    mnuFlat.Checked = False
    mnuWireframe.Checked = False
    Select Case Quality
        Case D3DRMRENDER_FLAT
            mnuFlat.Checked = True
        Case D3DRMRENDER_WIREFRAME
            mnuWireframe.Checked = True
    End Select
    
End Sub

Private Sub mnuFlat_Click()
    SetQuality D3DRMRENDER_FLAT
End Sub

Private Sub mnuWireframe_Click()
    SetQuality D3DRMRENDER_WIREFRAME
End Sub
