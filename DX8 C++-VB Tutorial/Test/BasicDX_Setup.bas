Attribute VB_Name = "BasicDX_Setup"
Option Explicit

Private g_DX As New DirectX8
Public g_D3DX As New D3DX8
Private g_D3D As Direct3D8
Private g_D3DD As Direct3DDevice8
Public g_Mesh As D3DXMesh
Private g_MeshMaterials() As D3DMATERIAL8
Private g_MeshTextures() As Direct3DTexture8
Private g_NumMaterials As Long
Private g_FPS As Single
Public EndLoop As Boolean

Public Function InitializeDX(hWnd As Long) As Boolean
 On Local Error GoTo ErrOut
 Dim D3DPP As D3DPRESENT_PARAMETERS
 Dim mode As D3DDISPLAYMODE
    
 Set g_D3D = g_DX.Direct3DCreate()
 If g_D3D Is Nothing Then Exit Function
 g_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
         
 D3DPP.Windowed = 1
 D3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD
 D3DPP.BackBufferFormat = mode.Format
 D3DPP.BackBufferCount = 1
 D3DPP.EnableAutoDepthStencil = 1
 D3DPP.AutoDepthStencilFormat = D3DFMT_D16

 Set g_D3DD = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
 If g_D3DD Is Nothing Then Exit Function
 g_D3DD.SetRenderState D3DRS_ZENABLE, 1
 SetupLights
 LoadMesh
 InitializeDX = True
 Exit Function
    
ErrOut: InitializeDX = False
End Function

Public Sub CleanupDX()

 Erase g_MeshTextures
 Erase g_MeshMaterials
 
 Set g_Mesh = Nothing
 Set g_D3DD = Nothing
 Set g_D3D = Nothing
 Set g_D3DX = Nothing
 Set g_DX = Nothing
 
End Sub

Private Function LoadMesh() As Boolean
 On Local Error GoTo ErrOut
 Dim MBuffer As D3DXBuffer
 Dim Texture As String
 Dim i As Long
    
 Set g_Mesh = g_D3DX.LoadMeshFromX(App.Path + "\Hand01.x", D3DXMESH_MANAGED, g_D3DD, Nothing, MBuffer, g_NumMaterials)
 If g_Mesh Is Nothing Then Exit Function

 ReDim g_MeshMaterials(g_NumMaterials - 1)
 ReDim g_MeshTextures(g_NumMaterials - 1)
    
 For i = 0 To g_NumMaterials - 1
  g_D3DX.BufferGetMaterial MBuffer, i, g_MeshMaterials(i)
  g_MeshMaterials(i).Ambient = g_MeshMaterials(i).diffuse
  Texture = g_D3DX.BufferGetTextureName(MBuffer, i)
  If Texture <> "" Then Set g_MeshTextures(i) = g_D3DX.CreateTextureFromFile(g_D3DD, App.Path + "\" + Texture)
 Next
 
 Set MBuffer = Nothing
 Change_FVF D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_VERTEX
 LoadMesh = True
 Exit Function
    
ErrOut: LoadMesh = False
End Function

Public Sub Render()
 Dim i As Long
 If g_D3DD Is Nothing Then Exit Sub
 
 Do
  g_D3DD.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1#, 0
  SetupMatrices
  g_D3DD.BeginScene
 
  For i = 0 To g_NumMaterials - 1
   g_D3DD.SetMaterial g_MeshMaterials(i)
   g_D3DD.SetTexture 0, g_MeshTextures(i)
   g_Mesh.DrawSubset i
  Next
 
  g_D3DD.EndScene
  g_D3DD.Present ByVal 0, ByVal 0, 0, ByVal 0
  Update_FPS
  DoEvents
 Loop Until EndLoop = True
 
End Sub

Private Sub SetupMatrices()
 Dim Projection As D3DMATRIX
 Dim World As D3DMATRIX
 Dim View As D3DMATRIX
 Static Angle As Single
 
 Angle = Angle + 0.01
 If Angle > 360 Then Angle = 0
 D3DXMatrixRotationX World, Angle
 g_D3DD.SetTransform D3DTS_WORLD, World
 
 D3DXMatrixLookAtLH View, Vector3(0#, 5#, -250#), Vector3(0#, 0#, 0#), Vector3(0#, 1#, 0#)
 g_D3DD.SetTransform D3DTS_VIEW, View
 D3DXMatrixPerspectiveFovLH Projection, 3.141 / 4, 1, 1, 1000
 g_D3DD.SetTransform D3DTS_PROJECTION, Projection

End Sub

Private Sub SetupLights()
 Dim Light As D3DLIGHT8
 
 Light.Type = D3DLIGHT_DIRECTIONAL
 Light.diffuse.r = 0.7
 Light.diffuse.g = 0.7
 Light.diffuse.b = 0.7
 Light.Position.Z = -100
 Light.Direction.X = 1#
 Light.Direction.Y = -1#
 Light.Direction.Z = 1#
 Light.Range = 1000#
    
 g_D3DD.SetLight 0, Light
 g_D3DD.LightEnable 0, 1
 g_D3DD.SetRenderState D3DRS_LIGHTING, 1
 
End Sub

Private Function Update_FPS()
 Static TimerEnd As Single
 Static TimerCurrent As Single
 Static FPS As Single
 Static i As Long
 
 i = i + 1
 If (i = 30) Then
  TimerCurrent = Timer
    If TimerCurrent <> TimerEnd Then
      FPS = 30 / (Timer - TimerEnd)
      TimerEnd = Timer
      i = 0
      g_FPS = FPS
    End If
 End If
 frmMain.lblFPS = "FPS - " & Format$(g_FPS, "###0.00")
 
End Function

Private Sub Change_FVF(FVF As Long)
 Dim Mesh As D3DXMesh
 On Local Error GoTo ErrOut

 If g_Mesh Is Nothing Then Exit Sub
 Set Mesh = g_Mesh.CloneMeshFVF(D3DXMESH_MANAGED, FVF, g_D3DD)
 Set g_Mesh = Nothing
 Set g_Mesh = Mesh
 Set Mesh = Nothing
 Exit Sub

ErrOut:
End Sub

Private Function Vector3(X As Single, Y As Single, Z As Single) As D3DVECTOR
 Vector3.X = X
 Vector3.Y = Y
 Vector3.Z = Z
End Function
