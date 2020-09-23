Attribute VB_Name = "SetupDemoAPI"
Option Explicit

Public Declare Function FlipMeshNormals Lib "SetupDemo.dll" (ByRef Vertices As D3DVERTEX, ByVal NumVertices As Long) As Boolean
Public Declare Function Distance2D Lib "SetupDemo.dll" (ByRef Vector1 As D3DVECTOR2, ByRef Vector2 As D3DVECTOR2) As Single
