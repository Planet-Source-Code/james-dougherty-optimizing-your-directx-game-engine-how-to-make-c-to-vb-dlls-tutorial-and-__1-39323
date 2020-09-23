VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   ".Dll Test"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRender 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stop Render"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdRender 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Start Render"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox Canvas 
      BackColor       =   &H00000000&
      Height          =   6750
      Left            =   120
      ScaleHeight     =   446
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   446
      TabIndex        =   2
      Top             =   1440
      Width           =   6750
   End
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Test Flip Normals"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Test Distance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblFPS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FPS = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblDist 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distance = ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      TabIndex        =   7
      Top             =   240
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vector2.Y = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vector1.Y = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vector2.X = 10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vector1.X = 50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRender_Click(Index As Integer)
 
 Select Case Index
  Case 0 'Start Render
   cmdTest(1).Enabled = True
   cmdRender(0).Enabled = False
   cmdRender(1).Enabled = True
   EndLoop = False
   Render
  Case 1 'Stop Render
   cmdTest(1).Enabled = False
   cmdRender(0).Enabled = True
   cmdRender(1).Enabled = False
   EndLoop = True
 End Select
 
 
End Sub

Private Sub cmdTest_Click(Index As Integer)
 
 'Look at the bottom
 Select Case Index
  Case 0
   TestDistance
  Case 1
   FlipNormals
 End Select
 
End Sub

Private Sub Form_Load()
 InitializeDX Canvas.hWnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 EndLoop = True
 CleanupDX
End Sub

'THESE 2 FUNCTIONS USE THE API'S
Public Sub TestDistance()
 Dim Distance As Single
 Dim Position1 As D3DVECTOR2
 Dim Position2 As D3DVECTOR2
 
 'Give some test values
 Position1.X = 50
 Position1.Y = 0
 Position2.X = 10
 Position2.Y = 0
 
 'Get the distance from the .dll
 Distance = Distance2D(Position1, Position2)
 'Display the distance
 lblDist = "Distance = " & Distance
 
 'The result should be 40
End Sub

Public Sub FlipNormals()
 On Local Error Resume Next
 Dim VB As Direct3DVertexBuffer8
 Dim Vertices() As D3DVERTEX
 Dim NumVertices As Long
 Dim Size As Long
 Dim i As Long
 
 'Get all the data
 Set VB = g_Mesh.GetVertexBuffer()
 Size = g_D3DX.GetFVFVertexSize(g_Mesh.GetFVF())
 NumVertices = g_Mesh.GetNumVertices()
 ReDim Vertices(NumVertices) As D3DVERTEX
 D3DVertexBuffer8GetData VB, 0, Size * NumVertices, 0, Vertices(0)
 
 'Call the .dll to flip the normals
 'Note: MAKE SURE THE FVF CONTAINS D3DFVF_VERTEX OR IT WILL COME OUT FUNNY
 FlipMeshNormals Vertices(0), NumVertices
 
 'set all the data
 D3DVertexBuffer8SetData VB, 0, Size * NumVertices, 0, Vertices(0)
 
 Set VB = Nothing
 Erase Vertices
 
End Sub
