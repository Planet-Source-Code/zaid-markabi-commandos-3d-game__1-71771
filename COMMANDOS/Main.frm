VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Set TV = New TVEngine
Set InputEngine = New TVInputEngine
Set CollisionResult = New TVCollisionResult
Set Screen2DText = New TVScreen2DText
Set Screen2DImmediate = New TVScreen2DImmediate

Dim x As String
Open App.Path + "\Data\Mission.txt" For Input As #1
Input #1, x
MissionNum = Int(x)
Close #1

MaxSpeed = 1 ' << set Game Speed here


Open App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\Setting.txt" For Input As #1
Input #1, x
Input #1, x
PlayerHealth = Int(x)
Input #1, x
Input #1, x
PlayerBombNum = Int(x)
Input #1, x
Input #1, x
PlayerWeapon = x
Close #1

TV.Init3DWindowedMode Me.hWnd
TV.SetAngleSystem TV_ANGLE_DEGREE

LoadTextures
LoadMap
BuildNodesOfPath
LoadPlayer
IntBomb
IntParticle
IntSound

MainLoop
End Sub

Private Sub Form_Resize()
TV.ResizeDevice
End Sub

Private Sub LoadMap()
Set Walls = Nothing
Set Walls = Scene.CreateMeshBuilder()
Set Map = Nothing
Set Map = Scene.CreateMeshBuilder()
Set GotoBlock = Nothing
Set GotoBlock = Scene.CreateMeshBuilder()
Set UpSide = Nothing
Set UpSide = Scene.CreateMeshBuilder()

UpSide.SetCollisionEnable False

GotoBlock.AddFloor GetTex("GoTo"), -30, -30, 30, 30, 1
GotoBlock.SetPosition 0, -5, 0

Dim XI As Integer, YI As Integer
Dim BlockName As String, BlockNameI As String
Dim TxtName As String, TxtName2 As String

Open App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\Map.txt" For Input As #1
Open App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\MapTx.txt" For Input As #3
Open App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\MapTx2.txt" For Input As #4
Open App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\MapI.txt" For Input As #5
Open App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\MapTxI.txt" For Input As #6

Input #1, ScaleMapX
Input #1, ScaleMapY

Map.AddFloor GetTex("TexF"), -500, -500, (ScaleMapX * 100) + 500, (ScaleMapY * 100) + 500, -1, ScaleMapX, ScaleMapY

For YI = 0 To ScaleMapY - 1
For XI = 0 To ScaleMapX - 1
Input #1, BlockName
Input #5, BlockNameI
Input #3, TxtName
Input #4, TxtName2
Input #6, TxtName3

Select Case BlockName

Case Is = "F":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Case Is = "I":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 50, (YI * 100), (XI * 100) + 50, (YI * 100) + 100, 100
Map.AddFloor GetTex("Tex" + TxtName3), (XI * 100) + 45, (YI * 100), (XI * 100) + 55, (YI * 100) + 100, 99
Case Is = "=":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100) + 50, (XI * 100) + 100, (YI * 100) + 50, 100
Map.AddFloor GetTex("Tex" + TxtName3), (XI * 100), (YI * 100) + 45, (XI * 100) + 100, (YI * 100) + 55, 99
Case Is = "{":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 50, (YI * 100) + 50, (XI * 100) + 100, (YI * 100) + 50, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 50, (YI * 100) + 50, (XI * 100) + 50, (YI * 100) + 100, 100
Case Is = "}":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100) + 50, (XI * 100) + 50, (YI * 100) + 50, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 50, (YI * 100) + 50, (XI * 100) + 50, (YI * 100) + 100, 100
Case Is = "[":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 50, (YI * 100), (XI * 100) + 50, (YI * 100) + 50, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 50, (YI * 100) + 50, (XI * 100) + 100, (YI * 100) + 50, 100
Case Is = "]":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 50, (YI * 100), (XI * 100) + 50, (YI * 100) + 50, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 50, (YI * 100) + 50, (XI * 100), (YI * 100) + 50, 100
Case Is = "^":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100) - 50, (XI * 100), (YI * 100) + 100, 50
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 100, (YI * 100) - 50, (XI * 100) + 100, (YI * 100) + 100, 50
Case Is = "v":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100), (XI * 100), (YI * 100) + 50, 50
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 100, (YI * 100), (XI * 100) + 100, (YI * 100) + 50, 50
Map.AddFloor GetTex("Tex" + TxtName3), (XI * 100) + 90, (YI * 100) - 10, (XI * 100) + 100, (YI * 100) + 50, 49
Map.AddFloor GetTex("Tex" + TxtName3), (XI * 100), (YI * 100) - 10, (XI * 100) + 10, (YI * 100) + 50, 49
Map.AddWall GetTex("TexD"), (XI * 100), (YI * 100) + 50, (XI * 100) + 100, (YI * 100) + 50, 100
AddTree Vector((XI * 100) - 50, 0, (YI * 100) + 25), True
AddTree Vector((XI * 100) + 150, 0, (YI * 100) + 25), True
Case Is = "b":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100, 50
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100), (XI * 100) + 100, (YI * 100), 50
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100), (XI * 100), (YI * 100) + 100, 50
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 100, (YI * 100), (XI * 100) + 100, (YI * 100) + 100, 50
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100) + 100, (XI * 100) + 100, (YI * 100) + 100, 50
Case Is = "B":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 25, (YI * 100) + 25, (XI * 100) + 75, (YI * 100) + 25, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 25, (YI * 100) + 25, (XI * 100) + 25, (YI * 100) + 75, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 75, (YI * 100) + 25, (XI * 100) + 75, (YI * 100) + 75, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 25, (YI * 100) + 75, (XI * 100) + 75, (YI * 100) + 75, 100
Map.AddFloor GetTex("Tex" + TxtName3), (XI * 100) + 25, (YI * 100) + 25, (XI * 100) + 75, (YI * 100) + 75, 100
Case Is = "p":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 25, (YI * 100) + 25, (XI * 100) + 75, (YI * 100) + 25, 50
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 25, (YI * 100) + 25, (XI * 100) + 25, (YI * 100) + 75, 50
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 75, (YI * 100) + 25, (XI * 100) + 75, (YI * 100) + 75, 50
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 25, (YI * 100) + 75, (XI * 100) + 75, (YI * 100) + 75, 50
Map.AddFloor GetTex("Tex" + TxtName3), (XI * 100) + 25, (YI * 100) + 25, (XI * 100) + 75, (YI * 100) + 75, 50
Case Is = "P":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100), (XI * 100) + 100, (YI * 100), 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100), (XI * 100), (YI * 100) + 100, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100) + 100, (YI * 100), (XI * 100) + 100, (YI * 100) + 100, 100
Walls.AddWall GetTex("Tex" + TxtName2), (XI * 100), (YI * 100) + 100, (XI * 100) + 100, (YI * 100) + 100, 100
Case Is = "C":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Set Object(ObjNum) = New TVActor2
Object(ObjNum).Load App.Path + "\Data\Model\truck.mdl"
Object(ObjNum).SetScale 0.7, 0.7, 0.7
Object(ObjNum).SetPosition (XI * 100) + 50, 0, (YI * 100) + 50
Walls.AddWall GetTex("NO"), (XI * 100) - 100, (YI * 100), (XI * 100) + 150, (YI * 100), 50
Walls.AddWall GetTex("NO"), (XI * 100) - 100, (YI * 100), (XI * 100) - 100, (YI * 100) + 100, 50
Walls.AddWall GetTex("NO"), (XI * 100) + 150, (YI * 100), (XI * 100) + 150, (YI * 100) + 100, 50
Walls.AddWall GetTex("NO"), (XI * 100) - 100, (YI * 100) + 100, (XI * 100) + 200, (YI * 100) + 100, 50
Object(ObjNum).SetName "Object" + Format(ObjNum)
ObjectDone(ObjNum) = False
ObjNum = ObjNum + 1
ObjDoneNum = ObjDoneNum + 1
Case Is = "T":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
AddTree Vector((XI * 100) - 50, 0, (YI * 100) - 50), True
Case Is = "t":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
AddTree Vector((XI * 100) - 50, 0, (YI * 100) - 50), False
Case Is = "H":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Set ItemO(ItemNum) = New TVMesh
Set ItemO(ItemNum) = Scene.CreateMeshBuilder()
ItemO(ItemNum).SetCollisionEnable False
ItemO(ItemNum).SetPosition (XI * 100) + 50, 51, (YI * 100) + 50
ItemO(ItemNum).CreateBox 50, 50, 50, False
ItemO(ItemNum).SetTexture GetTex("Health")
ItemMode(ItemNum) = "Health"
ItemNum = ItemNum + 1
Case Is = "W":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Set ItemO(ItemNum) = New TVMesh
Set ItemO(ItemNum) = Scene.CreateMeshBuilder()
ItemO(ItemNum).SetCollisionEnable False
ItemO(ItemNum).SetPosition (XI * 100) + 50, 51, (YI * 100) + 50
ItemO(ItemNum).CreateBox 50, 50, 50, False
ItemO(ItemNum).SetTexture GetTex("Weapon")
ItemMode(ItemNum) = "Weapon"
ItemNum = ItemNum + 1
Case Is = "Z":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Set ItemO(ItemNum) = New TVMesh
Set ItemO(ItemNum) = Scene.CreateMeshBuilder()
ItemO(ItemNum).SetCollisionEnable False
ItemO(ItemNum).SetPosition (XI * 100) + 50, 51, (YI * 100) + 50
ItemO(ItemNum).CreateBox 50, 50, 50, False
ItemO(ItemNum).SetTexture GetTex("Bomb")
ItemMode(ItemNum) = "Bomb"
ItemNum = ItemNum + 1
Case Is = "l":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Soor"), (XI * 100) + 50, (YI * 100), (XI * 100) + 50, (YI * 100) + 100, 50
Case Is = "-":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Soor"), (XI * 100), (YI * 100) + 50, (XI * 100) + 100, (YI * 100) + 50, 50
Case Is = "e":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Soor"), (XI * 100) + 50, (YI * 100) + 50, (XI * 100) + 100, (YI * 100) + 50, 50
Walls.AddWall GetTex("Soor"), (XI * 100) + 50, (YI * 100) + 50, (XI * 100) + 50, (YI * 100) + 100, 50
Case Is = "r":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Soor"), (XI * 100), (YI * 100) + 50, (XI * 100) + 50, (YI * 100) + 50, 50
Walls.AddWall GetTex("Soor"), (XI * 100) + 50, (YI * 100) + 50, (XI * 100) + 50, (YI * 100) + 100, 50
Case Is = "E":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Soor"), (XI * 100) + 50, (YI * 100), (XI * 100) + 50, (YI * 100) + 50, 50
Walls.AddWall GetTex("Soor"), (XI * 100) + 50, (YI * 100) + 50, (XI * 100) + 100, (YI * 100) + 50, 50
Case Is = "R":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Walls.AddWall GetTex("Soor"), (XI * 100) + 50, (YI * 100), (XI * 100) + 50, (YI * 100) + 50, 50
Walls.AddWall GetTex("Soor"), (XI * 100) + 50, (YI * 100) + 50, (XI * 100), (YI * 100) + 50, 50
Case Is = "M":
Map.AddFloor GetTex("Tex" + TxtName), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100
Set Enemy(EnemyNum) = New TVActor2
Set EnemyPath(EnemyNum) = New TVPath
Enemy(EnemyNum).Load App.Path + "\Data\Model\barney.mdl"
Enemy(EnemyNum).SetAnimationID 0
Enemy(EnemyNum).SetAnimationLoop True
Enemy(EnemyNum).PlayAnimation 10 * MaxSpeed
Enemy(EnemyNum).SetScale 0.9, 0.9, 0.9
EnemyAttack(EnemyNum) = False
EnemyPos(EnemyNum).x = (XI * 100) + 50
EnemyPos(EnemyNum).Y = 0
EnemyPos(EnemyNum).Z = (YI * 100) + 50
Enemy(EnemyNum).SetPosition EnemyPos(EnemyNum).x, 0, EnemyPos(EnemyNum).Z
EnemyDestination(EnemyNum) = EnemyPos(EnemyNum)
EnemyFoundNewPath(EnemyNum) = False
Enemy(EnemyNum).SetName Format(EnemyNum)
EnemyHealth(EnemyNum) = 100
EnemyNum = EnemyNum + 1
End Select

Select Case BlockNameI
Case Is = "T":
UpSide.AddFloor GetTex("Tex" + TxtName3), XI * 100, YI * 100, (XI * 100) + 100, (YI * 100) + 100, 100
Case Is = "{":
UpSide.AddFloor GetTex("Tex" + TxtName3), (XI * 100) + 50, (YI * 100) + 50, (XI * 100) + 100, (YI * 100) + 100, 100
Case Is = "}":
UpSide.AddFloor GetTex("Tex" + TxtName3), (XI * 100), (YI * 100) + 50, (XI * 100) + 50, (YI * 100) + 100, 100
Case Is = "[":
UpSide.AddFloor GetTex("Tex" + TxtName3), (XI * 100) + 50, (YI * 100), (XI * 100) + 100, (YI * 100) + 50, 100
Case Is = "]":
UpSide.AddFloor GetTex("Tex" + TxtName3), (XI * 100), (YI * 100), (XI * 100) + 50, (YI * 100) + 50, 100
Case Is = "<":
UpSide.AddFloor GetTex("Tex" + TxtName3), (XI * 100) + 50, (YI * 100), (XI * 100) + 100, (YI * 100) + 100, 100
Case Is = ">":
UpSide.AddFloor GetTex("Tex" + TxtName3), (XI * 100), (YI * 100), (XI * 100) + 50, (YI * 100) + 100, 100
Case Is = "v":
UpSide.AddFloor GetTex("Tex" + TxtName3), (XI * 100), (YI * 100), (XI * 100) + 100, (YI * 100) + 50, 100
Case Is = "^":
UpSide.AddFloor GetTex("Tex" + TxtName3), (XI * 100), (YI * 100) + 50, (XI * 100) + 100, (YI * 100) + 100, 100

End Select

Next
Next

Input #1, BlockName
PlayerPos.x = (Int(BlockName) * 100) - 48
Input #1, BlockName
PlayerPos.Z = (Int(BlockName) * 100) - 48

Close #1
Close #3
Close #4
Close #5
Close #6

EnemyNumSurvive = EnemyNum
End Sub

Private Sub LoadTextures()
Set Scene = New TVScene
Set TextureFactory = New TVTextureFactory

Dim TexturesNum As Integer
Dim i As Integer, x As String

Open App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\Textures.txt" For Input As #2
Input #2, TexturesNum
For i = 1 To TexturesNum
Input #2, x
If Right(x, 3) = "jpg" Then
Scene.LoadTexture App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\" + x, , , "Tex" + Left(x, 1)
Else
TextureFactory.LoadTexture App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\" + x, "Tex" + Left(x, 1), , , TV_COLORKEY_BLACK
End If
Next
Close #2

Scene.SetViewFrustum 70, 1000

TextureFactory.LoadTexture App.Path + "\Data\Gfx\GoTo.bmp", "GoTo", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Gfx\GoToL.bmp", "GoToL", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Gfx\Tree1.bmp", "TreeA", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Gfx\Tree2.bmp", "TreeB", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Gfx\NO.bmp", "NO", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Gfx\Health.bmp", "Health", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Gfx\Weapon.bmp", "Weapon", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Gfx\Soor.bmp", "Soor", , , TV_COLORKEY_MAGENTA
TextureFactory.LoadTexture App.Path + "\Data\Gfx\Bomb.bmp", "Bomb", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\Info.jpg", "Info"

Scene.LoadCursor App.Path + "\Data\Gfx\ARROW.BMP", TV_COLORKEY_BLACK, 32, 32
End Sub

Private Sub BuildNodesOfPath()
Set Path = New TVPath
Set AI = New TVAI
Dim x As Integer, Y As Integer

For x = -100 To (ScaleMapX + 1) * 100 Step 100
  For Y = -100 To (ScaleMapX + 1) * 100 Step 100
    AI.AddNode Vector(x + 50, 0, Y + 50)
  Next Y
Next x

GotoBlock.SetCollisionEnable False
UpSide.SetCollisionEnable False
Map.SetCollisionEnable False
Walls.SetCollisionEnable True

AI.SetFindPathParameters 250, 0
AI.CreateAIGraph
End Sub

Private Sub LoadPlayer()
Set Player = New TVActor2
Player.Load App.Path + "\Data\Model\scientist.mdl"
Player.SetAnimationID 13
Player.SetAnimationLoop True
Player.PlayAnimation 10 * MaxSpeed
Player.SetPosition PlayerPos.x, PlayerPos.Y, PlayerPos.Z
End Sub

Private Sub InputKey()
GetCursorPos WindowsMousePosition
InputEngine.SetMousePosition WindowsMousePosition.x, WindowsMousePosition.Y
InputEngine.GetAbsMouseState tmpMouseX, tmpMousey, tmpMouseB1, tmpMouseB2

If tmpMouseB2 <> 0 Then ' cancel command
FoundNewPath = False
Player.SetAnimationID 13
Player.SetAnimationLoop True
Player.PlayAnimation 10 * MaxSpeed
GotoBlock.SetPosition 0, -5, 0
PlayerPos = Player.GetPosition
PlayerDestination = PlayerPos
EndNode = PlayerPos
HurryUp = False
End If

If tmpMouseB1 <> 0 And TimeWait < 1 Then  ' command to go
TimeWait = 25
Map.SetCollisionEnable True
Walls.SetCollisionEnable False
Set CollisionResult = Scene.MousePicking(tmpMouseX, tmpMousey, TV_COLLIDE_MESH, TV_TESTTYPE_BOUNDINGBOX)
If CollisionResult.IsCollision Then

If FoundNewPath = True Then ' if we have to hurry up
HurryUp = True
Else
HurryUp = False
End If

EndNode.x = CollisionResult.GetImpactPoint.x
EndNode.Z = CollisionResult.GetImpactPoint.Z
EndNode.Y = 0
PlayerPos.Y = 0
Map.SetCollisionEnable False
Walls.SetCollisionEnable True

GotoBlock.SetTexture GetTex("GoToL")
GotoBlock.SetPosition EndNode.x, 1, EndNode.Z

If HurryUp = False Then ' walk or run
Player.SetAnimationID 0
Player.SetAnimationLoop True
Player.PlayAnimation 10 * MaxSpeed
Else
Player.SetAnimationID 2
Player.SetAnimationLoop True
Player.PlayAnimation 15 * MaxSpeed
End If

If GetDistance3D(PlayerPos.x, 0, PlayerPos.Z, EndNode.x, 0, EndNode.Z) > 99 And FoundNewPath = False Then

Set Path = New TVPath
FoundNewPath = AI.FindPath(PlayerPos, EndNode, Path)
PlayerPosPoint = 2

If Path.GetNodeCount > 2 Then
GotoBlock.SetTexture GetTex("GoTo")
PlayerDestination = Path.GetNode(PlayerPosPoint)

Else ' there are no path
 TimeWait = 50
 FoundNewPath = False
 Player.SetAnimationID 13
 Player.SetAnimationLoop True
 Player.PlayAnimation 10 * MaxSpeed
 GotoBlock.SetTexture GetTex("GoToL")
End If
End If
End If
Else
If TimeWait > 0 Then TimeWait = TimeWait - 1
End If

' Fire
If PlayerWeapon = True And tmpMouseB2 <> 0 Then
Set CollisionResult = Scene.MousePicking(tmpMouseX, tmpMousey, TV_COLLIDE_ACTOR2, TV_TESTTYPE_ACCURATETESTING)
If CollisionResult.IsCollision = True Then
Dim EnemyIndex As Integer
If IfItNumber(CollisionResult.GetCollisionActor2.GetName) = True Then
EnemyIndex = Int(CollisionResult.GetCollisionActor2.GetName)
If Walls.Collision(Player.GetPosition, Enemy(EnemyIndex).GetPosition, TV_TESTTYPE_ACCURATETESTING) = False Then
If EnemyHealth(EnemyIndex) > 0 Then
EnemyHealth(EnemyIndex) = EnemyHealth(EnemyIndex) - 1
AddBomb Vector(Enemy(EnemyIndex).GetPosition.x, Rnd * 60, Enemy(EnemyIndex).GetPosition.Z), Rnd * 0.3
If EnemyHealth(EnemyIndex) < 1 Then
Enemy(EnemyIndex).SetAnimationID 25 + Int(Rnd * 5)
Enemy(EnemyIndex).SetAnimationLoop False
Enemy(EnemyIndex).PlayAnimation 10 * MaxSpeed
EnemyNumSurvive = EnemyNumSurvive - 1
SFX(1 + CLng(Rnd) + CLng(Rnd)).Play
End If
End If
End If
End If
End If
End If

' Done Mission
If PlayerBombNum > 0 And tmpMouseB2 <> 0 Then
Set CollisionResult = Scene.MousePicking(tmpMouseX, tmpMousey, TV_COLLIDE_ACTOR2, TV_TESTTYPE_ACCURATETESTING)
If CollisionResult.IsCollision = True Then
Dim ItemIndex As Integer
If IfItItem(CollisionResult.GetCollisionActor2.GetName) = True Then
ItemIndex = Int(Right(CollisionResult.GetCollisionActor2.GetName, Len(CollisionResult.GetCollisionActor2.GetName) - 6))
If ObjectDone(ItemIndex) = False Then
ObjectDone(ItemIndex) = True
AddBomb Object(ItemIndex).GetPosition, 2
AddParticle Object(ItemIndex).GetPosition
PlayerBombNum = PlayerBombNum - 1
ObjDoneNum = ObjDoneNum - 1
SFX(4 + CLng(Rnd) + CLng(Rnd)).Play
End If
End If
End If
End If

If InputEngine.IsKeyPressed(TV_KEY_ESCAPE) Then
 EndGame = True
End If
End Sub

Private Sub Movement()
If FoundNewPath = True Then ' there are path , we have to move

If GetDistance3D(PlayerPos.x, 0, PlayerPos.Z, PlayerDestination.x, 0, PlayerDestination.Z) > 5 Then

If HurryUp = True Then ' if player WALK or RUN
PlayerPos = VAdd(Player.GetPosition, VScale(PlayerDirection, 3))
SFX(12).Play
Else
PlayerPos = VAdd(Player.GetPosition, VScale(PlayerDirection, 1))
SFX(11).Play
End If

PlayerDirection = VNormalize(VSubtract(PlayerDestination, Player.GetPosition))
Player.SetPosition PlayerPos.x, PlayerPos.Y, PlayerPos.Z
Player.Lookat Path.GetNode(PlayerPosPoint).x, 0, Path.GetNode(PlayerPosPoint).Z
Player.RotateY 90

Else ' go to next position

PlayerPosPoint = PlayerPosPoint + 1
If Not PlayerPosPoint > Path.GetNodeCount - 1 Then
PlayerDestination = Path.GetNode(PlayerPosPoint)
Else ' finished
FoundNewPath = False
Player.SetAnimationID 13
Player.SetAnimationLoop True
Player.PlayAnimation 10 * MaxSpeed
GotoBlock.SetPosition 0, -5, 0
End If
End If
End If

Scene.SetCamera PlayerPos.x - 50, 250, PlayerPos.Z - 100, PlayerPos.x, 0, PlayerPos.Z
End Sub

Sub Close3D()
Set Map = Nothing
Set Walls = Nothing
Set Path = Nothing
Set AI = Nothing
Set Scene = Nothing
Set Player = Nothing
Set TextureFactory = Nothing
Set TV = Nothing
End Sub

Sub AddTree(Pos As D3DVECTOR, Mode As Boolean)
If Mode = True Then
Set Tree(TreeNum) = Scene.CreateMeshBuilder
Tree(TreeNum).Load3DSMesh App.Path + "\Data\Model\Tree1.3ds"
Tree(TreeNum).SetTexture GetTex("TreeA")
Tree(TreeNum).SetColor RGBA(1, 1, 1, 1)
Tree(TreeNum).ScaleMesh 9, 9, 9
Tree(TreeNum).SetPosition Pos.x, 25, Pos.Z
Tree(TreeNum).SetRotation -90, Rnd * 180, 0
Else
Set Tree(TreeNum) = Scene.CreateMeshBuilder
Tree(TreeNum).Load3DSMesh App.Path + "\Data\Model\Tree2.3ds"
Tree(TreeNum).SetTexture GetTex("TreeB")
Tree(TreeNum).SetColor RGBA(1, 1, 1, 1)
Tree(TreeNum).ScaleMesh 8, 8, 8
Tree(TreeNum).SetPosition Pos.x, 80, Pos.Z
Tree(TreeNum).SetRotation -90, Rnd * 180, 0
End If

Tree(TreeNum).SetCollisionEnable False
TreeNum = TreeNum + 1
End Sub

Sub CollectItem()
Dim i As Integer
For i = 0 To ItemNum - 1
If GetDistance3D(PlayerPos.x, 0, PlayerPos.Z, ItemO(i).GetPosition.x, 0, ItemO(i).GetPosition.Z) < 20 Then

Select Case ItemMode(i)
Case Is = "Health"
PlayerHealth = PlayerHealth + 50
SFX(8).Play

Case Is = "Weapon"
PlayerWeapon = True
SFX(9).Play

Case Is = "Bomb"
PlayerBombNum = PlayerBombNum + 1
SFX(7).Play

End Select

ItemMode(i) = "None"
ItemO(i).SetPosition 0, -1000, 0

End If
Next
End Sub

Sub EnemyMovement()

Map.SetCollisionEnable False
Walls.SetCollisionEnable True

Dim i As Integer
For i = 0 To EnemyNum - 1

If EnemyHealth(i) > 0 Then

If EnemyAttack(i) = False Then

If Walls.Collision(Player.GetPosition, Enemy(i).GetPosition, TV_TESTTYPE_ACCURATETESTING) = False Then
EnemyAttack(i) = True
End If

Else

If EnemyFoundNewPath(i) = False Then
If GetDistance3D(PlayerPos.x, 0, PlayerPos.Z, EnemyPos(i).x, 0, EnemyPos(i).Z) > 99 And (TV.TickCount - WaitForNextFind) > 5000 Then
EnemyFoundNewPath(i) = AI.FindPath(Enemy(i).GetPosition, Player.GetPosition, EnemyPath(i))
EnemyPosPoint(i) = 1
WaitForNextFind = TV.TickCount
If EnemyPath(i).GetNodeCount > 1 Then
EnemyFoundNewPath(i) = True
EnemyDestination(i) = EnemyPath(i).GetNode(EnemyPosPoint(i))
Enemy(i).SetAnimationID 5
Enemy(i).SetAnimationLoop True
Enemy(i).PlayAnimation 15 * MaxSpeed
End If
End If

Else

If GetDistance3D(EnemyDestination(i).x, 0, EnemyDestination(i).Z, EnemyPos(i).x, 0, EnemyPos(i).Z) > 5 And GetDistance3D(PlayerPos.x, 0, PlayerPos.Z, EnemyPos(i).x, 0, EnemyPos(i).Z) > 25 Then
EnemyPos(i) = VAdd(Enemy(i).GetPosition, VScale(EnemyDirection(i), 3))
EnemyDirection(i) = VNormalize(VSubtract(EnemyDestination(i), Enemy(i).GetPosition))
Enemy(i).SetPosition EnemyPos(i).x, EnemyPos(i).Y, EnemyPos(i).Z
Enemy(i).Lookat EnemyPath(i).GetNode(EnemyPosPoint(i)).x, 0, EnemyPath(i).GetNode(EnemyPosPoint(i)).Z
Enemy(i).RotateY 90
SFX(13).Play
Else
EnemyPosPoint(i) = EnemyPosPoint(i) + 1
If Not EnemyPosPoint(i) > EnemyPath(i).GetNodeCount - 1 Then
EnemyDestination(i) = EnemyPath(i).GetNode(EnemyPosPoint(i))
Else
EnemyFoundNewPath(i) = False
Enemy(i).SetAnimationID 0
Enemy(i).SetAnimationLoop True
Enemy(i).PlayAnimation 10 * MaxSpeed
End If
End If

End If

End If

If GetDistance3D(PlayerPos.x, 0, PlayerPos.Z, EnemyPos(i).x, 0, EnemyPos(i).Z) < 30 Then
If Not Enemy(i).GetAnimation = 19 Then
Enemy(i).SetAnimationID 19
Enemy(i).SetAnimationLoop True
Enemy(i).PlayAnimation 10 * MaxSpeed
End If
PlayerHealth = PlayerHealth - 1
Else
If Enemy(i).GetAnimation = 19 Then
Enemy(i).SetAnimationID 0
Enemy(i).SetAnimationLoop True
Enemy(i).PlayAnimation 10 * MaxSpeed
End If
End If

Else

Dim i2 As Integer
For i2 = 0 To EnemyNum - 1
If EnemyHealth(i2) > 0 And GetDistance3D(Enemy(i).GetPosition.x, 0, Enemy(i).GetPosition.Z, Enemy(i2).GetPosition.x, 0, Enemy(i2).GetPosition.Z) < 300 Then
EnemyAttack(i2) = True
End If
Next

End If

Next
End Sub

Sub IntBomb()
On Error Resume Next
Dim i As Integer
For i = 1 To 8
TextureFactory.LoadTexture App.Path + "\Data\Gfx\expl" + Format(i) + ".jpg", "expl" + Format(i), , , TV_COLORKEY_BLACK
Next
For i = 0 To 9
Set Bomb(i) = Scene.CreateBillboard(GetTex("expl1"), -50, 0, -50, 100, 100)
BombFrame(i) = 0
BombIsFree(i) = True
Bomb(i).SetPosition -10000, -10000, -10000
Next
End Sub

Sub IntParticle()
On Error Resume Next
Dim i As Integer
TextureFactory.LoadTexture App.Path + "\Data\Gfx\smoke.bmp", "Smoke", , , TV_COLORKEY_BLACK
For i = 0 To 9
Set Particle(i) = New TVParticleSystem
ParticleIsFree(i) = True
Particle(i).ParticleSize = 100
Particle(i).SetEmittorMoveMode TV_EMITTOR_LERP
Next
End Sub

Sub AddParticle(Pos As D3DVECTOR)
On Error Resume Next
ParticleIsFree(ParticleFree) = False
Particle(ParticleFree).CreateBillboardSystem 100, 0, 100, Pos, 100, GetTex("Smoke")
Particle(ParticleFree).SetParticleAutoGenerateMode True, 1, 0.6, 0.2, 0.4, 0.3, 0, 10, 0, 20, 0.5, 20
Particle(ParticleFree).SetAlphaBlendingMode TV_CHANGE_ALPHA, D3DBLEND_SRCALPHA, D3DBLEND_ONE
Particle(ParticleFree).SetGeneratorSpeed 2
Dim i As Integer
For i = 0 To 9
If ParticleIsFree(i) = True Then
ParticleFree = i
Else
ParticleFree = 0
End If
Next
End Sub

Sub AddBomb(Pos As D3DVECTOR, Size As Single)
On Error Resume Next
BombIsFree(BombFree) = False
Bomb(BombFree).SetPosition Pos.x, Pos.Y, Pos.Z
Bomb(BombFree).ScaleMesh Size, Size, Size
Dim i As Integer
For i = 0 To 9
If BombIsFree(i) = True Then
BombFree = i
Else
BombFree = 0
End If
Next
End Sub

Sub UpdateBomb()
On Error Resume Next
Dim i As Integer
For i = 0 To 9
If BombIsFree(i) = False Then
If BombFrame(i) = 8 Then
BombFrame(i) = 0
BombIsFree(i) = True
Bomb(i).SetPosition -10000, -10000, -10000
Else
BombFrame(i) = BombFrame(i) + 1
Bomb(i).SetTexture GetTex("expl" + Format(BombFrame(i)))
Bomb(i).LookAtPoint Scene.GetCamera.GetPosition
Bomb(i).RotateX 90
End If
End If
Next
End Sub

Sub IntSound()
On Error Resume Next
SoundEngine.Init Me.hWnd
Set Listener = SoundEngine.Get3DListener
Dim i As Integer
Dim MusNum As Integer
Open App.Path + "\Data\Missions\Mission " + Format(MissionNum) + "\music.txt" For Input As #9
Input #9, MusNum
Close #9
SFX(0).Load App.Path + "\Data\Music\" + Format(MusNum) + ".wav"

SFX(0).Loop_ = True
SFX(0).Play
For i = 1 To 25
SFX(i).Load App.Path + "\Data\Sfx\" + Format(i) + ".wav"
SFX(i).Loop_ = False
Next
SFXMenu(0).Stop_
End Sub

Function IfItNumber(Text As String) As Boolean
On Error GoTo err
Dim i As Integer
i = Int(Text)
IfItNumber = True
Exit Function
err:
IfItNumber = False
End Function

Function IfItItem(Text As String) As Boolean
On Error GoTo err
If Left(Text, 6) = "Object" Then
Dim i As Integer
i = Int(Right(Text, Len(Text) - 6))
IfItItem = True
Else
IfItItem = False
End If
Exit Function
err:
IfItItem = False
End Function

Sub IfGameEnd()
On Error Resume Next
If PlayerHealth < 1 Then
SFX(0).Stop_
SFX(14).Play
LoseLVL.Show
End If

If EnemyNumSurvive = 0 And ObjDoneNum = 0 Then
SFX(0).Stop_
SFX(10).Play
WinLVL.Show
Open App.Path + "\Data\Mission.txt" For Output As #1
Write #1, MissionNum + 1
Close #1
End If

Unload Me
End Sub


Private Sub MainLoop()
Dim i As Integer
EndGame = False

Dim TextTexture As Long
Screen2DText.NormalFont_Create "I", "Arial", 16, True, False, False
TextTexture = Screen2DText.TextureFont_Create("font", "Arial", 16, True)

Me.Show

TimeStart = TV.TickCount
Do
DoEvents
TimeNow = (TV.TickCount - TimeStart) \ 1000
TV.Clear
Screen2DImmediate.DRAW_Texture GetTex("Info"), 0, 0, Me.ScaleWidth, Me.ScaleHeight
TV.RenderToScreen
If InputEngine.IsKeyPressed(TV_KEY_SPACE) = True Or TimeNow = 10 Then Exit Do
Loop

Do While EndGame = False
DoEvents

If PlayerHealth < 1 Then
EndGame = True
End If
If EnemyNumSurvive = 0 And ObjDoneNum = 0 Then
EndGame = True
End If

InputKey

Movement

EnemyMovement

CollectItem

UpdateBomb

If GameSpeed > MaxSpeed Then
GameSpeed = 0
TV.Clear

For i = 0 To TreeNum - 1
Tree(i).Render
Next

For i = 0 To ObjNum - 1
Object(i).Render
Next

For i = 0 To ItemNum - 1
ItemO(i).Render
ItemO(i).RotateY 1
Next

For i = 0 To EnemyNum - 1
Enemy(i).Render
Next

For i = 0 To 9
Bomb(i).Render
Next

Map.Render
Walls.Render
Player.Render
GotoBlock.Render

UpSide.SetCollisionEnable True
If UpSide.Collision(Scene.GetCamera.GetPosition, PlayerPos, TV_TESTTYPE_ACCURATETESTING) = False Then
UpSide.SetColor RGBA(1, 1, 1, 1)
Else
UpSide.SetColor RGBA(1, 1, 1, 0.4)
End If
UpSide.SetCollisionEnable False

UpSide.Render

Screen2DText.ACTION_BeginText
Screen2DText.NormalFont_DrawText "Health : " + Format(PlayerHealth), 25, 25, RGBA(1, 0, 0, 1), "I"
Screen2DText.NormalFont_DrawText "Enemies : " + Format(EnemyNumSurvive) + " of " + Format(EnemyNum), 25, 55, RGBA(1, 0, 0, 1), "I"
Screen2DText.NormalFont_DrawText "Missions : " + Format(ObjDoneNum) + " of " + Format(ObjNum), 25, 85, RGBA(1, 0, 0, 1), "I"
For i = 0 To EnemyNum - 1
If EnemyHealth(i) > 0 Then
Screen2DText.TextureFont_DrawBillboardText EnemyHealth(i), EnemyPos(i).x, 80, EnemyPos(i).Z, RGBA(1, 0, 0, 0.4), TextTexture, 5, 5
End If
Next
For i = 0 To ItemNum - 1
Screen2DText.TextureFont_DrawBillboardText ItemMode(i), ItemO(i).GetPosition.x, 80, ItemO(i).GetPosition.Z, RGBA(0, 0, 1, 1), TextTexture, 5, 5
Next
Screen2DText.ACTION_EndText

For i = 0 To 9
If ParticleIsFree(i) = False Then
  Particle(i).RenderParticles
  Particle(i).UpdateParticles TV_UPDATE_RANDOM_GENERATOR
End If
Next

TV.RenderToScreen

Else

GameSpeed = GameSpeed + 1

End If

Loop

Close3D
IfGameEnd
End Sub

