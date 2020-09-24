Attribute VB_Name = "GlobalDims"
' Programmed By [ Zaid Markabi ]
' ___________________________________________________________________________________________________
'|                                                                                                   |\_______________________
'|  ###############        ###         #####   ######                ######    #####                 |                        |\0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'| ##############         #####         ###     ##   ##               ######  #####                  |      Zaid Markabi      |=\ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|         ####          ### ###        ###     ##    ##              ##  ## ##  ##                  |                        |==\0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'|       ###            ###   ###       ###     ##     ##    #####    ##   ###   ##                  | zaidmarkabi@yahoo.com  |===\ 1 __________________________________  0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 1 0 0 0 1
'|     ###             ###########      ###     ##     ##   ####      ##    #    ##                  |                        |====|>| Development For Our Digital Life | 1 1 0 0 1 1 1 0 1 0 0 1 0 0 0 1 1 0 1 0
'|   ###              #############     ###     ##    ##              ##         ##      A R K A B I | VisualBasic Programmer |===/ 1|__________________________________| 0 1 1 0 1 0 0 0 1 1 1 0 1 0 1 1 0 1 0 0
'| ##############    ###         ###    ###     ##   ##               ##         ##     ############ |                        |==/0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'| ###############   ###         ###   #####   ######                ####       ####   ### 2009 ###  |Arabic Syrian Programmer|=/ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|                                                                                    ############   | _______________________|/0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'|___________________________________________________________________________________________________|/
' Em@il Me Please : zaidmarkabi@yahoo.com
' I hope to hear from you, my mate !
'
' _________________________________________________
'| THIS PROGRAM IS FREEWARE. ANY USE IN COMMERCIAL |
'| APPLICATIONS WITHOUT WRITTEN PERMISSION BY THE  |
'| AUTHOR IS PROHIBITED.                           |
'|_________________________________________________|
'
' COMMANDOS - Industrial Quality 3D
'
' This demo was developed using Visual Basic 6 , DirectX8 and TrueVision3D .
' You must have a 3D Graphic Accelerator Card and Sound card to play it.
'
' FIGHT COMBAT features include mesh animation, sizable viewport, true 3D sound, complex physics, fighter ai and more.
' System requirements: Windows 95/98/00/ME/XP/Vista, Microsoft DirectX8, Soundcard
'
' ENJOY !

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Global TV As New TVEngine
Global Scene As New TVScene
Global InputEngine As New TVInputEngine
Global EndGame As Boolean
Global TextureFactory As New TVTextureFactory
Global Screen2DText As TVScreen2DText
Global Screen2DImmediate As TVScreen2DImmediate

Global TimeStart As Long
Global TimeNow As Long

Global AI As TVAI
Global Path As TVPath
Global EndNode As D3DVECTOR
Global CollisionResult As TVCollisionResult

Global Bomb(9) As TVMesh
Global BombFrame(9) As Integer
Global BombIsFree(9) As Boolean
Global BombFree As Integer

Global Particle(9) As TVParticleSystem
Global ParticleIsFree(9) As Boolean
Global ParticleFree As Integer

Global Map As TVMesh
Global UpSide As TVMesh
Global Tree(99) As TVMesh
Global TreeNum As Integer
Global Object(99) As TVActor2
Global ObjectDone(99) As Boolean
Global ObjNum As Integer
Global ObjDoneNum As Integer
Global ItemO(99) As TVMesh
Global ItemMode(99) As String
Global ItemNum As Integer
Global ScaleMapX As Integer
Global ScaleMapY As Integer
Global Walls As TVMesh
Global MissionNum As Integer

Global GotoBlock As TVMesh
Global TimeWait As Integer

Global SoundEngine As New TVSoundEngine
Global Listener As TVListener
Global SFX(0 To 25) As New TVSoundWave3D
Global SFXMenu(0 To 5) As New TVSoundWave3D

Global Player As TVActor2
Global PlayerPos As D3DVECTOR
Global PlayerDestination As D3DVECTOR
Global PlayerDirection As D3DVECTOR
Global PlayerPosPoint As Integer
Global FoundNewPath As Boolean
Global HurryUp As Boolean

Global Enemy(99) As TVActor2
Global EnemyAttack(99) As Boolean
Global EnemyPos(99) As D3DVECTOR
Global EnemyDestination(99) As D3DVECTOR
Global EnemyDirection(99) As D3DVECTOR
Global EnemyFoundNewPath(99) As Boolean
Global EnemyPath(99) As TVPath
Global EnemyPosPoint(99) As Integer
Global EnemyHealth(99) As Integer
Global WaitForNextFind As Long
Global EnemyNum As Integer
Global EnemyNumSurvive As Integer

Global tmpMouseX As Long
Global tmpMousey As Long
Global tmpMouseB1 As Integer
Global tmpMouseB2 As Integer
Global WindowsMousePosition As POINTAPI

Global PlayerHealth As Integer
Global PlayerWeapon As Boolean
Global PlayerBombNum As Integer

Global GameSpeed As Integer
Global MaxSpeed As Integer

