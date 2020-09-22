Attribute VB_Name = "ESSSinf"
'ESSS

'Multi-form variables and specialized functions

Option Explicit
Public StarSize As Integer 'Star size
Public StarColor As Long 'Star color
Public PlanetNumber As Byte 'Number of planets in solar system
Public SSAsteroid As Integer, SSComet As Integer 'asteroid flag, comet flag
Public FileName As String 'Name of file to save
Public Age As Single 'Age of system
Public OrbitS As Byte 'Orbital Stability
Public LifeS As Byte 'Life situation
Public PlanetRad As Integer 'General planet radius
Public StartS As Byte 'start the solar system flag
Public Px(1 To 10) As Integer 'Global to save planet x locations
Public PRad(1 To 10) As Integer 'Global to save planet radii
Public PD(1 To 10) As Byte 'Global to save planet direction
Public PO(1 To 10) As Integer 'Global to save planet orbital radius
Public Py(1 To 10) As Single 'Global to save planet y locations
Public POA(1 To 10) As Integer 'Global to save planet Orbital angle
Public SelectedPlanet As Byte 'transfer the n of planet you clicked to editer
Public PlanetType(1 To 10) As String 'type of planet
Public OrbitalPlane(1 To 10) As Integer 'the angle at which a planet orbits
Public PlanetColor(1 To 10) As Long 'the color of the planet
Public PlanetAtmosphereColor(1 To 10) As Long 'the color of the planet's atmosphere
Public PlanetName(1 To 10) As String 'name of planet
Public StarName As String 'name of the system
Public CanSave As Byte 'determines if you can save or not
Public MassAugment(1 To 10) As Integer 'augments radius according to MassAugment/100
Public MagLevel(1 To 10) As Integer 'level of magnitivity
Public CometX As Single, CometY As Single, CometRadius As Integer, CometDir As Byte, CometName As String 'Comet variables
Public GroundElement1(1 To 10) As String, GroundElement2(1 To 10) As String, GroundElement3(1 To 10) As String 'all elements for non-atmosphereic parts of the planet
Public AirElement1(1 To 10) As String, AirElement2(1 To 10) As String, AirElement3(1 To 10) As String 'All elements for atmosphere of planet
Public PlanetTheme(1 To 10) As String 'the theme of the planet

Sub Main()
    'you can't run this program twice at the same time
    If App.PrevInstance = True Then End
    TopToolBarfrm.Show
End Sub

Public Function FormatPath(sPath As String) As String
    'this function gets your path to use
    FormatPath = IIf(Right$(sPath, 1) <> "\", sPath & "\", sPath)
End Function
                                                                                                        'EÞs¡lon Solar System Simulator
'ESSS                                                                                                        'Made By ßrian Adriance
