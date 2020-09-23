VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Messages"
   ClientHeight    =   3390
   ClientLeft      =   3360
   ClientTop       =   2460
   ClientWidth     =   8295
   Icon            =   "frmConfigura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8295
   Begin Messages1.ShellIcon ShellIcon1 
      Left            =   5520
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      ToolTipText     =   "Messages"
      Icon            =   "frmConfigura.frx":2372
      Visible         =   -1  'True
      SysMenu         =   0   'False
   End
   Begin VB.CheckBox chkMinimizeToTray 
      Caption         =   "Minimize on load"
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CheckBox chkLoadStartup 
      Caption         =   "Load at startup"
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdSaveConfig 
      Caption         =   "Save Config"
      Height          =   435
      Left            =   6360
      TabIndex        =   10
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   6240
      TabIndex        =   15
      Top             =   240
      Width           =   1815
      Begin VB.CheckBox chkUpdate 
         Caption         =   "Update"
         Height          =   215
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1065
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1320
         Tag             =   "0"
         Top             =   840
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.OptionButton optManual 
         Caption         =   "Manual"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Automatic"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Minutes:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdShowMessage 
      Caption         =   "Show Message"
      Height          =   435
      Left            =   3720
      TabIndex        =   11
      Top             =   2760
      Width           =   1635
   End
   Begin VB.CheckBox chkSquare 
      Caption         =   "Square (default)"
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   2100
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Yellow (default)"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Value           =   -1  'True
      Width           =   2160
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Blue"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1960
      Width           =   1560
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Red"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2240
      Width           =   1560
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Green"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   1560
   End
   Begin VB.CheckBox chkFade 
      Caption         =   "Fade"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Manual close"
      Height          =   390
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   2505
      Left            =   2655
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "frmConfigura.frx":46F4
      Top             =   150
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Auto close"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message"
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Caption         =   "Color"
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuShowManual 
         Caption         =   "Show message (manual)"
      End
      Begin VB.Menu mnuShowAuto 
         Caption         =   "Show message (automatic)"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "&Config"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'******************************************************************************
' Este programa fue diseñado originalmente para mostrar automáticamnete
' versículos biblícos. Puede editarse el archivo mensajes.txt para mostrar
' cualquier mensaje(s) que se quiera a intervalos de tiempo programados.
'
' This application was designed to show Bible verses. The file messages.txt can
' be edited to add or enter your own messages. Through interface you can define
' the time interval for display the messages.
'
'                   Basado en los programas / Based on
'                Alert Window Developed by Morgan Haueisen
'             Alert Window desarrollado por Morgan Haueisen
'
'              Daily Bible Verse Developed by Bruce Bowman
'            Daily Bible Verse desarrollado por Bruce Bowman
'
'                   DisplayMess Developed by Ernesto F.
'                 DisplayMess desarrollado por Ernesto F.
'
'      Add SysTray-Icon +++ Change Wallpaper Developed by Florian Egel
'     Add SysTray-Icon +++ Change Wallpaper desarrollado por Florian Egel
'
'                   AutoForm Developed by Ernesto Chapon
'                 AutoForm desarrollado por Ernesto Chapon
'
'       Routine for Read and Write INI file Developed by Niklas Spångberg
'    Rutina para Leer y Escribir archivo INI desarollada por Niklas Spångberg
'
'              CS Wallpaper Changer Developed by Shane Croft
'           CS Wallpaper Changer desarrollador por Shane Croft
'******************************************************************************

'Definir colección en memoria de la base de datos de versículos.
'The in-memory database of tips.
Dim Tips As New Collection

'Nombre del archivo de versículos
'Name of tips file
Const TIP_FILE = "messages.txt"

'Index en la colección de versículos del versículo a ser mostrado.
'Index in collection of tip currently being displayed.
Dim CurrentTip As Long

Private m_iBackColor As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Const ERROR_SUCCESS = 0&
Const REG_SZ = 1 ' Unicode nul terminated String
Const REG_DWORD = 4 ' 32-bit number

Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Private Sub chkLoadStartup_Click()
    
' Verifica Ejecutar al iniciar y añade o remueve entrada en registro
' Verify Load on StartUp and writes or delete Registry entry
    If chkLoadStartup.Value = 1 Then
        Call AddToRun("Mensajes", App.Path & "\" & App.EXEName & ".exe")
    End If
    If chkLoadStartup.Value = 0 Then
        Call RemoveFromRun("Mensajes")
    End If

End Sub

Private Sub chkUpdate_Click()
    
' Activa o desactiva la actualización automática
    Timer1.Tag = 0
    Timer1.Enabled = chkUpdate

End Sub

Private Sub cmdSaveConfig_Click()
    Dim X As Boolean
    
    ' Guardar desvanecer / Save Fade
    X = WriteIni("Main", "chkFade", chkFade.Value)
    ' Guardar rectángulo / Save Square
    X = WriteIni("Main", "chkSquare", chkSquare.Value)
    ' Guardar actualizar / Save Update
    X = WriteIni("Main", "chkUpdate", chkUpdate.Value)
    ' Guardar minutos / Save Minutes
    X = WriteIni("Main", "Text2", Text2.Text)
    ' Guardar cerrar automáticamente / Save Automatic close
    X = WriteIni("Main", "optAuto", optAuto.Value)
    ' Guardar cerrar manual / Save manual close
    X = WriteIni("Main", "optManual", optManual.Value)
    ' Guardar ejecutar al inicio / Save Run at startup
    X = WriteIni("Main", "chkLoadStartup", chkLoadStartup.Value)
    ' Guardar minimizar al iniciar / Save mimize to tray on load
    X = WriteIni("Main", "chkMinimizeToTray", chkMinimizeToTray.Value)

End Sub

Private Sub Command1_Click()
  
  Dim AlertWindow As frmAlertWindow
  Dim SMessage As String
  
    Set AlertWindow = New frmAlertWindow
    SMessage = Text1.Text & vbNewLine & vbNewLine & Format(Time, "Medium Time")
            
    AlertWindow.DisplayMessage SMessage, 15, _
        CBool(chkFade.Value), , CBool(chkSquare.Value), m_iBackColor, sMess
        
End Sub

Private Sub Command2_Click()
  Dim AlertWindow As frmAlertWindow
  Dim SMessage As String
    
    Set AlertWindow = New frmAlertWindow
    SMessage = Text1.Text & vbNewLine & vbNewLine & Format(Time, "Medium Time")
    
    AlertWindow.DisplayMessage SMessage, 0, _
        CBool(chkFade.Value), , CBool(chkSquare.Value), m_iBackColor, sAlert

End Sub

Private Sub cmdShowMessage_Click()
  Static bShowClose As Boolean
  Dim Frm As Form
  Dim SMessage As String
  
    If Not bShowClose Then
    SMessage = Text1.Text & vbNewLine & vbNewLine & Format(Time, "Medium Time")
        frmAlertWindow.DisplayMessage SMessage, -1, _
            CBool(chkFade.Value), False, CBool(chkSquare.Value), m_iBackColor, sAlert
        
        bShowClose = True
        cmdShowMessage.Caption = "Close message"
    Else
        bShowClose = False
        cmdShowMessage.Caption = "Show message"
        ' Esto son el necesario sí está permitiendo que más de 1 copia se
        ' muestra al mismo tiempo; de otra forma use frmAlertWindow.CloseActivate = True
        ' This is only necessary if you are allowing more than 1 copy to be
        ' shown at the same time; else use frmAlertWindow.CloseActivate = True
        For Each Frm In Forms
            If Frm.Name = "frmAlertWindow" Then
                Frm.CloseActivate = True
            End If
        Next Frm
    End If

End Sub

Private Sub ReloadTip()

Randomize
 'Lee el archivo de mensajes y muestra un mensaje al azar.
 'Read in the tips file and display a tip at random.
   If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        Text1.Text = "El archivo " & TIP_FILE & " no se encontró." & vbCrLf & vbCrLf & _
           "Cree un archivo llamado " & TIP_FILE & " usando el Block de Notas con 1 mensaje por línea. " & _
           "Guárdelo en el mismo directorio que la aplicación. "
    End If

' For english developers, uncomment this and comment above
   'If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        'Text1.Text = "The file " & TIP_FILE & " not found." & vbCrLf & vbCrLf & _
           '"Crate a text file and name it as " & TIP_FILE & " using Notepad, with only a tip in every line. " & _
           '"Save the file in the same location of the application."
    'End If

    m_iBackColor = &HC0FFFF

End Sub

Private Sub Form_Resize()
    
    If WindowState = 1 Then
        Hide
    Else
        Show
        'Actualiza el mensaje a mostrar
        'Reload tip
        ReloadTip
    End If
End Sub

Private Sub mnuAbout_Click()
  
  Dim AlertWindow As frmAlertWindow
  Dim SMessage As String
    
  SMessage = "Messages 1.0" & vbNewLine & vbNewLine & "Developed by" & vbNewLine & "DataFox" & vbNewLine & "www.datafox.com"
    
    Set AlertWindow = New frmAlertWindow
    SMessage = SMessage & vbNewLine & vbNewLine & Format(Time, "Medium Time")
    
    AlertWindow.DisplayMessage SMessage, 0, _
        CBool(chkFade.Value), , CBool(chkSquare.Value), m_iBackColor, sAlert

End Sub

Private Sub mnuConfig_Click()
    WindowState = 0: Show: AppActivate Caption
End Sub

Private Sub mnuExit_Click()
    Unload frmConfigura
End Sub

Private Sub mnuShowAuto_Click()
    Command1_Click
End Sub

Private Sub mnuShowManual_Click()
    Command2_Click
End Sub

Private Sub optColor_Click(Index As Integer)
    Select Case Index
    Case 0 ' Amarillo / Yellow
        m_iBackColor = &HC0FFFF
    Case 1 ' Azul / Blue
        m_iBackColor = RGB(160, 195, 255)
    Case 2 ' Rojo/ Red
        m_iBackColor = RGB(255, 200, 200)
    Case 3 ' Verde / Green
        m_iBackColor = RGB(200, 255, 200)
    End Select

End Sub

Private Sub Form_Load()

On Error Resume Next
    
' No ejecutar más de una instancia del programa / Don't let more than one copy run
    If App.PrevInstance = True Then End
    
    Randomize
    
 'Lee el archivo de mensajes y muestra un mensaje al azar.
 'Read in the tips file and display a tip at random.
   'If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        'Text1.Text = "El archivo " & TIP_FILE & " no se encontró." & vbCrLf & vbCrLf & _
           '"Cree un archivo llamado " & TIP_FILE & " usando el Block de Notas con 1 mensaje por línea. " & _
           '"Guárdelo en el mismo directorio que la aplicación. "
    'End If

' For english developers, uncomment this and comment above
   If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        Text1.Text = "The file " & TIP_FILE & " not found." & vbCrLf & vbCrLf & _
           "Crate a text file and name it as " & TIP_FILE & " using Notepad, with only a tip in every line. " & _
           "Save the file in the same location of the application."
    End If
    
' Leer archivo INI de configuración y cargar valores
' Read configuration INI file and load values
    With Me
        .chkFade = ReadIni("Main", "chkFade")
        .chkSquare = ReadIni("Main", "chkSquare")
        .chkUpdate = ReadIni("Main", "chkUpdate")
        .Text2 = ReadIni("Main", "Text2")
        .optAuto = ReadIni("Main", "optAuto")
        .optManual = ReadIni("Main", "optManual")
        .chkLoadStartup = ReadIni("Main", "chkLoadStartup")
        .chkMinimizeToTray = ReadIni("Main", "chkMinimizeToTray")
    End With
    
' Si está activo minimiza el formulario al iniciar
' If active, mimize to tray on load
    If chkMinimizeToTray.Value = 1 Then
        Me.Hide
    End If

' Definir color predeterminado / Default background color
    m_iBackColor = &HC0FFFF
    Command1_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

'Elimina el icono de la bandeja y limpia la memoria
'Removes icon from tray and clean up's memory
    ShellIcon1.Visible = False
    Set frmConfig = Nothing

End Sub

Private Sub DoNextTip()

    ' Seleccionar un mensaje al azar
    ' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' O puede ir en orden de los mensajes
    ' Or, you could cycle through the Tips in order
    
'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Mostrar.
    ' Show it.
    DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   'Cada mensaje leído del archivo / Each tip read in from file.
    Dim InFile As Integer   'Descripción del archivo / Descriptor for file.
    
    'Obtener la siguiente descripción del archivo
    'Obtain the next free file descriptor.
    InFile = FreeFile
    
    'Asegurarse que se ha especificado un archivo.
    'Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    'Asegurarse que existe el archivo antes de tratar de abrirlo.
    'Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    'Leer la colección desde el archivo de texto.
    'Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    'Mostrar un mensaje al azar.
    'Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Public Sub DisplayCurrentTip()

'Mostrar mensaje
'Show tip
    
    If Tips.Count > 0 Then
        Text1.Text = Tips.Item(CurrentTip)
    End If

End Sub

Private Sub ShellIcon1_Click(Button As Integer)

    PopupMenu mnuPopup

End Sub

Private Sub ShellIcon1_DblClick(Button As Integer)

' Icono en tray / Icon on tray
    If Button = 1 Then WindowState = 0: Show: AppActivate Caption

End Sub

Private Sub Text2_Validate(Cancel As Boolean)

' verifica un valor correcto en tiempo / Validate for a correct value in time
    Dim Text As String, Pos As Long
    Text = Val(Text2)
    Pos = InStr(Text, ",")
    If Pos Then Text = left(Text, Pos - 1) & "." & Mid(Text, Pos + 1)
    Text2 = Text

End Sub

Private Sub Timer1_Timer()
    
' Muestra mensaje / Show tip
    On Error Resume Next
    
    If Val(Text2) <= 0 Then Exit Sub
    
    Timer1.Tag = Timer1.Tag + 1
        
    If Timer1.Tag >= Int(Val(Text2) * 60) Then
        Timer1.Tag = 0
        If optManual.Value = True Then
            ReloadTip
            Command2_Click
        Else
            ReloadTip
            Command1_Click
        End If
    End If

End Sub

Public Sub AddToRun(ProgramName As String, FileToRun As String)
    
' Escribe una entrada en el registro para 'Ejecutar al iniciar'
' Add a program to the 'Run at Startup' registry keys
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, FileToRun)

End Sub

Public Sub RemoveFromRun(ProgramName As String)
    
' Elimina el programa de la entrada de registro 'Ejecutar al iniciar'
' Remove a program from the 'Run at Startup' registry keys
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName)

End Sub

Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strdata As String)

'EXAMPLE:
'Call savestring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String", text1.text)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)

End Sub

Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)

'EXAMPLE:
'Call DeleteValue(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
    Dim keyhand As Long
    Dim r As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

