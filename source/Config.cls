VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Config v2022.05.22
' хранение, загрузка и сохранение настроек
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================

Private Const iniFileName As String = "elvin_" & APP_NAME & ".ini"

'===============================================================================

Public FormFile As String
Public ExportResolution As Long
Public ExportJpegCompression As Long
Public ExportFolder As String
Public ExportFallbackFolder As String

'-------------------------------------------------------------------------------

Private Sub Assign(ByVal Action As iniAction)

    With iniFile.Create
    
        iniFilePrepare .Self, Action
        
        '.Section = "DEFAULT"
        .KeyVar "FormFile", FormFile
        .KeyVar "ExportResolution", ExportResolution, 300
        .KeyVar "ExportJpegCompression", ExportJpegCompression, 25
        .KeyVar "ExportFolder", ExportFolder
        .KeyVar "ExportFallbackFolder", ExportFallbackFolder, "c:\temp"
    
    End With
    
End Sub

'-------------------------------------------------------------------------------

'Опционально. Создаёт ini-файл заданного содержания (напр. с комментами).
'Если не нужно - заменить блок Case Else в iniFilePrepare на
'.File = GMSManager.UserGMSPath & iniFileName
'If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
Private Sub MakeDefaultFile(ByVal File As String)

    Dim Content As String
    
    Content = Content & "; кодировка Windows-1251" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "[DEFAULT]" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; путь к файлу бланка" & vbCrLf
    Content = Content & "FormFile=" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; разрешение экспорта" & vbCrLf
    Content = Content & "ExportResolution=300" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; сжатие джпега (0-100, меньше = выше качество)" & vbCrLf
    Content = Content & "ExportJpegCompression=25" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; папка для экспорта (если пусто - в папку документа)" & vbCrLf
    Content = Content & "ExportFolder=" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; папка для экспорта," & vbCrLf
    Content = Content & "; если текущий документ не имеет папки (новый, не сохранялся)" & vbCrLf
    Content = Content & "ExportFallbackFolder=c:\temp" & vbCrLf
    
    SaveStrToFile Content, File
    
End Sub

'===============================================================================
' # Constructors

Public Function Load() As Config
    Set Load = New Config
    Load.ForceLoad
End Function

Private Sub Class_Terminate()
    ForceSave
End Sub

'===============================================================================
' # Public

Public Sub ForceLoad()
    Assign iniRead
End Sub

Public Sub ForceSave()
    Assign iniWrite
End Sub

Public Property Get Self() As Config
    Set Self = Me
End Property

'===============================================================================

Private Sub iniFilePrepare(ByVal iniFile As iniFile, ByVal Action As iniAction)

    Dim iniFilePath As String

    With iniFile
        Select Case True
        Case FileExists(GMSManager.UserGMSPath & iniFileName)
            .File = Application.GMSManager.UserGMSPath & iniFileName
            .Action = Action
        Case FileExists(GMSManager.GMSPath & iniFileName)
            .File = Application.GMSManager.GMSPath & iniFileName
            .Action = Action
        Case Else
            .File = GMSManager.UserGMSPath & iniFileName
            MakeDefaultFile .File
            If Action = iniRead Then
                .Action = iniRead
            Else
                .Action = iniWrite
            End If
        End Select
    End With

End Sub

Public Function FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Function
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Public Sub SaveStrToFile( _
               ByVal Content As String, _
               ByVal File As String _
           )
    Dim FileNum As Long
    FileNum = VBA.FreeFile
    Open File For Output Access Write As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub
