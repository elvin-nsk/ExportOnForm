Attribute VB_Name = "ExportOnForm"
'===============================================================================
'   Макрос          : ExportOnForm
'   Версия          : 2022.04.26
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "ExportOnForm"

'===============================================================================

Private Const PlaceholderName As String = "*placeholder"

'===============================================================================

Sub Start()
    If RELEASE Then On Error GoTo Catch
    
    Dim Source As InputData
    Set Source = InputData.GetDocumentOrPage
    If Source.IsError Then Exit Sub
    
    lib_elvin.BoostStart Optimize:=RELEASE
    
    Main Source, Config.Load
    
Finally:
    lib_elvin.BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

'===============================================================================

Private Sub Main(ByVal Source As InputData, ByVal Cfg As Config)
    Dim ExportFile As IFileSpec
    Set ExportFile = GetExportFolderOrThrow(Source.Document, Cfg)
    Dim TempShapes As ShapeRange
    Set TempShapes = CollectShapesOnEnabledLayers(Source.Page)
    ThrowIf TempShapes.Count = 0, _
            "Страница пустая или все объекты на закрытых для печати слоях."
    Dim FormDoc As Document
    Set FormDoc = OpenDocumentAsCopy(Cfg.FormFile)
    Dim Placeholder As Shape
    Set Placeholder = GetPlaceholderOrThrow(FormDoc.ActivePage)
    Dim TempDoc As Document
    Set TempDoc = TempShapes.CreateDocumentFrom(RELEASE)
    Dim TempBitmap As Shape
    Set TempBitmap = RasterizeToFitPlaceholder( _
                         TempDoc.ActivePage, _
                         Placeholder.BoundingBox, , , Cfg.ExportResolution _
                     )
    Dim TempFile As String
    TempFile = SaveToTemp(TempBitmap)
    TempDoc.Close
    Import(TempFile, Placeholder.Layer) _
        .SetPositionEx cdrCenter, Placeholder.CenterX, Placeholder.CenterY
    ExportFile.Name = Source.Document.Name
    ExportFile.Ext = "jpg"
    ExportPageToJPEG FormDoc.ActivePage, ExportFile, _
                     Cfg.ExportResolution, Cfg.ExportJpegCompression
    FormDoc.Close
    VBA.Kill TempFile
End Sub

Private Function GetExportFolderOrThrow( _
                     ByVal Doc As Document, _
                     ByVal Cfg As Config _
                 ) As IFileSpec
    Dim Folder As String
    With Cfg
        Folder = .ExportFolder
        If Folder = "" Then Folder = Doc.FilePath
        If Folder = "" Then Folder = .ExportFallbackFolder
    End With
    If Folder = "" Then
        Throw "Ошибка пути экспорта."
    Else
        Set GetExportFolderOrThrow = FileSpec.Create
        GetExportFolderOrThrow.Path = Folder
    End If
End Function

Private Function GetPlaceholderOrThrow(ByVal Page As Page) As Shape
    Set GetPlaceholderOrThrow = Page.Shapes.FindShape(Name:=PlaceholderName)
    ThrowIf GetPlaceholderOrThrow Is Nothing, _
            "Не найден *placeholder на бланке."
End Function

Private Function CollectShapesOnEnabledLayers(ByVal Page As Page) As ShapeRange
    Set CollectShapesOnEnabledLayers = CreateShapeRange
    Dim Layer As Layer
    For Each Layer In Page.Layers
        If (Not Layer.IsSpecialLayer) And Layer.Printable Then
            CollectShapesOnEnabledLayers.AddRange Layer.Shapes.All
        End If
    Next Layer
End Function

Private Function RasterizeToFitPlaceholder( _
                    ByVal Page As Page, _
                    ByVal Placeholder As Rect, _
                    Optional ByVal Space As Double = 0, _
                    Optional ByVal Rotate As Boolean = False, _
                    Optional ByVal TargetDPI As Long = 150 _
                ) As Shape
    Dim ShapeRange As ShapeRange
    Set ShapeRange = Page.Shapes.All
    Dim PlaceholderSize As Rect
    Set PlaceholderSize = Placeholder.GetCopy
    Dim Ratio As Double
    With PlaceholderSize
        .Width = .Width - Space
        .Height = .Height - Space
        If Rotate Then
            Ratio = .Width / ShapeRange.SizeHeight
        Else
            Ratio = .Width / ShapeRange.SizeWidth
        End If
    End With
    Dim Result As Shape
    Set Result = _
            ShapeRange.ConvertToBitmapEx( _
                cdrRGBColorImage, , False, _
                VBA.CLng(TargetDPI * Ratio), _
                cdrNormalAntiAliasing, True, True, 100 _
            )
    If Rotate Then RasterizeToFitPlaceholder.Rotate -90
    Result.SetSize PlaceholderSize.Width
    With PlaceholderSize
        Result.SetBoundingBox .x, .y, .Width, .Height, True
    End With
    Result.PixelAlignedRendering = True
    Set RasterizeToFitPlaceholder = Result
End Function

Private Function SaveToTemp(ByVal BitmapShape As Shape) As String
    SaveToTemp = lib_elvin.GetTempFile
    BitmapShape.Bitmap.SaveAs(SaveToTemp, cdrTIFF, cdrCompressionLZW).Finish
End Function

Private Function Import( _
                     ByVal TempFile As String, _
                     ByVal ToLayer As Layer _
                 ) As Shape
    ToLayer.Import TempFile, cdrTIFF
    Set Import = ToLayer.Page.Parent.Parent.Selection
End Function

Private Sub ExportPageToJPEG( _
                ByVal Page As Page, _
                ByVal File As IFileSpec, _
                ByVal DPI As Long, _
                Optional ByVal Compression As Long = 25, _
                Optional ByVal AntiAliasingType As cdrAntiAliasingType _
                    = cdrNormalAntiAliasing, _
                Optional ByVal ImageType As cdrImageType = cdrRGBColorImage _
            )
    With Page.Parent.Parent.ExportBitmap( _
             File.ToString, cdrJPEG, cdrCurrentPage, ImageType, _
             0, 0, DPI, DPI, AntiAliasingType, _
             False, False, True, False, cdrCompressionNone _
         )
        .Compression = 25
        .Optimized = True
        .Finish
    End With
End Sub

'===============================================================================
' # тесты

Private Sub testSomething()
'
End Sub
