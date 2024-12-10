Imports System.IO
Imports DriveWorks.Components
Imports DriveWorks.Components.Tasks
Imports DriveWorks.Reporting
Imports DriveWorks.SolidWorks
Imports DriveWorks.SolidWorks.Generation
Imports DriveWorks.SolidWorks.Generation.Proxies
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

<GenerationTask("Add Dimension to Drawing View",
                "Adds a reference dimension between two named entities in a drawing view",
                "embedded://AddDimensionToDrawingViewPoC.AddDimToView.png",
                "Sherpa Tasks",
                GenerationTaskScope.Drawings,
                ComponentTaskSequenceLocation.Before Or ComponentTaskSequenceLocation.After)>
Public Class AddDimToView
    Inherits GenerationTask

    Private Const TASK_CLASS = "Model Generation Task"
    Private Const TASK_NAME As String = "Add Dimension to Drawing View"

    Private Const FIRST_DIM_REFERENCE_NAME_PARAM_NAME As String = "FirstDimReference"
    Private Const SECOND_DIM_REFERENCE_NAME_PARAM_NAME As String = "SecondDimReference"
    Private Const DIM_TYPE_PARAM_NAME As String = "Aligned"
    Private Const VIEW_NAME_PARAM_NAME As String = "ViewName"
    Private Const DIM_X_LOCATION As String = "DimXLocation"
    Private Const DIM_Y_LOCATION As String = "DimYLocation"


    Public Overrides ReadOnly Property Parameters As ComponentTaskParameterInfo()
        Get
            ' Declare any rules that can be written for this task
            Return New ComponentTaskParameterInfo() {
                        New ComponentTaskParameterInfo(FIRST_DIM_REFERENCE_NAME_PARAM_NAME,
                                                       "First Entity",
                                                       "First named entity to dimension between ex. 'LeftPlane@Module-1@TopLevelAsm'"),
                        New ComponentTaskParameterInfo(SECOND_DIM_REFERENCE_NAME_PARAM_NAME,
                                                       "Second Entity",
                                                       "Second named entity to dimension between ex. 'RightPlane@Module-2@TopLevelAsm'"),
                        New ComponentTaskParameterInfo(DIM_TYPE_PARAM_NAME,
                                                       "Dimension Type",
                                                       "Dimension type (choose from 'Horizontal', 'Vertical', or 'Aligned')"),
                        New ComponentTaskParameterInfo(VIEW_NAME_PARAM_NAME,
                                                       "View Name",
                                                       "Name of the drawing view ex. 'ElevationView'"),
                        New ComponentTaskParameterInfo(DIM_X_LOCATION,
                                                       "X Location of Dimension (m)",
                                                       "Horizontal location to place the dimension text in meters from the bottom left corner of the drawing sheet"),
                        New ComponentTaskParameterInfo(DIM_Y_LOCATION,
                                                       "Y Location of Dimension (m)",
                                                       "Vertical location to place the dimension text in meters from the bottom left corner of the drawing sheet")
                   }
        End Get
    End Property

    Public Overrides Sub Execute(model As SldWorksModelProxy, component As ReleasedComponent, generationSettings As GenerationSettings)

        Me.Report.BeginProcess(TASK_CLASS, TASK_NAME, "")

        Dim firstRefName = String.Empty
        If Not Me.Data.TryGetParameterValue(FIRST_DIM_REFERENCE_NAME_PARAM_NAME, firstRefName) Then
            SetExecutionResult(TaskExecutionResult.Failed, "No / Invalid value provided for the 'First Entity' rule.")
            Me.Report.EndProcess()
            Return
        End If

        Dim secondRefName = String.Empty
        If Not Me.Data.TryGetParameterValue(SECOND_DIM_REFERENCE_NAME_PARAM_NAME, secondRefName) Then
            SetExecutionResult(TaskExecutionResult.Failed, "No / Invalid value provided for the 'Second Entity' rule.")
            Me.Report.EndProcess()
            Return
        End If


        Dim viewName = String.Empty
        If Not Me.Data.TryGetParameterValue(VIEW_NAME_PARAM_NAME, viewName) Then
            SetExecutionResult(TaskExecutionResult.Failed, "No / Invalid value provided for the 'Drawing View' rule.")
            Me.Report.EndProcess()
            Return
        End If

        Dim dimType = "Aligned"
        If Not Me.Data.TryGetParameterValue(DIM_TYPE_PARAM_NAME, dimType) Then
            Me.Report.WriteEntry(ReportingLevel.Minimal,
                                 ReportEntryType.Information,
                                 TASK_NAME,
                                 TASK_CLASS,
                                 "No / Invalid value provided for the 'Dimension Type' rule. Assuming 'Aligned'.",
                                 Nothing)
        End If

        Dim dimLocX As Double
        Dim dimXString = String.Empty
        ' For some reason, TryGetParameterValueAsDouble was bringing anything less than 1 (ex. 0.2) over as 0
        ' So I switched to TryGetParameterValue and Double.TryParse
        If Not Me.Data.TryGetParameterValueAsDouble(DIM_X_LOCATION, dimLocX) Then
            Me.Report.WriteEntry(ReportingLevel.Minimal,
                                 ReportEntryType.Information,
                                 TASK_NAME,
                                 TASK_CLASS,
                                 "No / Invalid value provided for the 'X location of dimension' rule. Assuming '0'.",
                                 Nothing)
            'ElseIf Not Double.TryParse(dimXString, dimLocX) Then
            '    Me.Report.WriteEntry(ReportingLevel.Minimal,
            '                         ReportEntryType.Information,
            '                         TASK_NAME,
            '                         TASK_CLASS,
            '                         String.Format("Value provided for the 'X location of dimension' rule ('{0}') was not a number. Assuming '0'.", dimXString),
            '                         Nothing)
        End If

        Dim dimLocY As Double
        Dim dimYString = String.Empty
        If Not Me.Data.TryGetParameterValueAsDouble(DIM_Y_LOCATION, dimLocY) Then
            Me.Report.WriteEntry(ReportingLevel.Minimal,
                                 ReportEntryType.Information,
                                 TASK_NAME,
                                 TASK_CLASS,
                                 "No / Invalid value provided for the 'Y location of dimension' rule. Assuming '0'.",
                                 Nothing)
        End If

        Dim dimLocZ = 0

        ' Activate the drawing view
        If Not (model.Drawing.ActivateView(viewName)) Then
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Unable to activate view '{0}'", viewName))
            Return
        End If


        ' Select the entities
        If Not (SelectEntity(model, firstRefName, False)) Then
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Unable to select item '{0}'", firstRefName))
            Return
        End If

        If Not (SelectEntity(model, secondRefName, True)) Then
            Me.SetExecutionResult(TaskExecutionResult.Failed, String.Format("Unable to select item '{0}'", secondRefName))
            Return
        End If

        Dim displayDim As DisplayDimension
        Select Case dimType
            Case "Horizontal"
                displayDim = model.Model.AddHorizontalDimension2(dimLocX, dimLocY, dimLocZ)
            Case "Vertical"
                displayDim = model.Model.AddVerticalDimension2(dimLocX, dimLocY, dimLocZ)
            Case Else
                displayDim = model.Model.AddDimension2(dimLocX, dimLocY, dimLocZ)
        End Select
        displayDim.CenterText = True

    End Sub

    Private Function SelectEntity(model As SldWorksModelProxy,
                         entityName As String,
                         AddToSet As Boolean) As Boolean

        If Not (AddToSet) Then
            model.Model.ClearSelection2(All:=True)
        End If

        Dim extension = model.Extension
        ' All of these types are supported by SelectByID2, but may not be supported by Mate.
        ' We need to figure out which types to support in which order.
        Dim entityTypes = {"PLANE", "FACE", "AXIS", "EDGE", "VERTEX", "SILHOUETTE",
            "SKETCHSEGMENT", "SKETCHPOINT", "SKETCHPOINTFEAT", "SKETCH", "SKETCHTEXT", "SKETCHHATCH", "SKETCHREGION", "SKETCHCONTOUR", "SUBSKETCHDEF",
            "DATUMPOINT", "REFCURVE", "REFERENCECURVES", "REFSURFACE", "REFERENCE-EDGE", "POINTREF", "COORDSYS", "FRAMEPOINT",
            "EXTSKETCHSEGMENT", "EXTSKETCHPOINT", "EXTSKETCHTEXT", "HELIX"}   ',
        '"BODYFEATURE", "SURFACEBODY", "SOLIDBODY",
        '"OLEITEM", "EMBEDLINKDOC", "JOURNAL", "ATTRIBUTE", "OBJGROUP", "COMMENT", "PICTURE BODY", "SKETCHBITMAP",
        '"DRAWINGVIEW", "SECTIONLINE", "DETAILCIRCLE", "SECTIONTEXT", "SHEET", "BREAKLINE", "VIEWARROW", "ZONES", "TITLEBLOCK", "TITLEBLOCKTABLEFEAT",
        '"GTOL", "DIMENSIONS", "NOTE", "CENTERMARKS", "CENTERMARKSYMS", "CENTERLINE", "SFSYMBOL", "COSMETICWELDS", "MAGNETICLINES",
        '"DATUMTAG", "DTMTARG", "REFLINE", "CTHREAD", "HYPERLINK", "LEADER", "ANNVIEW",
        '"DOWLELSYM", "ANNOTATIONTABLES", "GENERALTABLEFEAT", "BLOCKDEF", "SUBSKETCHINST",
        '"COMPONENT", "MATE", "MATEGROUP", "MATEGROUPS", "MATESUPPLEMENT", "INCONTEXTFEAT", "INCONTEXTFEATS", "COMPPATTERN",
        '"WELD", "WELDBEADS",
        '"BOM", "BOMTEMP", "BOMFEATURE", "HOLETABLE", "HOLETABLEAXIS", "REVISIONTABLE", "REVISIONTABLEFEAT", "PUNCHTABLE", "HOLESERIES",
        '"EQNFOLDER", "IMPORTFOLDER", "FTRFOLDER", "BDYFOLDER", "DOCSFOLDER", "COMMENTSFOLDER", "SELECTIONSETFOLDER", "SUBSELECTIONSETNODE",
        '"EXPLODEDVIEWS", "EXPLODESTEPS", "EXPLODELINES", "CONFIGURATIONS", "VISUALSTATE",
        '"DCABINET", "ROUTEPOINT", "CONNECTIONPOINT", "ROUTEFABRICATED", "POSGROUP", "BROWSERITEM",
        '"LIGHTS", "MANIPULATOR", "CAMERAS",
        '"SIMULATION", "SIMULATION_ELEMENT",
        '"WELDMENT", "SUBWELDMENT", "WELDMENTTABLE",
        '"EVERYTHING", "LOCATIONS", "UNSUPPORTED",
        '"SWIFTANN", "SWIFTFEATURE", "SWIFTSCHEMA"

        For Each entityType In entityTypes
            If extension.SelectByID2(entityName, entityType, 0, 0, 0, AddToSet, 1, Nothing, swSelectOption_e.swSelectOptionDefault) Then
                ' Log to reporting to tell them what type we found
                Return True
            End If
        Next entityType

        ' Log to reporting that we could not select the entity
        Return False

    End Function

End Class
