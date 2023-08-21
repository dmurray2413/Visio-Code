Sub AddControlPointOnGeometry()
    'On Error Resume Next
    Dim sel As Selection
    Dim sec As Section
    Dim shp As Shape
    Dim dDist As Double
    Dim iCPRow As Integer
    Dim strName As String
    
    
    'always adds a new controlpoint along geometry1.path.  The name will always be based on the row number
    ' of the added ControlPoint
    
    'Use the variable gDist to set the initial point along the geometry
    dDist = 0.45
    dDist = IIf(dDist >= 0 And dDist <= 1, dDist, 0.45)
    
    ' Remove embedded whitespace from strName
    ' strName = Replace(strName, " ", "")
    
    Set sel = ActiveWindow.Selection    ' create "selection"
    
    For x = 1 To sel.Count              ' iterate all shapes (from first to last) in selection
        
        Set shp = sel(x)                ' set current selected shape
        With shp
            ''  User Section
            ''     User.uistrNameX             =(BeginX)
            ''     User.uistrNameY             =(BeginY)
            ''     User.uistrNamePos           =NEARESTPOINTONPATH(Geometry1.Path,User.uistrNameX,User.uistrNameY)
            ''     User.uistrNameAngle         =ANGLEALONGPATH(Geometry1.Path,User.uistrNamePos)
            ''     User.uistrNameDistAlongPath =NEARESTPOINTONPATH(Geometry1.Path,User.uistrNameX,User.uistrNameY)
            ''  Controls Section
            ''     Controls.strNameCP   =SETATREF(User.uistrNameX,SETATREFEXPR(Width*-2.787))*0+Scratch.X1
            ''     Controls.strNameCP   =SETATREF(User.uistrNameX,SETATREFEXPR(Width*-2.787))*0+(PntX(POINTALONGPATH(Geometry1.Path,User.uistrNamePos)))
            ''     Controls.strNameCP.Y =SETATREF(User.uistrNameY,SETATREFEXPR(Height*2.0178))*0+Scratch.Y1
            ''     Controls.strNameCP.Y =SETATREF(User.uistrNameY,SETATREFEXPR(Height*2.0178))*0+(PntY(POINTALONGPATH(Geometry1.Path,User.uistrNamePos)))
            ''  DemoScratchSection (Now User Section)
            ''     Scratch.X =GUARD(POINTALONGPATH(Geometry1.Path,User.uistrNamePos))
            ''     Scratch.Y =GUARD(POINTALONGPATH(Geometry1.Path,User.uistrNamePos))

            'iCPRow = .AddNamedRow(visSectionControls, "_", 0)
            'If Not .CellExists("Controls._", 0) Then
            If Not .CellExistsU("Controls._", 0) Then
                iCPRow = .AddNamedRow(visSectionControls, "", visTagDefault)
                ' iCPRow = .CellsRowIndexU("Controls._")
            End If
            
            ' Debug.Print .Section(visSectionControls).Count
            CPCount = shp.Section(visSectionControls).Count
            ' dDist = CPCount / 10
            strName = "CP" & Format(iCPRow + 1, "0#")
            .CellsSRC(visSectionControls, iCPRow, visUserValue).RowNameU = strName
            
            ' User.uistrNameX = (BeginX)
            If Not .CellExists("User." & "ui" & strName & "X", 1) Then
                irow = .AddNamedRow(visSectionUser, "ui" & strName & "X", 0)
            Else
                irow = .CellsRowIndexU("User." & "ui" & strName & "X")
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "ui" & strName & "X"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = "(Width*" & Format(dDist, "0.0#") & ")"
            .CellsSRC(visSectionUser, irow, visUserPrompt).FormulaForceU = ""
            
            ' User.uistrNameY = (BeginY)
            If Not .CellExists("User." & "ui" & strName & "Y", 1) Then
                irow = .AddNamedRow(visSectionUser, "ui" & strName & "Y", 0)
            Else
                irow = .CellsRowIndexU("User." & "ui" & strName & "Y")
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "ui" & strName & "Y"
            ' .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = "BeginY"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = "(height*0)"
            .CellsSRC(visSectionUser, irow, visUserPrompt).FormulaForceU = ""
            
            ' User.strNamePos = NEARESTPOINTONPATH(Geometry1.Path, User.uistrNameX, User.uistrNameY)
            If Not .CellExists("User." & strName & "Pos", 1) Then
                irow = .AddNamedRow(visSectionUser, strName & "Pos", 0)
            Else
                irow = .CellsRowIndexU("User." & strName & "Pos")
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = strName & "Pos"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = "NEARESTPOINTONPATH(Geometry1.Path, User.ui" & strName & "X, User.ui" & strName & "Y)"
            .CellsSRC(visSectionUser, irow, visUserPrompt).FormulaForceU = ""
            
            ' User.strNameAngle = ANGLEALONGPATH(Geometry1.Path, User.uistrNamePos)
            If Not .CellExists("User." & strName & "Angle", 1) Then
                irow = .AddNamedRow(visSectionUser, strName & "Angle", 0)
            Else
                irow = .CellsRowIndexU("User." & strName & "Angle")
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = strName & "Angle"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = "ANGLEALONGPATH(Geometry1.Path, User." & strName & "Pos)"
            .CellsSRC(visSectionUser, irow, visUserPrompt).FormulaForceU = ""
            
            ' User.strNameDistAlongPath = NEARESTPOINTONPATH(Geometry1.Path, User.uistrNameX, User.uistrNameY)
            If Not .CellExists("User." & strName & "DistAlongPath", 1) Then
                 irow = .AddNamedRow(visSectionUser, strName & "DistAlongPath", 0)
            Else
                irow = .CellsRowIndexU("User." & strName & "DistAlongPath")
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = strName & "DistAlongPath"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = "NEARESTPOINTONPATH(Geometry1.Path, User." & "ui" & strName & "X, User." & "ui" & strName & "Y)"
            .CellsSRC(visSectionUser, irow, visUserPrompt).FormulaForceU = ""

            '' Controls Section
            ''  Controls.strNameCP   =SETATREF(User.uistrNameX,SETATREFEXPR(Width*-2.787))*0+(PntX(POINTALONGPATH(Geometry1.Path,User.uistrNamePos)))
            'If Not .CellExists("Controls." & strName & "CP", 1) Then
            '    iCPRow = .AddNamedRow(visSectionControls, strName & "CP", 0)
            'Else
            iCPRow = .CellsRowIndexU("Controls." & strName)
            'End If
            .CellsSRC(visSectionControls, iCPRow, visUserValue).RowNameU = strName
            .CellsSRC(visSectionControls, iCPRow, visCtlX).FormulaForceU = "SETATREF(User.ui" & strName & "X, SETATREFEXPR(Width*-2.787))*0+(PntX(POINTALONGPATH(Geometry1.Path, User." & strName & "Pos)))"
            .CellsSRC(visSectionControls, iCPRow, visCtlXDyn).FormulaForceU = "Controls." & strName
            .CellsSRC(visSectionControls, iCPRow, visCtlYDyn).FormulaForceU = "Controls." & strName & ".Y"
            .CellsSRC(visSectionControls, iCPRow, visCtlY).FormulaForceU = "SETATREF(User.ui" & strName & "Y, SETATREFEXPR(Width*-2.787))*0+(PntY(POINTALONGPATH(Geometry1.Path, User." & strName & "Pos,Char.Size/2)))"
            .CellsSRC(visSectionControls, iCPRow, visCtlTip).FormulaForceU = """Relocate " + Replace(strName, "CP", "ControlPoint ") + """"
            .CellsSRC(visSectionControls, iCPRow, visCtlGlue).FormulaForceU = "FALSE"
            ' .CellsSRC(visSectionControls, iCPRow, visCtlXDyn).FormulaForceU = "Controls.Row_1"
            ' .CellsSRC(visSectionControls, iCPRow, visCtlYDyn).FormulaForceU = "Controls.Row_1.Y"
            ' .CellsSRC(visSectionControls, iCPRow, visCtlXCon).FormulaForceU = "0"
            ' .CellsSRC(visSectionControls, iCPRow, visCtlYCon).FormulaForceU = "0"
            ' .CellsSRC(visSectionControls, iCPRow, visCtlGlue).FormulaForceU = "TRUE"
            ' .CellsSRC(visSectionControls, iCPRow, visCtlType).FormulaForceU = "0"
            ' .CellsSRC(visSectionControls, iCPRow, visCtlTip).FormulaForceU = """"""

        End With ' shp

    Next x

End Sub '
