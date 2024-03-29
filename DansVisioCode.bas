' =======================================================================================
' Module Name:  Dans_Visio_Code (DansVisioCode.bas)
' Author:       DMurray
' Purpose:      My Visio Shape Customizations
' Version:      v2014-08-06
' =======================================================================================
'Version History
'20140806.01 Commenced trying to use Version Control
'20140806.02 Begin declaring and using PUBLIC Variables

' =======================================================================================
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = True
'Attribute VB_Name = "Dans_Visio_Code"

Const STD_FONTSIZE = "10 pt"
' Visio only accesses Fonts by number, which are variable from machine to machine, and day to day
' Const STD_FONTFACE = 72 '<-- don't use this anymore, instead use 'FONTTOID()' which is a SHAPESHEET function, and is not available in VBA
Const STD_FONTFACE = "FONTTOID(""Consolas"")" '<-- instead use this.  The font name is CaSe_SenSiTiVe!!! (bastards)
'Limit shape width to between 0.25 <=>1.00in
Const CON_SHP_WIDTH As String = "BOUND(0 ft 1 in,0,FALSE,0.25 in*ThePage!DrawingScale/ThePage!PageScale,1 in*ThePage!DrawingScale/ThePage!PageScale)"
'Limit shape height to between 0.25 <=>1.00in
Const CON_SHP_HEIGHT As String = "BOUND(0 ft 1 in,0,FALSE,0.25 in*ThePage!DrawingScale/ThePage!PageScale,1 in*ThePage!DrawingScale/ThePage!PageScale)"
'limit shape location to within printable page
'limit shape horizontal placement
Const CON_PINX As String = "BOUND(0 ft 8.5 in,0,FALSE,ThePage!PageLeftMargin*ThePage!DrawingScale/ThePage!PageScale+Width/2,ThePage!PageWidth-ThePage!PageLeftMargin*ThePage!DrawingScale/ThePage!PageScale-Width/2)"
'limit shape vertical placement
Const CON_PINY As String = "BOUND(0 ft 5.5 in,0,FALSE,ThePage!PageLeftMargin*ThePage!DrawingScale/ThePage!PageScale+Height/2,ThePage!PageHeight-ThePage!PageLeftMargin*ThePage!DrawingScale/ThePage!PageScale-Height/2)"
'Print Margins [2019-05-07]
Const CON_PRNT_MARGIN As String = "PageTopMargin*ThePage!DrawingScale/ThePage!PageScale/ThePage!ScaleX"
'Charcacter sizes [2019-05-09]
Const CON_CHAR_SIZE As String = "BOUND(,0,FALSE,8 pt/ThePage!ScaleX,24 pt/ThePage!ScaleX)"
Public Function QuoteMe(s As String) As String
  QuoteMe = Chr(34) & s & Chr(34)
End Function
Sub PageClassifications()
    Dim sel As Selection
    Dim sec As Section
    Dim shp As Shape
    
    Set pg = ActiveWindow.Selection    ' create "selection"

End Sub
Sub MakePEO()
    '2019-04-23--dam
    '2019-04-29--dam added '/droponpagescale' to character size
    
    Dim sel As Selection
    Dim sec As Section
    Dim shp As Shape
    
    Set sel = ActiveWindow.Selection    ' create "selection"
    
    For x = 1 To sel.Count              ' iterate all shapes (from first to last) in selection
        Set shp = sel(x)                ' set current selected shape

        ' Clear out any existing rows in the "User-defined Cells" section
        For y = shp.RowCount(visSectionUser) - 1 To 0 Step -1
            shp.DeleteRow visSectionUser, y ' in current shape delete current row in section "User-defined Cells"
        Next y

        ' Clear out any existing rows in the "Shape Data" section
        For y = shp.RowCount(visSectionProp) - 1 To 0 Step -1
            shp.DeleteRow visSectionProp, y ' in current shape delete current row in section "Shape Data"
        Next y

        ' Clear out any existing rows in the "Actions" section
        For y = shp.RowCount(visSectionAction) - 1 To 0 Step -1
            shp.DeleteRow visSectionAction, y ' in current shape delete current row in section "Actions"
        Next y

        With shp
        
        ' Add Shape Data(property) Fields
            
            ' 01 ICRSId ##Joint Enterprise Level Agreement (JELA)##
            If .CellExists("prop.ICRS_Id", 0) Then
                irow = .CellsRowIndex("prop.ICRS_ID")
            Else
                irow = .AddNamedRow(visSectionProp, "ICRS_ID", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "ICRS_ID"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """ICRS_ID"""
            .CellsSRC(visSectionProp, irow, visCustPropsPrompt).FormulaForceU = """Joint Enterprise Level Agreement (JELA)"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """01"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """n/a"""
            
            ' 01a Parent ICRS Id ##Joint Enterprise Level Agreement (JELA)##
            If .CellExists("prop.Parent_ICRS_Id", 0) Then
                irow = .CellsRowIndex("prop.Parent_ICRS_Id")
            Else
                irow = .AddNamedRow(visSectionProp, "Parent_ICRS_Id", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "Parent_ICRS_Id"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Parent ICRS Id"""
            .CellsSRC(visSectionProp, irow, visCustPropsPrompt).FormulaForceU = """Joint Enterprise Level Agreement (JELA)"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """01a"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """n/a"""
            
            ' 02 SerialNumber
            If .CellExists("prop.SerialNumber", 0) Then
                irow = .CellsRowIndex("prop.SerialNumber")
            Else
                irow = .AddNamedRow(visSectionProp, "SerialNumber", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "SerialNumber"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Serial Number"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """02"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """--"""
            
            ' 03 Location
            If .CellExists("prop.Location", 0) Then
                irow = .CellsRowIndex("prop.Location")
            Else
                irow = .AddNamedRow(visSectionProp, "Location", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "Location"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Location"""
            .CellsSRC(visSectionProp, irow, visCustPropsType).FormulaForceU = """1"""
            .CellsSRC(visSectionProp, irow, visCustPropsFormat).FormulaForceU = """;UNKNOWN"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = "INDEX(1,Prop.Location.Format)"
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """03"""
            
            ' 04 Building
            If .CellExists("prop.Building", 0) Then
                irow = .CellsRowIndex("prop.Building")
            Else
                irow = .AddNamedRow(visSectionProp, "Building", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "Building"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Building"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """04"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """--"""
            
            ' 05 Room
            If .CellExists("prop.Room", 0) Then
                irow = .CellsRowIndex("prop.Room")
            Else
                irow = .AddNamedRow(visSectionProp, "Room", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "Room"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Room"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """05"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """--"""
            
            ' 06 Manufacturer
            If .CellExists("prop.Manufacturer", 0) Then
                irow = .CellsRowIndex("prop.Manufacturer")
            Else
                irow = .AddNamedRow(visSectionProp, "Manufacturer", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "Manufacturer"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Manufacturer"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """06"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = "=""Mfr"""
            
            ' 07 PartNumber
            If .CellExists("prop.PartNumber", 0) Then
                irow = .CellsRowIndex("prop.PartNumber")
            Else
                irow = .AddNamedRow(visSectionProp, "PartNumber", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "PartNumber"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Part Number"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """07"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """no p/n"""
            
            ' 02b AdminInterface
            If .CellExists("prop.AdminInterface", 0) Then
                irow = .CellsRowIndex("prop.AdminInterface")
            Else
                irow = .AddNamedRow(visSectionProp, "AdminInterface", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "AdminInterface"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Admin Interface"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """02b"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = "=""0.0.0.0/0"""
            
            ' 02c IPAddress
            If .CellExists("prop.IPAddress", 0) Then
                irow = .CellsRowIndex("prop.IPAddress")
            Else
                irow = .AddNamedRow(visSectionProp, "IPAddress", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "IPAddress"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """IP Address"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """02c"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = "=""0.0.0.0/0"""
            
            ' 09 MACAddress
            If .CellExists("prop.MACAddress", 0) Then
                irow = .CellsRowIndex("prop.MACAddress")
            Else
                irow = .AddNamedRow(visSectionProp, "MACAddress", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "MACAddress"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """MAC Address"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """09"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = "=""00:00:00:00:00:00"""
             
            ' 02a NetworkName
            If .CellExists("prop.NetworkName", 0) Then
                irow = .CellsRowIndex("prop.NetworkName")
            Else
                irow = .AddNamedRow(visSectionProp, "NetworkName", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "NetworkName"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Network Name"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """02a"""
            .CellsSRC(visSectionProp, irow, visCustPropsPrompt).FormulaForceU = """Device/Host Name"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """Device Name"""
            
            ' 11 NumberOfPorts
            If .CellExists("prop.NumberOfPorts", 0) Then
                irow = .CellsRowIndex("prop.NumberOfPorts")
            Else
                irow = .AddNamedRow(visSectionProp, "NumberOfPorts", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "NumberOfPorts"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Number of Ports"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """11"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """n/a"""
            
            ' 12 CommunityString
            If .CellExists("prop.CommunityString", 0) Then
                irow = .CellsRowIndex("prop.CommunityString")
            Else
                irow = .AddNamedRow(visSectionProp, "CommunityString", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "CommunityString"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Community String"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """12"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """PUBLIC"""
            
            ' 13 OSVersion
            If .CellExists("prop.OSVersion", 0) Then
                irow = .CellsRowIndex("prop.OSVersion")
            Else
                irow = .AddNamedRow(visSectionProp, "OSVersion", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "OSVersion"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """OS Version"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """13"""
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """n/a"""
            
            ' Hide NetworkDescription
            If .CellExists("prop.NetworkDescription", 0) Then
                irow = .CellsRowIndex("prop.NetworkDescription")
            Else
                irow = .AddNamedRow(visSectionProp, "NetworkDescription", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "NetworkDescription"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Network Description"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """"""
            .CellsSRC(visSectionProp, irow, visCustPropsInvis).FormulaForceU = "TRUE"
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """n/a"""
            
            ' Hide SubnetMask
            If .CellExists("prop.SubnetMask", 0) Then
                irow = .CellsRowIndex("prop.SubnetMask")
            Else
                irow = .AddNamedRow(visSectionProp, "SubnetMask", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "SubnetMask"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Subnet Mask"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """"""
            .CellsSRC(visSectionProp, irow, visCustPropsInvis).FormulaForceU = "TRUE"
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """255.255.255.0"""
            
            ' Hide ProductDescription
            If .CellExists("prop.ProductDescription", 0) Then
                irow = .CellsRowIndex("prop.ProductDescription")
            Else
                irow = .AddNamedRow(visSectionProp, "ProductDescription", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "ProductDescription"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Product Description"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """"""
            .CellsSRC(visSectionProp, irow, visCustPropsInvis).FormulaForceU = "TRUE"
            .CellsSRC(visSectionProp, irow, visCustPropsValue).FormulaForceU = """n/a"""
            
            ' Hide ProductNumber
            If .CellExists("prop.ProductNumber", 0) Then
                irow = .CellsRowIndex("prop.ProductNumber")
            Else
                irow = .AddNamedRow(visSectionProp, "ProductNumber", 0)
            End If
            .CellsSRC(visSectionProp, irow, visCustPropsValue).RowNameU = "ProductNumber"
            .CellsSRC(visSectionProp, irow, visCustPropsLabel).FormulaForceU = """Product Number"""
            .CellsSRC(visSectionProp, irow, visCustPropsSortKey).FormulaForceU = """"""
            .CellsSRC(visSectionProp, irow, visCustPropsInvis).FormulaForceU = "TRUE"
            
        ' Add User Fields
            ' AntiScale
            If .CellExists("user.AntiScale", 0) Then
                irow = .CellsRowIndex("user.AntiScale")
            Else
                irow = .AddNamedRow(visSectionUser, "AntiScale", 0)
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "AntiScale"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
                "IF(DropOnPageScale=1,User.AntiScale.Prompt,DropOnPageScale)"
            .CellsSRC(visSectionUser, irow, visUserPrompt).FormulaForceU = _
                """IF(DropOnPageScale=1,User.AntiScale.Prompt,DropOnPageScale)"""
                        
            ' HasText
            If .CellExists("user.HasText", 0) Then
                irow = .CellsRowIndex("user.HasText")
            Else
                irow = .AddNamedRow(visSectionUser, "HasText", 0)
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "HasText"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
                "NOT(OR(HideText,STRSAME(SHAPETEXT(TheText),"""")))"
            
            ' DefaultLabel
            If .CellExists("user.DefaultLabel", 0) Then
                irow = .CellsRowIndex("user.DefaultLabel")
            Else
                irow = .AddNamedRow(visSectionUser, "DefaultLabel", 0)
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "DefaultLabel"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
                "GUARD(Prop.NetworkName&CHAR(10)&Prop.AdminInterface&CHAR(10)&Prop.Manufacturer&"" ""&Prop.PartNumber&CHAR(10)&""OS:""&Prop.OSVersion)"
            
            ' JELALabel
            If .CellExists("user.JELALabel", 0) Then
                irow = .CellsRowIndex("user.JELALabel")
            Else
                irow = .AddNamedRow(visSectionUser, "JELALabel", 0)
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "JELALabel"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
                "GUARD(""ICRS Id:""&Prop.ICRS_Id&CHAR(10)&Prop.NetworkName&CHAR(10)&Prop.AdminInterface&CHAR(10)&Prop.Manufacturer&"" ""&Prop.PartNumber&CHAR(10)&""OS:""&Prop.OSVersion)"
            
            ' NoIPLabel
            If .CellExists("user.NoIPLabel", 0) Then
                irow = .CellsRowIndex("user.NoIPLabel")
            Else
                irow = .AddNamedRow(visSectionUser, "NoIPLabel", 0)
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "NoIPLabel"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
                "GUARD(Prop.NetworkName&CHAR(10)&Prop.Manufacturer&"" ""&Prop.PartNumber&CHAR(10)&""OS:""&Prop.OSVersion)"
            
            ' LabelPos
            If .CellExists("user.LabelPos", 0) Then
                irow = .CellsRowIndex("user.LabelPos")
            Else
                irow = .AddNamedRow(visSectionUser, "LabelPos", 0)
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "LabelPos"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
                """Bottom"""
        
            ' LabelType
            If .CellExists("user.LabelType", 0) Then
               irow = .CellsRowIndex("user.LabelType")
            Else
                irow = .AddNamedRow(visSectionUser, "LabelType", 0)
            End If
            .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "LabelType"
            .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
                """NoIPLabel"""
        
        ' Create Label
            'uncomment below to clear any previous label
             If .CharCount <> 0 Then
               If MsgBox("Replace " & .Name & " Label?", vbYesNo + vbQuestion, "Replace " & .Name & " Label?") = vbYes Then
                 .CellsSRC(visSectionObject, visRowLock, visLockTextEdit).FormulaForceU = False
                 Set priorlabel = .Characters
                 priorlabel.Begin = 0
                 priorlabel.End = .Characters.CharCount
                 priorlabel.Delete
                 Set priorlabel = Nothing
               End If
             End If
            'uncomment above to clear any previous label

            .CellsSRC(visSectionObject, visRowLock, visLockTextEdit).FormulaForceU = False
            Set NewLabel = .Characters
            NewLabel.Begin = 0
            NewLabel.End = 0
            NewLabel.AddCustomFieldU "=User.DefaultLabel", visFmtNumGenNoUnits
            
        ' Actions
            irow = .AddNamedRow(visSectionAction, "LabelType", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelType"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """Label Type"""
            .CellsSRC(visSectionAction, irow, visActionBeginGroup).FormulaForceU = "true"

            irow = .AddNamedRow(visSectionAction, "JELALabel", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "JELALabel"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&JELA"""
            .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = _
                "GUARD(SETF(""Fields.Value"",""User.JELALabel"")+SETF(GetRef(User.LabelType),""""""JELA"""""")+SETF(GetRef(HideText),FALSE))"
            .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelType,""JELA"")"
            .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"

            irow = .AddNamedRow(visSectionAction, "NoIPLabel", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "NoIPLabel"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&NoIP"""
            .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = _
                "GUARD(SETF(""Fields.Value"",""User.NoIPLabel"")+SETF(GetRef(User.LabelType),""""""NoIP"""""")+SETF(GetRef(HideText),FALSE))"
            .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelType,""NoIP"")"
            .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"

            irow = .AddNamedRow(visSectionAction, "DefaultLabel", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "DefaultLabel"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Default"""
            .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = _
                "GUARD(SETF(""Fields.Value"",""User.DefaultLabel"")+SETF(GetRef(User.LabelType),""""""Default"""""")+SETF(GetRef(HideText),FALSE))"
            .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelType,""Default"")"
            .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"

            irow = .AddNamedRow(visSectionAction, "NoLabel", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "NoLabel"
            .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "SETF(GetRef(HideText),NOT(HideText))+SETF(GetRef(User.LabelType),""""""None"""""")"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """N&one"""
            .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "HideText"
            .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
            
            irow = .AddNamedRow(visSectionAction, "LabelPos", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelPos"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """Position"""
            .CellsSRC(visSectionAction, irow, visActionBeginGroup).FormulaForceU = "true"
            
            irow = .AddNamedRow(visSectionAction, "LabelTop", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelTop"
            .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = _
                "GUARD(SETF(""Controls.TextPosition"",""Width/2"")+SETF(""Controls.TextPosition.Y"",""height+Txtheight/2"")+SETF(""Para.HorzAlign"",""1"")+SETF(GetRef(User.LabelPos),""""""Top""""""))"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Top"""
            .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Top"")"
            .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
            .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
            
            irow = .AddNamedRow(visSectionAction, "LabelBottom", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelBottom"
            .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = _
                "GUARD(SETF(""Controls.TextPosition"",""Width/2"")+SETF(""Controls.TextPosition.Y"",""-Txtheight/2"")+SETF(""Para.HorzAlign"",""1"")+SETF(GetRef(User.LabelPos),""""""Bottom""""""))"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Bottom"""
            .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Bottom"")"
            .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
            .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
            
            irow = .AddNamedRow(visSectionAction, "LabelLeft", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelLeft"
            .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = _
                "GUARD(SETF(""Controls.TextPosition"",""-txtWidth/2"")+SETF(""Controls.TextPosition.Y"",""height/2"")+SETF(""Para.HorzAlign"",""2"")+SETF(GetRef(User.LabelPos),""""""Left""""""))"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Left"""
            .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Left"")"
            .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
            .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
            
            irow = .AddNamedRow(visSectionAction, "LabelRight", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelRight"
            .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = _
                "GUARD(SETF(""Controls.TextPosition"",""Width+txtWidth/2"")+SETF(""Controls.TextPosition.Y"",""height/2"")+SETF(""Para.HorzAlign"",""0"")+SETF(GetRef(User.LabelPos),""""""Right""""""))"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Right"""
            .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Right"")"
            .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
            .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
            
            irow = .AddNamedRow(visSectionAction, "HideLabel", 0)
            .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "HideLabel"
            .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "SETF(GetRef(HideText),NOT(HideText))"
            .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """Hidden"""
            .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "HideText"
            .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
            
        ' Standard formatting
            .CellsSRC(visSectionParagraph, 0, visSpaceLine).FormulaForceU = "-100%"
            .CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaForceU = "1 pt"
            .CellsSRC(visSectionObject, visRowText, visTxtBlkRightMargin).FormulaForceU = "1 pt"
            .CellsSRC(visSectionObject, visRowText, visTxtBlkTopMargin).FormulaForceU = "1 pt"
            .CellsSRC(visSectionObject, visRowText, visTxtBlkBottomMargin).FormulaForceU = "1 pt"
            .CellsSRC(visSectionObject, visRowTextXForm, visXFormWidth).FormulaU = "(TEXTWIDTH(TheText))"
            .CellsSRC(visSectionObject, visRowTextXForm, visXFormHeight).FormulaU = "(TEXTHEIGHT(TheText,TEXTWIDTH(TheText)))"
            
        ' reset Label Position
            If Not .CellExists("Controls.TextPosition", 1) Then
                'Debug.Assert .AddNamedRow(visSectionControls, "Row_1", 0)
                irow = .AddNamedRow(visSectionControls, "Row_1", 0)
            End If
        ' reset Label Size
            .CellsSRC(visSectionControls, 0, visCtlX).FormulaForceU = "Width*0.5"
            .CellsSRC(visSectionControls, 0, visCtlY).FormulaForceU = "-TxtHeight*0.5"
            .CellsSRC(visSectionObject, visRowTextXForm, visXFormPinX).FormulaForceU = "SETATREF(Controls.TextPosition)"
            .CellsSRC(visSectionObject, visRowTextXForm, visXFormPinY).FormulaForceU = "SETATREF(Controls.TextPosition.Y)"
            
            'reset standard font to scaled 10pt, Consolas
            .CellsSRC(visSectionCharacter, 0, visCharacterFont).FormulaForceU = "=FONTTOID(""Consolas"")"
            'added /droponpagescale  20190429-dam
              ' .CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaForceU = "10 pt*(Width/1 in)"
              ' .CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaForceU = "10 pt*(Width/1 in)/droponpagescale"
            .CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaForceU = "10 pt"
            
            'reset text block to 25% opacity, with default tabs at 1/10th of the shape width
            .CellsSRC(visSectionObject, visRowText, visTxtBlkBkgndTrans).FormulaForceU = "25%"
            .CellsSRC(visSectionObject, visRowText, visTxtBlkDefaultTabStop).FormulaForceU = "Width/10"
        
            'set comment field/hover text
            .CellsSRC(visSectionObject, visRowMisc, visComment).FormulaForceU = _
                "GUARD(Prop.NetworkName&char(13)&char(10)&Prop.AdminInterface)"
        
            'Reset Shape Text Protection
            .CellsSRC(visSectionObject, visRowLock, visLockTextEdit).FormulaForceU = True
            
            'Enable Double-Click to show properties
            .CellsSRC(visSectionObject, visRowEvent, visEvtCellDblClick).FormulaForceU = "DOCMD(1312)"
        End With ' shp
    Next x
End Sub ' MakePEO
Sub ApplyLabelActions()
    On Error Resume Next
    Dim sel As Selection
    Dim sec As Section
    Dim shp As Shape
    
    Set sel = ActiveWindow.Selection    ' create "selection"
    
    For x = 1 To sel.Count              ' iterate all shapes (from first to last) in selection
        Set shp = sel(x)                ' set current selected shape

        'ClearUserDefinedCells shp
        'ClearShapeData shp
        ' ClearShapeActions shp
        ResetLabelPosition shp
        AddUserFields shp
        AddLabelActions shp

    Next x
End Sub
Private Sub AddLabelActionRows(shp As Shape)
    On Error Resume Next
    With shp
        irow = .AddNamedRow(visSectionAction, "LabelPos", 0)
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelPos"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """Position"""
        .CellsSRC(visSectionAction, irow, visActionBeginGroup).FormulaForceU = "true"
        
        irow = .AddNamedRow(visSectionAction, "LabelTop", 0)
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelTop"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "GUARD(SETF(""Controls.TextPosition"",""Width/2"")+SETF(""Controls.TextPosition.Y"",""height+Txtheight/2"")+SETF(""Para.HorzAlign"",""1"")+SETF(GetRef(User.LabelPos),""""""Top""""""))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Top"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Top"")"
        .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
        
        irow = .AddNamedRow(visSectionAction, "LabelBottom", 0)
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelBottom"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "GUARD(SETF(""Controls.TextPosition"",""Width/2"")+SETF(""Controls.TextPosition.Y"",""-Txtheight/2"")+SETF(""Para.HorzAlign"",""1"")+SETF(GetRef(User.LabelPos),""""""Bottom""""""))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Bottom"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Bottom"")"
        .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
        
        irow = .AddNamedRow(visSectionAction, "LabelLeft", 0)
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelLeft"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "GUARD(SETF(""Controls.TextPosition"",""-txtWidth/2"")+SETF(""Controls.TextPosition.Y"",""height/2"")+SETF(""Para.HorzAlign"",""2"")+SETF(GetRef(User.LabelPos),""""""Left""""""))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Left"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Left"")"
        .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
        
        irow = .AddNamedRow(visSectionAction, "LabelRight", 0)
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelRight"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "GUARD(SETF(""Controls.TextPosition"",""Width+txtWidth/2"")+SETF(""Controls.TextPosition.Y"",""height/2"")+SETF(""Para.HorzAlign"",""0"")+SETF(GetRef(User.LabelPos),""""""Right""""""))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Right"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Right"")"
        .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
        
        irow = .AddNamedRow(visSectionAction, "HideLabel", 0)
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "HideLabel"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "SETF(GetRef(HideText),NOT(HideText))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """Hidden"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
    End With ' shp
End Sub ' AddLabelActionRows
Private Sub AddLabelActions(shp As Shape)
    On Error Resume Next
    With shp
        If .CellExists("Actions.LabelPos", 0) Then
            irow = .CellsRowIndex("Actions.LabelPos")
        Else
            irow = .AddNamedRow(visSectionAction, "LabelPos", 0)
        End If
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelPos"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """Label Position"""
        .CellsSRC(visSectionAction, irow, visActionBeginGroup).FormulaForceU = "true"
        
        'LabelCenter
        If .CellExists("Actions.LabelCenter", 0) Then
            irow = .CellsRowIndex("Actions.LabelCenter")
        Else
            irow = .AddNamedRow(visSectionAction, "LabelCenter", 0)
        End If
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelCenter"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "GUARD(SETF(""Controls.TextPosition"",""Width/2"")+SETF(""Controls.TextPosition.Y"",""height/2"")+SETF(""Para.HorzAlign"",""1"")+SETF(GetRef(User.LabelPos),""""""Center""""""))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Center"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Center"")"
        .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
        
        'LabelTop
        If .CellExists("Actions.LabelTop", 0) Then
            irow = .CellsRowIndex("Actions.LabelTop")
        Else
            irow = .AddNamedRow(visSectionAction, "LabelTop", 0)
        End If
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelTop"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "GUARD(SETF(""Controls.TextPosition"",""Width/2"")+SETF(""Controls.TextPosition.Y"",""height+Txtheight/2"")+SETF(""Para.HorzAlign"",""1"")+SETF(GetRef(User.LabelPos),""""""Top""""""))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Top"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Top"")"
        .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
        
        'LabelBottom
        If .CellExists("Actions.LabelBottom", 0) Then
            irow = .CellsRowIndex("Actions.LabelBottom")
        Else
            irow = .AddNamedRow(visSectionAction, "LabelBottom", 0)
        End If
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelBottom"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "GUARD(SETF(""Controls.TextPosition"",""Width/2"")+SETF(""Controls.TextPosition.Y"",""-Txtheight/2"")+SETF(""Para.HorzAlign"",""1"")+SETF(GetRef(User.LabelPos),""""""Bottom""""""))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Bottom"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Bottom"")"
        .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
        
        'LabelLeft
        If .CellExists("Actions.LabelLeft", 0) Then
            irow = .CellsRowIndex("Actions.LabelLeft")
        Else
            irow = .AddNamedRow(visSectionAction, "LabelLeft", 0)
        End If
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelLeft"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "GUARD(SETF(""Controls.TextPosition"",""-txtWidth/2"")+SETF(""Controls.TextPosition.Y"",""height/2"")+SETF(""Para.HorzAlign"",""2"")+SETF(GetRef(User.LabelPos),""""""Left""""""))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Left"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Left"")"
        .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
        
        'LabelRight
        If .CellExists("Actions.LabelRight", 0) Then
            irow = .CellsRowIndex("Actions.LabelRight")
        Else
            irow = .AddNamedRow(visSectionAction, "LabelRight", 0)
        End If
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "LabelRight"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "GUARD(SETF(""Controls.TextPosition"",""Width+txtWidth/2"")+SETF(""Controls.TextPosition.Y"",""height/2"")+SETF(""Para.HorzAlign"",""0"")+SETF(GetRef(User.LabelPos),""""""Right""""""))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """&Right"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "STRSAME(User.LabelPos,""Right"")"
        .CellsSRC(visSectionAction, irow, visActionDisabled).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
        
        'HideLabel
        If .CellExists("Actions.HideLabel", 0) Then
            irow = .CellsRowIndex("Actions.HideLabel")
        Else
            irow = .AddNamedRow(visSectionAction, "HideLabel", 0)
        End If
        .CellsSRC(visSectionAction, irow, visActionMenu).RowNameU = "HideLabel"
        .CellsSRC(visSectionAction, irow, visActionAction).FormulaForceU = "SETF(GetRef(HideText),NOT(HideText))"
        .CellsSRC(visSectionAction, irow, visActionMenu).FormulaForceU = """Hidden"""
        .CellsSRC(visSectionAction, irow, visActionChecked).FormulaForceU = "HideText"
        .CellsSRC(visSectionAction, irow, visActionFlyoutChild).FormulaForceU = "TRUE"
    End With ' shp
End Sub 'AddLabelActions
Private Sub AddUserFields(shp As Shape)
    On Error Resume Next
    With shp
        ' Add User Fields
        ' AntiScale
        If .CellExists("user.AntiScale", 0) Then
            irow = .CellsRowIndex("user.AntiScale")
        Else
            'irow = .CellsRowIndexU("user.AntiScale")
            irow = .AddNamedRow(visSectionUser, "AntiScale", 0)
        End If
        .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "AntiScale"
        .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
            "IF(DropOnPageScale=1,User.AntiScale.Prompt,DropOnPageScale)"
        .CellsSRC(visSectionUser, irow, visUserPrompt).FormulaForceU = _
            """IF(DropOnPageScale=1,User.AntiScale.Prompt,DropOnPageScale)"""
                    
        ' HasText
        If .CellExists("user.HasText", 0) Then
            irow = .CellsRowIndex("user.HasText")
        Else
            irow = .CellsRowIndexU("user.HasText")
            'irow = .AddNamedRow(visSectionUser, "HasText", 0)
        End If
        .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "HasText"
        .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
            "NOT(OR(HideText,STRSAME(SHAPETEXT(TheText),"""")))"
        
        ' LabelPos
        If .CellExists("user.LabelPos", 0) Then
            irow = .CellsRowIndex("user.LabelPos")
        Else
            irow = .CellsRowIndexU("user.LabelPos")
            'irow = .AddNamedRow(visSectionUser, "LabelPos", 0)
        End If
        .CellsSRC(visSectionUser, irow, visUserValue).RowNameU = "LabelPos"
        .CellsSRC(visSectionUser, irow, visUserValue).FormulaForceU = _
            """Bottom"""
    End With ' shp
End Sub ' AddUserFields
Private Sub ClearShapeData(shp As Shape)
    On Error Resume Next
    With shp
        ' Clear out any existing rows in the "Shape Data" section
        For y = .RowCount(visSectionProp) - 1 To 0 Step -1
            ' in current shape delete current row in section "Shape Data"
            .DeleteRow visSectionProp, y
        Next y
    End With ' shp
End Sub ' ClearShapeData
Private Sub ClearShapeActions(shp As Shape)
    On Error Resume Next
    With shp
        ' Clear out any existing rows in the "Actions" section
        For y = .RowCount(visSectionAction) - 1 To 0 Step -1
            ' in current shape delete current row in section "Actions"
            .DeleteRow visSectionAction, y
        Next y
    End With ' shp
End Sub ' ClearShapeActions
Private Sub ClearUserDefinedCells(shp As Shape)
    On Error Resume Next
    With shp
        ' Clear out any existing rows in the "User-defined Cells" section
        For y = .RowCount(visSectionUser) - 1 To 0 Step -1
            ' in current shape delete current row in section "User-defined Cells"
            .DeleteRow visSectionUser, y
        Next y
    End With ' shp
End Sub ' ClearUserDefinedCells
Private Sub Default_Format()
  With ActiveWindow.Selection(1)
    On Error Resume Next
    
    .DeleteSection visSectionAction
    
    Default_Format
    
    'Add User Fields
    If Not .CellExists("user.Hypot", 1) Then
    .AddNamedRow visSectionUser, "Hypot", 0
    End If
    .Cells("user.Hypot").Formula = "SQRT(Height^2+Width^2)/1 in"
    
    If Not .CellExists("user.FontSize", 1) Then
    .AddNamedRow visSectionUser, "FontSize", 0
    End If
    .Cells("user.FontSize").Formula = "MIN(STD_FONTSIZE,STD_FONTSIZE*User.Hypot)"
    
    If Not .CellExists("user.TextWidth", 1) Then
    .AddNamedRow visSectionUser, "TextWidth", 0
    End If
    .Cells("user.TextWidth").Formula = "TEXTWIDTH(TheText)"
    
    If Not .CellExists("user.TextAt", 1) Then
    .AddNamedRow visSectionUser, "TextAt", 0
    End If
    .Cells("user.TextAt").FormulaU = "=""Center"""
    
    'CenterText
    If Not .CellExists("Actions.CenterText", 1) Then
    .AddNamedRow visSectionAction, "CenterText", 0
    End If
    .Cells("Actions.CenterText").FormulaForceU = sL_Menu
    .Cells("Actions.CenterText.Action").Formula = sL_Action
    .Cells("Actions.CenterText.Checked").Formula = "=STRSAME(User.TextAt,""Center"")"
    
    'LeftText
    If Not .CellExists("Actions.LeftText", 1) Then
    .AddNamedRow visSectionAction, "LeftText", 0
    End If
    .Cells("Actions.LeftText").FormulaForceU = sL_Menu
    .Cells("Actions.LeftText.Action").Formula = sL_Action
    .Cells("Actions.LeftText.Checked").Formula = "=STRSAME(User.TextAt,""Left"")"
    
    'RightText
    If Not .CellExists("Actions.LeftText", 1) Then
    .AddNamedRow visSectionAction, "RightText", 0
    End If
    .Cells("Actions.RightText").FormulaForceU = sR_Menu
    .Cells("Actions.RightText.Action").Formula = sR_Action
    .Cells("Actions.RightText.Checked").Formula = "=STRSAME(User.TextAt,""Right"")"
    
    'TopText
    If Not .CellExists("Actions.TopText", 1) Then
    .AddNamedRow visSectionAction, "TopText", 0
    End If
    .Cells("Actions.TopText").FormulaForceU = sT_Menu
    .Cells("Actions.TopText.Action").Formula = sT_Action
    .Cells("Actions.TopText.Checked").Formula = "=STRSAME(User.TextAt,""Top"")"
    
    'BottomText
    If Not .CellExists("Actions.BottomText", 1) Then
    .AddNamedRow visSectionAction, "BottomText", 0
    End If
    .Cells("Actions.BottomText").FormulaForceU = sB_Menu
    .Cells("Actions.BottomText.Action").Formula = sB_Action
    .Cells("Actions.BottomText.Checked").Formula = "=STRSAME(User.TextAt,""Bottom"")"
    
    'CenterText
    If Not .CellExists("Actions.CenterText", 1) Then
    .AddNamedRow visSectionAction, "CenterText", 0
    End If
    .Cells("Actions.CenterText").FormulaForceU = sC_Menu
    .Cells("Actions.CenterText.Action").Formula = sC_Action
    .Cells("Actions.CenterText.Checked").Formula = "=STRSAME(User.TextAt,""Center"")"
End With
End Sub
Private Sub ResetLabelPosition(shp As Shape)
    On Error Resume Next
    With shp
        ' reset Label Position
        If Not .CellExists("controls.TextPosition", 1) Then
            'Debug.Assert .AddNamedRow(visSectionControls, "TextPosition", 0)
            irow = .AddNamedRow(visSectionControls, "TextPosition", 0)
        Else
            irow = .CellsRowIndexU("Controls.TextPosition")
        End If
        ' reset Label Size
        .CellsSRC(visSectionControls, irow, visCtlX).FormulaForceU = "Width*0.5"
        .CellsSRC(visSectionControls, irow, visCtlY).FormulaForceU = "-TxtHeight*0.5"
        .CellsSRC(visSectionObject, visRowTextXForm, visXFormPinX).FormulaForceU = "SETATREF(Controls.TextPosition)"
        .CellsSRC(visSectionObject, visRowTextXForm, visXFormPinY).FormulaForceU = "SETATREF(Controls.TextPosition.Y)"
        
        'reset standard font to scaled 10pt, Consolas
        ' .CellsSRC(visSectionCharacter, 0, visCharacterFont).FormulaForceU = "=FONTTOID(""Consolas"")"
        'added /droponpagescale  20190429-dam
          ' .CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaForceU = "10 pt*(Width/1 in)"
          ' .CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaForceU = "10 pt*(Width/1 in)/droponpagescale"
        
        'reset text block to 25% opacity, with default tabs at 1/10th of the shape width
        .CellsSRC(visSectionObject, visRowText, visTxtBlkBkgndTrans).FormulaForceU = "25%"
        .CellsSRC(visSectionObject, visRowText, visTxtBlkDefaultTabStop).FormulaForceU = "Width/10"
    
        ' Standard formatting
        .CellsSRC(visSectionParagraph, 0, visSpaceLine).FormulaForceU = "-100%"
        .CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaForceU = "1 pt"
        .CellsSRC(visSectionObject, visRowText, visTxtBlkRightMargin).FormulaForceU = "1 pt"
        .CellsSRC(visSectionObject, visRowText, visTxtBlkTopMargin).FormulaForceU = "1 pt"
        .CellsSRC(visSectionObject, visRowText, visTxtBlkBottomMargin).FormulaForceU = "1 pt"
        .CellsSRC(visSectionObject, visRowTextXForm, visXFormWidth).FormulaU = "(TEXTWIDTH(TheText))"
        .CellsSRC(visSectionObject, visRowTextXForm, visXFormHeight).FormulaU = "(TEXTHEIGHT(TheText,TEXTWIDTH(TheText)))"
    End With ' shp
End Sub ' ResetLabelPosition
Private Sub ResetFontSizeMenuItem()
  sCommon = "SETF(""User.Hypot"",""SQRT(Height^2+Width^2)/1 in"")+" & _
   "SETF(""User.FontSize"",""MIN(10 pt,10 pt*User.Hypot)"" )+" & _
   "SETF(""User.TextWidth"",""TEXTWIDTH(TheText)"" )+" & _
   "SETF(""TxtPinX"",""Controls.TextPosition"" )+" & _
   "SETF(""TxtPiny"",""Controls.TextPosition.Y"" )+" & _
   "SETF(""Char.size"",""user.FontSize"" )+" & _
   "SETF(""Para.IndFirst"",""Width*-0.0833"" )+" & _
   "SETF(""Para.IndLeft"",""Width*0.0833"" )+" & _
   "SETF(""Para.SpLine"",""-100%"" )+" & _
   "SETF(""LeftMargin"",""1 pt"" )+" & _
   "SETF(""RightMargin"",""1 pt"" )+" & _
   "SETF(""TopMargin"",""1 pt"" )+" & _
   "SETF(""BottomMargin"",""1 pt"" )+" & _
   "SETF(""TxtWidth"",""MIN(TEXTWIDTH(TheText),Width )"" )+" & _
   "SETF(""TxtHeight"",""TEXTHEIGHT(TheText,TxtWidth)"" )+" & _
   "SETF(""TxtAngle"",""IF(BITXOR(FlipX,FlipY),1,-1)*Angle"")+" & _
   "SETF(""TxtLocPinX"",""TxtWidth/2"" )+" & _
   "SETF(""TxtLocPiny"",""TxtHeight/2"" )"
  sRF_Menu = QuoteMe("Reset FontSize")
  sRF_Action = "SETF(""User.Hypot"",""SQRT(Height^2+Width^2)/1 in"")+" & _
   "SETF(""User.FontSize"",""MIN(10 pt,10 pt*User.Hypot)"" )+" & _
   "SETF(""User.TextWidth"",""TEXTWIDTH(TheText)"" )+" & _
   "SETF(""Char.size"",""user.fontsize"")+" & _
   "SETF(""Para.IndFirst"",""Width*-0.0833"" )+" & _
   "SETF(""Para.IndLeft"",""Width*0.0833"" )+" & _
   "SETF(""Para.SpLine"",""-100%"" )+" & _
   "SETF(""LeftMargin"",""1 pt"" )+" & _
   "SETF(""RightMargin"",""1 pt"" )+" & _
   "SETF(""TopMargin"",""1 pt"" )+" & _
   "SETF(""BottomMargin"",""1 pt"" )+" & _
   "SETF(""TextBkgndTrans"",""30%"" )+" & _
   "SETF(""TextBkgnd"",""0"" )+" & _
   "SETF(""TxtPinX"",""Controls.TextPosition"" )+" & _
   "SETF(""TxtPiny"",""Controls.TextPosition.Y"" )+" & _
   "SETF(""TxtLocPinX"",""TxtWidth/2"" )+" & _
   "SETF(""TxtLocPiny"",""TxtHeight/2"" )"
  With ActiveWindow.Selection(1)
    On Error Resume Next
    Default_Format
    .AddNamedRow visSectionUser, "Hypot", 0
    .AddNamedRow visSectionUser, "FontSize", 0
    .AddNamedRow visSectionUser, "TextWidth", 0
    .AddNamedRow visSectionUser, "TextAt", 0
    .AddNamedRow visSectionAction, "ResetFontSize", 0
    .Cells("Actions.ResetFontSize").FormulaForceU = sRF_Menu
    .Cells("Actions.ResetFontSize.Action").Formula = sRF_Action
  End With
End Sub
Private Sub TestActionAdder()
  Dim sTmp As String
  Dim sMnu As String
  sTmp = "=This Is A Test"
  sMnu = "Test"
  With ActiveWindow.Selection(1)
    ' This updates the visible menu cell, Leave it blank if for 'internal' use
    ' only
    .Cells("Actions.NewActionRow").FormulaForceU = QuoteMe(sMnu)
    'Update the action cell (iTheRow MUST be an integer)
    .Cells("Actions.NewActionRow.Action").Formula = QuoteMe(sTmp)
  End With
End Sub
Public Sub SortStencil()
Dim i As Integer
Dim j As Integer
Dim pntr As Integer
Dim lastmst As String
Dim mstrcnt As Integer
mastrcnt = ThisDocument.Masters.Count
For i = 1 To mastrcnt
    lastmst = ""
    pntr = 0
    For j = i To mastrcnt
        If ThisDocument.Masters(j).Type = 1 Then
            If ThisDocument.Masters(j).Name > lastmst Then
                lastmst = ThisDocument.Masters(j).Name
                pntr = j
            End If
        End If
    Next j
 
    If pntr > 0 Then ThisDocument.Masters(pntr).IndexInStencil = 1
Next i

End Sub

Private Sub DoSomethingWithAllSelectedShapes()
    ' D.Murray
    ' see https://learn.microsoft.com/en-us/office/vba/api/visio(enumerations) for native constant enumerations
    Dim sel As Selection
    Dim sec As Section
    Dim shp As Shape
    
    Set sel = ActiveWindow.Selection    ' create "selection"
    
    For c = 1 To sel.Count              ' iterate all shapes (from first to last) in selection
        Set shp = sel(c)                ' set current selected shape
        
        With shp                        ' do something with the current shape
            '                           'how many control rows?
            ' Debug.Print .RowCount(visSectionControls)
            
            '                           'Make sure we have a TextPosition section
            ' reset Label Size
            .CellsU("TxtWidth").FormulaU = "TEXTWIDTH(TheText)"
            .CellsU("TxtHeight").FormulaU = "TEXTHEIGHT(TheText,TxtWidth)"
        
            '                           'MakeSure we have a TextPosition control
            If Not .CellExists("controls.TextPosition", 1) Then
                irow = .AddNamedRow(visSectionControls, "TextPosition", 0)
            End If
            ' reset Label Size
            .CellsU("Controls.TextPosition").FormulaU = "Width*0.5"
            .CellsU("Controls.TextPosition.Y").FormulaU = "Height*0.5"
            .CellsU("TxtPinX").FormulaU = "Controls.TextPosition"
            .CellsU("TxtPinY").FormulaU = "Controls.TextPosition.Y"
        
            ' Label Ethernet Controls
            SpacingPctg = (.CellsU("width") / (.RowCount(visSectionControls) - 1))
            
            Debug.Print .RowCount(visSectionControls)
            For CR = 0 To .RowCount(visSectionControls) - 1
                Debug.Print "CR: " & CR & " " & .CellsSRC(visSectionControls, CR, 1).Name
                .CellsSRC(visSectionControls, CR, 1).Name = "Device_" & (CR + 1)
            Next CR
        End With                        ' do something with the current shape
        
    Next 'c                             ' iterate all shapes (from first to last) in selection
        
End Sub
