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
