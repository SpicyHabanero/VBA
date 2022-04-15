Attribute VB_Name = "AutoWork"
Public Out
Public Comm
Public USDINT
Public USDPRI
Public EURINT
Public EURPRI
Public JPYINT
Public JPYPRI
Public strDate

'Big boy - heavy lifting and automate my job :)
Sub DailyRec()

Dim exesht As Worksheet
Dim SaveDir As String
Dim file1 As String
Dim file3 As String
Dim file12 As String
Dim sPrinter As String
Dim sDefaultPrinter As String

    'Turn off stuff
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    sDefaultPrinter = Application.ActivePrinter ' store default printer
    sPrinter = GetPrinterFullName("ANA-PR101")
    If sPrinter = vbNullString Then ' no match
        'Do nothing
    Else
        Application.ActivePrinter = sPrinter
    End If

    OrgPath = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\"
    Const SOME_PATH As String = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\"
    Const WSOME_PATH As String = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\WSO Reports\"
    PathName = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\"
    TempPath = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\"

    FolderN = ThisWorkbook.Worksheets("Sheet1").Range("S4")
    sFolderName = Format(FolderN, "mmddyyyy")

    AZ1Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding\On processing\" & sFolderName
    AZ2Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 2\On processing\" & sFolderName
    AZ3Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 3\On processing\" & sFolderName
    AZ5Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 5\On processing\" & sFolderName
    AZ6Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 6\On processing\" & sFolderName
    AZ7Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 7\On processing\" & sFolderName
    AZ8Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 8\On processing\" & sFolderName
    AZ9Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 9\On processing\" & sFolderName
    AZ12Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 12\On processing\" & sFolderName

    If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then

    strDate = InputBox("Please enter last Waterfall date (Date Format MM/DD/YYYY):", "Waterfall Date", Format(Now(), "mm/dd/yyyy"))

        If IsDate(strDate) Then
            strDate = Format(CDate(strDate), "mm/dd/yyyy")
        Else
            MsgBox "Wrong date format, please re run macro"
            Exit Sub
        End If

    End If

    file1 = Dir$(SOME_PATH & "AZBFUND1" & "_*")
    If (Len(file1) > 0) Then

        FPath = OrgPath & file1
        Name1 = "AZBF1 "
        Workbooks.Open FPath

        strFolderExists = Dir(AZ1Path, vbDirectory)

        If strFolderExists = "" Then
            MkDir AZ1Path
        End If

        Call Holdings
        
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ1Path & "\" & Name1 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("USD Collection").Activate

        Call USDColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ1Path & "\" & Name1 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

        If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then
            ActiveWorkbook.Sheets("USD Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ1Path & "\" & Name1 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If

        ActiveWorkbook.Sheets("USD Payment").Activate
        LastRowP = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        USDPMT = ActiveSheet.Range("R" & LastRowP)
        ActiveWorkbook.SaveAs "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding\" & ActiveWorkbook.Name
     
  '-----Count Unapplied Wires------------------------------

        With ActiveWorkbook.Sheets("Additional Data")
            unAZBF1 = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 1
        End With

        ActiveWorkbook.Close

        file1 = Dir$(WSOME_PATH & "AZBF1 Commitment" & "_*.csv")
            If (Len(file1) > 0) Then
                Workbooks.Open WSOME_PATH & file1
                Out1 = ActiveSheet.Range("AA" & Rows.Count).End(xlUp).Row
                SUM1 = WorksheetFunction.Sum(Range("AA2:AA" & Out1))
                ActiveWorkbook.Close
            End If

        Workbooks.Open "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\AZB Funding Daily Reconciliation.xls"
        Workbooks("AZB Funding Daily Reconciliation.xls").Sheets("Summary").Range("D7").Value = WorksheetFunction.Round(Out, 2)
        Workbooks("AZB Funding Daily Reconciliation.xls").Sheets("Summary").Range("D8").Value = WorksheetFunction.Round(Comm, 2)
        Workbooks("AZB Funding Daily Reconciliation.xls").Sheets("Summary").Range("D18").Value = WorksheetFunction.Round(USDPRI, 2)
        Workbooks("AZB Funding Daily Reconciliation.xls").Sheets("Summary").Range("D19").Value = WorksheetFunction.Round(USDINT, 2)
        Workbooks("AZB Funding Daily Reconciliation.xls").Sheets("Summary").Range("D20").Value = WorksheetFunction.Round(USDPMT, 2)
        Workbooks("AZB Funding Daily Reconciliation.xls").Sheets("Summary").Range("C7").Value = WorksheetFunction.Round(SUM1, 2)
        Workbooks("AZB Funding Daily Reconciliation.xls").Sheets("Summary").Range("C8").Value = WorksheetFunction.Round(SUM1, 2)

        ActiveWorkbook.SaveAs AZ1Path & "\" & ActiveWorkbook.Name

    End If

    file2 = Dir$(SOME_PATH & "AZBFUND2" & "_*")
        If (Len(file2) > 0) Then
        
            FPath2 = OrgPath & file2
            Name2 = "AZBF2 "
            Workbooks.Open FPath2
            strFolderExists = Dir(AZ2Path, vbDirectory)

            If strFolderExists = "" Then
                MkDir AZ2Path
            End If

            Call Holdings
            
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ2Path & "\" & Name2 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            ActiveWorkbook.Sheets("USD Collection").Activate

            Call USDColl
            
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ2Path & "\" & Name2 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then
                ActiveWorkbook.Sheets("Payment").Activate
                Call PMT
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ2Path & "\" & Name2 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            End If

            ActiveWorkbook.Sheets("Payment").Activate
            LastRowP1 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
            ActiveWorkbook.SaveAs "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 2\" & ActiveWorkbook.Name
        '-----Count Unapplied Wires------------------------------
            With ActiveWorkbook.Sheets("Additional Data")
                unAZBF2 = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 1
            End With

            ActiveWorkbook.Close

            file2 = Dir$(WSOME_PATH & "AZBF2 Commitment" & "_*.csv")
                If (Len(file2) > 0) Then
                    Workbooks.Open WSOME_PATH & file2
                    Out2 = ActiveSheet.Range("AA" & Rows.Count).End(xlUp).Row
                    SUM2 = WorksheetFunction.Sum(Range("AA2:AA" & Out2))
                    ActiveWorkbook.Close
                End If

            Workbooks.Open "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\AZB Funding 2 Daily Reconciliation.xls"
            
            Workbooks("AZB Funding 2 Daily Reconciliation.xls").Sheets("Summary").Range("D7").Value = WorksheetFunction.Round(Out, 2)
            Workbooks("AZB Funding 2 Daily Reconciliation.xls").Sheets("Summary").Range("D8").Value = WorksheetFunction.Round(Comm, 2)
            Workbooks("AZB Funding 2 Daily Reconciliation.xls").Sheets("Summary").Range("D18").Value = WorksheetFunction.Round(USDPRI, 2)
            Workbooks("AZB Funding 2 Daily Reconciliation.xls").Sheets("Summary").Range("D19").Value = WorksheetFunction.Round(USDINT, 2)
            Workbooks("AZB Funding 2 Daily Reconciliation.xls").Sheets("Summary").Range("D20").Value = WorksheetFunction.Round(USDPMT1, 2)
            Workbooks("AZB Funding 2 Daily Reconciliation.xls").Sheets("Summary").Range("C7").Value = WorksheetFunction.Round(SUM2, 2)
            Workbooks("AZB Funding 2 Daily Reconciliation.xls").Sheets("Summary").Range("C8").Value = WorksheetFunction.Round(SUM2, 2)

            ActiveWorkbook.SaveAs AZ2Path & "\" & ActiveWorkbook.Name

    End If

    file3 = Dir$(SOME_PATH & "AZBFUND3" & "_*")
    If (Len(file3) > 0) Then
            Name3 = "AZBF3 "
            FPath3 = OrgPath & file3
            Workbooks.Open FPath3
            strFolderExists = Dir(AZ3Path, vbDirectory)

        If strFolderExists = "" Then
            MkDir AZ3Path
        End If

        Call Holdings

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ3Path & "\" & Name3 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("USD Collection").Activate

        Call USDColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ3Path & "\" & Name3 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("EUR Collection").Activate

        Call EURColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ3Path & "\" & Name3 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

        If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then
            ActiveWorkbook.Sheets("Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ3Path & "\" & Name3 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            ActiveWorkbook.Sheets("EUR Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ3Path & "\" & Name3 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If

        ActiveWorkbook.Sheets("Payment").Activate
        LastRowP2 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        USDPMT2 = ActiveSheet.Range("R" & LastRowP2)
        ActiveWorkbook.Sheets("EUR Payment").Activate
        LastRowP3 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        EURPMT3 = ActiveSheet.Range("R" & LastRowP3)

        ActiveWorkbook.SaveAs "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 3\" & ActiveWorkbook.Name

'-----Count Unapplied Wires------------------------------

        With ActiveWorkbook.Sheets("Additional Data")
            unAZBF3 = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 3
        End With
        
'---------------------------------------------------------
'MsgBox ("Unapplied: " & unAZBF3)

        ActiveWorkbook.Close
        file3 = Dir$(WSOME_PATH & "AZBF3 Commitment" & "_*.csv")
            If (Len(file3) > 0) Then
                Workbooks.Open WSOME_PATH & file3
                Out3 = ActiveSheet.Range("AA" & Rows.Count).End(xlUp).Row
                SUM3 = WorksheetFunction.Sum(Range("AA2:AA" & Out3))
                ActiveWorkbook.Close
            End If

        Workbooks.Open "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\AZB Funding 3 Daily Reconciliation.xls"

        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("D7").Value = WorksheetFunction.Round(Out, 2)
        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("D8").Value = WorksheetFunction.Round(Comm, 2)
        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("D18").Value = WorksheetFunction.Round(USDPRI, 2)
        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("D19").Value = WorksheetFunction.Round(USDINT, 2)
        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("D22").Value = WorksheetFunction.Round(EURPRI, 2)
        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("D23").Value = WorksheetFunction.Round(EURINT, 2)
        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("D20").Value = WorksheetFunction.Round(USDPMT2, 2)
        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("D24").Value = WorksheetFunction.Round(EURPMT3, 2)
        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("C7").Value = WorksheetFunction.Round(SUM3, 2)
        Workbooks("AZB Funding 3 Daily Reconciliation.xls").Sheets("Summary").Range("C8").Value = WorksheetFunction.Round(SUM3, 2)

        ActiveWorkbook.SaveAs AZ3Path & "\" & ActiveWorkbook.Name
        
    End If

    file5 = Dir$(SOME_PATH & "AZBFUND5" & "_*")
    If (Len(file5) > 0) Then
        FPath5 = OrgPath & file5
        Name5 = "AZBF5 "
        Workbooks.Open FPath5
        strFolderExists = Dir(AZ5Path, vbDirectory)

        If strFolderExists = "" Then
            MkDir AZ5Path
        End If

        Call Holdings

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ5Path & "\" & Name5 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("USD Collection").Activate

        Call USDColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ5Path & "\" & Name5 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

        If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then
            ActiveWorkbook.Sheets("USD Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ5Path & "\" & Name5 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If

        ActiveWorkbook.Sheets("USD Payment").Activate
        LastRowP5 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        USDPMT5 = ActiveSheet.Range("R" & LastRowP5)

        ActiveWorkbook.SaveAs "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 5\" & ActiveWorkbook.Name

'-----Count Unapplied Wires------------------------------

        With ActiveWorkbook.Sheets("Additional Data")
            unAZBF5 = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 1
        End With

        ActiveWorkbook.Close

        file5 = Dir$(WSOME_PATH & "AZBF5 Commitment" & "_*.csv")
            If (Len(file5) > 0) Then
                Workbooks.Open WSOME_PATH & file5
                Out5 = ActiveSheet.Range("AA" & Rows.Count).End(xlUp).Row
                SUM5 = WorksheetFunction.Sum(Range("AA2:AA" & Out5))
                ActiveWorkbook.Close
            End If
            
        Workbooks.Open "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\AZB Funding 5 Daily Reconciliation.xls"

        Workbooks("AZB Funding 5 Daily Reconciliation.xls").Sheets("Summary").Range("D7").Value = WorksheetFunction.Round(Out, 2)
        Workbooks("AZB Funding 5 Daily Reconciliation.xls").Sheets("Summary").Range("D8").Value = WorksheetFunction.Round(Comm, 2)
        Workbooks("AZB Funding 5 Daily Reconciliation.xls").Sheets("Summary").Range("D18").Value = WorksheetFunction.Round(USDPRI, 2)
        Workbooks("AZB Funding 5 Daily Reconciliation.xls").Sheets("Summary").Range("D19").Value = WorksheetFunction.Round(USDINT, 2)
        Workbooks("AZB Funding 5 Daily Reconciliation.xls").Sheets("Summary").Range("D20").Value = WorksheetFunction.Round(USDPMT5, 2)
        Workbooks("AZB Funding 5 Daily Reconciliation.xls").Sheets("Summary").Range("C7").Value = WorksheetFunction.Round(SUM5, 2)
        Workbooks("AZB Funding 5 Daily Reconciliation.xls").Sheets("Summary").Range("C8").Value = WorksheetFunction.Round(SUM5, 2)

        ActiveWorkbook.SaveAs AZ5Path & "\" & ActiveWorkbook.Name

    End If

    file6 = Dir$(SOME_PATH & "AZBFUND6" & "_*")
    If (Len(file6) > 0) Then
        Name6 = "AZBF6 "
        FPath6 = OrgPath & file6
        Workbooks.Open FPath6
        strFolderExists = Dir(AZ6Path, vbDirectory)
        
        If strFolderExists = "" Then
            MkDir AZ6Path
        End If

        Call Holdings

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ6Path & "\" & Name6 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("USD Collection").Activate

        Call USDColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ6Path & "\" & Name6 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("EUR Collection").Activate

        Call EURColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ6Path & "\" & Name6 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

        If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then
            ActiveWorkbook.Sheets("USD Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ6Path & "\" & Name6 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            ActiveWorkbook.Sheets("EUR Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ6Path & "\" & Name6 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If

        ActiveWorkbook.Sheets("USD Payment").Activate
        LastRowP6 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        USDPMT6 = ActiveSheet.Range("R" & LastRowP6)
        ActiveWorkbook.Sheets("EUR Payment").Activate
        LastERowP6 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        EURPMT6 = ActiveSheet.Range("R" & LastERowP6)

        ActiveWorkbook.SaveAs "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 6\" & ActiveWorkbook.Name

'-----Count Unapplied Wires------------------------------
        With ActiveWorkbook.Sheets("Additional Data")
            unAZBF6 = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 1
        End With

        ActiveWorkbook.Close

        file6 = Dir$(WSOME_PATH & "AZBF6 Commitment" & "_*.csv")
        If (Len(file6) > 0) Then
            Workbooks.Open WSOME_PATH & file6
            Out6 = ActiveSheet.Range("AA" & Rows.Count).End(xlUp).Row
            SUM6 = WorksheetFunction.Sum(Range("AA2:AA" & Out6))
            ActiveWorkbook.Close
        End If

        Workbooks.Open "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\AZB Funding 6 Daily Reconciliation.xls"

        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("D7").Value = WorksheetFunction.Round(Out, 2)
        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("D8").Value = WorksheetFunction.Round(Comm, 2)
        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("D18").Value = WorksheetFunction.Round(USDPRI, 2)
        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("D19").Value = WorksheetFunction.Round(USDINT, 2)
        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("D22").Value = WorksheetFunction.Round(EURPRI, 2)
        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("D23").Value = WorksheetFunction.Round(EURINT, 2)
        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("D20").Value = WorksheetFunction.Round(USDPMT6, 2)
        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("D24").Value = WorksheetFunction.Round(EURPMT6, 2)
        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("C7").Value = WorksheetFunction.Round(SUM6, 2)
        Workbooks("AZB Funding 6 Daily Reconciliation.xls").Sheets("Summary").Range("C8").Value = WorksheetFunction.Round(SUM6, 2)

        ActiveWorkbook.SaveAs AZ6Path & "\" & ActiveWorkbook.Name

    End If

    file7 = Dir$(SOME_PATH & "AZBFUND7" & "_*")
    If (Len(file7) > 0) Then
        Name7 = "AZBF7 "
        FPath7 = OrgPath & file7
        Workbooks.Open FPath7
        strFolderExists = Dir(AZ7Path, vbDirectory)
        
        If strFolderExists = "" Then
            MkDir AZ7Path
        End If

        Call Holdings

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ7Path & "\" & Name7 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("USD Collection").Activate

        Call USDColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ7Path & "\" & Name7 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("EUR Collection").Activate

        Call EURColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ7Path & "\" & Name7 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

        If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then
            ActiveWorkbook.Sheets("USD Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ7Path & "\" & Name7 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            ActiveWorkbook.Sheets("EUR Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ7Path & "\" & Name7 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If

        ActiveWorkbook.Sheets("USD Payment").Activate
        LastRowP7 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        USDPMT7 = ActiveSheet.Range("R" & LastRowP7)
        ActiveWorkbook.Sheets("EUR Payment").Activate
        LastERowP7 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        EURPMT7 = ActiveSheet.Range("R" & LastERowP7)
        ActiveWorkbook.SaveAs "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 7\" & ActiveWorkbook.Name

'-----Count Unapplied Wires------------------------------
        With ActiveWorkbook.Sheets("Additional Data")
            unAZBF7 = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 1
        End With

        ActiveWorkbook.Close

        file7 = Dir$(WSOME_PATH & "AZBF7 Commitment" & "_*.csv")
            If (Len(file7) > 0) Then
                Workbooks.Open WSOME_PATH & file7
                Out7 = ActiveSheet.Range("AA" & Rows.Count).End(xlUp).Row
                SUM7 = WorksheetFunction.Sum(Range("AA2:AA" & Out7))
                ActiveWorkbook.Close
            End If

        Workbooks.Open "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\AZB Funding 7 Daily Reconciliation.xls"

        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("D7").Value = WorksheetFunction.Round(Out, 2)
        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("D8").Value = WorksheetFunction.Round(Comm, 2)
        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("D18").Value = WorksheetFunction.Round(USDPRI, 2)
        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("D19").Value = WorksheetFunction.Round(USDINT, 2)
        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("D22").Value = WorksheetFunction.Round(EURPRI, 2)
        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("D23").Value = WorksheetFunction.Round(EURINT, 2)
        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("D20").Value = WorksheetFunction.Round(USDPMT7, 2)
        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("D24").Value = WorksheetFunction.Round(EURPMT7, 2)
        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("C7").Value = WorksheetFunction.Round(SUM7, 2)
        Workbooks("AZB Funding 7 Daily Reconciliation.xls").Sheets("Summary").Range("C8").Value = WorksheetFunction.Round(SUM7, 2)

        ActiveWorkbook.SaveAs AZ7Path & "\" & ActiveWorkbook.Name

    End If

    file8 = Dir$(SOME_PATH & "AZBFUND8" & "_*")
    If (Len(file8) > 0) Then
        Name8 = "AZBF8 "
        FPath8 = OrgPath & file8
        Workbooks.Open FPath8
        strFolderExists = Dir(AZ8Path, vbDirectory)

        If strFolderExists = "" Then
            MkDir AZ8Path
        End If

        Call Holdings

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ8Path & "\" & Name8 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("USD Collection").Activate

        Call USDColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ8Path & "\" & Name8 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("EUR Collection").Activate

        Call EURColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ8Path & "\" & Name8 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
        If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then
            ActiveWorkbook.Sheets("USD Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ8Path & "\" & Name8 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            ActiveWorkbook.Sheets("EUR Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ8Path & "\" & Name8 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If

        ActiveWorkbook.Sheets("USD Payment").Activate
        LastRowP8 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        USDPMT8 = ActiveSheet.Range("R" & LastRowP8)
        ActiveWorkbook.Sheets("EUR Payment").Activate
        LastRowP8 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        EURPMT8 = ActiveSheet.Range("R" & LastRowP8)

        ActiveWorkbook.SaveAs "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 8\" & ActiveWorkbook.Name

'-----Count Unapplied Wires------------------------------
        With ActiveWorkbook.Sheets("Additional Data")
            unAZBF8 = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 1
        End With

        ActiveWorkbook.Close

        file8 = Dir$(WSOME_PATH & "AZBF8 Commitment" & "_*.csv")
            If (Len(file8) > 0) Then
                Workbooks.Open WSOME_PATH & file8
                Out8 = ActiveSheet.Range("AA" & Rows.Count).End(xlUp).Row
                SUM8 = WorksheetFunction.Sum(Range("AA2:AA" & Out8))
                ActiveWorkbook.Close
            End If

        Workbooks.Open "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\AZB Funding 8 Daily Reconciliation.xls"

        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("D7").Value = WorksheetFunction.Round(Out, 2)
        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("D8").Value = WorksheetFunction.Round(Comm, 2)
        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("D18").Value = WorksheetFunction.Round(USDPRI, 2)
        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("D19").Value = WorksheetFunction.Round(USDINT, 2)
        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("D22").Value = WorksheetFunction.Round(EURPRI, 2)
        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("D23").Value = WorksheetFunction.Round(EURINT, 2)
        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("D20").Value = WorksheetFunction.Round(USDPMT8, 2)
        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("D24").Value = WorksheetFunction.Round(EURPMT8, 2)
        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("C7").Value = WorksheetFunction.Round(SUM8, 2)
        Workbooks("AZB Funding 8 Daily Reconciliation.xls").Sheets("Summary").Range("C8").Value = WorksheetFunction.Round(SUM8, 2)

        ActiveWorkbook.SaveAs AZ8Path & "\" & ActiveWorkbook.Name

    End If

file9 = Dir$(SOME_PATH & "AZBFUND9" & "_*")
    If (Len(file9) > 0) Then
        Name9 = "AZBF9 "
        FPath9 = OrgPath & file9
        Workbooks.Open FPath9
        strFolderExists = Dir(AZ9Path, vbDirectory)

        If strFolderExists = "" Then
            MkDir AZ9Path
        End If

        Call Holdings

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ9Path & "\" & Name9 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("USD Collection Accounts").Activate

        Call USDColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ9Path & "\" & Name9 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("EUR Collection Accounts").Activate

        Call EURColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ9Path & "\" & Name9 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

'------------No more JPY--------------------
        'ActiveWorkbook.Sheets("JPY Collection Accounts").Activate
        'Call JPYColl
        'ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=AZ9Path & "\" & Name9 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

        If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then
            ActiveWorkbook.Sheets("EUR Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ9Path & "\" & Name9 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            ActiveWorkbook.Sheets("USD Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ9Path & "\" & Name9 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            'ActiveWorkbook.Sheets("JPY Payment").Activate
            'Call PMT
            'ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=AZ9Path & "\" & Name9 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If

        ActiveWorkbook.Sheets("USD Payment").Activate
        LastRowP3 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        USDPMT3 = ActiveSheet.Range("R" & LastRowP3)
        ActiveWorkbook.Sheets("EUR Payment").Activate
        LastRowP4 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        EURPMT4 = ActiveSheet.Range("R" & LastRowP4)
        'ActiveWorkbook.Sheets("JPY Payment").Activate
        'LastRowP5 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        'JPYPMT = ActiveSheet.Range("R" & LastRowP5)

        ActiveWorkbook.SaveAs "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 9\" & ActiveWorkbook.Name

'-----Count Unapplied Wires------------------------------
        With ActiveWorkbook.Sheets("Additional Data")
            unAZBF9 = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 1
        End With
        
        ActiveWorkbook.Close

        file9 = Dir$(WSOME_PATH & "AZBF9 Commitment" & "_*.csv")
        If (Len(file9) > 0) Then
            Workbooks.Open WSOME_PATH & file9
            Out9 = ActiveSheet.Range("AA" & Rows.Count).End(xlUp).Row
            SUM9 = WorksheetFunction.Sum(Range("AA2:AA" & Out9))
            ActiveWorkbook.Close
        End If

        Workbooks.Open "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\AZB Funding 9 Daily Reconciliation.xls"

        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D7").Value = WorksheetFunction.Round(Out, 2)
        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D8").Value = WorksheetFunction.Round(Comm, 2)
        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D18").Value = WorksheetFunction.Round(USDPRI, 2)
        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D19").Value = WorksheetFunction.Round(USDINT, 2)
        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D22").Value = WorksheetFunction.Round(EURPRI, 2)
        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D23").Value = WorksheetFunction.Round(EURINT, 2)
        'Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D26").Value = WorksheetFunction.Round(JPYPRI, 2)
        'Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D27").Value = WorksheetFunction.Round(JPYINT, 2)
        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D20").Value = WorksheetFunction.Round(USDPMT3, 2)
        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D24").Value = WorksheetFunction.Round(EURPMT4, 2)
        'Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("D28").Value = WorksheetFunction.Round(JPYPMT, 2)
        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("C7").Value = WorksheetFunction.Round(SUM9, 2)
        Workbooks("AZB Funding 9 Daily Reconciliation.xls").Sheets("Summary").Range("C8").Value = WorksheetFunction.Round(SUM9, 2)

        ActiveWorkbook.SaveAs AZ9Path & "\" & ActiveWorkbook.Name

    End If

    file12 = Dir$(SOME_PATH & "AZBFND12" & "_*")
    If (Len(file12) > 0) Then
        Name12 = "AZBF12 "
        FPath12 = OrgPath & file12
        Workbooks.Open FPath12
        strFolderExists = Dir(AZ12Path, vbDirectory)

        If strFolderExists = "" Then
            MkDir AZ12Path
        End If

        Call Holdings

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ12Path & "\" & Name12 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("USD Collection Accounts").Activate

        Call USDColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ12Path & "\" & Name12 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        ActiveWorkbook.Sheets("EUR Collection Accounts").Activate

        Call EURColl

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ12Path & "\" & Name12 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

        If ThisWorkbook.Worksheets("Sheet1").Range("D6") = "Yes" Then
            ActiveWorkbook.Sheets("EUR Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ12Path & "\" & Name12 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            ActiveWorkbook.Sheets("USD Payment").Activate
            Call PMT
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=AZ12Path & "\" & Name12 & ActiveSheet.Name & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        End If

        ActiveWorkbook.Sheets("USD Payment").Activate
        LastRowP4 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        USDPMT4 = ActiveSheet.Range("R" & LastRowP4)
        ActiveWorkbook.Sheets("EUR Payment").Activate
        LastRowP5 = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
        EURPMT5 = ActiveSheet.Range("R" & LastRowP5)

        ActiveWorkbook.SaveAs "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 12\" & ActiveWorkbook.Name

'-----Count Unapplied Wires------------------------------
        With ActiveWorkbook.Sheets("Additional Data")
            unAZBF12 = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 1
        End With

        ActiveWorkbook.Close

        file12 = Dir$(WSOME_PATH & "AZBF12 Commitment" & "*.csv")
        If (Len(file12) > 0) Then
            Workbooks.Open WSOME_PATH & file12
            Out12 = ActiveSheet.Range("AA" & Rows.Count).End(xlUp).Row
            SUM12 = WorksheetFunction.Sum(Range("AA2:AA" & Out12))
            ActiveWorkbook.Close
        End If

        Workbooks.Open "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\Template\AZB Funding 12 Daily Reconciliation.xls"

        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("D7").Value = WorksheetFunction.Round(Out, 2)
        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("D8").Value = WorksheetFunction.Round(Comm, 2)
        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("D18").Value = WorksheetFunction.Round(USDPRI, 2)
        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("D19").Value = WorksheetFunction.Round(USDINT, 2)
        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("D22").Value = WorksheetFunction.Round(EURPRI, 2)
        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("D23").Value = WorksheetFunction.Round(EURINT, 2)
        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("D20").Value = WorksheetFunction.Round(USDPMT4, 2)
        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("D24").Value = WorksheetFunction.Round(EURPMT5, 2)
        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("C7").Value = WorksheetFunction.Round(SUM12, 2)
        Workbooks("AZB Funding 12 Daily Reconciliation.xls").Sheets("Summary").Range("C8").Value = WorksheetFunction.Round(SUM12, 2)

        ActiveWorkbook.SaveAs AZ12Path & "\" & ActiveWorkbook.Name

    End If

    Call WSOCommit

    ThisWorkbook.Worksheets("Sheet1").Range("B4").ClearContents
    ThisWorkbook.Worksheets("Sheet1").Range("D4").ClearContents
    ThisWorkbook.Worksheets("Sheet1").Range("D6").ClearContents

    Call DELExcel

    MsgBox "UNAPPLIED WIRES" & vbCrLf & vbCrLf & "AZBF1: " & unAZBF1 & vbCrLf & "AZBF2: " & unAZBF2 & vbCrLf & "AZBF3: " & unAZBF3 & vbCrLf & "AZBF5: " & unAZBF5 & vbCrLf & "AZBF6: " & unAZBF6 & vbCrLf & "AZBF7: " & unAZBF7 & vbCrLf & "AZBF8: " & unAZBF8 & vbCrLf & "AZBF9: " & unAZBF9 & vbCrLf & "AZBF12: " & unAZBF12

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

'Setup Holding tab to grab values & setup for print
Sub Holdings()

        ActiveWorkbook.Sheets("US Bank Holdings").Activate
        ActiveSheet.Columns("A:B").Hidden = True
        ActiveSheet.Columns("D").Hidden = True
        ActiveSheet.Columns("F:I").Hidden = True
        ActiveSheet.Columns("L").Hidden = True
        ActiveSheet.Columns("P:S").Hidden = True
        ActiveSheet.Columns("X:AG").Hidden = True
        ActiveSheet.Columns("K").ColumnWidth = 25
        ActiveSheet.Columns("AI").ColumnWidth = 16

        ThisWorkbook.Worksheets("Sheet1").Range("J2").Copy ActiveSheet.Range("AI5")

        MaxRow = ActiveSheet.Range("AG" & Rows.Count).End(xlUp).Row

        ActiveSheet.Range("AI6:AI" & MaxRow).FormulaR1C1 = "=RC[-22]*RC[-20]"
        ActiveSheet.Range("AI6:AI" & MaxRow).NumberFormatLocal = "#,##0.00"
 
        With Range("AI6").End(xlDown).Offset(2, 0) = "=SUM(" & Range(.Address, .End(xlDown)).Address(False, False) & ")"
            .End(xlDown).Offset(2, 0).NumberFormatLocal = "#,##0.00"
        End With

        ThisWorkbook.Worksheets("Sheet1").Range("D4").Copy ActiveSheet.Range("J" & MaxRow + 3)
        ActiveSheet.Range("J" & MaxRow + 3).VerticalAlignment = xlCenter
        ActiveSheet.Range("J" & MaxRow + 3 & ":" & "J" & MaxRow + 5).Merge
        ActiveSheet.Range("J" & MaxRow + 3 & ":" & "J" & MaxRow + 5).Borders.LineStyle = True

        ThisWorkbook.Worksheets("Sheet1").Range("K5").Copy ActiveSheet.Range("K" & MaxRow + 3)
        ThisWorkbook.Worksheets("Sheet1").Range("K6").Copy ActiveSheet.Range("K" & MaxRow + 4)
        ThisWorkbook.Worksheets("Sheet1").Range("K7").Copy ActiveSheet.Range("K" & MaxRow + 5)
        ThisWorkbook.Worksheets("Sheet1").Range("L4").Copy ActiveSheet.Range("M" & MaxRow + 2)
        ThisWorkbook.Worksheets("Sheet1").Range("L5").Copy ActiveSheet.Range("M" & MaxRow + 3)
        ThisWorkbook.Worksheets("Sheet1").Range("L6").Copy ActiveSheet.Range("M" & MaxRow + 4)
        ThisWorkbook.Worksheets("Sheet1").Range("L7").Copy ActiveSheet.Range("M" & MaxRow + 5)

        ActiveSheet.Range("M" & MaxRow + 3).HorizontalAlignment = xlRight
        ActiveSheet.Range("M" & MaxRow + 4).HorizontalAlignment = xlRight
        ActiveSheet.Range("M" & MaxRow + 5).HorizontalAlignment = xlRight

        With Range("M6").End(xlDown).Offset(1, 0) = "=SUM(" & Range(.Address, .End(xlDown)).Address(False, False) & ")"
            'Stuff
        End With

        Range("M" & MaxRow + 3) = Range("M" & MaxRow + 3).Offset(-2, 0).Value
        Range("M" & MaxRow + 3).Offset(-2, 0).ClearContents
        Range("M" & MaxRow + 3).NumberFormatLocal = "#,##0.00"
        Range("M" & MaxRow + 5) = Range("AI6").End(xlDown).Offset(2, 0)

        Comm = Range("M" & MaxRow + 3)
        Out = Range("M" & MaxRow + 4)

        Application.PrintCommunication = False

        With ActiveSheet.PageSetup
            .PrintArea = ActiveSheet.Range("C5:AI" & MaxRow + 5).Address
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .CenterFooter = "&F &A"
            .LeftMargin = Application.InchesToPoints(0.7)
            .RightMargin = Application.InchesToPoints(0.7)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 0
            .CenterHorizontally = True
        End With

        Application.PrintCommunication = True

End Sub

'Setup USD tab to grab values & setup for print
Sub USDColl()

    LastRowS = ActiveSheet.Range("S" & Rows.Count).End(xlUp).Row
    USDINT = ActiveSheet.Range("S" & LastRowS)
    LastRowR = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
    USDPRI = ActiveSheet.Range("R" & LastRowR)

    USDate = Format(CDate(Workbooks("Daily RecV7.xlsm").Worksheets("Sheet1").Range("B4")), "mm/dd/yyyy")

    ActiveSheet.Range("A:Y").AutoFilter Field:=5, Criteria1:=">=" & USDate
    ActiveSheet.Columns("A").Hidden = True
    ActiveSheet.Columns("F:I").Hidden = True
    ActiveSheet.Columns("L:M").Hidden = True
    ActiveSheet.Columns("P").Hidden = True
    ActiveSheet.Columns("T:X").Hidden = True

    MaxRow4 = ActiveSheet.Range("E" & Rows.Count).End(xlUp).Row

    With ActiveSheet.PageSetup
        .PrintArea = ActiveSheet.Range("B5:S" & MaxRow4).Address
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .CenterFooter = "&F &A"
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With

End Sub

'Setup EUR tab to grab values & setup for print
Sub EURColl()

    LastERowS = ActiveSheet.Range("S" & Rows.Count).End(xlUp).Row
    EURINT = ActiveSheet.Range("S" & LastERowS)
    LastERowR = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
    EURPRI = ActiveSheet.Range("R" & LastERowR)
    EUDate = Format(CDate(Workbooks("Daily RecV7.xlsm").Worksheets("Sheet1").Range("B4")), "mm/dd/yyyy")

    ActiveSheet.Range("A:Y").AutoFilter Field:=5, Criteria1:=">=" & EUDate

    LastDate = Range("E5").CurrentRegion.SpecialCells(xlCellTypeVisible).Rows.Count

    If LastDate = 1 Then
        ActiveSheet.ShowAllData
        NewDT = ActiveSheet.Range("E" & Rows.Count).End(xlUp).Row
        NewDate = ActiveSheet.Range("E" & NewDT - 1)
        ActiveSheet.Range("A:Y").AutoFilter Field:=5, Criteria1:=">=" & NewDate
    End If

    ActiveSheet.Columns("A").Hidden = True
    ActiveSheet.Columns("F:I").Hidden = True
    ActiveSheet.Columns("L:M").Hidden = True
    ActiveSheet.Columns("P").Hidden = True
    ActiveSheet.Columns("T:X").Hidden = True

    MaxRow4 = ActiveSheet.Range("E" & Rows.Count).End(xlUp).Row

    With ActiveSheet.PageSetup
        .PrintArea = ActiveSheet.Range("B5:S" & MaxRow4).Address
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .CenterFooter = "&F &A"
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With

End Sub

'Setup JPY tab to grab values & setup for print
Sub JPYColl()

    LastERowS = ActiveSheet.Range("S" & Rows.Count).End(xlUp).Row
    JPYINT = ActiveSheet.Range("S" & LastERowS)
    LastERowR = ActiveSheet.Range("R" & Rows.Count).End(xlUp).Row
    JPYPRI = ActiveSheet.Range("R" & LastERowR)
    JPDate = Format(CDate(Workbooks("Daily RecV7.xlsm").Worksheets("Sheet1").Range("B4")), "mm/dd/yyyy")

    ActiveSheet.Range("A:Y").AutoFilter Field:=5, Criteria1:=">=" & JPDate

    LastJDate = Range("E5").CurrentRegion.SpecialCells(xlCellTypeVisible).Rows.Count

    If LastJDate = 1 Then
        ActiveSheet.ShowAllData
        NewJDT = ActiveSheet.Range("E" & Rows.Count).End(xlUp).Row
        NewJDate = ActiveSheet.Range("E" & NewJDT - 1)
        ActiveSheet.Range("A:Y").AutoFilter Field:=5, Criteria1:=">=" & NewJDate
    End If

    ActiveSheet.Columns("A").Hidden = True
    ActiveSheet.Columns("F:I").Hidden = True
    ActiveSheet.Columns("L:M").Hidden = True
    ActiveSheet.Columns("P").Hidden = True
    ActiveSheet.Columns("T:X").Hidden = True

    MaxRow4 = ActiveSheet.Range("E" & Rows.Count).End(xlUp).Row

    With ActiveSheet.PageSetup
        .PrintArea = ActiveSheet.Range("B5:S" & MaxRow4).Address
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .CenterFooter = "&F &A"
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With

End Sub

'Setup PMT tabs for print
Sub PMT()

    ActiveSheet.Range("A:Y").AutoFilter Field:=5, Operator:=xlOr, Criteria2:=">=" & strDate

    'Array(1, strDate)

    ActiveSheet.Columns("A").Hidden = True
    ActiveSheet.Columns("F:I").Hidden = True
    ActiveSheet.Columns("L:M").Hidden = True
    ActiveSheet.Columns("P").Hidden = True
    ActiveSheet.Columns("T:X").Hidden = True

    MaxRow4 = ActiveSheet.Range("E" & Rows.Count).End(xlUp).Row

    Application.PrintCommunication = False
    
    With ActiveSheet.PageSetup
              .PrintArea = ActiveSheet.Range("B5:S" & MaxRow4).Address
              .Orientation = xlLandscape
              .PaperSize = xlPaperA4
              .CenterFooter = "&F &A"
              .Zoom = Falsecreate
              .FitToPagesWide = 1
              .FitToPagesTall = 1
              .CenterHorizontally = True
    End With

    Application.PrintCommunication = True

End Sub

'Remove files for next batch
Sub DELExcel()

    On Error Resume Next
    Kill "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Macro\*.xlsm*"

    On Error Resume Next
    Kill "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\WSO Reports\\*.csv*"

End Sub

'Ship autodropped reports to new file path
Sub WSOCommit()

    Const WSOME_PATH As String = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\WSO Reports\"

    AZ1Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding\On processing\" & sFolderName
    AZ2Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 2\On processing\" & sFolderName
    AZ3Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 3\On processing\" & sFolderName
    AZ5Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 5\On processing\" & sFolderName
    AZ6Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 6\On processing\" & sFolderName
    AZ7Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 7\On processing\" & sFolderName
    AZ8Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 8\On processing\" & sFolderName
    AZ9Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 9\On processing\" & sFolderName
    AZ12Path = "\\ads.aozora.lan\nyi\SH01\ANA_01\BMG\Daily Operation\Daily Materials\Trustee Report (Daily)\AZB Funding 12\On processing\" & sFolderName

    FolderN = ThisWorkbook.Worksheets("Sheet1").Range("S4")
    sFolderName = Format(FolderN, "mmddyyyy")

    file1 = Dir$(WSOME_PATH & "AZBF1 Commitment" & "_*.pdf")
        If (Len(file1) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file1 As AZ1Path & sFolderName & "\" & file1
        End If

    file2 = Dir$(WSOME_PATH & "AZBF2 Commitment" & "_*.pdf")
        If (Len(file2) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file2 As AZ2Path & sFolderName & "\" & file2
        End If

    file3 = Dir$(WSOME_PATH & "AZBF3 Commitment" & "_*.pdf")
        If (Len(file3) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file3 As AZ3Path & sFolderName & "\" & file3
        End If

    file5 = Dir$(WSOME_PATH & "AZBF5 Commitment" & "_*.pdf")
        If (Len(file5) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file5 As AZ5Path & sFolderName & "\" & file5
        End If

    file6 = Dir$(WSOME_PATH & "AZBF6 Commitment" & "_*.pdf")
        If (Len(file6) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file6 As AZ6Path & sFolderName & "\" & file6
        End If

    file7 = Dir$(WSOME_PATH & "AZBF7 Commitment" & "_*.pdf")
        If (Len(file7) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file7 As AZ7Path & sFolderName & "\" & file7
        End If

    file8 = Dir$(WSOME_PATH & "AZBF8 Commitment" & "_*.pdf")
        If (Len(file8) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file8 As AZ8Path & sFolderName & "\" & file8
        End If

    file9 = Dir$(WSOME_PATH & "AZBF9 Commitment" & "_*.pdf")
        If (Len(file9) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file9 As AZ9Path & sFolderName & "\" & file9
        End If

    file12 = Dir$(WSOME_PATH & "AZBF12 Commitment" & "_*.pdf")
        If (Len(file12) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file12 As AZ12Path & sFolderName & "\" & file12
        End If

'------------------------------------------------------------------------------

    file1 = Dir$(WSOME_PATH & "Upcoming1_" & "*.pdf")
        If (Len(file1) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file1 As AZ1Path & sFolderName & "\" & file1
        End If

    file2 = Dir$(WSOME_PATH & "Upcoming2_" & "*.pdf")
        If (Len(file2) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file2 As AZ2Path & sFolderName & "\" & file2
        End If

    file3 = Dir$(WSOME_PATH & "Upcoming3_" & "*.pdf")
        If (Len(file3) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file3 As AZ3Path & sFolderName & "\" & file3
        End If

    file5 = Dir$(WSOME_PATH & "Upcoming5_" & "*.pdf")
        If (Len(file5) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file5 As AZ5Path & sFolderName & "\" & file5
        End If

    file6 = Dir$(WSOME_PATH & "Upcoming6_" & "*.pdf")
        If (Len(file6) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file6 As AZ6Path & sFolderName & "\" & file6
        End If

    file7 = Dir$(WSOME_PATH & "Upcoming7_" & "*.pdf")
        If (Len(file7) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file7 As AZ7Path & sFolderName & "\" & file7
        End If

    file8 = Dir$(WSOME_PATH & "Upcoming8_" & "*.pdf")
        If (Len(file8) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file8 As AZ8Path & sFolderName & "\" & file8
        End If

    file9 = Dir$(WSOME_PATH & "Upcoming9_" & "*.pdf")
        If (Len(file9) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file9 As AZ9Path & sFolderName & "\" & file9
        End If

    file12 = Dir$(WSOME_PATH & "Upcoming12_" & "*.pdf")
        If (Len(file12) > 0) Then
            On Error Resume Next
            Name WSOME_PATH & file12 As AZ12Path & sFolderName & "\" & file12
        End If

End Sub

'Set printer as excel has my printer locked
Public Function GetPrinterFullName(Printer As String) As String


    Const HKEY_CURRENT_USER = &H80000001
    Dim regobj As Object
    Dim aTypes As Variant
    Dim aDevices As Variant
    Dim vDevice As Variant
    Dim sValue As String
    Dim v As Variant
    Dim sLocaleOn As String

    ' get locale "on" from current activeprinter
    v = Split(Application.ActivePrinter, Space(1))
    sLocaleOn = Space(1) & CStr(v(UBound(v) - 1)) & Space(1)

    ' connect to WMI registry provider on current machine with current user
    Set regobj = GetObject("WINMGMTS:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    ' get the Devices from the registry
    regobj.EnumValues HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", aDevices, aTypes

    ' find Printer and create full name
    For Each vDevice In aDevices

        ' get port of device
        regobj.GetStringValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", vDevice, sValue

        ' select device
        If Left(vDevice, Len(Printer)) = Printer Then ' match!
            ' create localized printername
            GetPrinterFullName = vDevice & sLocaleOn & Split(sValue, ",")(1)
            Exit Function
        End If

    Next

    ' at this point no match found
    GetPrinterFullName = vbNullString

End Function

'Count rows of wires
Sub CountUnappliedWires()

    With ActiveWorkbook.Sheets("Additional Data")
        recct = .Range("AB4", .Range("AB" & .Rows.Count).End(xlUp)).Rows.Count - 1
    End With

End Sub

