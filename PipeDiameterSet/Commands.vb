Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime


Public Class Commands

    <CommandMethod("PipeWithDiameter")> _
    Public Sub PipeWithDiameter()
        '' Get the current database and start the Transaction Manager
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")


        '' Prompt for the start point
        pPtOpts.Message = vbLf & "Enter the start point of the line: "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptStart As Point3d = pPtRes.Value

        '' Exit if the user presses ESC or cancels the command
        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' Prompt for the end point
        pPtOpts.Message = vbLf & "Enter the end point of the line: "
        pPtOpts.UseBasePoint = True
        pPtOpts.BasePoint = ptStart
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptEnd As Point3d = pPtRes.Value

        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim acBlkTbl As BlockTable
            Dim acBlkTblRec As BlockTableRecord

            '' Open Model space for write
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, _
                                         OpenMode.ForRead)

            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                            OpenMode.ForWrite)

            '' Define the new line
            Dim acLine As Line = New Line(ptStart, ptEnd)
            acLine.SetDatabaseDefaults()

            '' Add the line to the drawing
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            '' Zoom to the extents or limits of the drawing
            'acDoc.SendStringToExecute("._zoom _all ", True, False, False)
            '///////////////////////////////////////////////////////////////////
            Dim pStrOpts0 As PromptStringOptions = New PromptStringOptions(vbLf & _
                                                              "Enter the Text ∅")
            pStrOpts0.AllowSpaces = True
            Dim pStrRes0 As PromptResult = acDoc.Editor.GetString(pStrOpts0)
            'Dim textString As String = "Left"
            Dim textString As String = "%%C" & " " & pStrRes0.StringResult
            Dim textAlign As Integer = TextHorizontalMode.TextCenter
            Dim acText As DBText = New DBText()
            'acText.Height = 5
            '////////////
            Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & _
                                                               "Enter Font Size<" & FontSize & ">: ")
            Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)

            If pStrRes.StringResult = "" Then

            Else
                FontSize = CDbl(pStrRes.StringResult.ToString)
            End If
            'Dim s As New Form1
            's.Show()

            acText.Height = FontSize
            '////////////
            acText.TextString = textString
            acText.HorizontalMode = TextHorizontalMode.TextCenter
            Dim pp As Point3d = New Point3d(((ptEnd.X + ptStart.X) / 2), ((ptEnd.Y + ptStart.Y) / 2), ptEnd.Z)
            acText.Position = pp
            If acText.HorizontalMode <> TextHorizontalMode.TextLeft Then
                acText.AlignmentPoint = pp
                acText.Rotation = Math.Atan((ptEnd.Y - ptStart.Y) / (ptEnd.X - ptStart.X))
            End If

            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            ' '' Create a point over the alignment point of the text
            'Dim acPtAlign As Point3d = ptEnd
            'Dim acPtIns As Point3d = ptStart
            'Dim acPoint As DBPoint = New DBPoint(acPtAlign)
            'acPoint.ColorIndex = 1

            'acBlkTblRec.AppendEntity(acPoint)
            'acTrans.AddNewlyCreatedDBObject(acPoint, True)

            ' '' Adjust the insertion and alignment points
            'acPtIns = New Point3d(acPtIns.X, acPtIns.Y + 3, 0)
            'acPtAlign = acPtIns

            ' '' Set the point style to crosshair
            'Application.SetSystemVariable("PDMODE", 2)
            '///////////////////////////////////////////////////////////////////
            '' Commit the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub

    <CommandMethod("PipeWithDiaWithOption")> _
    Public Sub PipeWithDiaWithOption()
        '' Get the current database and start the Transaction Manager
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")


        '' Prompt for the start point
        pPtOpts.Message = vbLf & "Enter the start point of the line: "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptStart As Point3d = pPtRes.Value

        '' Exit if the user presses ESC or cancels the command
        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' Prompt for the end point
        pPtOpts.Message = vbLf & "Enter the end point of the line: "
        pPtOpts.UseBasePoint = True
        pPtOpts.BasePoint = ptStart
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptEnd As Point3d = pPtRes.Value

        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim acBlkTbl As BlockTable
            Dim acBlkTblRec As BlockTableRecord

            '' Open Model space for write
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, _
                                         OpenMode.ForRead)

            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                            OpenMode.ForWrite)

            '' Define the new line
            Dim acLine As Line = New Line(ptStart, ptEnd)
            acLine.SetDatabaseDefaults()

            '' Add the line to the drawing
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            '' Zoom to the extents or limits of the drawing
            'acDoc.SendStringToExecute("._zoom _all ", True, False, False)
            '///////////////////////////////////////////////////////////////////
            Dim pStrOpts0 As PromptStringOptions = New PromptStringOptions(vbLf & _
                                                              "Enter the Text ∅ ")
            pStrOpts0.AllowSpaces = True
            pStrOpts0.Keywords.Add("@WC")
            Dim pStrRes0 As PromptResult = acDoc.Editor.GetString(pStrOpts0)
            'Dim textString As String = "Left"
            Dim textString As String = "%%C" & " " & pStrRes0.StringResult
            Dim textAlign As Integer = TextHorizontalMode.TextCenter
            Dim acText As DBText = New DBText()
            'acText.Height = 5
            '////////////
            Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & _
                                                               "Enter Font Size<" & FontSize & ">: ")
            Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)

            If pStrRes.StringResult = "" Then

            Else
                FontSize = CDbl(pStrRes.StringResult.ToString)
            End If
            'Dim s As New Form1
            's.Show()

            acText.Height = FontSize
            '////////////
            acText.TextString = textString
            acText.HorizontalMode = TextHorizontalMode.TextCenter
            Dim pp As Point3d = New Point3d(((ptEnd.X + ptStart.X) / 2), ((ptEnd.Y + ptStart.Y) / 2), ptEnd.Z)
            acText.Position = pp
            If acText.HorizontalMode <> TextHorizontalMode.TextLeft Then
                acText.AlignmentPoint = pp
                acText.Rotation = Math.Atan((ptEnd.Y - ptStart.Y) / (ptEnd.X - ptStart.X))
            End If

            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            ' '' Create a point over the alignment point of the text
            'Dim acPtAlign As Point3d = ptEnd
            'Dim acPtIns As Point3d = ptStart
            'Dim acPoint As DBPoint = New DBPoint(acPtAlign)
            'acPoint.ColorIndex = 1

            'acBlkTblRec.AppendEntity(acPoint)
            'acTrans.AddNewlyCreatedDBObject(acPoint, True)

            ' '' Adjust the insertion and alignment points
            'acPtIns = New Point3d(acPtIns.X, acPtIns.Y + 3, 0)
            'acPtAlign = acPtIns

            ' '' Set the point style to crosshair
            'Application.SetSystemVariable("PDMODE", 2)
            '///////////////////////////////////////////////////////////////////
            '' Commit the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub

    <CommandMethod("DiaOverLine")> _
    Public Sub DiaOverLine()
        '' Get the current database and start the Transaction Manager
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")


        '' Prompt for the start point
        pPtOpts.Message = vbLf & "Enter the start point of the line: "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptStart As Point3d = pPtRes.Value

        '' Exit if the user presses ESC or cancels the command
        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' Prompt for the end point
        pPtOpts.Message = vbLf & "Enter the end point of the line: "
        pPtOpts.UseBasePoint = True
        pPtOpts.BasePoint = ptStart
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptEnd As Point3d = pPtRes.Value

        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim acBlkTbl As BlockTable
            Dim acBlkTblRec As BlockTableRecord

            '' Open Model space for write
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, _
                                         OpenMode.ForRead)

            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                            OpenMode.ForWrite)

            '' Define the new line
            Dim acLine As Line = New Line(ptStart, ptEnd)
            Dim LineLength As Double = acLine.Length
            acLine.SetDatabaseDefaults()

            '' Add the line to the drawing
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)

            'Dim textString As String = "Left"
            Dim textString As String = Format(LineLength, "0.00").ToString
            Dim textAlign As Integer = TextHorizontalMode.TextCenter
            Dim acText As DBText = New DBText()
            '////////////
            Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & _
                                                               "Enter Font Size<" & FontSizeX & ">: ")
            Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)

            If pStrRes.StringResult = "" Then

            Else
                FontSizeX = CDbl(pStrRes.StringResult.ToString)
            End If
            'Dim s As New Form1
            's.Show()

            acText.Height = FontSizeX
            '////////////
            acText.TextString = textString
            acText.HorizontalMode = TextHorizontalMode.TextCenter
            Dim pp As Point3d = New Point3d(((ptEnd.X + ptStart.X) / 2), ((ptEnd.Y + ptStart.Y) / 2), ptEnd.Z)
            acText.Position = pp
            If acText.HorizontalMode <> TextHorizontalMode.TextLeft Then
                acText.AlignmentPoint = pp
                acText.Rotation = Math.Atan((ptEnd.Y - ptStart.Y) / (ptEnd.X - ptStart.X))
            End If

            acBlkTblRec.AppendEntity(acText)
            acTrans.AddNewlyCreatedDBObject(acText, True)
            ' '' Create a point over the alignment point of the text
            'Dim acPtAlign As Point3d = ptEnd
            'Dim acPtIns As Point3d = ptStart
            'Dim acPoint As DBPoint = New DBPoint(acPtAlign)
            'acPoint.ColorIndex = 1

            'acBlkTblRec.AppendEntity(acPoint)
            'acTrans.AddNewlyCreatedDBObject(acPoint, True)

            ' '' Adjust the insertion and alignment points
            'acPtIns = New Point3d(acPtIns.X, acPtIns.Y + 3, 0)
            'acPtAlign = acPtIns

            ' '' Set the point style to crosshair
            'Application.SetSystemVariable("PDMODE", 2)
            '///////////////////////////////////////////////////////////////////
            '' Commit the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
End Class
