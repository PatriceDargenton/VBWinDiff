
' Fichier modUtil.vb
' ------------------

Module modUtil

    Public Sub AfficherMsgErreur2(ByRef Ex As Exception,
        Optional sTitreFct$ = "", Optional sInfo$ = "",
        Optional sDetailMsgErr$ = "",
        Optional bCopierMsgPressePapier As Boolean = True,
        Optional ByRef sMsgErrFinal$ = "")

        If Not Cursor.Current.Equals(Cursors.Default) Then _
            Cursor.Current = Cursors.Default
        Dim sMsg$ = ""
        If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
        If sInfo <> "" Then sMsg &= vbCrLf & sInfo
        If sDetailMsgErr <> "" Then sMsg &= vbCrLf & sDetailMsgErr
        If Ex.Message <> "" Then
            sMsg &= vbCrLf & Ex.Message.Trim
            If Not IsNothing(Ex.InnerException) Then _
                sMsg &= vbCrLf & Ex.InnerException.Message
        End If
        If bCopierMsgPressePapier Then CopierPressePapier(sMsg)
        sMsgErrFinal = sMsg
        MsgBox(sMsg, MsgBoxStyle.Critical)

    End Sub

    Public Sub Sablier(Optional ByRef bDesactiver As Boolean = False)

        If bDesactiver Then
            Cursor.Current = Cursors.Default
        Else
            Cursor.Current = Cursors.WaitCursor
        End If

    End Sub

    Public Sub CopierPressePapier(sInfo$)

        ' Copier des informations dans le presse-papier de Windows
        ' (elles resteront jusqu'à ce que l'application soit fermée)

        Try
            Dim dataObj As New DataObject
            dataObj.SetData(DataFormats.Text, sInfo)
            Clipboard.SetDataObject(dataObj)
        Catch ex As Exception
            ' Le presse-papier peut être indisponible
            AfficherMsgErreur2(ex, "CopierPressePapier",
                bCopierMsgPressePapier:=False)
        End Try

    End Sub

    Public Sub TraiterMsgSysteme_DoEvents()

        Try
            Application.DoEvents() ' Peut planter avec OWC : Try Catch nécessaire
        Catch
        End Try

    End Sub

    Public Sub LibererRessourceDotNet()

        GC.Collect()
        GC.WaitForPendingFinalizers()
        TraiterMsgSysteme_DoEvents()

    End Sub

End Module