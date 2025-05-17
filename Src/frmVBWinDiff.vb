
' Fichier frmVBWinDiff.vb : Interface d'options pour le comparateur WinDiff et WinMerge
' --------------------

' Conventions de nommage des variables :
' ------------------------------------
' b pour Boolean (booléen vrai ou faux)
' i pour Integer : % (en VB .Net, l'entier a la capacité du VB6.Long)
' l pour Long : &
' r pour nombre Réel (Single!, Double# ou Decimal : D)
' s pour String : $
' c pour Char ou Byte
' d pour Date
' a pour Array (tableau) : ()
' o pour Object : objet instancié localement
' refX pour reference à un objet X préexistant qui n'est pas sensé être fermé
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)
' frm pour Form
' cls pour Classe
' mod pour Module
' ...
' ------------------------------------

Imports System.Text ' Pour StringBuilder

Public Class frmVBWinDiff

#Region "Constantes"

    Private Const sNomExeWinDiff$ = "WinDiff.exe"
    Private Const sNom_Exe_WinDiff$ = "WinDiff._exe_"
    Private Const sExeVBWinDiff$ = "VBWinDiff.exe"
    Private Const sLienExeVBWinDiff$ = sExeVBWinDiff & ".lnk"

#End Region

#Region "Configuration"

    ' Taille de la page en octets
    Private Const iTaillePage% = 50000 ' Prod.

#End Region

#Region "Interface"

    Public m_sCheminFichier1$ = ""
    Public m_sCheminFichier2$ = ""

#End Region

#Region "Déclarations"

    Private m_iNbPages% = 1
    Private m_iNumPage% = 1
    Private m_sCheminWinDiff$ = ""
    Private m_sCheminWinMerge$ = ""
    Private m_sCheminTDTH$ = ""

#End Region

#Region "Initialisations"

    Private Sub frmVBWinDiff_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' modUtilFichier peut maintenant être compilé dans une dll
        DefinirTitreApplication(sTitreMsg)

        Dim sCheminWinDiff$ = Application.StartupPath & "\" & sNomExeWinDiff
        If bFichierExiste(sCheminWinDiff, bPrompt:=True) Then
            lbAlgo.Items.Add("WinDiff")
            m_sCheminWinDiff = sCheminWinDiff
        End If

        If bLireCleBRWinMerge() Then lbAlgo.Items.Add("WinMerge")

        Dim sCheminTDTH$ = Application.StartupPath & "\TextDiffToHtml\TextDiffToHtml.exe"
        If bFichierExiste(sCheminTDTH, bPrompt:=True) Then
            lbAlgo.Items.Add("TextDiffToHtml")
            m_sCheminTDTH = sCheminTDTH
        End If

    End Sub

    Private Sub frmVBWinDiff_Shown(sender As Object, e As EventArgs) _
        Handles Me.Shown

        Dim sVersion$ = " - V" & sVersionAppli & " (" & sDateVersionAppli & ")"
        Dim sDebug$ = " - Debug"
        Dim sTxt$ = Me.Text & sVersion
        If bDebug Then sTxt &= sDebug
        Me.Text = sTxt

        Dim bModeConfig As Boolean = False
        If Me.m_sCheminFichier1.Length = 0 Then
            bModeConfig = True
        Else
            If Not bFichierExiste(Me.m_sCheminFichier1,
                bPrompt:=True) Then bModeConfig = True
            If Not bFichierExiste(Me.m_sCheminFichier2,
                bPrompt:=True) Then bModeConfig = True
        End If

        If bModeConfig And bRelease Then
            Me.cmdAjouterRaccourci.Visible = True
            Me.cmdEnleverRaccourci.Visible = True
            Me.cmdComp.Visible = False
            Me.cmdAnnuler.Visible = False
            Me.lblChemin1.Text = ""
            Me.lblChemin2.Text = ""
            VerifierRaccourci()
            Exit Sub
        Else
            Me.cmdAjouterRaccourci.Visible = False
            Me.cmdEnleverRaccourci.Visible = False
        End If

        ' bDebug
        Dim sCheminFichier1$, sCheminFichier2$
        sCheminFichier1 = m_sCheminFichier1
        sCheminFichier2 = m_sCheminFichier2

        If bDebug Then
            'sCheminFichier1 = Application.StartupPath & "\Fichier1.txt"
            'sCheminFichier2 = Application.StartupPath & "\Fichier2.txt"
        End If

        Me.lbAlgo.Text = "WinMerge"

        ' 04/01/2014 On pagine le second fichier par rapport aux tronçons du 1er, de taille fixe
        '  mieux vaut tjrs inverser
        If Not String.IsNullOrEmpty(sCheminFichier1) AndAlso
           Not String.IsNullOrEmpty(sCheminFichier2) AndAlso
           bFichierExiste(sCheminFichier1) AndAlso
           bFichierExiste(sCheminFichier2) Then
            Dim lLong1& = (New IO.FileInfo(sCheminFichier1)).Length
            Dim lLong2& = (New IO.FileInfo(sCheminFichier2)).Length
            If lLong2 < lLong1 Then
                Dim sCheminTmp$ = sCheminFichier1
                sCheminFichier1 = sCheminFichier2
                sCheminFichier2 = sCheminTmp
                m_sCheminFichier1 = sCheminFichier1
                m_sCheminFichier2 = sCheminFichier2
            End If
            ' 04/01/2014
            ' 09/05/2014 Paginer ssi Windiff (WinMerge : pas besoin)
            If Me.lbAlgo.Text = "WinDiff" AndAlso
                (lLong1 > iTaillePage OrElse lLong2 > iTaillePage) Then Me.chkPaginer.Checked = True
        End If

        If bDebug Then
            Me.chkTout.Checked = False
            Me.chkAccents.Checked = False
            Me.chkPonctuation.Checked = False
            Me.chkCasse.Checked = False
            Me.chkEspacesInsec.Checked = False
            Me.chkEspaces.Checked = False
            Me.chkQuotes.Checked = False
            Me.chkInfo.Checked = True
            Me.chkPhrases.Checked = False
            Me.chkPaginer.Checked = False
            Me.chkRatio.Checked = False
            Me.chkParag.Checked = False
            Me.chkNum.Checked = False
            Me.lbAlgo.Text = "WinMerge"
            'Me.lbAlgo.Text = "TextDiffToHtml"
        End If

        Me.lblChemin1.Text = sCheminFichier1
        Me.lblChemin2.Text = sCheminFichier2

    End Sub

    Private Sub AfficherMessage(sMsg$)
        Me.sbStatusBar.Text = sMsg
        Application.DoEvents()
    End Sub

#End Region

#Region "Conversion"

    Private Function bConfirmerTailleFichier(sCheminFichier$) As Boolean

        Dim lTaille& = (New IO.FileInfo(sCheminFichier)).Length
        ' Afficher un avertissement ssi la taille risque vraiment de faire planter WinDiff
        If lTaille <= iTaillePage * 5 Then Return True
        If MsgBoxResult.Cancel = MsgBox("La taille du fichier (" &
            sFormaterTailleOctets(lTaille) & ") dépasse la taille limite conseillée (" &
            sFormaterTailleOctets(CLng(iTaillePage * 1.024), bSupprimerPt0:=True) &
            ") pour WinDiff :" & vbLf & sCheminFichier & vbLf &
            "Etes-vous sûr de vouloir comparer sans pagination ?",
            MsgBoxStyle.OkCancel Or MsgBoxStyle.Exclamation, sTitreMsg) Then Return False
        Return True

    End Function

    Private Sub Comparer()

        Sablier() ' 04/07/2022
        Me.cmdAnnuler.Enabled = True
        Me.cmdComp.Enabled = False

        Dim sCheminFichier1$ = Me.lblChemin1.Text
        Dim sCheminFichier2$ = Me.lblChemin2.Text
        Dim sCheminFichier1Orig$ = sCheminFichier1
        Dim sCheminFichier2Orig$ = sCheminFichier2

        Dim sChemin$ = Application.StartupPath
        If Not bSupprimerFichiersFiltres(sChemin, sFiltreTmp) Then GoTo Fin
        If Not bSupprimerFichiersFiltres(sChemin, sFiltreFusion) Then GoTo Fin

        If Not bFichierExiste(sCheminFichier1, bPrompt:=True) Then GoTo Fin
        If Not bFichierExiste(sCheminFichier2, bPrompt:=True) Then GoTo Fin

        Const bDebugSplit As Boolean = False

        Dim sbSrc1, sbSrc2 As StringBuilder
        Dim iIdxSrcOrig1%, iIdxSrcOrig2%

        ' 03/09/2022
        Dim sEncodage1$ = "", sEncodage2$ = ""
        ' Problème : parfois il faut laisser par défaut, parfois UTF8, comment choisir ?
        'Dim encod1 As Encoding = LireEncodage(sCheminFichier1, sEncodage1, bEncodageParDefautUTF8:=True)
        'Dim encod2 As Encoding = LireEncodage(sCheminFichier2, sEncodage2, bEncodageParDefautUTF8:=True)
        ' Solution : https://github.com/AutoItConsulting/text-encoding-detect
        Dim encod1 As Encoding = LireEncodageTED(sCheminFichier1, sEncodage1, bEncodageParDefaut:=True)
        Dim encod2 As Encoding = LireEncodageTED(sCheminFichier2, sEncodage2, bEncodageParDefaut:=True)

        ' 04/01/2014 Paginer ici une fois pour toutes
        If Me.chkPaginer.Checked Then

            Dim iNbPages% = Me.m_iNbPages
            Dim dico1Pages As Dictionary(Of Integer, clsPage) = Nothing
            Dim dico2Pages As Dictionary(Of Integer, clsPage) = Nothing
            If Not bPaginerFichiers(sCheminFichier1, sCheminFichier2,
                iTaillePage, iNbPages, dico1Pages, dico2Pages, Me.chkRatio.Checked,
                encod1, encod2) Then GoTo Fin
            Me.m_iNbPages = iNbPages
            sbSrc1 = dico1Pages(Me.m_iNumPage - 1).sbPage
            sbSrc2 = dico2Pages(Me.m_iNumPage - 1).sbPage
            iIdxSrcOrig1 = dico1Pages(Me.m_iNumPage - 1).iIndexSrc
            iIdxSrcOrig2 = dico2Pages(Me.m_iNumPage - 1).iIndexSrc

        Else

            If Me.lbAlgo.Text = "WinDiff" Then
                If Not bConfirmerTailleFichier(sCheminFichier1) Then GoTo Fin
                If Not bConfirmerTailleFichier(sCheminFichier2) Then GoTo Fin
            End If
            sbSrc1 = sbLireFichier(sCheminFichier1, encod1)
            sbSrc2 = sbLireFichier(sCheminFichier2, encod2)
            'If bDebugSplit Then Debug.WriteLine("Lecture :")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°1 : [" & sbSrc1.ToString & "]")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°2 : [" & sbSrc2.ToString & "]")
            iIdxSrcOrig1 = 0 : iIdxSrcOrig2 = 0

        End If

        Dim iNbEcritures% = 0
        If Not Me.chkEspacesInsec.Checked Then iNbEcritures += 1
        If Not Me.chkEspaces.Checked Then iNbEcritures += 1
        If Not Me.chkAccents.Checked Then iNbEcritures += 1
        If Not Me.chkCasse.Checked Then iNbEcritures += 1
        If Not Me.chkParag.Checked Then iNbEcritures += 1 ' 17/05/2025
        If Not Me.chkPonctuation.Checked Then iNbEcritures += 1
        If Not Me.chkQuotes.Checked Then iNbEcritures += 1
        Dim iNbEcrituresTot% = iNbEcritures

        Dim bEcriture As Boolean = False
        If Me.chkInfo.Checked OrElse iNbEcritures > 0 Then bEcriture = True

        Dim sbDest1 As StringBuilder = Nothing
        Dim sbDest2 As StringBuilder = Nothing

        Dim sbSrcOrig1 As New StringBuilder
        sbSrcOrig1.Append(sbSrc1)
        Dim sbSrcOrig2 As New StringBuilder
        sbSrcOrig2.Append(sbSrc2)

        If Not Me.chkQuotes.Checked Then
            NormaliserQuotes(sbSrc1, sbDest1) : sbSrc1 = sbDest1
            NormaliserQuotes(sbSrc2, sbDest2) : sbSrc2 = sbDest2
            iNbEcritures -= 1 ' Décrémenter le nombre d'écriture restantes
            'If bDebugSplit Then Debug.WriteLine("Quotes :")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°1 : [" & sbSrc1.ToString & "]")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°2 : [" & sbSrc2.ToString & "]")
        End If

        If Not Me.chkAccents.Checked Then
            ' 26/11/2021 Ne pas enlever la casse ici si on ne le demande pas
            Dim bMinuscule As Boolean = Not Me.chkCasse.Checked
            If Not bEnleverAccents(sCheminFichier1, sbSrc1, sbDest1, bMinuscule) Then GoTo Fin
            If Not bEnleverAccents(sCheminFichier2, sbSrc2, sbDest2, bMinuscule) Then GoTo Fin
            sbSrc1 = sbDest1 : sbSrc2 = sbDest2
            iNbEcritures -= 1 ' Décrémenter le nombre d'écriture restantes
            If bDebugSplit Then Debug.WriteLine("Accents :")
            If bDebugSplit Then Debug.WriteLine("Fichier n°1 : [" & sbSrc1.ToString & "]")
            If bDebugSplit Then Debug.WriteLine("Fichier n°2 : [" & sbSrc2.ToString & "]")
        End If

        If Not Me.chkEspacesInsec.Checked Then
            EnleverEspInsec(sCheminFichier1, sbSrc1, sbDest1) : sbSrc1 = sbDest1
            EnleverEspInsec(sCheminFichier2, sbSrc2, sbDest2) : sbSrc2 = sbDest2
            iNbEcritures -= 1 ' Décrémenter le nombre d'écriture restantes
            'If bDebugSplit Then Debug.WriteLine("EspacesInsec :")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°1 : [" & sbSrc1.ToString & "]")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°2 : [" & sbSrc2.ToString & "]")
        End If

        If Not Me.chkEspaces.Checked Then ' 03/05/2014
            EnleverEspaces(sCheminFichier1, sbSrc1, sbDest1) : sbSrc1 = sbDest1
            EnleverEspaces(sCheminFichier2, sbSrc2, sbDest2) : sbSrc2 = sbDest2
            iNbEcritures -= 1 ' Décrémenter le nombre d'écriture restantes
            'If bDebugSplit Then Debug.WriteLine("Espaces :")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°1 : [" & sbSrc1.ToString & "]")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°2 : [" & sbSrc2.ToString & "]")
        End If

        If Not Me.chkCasse.Checked Then
            If Not bEnleverMajuscules(sCheminFichier1, sbSrc1, sbDest1) Then GoTo Fin
            If Not bEnleverMajuscules(sCheminFichier2, sbSrc2, sbDest2) Then GoTo Fin
            sbSrc1 = sbDest1 : sbSrc2 = sbDest2
            iNbEcritures -= 1 ' Décrémenter le nombre d'écriture restantes
            'If bDebugSplit Then Debug.WriteLine("Casse :")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°1 : [" & sbSrc1.ToString & "]")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°2 : [" & sbSrc2.ToString & "]")
        End If

        ' 17/05/2025
        If Not Me.chkParag.Checked AndAlso Me.chkPonctuation.Checked Then
            If Not bDecouperParagraphesEnPhrasesAvecPonctuation(sCheminFichier1, sbSrc1, sbDest1) Then GoTo Fin
            If Not bDecouperParagraphesEnPhrasesAvecPonctuation(sCheminFichier2, sbSrc2, sbDest2) Then GoTo Fin
            sbSrc1 = sbDest1 : sbSrc2 = sbDest2
            iNbEcritures -= 1 ' Décrémenter le nombre d'écriture restantes
        End If

        If Not Me.chkPonctuation.Checked Then
            ' Option Mots possible ssi on ignore la ponctuation
            Dim bOptionComparerMots As Boolean = Not Me.chkPhrases.Checked
            Dim bOptionComparerParag As Boolean = Me.chkParag.Checked
            Dim bOptionComparerNum As Boolean = Me.chkNum.Checked
            Dim sCheminDest1$ = "", sCheminDest2$ = ""
            If Not bEnleverPonctuation(sCheminFichier1, sbSrc1, sbDest1,
                bOptionComparerMots, bOptionComparerParag, bOptionComparerNum) Then GoTo Fin
            If Not bEnleverPonctuation(sCheminFichier2, sbSrc2, sbDest2,
                bOptionComparerMots, bOptionComparerParag, bOptionComparerNum) Then GoTo Fin
            sbSrc1 = sbDest1 : sbSrc2 = sbDest2
            iNbEcritures -= 1 ' Décrémenter le nombre d'écriture restantes
            'If bDebugSplit Then Debug.WriteLine("Ponctuation :")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°1 : [" & sbSrc1.ToString & "]")
            'If bDebugSplit Then Debug.WriteLine("Fichier n°2 : [" & sbSrc2.ToString & "]")
        End If

        If bEcriture Then
            sCheminFichier1 = Application.StartupPath & "\" & sFichier & "1" & sExtTxt
            sCheminFichier2 = Application.StartupPath & "\" & sFichier & "2" & sExtTxt
            ' Si on a demandé les infos avec tout coché, il faut lire une 1ère fois
            If iNbEcrituresTot = 0 Then
                'sbDest1 = sbLireFichier(sCheminFichier1Orig)
                'sbDest2 = sbLireFichier(sCheminFichier2Orig)
                ' 10/07/2022
                sbDest1 = sbLireFichier(sCheminFichier1Orig, encod1)
                sbDest2 = sbLireFichier(sCheminFichier2Orig, encod2)
            End If
            If Not bEcrireFichiers(sbDest1, sbDest2,
                sCheminFichier1, sCheminFichier1Orig,
                sCheminFichier2, sCheminFichier2Orig, iIdxSrcOrig1, iIdxSrcOrig2,
                sbSrcOrig1, sbSrcOrig2, sEncodage1, sEncodage2) Then GoTo Fin
        End If



        Const sGm$ = """"
        Dim sCmd$ = sGm & sCheminFichier1 & sGm & " " & sGm & sCheminFichier2 & sGm
        Dim p As New Process

        Select Case Me.lbAlgo.Text ' 17/05/2025
            Case "WinDiff"
                p.StartInfo = New ProcessStartInfo(m_sCheminWinDiff)
            Case "WinMerge"
                p.StartInfo = New ProcessStartInfo(m_sCheminWinMerge)
            Case "TextDiffToHtml"
                p.StartInfo = New ProcessStartInfo(m_sCheminTDTH)
        End Select

        p.StartInfo.Arguments = sCmd
        ' Il faut indiquer le chemin de l'exe si on n'utilise pas le shell
        'p.StartInfo.UseShellExecute = False
        'If bMax Then p.StartInfo.WindowStyle = ProcessWindowStyle.Maximized
        p.Start()

Fin:
        ActivationCmdPage()
        Me.cmdComp.Enabled = True
        Me.cmdAnnuler.Enabled = False
        Sablier(bDesactiver:=True) ' 04/07/2022

    End Sub

    Private Function bEcrireFichiers(
        ByRef sbPage1 As StringBuilder, ByRef sbPage2 As StringBuilder,
        sCheminSrc1$, sCheminSrcOrig1$,
        sCheminSrc2$, sCheminSrcOrig2$, iIdxSrcOrig1%, iIdxSrcOrig2%,
        sbSrc1Orig As StringBuilder, sbSrc2Orig As StringBuilder,
        sEncodage1$, sEncodage2$) As Boolean

        For iNumFichier As Integer = 1 To 2
            Dim sCheminDest$ = Application.StartupPath & "\" & sFichier & iNumFichier & sExtTxt
            Dim sbPage As StringBuilder = sbPage1
            Dim sbSrcOrig As StringBuilder = sbSrc1Orig
            Dim sCheminSrc$ = sCheminSrc1
            Dim sSrcOrig$ = sCheminSrcOrig1
            Dim iIdxSrcOrig% = iIdxSrcOrig1
            Dim sEncodage$ = sEncodage1
            If iNumFichier = 2 Then
                sbPage = sbPage2
                sCheminSrc = sCheminSrc2
                sSrcOrig = sCheminSrcOrig2
                iIdxSrcOrig = iIdxSrcOrig2
                sbSrcOrig = sbSrc2Orig
                sEncodage = sEncodage2
            End If

            If Me.chkInfo.Checked Then
                Const bAfficherFichier As Boolean = True
                Dim bAfficherTailleFinale As Boolean = Not Me.chkParag.Checked
                Dim iNumPage% = m_iNumPage
                Dim iNbPages% = m_iNbPages
                If Not Me.chkPaginer.Checked Then iNumPage = 1 : iNbPages = 1 ' 12/01/2014
                Dim sbDest As StringBuilder = Nothing
                If Not bAjouterInfo(iNumFichier, sCheminSrc, sSrcOrig, sEncodage,
                    sbPage, sbSrcOrig, sbDest,
                    iNumPage, iNbPages, iIdxSrcOrig,
                    bAfficherFichier, bAfficherTailleFinale) Then Return False
                sbPage = sbDest
            End If

            ' 26/01/2014
            Dim bFusionMC As Boolean = Not Me.chkParag.Checked
            If bFusionMC Then
                Dim sbDest As StringBuilder = Nothing
                FusionnerMotsCoupes(sbPage, sbDest, iNumFichier, sSrcOrig)
                sbPage = sbDest
            End If

            If Not bEcrireFichier(sCheminDest, sbPage) Then Return False

            If bDebug AndAlso Me.chkPaginer.Checked Then
                Dim sDest2$ = Application.StartupPath &
                    "\" & sFichier & iNumFichier & "_" & sPage & Me.m_iNumPage & sExtTxt
                If Not bEcrireFichier(sDest2, sbPage) Then Return False
            End If

        Next

        Return True

    End Function

#End Region

#Region "Gestion des événements"

    Private Sub cmdComp_Click(sender As Object, e As EventArgs) _
        Handles cmdComp.Click

        Comparer()

    End Sub

    Private Sub chkTout_Click(sender As Object, e As EventArgs) _
        Handles chkTout.Click

        ' Tout implique les espaces inséc., la casse, les accents et la ponctuation,
        '  et vice versa : GererChkTout
        Me.chkEspacesInsec.Checked = Me.chkTout.Checked
        Me.chkEspaces.Checked = Me.chkTout.Checked
        Me.chkCasse.Checked = Me.chkTout.Checked
        Me.chkAccents.Checked = Me.chkTout.Checked
        Me.chkPonctuation.Checked = Me.chkTout.Checked
        Me.chkQuotes.Checked = Me.chkTout.Checked
        Me.chkNum.Checked = Me.chkTout.Checked
        Me.chkPhrases.Checked = Me.chkTout.Checked
        Me.chkParag.Checked = Me.chkTout.Checked
        GererActivationPhrasesEtParag()

    End Sub

    Private Sub GererChkTout()

        If Me.chkEspacesInsec.Checked AndAlso Me.chkEspaces.Checked AndAlso
           Me.chkCasse.Checked AndAlso
           Me.chkAccents.Checked AndAlso Me.chkPonctuation.Checked AndAlso
           Me.chkQuotes.Checked AndAlso Me.chkNum.Checked Then
            Me.chkTout.Checked = True
        Else
            Me.chkTout.Checked = False
        End If

    End Sub

    Private Sub chkEspacesInsec_Click(sender As Object, e As EventArgs) _
        Handles chkEspacesInsec.Click
        ' La détection des espaces insécables ne fonctionne que si l'on conserve la ponctuation
        If Me.chkEspacesInsec.Checked Then Me.chkPonctuation.Checked = True
        GererChkTout()
    End Sub

    Private Sub chkEspaces_Click(sender As Object, e As EventArgs) _
        Handles chkEspaces.Click
        ' La détection des espaces ne fonctionne que si l'on conserve la ponctuation
        If Me.chkEspaces.Checked Then Me.chkPonctuation.Checked = True
        GererChkTout()
    End Sub

    Private Sub chkCasse_Click(sender As Object, e As EventArgs) Handles chkCasse.Click
        GererChkTout()
    End Sub

    Private Sub chkAccents_Click(sender As Object, e As EventArgs) Handles chkAccents.Click
        GererChkTout()
    End Sub

    Private Sub chkQuotes_Click(sender As Object, e As System.EventArgs) Handles chkQuotes.Click
        GererChkTout()
    End Sub

    Private Sub chkNum_Click(sender As Object, e As EventArgs) Handles chkNum.Click
        ' La possibilité d'ignorer les numériques ne fonctionne que si
        '  l'on retire la ponctuation et que l'on compare mot à mot
        If Not Me.chkNum.Checked Then _
            Me.chkPonctuation.Checked = False : Me.chkPhrases.Checked = False
        GererChkTout()
    End Sub

    Private Sub chkPonctuation_Click(sender As Object, e As EventArgs) _
        Handles chkPonctuation.Click
        GererChkTout()
    End Sub

    Private Sub chkPonctuation_CheckedChanged(sender As Object, e As EventArgs) _
        Handles chkPonctuation.CheckedChanged
        GererActivationPhrasesEtParag()
    End Sub

    Private Sub GererActivationPhrasesEtParag()

        ' L'option de découpage des phrases en mots n'est possible que si on ignore la ponctuation
        '  donc si on coche la ponctuation, on doit désactiver le découpage en mots (chkPhrases = True)
        If Me.chkPonctuation.Checked AndAlso Not Me.chkPhrases.Checked Then _
            Me.chkPhrases.Checked = True

        ' 17/05/2025 Commenté
        ' Pareil pour les paragraphes
        'If Me.chkPonctuation.Checked AndAlso Not Me.chkParag.Checked Then _
        '    Me.chkParag.Checked = True
        ' Pareil pour les numériques 03/06/2018

        If Me.chkPonctuation.Checked AndAlso Not Me.chkNum.Checked Then _
            Me.chkNum.Checked = True

    End Sub

    Private Sub chkPhrases_Click(sender As Object, e As EventArgs) _
        Handles chkPhrases.Click

        ' Si on compare mot à mot, alors décocher la ponctuation
        ' (car le mode mot à mot est lancé uniquement dans ce cas)
        If Not Me.chkPhrases.Checked AndAlso Me.chkPonctuation.Checked Then _
            Me.chkPonctuation.Checked = False ' 04/01/2014 Sens unique

        ' 03/06/2018 Si on coche les phrases, on ne peut pas ignorer les numériques
        If Me.chkPhrases.Checked Then Me.chkNum.Checked = True

    End Sub

    Private Sub chkPhrases_CheckedChanged(sender As Object, e As EventArgs) _
        Handles chkPhrases.CheckedChanged
        ActivationCmdPage()
    End Sub

    Private Sub chkParag_Click(sender As Object, e As EventArgs) Handles chkParag.Click

        ' 17/05/2025 Commenté
        ' Si on ignore les paragraphes dans le mode mot à mot, alors décocher la ponctuation
        ' (car le mode mot à mot est lancé uniquement dans ce cas)
        'If Not Me.chkParag.Checked AndAlso Me.chkPonctuation.Checked Then _
        '    Me.chkPonctuation.Checked = False : Me.chkPhrases.Checked = False

    End Sub

    'Private Sub chkParag_CheckedChanged(sender As Object, e As EventArgs) Handles chkParag.CheckedChanged
    'End Sub

    Private Sub chkPaginer_CheckedChanged(sender As Object, e As EventArgs) _
        Handles chkPaginer.CheckedChanged

        ActivationCmdPage()

    End Sub

    Private Sub cmdPagePreced_Click(sender As Object, e As EventArgs) _
        Handles cmdPagePreced.Click

        Me.m_iNumPage -= 1
        ActivationCmdPage()
        If Me.m_iNumPage = 1 Then Me.cmdPageSuiv.Select()

    End Sub

    Private Sub cmdPageSuiv_Click(sender As Object, e As EventArgs) _
        Handles cmdPageSuiv.Click

        Me.m_iNumPage += 1
        ActivationCmdPage()
        If Me.m_iNumPage = Me.m_iNbPages Then Me.cmdPagePreced.Select()

    End Sub

    Private Sub ActivationCmdPage()

        If Me.chkPaginer.Checked Then
            Me.cmdPagePreced.Enabled = (Me.m_iNumPage > 1)
            Me.cmdPageSuiv.Enabled = (Me.m_iNumPage < Me.m_iNbPages)
            Me.lblNumPage.Enabled = True
            Me.chkRatio.Enabled = True
        Else
            Me.lblNumPage.Enabled = False
            Me.cmdPagePreced.Enabled = False
            Me.cmdPageSuiv.Enabled = False
            Me.chkRatio.Enabled = False
        End If
        Me.lblNumPage.Text = Me.m_iNumPage & "/" & Me.m_iNbPages

    End Sub

    Private Function bLireCleBRWinMerge() As Boolean

        Dim sCheminWinMerge$ = ""
        If Not bCleRegistreCUExiste("SOFTWARE\Thingamahoochie\WinMerge",
            "Executable", sCheminWinMerge) Then Return False
        ' Par défaut : "C:\Program Files\WinMerge\WinMergeU.exe"
        m_sCheminWinMerge = sCheminWinMerge
        If m_sCheminWinMerge.Length = 0 Then Return False
        If Not bFichierExiste(m_sCheminWinMerge, bPrompt:=True) Then Return False
        Return True

    End Function

#End Region

#Region "Gestion du raccourci"

    Private m_sCheminRaccourci$ =
        Environment.GetFolderPath(Environment.SpecialFolder.SendTo) & "\" & sLienExeVBWinDiff

    Private Sub VerifierRaccourci()

        If bFichierExiste(m_sCheminRaccourci) Then
            Me.cmdAjouterRaccourci.Enabled = False
            Me.cmdEnleverRaccourci.Enabled = True
        Else
            Me.cmdAjouterRaccourci.Enabled = True
            Me.cmdEnleverRaccourci.Enabled = False
        End If

    End Sub

    Private Sub cmdAjouterRaccourci_Click(sender As Object, e As EventArgs) _
        Handles cmdAjouterRaccourci.Click

        Dim sLien$ = m_sCheminRaccourci
        Dim sCibleFinale$ = Application.StartupPath & "\" & sExeVBWinDiff
        CreerRaccourci(sLien, sCibleFinale)
        VerifierRaccourci()

    End Sub

    Private Sub cmdEnleverRaccourci_Click(sender As Object, e As EventArgs) _
         Handles cmdEnleverRaccourci.Click

        If Not bFichierExiste(m_sCheminRaccourci) Then Exit Sub
        If Not bSupprimerFichier(m_sCheminRaccourci) Then Exit Sub
        VerifierRaccourci()

    End Sub

#End Region

End Class