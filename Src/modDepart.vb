
' Fichier modDepart.vb
' --------------------

Module modDepart

    Public ReadOnly sNomAppli$ = My.Application.Info.Title
    Public ReadOnly sTitreMsg$ = sNomAppli
    Public m_sTitreMsg$ = sTitreMsg
    Public Const sTitreMsgDescription$ = " : Interface d'options pour le comparateur WinDiff et WinMerge"
    Public Const sDateVersionAppli$ = "17/05/2025"

    Public ReadOnly sVersionAppli$ =
        My.Application.Info.Version.Major & "." &
        My.Application.Info.Version.Minor &
        My.Application.Info.Version.Build

    Public Sub DefinirTitreApplication(sTitreMsg As String)
        m_sTitreMsg = sTitreMsg
    End Sub

    Public Sub Main()

        ' S'il n'y a aucune gestion d'erreur, on peut déboguer dans l'IDE
        ' Sinon, ce n'est pas pratique de retrouver la ligne du bug :
        '  il faut cocher Levé (Thrown) dans le menu Déboguer:Exceptions... pour les 2 lignes
        ' Dans ce cas, il peut y avoir beaucoup d'interruptions selon la logique 
        '   de programmation : mieux vaut prévenir les erreurs que de les traiter,
        '   sinon utiliser l'attribut de fonction <System.Diagnostics.DebuggerStepThrough()>
        If bDebug Then Depart() : Exit Sub

        ' Attention : En mode Release il faut un Try Catch ici  
        '  car sinon il n'y a pas de gestion d'erreur !
        ' (.Net renvoie un message d'erreur équivalent 
        '  à un plantage complet sans explication)
        Try
            Depart()
        Catch ex As Exception
            AfficherMsgErreur2(ex, "Depart " & sTitreMsg)
        End Try

    End Sub

    Public Sub Depart()

        ' On peut démarrer l'application sur la feuille, ou bien sur la procédure 
        '  Main() si on veut pouvoir détecter l'absence de la dll sans plantage

        ' Extraire les options passées en argument de la ligne de commande
        ' Cette fct ne marche pas avec des chemins contenant des espaces, même entre guillemets
        'Dim asArgs$() = Environment.GetCommandLineArgs()
        Dim sArg0$ = Microsoft.VisualBasic.Interaction.Command
        Dim sCheminFichier1$ = ""
        Dim sCheminFichier2$ = ""
        'Dim iTypeComp As frmVBWinDiff.TypeComp = frmVBWinDiff.TypeComp.xxx
        Dim bSyntaxeOk As Boolean = False
        Dim iNbArguments% = 0

        If sArg0 <> "" Then
            Dim asArgs$() = asArgLigneCmd(sArg0)
            iNbArguments = UBound(asArgs) + 1
            If iNbArguments = 2 Then bSyntaxeOk = True
            If Not bSyntaxeOk Then GoTo Suite
            sCheminFichier1 = asArgs(0)
            If Not bFichierExiste(sCheminFichier1, bPrompt:=True) Then _
                bSyntaxeOk = False : GoTo Suite
            sCheminFichier2 = asArgs(1)
            If Not bFichierExiste(sCheminFichier2, bPrompt:=True) Then _
                bSyntaxeOk = False
        End If
Suite:
        If bRelease And Not bSyntaxeOk Then
            MsgBox(
                "Syntaxe : Chemin des deux fichiers textes à comparer" & vbCrLf &
                "Sinon ajouter le raccourci via le menu dédié suivant" & vbCrLf &
                " et envoyer deux fichiers à comparer vers VBWinDiff" & vbCrLf &
                " via l'explorateur de fichier de Windows.",
                MsgBoxStyle.Information, sTitreMsg & sTitreMsgDescription)
            If iNbArguments > 0 Then Exit Sub
        End If

        Dim oFrm As New frmVBWinDiff
        oFrm.m_sCheminFichier1 = sCheminFichier1
        oFrm.m_sCheminFichier2 = sCheminFichier2
        'oFrm.m_iTypeConv = iTypeConv
        Application.Run(oFrm)

    End Sub

End Module