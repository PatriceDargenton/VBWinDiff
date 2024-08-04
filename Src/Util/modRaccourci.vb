
Option Strict Off ' Pour CreateObject("WScript.Shell")

Module modRaccourci

    Sub CreerRaccourci(ByRef sCheminRaccourci$, ByRef sCheminCible$,
        Optional ByRef sRepertoirDeTravail$ = "",
        Optional ByRef iStyleWindows% = 4,
        Optional ByRef sCheminIcone$ = "",
        Optional ByRef iIndexIcone% = 0,
        Optional ByRef sArguments$ = "")

        ' Fonction pour créer un raccourci

        ' Paramètres :
        ' ----------
        ' sCheminRaccourci : Chemin du raccourci 
        ' (ex.: C:\Documents and Settings\[MonCompteUtilisateur]\SendTo\VBWinDiff.exe.lnk)
        ' sCheminCible : La cible du raccourci (ex.: C:\Tmp\VBWinDiff.exe)

        ' Paramètres Facultatifs :
        ' ----------------------
        ' sRepertoirDeTravail : Répertoire d'exécution, par defaut le répertoire 
        '  contenant l'exécutable (ex.: C:\Tmp)
        ' iStyleWindows : Comment est affiché le programme : normal, reduit, agrandi... 
        '  Par defaut: normal (comme pour shell en VB, ex.: 4 = normal)
        ' sCheminIcone : Chemin d'acces de l'icone, par defaut l'icone de 
        '  l'exécutable cible (sinon aucun) (ex.: C:\Tmp\VBWinDiff.ico)
        ' iIndexIcone    : L'index de l'icone dans le fichier
        ' ----------------------

        ' Si il n'y a le .lnk à la fin on l'ajoute
        If Right(sCheminRaccourci, 4).ToLower <> ".lnk" Then sCheminRaccourci &= ".lnk"

        If sRepertoirDeTravail.Length = 0 Then _
            sRepertoirDeTravail = IO.Path.GetDirectoryName(sCheminCible)

        ' Si un n'y a pas d'icone, on prend l'icone de l'exécutable cible ou rien
        If sCheminIcone.Length = 0 Then sCheminIcone = sCheminCible

        Dim oWSHShell As Object ' Pour Créer le raccourci
        Dim oShortcut As Object ' Raccourci

        oWSHShell = CreateObject("WScript.Shell") ' on crée un objet Shell

        ' Création d'un objet raccourci
        oShortcut = oWSHShell.CreateShortcut(sCheminRaccourci)

        ' Paramétrage du raccourci
        oShortcut.TargetPath = sCheminCible
        oShortcut.Arguments = sArguments
        oShortcut.WorkingDirectory = sRepertoirDeTravail
        oShortcut.WindowStyle = iStyleWindows
        ' ExpandEnvironmentStrings permet de traiter des variables de chemin
        '  telles que %windir% par exemple
        oShortcut.IconLocation =
            oWSHShell.ExpandEnvironmentStrings(sCheminIcone & ", " & iIndexIcone)

        oShortcut.Save()

        oShortcut = Nothing
        oWSHShell = Nothing

    End Sub

End Module