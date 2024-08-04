
Option Strict Off ' Pour CreateObject("WScript.Shell")

Module modRaccourci

    Sub CreerRaccourci(ByRef sCheminRaccourci$, ByRef sCheminCible$,
        Optional ByRef sRepertoirDeTravail$ = "",
        Optional ByRef iStyleWindows% = 4,
        Optional ByRef sCheminIcone$ = "",
        Optional ByRef iIndexIcone% = 0,
        Optional ByRef sArguments$ = "")

        ' Fonction pour cr�er un raccourci

        ' Param�tres :
        ' ----------
        ' sCheminRaccourci : Chemin du raccourci 
        ' (ex.: C:\Documents and Settings\[MonCompteUtilisateur]\SendTo\VBWinDiff.exe.lnk)
        ' sCheminCible : La cible du raccourci (ex.: C:\Tmp\VBWinDiff.exe)

        ' Param�tres Facultatifs :
        ' ----------------------
        ' sRepertoirDeTravail : R�pertoire d'ex�cution, par defaut le r�pertoire 
        '  contenant l'ex�cutable (ex.: C:\Tmp)
        ' iStyleWindows : Comment est affich� le programme : normal, reduit, agrandi... 
        '  Par defaut: normal (comme pour shell en VB, ex.: 4 = normal)
        ' sCheminIcone : Chemin d'acces de l'icone, par defaut l'icone de 
        '  l'ex�cutable cible (sinon aucun) (ex.: C:\Tmp\VBWinDiff.ico)
        ' iIndexIcone    : L'index de l'icone dans le fichier
        ' ----------------------

        ' Si il n'y a le .lnk � la fin on l'ajoute
        If Right(sCheminRaccourci, 4).ToLower <> ".lnk" Then sCheminRaccourci &= ".lnk"

        If sRepertoirDeTravail.Length = 0 Then _
            sRepertoirDeTravail = IO.Path.GetDirectoryName(sCheminCible)

        ' Si un n'y a pas d'icone, on prend l'icone de l'ex�cutable cible ou rien
        If sCheminIcone.Length = 0 Then sCheminIcone = sCheminCible

        Dim oWSHShell As Object ' Pour Cr�er le raccourci
        Dim oShortcut As Object ' Raccourci

        oWSHShell = CreateObject("WScript.Shell") ' on cr�e un objet Shell

        ' Cr�ation d'un objet raccourci
        oShortcut = oWSHShell.CreateShortcut(sCheminRaccourci)

        ' Param�trage du raccourci
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