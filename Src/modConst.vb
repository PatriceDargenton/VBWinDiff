
Module modConst

#If DEBUG Then
    Public Const bDebug As Boolean = True
    Public Const bRelease As Boolean = False
#Else
    Public Const bDebug As Boolean = False
    Public Const bRelease As Boolean = True
#End If

    Public Const sExtTxt$ = ".txt"
    Public Const sFusion$ = "Fusion"
    Public Const sFichier$ = "Fichier"
    Public Const sPage$ = "Page"
    Public Const sOrig$ = "Orig"
    'Public Const sFiltrePages$ = "Fichier?_Page?.txt"
    Public Const sFiltreTmp$ = "Fichier?_*" & sExtTxt
    Public Const sFiltreFusion$ = sFusion & "?" & sExtTxt

    Public Const sListeSeparateursPhrase$ = ".:?!;|¡¿"

    ' Normalisation des quotes
    Public Const iCodeASCIIGuillemet% = 34 ' "
    Public Const iCodeASCIIGuillemetOuvrant% = 171 ' « ' Rétabli le 18/11/2018
    Public Const iCodeASCIIGuillemetFermant% = 187 ' » ' Rétabli le 18/11/2018
    Public Const iCodeASCIIGuillemetOuvrant3% = 147 ' “
    Public Const iCodeASCIIGuillemetFermant3% = 148 ' ”
    Public Const iCodeASCIIQuote% = 39 '
    Public Const iCodeASCIIQuote2% = 27 '
    Public Const iCodeASCIIGuillemetOuvrant2% = 145 ' ‘
    Public Const iCodeASCIIGuillemetFermant2% = 146 ' ’
    'Public Const iCodeASCIIGuillemetOuvrant4% = 96 ' `
    Public Const iCodeASCIIGuillemetFermant4% = 180 ' ´

    Public Const iCodeASCIIEspaceInsecable% = 160 ' Non-breaking space &nbsp;
    Public Const iCodeUTF16EspaceFineInsecable% = 8239 ' Alt+8239 = 0x202F = espace fine insécable

    Public Const iCodeASCIITiretMoyen% = 150 ' –

    Const cChar3P As Char = "…"c ' 17/05/2025
    Public Const sChar3P$ = cChar3P

    Public Const iIndiceNulString% = -1

End Module