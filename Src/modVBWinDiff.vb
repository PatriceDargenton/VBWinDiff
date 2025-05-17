
'Imports System.Text.Encoding ' Pour GetEncoding
Imports System.Text ' Pour StringBuilder
Imports System.Text.RegularExpressions ' Pour Regex

Module modVBWinDiff

#Region "Classes page et mot"

    Public Class clsPage
        Public iIndexSrc% = 0 ' Index de départ dans le texte d'origine
        Public sbPage As StringBuilder
        Public Sub New(iIndexSrc0%, sbPage0 As StringBuilder)
            iIndexSrc = iIndexSrc0
            sbPage = sbPage0
        End Sub
    End Class

    Public Class clsMot
        Public sMotConcat$, sMot1$, sMot2$
        Public iNbOccConcat%, iNbOcc1%, iNbOcc2%
    End Class

#End Region

    Public Function bEnleverAccents(sSrc$,
        ByRef sbSrc As StringBuilder,
        ByRef sbDest As StringBuilder, bMinuscule As Boolean) As Boolean

        If IsNothing(sbSrc) Then sbSrc = sbLireFichier(sSrc)
        If IsNothing(sbDest) Then sbDest = New StringBuilder

        sbDest = sbEnleverAccents(sbSrc, bMinuscule)

        bEnleverAccents = True

    End Function

    Private Sub LireInfoFichier(sCheminFichier$,
        ByRef sTaille$, ByRef sDate$, ByRef lTaille&)

        Dim fi As New IO.FileInfo(sCheminFichier)
        lTaille = fi.Length
        Dim sTailleFichier$ = sFormaterTailleOctets(lTaille)
        Dim sTailleFichierDetail$ = sFormaterTailleOctets(lTaille, bDetail:=True)
        ' Attention à l'heure de la date : l'explorateur de Windows XP
        '  enlève 1 heure si l'on est passé à l'heure d'hiver depuis la date à afficher
        '  c'est n'importe quoi !
        ' Heureusement fi.LastWriteTime affiche toujours la bonne heure (et la même heure)
        sTaille = sTailleFichierDetail
        sDate = fi.LastWriteTime.ToString

    End Sub

    Public Function bAjouterInfo(iNumFichier%, sSrc$, sSrcOrig$, sEncodage$,
        ByRef sbSrc As StringBuilder,
        ByRef sbSrcOrig As StringBuilder,
        ByRef sbDest As StringBuilder,
        Optional iNumPage% = 0, Optional iNbPages% = 0,
        Optional iIdxSrcOrig% = -1,
        Optional bAfficherFichier As Boolean = True,
        Optional bAfficherTailleFinale As Boolean = False) As Boolean

        If IsNothing(sbSrc) Then
            If bDebug Then Stop
            Return False
        End If
        If IsNothing(sbSrcOrig) Then
            If bDebug Then Stop
            Return False
        End If

        If IsNothing(sbDest) Then sbDest = New StringBuilder

        Dim sTailleFichier$ = "", sDateFichier$ = ""
        Dim lTailleFichier& = 0
        LireInfoFichier(sSrcOrig, sTailleFichier, sDateFichier, lTailleFichier)

        Dim sbTmp As New StringBuilder
        ' Mettre le fichier en 1er, car le dossier sera tjrs le même
        If bAfficherFichier Then
            sbTmp.AppendLine("Fichier n°" & iNumFichier & " : " &
                IO.Path.GetFileName(sSrcOrig) & " : " &
                IO.Path.GetDirectoryName(sSrcOrig))
            sbTmp.AppendLine("Taille = " & sTailleFichier & ", Date = " & sDateFichier)
        End If

        If iNbPages > 1 Then
            sbTmp.AppendLine("Page " & iNumPage & "/" & iNbPages)

            ' Afficher la taille de la page en octets :
            '  Section = [octet de départ - octet de fin]
            Dim iTaillePage% = 0
            Dim lLong& = 0
            If iIdxSrcOrig > lTailleFichier Then
                If bDebug Then Stop
            End If
            lLong = iIdxSrcOrig + sbSrcOrig.Length
            If lLong > lTailleFichier Then Stop
            sbTmp.Append("Section = [" & iIdxSrcOrig & " - " & lLong & "[")
            iTaillePage = CInt(lLong - iIdxSrcOrig)
            If sbSrcOrig.Length > iTaillePage Then
                If bDebug Then Stop
            End If
            sbTmp.Append(" : " & sFormaterTailleOctets(iTaillePage, bDetail:=True))

            ' Total déjà découpé :
            sbTmp.Append(" : " & sFormaterTailleOctets(lLong, bDetail:=True))

            ' Pourcentage déjà découpé :
            Dim rPC! = CSng(lLong / lTailleFichier)
            sbTmp.Append(" : " & rPC.ToString("0.00%"))

            sbTmp.Append(vbCrLf)
        End If

        ' 26/01/2014 Si on décoche l'option Paragraphe, alors afficher la taille finale
        '  après les traitements, pour vérifier rapidement si les textes sont de même
        '  longueur (si on ne compare que les mots sans les sauts de ligne par ex.)
        If bAfficherTailleFinale Then
            Dim lTailleFinale& = sbSrc.Length
            Dim sTailleFinale$ = sFormaterTailleOctets(lTailleFinale, bDetail:=True)
            sbTmp.AppendLine("Taille finale = " & sTailleFinale)
        End If

        ' 10/07/2022
        sbTmp.AppendLine("Encodage = " & sEncodage)

        sbDest = sbTmp.Append(sbSrc)

        bAjouterInfo = True

    End Function

    Public Function bPaginerFichiers(sCheminFichier1$, sCheminFichier2$, iTaillePage%, ByRef iNbPages%,
        ByRef dico1Pages As Dictionary(Of Integer, clsPage),
        ByRef dico2Pages As Dictionary(Of Integer, clsPage),
        bAppliquerRatio As Boolean, encodage1 As Encoding, encodage2 As Encoding) As Boolean

        Dim sbSrc1 As StringBuilder = sbLireFichier(sCheminFichier1, encodage1)
        Dim sbSrc2 As StringBuilder = sbLireFichier(sCheminFichier2, encodage2)
        Dim iLongSrc1% = sbSrc1.Length
        Dim iLongSrc2% = sbSrc2.Length

        Dim iLongMax12% = 0
        Dim bLongMax2 As Boolean = False
        If iLongSrc1 > iLongMax12 Then iLongMax12 = iLongSrc1
        If iLongSrc2 >= iLongMax12 Then iLongMax12 = iLongSrc2 : bLongMax2 = True
        If Not bLongMax2 Then
            If bDebug Then Stop ' Ce n'est plus possible grâce à la permutation des 2 fichiers
        End If

        iNbPages = CInt(iLongMax12 \ CLng(iTaillePage))
        Dim lReste& = iLongMax12 Mod iTaillePage
        If lReste > 0 Then iNbPages += 1

        Dim rRatio! = 1.0!
        If bAppliquerRatio AndAlso iLongSrc1 > 0 Then
            rRatio = CSng(iLongSrc2 / iLongSrc1)
        End If

        dico1Pages = New Dictionary(Of Integer, clsPage)
        dico2Pages = New Dictionary(Of Integer, clsPage)
        Dim iCumulPage1% = 0
        Dim iCumulPage2% = 0

        Dim iNumPage%
        For iNumPage = 0 To iNbPages - 1

            ' Pagination
            Dim iLong1% = iTaillePage
            Dim iLong2% = iTaillePage
            If bAppliquerRatio Then
                iLong1 = CInt(iLong1 / rRatio)
                ' Gestion de l'arrondi
                If iNumPage = iNbPages - 1 AndAlso iCumulPage1 + iLong1 < iLongSrc1 Then
                    iLong1 = iLongSrc1 - iCumulPage1
                End If
            End If
            Dim iIdxSrc1% = iCumulPage1
            Dim iIdxSrc2% = iCumulPage2

            Dim sbDestPage1 As StringBuilder = Nothing
            Paginer(iIdxSrc1, iLong1, sbSrc1, sbDestPage1)
            dico1Pages.Add(iNumPage, New clsPage(iIdxSrc1, sbDestPage1))

            Dim sbDestPage2 As StringBuilder = Nothing
            Paginer(iIdxSrc2, iLong2, sbSrc2, sbDestPage2)
            dico2Pages.Add(iNumPage, New clsPage(iIdxSrc2, sbDestPage2))

            Dim iLongDest1% = sbDestPage1.Length
            Dim iLongDest2% = sbDestPage2.Length
            iCumulPage1 += iLongDest1
            iCumulPage2 += iLongDest2
            'Debug.WriteLine("Page n°" & iNumPage + 1 & " :")
            'Debug.WriteLine("Fichier 1 : " & iLongDest1 & " : " & iCumulPage1 & "/" & iLongSrc1)
            'Debug.WriteLine("Fichier 2 : " & iLongDest2 & " : " & iCumulPage2 & "/" & iLongSrc2)

        Next

        Return True

    End Function

    Private Sub Paginer(iIdxSrc%, iTailleTroncon%, sbSrc As StringBuilder,
        ByRef sbDestPage As StringBuilder)

        Dim iLongSb% = sbSrc.Length
        Dim iMemTailleTroncon% = iTailleTroncon
        If iIdxSrc + iMemTailleTroncon > iLongSb Then
            iTailleTroncon = iLongSb - iIdxSrc
            If iTailleTroncon < 0 Then
                iTailleTroncon = 0
                If iIdxSrc > iLongSb Then
                    ' De toute façon, le fichier sera vide ici (iLong = 0)
                    '  c'est juste pour éviter un dépassement
                    iIdxSrc = iLongSb
                End If
            End If
        End If

        Dim ac As Char() = Nothing
        ReDim ac(0 To iTailleTroncon - 1)
        sbSrc.CopyTo(iIdxSrc, ac, 0, iTailleTroncon)
        sbDestPage = New StringBuilder
        For Each cCar As Char In ac
            sbDestPage.Append(cCar)
        Next

    End Sub

    Public Sub EnleverEspInsec(sSrc$,
        ByRef sbSrc As StringBuilder,
        ByRef sbDest As StringBuilder)

        If IsNothing(sbSrc) Then sbSrc = sbLireFichier(sSrc)
        sbDest = sbSrc.Replace(Chr(iCodeASCIIEspaceInsecable), " "c)
        sbDest = sbSrc.Replace(ChrW(iCodeUTF16EspaceFineInsecable), " "c) ' 15/09/2018

        ' 05/07/2024 Remplacer les tirets moyens (–) par des tirets courts (-)
        sbDest = sbSrc.Replace(" " & Chr(iCodeASCIITiretMoyen) & " ", " - ")
        sbDest = sbSrc.Replace(Chr(iCodeASCIITiretMoyen) & " ", "- ")
        sbDest = sbSrc.Replace(" " & Chr(iCodeASCIITiretMoyen), " -")

        sbDest = sbSrc.Replace(sChar3P, "...") ' 17/05/2025

    End Sub

    Public Sub EnleverEspaces(sSrc$,
        ByRef sbSrc As StringBuilder, ByRef sbDest As StringBuilder)

        If IsNothing(sbSrc) Then sbSrc = sbLireFichier(sSrc)
        sbDest = New StringBuilder

        ' Découpage par paragraphe
        Dim asParag$() = sbSrc.ToString.Split(CChar(vbCrLf))
        For Each sParag As String In asParag

            Dim sParagTrim$ = sParag.Trim

            ' 12/12/2015 Supprimer les doubles saut de ligne si on coche Espace
            If String.IsNullOrEmpty(sParagTrim) Then Continue For

            sbDest.AppendLine(sParagTrim)
        Next

    End Sub

    Public Function bEnleverMajuscules(sSrc$,
        ByRef sbSrc As StringBuilder,
        ByRef sbDest As StringBuilder) As Boolean

        If IsNothing(sbSrc) Then sbSrc = sbLireFichier(sSrc)

        sbDest = New StringBuilder
        sbDest.Append(sbSrc.ToString.ToLower)

        bEnleverMajuscules = True

    End Function

    ' 17/05/2025
    Public Function bDecouperParagraphesEnPhrasesAvecPonctuation(sSrc$,
        ByRef sbSrc As StringBuilder,
        ByRef sbDest As StringBuilder) As Boolean

        If IsNothing(sbSrc) Then sbSrc = sbLireFichier(sSrc)

        sbDest = New StringBuilder

        ' Utilisation d'une expression régulière pour découper en phrases
        Dim sGm$ = Chr(iCodeASCIIGuillemet)
        Dim pattern As String = "(?<=[\.!\?;:—]|\." + sGm + ")\s+" 'Avec Guillemets
        Dim phrases() As String = Regex.Split(sbSrc.ToString, pattern)

        ' Affichage des phrases
        For Each phrase As String In phrases
            sbDest.AppendLine(phrase)
        Next

        Return True

    End Function

    Public Function bEnleverPonctuation(sSrc$,
        ByRef sbSrc As StringBuilder,
        ByRef sbDest As StringBuilder,
        bOptionComparerMots As Boolean,
        bOptionComparerParag As Boolean,
        bOptionComparerNum As Boolean) As Boolean

        If IsNothing(sbSrc) Then sbSrc = sbLireFichier(sSrc)
        sbDest = New StringBuilder

        ' Ne marche pas :
        'sbDest.Append(sbSrc.Replace("-;".ToCharArray, ""))

        ' Rechercher tous les mots de la chaine : \w+
        Const sRechMots$ = "\w+"

        ' Découpage par phrase
        Dim bSupprDblSautDeLignes As Boolean = Not bOptionComparerParag
        'Const bOptionComparerMots As Boolean = True
        Dim acSepPhrase() As Char = sListeSeparateursPhrase.ToCharArray
        Dim asParag$() = sbSrc.ToString.Split(CChar(vbCrLf))

        Const bDebugSplit As Boolean = False
        If bDebugSplit Then Debug.WriteLine("-[" & sbSrc.ToString & "]-")

        Dim bSautDejaFait As Boolean = False
        Dim iNumParag% = 0
        For Each sParag As String In asParag
            iNumParag += 1

            If bDebugSplit Then Debug.WriteLine("Parag. n°" & iNumParag & " : [" & sParag & "]")

            'Dim bParagVide As Boolean
            'bParagVide = False
            If sParag.Length = 0 Then Continue For
            'If sParag.Length = 0 Then bParagVide = True : GoTo ParagSuiv 

            'Dim bSautDeParagDejaFait As Boolean
            'bSautDeParagDejaFait = False
            'Dim c As Char = sParag.Chars(0)
            'If c = vbLf Then
            '    If bDebugSplit Then Debug.WriteLine("vbLf")
            '    sbDest.Append(vbCrLf)
            '    bSautDeParagDejaFait = True
            'End If

            Dim asPhrases$() = sParag.Split(acSepPhrase)
            Dim iNumPhrase% = 0
            For Each sPhrase As String In asPhrases
                iNumPhrase += 1

                'Dim bPhraseVide As Boolean
                'bPhraseVide = False
                If sPhrase.Length = 0 Then Continue For
                'If sPhrase.Length = 0 Then bPhraseVide = True : GoTo PhraseSuiv 

                ' 26/11/2021 Si une phrase ne contient qu'un " alors ignorer
                Dim iLen% = sPhrase.Length
                If iLen = 1 Then
                    Dim c1 As Char = sPhrase.Chars(0)
                    If Asc(c1) = iCodeASCIIGuillemet Then Continue For
                End If

                Dim matches As MatchCollection = Regex.Matches(sPhrase, sRechMots)
                For i As Integer = 0 To matches.Count - 1
                    bSautDejaFait = False
                    Dim sMot$ = matches(i).ToString

                    ' 19/01/2018
                    If Not bOptionComparerNum AndAlso bOptionComparerMots AndAlso
                        IsNumeric(sMot) Then
                        bSautDejaFait = True ' Eviter un saut de ligne, puisque le mot est ignoré
                        Continue For
                    End If

                    sbDest.Append(sMot)
                    If bOptionComparerMots Then
                        sbDest.Append(vbCrLf)
                        bSautDejaFait = True
                    Else
                        sbDest.Append(" ")
                    End If
                Next

                'PhraseSuiv:
                If Not bSupprDblSautDeLignes OrElse Not bSautDejaFait Then sbDest.Append(vbCrLf)

            Next

            'ParagSuiv:
            If Not bSupprDblSautDeLignes OrElse Not bSautDejaFait Then sbDest.Append(vbCrLf)

        Next

        bEnleverPonctuation = True

    End Function

    Public Sub FusionnerMotsCoupes(sbMotsSrc As StringBuilder,
        ByRef sbMotsDest As StringBuilder, iNumFichier%, sCheminSrcOrig$)

        ' Fusionner les mots coupés éventuels dans la mesure où
        '  un mot concaténé est plus fréquent que chacun des tronçons

        Dim sb As New StringBuilder()

        Const bDebug0 As Boolean = False
        If bDebug0 Then
            Debug.WriteLine("")
            Debug.WriteLine("Fusion du fichier :" & sbMotsSrc.Length)
        End If

        sbMotsDest = New StringBuilder
        Dim asLignes$() = sbMotsSrc.ToString.Split(CChar(vbCrLf))
        Dim dico As New Dictionary(Of String, Integer) ' sClé : sMot -> iNbMots
        ' Compter la fréquence de chaque mot
        For Each sMot As String In asLignes
            Dim sMot2$ = sMot.Trim
            If dico.ContainsKey(sMot2) Then
                dico(sMot2) += 1
            Else
                dico(sMot2) = 1
            End If
        Next
        ' Vérifier si un mot concaténé avec le suivant est plus fréquent
        Dim dicoVerif As New DicoTri(Of String, clsMot) ' sClé : sMotConcat -> clsMot
        Dim iNbMots% = asLignes.GetUpperBound(0)
        Dim iNumMot% = 0
        Dim bFusion As Boolean = False
        Do While iNumMot < iNbMots
            bFusion = False
            Dim sMot$ = asLignes(iNumMot).Trim
            Dim sMotSuiv$ = asLignes(iNumMot + 1).Trim
            If sMot.Length <= 1 OrElse sMotSuiv.Length <= 1 Then
                sbMotsDest.AppendLine(sMot)
                GoTo MotSuivant
            End If
            Dim sMotConcat$ = sMot & sMotSuiv
            Dim iNbOccMot% = dico(sMot)
            Dim iNbOccMotSuiv% = dico(sMotSuiv)
            If Not dico.ContainsKey(sMotConcat) Then
                sbMotsDest.AppendLine(sMot)
                GoTo MotSuivant
            End If
            Dim iNbOccMotConcat% = dico(sMotConcat)
            If iNbOccMot < iNbOccMotConcat AndAlso
               iNbOccMotSuiv < iNbOccMotConcat Then
                If bDebug0 Then _
                    Debug.WriteLine("Mot coupé potentiel : " &
                        sMotConcat & "(" & iNbOccMotConcat & ") " &
                        sMot & "(" & iNbOccMot & ") " &
                        sMotSuiv & "(" & iNbOccMotSuiv & ")")
                sbMotsDest.AppendLine(sMotConcat)
                Dim mot As New clsMot
                mot.sMotConcat = sMotConcat
                mot.iNbOccConcat = iNbOccMotConcat
                mot.sMot1 = sMot : mot.iNbOcc1 = iNbOccMot
                mot.sMot2 = sMotSuiv : mot.iNbOcc2 = iNbOccMotSuiv
                If dicoVerif.ContainsKey(sMotConcat) Then
                    ' Conserver la taille max.
                    Dim mot0 As clsMot = dicoVerif(sMotConcat)
                    If iNbOccMotConcat > mot0.iNbOccConcat Then mot0.iNbOccConcat = iNbOccMotConcat
                Else
                    dicoVerif.Add(sMotConcat, mot)
                End If
                iNumMot += 1
                bFusion = True
            Else
                sbMotsDest.AppendLine(sMot)
            End If
MotSuivant:
            iNumMot += 1
        Loop
        ' Ajouter le dernier mot le cas échéant
        If Not bFusion Then
            Dim sMot$ = asLignes(iNumMot).Trim
            If sMot.Length > 0 Then sbMotsDest.AppendLine(sMot)
        End If

        'If Not bDebug0 Then Exit Sub
        If bDebug0 Then
            Debug.WriteLine("")
            Debug.WriteLine("Tri par fréquence décroissante :")
        End If
        For Each mot As clsMot In dicoVerif.Trier(
            "iNbOccConcat DESC, iNbOcc1 DESC, iNbOcc2 DESC, sMotConcat")
            If bDebug0 Then _
                Debug.WriteLine("Mot coupé potentiel : " &
                    mot.sMotConcat & "(" & mot.iNbOccConcat & ") " &
                    mot.sMot1 & "(" & mot.iNbOcc1 & ") " &
                    mot.sMot2 & "(" & mot.iNbOcc2 & ")")
            If sb.Length = 0 Then
                sb.AppendLine("Fusion du fichier " & iNumFichier & " : " & sCheminSrcOrig)
                sb.AppendLine("(Occurrences du mot fusionné : occurrences du tronçon de début-occurrences du tronçon de fin)")
            End If
            sb.AppendLine(mot.sMotConcat & " : " & mot.sMot1 & "-" & mot.sMot2 &
                " (" & mot.iNbOccConcat & " : " & mot.iNbOcc1 & "-" & mot.iNbOcc2 & ")")
        Next

        Dim sCheminRapportFusion$ = Application.StartupPath & "\" &
            sFusion & iNumFichier & sExtTxt
        Dim sCheminRapportOrig$ = Application.StartupPath & "\" &
            sFichier & iNumFichier & "_" & sOrig & sExtTxt
        If sb.Length = 0 Then
            bSupprimerFichier(sCheminRapportFusion)
            bSupprimerFichier(sCheminRapportOrig)
            Exit Sub
        End If
        bEcrireFichier(sCheminRapportFusion, sb)
        bEcrireFichier(sCheminRapportOrig, sbMotsSrc)

    End Sub

    Public Sub NormaliserQuotes(sbSrc As StringBuilder, ByRef sbDest As StringBuilder)

        Dim sSepQuote$ = Chr(iCodeASCIIQuote)
        Dim sSepQuote2$ = Chr(iCodeASCIIQuote2)
        Dim sSepGmO2$ = Chr(iCodeASCIIGuillemetOuvrant2) ' ‘ ' 23/11/2014
        Dim sSepGmF2$ = Chr(iCodeASCIIGuillemetFermant2) ' ’
        Dim sSepGmF4$ = Chr(iCodeASCIIGuillemetFermant4) ' ´

        ' 18/11/2018 Dans ce cas, il doit y avoir un espace aussi
        '  Solution possible : commencer par remplacer avec espace (l'espace sera supprimé),
        '   puis sans espace (tester aussi l'espace insécable et l'espace fine insécable)
        Dim sSepGmO1$ = Chr(iCodeASCIIGuillemetOuvrant) ' «
        Dim sSepGmF1$ = Chr(iCodeASCIIGuillemetFermant) ' »
        Dim sSepGmO1E$ = sSepGmO1 & " "
        Dim sSepGmF1E$ = " " & sSepGmF1
        Dim sEspInsec$ = Chr(iCodeASCIIEspaceInsecable)
        Dim sEspInsecF$ = ChrW(iCodeUTF16EspaceFineInsecable)
        Dim sSepGmO1EI$ = sSepGmO1 & sEspInsec
        Dim sSepGmF1EI$ = sEspInsec & sSepGmF1
        Dim sSepGmO1EFI$ = sSepGmO1 & sEspInsecF
        Dim sSepGmF1EFI$ = sEspInsecF & sSepGmF1

        ' 20/07/2014
        Dim sSepGuill$ = Chr(iCodeASCIIGuillemet) ' "
        Dim sSepGmO3$ = Chr(iCodeASCIIGuillemetOuvrant3) ' “
        Dim sSepGmF3$ = Chr(iCodeASCIIGuillemetFermant3) ' ”

        sbDest = sbSrc.Replace(sSepQuote2, sSepQuote).
            Replace(sSepGmO2, sSepQuote).Replace(sSepGmF2, sSepQuote).
            Replace(sSepGmO3, sSepGuill).Replace(sSepGmF3, sSepGuill).
            Replace(sSepGmF4, sSepQuote).
            Replace(sSepGmO1EFI, sSepGuill).Replace(sSepGmF1EFI, sSepGuill).
            Replace(sSepGmO1EI, sSepGuill).Replace(sSepGmF1EI, sSepGuill).
            Replace(sSepGmO1E, sSepGuill).Replace(sSepGmF1E, sSepGuill).
            Replace(sSepGmO1, sSepGuill).Replace(sSepGmF1, sSepGuill) ' 18/11/2018

        'Dim sDest$ = sbDest.ToString
        'Dim iLong% = sDest.Length - 1
        'Debug.WriteLine(sDest)
        'Debug.WriteLine("Ouvrant : " & sDest.Chars(0) & " = " & Asc(sDest.Chars(0)))
        'Debug.WriteLine("Fermant : " & sDest.Chars(iLong) & " = " & Asc(sDest.Chars(iLong)))

    End Sub

End Module