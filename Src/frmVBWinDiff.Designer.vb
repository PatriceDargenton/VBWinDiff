<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVBWinDiff : Inherits Form

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()
        If bDebug Then Me.StartPosition = FormStartPosition.CenterScreen

    End Sub

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'éditeur de code.
    Friend WithEvents sbStatusBar As System.Windows.Forms.StatusBar
    Friend WithEvents cmdComp As System.Windows.Forms.Button
    Friend WithEvents cmdAnnuler As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVBWinDiff))
        Me.sbStatusBar = New System.Windows.Forms.StatusBar()
        Me.cmdComp = New System.Windows.Forms.Button()
        Me.cmdAnnuler = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdAjouterRaccourci = New System.Windows.Forms.Button()
        Me.cmdEnleverRaccourci = New System.Windows.Forms.Button()
        Me.chkInfo = New System.Windows.Forms.CheckBox()
        Me.chkPhrases = New System.Windows.Forms.CheckBox()
        Me.chkCasse = New System.Windows.Forms.CheckBox()
        Me.chkAccents = New System.Windows.Forms.CheckBox()
        Me.chkEspacesInsec = New System.Windows.Forms.CheckBox()
        Me.chkPonctuation = New System.Windows.Forms.CheckBox()
        Me.chkTout = New System.Windows.Forms.CheckBox()
        Me.cmdPagePreced = New System.Windows.Forms.Button()
        Me.cmdPageSuiv = New System.Windows.Forms.Button()
        Me.chkPaginer = New System.Windows.Forms.CheckBox()
        Me.chkRatio = New System.Windows.Forms.CheckBox()
        Me.chkParag = New System.Windows.Forms.CheckBox()
        Me.chkQuotes = New System.Windows.Forms.CheckBox()
        Me.chkEspaces = New System.Windows.Forms.CheckBox()
        Me.chkNum = New System.Windows.Forms.CheckBox()
        Me.lbAlgo = New System.Windows.Forms.ListBox()
        Me.lblChemin1 = New System.Windows.Forms.Label()
        Me.lblChemin2 = New System.Windows.Forms.Label()
        Me.lblNumPage = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'sbStatusBar
        '
        Me.sbStatusBar.Location = New System.Drawing.Point(0, 222)
        Me.sbStatusBar.Name = "sbStatusBar"
        Me.sbStatusBar.Size = New System.Drawing.Size(695, 22)
        Me.sbStatusBar.TabIndex = 0
        '
        'cmdComp
        '
        Me.cmdComp.Location = New System.Drawing.Point(321, 45)
        Me.cmdComp.Name = "cmdComp"
        Me.cmdComp.Size = New System.Drawing.Size(103, 32)
        Me.cmdComp.TabIndex = 1
        Me.cmdComp.Text = "Comparer"
        Me.ToolTip1.SetToolTip(Me.cmdComp, "Comparer les fichiers via WinDiff avec ces options")
        '
        'cmdAnnuler
        '
        Me.cmdAnnuler.Enabled = False
        Me.cmdAnnuler.Location = New System.Drawing.Point(459, 45)
        Me.cmdAnnuler.Name = "cmdAnnuler"
        Me.cmdAnnuler.Size = New System.Drawing.Size(103, 32)
        Me.cmdAnnuler.TabIndex = 2
        Me.cmdAnnuler.Text = "Annuler"
        Me.ToolTip1.SetToolTip(Me.cmdAnnuler, "Interrompre la requête en cours, et renvoyer les données déjà récupérées")
        '
        'cmdAjouterRaccourci
        '
        Me.cmdAjouterRaccourci.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAjouterRaccourci.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAjouterRaccourci.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAjouterRaccourci.Location = New System.Drawing.Point(321, 83)
        Me.cmdAjouterRaccourci.Name = "cmdAjouterRaccourci"
        Me.cmdAjouterRaccourci.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAjouterRaccourci.Size = New System.Drawing.Size(103, 25)
        Me.cmdAjouterRaccourci.TabIndex = 3
        Me.cmdAjouterRaccourci.Text = "Ajouter raccourci"
        Me.ToolTip1.SetToolTip(Me.cmdAjouterRaccourci, resources.GetString("cmdAjouterRaccourci.ToolTip"))
        Me.cmdAjouterRaccourci.UseVisualStyleBackColor = False
        '
        'cmdEnleverRaccourci
        '
        Me.cmdEnleverRaccourci.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEnleverRaccourci.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEnleverRaccourci.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnleverRaccourci.Location = New System.Drawing.Point(459, 83)
        Me.cmdEnleverRaccourci.Name = "cmdEnleverRaccourci"
        Me.cmdEnleverRaccourci.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEnleverRaccourci.Size = New System.Drawing.Size(103, 25)
        Me.cmdEnleverRaccourci.TabIndex = 4
        Me.cmdEnleverRaccourci.Text = "Enlever raccourci"
        Me.ToolTip1.SetToolTip(Me.cmdEnleverRaccourci, "Enlever le raccourci ""Envoyer vers"" (SendTo) vers VBWinDiff (depuis Windows Vista" &
        ", il faut préalablement lancer l'application en tant qu'admin. ou être admin.)")
        Me.cmdEnleverRaccourci.UseVisualStyleBackColor = False
        '
        'chkInfo
        '
        Me.chkInfo.AutoSize = True
        Me.chkInfo.Checked = True
        Me.chkInfo.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkInfo.Location = New System.Drawing.Point(321, 10)
        Me.chkInfo.Name = "chkInfo"
        Me.chkInfo.Size = New System.Drawing.Size(78, 17)
        Me.chkInfo.TabIndex = 19
        Me.chkInfo.Text = "Information"
        Me.ToolTip1.SetToolTip(Me.chkInfo, "Inclure les informations dans les fichiers textes pour faciliter le repérage des " &
        "deux fichiers dans WinDiff")
        Me.chkInfo.UseVisualStyleBackColor = True
        '
        'chkPhrases
        '
        Me.chkPhrases.AutoSize = True
        Me.chkPhrases.Checked = True
        Me.chkPhrases.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPhrases.Location = New System.Drawing.Point(104, 113)
        Me.chkPhrases.Name = "chkPhrases"
        Me.chkPhrases.Size = New System.Drawing.Size(64, 17)
        Me.chkPhrases.TabIndex = 11
        Me.chkPhrases.Text = "Phrases"
        Me.ToolTip1.SetToolTip(Me.chkPhrases, "Cocher pour comparer les phrases entières, sinon comparer le détail des mots (mai" &
        "s sans la ponctuation alors).")
        Me.chkPhrases.UseVisualStyleBackColor = True
        '
        'chkCasse
        '
        Me.chkCasse.AutoSize = True
        Me.chkCasse.Checked = True
        Me.chkCasse.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCasse.Location = New System.Drawing.Point(15, 73)
        Me.chkCasse.Name = "chkCasse"
        Me.chkCasse.Size = New System.Drawing.Size(55, 17)
        Me.chkCasse.TabIndex = 8
        Me.chkCasse.Text = "Casse"
        Me.ToolTip1.SetToolTip(Me.chkCasse, "Cocher pour prendre en compte la casse (majuscule/minuscule), sinon l'ignorer.")
        Me.chkCasse.UseVisualStyleBackColor = True
        '
        'chkAccents
        '
        Me.chkAccents.AutoSize = True
        Me.chkAccents.Checked = True
        Me.chkAccents.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAccents.Location = New System.Drawing.Point(15, 93)
        Me.chkAccents.Name = "chkAccents"
        Me.chkAccents.Size = New System.Drawing.Size(65, 17)
        Me.chkAccents.TabIndex = 9
        Me.chkAccents.Text = "Accents"
        Me.ToolTip1.SetToolTip(Me.chkAccents, "Cocher pour prendre en compte les accents, sinon les ignorer.")
        Me.chkAccents.UseVisualStyleBackColor = True
        '
        'chkEspacesInsec
        '
        Me.chkEspacesInsec.AutoSize = True
        Me.chkEspacesInsec.Location = New System.Drawing.Point(15, 30)
        Me.chkEspacesInsec.Name = "chkEspacesInsec"
        Me.chkEspacesInsec.Size = New System.Drawing.Size(118, 17)
        Me.chkEspacesInsec.TabIndex = 6
        Me.chkEspacesInsec.Text = "Espaces insécabes"
        Me.ToolTip1.SetToolTip(Me.chkEspacesInsec, "Cocher pour prendre en compte les espaces insécables (cela implique de conserver " &
        "la ponctuation), sinon les ignorer.")
        Me.chkEspacesInsec.UseVisualStyleBackColor = True
        '
        'chkPonctuation
        '
        Me.chkPonctuation.AutoSize = True
        Me.chkPonctuation.Checked = True
        Me.chkPonctuation.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPonctuation.Location = New System.Drawing.Point(15, 113)
        Me.chkPonctuation.Name = "chkPonctuation"
        Me.chkPonctuation.Size = New System.Drawing.Size(83, 17)
        Me.chkPonctuation.TabIndex = 10
        Me.chkPonctuation.Text = "Ponctuation"
        Me.ToolTip1.SetToolTip(Me.chkPonctuation, "Cocher pour prendre en compte la ponctuation, sinon l'ignorer.")
        Me.chkPonctuation.UseVisualStyleBackColor = True
        '
        'chkTout
        '
        Me.chkTout.AutoSize = True
        Me.chkTout.Location = New System.Drawing.Point(15, 10)
        Me.chkTout.Name = "chkTout"
        Me.chkTout.Size = New System.Drawing.Size(48, 17)
        Me.chkTout.TabIndex = 5
        Me.chkTout.Text = "Tout"
        Me.ToolTip1.SetToolTip(Me.chkTout, "Tout cocher/décocher")
        Me.chkTout.UseVisualStyleBackColor = True
        '
        'cmdPagePreced
        '
        Me.cmdPagePreced.Enabled = False
        Me.cmdPagePreced.Font = New System.Drawing.Font("Wingdings", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.cmdPagePreced.Location = New System.Drawing.Point(175, 46)
        Me.cmdPagePreced.Name = "cmdPagePreced"
        Me.cmdPagePreced.Size = New System.Drawing.Size(24, 24)
        Me.cmdPagePreced.TabIndex = 16
        Me.cmdPagePreced.Text = ""
        Me.ToolTip1.SetToolTip(Me.cmdPagePreced, "Comparer la page précédente")
        Me.cmdPagePreced.UseVisualStyleBackColor = True
        '
        'cmdPageSuiv
        '
        Me.cmdPageSuiv.Enabled = False
        Me.cmdPageSuiv.Font = New System.Drawing.Font("Wingdings", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.cmdPageSuiv.Location = New System.Drawing.Point(256, 46)
        Me.cmdPageSuiv.Name = "cmdPageSuiv"
        Me.cmdPageSuiv.Size = New System.Drawing.Size(24, 24)
        Me.cmdPageSuiv.TabIndex = 18
        Me.cmdPageSuiv.Text = ""
        Me.ToolTip1.SetToolTip(Me.cmdPageSuiv, "Comparer la page suivante")
        Me.cmdPageSuiv.UseVisualStyleBackColor = True
        '
        'chkPaginer
        '
        Me.chkPaginer.AutoSize = True
        Me.chkPaginer.Location = New System.Drawing.Point(175, 23)
        Me.chkPaginer.Name = "chkPaginer"
        Me.chkPaginer.Size = New System.Drawing.Size(62, 17)
        Me.chkPaginer.TabIndex = 14
        Me.chkPaginer.Text = "Paginer"
        Me.ToolTip1.SetToolTip(Me.chkPaginer, "Découper en pages pour accélérer WinDiff (utile dans le cas de la comparaison mot" &
        " à mot)")
        Me.chkPaginer.UseVisualStyleBackColor = True
        '
        'chkRatio
        '
        Me.chkRatio.AutoSize = True
        Me.chkRatio.Location = New System.Drawing.Point(243, 23)
        Me.chkRatio.Name = "chkRatio"
        Me.chkRatio.Size = New System.Drawing.Size(51, 17)
        Me.chkRatio.TabIndex = 15
        Me.chkRatio.Text = "Ratio"
        Me.ToolTip1.SetToolTip(Me.chkRatio, "Appliquer un ratio de façon à comparer deux textes comme s'ils avaient un contenu" &
        " identique (par exemple si l'un a des retours à la ligne à chaque ligne)")
        Me.chkRatio.UseVisualStyleBackColor = True
        '
        'chkParag
        '
        Me.chkParag.AutoSize = True
        Me.chkParag.Checked = True
        Me.chkParag.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkParag.Location = New System.Drawing.Point(173, 113)
        Me.chkParag.Name = "chkParag"
        Me.chkParag.Size = New System.Drawing.Size(86, 17)
        Me.chkParag.TabIndex = 12
        Me.chkParag.Text = "Paragraphes"
        Me.ToolTip1.SetToolTip(Me.chkParag, "Cocher pour comparer les paragraphes tels quels, sinon comparer les phrases indép" &
        "endamment des paragraphes.")
        Me.chkParag.UseVisualStyleBackColor = True
        '
        'chkQuotes
        '
        Me.chkQuotes.AutoSize = True
        Me.chkQuotes.Location = New System.Drawing.Point(15, 136)
        Me.chkQuotes.Name = "chkQuotes"
        Me.chkQuotes.Size = New System.Drawing.Size(60, 17)
        Me.chkQuotes.TabIndex = 13
        Me.chkQuotes.Text = "Quotes"
        Me.ToolTip1.SetToolTip(Me.chkQuotes, "Décocher pour normaliser les quotes (apostrophes).")
        Me.chkQuotes.UseVisualStyleBackColor = True
        '
        'chkEspaces
        '
        Me.chkEspaces.AutoSize = True
        Me.chkEspaces.Location = New System.Drawing.Point(15, 50)
        Me.chkEspaces.Name = "chkEspaces"
        Me.chkEspaces.Size = New System.Drawing.Size(67, 17)
        Me.chkEspaces.TabIndex = 7
        Me.chkEspaces.Text = "Espaces"
        Me.ToolTip1.SetToolTip(Me.chkEspaces, "Cocher pour prendre en compte les espaces en début ou fin de paragraphe ainsi que" &
        " les sauts de ligne multiples (cela implique de conserver la ponctuation), sinon" &
        " les ignorer.")
        Me.chkEspaces.UseVisualStyleBackColor = True
        '
        'chkNum
        '
        Me.chkNum.AutoSize = True
        Me.chkNum.Checked = True
        Me.chkNum.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkNum.Location = New System.Drawing.Point(265, 114)
        Me.chkNum.Name = "chkNum"
        Me.chkNum.Size = New System.Drawing.Size(82, 17)
        Me.chkNum.TabIndex = 23
        Me.chkNum.Text = "Numériques"
        Me.ToolTip1.SetToolTip(Me.chkNum, "Cocher pour prendre en compte les numériques (n° de page, ...), sinon les ignorer" &
        ".")
        Me.chkNum.UseVisualStyleBackColor = True
        '
        'lbAlgo
        '
        Me.lbAlgo.FormattingEnabled = True
        Me.lbAlgo.Location = New System.Drawing.Point(584, 34)
        Me.lbAlgo.Name = "lbAlgo"
        Me.lbAlgo.Size = New System.Drawing.Size(86, 43)
        Me.lbAlgo.TabIndex = 24
        Me.ToolTip1.SetToolTip(Me.lbAlgo, "Algorithme de recherche des différences")
        '
        'lblChemin1
        '
        Me.lblChemin1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblChemin1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblChemin1.Location = New System.Drawing.Point(12, 168)
        Me.lblChemin1.Name = "lblChemin1"
        Me.lblChemin1.Size = New System.Drawing.Size(671, 15)
        Me.lblChemin1.TabIndex = 21
        Me.lblChemin1.Text = "Chemin1"
        '
        'lblChemin2
        '
        Me.lblChemin2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblChemin2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblChemin2.Location = New System.Drawing.Point(12, 192)
        Me.lblChemin2.Name = "lblChemin2"
        Me.lblChemin2.Size = New System.Drawing.Size(671, 15)
        Me.lblChemin2.TabIndex = 22
        Me.lblChemin2.Text = "Chemin2"
        '
        'lblNumPage
        '
        Me.lblNumPage.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblNumPage.Enabled = False
        Me.lblNumPage.Location = New System.Drawing.Point(205, 48)
        Me.lblNumPage.Name = "lblNumPage"
        Me.lblNumPage.Size = New System.Drawing.Size(45, 19)
        Me.lblNumPage.TabIndex = 17
        Me.lblNumPage.Text = "1/1"
        '
        'frmVBWinDiff
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(695, 244)
        Me.Controls.Add(Me.lbAlgo)
        Me.Controls.Add(Me.chkNum)
        Me.Controls.Add(Me.chkEspaces)
        Me.Controls.Add(Me.chkQuotes)
        Me.Controls.Add(Me.chkParag)
        Me.Controls.Add(Me.chkRatio)
        Me.Controls.Add(Me.chkPaginer)
        Me.Controls.Add(Me.lblNumPage)
        Me.Controls.Add(Me.cmdPagePreced)
        Me.Controls.Add(Me.cmdPageSuiv)
        Me.Controls.Add(Me.chkTout)
        Me.Controls.Add(Me.chkPhrases)
        Me.Controls.Add(Me.lblChemin2)
        Me.Controls.Add(Me.lblChemin1)
        Me.Controls.Add(Me.chkInfo)
        Me.Controls.Add(Me.chkPonctuation)
        Me.Controls.Add(Me.chkEspacesInsec)
        Me.Controls.Add(Me.chkAccents)
        Me.Controls.Add(Me.chkCasse)
        Me.Controls.Add(Me.cmdAjouterRaccourci)
        Me.Controls.Add(Me.cmdEnleverRaccourci)
        Me.Controls.Add(Me.cmdAnnuler)
        Me.Controls.Add(Me.cmdComp)
        Me.Controls.Add(Me.sbStatusBar)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmVBWinDiff"
        Me.Text = "VBWinDiff : Interface d'options pour le comparateur WinDiff"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents cmdAjouterRaccourci As System.Windows.Forms.Button
    Public WithEvents cmdEnleverRaccourci As System.Windows.Forms.Button
    Friend WithEvents chkCasse As System.Windows.Forms.CheckBox
    Friend WithEvents chkAccents As System.Windows.Forms.CheckBox
    Friend WithEvents chkEspacesInsec As System.Windows.Forms.CheckBox
    Friend WithEvents chkPonctuation As System.Windows.Forms.CheckBox
    Friend WithEvents chkInfo As System.Windows.Forms.CheckBox
    Friend WithEvents lblChemin1 As System.Windows.Forms.Label
    Friend WithEvents lblChemin2 As System.Windows.Forms.Label
    Friend WithEvents chkPhrases As System.Windows.Forms.CheckBox
    Friend WithEvents chkTout As System.Windows.Forms.CheckBox
    Friend WithEvents cmdPagePreced As System.Windows.Forms.Button
    Friend WithEvents cmdPageSuiv As System.Windows.Forms.Button
    Friend WithEvents lblNumPage As System.Windows.Forms.Label
    Friend WithEvents chkPaginer As System.Windows.Forms.CheckBox
    Friend WithEvents chkRatio As System.Windows.Forms.CheckBox
    Friend WithEvents chkParag As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuotes As System.Windows.Forms.CheckBox
    Friend WithEvents chkEspaces As System.Windows.Forms.CheckBox
    Friend WithEvents chkNum As System.Windows.Forms.CheckBox
    Friend WithEvents lbAlgo As ListBox
End Class
