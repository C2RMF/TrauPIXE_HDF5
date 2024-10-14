Imports System.IO
Imports System.Numerics

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_Traupixe_H5_2024
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.LstPar_Mat = New System.Windows.Forms.ListBox()
        Me.LstPar_Trc = New System.Windows.Forms.ListBox()
        Me.ComboBoxDrive = New System.Windows.Forms.ComboBox()
        Me.trvFolders = New System.Windows.Forms.TreeView()
        Me.LvFiles = New System.Windows.Forms.ListView()
        Me.BtRefresh = New System.Windows.Forms.Button()
        Me.CbDetMat = New System.Windows.Forms.ComboBox()
        Me.Par_det0 = New System.Windows.Forms.TextBox()
        Me.Par_det1 = New System.Windows.Forms.TextBox()
        Me.Check_det0 = New System.Windows.Forms.CheckBox()
        Me.Check_det1 = New System.Windows.Forms.CheckBox()
        Me.Check_det2 = New System.Windows.Forms.CheckBox()
        Me.Par_det2 = New System.Windows.Forms.TextBox()
        Me.Par_det5 = New System.Windows.Forms.TextBox()
        Me.Check_det5 = New System.Windows.Forms.CheckBox()
        Me.Check_det4 = New System.Windows.Forms.CheckBox()
        Me.Check_det3 = New System.Windows.Forms.CheckBox()
        Me.Par_det4 = New System.Windows.Forms.TextBox()
        Me.Par_det3 = New System.Windows.Forms.TextBox()
        Me.Par_det8 = New System.Windows.Forms.TextBox()
        Me.Check_det8 = New System.Windows.Forms.CheckBox()
        Me.Check_det7 = New System.Windows.Forms.CheckBox()
        Me.Check_det6 = New System.Windows.Forms.CheckBox()
        Me.Par_det7 = New System.Windows.Forms.TextBox()
        Me.Par_det6 = New System.Windows.Forms.TextBox()
        Me.Par_Mat = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Nb_Proc = New System.Windows.Forms.TextBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.ComboBox_Type_Calc = New System.Windows.Forms.ComboBox()
        Me.Progress = New System.Windows.Forms.TextBox()
        Me.Tps_Calc = New System.Windows.Forms.TextBox()
        Me.MyPause = New System.Windows.Forms.Button()
        Me.Text_Status = New System.Windows.Forms.TextBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.mnuOxyde = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOxydeOUI = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOxydeNON = New System.Windows.Forms.ToolStripMenuItem()
        Me.LODToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GupixLODNWrite0ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WriteAllValuesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RoundConcentrationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SkipPbMatrixToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Pivot_det0 = New System.Windows.Forms.TextBox()
        Me.Pivot_det1 = New System.Windows.Forms.TextBox()
        Me.Pivot_det2 = New System.Windows.Forms.TextBox()
        Me.Pivot_det5 = New System.Windows.Forms.TextBox()
        Me.Pivot_det4 = New System.Windows.Forms.TextBox()
        Me.Pivot_det3 = New System.Windows.Forms.TextBox()
        Me.Pivot_det8 = New System.Windows.Forms.TextBox()
        Me.Pivot_det7 = New System.Windows.Forms.TextBox()
        Me.Pivot_det6 = New System.Windows.Forms.TextBox()
        Me.Same_Z = New System.Windows.Forms.CheckBox()
        Me.TextXLS = New System.Windows.Forms.TextBox()
        Me.LabelAppend = New System.Windows.Forms.Label()
        Me.LabelNew = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Check_Trc_As_Oxy = New System.Windows.Forms.CheckBox()
        Me.Text_Lst_Ox_Trc = New System.Windows.Forms.TextBox()
        Me.Ck_AllAsOxy = New System.Windows.Forms.CheckBox()
        Me.ListFileInit = New System.Windows.Forms.ListView()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Chk_RoundValue = New System.Windows.Forms.CheckBox()
        Me.Adjust_Filter = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.Calcul = New System.Windows.Forms.TabPage()
        Me.chb_skip_elem = New System.Windows.Forms.CheckBox()
        Me.txt_skip_elem = New System.Windows.Forms.TextBox()
        Me.chk_external_ok = New System.Windows.Forms.CheckBox()
        Me.Text_gamma = New System.Windows.Forms.TextBox()
        Me.Tab_Adjust = New System.Windows.Forms.TabPage()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextF_Z = New System.Windows.Forms.TextBox()
        Me.TextF_Step = New System.Windows.Forms.TextBox()
        Me.TextF_To = New System.Windows.Forms.TextBox()
        Me.ComboBox_Type_F = New System.Windows.Forms.ComboBox()
        Me.TextF_From = New System.Windows.Forms.TextBox()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Button_Extract = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Box_txtFiltre = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ListBox_HDF5 = New System.Windows.Forms.ListBox()
        Me.TxtBox_HDF5_File = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button_Run2 = New System.Windows.Forms.Button()
        Me.TextProcessIf = New System.Windows.Forms.TextBox()
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.Calcul.SuspendLayout()
        Me.Tab_Adjust.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'LstPar_Mat
        '
        Me.LstPar_Mat.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.LstPar_Mat.FormattingEnabled = True
        Me.LstPar_Mat.Location = New System.Drawing.Point(7, 36)
        Me.LstPar_Mat.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.LstPar_Mat.Name = "LstPar_Mat"
        Me.LstPar_Mat.Size = New System.Drawing.Size(330, 30)
        Me.LstPar_Mat.TabIndex = 7
        '
        'LstPar_Trc
        '
        Me.LstPar_Trc.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.LstPar_Trc.FormattingEnabled = True
        Me.LstPar_Trc.Location = New System.Drawing.Point(7, 207)
        Me.LstPar_Trc.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.LstPar_Trc.Name = "LstPar_Trc"
        Me.LstPar_Trc.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.LstPar_Trc.Size = New System.Drawing.Size(336, 30)
        Me.LstPar_Trc.TabIndex = 9
        '
        'ComboBoxDrive
        '
        Me.ComboBoxDrive.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.ComboBoxDrive.FormattingEnabled = True
        Me.ComboBoxDrive.Location = New System.Drawing.Point(7, 36)
        Me.ComboBoxDrive.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ComboBoxDrive.Name = "ComboBoxDrive"
        Me.ComboBoxDrive.Size = New System.Drawing.Size(301, 21)
        Me.ComboBoxDrive.TabIndex = 10
        '
        'trvFolders
        '
        Me.trvFolders.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.trvFolders.Location = New System.Drawing.Point(4, 117)
        Me.trvFolders.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.trvFolders.Name = "trvFolders"
        Me.trvFolders.ShowPlusMinus = False
        Me.trvFolders.ShowRootLines = False
        Me.trvFolders.Size = New System.Drawing.Size(304, 394)
        Me.trvFolders.TabIndex = 11
        '
        'LvFiles
        '
        Me.LvFiles.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.LvFiles.GridLines = True
        Me.LvFiles.Location = New System.Drawing.Point(316, 117)
        Me.LvFiles.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.LvFiles.Name = "LvFiles"
        Me.LvFiles.Size = New System.Drawing.Size(368, 394)
        Me.LvFiles.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.LvFiles.TabIndex = 12
        Me.LvFiles.UseCompatibleStateImageBehavior = False
        Me.LvFiles.View = System.Windows.Forms.View.List
        '
        'BtRefresh
        '
        Me.BtRefresh.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.BtRefresh.Location = New System.Drawing.Point(317, 33)
        Me.BtRefresh.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.BtRefresh.Name = "BtRefresh"
        Me.BtRefresh.Size = New System.Drawing.Size(132, 25)
        Me.BtRefresh.TabIndex = 13
        Me.BtRefresh.Text = "Refresh spectra"
        Me.BtRefresh.UseVisualStyleBackColor = True
        '
        'CbDetMat
        '
        Me.CbDetMat.FormattingEnabled = True
        Me.CbDetMat.Location = New System.Drawing.Point(7, 6)
        Me.CbDetMat.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CbDetMat.Name = "CbDetMat"
        Me.CbDetMat.Size = New System.Drawing.Size(59, 23)
        Me.CbDetMat.TabIndex = 14
        Me.CbDetMat.Text = "X0"
        '
        'Par_det0
        '
        Me.Par_det0.Location = New System.Drawing.Point(72, 6)
        Me.Par_det0.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_det0.Name = "Par_det0"
        Me.Par_det0.Size = New System.Drawing.Size(189, 23)
        Me.Par_det0.TabIndex = 16
        '
        'Par_det1
        '
        Me.Par_det1.Location = New System.Drawing.Point(72, 28)
        Me.Par_det1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_det1.Name = "Par_det1"
        Me.Par_det1.Size = New System.Drawing.Size(189, 23)
        Me.Par_det1.TabIndex = 17
        '
        'Check_det0
        '
        Me.Check_det0.AutoSize = True
        Me.Check_det0.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Check_det0.Location = New System.Drawing.Point(4, 8)
        Me.Check_det0.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_det0.Name = "Check_det0"
        Me.Check_det0.Size = New System.Drawing.Size(46, 19)
        Me.Check_det0.TabIndex = 22
        Me.Check_det0.Text = "BE0"
        Me.Check_det0.UseVisualStyleBackColor = True
        '
        'Check_det1
        '
        Me.Check_det1.AutoSize = True
        Me.Check_det1.Location = New System.Drawing.Point(4, 29)
        Me.Check_det1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_det1.Name = "Check_det1"
        Me.Check_det1.Size = New System.Drawing.Size(47, 19)
        Me.Check_det1.TabIndex = 23
        Me.Check_det1.Text = "HE1"
        Me.Check_det1.UseVisualStyleBackColor = True
        '
        'Check_det2
        '
        Me.Check_det2.AutoSize = True
        Me.Check_det2.Location = New System.Drawing.Point(4, 51)
        Me.Check_det2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_det2.Name = "Check_det2"
        Me.Check_det2.Size = New System.Drawing.Size(47, 19)
        Me.Check_det2.TabIndex = 24
        Me.Check_det2.Text = "HE2"
        Me.Check_det2.UseVisualStyleBackColor = True
        '
        'Par_det2
        '
        Me.Par_det2.Location = New System.Drawing.Point(72, 50)
        Me.Par_det2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_det2.Name = "Par_det2"
        Me.Par_det2.Size = New System.Drawing.Size(189, 23)
        Me.Par_det2.TabIndex = 25
        '
        'Par_det5
        '
        Me.Par_det5.Location = New System.Drawing.Point(72, 115)
        Me.Par_det5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_det5.Name = "Par_det5"
        Me.Par_det5.Size = New System.Drawing.Size(189, 23)
        Me.Par_det5.TabIndex = 31
        '
        'Check_det5
        '
        Me.Check_det5.AutoSize = True
        Me.Check_det5.BackColor = System.Drawing.Color.LightGray
        Me.Check_det5.Location = New System.Drawing.Point(4, 116)
        Me.Check_det5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_det5.Name = "Check_det5"
        Me.Check_det5.Size = New System.Drawing.Size(54, 19)
        Me.Check_det5.TabIndex = 30
        Me.Check_det5.Text = "HE10"
        Me.Check_det5.UseVisualStyleBackColor = False
        '
        'Check_det4
        '
        Me.Check_det4.AutoSize = True
        Me.Check_det4.Location = New System.Drawing.Point(4, 94)
        Me.Check_det4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_det4.Name = "Check_det4"
        Me.Check_det4.Size = New System.Drawing.Size(47, 19)
        Me.Check_det4.TabIndex = 29
        Me.Check_det4.Text = "HE4"
        Me.Check_det4.UseVisualStyleBackColor = True
        '
        'Check_det3
        '
        Me.Check_det3.AutoSize = True
        Me.Check_det3.Location = New System.Drawing.Point(4, 73)
        Me.Check_det3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_det3.Name = "Check_det3"
        Me.Check_det3.Size = New System.Drawing.Size(47, 19)
        Me.Check_det3.TabIndex = 28
        Me.Check_det3.Text = "HE3"
        Me.Check_det3.UseVisualStyleBackColor = True
        '
        'Par_det4
        '
        Me.Par_det4.Location = New System.Drawing.Point(72, 93)
        Me.Par_det4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_det4.Name = "Par_det4"
        Me.Par_det4.Size = New System.Drawing.Size(189, 23)
        Me.Par_det4.TabIndex = 27
        '
        'Par_det3
        '
        Me.Par_det3.Location = New System.Drawing.Point(72, 71)
        Me.Par_det3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_det3.Name = "Par_det3"
        Me.Par_det3.Size = New System.Drawing.Size(189, 23)
        Me.Par_det3.TabIndex = 26
        '
        'Par_det8
        '
        Me.Par_det8.Location = New System.Drawing.Point(72, 180)
        Me.Par_det8.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_det8.Name = "Par_det8"
        Me.Par_det8.Size = New System.Drawing.Size(189, 23)
        Me.Par_det8.TabIndex = 37
        '
        'Check_det8
        '
        Me.Check_det8.AutoSize = True
        Me.Check_det8.Location = New System.Drawing.Point(4, 182)
        Me.Check_det8.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_det8.Name = "Check_det8"
        Me.Check_det8.Size = New System.Drawing.Size(54, 19)
        Me.Check_det8.TabIndex = 36
        Me.Check_det8.Text = "HE13"
        Me.Check_det8.UseVisualStyleBackColor = True
        '
        'Check_det7
        '
        Me.Check_det7.AutoSize = True
        Me.Check_det7.Location = New System.Drawing.Point(4, 160)
        Me.Check_det7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_det7.Name = "Check_det7"
        Me.Check_det7.Size = New System.Drawing.Size(54, 19)
        Me.Check_det7.TabIndex = 35
        Me.Check_det7.Text = "HE12"
        Me.Check_det7.UseVisualStyleBackColor = True
        '
        'Check_det6
        '
        Me.Check_det6.AutoSize = True
        Me.Check_det6.Location = New System.Drawing.Point(4, 138)
        Me.Check_det6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_det6.Name = "Check_det6"
        Me.Check_det6.Size = New System.Drawing.Size(54, 19)
        Me.Check_det6.TabIndex = 34
        Me.Check_det6.Text = "HE11"
        Me.Check_det6.UseVisualStyleBackColor = True
        '
        'Par_det7
        '
        Me.Par_det7.Location = New System.Drawing.Point(72, 158)
        Me.Par_det7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_det7.Name = "Par_det7"
        Me.Par_det7.Size = New System.Drawing.Size(189, 23)
        Me.Par_det7.TabIndex = 33
        '
        'Par_det6
        '
        Me.Par_det6.Location = New System.Drawing.Point(72, 136)
        Me.Par_det6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_det6.Name = "Par_det6"
        Me.Par_det6.Size = New System.Drawing.Size(189, 23)
        Me.Par_det6.TabIndex = 32
        '
        'Par_Mat
        '
        Me.Par_Mat.Location = New System.Drawing.Point(75, 6)
        Me.Par_Mat.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Par_Mat.Name = "Par_Mat"
        Me.Par_Mat.Size = New System.Drawing.Size(261, 23)
        Me.Par_Mat.TabIndex = 38
        Me.Par_Mat.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Button2.Location = New System.Drawing.Point(8, 86)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(102, 38)
        Me.Button2.TabIndex = 39
        Me.Button2.Text = "Run"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Gray
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(700, 25)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 17)
        Me.Label1.TabIndex = 40
        Me.Label1.Text = "Matrix"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.LightGray
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label2.Location = New System.Drawing.Point(703, 123)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 17)
        Me.Label2.TabIndex = 41
        Me.Label2.Text = "Trace"
        '
        'Nb_Proc
        '
        Me.Nb_Proc.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Nb_Proc.Location = New System.Drawing.Point(307, 128)
        Me.Nb_Proc.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Nb_Proc.Name = "Nb_Proc"
        Me.Nb_Proc.Size = New System.Drawing.Size(20, 21)
        Me.Nb_Proc.TabIndex = 43
        Me.Nb_Proc.Text = "4"
        Me.Nb_Proc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(78, 538)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(606, 16)
        Me.ProgressBar1.TabIndex = 44
        '
        'ComboBox_Type_Calc
        '
        Me.ComboBox_Type_Calc.FormattingEnabled = True
        Me.ComboBox_Type_Calc.Items.AddRange(New Object() {"Ponctual", "EDF Map"})
        Me.ComboBox_Type_Calc.Location = New System.Drawing.Point(522, 62)
        Me.ComboBox_Type_Calc.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ComboBox_Type_Calc.Name = "ComboBox_Type_Calc"
        Me.ComboBox_Type_Calc.Size = New System.Drawing.Size(76, 23)
        Me.ComboBox_Type_Calc.TabIndex = 45
        Me.ComboBox_Type_Calc.Text = "Ponctual"
        Me.ComboBox_Type_Calc.Visible = False
        '
        'Progress
        '
        Me.Progress.Enabled = False
        Me.Progress.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Progress.Location = New System.Drawing.Point(131, 92)
        Me.Progress.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Progress.Name = "Progress"
        Me.Progress.Size = New System.Drawing.Size(74, 29)
        Me.Progress.TabIndex = 46
        Me.Progress.Text = "0 / 0"
        Me.Progress.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Tps_Calc
        '
        Me.Tps_Calc.Location = New System.Drawing.Point(125, 125)
        Me.Tps_Calc.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Tps_Calc.Name = "Tps_Calc"
        Me.Tps_Calc.Size = New System.Drawing.Size(88, 23)
        Me.Tps_Calc.TabIndex = 47
        Me.Tps_Calc.Text = "??:??:??"
        Me.Tps_Calc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.Tps_Calc.Visible = False
        '
        'MyPause
        '
        Me.MyPause.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.MyPause.Location = New System.Drawing.Point(228, 86)
        Me.MyPause.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MyPause.Name = "MyPause"
        Me.MyPause.Size = New System.Drawing.Size(102, 38)
        Me.MyPause.TabIndex = 48
        Me.MyPause.Text = "Pause"
        Me.MyPause.UseVisualStyleBackColor = True
        '
        'Text_Status
        '
        Me.Text_Status.Location = New System.Drawing.Point(116, 633)
        Me.Text_Status.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Text_Status.Name = "Text_Status"
        Me.Text_Status.Size = New System.Drawing.Size(578, 23)
        Me.Text_Status.TabIndex = 49
        Me.Text_Status.Text = "??:??:??"
        Me.Text_Status.Visible = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOxyde, Me.LODToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(7, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(1055, 24)
        Me.MenuStrip1.TabIndex = 50
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'mnuOxyde
        '
        Me.mnuOxyde.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOxydeOUI, Me.mnuOxydeNON})
        Me.mnuOxyde.Name = "mnuOxyde"
        Me.mnuOxyde.Size = New System.Drawing.Size(50, 20)
        Me.mnuOxyde.Text = "Oxide"
        '
        'mnuOxydeOUI
        '
        Me.mnuOxydeOUI.Checked = True
        Me.mnuOxydeOUI.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnuOxydeOUI.Name = "mnuOxydeOUI"
        Me.mnuOxydeOUI.Size = New System.Drawing.Size(91, 22)
        Me.mnuOxydeOUI.Text = "Yes"
        '
        'mnuOxydeNON
        '
        Me.mnuOxydeNON.Name = "mnuOxydeNON"
        Me.mnuOxydeNON.Size = New System.Drawing.Size(91, 22)
        Me.mnuOxydeNON.Text = "No"
        '
        'LODToolStripMenuItem
        '
        Me.LODToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GupixLODNWrite0ToolStripMenuItem, Me.WriteAllValuesToolStripMenuItem, Me.RoundConcentrationToolStripMenuItem, Me.SkipPbMatrixToolStripMenuItem})
        Me.LODToolStripMenuItem.Name = "LODToolStripMenuItem"
        Me.LODToolStripMenuItem.Size = New System.Drawing.Size(49, 20)
        Me.LODToolStripMenuItem.Text = "Setup"
        '
        'GupixLODNWrite0ToolStripMenuItem
        '
        Me.GupixLODNWrite0ToolStripMenuItem.Name = "GupixLODNWrite0ToolStripMenuItem"
        Me.GupixLODNWrite0ToolStripMenuItem.Size = New System.Drawing.Size(192, 22)
        Me.GupixLODNWrite0ToolStripMenuItem.Text = "Gupix LOD = N write 0"
        '
        'WriteAllValuesToolStripMenuItem
        '
        Me.WriteAllValuesToolStripMenuItem.Checked = True
        Me.WriteAllValuesToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.WriteAllValuesToolStripMenuItem.Name = "WriteAllValuesToolStripMenuItem"
        Me.WriteAllValuesToolStripMenuItem.Size = New System.Drawing.Size(192, 22)
        Me.WriteAllValuesToolStripMenuItem.Text = "Write all values"
        '
        'RoundConcentrationToolStripMenuItem
        '
        Me.RoundConcentrationToolStripMenuItem.Checked = True
        Me.RoundConcentrationToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.RoundConcentrationToolStripMenuItem.Name = "RoundConcentrationToolStripMenuItem"
        Me.RoundConcentrationToolStripMenuItem.Size = New System.Drawing.Size(192, 22)
        Me.RoundConcentrationToolStripMenuItem.Text = "Round Concentration"
        '
        'SkipPbMatrixToolStripMenuItem
        '
        Me.SkipPbMatrixToolStripMenuItem.Name = "SkipPbMatrixToolStripMenuItem"
        Me.SkipPbMatrixToolStripMenuItem.Size = New System.Drawing.Size(192, 22)
        Me.SkipPbMatrixToolStripMenuItem.Text = "Skip element in matrix"
        '
        'Pivot_det0
        '
        Me.Pivot_det0.Location = New System.Drawing.Point(268, 6)
        Me.Pivot_det0.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Pivot_det0.Name = "Pivot_det0"
        Me.Pivot_det0.Size = New System.Drawing.Size(75, 23)
        Me.Pivot_det0.TabIndex = 51
        Me.Pivot_det0.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Pivot_det1
        '
        Me.Pivot_det1.Location = New System.Drawing.Point(268, 28)
        Me.Pivot_det1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Pivot_det1.Name = "Pivot_det1"
        Me.Pivot_det1.Size = New System.Drawing.Size(75, 23)
        Me.Pivot_det1.TabIndex = 53
        Me.Pivot_det1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Pivot_det2
        '
        Me.Pivot_det2.Location = New System.Drawing.Point(268, 50)
        Me.Pivot_det2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Pivot_det2.Name = "Pivot_det2"
        Me.Pivot_det2.Size = New System.Drawing.Size(75, 23)
        Me.Pivot_det2.TabIndex = 54
        Me.Pivot_det2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Pivot_det5
        '
        Me.Pivot_det5.Location = New System.Drawing.Point(268, 115)
        Me.Pivot_det5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Pivot_det5.Name = "Pivot_det5"
        Me.Pivot_det5.Size = New System.Drawing.Size(75, 23)
        Me.Pivot_det5.TabIndex = 57
        Me.Pivot_det5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Pivot_det4
        '
        Me.Pivot_det4.Location = New System.Drawing.Point(268, 93)
        Me.Pivot_det4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Pivot_det4.Name = "Pivot_det4"
        Me.Pivot_det4.Size = New System.Drawing.Size(75, 23)
        Me.Pivot_det4.TabIndex = 56
        Me.Pivot_det4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Pivot_det3
        '
        Me.Pivot_det3.Location = New System.Drawing.Point(268, 71)
        Me.Pivot_det3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Pivot_det3.Name = "Pivot_det3"
        Me.Pivot_det3.Size = New System.Drawing.Size(75, 23)
        Me.Pivot_det3.TabIndex = 55
        Me.Pivot_det3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Pivot_det8
        '
        Me.Pivot_det8.Location = New System.Drawing.Point(268, 180)
        Me.Pivot_det8.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Pivot_det8.Name = "Pivot_det8"
        Me.Pivot_det8.Size = New System.Drawing.Size(75, 23)
        Me.Pivot_det8.TabIndex = 60
        Me.Pivot_det8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Pivot_det7
        '
        Me.Pivot_det7.Location = New System.Drawing.Point(268, 158)
        Me.Pivot_det7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Pivot_det7.Name = "Pivot_det7"
        Me.Pivot_det7.Size = New System.Drawing.Size(75, 23)
        Me.Pivot_det7.TabIndex = 59
        Me.Pivot_det7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Pivot_det6
        '
        Me.Pivot_det6.AccessibleDescription = ""
        Me.Pivot_det6.Location = New System.Drawing.Point(268, 136)
        Me.Pivot_det6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Pivot_det6.Name = "Pivot_det6"
        Me.Pivot_det6.Size = New System.Drawing.Size(75, 23)
        Me.Pivot_det6.TabIndex = 58
        Me.Pivot_det6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Same_Z
        '
        Me.Same_Z.AutoSize = True
        Me.Same_Z.Location = New System.Drawing.Point(624, 64)
        Me.Same_Z.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Same_Z.Name = "Same_Z"
        Me.Same_Z.Size = New System.Drawing.Size(55, 19)
        Me.Same_Z.TabIndex = 61
        Me.Same_Z.Text = "Same"
        Me.Same_Z.UseVisualStyleBackColor = True
        Me.Same_Z.Visible = False
        '
        'TextXLS
        '
        Me.TextXLS.Enabled = False
        Me.TextXLS.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.TextXLS.Location = New System.Drawing.Point(78, 516)
        Me.TextXLS.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TextXLS.Name = "TextXLS"
        Me.TextXLS.Size = New System.Drawing.Size(606, 22)
        Me.TextXLS.TabIndex = 62
        Me.TextXLS.Text = "Gupix-?????_Mat-??_Trc-??_Pivot-??.xls"
        '
        'LabelAppend
        '
        Me.LabelAppend.AutoSize = True
        Me.LabelAppend.Location = New System.Drawing.Point(4, 518)
        Me.LabelAppend.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelAppend.Name = "LabelAppend"
        Me.LabelAppend.Size = New System.Drawing.Size(63, 15)
        Me.LabelAppend.TabIndex = 64
        Me.LabelAppend.Text = "Append to"
        '
        'LabelNew
        '
        Me.LabelNew.AutoSize = True
        Me.LabelNew.Location = New System.Drawing.Point(4, 518)
        Me.LabelNew.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LabelNew.Name = "LabelNew"
        Me.LabelNew.Size = New System.Drawing.Size(62, 15)
        Me.LabelNew.TabIndex = 65
        Me.LabelNew.Text = "Create       "
        '
        'Check_Trc_As_Oxy
        '
        Me.Check_Trc_As_Oxy.AutoSize = True
        Me.Check_Trc_As_Oxy.Checked = True
        Me.Check_Trc_As_Oxy.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Check_Trc_As_Oxy.Location = New System.Drawing.Point(13, 8)
        Me.Check_Trc_As_Oxy.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Check_Trc_As_Oxy.Name = "Check_Trc_As_Oxy"
        Me.Check_Trc_As_Oxy.Size = New System.Drawing.Size(81, 19)
        Me.Check_Trc_As_Oxy.TabIndex = 74
        Me.Check_Trc_As_Oxy.Text = "Z as oxide"
        Me.Check_Trc_As_Oxy.UseVisualStyleBackColor = True
        '
        'Text_Lst_Ox_Trc
        '
        Me.Text_Lst_Ox_Trc.Location = New System.Drawing.Point(9, 28)
        Me.Text_Lst_Ox_Trc.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Text_Lst_Ox_Trc.Name = "Text_Lst_Ox_Trc"
        Me.Text_Lst_Ox_Trc.Size = New System.Drawing.Size(182, 23)
        Me.Text_Lst_Ox_Trc.TabIndex = 75
        Me.Text_Lst_Ox_Trc.Text = "19,20,25,26,29,82"
        '
        'Ck_AllAsOxy
        '
        Me.Ck_AllAsOxy.AutoSize = True
        Me.Ck_AllAsOxy.Location = New System.Drawing.Point(102, 8)
        Me.Ck_AllAsOxy.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Ck_AllAsOxy.Name = "Ck_AllAsOxy"
        Me.Ck_AllAsOxy.Size = New System.Drawing.Size(99, 19)
        Me.Ck_AllAsOxy.TabIndex = 76
        Me.Ck_AllAsOxy.Text = "All Z as oxide"
        Me.Ck_AllAsOxy.UseVisualStyleBackColor = True
        '
        'ListFileInit
        '
        Me.ListFileInit.GridLines = True
        Me.ListFileInit.Location = New System.Drawing.Point(1057, 146)
        Me.ListFileInit.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ListFileInit.Name = "ListFileInit"
        Me.ListFileInit.Size = New System.Drawing.Size(123, 159)
        Me.ListFileInit.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.ListFileInit.TabIndex = 79
        Me.ListFileInit.UseCompatibleStateImageBehavior = False
        Me.ListFileInit.View = System.Windows.Forms.View.List
        Me.ListFileInit.Visible = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2, Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 561)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 16, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1055, 22)
        Me.StatusStrip1.TabIndex = 80
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Arial Narrow", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.ToolStripStatusLabel1.Margin = New System.Windows.Forms.Padding(200, 3, 0, 2)
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(254, 17)
        Me.ToolStripStatusLabel1.Text = "                             Status TRAUPIXE                         "
        '
        'Chk_RoundValue
        '
        Me.Chk_RoundValue.AutoSize = True
        Me.Chk_RoundValue.Checked = True
        Me.Chk_RoundValue.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Chk_RoundValue.Location = New System.Drawing.Point(8, 126)
        Me.Chk_RoundValue.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Chk_RoundValue.Name = "Chk_RoundValue"
        Me.Chk_RoundValue.Size = New System.Drawing.Size(100, 19)
        Me.Chk_RoundValue.TabIndex = 81
        Me.Chk_RoundValue.Text = "Round values"
        Me.Chk_RoundValue.UseVisualStyleBackColor = True
        '
        'Adjust_Filter
        '
        Me.Adjust_Filter.Font = New System.Drawing.Font("Calibri", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Adjust_Filter.Location = New System.Drawing.Point(148, 64)
        Me.Adjust_Filter.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Adjust_Filter.Name = "Adjust_Filter"
        Me.Adjust_Filter.Size = New System.Drawing.Size(140, 38)
        Me.Adjust_Filter.TabIndex = 82
        Me.Adjust_Filter.Text = "Adjust Filter"
        Me.Adjust_Filter.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.TabControl1.Controls.Add(Me.Calcul)
        Me.TabControl1.Controls.Add(Me.Tab_Adjust)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.TabControl1.Location = New System.Drawing.Point(700, 394)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(350, 183)
        Me.TabControl1.TabIndex = 83
        '
        'Calcul
        '
        Me.Calcul.Controls.Add(Me.chb_skip_elem)
        Me.Calcul.Controls.Add(Me.txt_skip_elem)
        Me.Calcul.Controls.Add(Me.chk_external_ok)
        Me.Calcul.Controls.Add(Me.Text_gamma)
        Me.Calcul.Controls.Add(Me.Button2)
        Me.Calcul.Controls.Add(Me.Nb_Proc)
        Me.Calcul.Controls.Add(Me.Chk_RoundValue)
        Me.Calcul.Controls.Add(Me.MyPause)
        Me.Calcul.Controls.Add(Me.Check_Trc_As_Oxy)
        Me.Calcul.Controls.Add(Me.Text_Lst_Ox_Trc)
        Me.Calcul.Controls.Add(Me.Ck_AllAsOxy)
        Me.Calcul.Controls.Add(Me.Progress)
        Me.Calcul.Controls.Add(Me.Tps_Calc)
        Me.Calcul.Location = New System.Drawing.Point(4, 27)
        Me.Calcul.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Calcul.Name = "Calcul"
        Me.Calcul.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Calcul.Size = New System.Drawing.Size(342, 152)
        Me.Calcul.TabIndex = 0
        Me.Calcul.Text = "Process Spectra"
        Me.Calcul.UseVisualStyleBackColor = True
        '
        'chb_skip_elem
        '
        Me.chb_skip_elem.AutoSize = True
        Me.chb_skip_elem.Location = New System.Drawing.Point(242, 8)
        Me.chb_skip_elem.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.chb_skip_elem.Name = "chb_skip_elem"
        Me.chb_skip_elem.Size = New System.Drawing.Size(81, 19)
        Me.chb_skip_elem.TabIndex = 84
        Me.chb_skip_elem.Text = "skip Z mtx"
        Me.chb_skip_elem.UseVisualStyleBackColor = True
        Me.chb_skip_elem.Visible = False
        '
        'txt_skip_elem
        '
        Me.txt_skip_elem.Location = New System.Drawing.Point(242, 28)
        Me.txt_skip_elem.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txt_skip_elem.Name = "txt_skip_elem"
        Me.txt_skip_elem.Size = New System.Drawing.Size(77, 23)
        Me.txt_skip_elem.TabIndex = 83
        Me.txt_skip_elem.Text = "82"
        Me.txt_skip_elem.Visible = False
        '
        'chk_external_ok
        '
        Me.chk_external_ok.AutoSize = True
        Me.chk_external_ok.Checked = True
        Me.chk_external_ok.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk_external_ok.Location = New System.Drawing.Point(294, 60)
        Me.chk_external_ok.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.chk_external_ok.Name = "chk_external_ok"
        Me.chk_external_ok.Size = New System.Drawing.Size(43, 19)
        Me.chk_external_ok.TabIndex = 77
        Me.chk_external_ok.Text = "ON"
        Me.chk_external_ok.UseVisualStyleBackColor = True
        Me.chk_external_ok.Visible = False
        '
        'Text_gamma
        '
        Me.Text_gamma.Location = New System.Drawing.Point(8, 57)
        Me.Text_gamma.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Text_gamma.Name = "Text_gamma"
        Me.Text_gamma.Size = New System.Drawing.Size(278, 23)
        Me.Text_gamma.TabIndex = 82
        Me.Text_gamma.Text = "File ""external-conc.csv"" not found"
        '
        'Tab_Adjust
        '
        Me.Tab_Adjust.Controls.Add(Me.Label10)
        Me.Tab_Adjust.Controls.Add(Me.Label9)
        Me.Tab_Adjust.Controls.Add(Me.Label8)
        Me.Tab_Adjust.Controls.Add(Me.Label7)
        Me.Tab_Adjust.Controls.Add(Me.Label6)
        Me.Tab_Adjust.Controls.Add(Me.Label5)
        Me.Tab_Adjust.Controls.Add(Me.TextF_Z)
        Me.Tab_Adjust.Controls.Add(Me.TextF_Step)
        Me.Tab_Adjust.Controls.Add(Me.TextF_To)
        Me.Tab_Adjust.Controls.Add(Me.ComboBox_Type_F)
        Me.Tab_Adjust.Controls.Add(Me.TextF_From)
        Me.Tab_Adjust.Controls.Add(Me.Adjust_Filter)
        Me.Tab_Adjust.Location = New System.Drawing.Point(4, 27)
        Me.Tab_Adjust.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Tab_Adjust.Name = "Tab_Adjust"
        Me.Tab_Adjust.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Tab_Adjust.Size = New System.Drawing.Size(342, 152)
        Me.Tab_Adjust.TabIndex = 1
        Me.Tab_Adjust.Text = "Adjust Absorbers"
        Me.Tab_Adjust.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(9, 85)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(120, 15)
        Me.Label10.TabIndex = 88
        Me.Label10.Text = "mm / inert gas or Air"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 73)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(122, 15)
        Me.Label9.TabIndex = 87
        Me.Label9.Text = "µm / solid absorbers"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(114, 14)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(66, 15)
        Me.Label8.TabIndex = 85
        Me.Label8.Text = "Z absorber"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(198, 43)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(30, 15)
        Me.Label7.TabIndex = 85
        Me.Label7.Text = "Step"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(98, 41)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(19, 15)
        Me.Label6.TabIndex = 85
        Me.Label6.Text = "To"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 43)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(35, 15)
        Me.Label5.TabIndex = 84
        Me.Label5.Text = "From"
        '
        'TextF_Z
        '
        Me.TextF_Z.Location = New System.Drawing.Point(189, 10)
        Me.TextF_Z.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TextF_Z.Name = "TextF_Z"
        Me.TextF_Z.Size = New System.Drawing.Size(42, 23)
        Me.TextF_Z.TabIndex = 83
        Me.TextF_Z.Text = "102"
        Me.TextF_Z.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextF_Step
        '
        Me.TextF_Step.Location = New System.Drawing.Point(233, 38)
        Me.TextF_Step.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TextF_Step.Name = "TextF_Step"
        Me.TextF_Step.Size = New System.Drawing.Size(61, 23)
        Me.TextF_Step.TabIndex = 86
        Me.TextF_Step.Text = "0.1"
        Me.TextF_Step.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextF_To
        '
        Me.TextF_To.Location = New System.Drawing.Point(128, 38)
        Me.TextF_To.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TextF_To.Name = "TextF_To"
        Me.TextF_To.Size = New System.Drawing.Size(48, 23)
        Me.TextF_To.TabIndex = 85
        Me.TextF_To.Text = "1"
        Me.TextF_To.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ComboBox_Type_F
        '
        Me.ComboBox_Type_F.FormattingEnabled = True
        Me.ComboBox_Type_F.Location = New System.Drawing.Point(16, 10)
        Me.ComboBox_Type_F.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ComboBox_Type_F.Name = "ComboBox_Type_F"
        Me.ComboBox_Type_F.Size = New System.Drawing.Size(80, 23)
        Me.ComboBox_Type_F.TabIndex = 10
        Me.ComboBox_Type_F.TabStop = False
        '
        'TextF_From
        '
        Me.TextF_From.Location = New System.Drawing.Point(49, 38)
        Me.TextF_From.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TextF_From.Name = "TextF_From"
        Me.TextF_From.Size = New System.Drawing.Size(47, 23)
        Me.TextF_From.TabIndex = 84
        Me.TextF_From.Text = "0.1"
        Me.TextF_From.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Button_Extract)
        Me.TabPage1.Location = New System.Drawing.Point(4, 27)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(342, 152)
        Me.TabPage1.TabIndex = 2
        Me.TabPage1.Text = "Extract spectra"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Button_Extract
        '
        Me.Button_Extract.Location = New System.Drawing.Point(30, 22)
        Me.Button_Extract.Name = "Button_Extract"
        Me.Button_Extract.Size = New System.Drawing.Size(144, 44)
        Me.Button_Extract.TabIndex = 0
        Me.Button_Extract.Text = "Extract hdf5 for GUPIX"
        Me.Button_Extract.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(1064, 310)
        Me.Button6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(90, 44)
        Me.Button6.TabIndex = 84
        Me.Button6.Text = "Button6"
        Me.Button6.UseVisualStyleBackColor = True
        Me.Button6.Visible = False
        '
        'ToolTip1
        '
        Me.ToolTip1.AutomaticDelay = 400
        Me.ToolTip1.AutoPopDelay = 6000
        Me.ToolTip1.InitialDelay = 400
        Me.ToolTip1.ReshowDelay = 80
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(1059, 91)
        Me.Button7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(85, 45)
        Me.Button7.TabIndex = 86
        Me.Button7.Text = "Button7"
        Me.Button7.UseVisualStyleBackColor = True
        Me.Button7.Visible = False
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(1059, 377)
        Me.Button8.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(90, 44)
        Me.Button8.TabIndex = 87
        Me.Button8.Text = "Button8"
        Me.Button8.UseVisualStyleBackColor = True
        Me.Button8.Visible = False
        '
        'Box_txtFiltre
        '
        Me.Box_txtFiltre.AcceptsReturn = True
        Me.Box_txtFiltre.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Box_txtFiltre.Location = New System.Drawing.Point(567, 34)
        Me.Box_txtFiltre.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Box_txtFiltre.Name = "Box_txtFiltre"
        Me.Box_txtFiltre.Size = New System.Drawing.Size(112, 23)
        Me.Box_txtFiltre.TabIndex = 88
        Me.Box_txtFiltre.Text = "*"
        Me.Box_txtFiltre.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Label11.Location = New System.Drawing.Point(477, 38)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(86, 15)
        Me.Label11.TabIndex = 89
        Me.Label11.Text = "Filter filename"
        '
        'ListBox_HDF5
        '
        Me.ListBox_HDF5.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.ListBox_HDF5.FormattingEnabled = True
        Me.ListBox_HDF5.Location = New System.Drawing.Point(7, 64)
        Me.ListBox_HDF5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ListBox_HDF5.Name = "ListBox_HDF5"
        Me.ListBox_HDF5.Size = New System.Drawing.Size(301, 30)
        Me.ListBox_HDF5.TabIndex = 90
        '
        'TxtBox_HDF5_File
        '
        Me.TxtBox_HDF5_File.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.TxtBox_HDF5_File.Location = New System.Drawing.Point(317, 89)
        Me.TxtBox_HDF5_File.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TxtBox_HDF5_File.Name = "TxtBox_HDF5_File"
        Me.TxtBox_HDF5_File.Size = New System.Drawing.Size(367, 21)
        Me.TxtBox_HDF5_File.TabIndex = 91
        Me.TxtBox_HDF5_File.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightGray
        Me.Panel1.Controls.Add(Me.Pivot_det8)
        Me.Panel1.Controls.Add(Me.Pivot_det7)
        Me.Panel1.Controls.Add(Me.Pivot_det6)
        Me.Panel1.Controls.Add(Me.Pivot_det5)
        Me.Panel1.Controls.Add(Me.Pivot_det4)
        Me.Panel1.Controls.Add(Me.Pivot_det3)
        Me.Panel1.Controls.Add(Me.Pivot_det2)
        Me.Panel1.Controls.Add(Me.Pivot_det1)
        Me.Panel1.Controls.Add(Me.Pivot_det0)
        Me.Panel1.Controls.Add(Me.Par_det8)
        Me.Panel1.Controls.Add(Me.Check_det8)
        Me.Panel1.Controls.Add(Me.Check_det7)
        Me.Panel1.Controls.Add(Me.Check_det6)
        Me.Panel1.Controls.Add(Me.Par_det7)
        Me.Panel1.Controls.Add(Me.Par_det6)
        Me.Panel1.Controls.Add(Me.Par_det5)
        Me.Panel1.Controls.Add(Me.Check_det5)
        Me.Panel1.Controls.Add(Me.Check_det4)
        Me.Panel1.Controls.Add(Me.Check_det3)
        Me.Panel1.Controls.Add(Me.Par_det4)
        Me.Panel1.Controls.Add(Me.Par_det3)
        Me.Panel1.Controls.Add(Me.Par_det2)
        Me.Panel1.Controls.Add(Me.Check_det2)
        Me.Panel1.Controls.Add(Me.Check_det1)
        Me.Panel1.Controls.Add(Me.Check_det0)
        Me.Panel1.Controls.Add(Me.Par_det1)
        Me.Panel1.Controls.Add(Me.Par_det0)
        Me.Panel1.Controls.Add(Me.LstPar_Trc)
        Me.Panel1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Panel1.Location = New System.Drawing.Point(700, 140)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(350, 246)
        Me.Panel1.TabIndex = 92
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Gray
        Me.Panel2.Controls.Add(Me.Par_Mat)
        Me.Panel2.Controls.Add(Me.CbDetMat)
        Me.Panel2.Controls.Add(Me.LstPar_Mat)
        Me.Panel2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Panel2.Location = New System.Drawing.Point(700, 41)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(350, 75)
        Me.Panel2.TabIndex = 93
        '
        'Button_Run2
        '
        Me.Button_Run2.Font = New System.Drawing.Font("Calibri", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.Button_Run2.Location = New System.Drawing.Point(316, 64)
        Me.Button_Run2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Button_Run2.Name = "Button_Run2"
        Me.Button_Run2.Size = New System.Drawing.Size(77, 20)
        Me.Button_Run2.TabIndex = 94
        Me.Button_Run2.Text = "Run"
        Me.Button_Run2.UseVisualStyleBackColor = True
        '
        'TextProcessIf
        '
        Me.TextProcessIf.Location = New System.Drawing.Point(413, 64)
        Me.TextProcessIf.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextProcessIf.Name = "TextProcessIf"
        Me.TextProcessIf.Size = New System.Drawing.Size(98, 23)
        Me.TextProcessIf.TabIndex = 95
        Me.TextProcessIf.Visible = False
        '
        'Form_Traupixe_H5_2024
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1055, 583)
        Me.Controls.Add(Me.TextProcessIf)
        Me.Controls.Add(Me.Button_Run2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TxtBox_HDF5_File)
        Me.Controls.Add(Me.ListBox_HDF5)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Box_txtFiltre)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.ListFileInit)
        Me.Controls.Add(Me.LabelNew)
        Me.Controls.Add(Me.LabelAppend)
        Me.Controls.Add(Me.TextXLS)
        Me.Controls.Add(Me.Same_Z)
        Me.Controls.Add(Me.Text_Status)
        Me.Controls.Add(Me.ComboBox_Type_Calc)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.BtRefresh)
        Me.Controls.Add(Me.LvFiles)
        Me.Controls.Add(Me.trvFolders)
        Me.Controls.Add(Me.ComboBoxDrive)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.MaximizeBox = False
        Me.Name = "Form_Traupixe_H5_2024"
        Me.Text = "TrauPIXE 25.10 - C2RMF"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.Calcul.ResumeLayout(False)
        Me.Calcul.PerformLayout()
        Me.Tab_Adjust.ResumeLayout(False)
        Me.Tab_Adjust.PerformLayout()
        Me.TabPage1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LstPar_Mat As System.Windows.Forms.ListBox
    Friend WithEvents LstPar_Trc As System.Windows.Forms.ListBox
    Friend WithEvents ComboBoxDrive As System.Windows.Forms.ComboBox
    Friend WithEvents trvFolders As System.Windows.Forms.TreeView


    Private Sub Form_Main_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Dim drives As System.Collections.ObjectModel.ReadOnlyCollection(Of IO.DriveInfo) = My.Computer.FileSystem.Drives
        Dim rootDir As String = String.Empty
        Dim SR As StreamReader
        Dim SplitText() As String

        ToolTip1.SetToolTip(Text_Lst_Ox_Trc, "Set Z as oxyde only for 'S Sheets : S_Conc 100% & S_Conc ppm, '")
        ToolTip1.SetToolTip(Pivot_det0, "Z of the PIVOT element, Single or multiple separate Z with ',' (e.g.: 26,20)" & vbCrLf & "Select the one with lower Total error (Quadratic sum of Fit & Pivot Error)")
        ToolTip1.SetToolTip(Pivot_det1, "Z of the PIVOT element, Single or multiple separate Z with ',' (e.g.: 26,20)" & vbCrLf & "Select the one with lower Total error (Quadratic sum of Fit & Pivot Error)")
        ToolTip1.SetToolTip(Pivot_det2, "Z of the PIVOT element, Single or multiple separate Z with ',' (e.g.: 26,20)" & vbCrLf & "Select the one with lower Total error (Quadratic sum of Fit & Pivot Error)")
        ToolTip1.SetToolTip(Pivot_det3, "Z of the PIVOT element, Single or multiple separate Z with ',' (e.g.: 26,20)" & vbCrLf & "Select the one with lower Total error (Quadratic sum of Fit & Pivot Error)")
        ToolTip1.SetToolTip(Pivot_det4, "Z of the PIVOT element, Single or multiple separate Z with ',' (e.g.: 26,20)" & vbCrLf & "Select the one with lower Total error (Quadratic sum of Fit & Pivot Error)")
        ToolTip1.SetToolTip(Pivot_det6, "Z of the PIVOT element, Single or multiple separate Z with ',' (e.g.: 26,20)" & vbCrLf & "Select the one with lower Total error (Quadratic sum of Fit & Pivot Error)")
        ToolTip1.SetToolTip(Pivot_det7, "Z of the PIVOT element, Single or multiple separate Z with ',' (e.g.: 26,20)" & vbCrLf & "Select the one with lower Total error (Quadratic sum of Fit & Pivot Error)")
        ToolTip1.SetToolTip(Pivot_det8, "Z of the PIVOT element, Single or multiple separate Z with ',' (e.g.: 26,20)" & vbCrLf & "Select the one with lower Total error (Quadratic sum of Fit & Pivot Error)")
        ToolTip1.SetToolTip(Pivot_det5, "Z of the PIVOT element, Single or multiple separate Z with ',' (e.g.: 26,20)" & vbCrLf & "Select the one with lower Total error (Quadratic sum of Fit & Pivot Error)")

        ' ToolTip1.SetToolTip(txtAge, "Enter Your age")
        'ToolTip1.SetToolTip(btnSave, "Save record")


        For Each Drive_Name In DriveInfo.GetDrives()
            'For Each disque As String In Directory.GetLogicalDrives()
            If Drive_Name.IsReady = True Then
                Try
                    ComboBoxDrive.Items.Add(Drive_Name.Name & Drive_Name.VolumeLabel)
                Catch ex As Exception
                    ComboBoxDrive.Items.Add(Drive_Name.Name)
                End Try
                If Drive_Name.Name = "C:\" Then
                    ComboBoxDrive.SelectedIndex = 0
                    Myinit = False
                    ' Exit For
                End If

            End If

            ComboBoxDrive.SelectedIndex = 0
            Myinit = False
            ' ComboBoxDrive.Items.Add(disque)
        Next


        Try

            Me.Chemin_GupixWin = "c:\gupixwin\gupix"
            SR = File.OpenText(Environment.CurrentDirectory & "\Config_Traupixe.ini")
            Dim MyConfig = SR.ReadToEnd
            SplitText = Split(MyConfig, vbCrLf)
            Dim ind1 = 0
            Dim str_hdf5 As String
            Dim h5_analyses_location As String

            Me.chb_skip_elem.Checked = False
            Me.chb_skip_elem.Visible = False
            Me.B_skip_elem_mtx = False

            For Each Str As String In SplitText
                str_hdf5 = Str
                Select Case Str

                    Case "[GUPIX PATH]"
                        Me.Chemin_GupixWin = SplitText(ind1 + 1)
                    Case "[DET0]"
                        Me.Ext_Mat = SplitText(ind1 + 1)
                        Me.CbDetMat.Items.Add(SplitText(ind1 + 1))
                        Me.Check_det0.Text = SplitText(ind1 + 1)
                        Me.Ext_Trc0 = SplitText(ind1 + 1)
                    Case "[DET1]"
                        Me.Ext_Trc1 = SplitText(ind1 + 1)
                        Me.CbDetMat.Items.Add(Ext_Trc1)
                        Me.Check_det1.Text = SplitText(ind1 + 1)
                        'Check_BE0.Text = Me.Ext_Trc0
                    Case "[DET2]"
                        Me.Ext_Trc2 = SplitText(ind1 + 1)
                        Me.CbDetMat.Items.Add(Ext_Trc2)
                        Me.Check_det2.Text = SplitText(ind1 + 1)
                    Case "[DET3]"
                        Me.Ext_Trc3 = SplitText(ind1 + 1)
                        Me.CbDetMat.Items.Add(Ext_Trc3)
                        Me.Check_det3.Text = SplitText(ind1 + 1)
                    Case "[DET4]"
                        Me.Ext_Trc4 = SplitText(ind1 + 1)
                        Me.CbDetMat.Items.Add(Ext_Trc4)
                        Me.Check_det4.Text = SplitText(ind1 + 1)
                    Case "[DET5]"
                        Me.Ext_Trc5 = SplitText(ind1 + 1)
                        Me.CbDetMat.Items.Add(Ext_Trc5)
                        Me.Check_det5.Text = SplitText(ind1 + 1)
                    Case "[DET6]"
                        Me.Ext_Trc6 = SplitText(ind1 + 1)
                        Me.CbDetMat.Items.Add(Ext_Trc6)
                        Me.Check_det6.Text = SplitText(ind1 + 1)
                    Case "[DET7]"
                        Me.Ext_Trc7 = SplitText(ind1 + 1)
                        Me.CbDetMat.Items.Add(Ext_Trc7)
                        Me.Check_det7.Text = SplitText(ind1 + 1)
                    Case "[DET8]"
                        Me.Ext_Trc8 = SplitText(ind1 + 1)
                        Me.CbDetMat.Items.Add(Ext_Trc8)
                        Me.Check_det8.Text = SplitText(ind1 + 1)
                    Case "[HDF5]"

                        '                       [analyses_location]
                        '/
                        '[analyses_attributes]
                        '                       ref Object
                        '   start Date
                        '[analyses_dataset-name]
                        '"x0,x1,x2,x3,x4,x10,x11,x12,x13"
                        '[dataset_attributes]
                        '                       acquisition time
                        '   experiment Information
                        '   Month()
                        '                       seconds since midnight
                        '   spectrum sum
                        '   user comment
                        '   Year()

                        'ind1 += 1
                        'Do
                        '    str_hdf5 = SplitText(ind1 + 1)
                        '    Select Case str_hdf5
                        '        Case "[analyses_location]"
                        '            h5_analyses_location = str_hdf5 = SplitText(ind1 + 2)
                        '        Case "[analyses_attributes]"


                        '    End Select

                        '    ind1 += 1
                        'Loop While str_hdf5 <> "[\HDF5]"
                End Select
                ind1 += 1
            Next

        Catch ex As Exception
            Me.Chemin_GupixWin = "c:\gupixwin\gupix"
            Me.Ext_Mat = "X0"
            Me.CbDetMat.Items.Add("X0")
            Me.Check_det0.Text = "X0"
            Me.Ext_Trc0 = "X0"

            Me.Ext_Trc1 = "X1"
            Me.CbDetMat.Items.Add("X1")
            Me.Check_det1.Text = "X1"

            Me.Ext_Trc2 = "X2"
            Me.CbDetMat.Items.Add("X2")
            Me.Check_det2.Text = "X2"

            Me.Ext_Trc3 = "X3"
            Me.CbDetMat.Items.Add("X3")
            Me.Check_det3.Text = "X3"

            Me.Ext_Trc4 = "X4"
            Me.CbDetMat.Items.Add("X4")
            Me.Check_det4.Text = "X4"

            Me.Ext_Trc5 = "X10"
            Me.CbDetMat.Items.Add("X10")
            Me.Check_det5.Text = "X10"

            Me.Ext_Trc6 = "X11"
            Me.CbDetMat.Items.Add("X12")
            Me.Check_det6.Text = "X12"

            Me.Ext_Trc7 = "X12"
            Me.CbDetMat.Items.Add("X12")
            Me.Check_det7.Text = "X12"

            Me.Ext_Trc8 = "X13"
            Me.CbDetMat.Items.Add("X13")
            Me.Check_det8.Text = "X13"
            Me.Ext_Trc8 = "X13"
            Me.CbDetMat.Items.Add("X1")
            Me.Check_det1.Text = "X1"

        End Try
        Me.CbDetMat.Text = Me.Ext_Mat
        Dim factor
        Dim screenwidth As Integer = My.Computer.Screen.Bounds.Width
        'screenwidth on my computer is 1600. The controls have been designed for that width.
        factor = screenwidth / 1600 'this is screenwidth of computer where the program has been installed.
        If factor < 1 Then
            Button_Run2.Visible = True
        Else
            Button_Run2.Visible = False
        End If


    End Sub

    Friend WithEvents LvFiles As System.Windows.Forms.ListView
    Friend WithEvents BtRefresh As System.Windows.Forms.Button
    Friend WithEvents CbDetMat As System.Windows.Forms.ComboBox
    Friend WithEvents Par_det0 As System.Windows.Forms.TextBox
    Friend WithEvents Par_det1 As System.Windows.Forms.TextBox
    Friend WithEvents Check_det0 As System.Windows.Forms.CheckBox
    Friend WithEvents Check_det1 As System.Windows.Forms.CheckBox
    Friend WithEvents Check_det2 As System.Windows.Forms.CheckBox
    Friend WithEvents Par_det2 As System.Windows.Forms.TextBox
    Friend WithEvents Par_det5 As System.Windows.Forms.TextBox
    Friend WithEvents Check_det5 As System.Windows.Forms.CheckBox
    Friend WithEvents Check_det4 As System.Windows.Forms.CheckBox
    Friend WithEvents Check_det3 As System.Windows.Forms.CheckBox
    Friend WithEvents Par_det4 As System.Windows.Forms.TextBox
    Friend WithEvents Par_det3 As System.Windows.Forms.TextBox
    Friend WithEvents Par_det8 As System.Windows.Forms.TextBox
    Friend WithEvents Check_det8 As System.Windows.Forms.CheckBox
    Friend WithEvents Check_det7 As System.Windows.Forms.CheckBox
    Friend WithEvents Check_det6 As System.Windows.Forms.CheckBox
    Friend WithEvents Par_det7 As System.Windows.Forms.TextBox
    Friend WithEvents Par_det6 As System.Windows.Forms.TextBox
    Friend WithEvents Par_Mat As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Nb_Proc As System.Windows.Forms.TextBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ComboBox_Type_Calc As System.Windows.Forms.ComboBox
    Friend WithEvents Progress As System.Windows.Forms.TextBox
    Friend WithEvents Tps_Calc As System.Windows.Forms.TextBox
    Friend WithEvents MyPause As System.Windows.Forms.Button
    Friend WithEvents Text_Status As System.Windows.Forms.TextBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuOxyde As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOxydeOUI As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOxydeNON As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LODToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GupixLODNWrite0ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents WriteAllValuesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Pivot_det0 As System.Windows.Forms.TextBox
    Friend WithEvents Pivot_det1 As System.Windows.Forms.TextBox
    Friend WithEvents Pivot_det2 As System.Windows.Forms.TextBox
    Friend WithEvents Pivot_det5 As System.Windows.Forms.TextBox
    Friend WithEvents Pivot_det4 As System.Windows.Forms.TextBox
    Friend WithEvents Pivot_det3 As System.Windows.Forms.TextBox
    Friend WithEvents Pivot_det8 As System.Windows.Forms.TextBox
    Friend WithEvents Pivot_det7 As System.Windows.Forms.TextBox
    Friend WithEvents Pivot_det6 As System.Windows.Forms.TextBox
    Friend WithEvents Same_Z As System.Windows.Forms.CheckBox
    Friend WithEvents TextXLS As System.Windows.Forms.TextBox
    Friend WithEvents LabelAppend As System.Windows.Forms.Label
    Friend WithEvents LabelNew As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Check_Trc_As_Oxy As System.Windows.Forms.CheckBox
    Friend WithEvents Text_Lst_Ox_Trc As System.Windows.Forms.TextBox
    Friend WithEvents Ck_AllAsOxy As System.Windows.Forms.CheckBox
    Friend WithEvents ListFileInit As System.Windows.Forms.ListView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents RoundConcentrationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Chk_RoundValue As System.Windows.Forms.CheckBox
    Friend WithEvents Adjust_Filter As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents Calcul As System.Windows.Forms.TabPage
    Friend WithEvents Tab_Adjust As System.Windows.Forms.TabPage
    Friend WithEvents TextF_From As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox_Type_F As System.Windows.Forms.ComboBox
    Friend WithEvents TextF_Step As System.Windows.Forms.TextBox
    Friend WithEvents TextF_To As System.Windows.Forms.TextBox
    Friend WithEvents TextF_Z As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Box_txtFiltre As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ListBox_HDF5 As ListBox
    Friend WithEvents TxtBox_HDF5_File As TextBox
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents Button_Extract As Button
    Friend WithEvents Text_gamma As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Button_Run2 As Button
    Friend WithEvents TextProcessIf As TextBox
    Public WithEvents SkipPbMatrixToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents chk_external_ok As CheckBox
    Friend WithEvents chb_skip_elem As CheckBox
    Friend WithEvents txt_skip_elem As TextBox
End Class
