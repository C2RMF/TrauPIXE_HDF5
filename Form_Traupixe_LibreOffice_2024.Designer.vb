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
        components = New ComponentModel.Container()
        LstPar_Mat = New ListBox()
        LstPar_Trc = New ListBox()
        ComboBoxDrive = New ComboBox()
        trvFolders = New TreeView()
        LvFiles = New ListView()
        BtRefresh = New Button()
        CbDetMat = New ComboBox()
        Par_det0 = New TextBox()
        Par_det1 = New TextBox()
        Check_det0 = New CheckBox()
        Check_det1 = New CheckBox()
        Check_det2 = New CheckBox()
        Par_det2 = New TextBox()
        Par_det5 = New TextBox()
        Check_det5 = New CheckBox()
        Check_det4 = New CheckBox()
        Check_det3 = New CheckBox()
        Par_det4 = New TextBox()
        Par_det3 = New TextBox()
        Par_det8 = New TextBox()
        Check_det8 = New CheckBox()
        Check_det7 = New CheckBox()
        Check_det6 = New CheckBox()
        Par_det7 = New TextBox()
        Par_det6 = New TextBox()
        Par_Mat = New TextBox()
        Button2 = New Button()
        Label1 = New Label()
        Label2 = New Label()
        Nb_Proc = New TextBox()
        ProgressBar1 = New ProgressBar()
        ComboBox_Type_Calc = New ComboBox()
        Progress = New TextBox()
        Tps_Calc = New TextBox()
        MyPause = New Button()
        Text_Status = New TextBox()
        MenuStrip1 = New MenuStrip()
        mnuOxyde = New ToolStripMenuItem()
        mnuOxydeOUI = New ToolStripMenuItem()
        mnuOxydeNON = New ToolStripMenuItem()
        LODToolStripMenuItem = New ToolStripMenuItem()
        GupixLODNWrite0ToolStripMenuItem = New ToolStripMenuItem()
        WriteAllValuesToolStripMenuItem = New ToolStripMenuItem()
        RoundConcentrationToolStripMenuItem = New ToolStripMenuItem()
        SkipPbMatrixToolStripMenuItem = New ToolStripMenuItem()
        Pivot_det0 = New TextBox()
        Pivot_det1 = New TextBox()
        Pivot_det2 = New TextBox()
        Pivot_det5 = New TextBox()
        Pivot_det4 = New TextBox()
        Pivot_det3 = New TextBox()
        Pivot_det8 = New TextBox()
        Pivot_det7 = New TextBox()
        Pivot_det6 = New TextBox()
        Same_Z = New CheckBox()
        TextXLS = New TextBox()
        LabelAppend = New Label()
        LabelNew = New Label()
        Timer1 = New Timer(components)
        Check_Trc_As_Oxy = New CheckBox()
        Text_Lst_Ox_Trc = New TextBox()
        Ck_AllAsOxy = New CheckBox()
        ListFileInit = New ListView()
        StatusStrip1 = New StatusStrip()
        ToolStripStatusLabel2 = New ToolStripStatusLabel()
        ToolStripStatusLabel1 = New ToolStripStatusLabel()
        Chk_RoundValue = New CheckBox()
        Adjust_Filter = New Button()
        TabControl1 = New TabControl()
        Calcul = New TabPage()
        chb_skip_elem = New CheckBox()
        txt_skip_elem = New TextBox()
        chk_external_ok = New CheckBox()
        Text_gamma = New TextBox()
        Tab_Adjust = New TabPage()
        Label10 = New Label()
        Label9 = New Label()
        Label8 = New Label()
        Label7 = New Label()
        Label6 = New Label()
        Label5 = New Label()
        TextF_Z = New TextBox()
        TextF_Step = New TextBox()
        TextF_To = New TextBox()
        ComboBox_Type_F = New ComboBox()
        TextF_From = New TextBox()
        TabPage1 = New TabPage()
        Button_Extract = New Button()
        Button6 = New Button()
        ToolTip1 = New ToolTip(components)
        Button7 = New Button()
        Button8 = New Button()
        Box_txtFiltre = New TextBox()
        Label11 = New Label()
        ListBox_HDF5 = New ListBox()
        TxtBox_HDF5_File = New TextBox()
        Panel1 = New Panel()
        Panel2 = New Panel()
        Button_Run2 = New Button()
        TextProcessIf = New TextBox()
        Button1 = New Button()
        Label_QFile = New Label()
        Label3 = New Label()
        Label4 = New Label()
        TextBox_hdf5_grps = New TextBox()
        Label12 = New Label()
        MenuStrip1.SuspendLayout()
        StatusStrip1.SuspendLayout()
        TabControl1.SuspendLayout()
        Calcul.SuspendLayout()
        Tab_Adjust.SuspendLayout()
        TabPage1.SuspendLayout()
        Panel1.SuspendLayout()
        Panel2.SuspendLayout()
        SuspendLayout()
        ' 
        ' LstPar_Mat
        ' 
        LstPar_Mat.Font = New Font("Calibri", 8F)
        LstPar_Mat.FormattingEnabled = True
        LstPar_Mat.Location = New Point(8, 28)
        LstPar_Mat.Margin = New Padding(5, 4, 5, 4)
        LstPar_Mat.Name = "LstPar_Mat"
        LstPar_Mat.Size = New Size(377, 30)
        LstPar_Mat.TabIndex = 7
        ' 
        ' LstPar_Trc
        ' 
        LstPar_Trc.Font = New Font("Calibri", 8F)
        LstPar_Trc.FormattingEnabled = True
        LstPar_Trc.Location = New Point(8, 235)
        LstPar_Trc.Margin = New Padding(5, 4, 5, 4)
        LstPar_Trc.Name = "LstPar_Trc"
        LstPar_Trc.SelectionMode = SelectionMode.MultiExtended
        LstPar_Trc.Size = New Size(383, 43)
        LstPar_Trc.TabIndex = 9
        ' 
        ' ComboBoxDrive
        ' 
        ComboBoxDrive.Font = New Font("Calibri", 8.25F)
        ComboBoxDrive.FormattingEnabled = True
        ComboBoxDrive.Location = New Point(11, 28)
        ComboBoxDrive.Margin = New Padding(5, 4, 5, 4)
        ComboBoxDrive.Name = "ComboBoxDrive"
        ComboBoxDrive.Size = New Size(194, 21)
        ComboBoxDrive.TabIndex = 10
        ' 
        ' trvFolders
        ' 
        trvFolders.Font = New Font("Calibri", 8.25F)
        trvFolders.Location = New Point(11, 58)
        trvFolders.Margin = New Padding(5, 4, 5, 4)
        trvFolders.Name = "trvFolders"
        trvFolders.ShowPlusMinus = False
        trvFolders.ShowRootLines = False
        trvFolders.Size = New Size(339, 462)
        trvFolders.TabIndex = 11
        ' 
        ' LvFiles
        ' 
        LvFiles.Font = New Font("Calibri", 8.25F)
        LvFiles.GridLines = True
        LvFiles.Location = New Point(361, 58)
        LvFiles.Margin = New Padding(5, 4, 5, 4)
        LvFiles.Name = "LvFiles"
        LvFiles.Size = New Size(420, 356)
        LvFiles.Sorting = SortOrder.Ascending
        LvFiles.TabIndex = 12
        LvFiles.UseCompatibleStateImageBehavior = False
        LvFiles.View = View.List
        ' 
        ' BtRefresh
        ' 
        BtRefresh.Font = New Font("Calibri", 8F)
        BtRefresh.Location = New Point(361, 25)
        BtRefresh.Margin = New Padding(5, 4, 5, 4)
        BtRefresh.Name = "BtRefresh"
        BtRefresh.Size = New Size(119, 29)
        BtRefresh.TabIndex = 13
        BtRefresh.Text = "Refresh spectra"
        BtRefresh.UseVisualStyleBackColor = True
        ' 
        ' CbDetMat
        ' 
        CbDetMat.Font = New Font("Calibri", 8F)
        CbDetMat.FormattingEnabled = True
        CbDetMat.Location = New Point(8, 4)
        CbDetMat.Margin = New Padding(5, 4, 5, 4)
        CbDetMat.Name = "CbDetMat"
        CbDetMat.Size = New Size(67, 21)
        CbDetMat.TabIndex = 14
        CbDetMat.Text = "X0"
        ' 
        ' Par_det0
        ' 
        Par_det0.Font = New Font("Calibri", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Par_det0.Location = New Point(82, 3)
        Par_det0.Margin = New Padding(5, 4, 5, 4)
        Par_det0.Name = "Par_det0"
        Par_det0.Size = New Size(215, 21)
        Par_det0.TabIndex = 16
        ' 
        ' Par_det1
        ' 
        Par_det1.Font = New Font("Calibri", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Par_det1.Location = New Point(82, 28)
        Par_det1.Margin = New Padding(5, 4, 5, 4)
        Par_det1.Name = "Par_det1"
        Par_det1.Size = New Size(215, 21)
        Par_det1.TabIndex = 17
        ' 
        ' Check_det0
        ' 
        Check_det0.AutoSize = True
        Check_det0.Font = New Font("Calibri", 9.75F)
        Check_det0.Location = New Point(10, 5)
        Check_det0.Margin = New Padding(5, 4, 5, 4)
        Check_det0.Name = "Check_det0"
        Check_det0.Size = New Size(46, 19)
        Check_det0.TabIndex = 22
        Check_det0.Text = "BE0"
        Check_det0.UseVisualStyleBackColor = True
        ' 
        ' Check_det1
        ' 
        Check_det1.AutoSize = True
        Check_det1.Location = New Point(10, 30)
        Check_det1.Margin = New Padding(5, 4, 5, 4)
        Check_det1.Name = "Check_det1"
        Check_det1.Size = New Size(47, 19)
        Check_det1.TabIndex = 23
        Check_det1.Text = "HE1"
        Check_det1.UseVisualStyleBackColor = True
        ' 
        ' Check_det2
        ' 
        Check_det2.AutoSize = True
        Check_det2.Location = New Point(10, 55)
        Check_det2.Margin = New Padding(5, 4, 5, 4)
        Check_det2.Name = "Check_det2"
        Check_det2.Size = New Size(47, 19)
        Check_det2.TabIndex = 24
        Check_det2.Text = "HE2"
        Check_det2.UseVisualStyleBackColor = True
        ' 
        ' Par_det2
        ' 
        Par_det2.Font = New Font("Calibri", 8F)
        Par_det2.Location = New Point(82, 53)
        Par_det2.Margin = New Padding(5, 4, 5, 4)
        Par_det2.Name = "Par_det2"
        Par_det2.Size = New Size(215, 21)
        Par_det2.TabIndex = 25
        ' 
        ' Par_det5
        ' 
        Par_det5.Font = New Font("Calibri", 8F)
        Par_det5.Location = New Point(82, 128)
        Par_det5.Margin = New Padding(5, 4, 5, 4)
        Par_det5.Name = "Par_det5"
        Par_det5.Size = New Size(215, 21)
        Par_det5.TabIndex = 31
        ' 
        ' Check_det5
        ' 
        Check_det5.AutoSize = True
        Check_det5.BackColor = Color.LightGray
        Check_det5.Location = New Point(10, 130)
        Check_det5.Margin = New Padding(5, 4, 5, 4)
        Check_det5.Name = "Check_det5"
        Check_det5.Size = New Size(54, 19)
        Check_det5.TabIndex = 30
        Check_det5.Text = "HE10"
        Check_det5.UseVisualStyleBackColor = False
        ' 
        ' Check_det4
        ' 
        Check_det4.AutoSize = True
        Check_det4.Location = New Point(10, 105)
        Check_det4.Margin = New Padding(5, 4, 5, 4)
        Check_det4.Name = "Check_det4"
        Check_det4.Size = New Size(47, 19)
        Check_det4.TabIndex = 29
        Check_det4.Text = "HE4"
        Check_det4.UseVisualStyleBackColor = True
        ' 
        ' Check_det3
        ' 
        Check_det3.AutoSize = True
        Check_det3.Location = New Point(10, 80)
        Check_det3.Margin = New Padding(5, 4, 5, 4)
        Check_det3.Name = "Check_det3"
        Check_det3.Size = New Size(47, 19)
        Check_det3.TabIndex = 28
        Check_det3.Text = "HE3"
        Check_det3.UseVisualStyleBackColor = True
        ' 
        ' Par_det4
        ' 
        Par_det4.Font = New Font("Calibri", 8F)
        Par_det4.Location = New Point(82, 103)
        Par_det4.Margin = New Padding(5, 4, 5, 4)
        Par_det4.Name = "Par_det4"
        Par_det4.Size = New Size(215, 21)
        Par_det4.TabIndex = 27
        ' 
        ' Par_det3
        ' 
        Par_det3.Font = New Font("Calibri", 8F)
        Par_det3.Location = New Point(82, 77)
        Par_det3.Margin = New Padding(5, 4, 5, 4)
        Par_det3.Name = "Par_det3"
        Par_det3.Size = New Size(215, 21)
        Par_det3.TabIndex = 26
        ' 
        ' Par_det8
        ' 
        Par_det8.Font = New Font("Calibri", 8F)
        Par_det8.Location = New Point(82, 204)
        Par_det8.Margin = New Padding(5, 4, 5, 4)
        Par_det8.Name = "Par_det8"
        Par_det8.Size = New Size(215, 21)
        Par_det8.TabIndex = 37
        ' 
        ' Check_det8
        ' 
        Check_det8.AutoSize = True
        Check_det8.Location = New Point(10, 205)
        Check_det8.Margin = New Padding(5, 4, 5, 4)
        Check_det8.Name = "Check_det8"
        Check_det8.Size = New Size(54, 19)
        Check_det8.TabIndex = 36
        Check_det8.Text = "HE13"
        Check_det8.UseVisualStyleBackColor = True
        ' 
        ' Check_det7
        ' 
        Check_det7.AutoSize = True
        Check_det7.Location = New Point(10, 180)
        Check_det7.Margin = New Padding(5, 4, 5, 4)
        Check_det7.Name = "Check_det7"
        Check_det7.Size = New Size(54, 19)
        Check_det7.TabIndex = 35
        Check_det7.Text = "HE12"
        Check_det7.UseVisualStyleBackColor = True
        ' 
        ' Check_det6
        ' 
        Check_det6.AutoSize = True
        Check_det6.Location = New Point(10, 155)
        Check_det6.Margin = New Padding(5, 4, 5, 4)
        Check_det6.Name = "Check_det6"
        Check_det6.Size = New Size(54, 19)
        Check_det6.TabIndex = 34
        Check_det6.Text = "HE11"
        Check_det6.UseVisualStyleBackColor = True
        ' 
        ' Par_det7
        ' 
        Par_det7.Font = New Font("Calibri", 8F)
        Par_det7.Location = New Point(82, 179)
        Par_det7.Margin = New Padding(5, 4, 5, 4)
        Par_det7.Name = "Par_det7"
        Par_det7.Size = New Size(215, 21)
        Par_det7.TabIndex = 33
        ' 
        ' Par_det6
        ' 
        Par_det6.Font = New Font("Calibri", 8F)
        Par_det6.Location = New Point(82, 152)
        Par_det6.Margin = New Padding(5, 4, 5, 4)
        Par_det6.Name = "Par_det6"
        Par_det6.Size = New Size(215, 21)
        Par_det6.TabIndex = 32
        ' 
        ' Par_Mat
        ' 
        Par_Mat.Font = New Font("Calibri", 8F)
        Par_Mat.Location = New Point(86, 4)
        Par_Mat.Margin = New Padding(5, 4, 5, 4)
        Par_Mat.Name = "Par_Mat"
        Par_Mat.Size = New Size(298, 21)
        Par_Mat.TabIndex = 38
        Par_Mat.TextAlign = HorizontalAlignment.Center
        ' 
        ' Button2
        ' 
        Button2.Font = New Font("Calibri", 13F)
        Button2.Location = New Point(414, 5)
        Button2.Margin = New Padding(5, 4, 5, 4)
        Button2.Name = "Button2"
        Button2.Size = New Size(117, 37)
        Button2.TabIndex = 39
        Button2.Text = "Run"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.BackColor = Color.Gray
        Label1.Font = New Font("Segoe UI", 8F, FontStyle.Bold)
        Label1.ForeColor = Color.White
        Label1.Location = New Point(800, 26)
        Label1.Margin = New Padding(5, 0, 5, 0)
        Label1.Name = "Label1"
        Label1.Size = New Size(41, 13)
        Label1.TabIndex = 40
        Label1.Text = "Matrix"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.BackColor = Color.LightGray
        Label2.Font = New Font("Segoe UI", 8F, FontStyle.Bold)
        Label2.Location = New Point(802, 113)
        Label2.Margin = New Padding(5, 0, 5, 0)
        Label2.Name = "Label2"
        Label2.Size = New Size(33, 13)
        Label2.TabIndex = 41
        Label2.Text = "Trace"
        ' 
        ' Nb_Proc
        ' 
        Nb_Proc.Font = New Font("Calibri", 8.25F)
        Nb_Proc.Location = New Point(805, 45)
        Nb_Proc.Margin = New Padding(5, 4, 5, 4)
        Nb_Proc.Name = "Nb_Proc"
        Nb_Proc.Size = New Size(22, 21)
        Nb_Proc.TabIndex = 43
        Nb_Proc.Text = "4"
        Nb_Proc.TextAlign = HorizontalAlignment.Center
        ' 
        ' ProgressBar1
        ' 
        ProgressBar1.Location = New Point(532, 524)
        ProgressBar1.Margin = New Padding(5, 4, 5, 4)
        ProgressBar1.Name = "ProgressBar1"
        ProgressBar1.Size = New Size(668, 24)
        ProgressBar1.TabIndex = 44
        ' 
        ' ComboBox_Type_Calc
        ' 
        ComboBox_Type_Calc.Enabled = False
        ComboBox_Type_Calc.Font = New Font("Calibri", 8.0F)
        ComboBox_Type_Calc.FormattingEnabled = True
        ComboBox_Type_Calc.Items.AddRange(New Object() {"Ponctual hdf5 ", "Map hdf5 "})
        ComboBox_Type_Calc.Location = New Point(231, 342)
        ComboBox_Type_Calc.Margin = New Padding(5, 4, 5, 4)
        ComboBox_Type_Calc.Name = "ComboBox_Type_Calc"
        ComboBox_Type_Calc.Size = New Size(118, 21)
        ComboBox_Type_Calc.TabIndex = 45
        ComboBox_Type_Calc.Text = "Ponctual hdf5 "
        ComboBox_Type_Calc.Visible = False
        ' 
        ' Progress
        ' 
        Progress.Enabled = False
        Progress.Font = New Font("Segoe UI", 12.0F)
        Progress.Location = New Point(555, 11)
        Progress.Margin = New Padding(5, 4, 5, 4)
        Progress.Name = "Progress"
        Progress.Size = New Size(84, 29)
        Progress.TabIndex = 46
        Progress.Text = "0 / 0"
        Progress.TextAlign = HorizontalAlignment.Center
        ' 
        ' Tps_Calc
        ' 
        Tps_Calc.Location = New Point(547, 45)
        Tps_Calc.Margin = New Padding(5, 4, 5, 4)
        Tps_Calc.Name = "Tps_Calc"
        Tps_Calc.Size = New Size(100, 21)
        Tps_Calc.TabIndex = 47
        Tps_Calc.Text = "??:??:??"
        Tps_Calc.TextAlign = HorizontalAlignment.Center
        Tps_Calc.Visible = False
        ' 
        ' MyPause
        ' 
        MyPause.Font = New Font("Calibri", 13.0F)
        MyPause.Location = New Point(666, 5)
        MyPause.Margin = New Padding(5, 4, 5, 4)
        MyPause.Name = "MyPause"
        MyPause.Size = New Size(117, 37)
        MyPause.TabIndex = 48
        MyPause.Text = "Pause"
        MyPause.UseVisualStyleBackColor = True
        ' 
        ' Text_Status
        ' 
        Text_Status.Location = New Point(133, 844)
        Text_Status.Margin = New Padding(5, 4, 5, 4)
        Text_Status.Name = "Text_Status"
        Text_Status.Size = New Size(660, 21)
        Text_Status.TabIndex = 49
        Text_Status.Text = "??:??:??"
        Text_Status.Visible = False
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Font = New Font("Segoe UI", 8.0F)
        MenuStrip1.ImageScalingSize = New Size(24, 24)
        MenuStrip1.Items.AddRange(New ToolStripItem() {mnuOxyde, LODToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Padding = New Padding(8, 3, 0, 3)
        MenuStrip1.Size = New Size(1320, 24)
        MenuStrip1.TabIndex = 50
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' mnuOxyde
        ' 
        mnuOxyde.DropDownItems.AddRange(New ToolStripItem() {mnuOxydeOUI, mnuOxydeNON})
        mnuOxyde.Name = "mnuOxyde"
        mnuOxyde.Size = New Size(49, 18)
        mnuOxyde.Text = "Oxide"
        ' 
        ' mnuOxydeOUI
        ' 
        mnuOxydeOUI.Checked = True
        mnuOxydeOUI.CheckState = CheckState.Checked
        mnuOxydeOUI.Name = "mnuOxydeOUI"
        mnuOxydeOUI.Size = New Size(89, 22)
        mnuOxydeOUI.Text = "Yes"
        ' 
        ' mnuOxydeNON
        ' 
        mnuOxydeNON.Name = "mnuOxydeNON"
        mnuOxydeNON.Size = New Size(89, 22)
        mnuOxydeNON.Text = "No"
        ' 
        ' LODToolStripMenuItem
        ' 
        LODToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {GupixLODNWrite0ToolStripMenuItem, WriteAllValuesToolStripMenuItem, RoundConcentrationToolStripMenuItem, SkipPbMatrixToolStripMenuItem})
        LODToolStripMenuItem.Name = "LODToolStripMenuItem"
        LODToolStripMenuItem.Size = New Size(49, 18)
        LODToolStripMenuItem.Text = "Setup"
        ' 
        ' GupixLODNWrite0ToolStripMenuItem
        ' 
        GupixLODNWrite0ToolStripMenuItem.Name = "GupixLODNWrite0ToolStripMenuItem"
        GupixLODNWrite0ToolStripMenuItem.Size = New Size(189, 22)
        GupixLODNWrite0ToolStripMenuItem.Text = "Gupix LOD = N write 0"
        ' 
        ' WriteAllValuesToolStripMenuItem
        ' 
        WriteAllValuesToolStripMenuItem.Checked = True
        WriteAllValuesToolStripMenuItem.CheckState = CheckState.Checked
        WriteAllValuesToolStripMenuItem.Name = "WriteAllValuesToolStripMenuItem"
        WriteAllValuesToolStripMenuItem.Size = New Size(189, 22)
        WriteAllValuesToolStripMenuItem.Text = "Write all values"
        ' 
        ' RoundConcentrationToolStripMenuItem
        ' 
        RoundConcentrationToolStripMenuItem.Checked = True
        RoundConcentrationToolStripMenuItem.CheckState = CheckState.Checked
        RoundConcentrationToolStripMenuItem.Name = "RoundConcentrationToolStripMenuItem"
        RoundConcentrationToolStripMenuItem.Size = New Size(189, 22)
        RoundConcentrationToolStripMenuItem.Text = "Round Concentration"
        ' 
        ' SkipPbMatrixToolStripMenuItem
        ' 
        SkipPbMatrixToolStripMenuItem.Name = "SkipPbMatrixToolStripMenuItem"
        SkipPbMatrixToolStripMenuItem.Size = New Size(189, 22)
        SkipPbMatrixToolStripMenuItem.Text = "Skip element in matrix"
        ' 
        ' Pivot_det0
        ' 
        Pivot_det0.Font = New Font("Calibri", 8.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Pivot_det0.Location = New Point(306, 3)
        Pivot_det0.Margin = New Padding(5, 4, 5, 4)
        Pivot_det0.Name = "Pivot_det0"
        Pivot_det0.Size = New Size(85, 21)
        Pivot_det0.TabIndex = 51
        Pivot_det0.TextAlign = HorizontalAlignment.Center
        ' 
        ' Pivot_det1
        ' 
        Pivot_det1.Font = New Font("Calibri", 8.0F)
        Pivot_det1.Location = New Point(306, 28)
        Pivot_det1.Margin = New Padding(5, 4, 5, 4)
        Pivot_det1.Name = "Pivot_det1"
        Pivot_det1.Size = New Size(85, 21)
        Pivot_det1.TabIndex = 53
        Pivot_det1.TextAlign = HorizontalAlignment.Center
        ' 
        ' Pivot_det2
        ' 
        Pivot_det2.Font = New Font("Calibri", 8.0F)
        Pivot_det2.Location = New Point(306, 53)
        Pivot_det2.Margin = New Padding(5, 4, 5, 4)
        Pivot_det2.Name = "Pivot_det2"
        Pivot_det2.Size = New Size(85, 21)
        Pivot_det2.TabIndex = 54
        Pivot_det2.TextAlign = HorizontalAlignment.Center
        ' 
        ' Pivot_det5
        ' 
        Pivot_det5.Font = New Font("Calibri", 8.0F)
        Pivot_det5.Location = New Point(306, 128)
        Pivot_det5.Margin = New Padding(5, 4, 5, 4)
        Pivot_det5.Name = "Pivot_det5"
        Pivot_det5.Size = New Size(85, 21)
        Pivot_det5.TabIndex = 57
        Pivot_det5.TextAlign = HorizontalAlignment.Center
        ' 
        ' Pivot_det4
        ' 
        Pivot_det4.Font = New Font("Calibri", 8.0F)
        Pivot_det4.Location = New Point(306, 103)
        Pivot_det4.Margin = New Padding(5, 4, 5, 4)
        Pivot_det4.Name = "Pivot_det4"
        Pivot_det4.Size = New Size(85, 21)
        Pivot_det4.TabIndex = 56
        Pivot_det4.TextAlign = HorizontalAlignment.Center
        ' 
        ' Pivot_det3
        ' 
        Pivot_det3.Font = New Font("Calibri", 8.0F)
        Pivot_det3.Location = New Point(306, 77)
        Pivot_det3.Margin = New Padding(5, 4, 5, 4)
        Pivot_det3.Name = "Pivot_det3"
        Pivot_det3.Size = New Size(85, 21)
        Pivot_det3.TabIndex = 55
        Pivot_det3.TextAlign = HorizontalAlignment.Center
        ' 
        ' Pivot_det8
        ' 
        Pivot_det8.Font = New Font("Calibri", 8.0F)
        Pivot_det8.Location = New Point(306, 204)
        Pivot_det8.Margin = New Padding(5, 4, 5, 4)
        Pivot_det8.Name = "Pivot_det8"
        Pivot_det8.Size = New Size(85, 21)
        Pivot_det8.TabIndex = 60
        Pivot_det8.TextAlign = HorizontalAlignment.Center
        ' 
        ' Pivot_det7
        ' 
        Pivot_det7.Font = New Font("Calibri", 8.0F)
        Pivot_det7.Location = New Point(306, 179)
        Pivot_det7.Margin = New Padding(5, 4, 5, 4)
        Pivot_det7.Name = "Pivot_det7"
        Pivot_det7.Size = New Size(85, 21)
        Pivot_det7.TabIndex = 59
        Pivot_det7.TextAlign = HorizontalAlignment.Center
        ' 
        ' Pivot_det6
        ' 
        Pivot_det6.AccessibleDescription = ""
        Pivot_det6.Font = New Font("Calibri", 8.0F)
        Pivot_det6.Location = New Point(306, 152)
        Pivot_det6.Margin = New Padding(5, 4, 5, 4)
        Pivot_det6.Name = "Pivot_det6"
        Pivot_det6.Size = New Size(85, 21)
        Pivot_det6.TabIndex = 58
        Pivot_det6.TextAlign = HorizontalAlignment.Center
        ' 
        ' Same_Z
        ' 
        Same_Z.AutoSize = True
        Same_Z.Location = New Point(1120, 24)
        Same_Z.Margin = New Padding(5, 4, 5, 4)
        Same_Z.Name = "Same_Z"
        Same_Z.Size = New Size(52, 17)
        Same_Z.TabIndex = 61
        Same_Z.Text = "Same"
        Same_Z.UseVisualStyleBackColor = True
        Same_Z.Visible = False
        ' 
        ' TextXLS
        ' 
        TextXLS.Enabled = False
        TextXLS.Font = New Font("Arial Narrow", 10.0F)
        TextXLS.Location = New Point(85, 525)
        TextXLS.Margin = New Padding(5, 4, 5, 4)
        TextXLS.Name = "TextXLS"
        TextXLS.Size = New Size(411, 23)
        TextXLS.TabIndex = 62
        TextXLS.Text = "Gupix-?????_Mat-??_Trc-??_Pivot-??.xls"
        ' 
        ' LabelAppend
        ' 
        LabelAppend.AutoSize = True
        LabelAppend.Font = New Font("Segoe UI", 8.0F)
        LabelAppend.Location = New Point(14, 532)
        LabelAppend.Margin = New Padding(5, 0, 5, 0)
        LabelAppend.Name = "LabelAppend"
        LabelAppend.Size = New Size(62, 13)
        LabelAppend.TabIndex = 64
        LabelAppend.Text = "Append to"
        ' 
        ' LabelNew
        ' 
        LabelNew.AutoSize = True
        LabelNew.Font = New Font("Segoe UI", 8.0F)
        LabelNew.Location = New Point(15, 532)
        LabelNew.Margin = New Padding(5, 0, 5, 0)
        LabelNew.Name = "LabelNew"
        LabelNew.Size = New Size(61, 13)
        LabelNew.TabIndex = 65
        LabelNew.Text = "Create       "
        ' 
        ' Check_Trc_As_Oxy
        ' 
        Check_Trc_As_Oxy.AutoSize = True
        Check_Trc_As_Oxy.Checked = True
        Check_Trc_As_Oxy.CheckState = CheckState.Checked
        Check_Trc_As_Oxy.Font = New Font("Calibri", 8.0F)
        Check_Trc_As_Oxy.Location = New Point(15, 5)
        Check_Trc_As_Oxy.Margin = New Padding(5, 4, 5, 4)
        Check_Trc_As_Oxy.Name = "Check_Trc_As_Oxy"
        Check_Trc_As_Oxy.Size = New Size(72, 17)
        Check_Trc_As_Oxy.TabIndex = 74
        Check_Trc_As_Oxy.Text = "Z as oxide"
        Check_Trc_As_Oxy.UseVisualStyleBackColor = True
        ' 
        ' Text_Lst_Ox_Trc
        ' 
        Text_Lst_Ox_Trc.Font = New Font("Calibri", 8.0F)
        Text_Lst_Ox_Trc.Location = New Point(10, 21)
        Text_Lst_Ox_Trc.Margin = New Padding(5, 4, 5, 4)
        Text_Lst_Ox_Trc.Name = "Text_Lst_Ox_Trc"
        Text_Lst_Ox_Trc.Size = New Size(207, 21)
        Text_Lst_Ox_Trc.TabIndex = 75
        Text_Lst_Ox_Trc.Text = "19,20,25,26,29,82"
        ' 
        ' Ck_AllAsOxy
        ' 
        Ck_AllAsOxy.AutoSize = True
        Ck_AllAsOxy.Font = New Font("Calibri", 8.0F)
        Ck_AllAsOxy.Location = New Point(117, 5)
        Ck_AllAsOxy.Margin = New Padding(5, 4, 5, 4)
        Ck_AllAsOxy.Name = "Ck_AllAsOxy"
        Ck_AllAsOxy.Size = New Size(86, 17)
        Ck_AllAsOxy.TabIndex = 76
        Ck_AllAsOxy.Text = "All Z as oxide"
        Ck_AllAsOxy.UseVisualStyleBackColor = True
        ' 
        ' ListFileInit
        ' 
        ListFileInit.GridLines = True
        ListFileInit.Location = New Point(1209, 120)
        ListFileInit.Margin = New Padding(5, 4, 5, 4)
        ListFileInit.Name = "ListFileInit"
        ListFileInit.Size = New Size(111, 211)
        ListFileInit.Sorting = SortOrder.Ascending
        ListFileInit.TabIndex = 79
        ListFileInit.UseCompatibleStateImageBehavior = False
        ListFileInit.View = View.List
        ListFileInit.Visible = False
        ' 
        ' StatusStrip1
        ' 
        StatusStrip1.ImageScalingSize = New Size(24, 24)
        StatusStrip1.Items.AddRange(New ToolStripItem() {ToolStripStatusLabel2, ToolStripStatusLabel1})
        StatusStrip1.Location = New Point(0, 865)
        StatusStrip1.Name = "StatusStrip1"
        StatusStrip1.Padding = New Padding(1, 0, 18, 0)
        StatusStrip1.Size = New Size(1320, 22)
        StatusStrip1.TabIndex = 80
        StatusStrip1.Text = "StatusStrip1"
        ' 
        ' ToolStripStatusLabel2
        ' 
        ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        ToolStripStatusLabel2.Size = New Size(0, 17)
        ' 
        ' ToolStripStatusLabel1
        ' 
        ToolStripStatusLabel1.Font = New Font("Arial Narrow", 9.75F)
        ToolStripStatusLabel1.Margin = New Padding(200, 3, 0, 2)
        ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        ToolStripStatusLabel1.Size = New Size(254, 17)
        ToolStripStatusLabel1.Text = "                             Status TRAUPIXE                         "
        ' 
        ' Chk_RoundValue
        ' 
        Chk_RoundValue.AutoSize = True
        Chk_RoundValue.Checked = True
        Chk_RoundValue.CheckState = CheckState.Checked
        Chk_RoundValue.Font = New Font("Calibri", 8.0F)
        Chk_RoundValue.Location = New Point(421, 48)
        Chk_RoundValue.Margin = New Padding(5, 4, 5, 4)
        Chk_RoundValue.Name = "Chk_RoundValue"
        Chk_RoundValue.Size = New Size(89, 17)
        Chk_RoundValue.TabIndex = 81
        Chk_RoundValue.Text = "Round values"
        Chk_RoundValue.UseVisualStyleBackColor = True
        ' 
        ' Adjust_Filter
        ' 
        Adjust_Filter.Font = New Font("Calibri", 14.0F)
        Adjust_Filter.Location = New Point(583, 10)
        Adjust_Filter.Margin = New Padding(5, 4, 5, 4)
        Adjust_Filter.Name = "Adjust_Filter"
        Adjust_Filter.Size = New Size(160, 40)
        Adjust_Filter.TabIndex = 82
        Adjust_Filter.Text = "Adjust Filter"
        Adjust_Filter.UseVisualStyleBackColor = True
        ' 
        ' TabControl1
        ' 
        TabControl1.Appearance = TabAppearance.FlatButtons
        TabControl1.Controls.Add(Calcul)
        TabControl1.Controls.Add(Tab_Adjust)
        TabControl1.Controls.Add(TabPage1)
        TabControl1.Font = New Font("Calibri", 8.0F)
        TabControl1.Location = New Point(360, 423)
        TabControl1.Margin = New Padding(5, 4, 5, 4)
        TabControl1.Name = "TabControl1"
        TabControl1.SelectedIndex = 0
        TabControl1.Size = New Size(840, 97)
        TabControl1.TabIndex = 83
        ' 
        ' Calcul
        ' 
        Calcul.BackColor = Color.Silver
        Calcul.Controls.Add(chb_skip_elem)
        Calcul.Controls.Add(txt_skip_elem)
        Calcul.Controls.Add(chk_external_ok)
        Calcul.Controls.Add(Text_gamma)
        Calcul.Controls.Add(Button2)
        Calcul.Controls.Add(Nb_Proc)
        Calcul.Controls.Add(Chk_RoundValue)
        Calcul.Controls.Add(MyPause)
        Calcul.Controls.Add(Check_Trc_As_Oxy)
        Calcul.Controls.Add(Text_Lst_Ox_Trc)
        Calcul.Controls.Add(Ck_AllAsOxy)
        Calcul.Controls.Add(Progress)
        Calcul.Controls.Add(Tps_Calc)
        Calcul.Location = New Point(4, 25)
        Calcul.Margin = New Padding(5, 4, 5, 4)
        Calcul.Name = "Calcul"
        Calcul.Padding = New Padding(5, 4, 5, 4)
        Calcul.Size = New Size(832, 68)
        Calcul.TabIndex = 0
        Calcul.Text = "Process Spectra"
        ' 
        ' chb_skip_elem
        ' 
        chb_skip_elem.AutoSize = True
        chb_skip_elem.Font = New Font("Calibri", 8.0F)
        chb_skip_elem.Location = New Point(239, 5)
        chb_skip_elem.Margin = New Padding(5, 4, 5, 4)
        chb_skip_elem.Name = "chb_skip_elem"
        chb_skip_elem.Size = New Size(72, 17)
        chb_skip_elem.TabIndex = 84
        chb_skip_elem.Text = "skip Z mtx"
        chb_skip_elem.UseVisualStyleBackColor = True
        chb_skip_elem.Visible = False
        ' 
        ' txt_skip_elem
        ' 
        txt_skip_elem.Font = New Font("Calibri", 8.0F)
        txt_skip_elem.Location = New Point(239, 21)
        txt_skip_elem.Margin = New Padding(5, 4, 5, 4)
        txt_skip_elem.Name = "txt_skip_elem"
        txt_skip_elem.Size = New Size(87, 21)
        txt_skip_elem.TabIndex = 83
        txt_skip_elem.Text = "82"
        txt_skip_elem.Visible = False
        ' 
        ' chk_external_ok
        ' 
        chk_external_ok.AutoSize = True
        chk_external_ok.Checked = True
        chk_external_ok.CheckState = CheckState.Checked
        chk_external_ok.Font = New Font("Calibri", 8.0F)
        chk_external_ok.Location = New Point(336, 48)
        chk_external_ok.Margin = New Padding(5, 4, 5, 4)
        chk_external_ok.Name = "chk_external_ok"
        chk_external_ok.Size = New Size(40, 17)
        chk_external_ok.TabIndex = 77
        chk_external_ok.Text = "ON"
        chk_external_ok.UseVisualStyleBackColor = True
        chk_external_ok.Visible = False
        ' 
        ' Text_gamma
        ' 
        Text_gamma.Font = New Font("Calibri", 8.0F)
        Text_gamma.Location = New Point(11, 45)
        Text_gamma.Margin = New Padding(5, 4, 5, 4)
        Text_gamma.Name = "Text_gamma"
        Text_gamma.Size = New Size(317, 21)
        Text_gamma.TabIndex = 82
        Text_gamma.Text = "File ""external-conc.csv"" not found"
        ' 
        ' Tab_Adjust
        ' 
        Tab_Adjust.Controls.Add(Label10)
        Tab_Adjust.Controls.Add(Label9)
        Tab_Adjust.Controls.Add(Label8)
        Tab_Adjust.Controls.Add(Label7)
        Tab_Adjust.Controls.Add(Label6)
        Tab_Adjust.Controls.Add(Label5)
        Tab_Adjust.Controls.Add(TextF_Z)
        Tab_Adjust.Controls.Add(TextF_Step)
        Tab_Adjust.Controls.Add(TextF_To)
        Tab_Adjust.Controls.Add(ComboBox_Type_F)
        Tab_Adjust.Controls.Add(TextF_From)
        Tab_Adjust.Controls.Add(Adjust_Filter)
        Tab_Adjust.Location = New Point(4, 25)
        Tab_Adjust.Margin = New Padding(5, 4, 5, 4)
        Tab_Adjust.Name = "Tab_Adjust"
        Tab_Adjust.Padding = New Padding(5, 4, 5, 4)
        Tab_Adjust.Size = New Size(832, 68)
        Tab_Adjust.TabIndex = 1
        Tab_Adjust.Text = "Adjust Absorbers"
        Tab_Adjust.UseVisualStyleBackColor = True
        ' 
        ' Label10
        ' 
        Label10.AutoSize = True
        Label10.Location = New Point(440, 37)
        Label10.Margin = New Padding(5, 0, 5, 0)
        Label10.Name = "Label10"
        Label10.Size = New Size(101, 13)
        Label10.TabIndex = 88
        Label10.Text = "mm / inert gas or Air"
        ' 
        ' Label9
        ' 
        Label9.AutoSize = True
        Label9.Location = New Point(439, 23)
        Label9.Margin = New Padding(5, 0, 5, 0)
        Label9.Name = "Label9"
        Label9.Size = New Size(103, 13)
        Label9.TabIndex = 87
        Label9.Text = "µm / solid absorbers"
        ' 
        ' Label8
        ' 
        Label8.AutoSize = True
        Label8.Location = New Point(130, 19)
        Label8.Margin = New Padding(5, 0, 5, 0)
        Label8.Name = "Label8"
        Label8.Size = New Size(57, 13)
        Label8.TabIndex = 85
        Label8.Text = "Z absorber"
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.Location = New Point(226, 46)
        Label7.Margin = New Padding(5, 0, 5, 0)
        Label7.Name = "Label7"
        Label7.Size = New Size(28, 13)
        Label7.TabIndex = 85
        Label7.Text = "Step"
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(112, 44)
        Label6.Margin = New Padding(5, 0, 5, 0)
        Label6.Name = "Label6"
        Label6.Size = New Size(17, 13)
        Label6.TabIndex = 85
        Label6.Text = "To"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(9, 46)
        Label5.Margin = New Padding(5, 0, 5, 0)
        Label5.Name = "Label5"
        Label5.Size = New Size(31, 13)
        Label5.TabIndex = 84
        Label5.Text = "From"
        ' 
        ' TextF_Z
        ' 
        TextF_Z.Location = New Point(216, 13)
        TextF_Z.Margin = New Padding(5, 4, 5, 4)
        TextF_Z.Name = "TextF_Z"
        TextF_Z.Size = New Size(47, 21)
        TextF_Z.TabIndex = 83
        TextF_Z.Text = "102"
        TextF_Z.TextAlign = HorizontalAlignment.Center
        ' 
        ' TextF_Step
        ' 
        TextF_Step.Location = New Point(266, 40)
        TextF_Step.Margin = New Padding(5, 4, 5, 4)
        TextF_Step.Name = "TextF_Step"
        TextF_Step.Size = New Size(69, 21)
        TextF_Step.TabIndex = 86
        TextF_Step.Text = "0.1"
        TextF_Step.TextAlign = HorizontalAlignment.Center
        ' 
        ' TextF_To
        ' 
        TextF_To.Location = New Point(146, 40)
        TextF_To.Margin = New Padding(5, 4, 5, 4)
        TextF_To.Name = "TextF_To"
        TextF_To.Size = New Size(54, 21)
        TextF_To.TabIndex = 85
        TextF_To.Text = "1"
        TextF_To.TextAlign = HorizontalAlignment.Center
        ' 
        ' ComboBox_Type_F
        ' 
        ComboBox_Type_F.FormattingEnabled = True
        ComboBox_Type_F.Location = New Point(18, 13)
        ComboBox_Type_F.Margin = New Padding(5, 4, 5, 4)
        ComboBox_Type_F.Name = "ComboBox_Type_F"
        ComboBox_Type_F.Size = New Size(91, 21)
        ComboBox_Type_F.TabIndex = 10
        ComboBox_Type_F.TabStop = False
        ' 
        ' TextF_From
        ' 
        TextF_From.Location = New Point(56, 40)
        TextF_From.Margin = New Padding(5, 4, 5, 4)
        TextF_From.Name = "TextF_From"
        TextF_From.Size = New Size(53, 21)
        TextF_From.TabIndex = 84
        TextF_From.Text = "0.1"
        TextF_From.TextAlign = HorizontalAlignment.Center
        ' 
        ' TabPage1
        ' 
        TabPage1.Controls.Add(Button_Extract)
        TabPage1.Location = New Point(4, 25)
        TabPage1.Margin = New Padding(3, 4, 3, 4)
        TabPage1.Name = "TabPage1"
        TabPage1.Padding = New Padding(3, 4, 3, 4)
        TabPage1.Size = New Size(832, 68)
        TabPage1.TabIndex = 2
        TabPage1.Text = "Extract hdf5 to ASCII"
        TabPage1.UseVisualStyleBackColor = True
        ' 
        ' Button_Extract
        ' 
        Button_Extract.Font = New Font("Calibri", 12.0F)
        Button_Extract.Location = New Point(290, 11)
        Button_Extract.Margin = New Padding(3, 4, 3, 4)
        Button_Extract.Name = "Button_Extract"
        Button_Extract.Size = New Size(193, 41)
        Button_Extract.TabIndex = 0
        Button_Extract.Text = "Extract hdf5 to GUPIX files"
        Button_Extract.UseVisualStyleBackColor = True
        ' 
        ' Button6
        ' 
        Button6.Location = New Point(1217, 338)
        Button6.Margin = New Padding(5, 4, 5, 4)
        Button6.Name = "Button6"
        Button6.Size = New Size(103, 32)
        Button6.TabIndex = 84
        Button6.Text = "Button6"
        Button6.UseVisualStyleBackColor = True
        Button6.UseWaitCursor = True
        Button6.Visible = False
        ' 
        ' ToolTip1
        ' 
        ToolTip1.AutomaticDelay = 400
        ToolTip1.AutoPopDelay = 6000
        ToolTip1.InitialDelay = 400
        ToolTip1.ReshowDelay = 80
        ' 
        ' Button7
        ' 
        Button7.Location = New Point(1209, 65)
        Button7.Margin = New Padding(5, 4, 5, 4)
        Button7.Name = "Button7"
        Button7.Size = New Size(97, 36)
        Button7.TabIndex = 86
        Button7.Text = "Button7"
        Button7.UseVisualStyleBackColor = True
        Button7.Visible = False
        ' 
        ' Button8
        ' 
        Button8.Location = New Point(1217, 379)
        Button8.Margin = New Padding(5, 4, 5, 4)
        Button8.Name = "Button8"
        Button8.Size = New Size(103, 42)
        Button8.TabIndex = 87
        Button8.Text = "Button8"
        Button8.UseVisualStyleBackColor = True
        Button8.Visible = False
        ' 
        ' Box_txtFiltre
        ' 
        Box_txtFiltre.AcceptsReturn = True
        Box_txtFiltre.Font = New Font("Calibri", 8.0F)
        Box_txtFiltre.Location = New Point(654, 33)
        Box_txtFiltre.Margin = New Padding(5, 4, 5, 4)
        Box_txtFiltre.Name = "Box_txtFiltre"
        Box_txtFiltre.Size = New Size(127, 21)
        Box_txtFiltre.TabIndex = 88
        Box_txtFiltre.Text = "*"
        Box_txtFiltre.TextAlign = HorizontalAlignment.Center
        ' 
        ' Label11
        ' 
        Label11.AutoSize = True
        Label11.Font = New Font("Calibri", 8.0F)
        Label11.Location = New Point(655, 21)
        Label11.Margin = New Padding(5, 0, 5, 0)
        Label11.Name = "Label11"
        Label11.Size = New Size(76, 13)
        Label11.TabIndex = 89
        Label11.Text = "Filter filename"
        ' 
        ' ListBox_HDF5
        ' 
        ListBox_HDF5.Font = New Font("Calibri", 8.25F)
        ListBox_HDF5.FormattingEnabled = True
        ListBox_HDF5.Location = New Point(11, 368)
        ListBox_HDF5.Margin = New Padding(5, 4, 5, 4)
        ListBox_HDF5.Name = "ListBox_HDF5"
        ListBox_HDF5.Size = New Size(339, 69)
        ListBox_HDF5.TabIndex = 90
        ' 
        ' TxtBox_HDF5_File
        ' 
        TxtBox_HDF5_File.Font = New Font("Calibri", 9.0F)
        TxtBox_HDF5_File.Location = New Point(11, 455)
        TxtBox_HDF5_File.Margin = New Padding(5, 4, 5, 4)
        TxtBox_HDF5_File.Name = "TxtBox_HDF5_File"
        TxtBox_HDF5_File.Size = New Size(339, 22)
        TxtBox_HDF5_File.TabIndex = 91
        TxtBox_HDF5_File.TextAlign = HorizontalAlignment.Center
        ' 
        ' Panel1
        ' 
        Panel1.BackColor = Color.LightGray
        Panel1.Controls.Add(Pivot_det8)
        Panel1.Controls.Add(Pivot_det7)
        Panel1.Controls.Add(Pivot_det6)
        Panel1.Controls.Add(Pivot_det5)
        Panel1.Controls.Add(Pivot_det4)
        Panel1.Controls.Add(Pivot_det3)
        Panel1.Controls.Add(Pivot_det2)
        Panel1.Controls.Add(Pivot_det1)
        Panel1.Controls.Add(Pivot_det0)
        Panel1.Controls.Add(Par_det8)
        Panel1.Controls.Add(Check_det8)
        Panel1.Controls.Add(Check_det7)
        Panel1.Controls.Add(Check_det6)
        Panel1.Controls.Add(Par_det7)
        Panel1.Controls.Add(Par_det6)
        Panel1.Controls.Add(Par_det5)
        Panel1.Controls.Add(Check_det5)
        Panel1.Controls.Add(Check_det4)
        Panel1.Controls.Add(Check_det3)
        Panel1.Controls.Add(Par_det4)
        Panel1.Controls.Add(Par_det3)
        Panel1.Controls.Add(Par_det2)
        Panel1.Controls.Add(Check_det2)
        Panel1.Controls.Add(Check_det1)
        Panel1.Controls.Add(Check_det0)
        Panel1.Controls.Add(Par_det1)
        Panel1.Controls.Add(Par_det0)
        Panel1.Controls.Add(LstPar_Trc)
        Panel1.Font = New Font("Calibri", 9.75F)
        Panel1.Location = New Point(801, 129)
        Panel1.Margin = New Padding(3, 4, 3, 4)
        Panel1.Name = "Panel1"
        Panel1.Size = New Size(400, 285)
        Panel1.TabIndex = 92
        ' 
        ' Panel2
        ' 
        Panel2.BackColor = Color.Gray
        Panel2.Controls.Add(Par_Mat)
        Panel2.Controls.Add(CbDetMat)
        Panel2.Controls.Add(LstPar_Mat)
        Panel2.Font = New Font("Calibri", 9.75F)
        Panel2.Location = New Point(800, 43)
        Panel2.Margin = New Padding(3, 4, 3, 4)
        Panel2.Name = "Panel2"
        Panel2.Size = New Size(400, 62)
        Panel2.TabIndex = 93
        ' 
        ' Button_Run2
        ' 
        Button_Run2.Font = New Font("Calibri", 8.0F)
        Button_Run2.Location = New Point(487, 25)
        Button_Run2.Margin = New Padding(5, 4, 5, 4)
        Button_Run2.Name = "Button_Run2"
        Button_Run2.Size = New Size(71, 29)
        Button_Run2.TabIndex = 94
        Button_Run2.Text = "Run"
        Button_Run2.UseVisualStyleBackColor = True
        ' 
        ' TextProcessIf
        ' 
        TextProcessIf.Font = New Font("Calibri", 8.0F)
        TextProcessIf.Location = New Point(566, 33)
        TextProcessIf.Name = "TextProcessIf"
        TextProcessIf.Size = New Size(79, 21)
        TextProcessIf.TabIndex = 95
        TextProcessIf.Visible = False
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(1217, 430)
        Button1.Margin = New Padding(3, 4, 3, 4)
        Button1.Name = "Button1"
        Button1.Size = New Size(98, 34)
        Button1.TabIndex = 96
        Button1.Text = "Button1"
        Button1.UseVisualStyleBackColor = True
        Button1.Visible = False
        ' 
        ' Label_QFile
        ' 
        Label_QFile.AutoSize = True
        Label_QFile.BackColor = Color.Transparent
        Label_QFile.Font = New Font("Segoe UI", 8.0F, FontStyle.Bold)
        Label_QFile.ForeColor = Color.Black
        Label_QFile.Location = New Point(857, 26)
        Label_QFile.Margin = New Padding(5, 0, 5, 0)
        Label_QFile.Name = "Label_QFile"
        Label_QFile.Size = New Size(37, 13)
        Label_QFile.TabIndex = 97
        Label_QFile.Text = "Q-File"
        Label_QFile.Visible = False
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Font = New Font("Calibri", 8.0F)
        Label3.Location = New Point(11, 441)
        Label3.Margin = New Padding(5, 0, 5, 0)
        Label3.Name = "Label3"
        Label3.Size = New Size(88, 13)
        Label3.TabIndex = 98
        Label3.Text = "hdf5 selected file"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Font = New Font("Calibri", 8.0F)
        Label4.Location = New Point(14, 354)
        Label4.Margin = New Padding(5, 0, 5, 0)
        Label4.Name = "Label4"
        Label4.Size = New Size(50, 13)
        Label4.TabIndex = 99
        Label4.Text = "hdf5 files"
        ' 
        ' TextBox_hdf5_grps
        ' 
        TextBox_hdf5_grps.Enabled = False
        TextBox_hdf5_grps.Font = New Font("Calibri", 9.0F)
        TextBox_hdf5_grps.Location = New Point(11, 493)
        TextBox_hdf5_grps.Margin = New Padding(5, 4, 5, 4)
        TextBox_hdf5_grps.Multiline = True
        TextBox_hdf5_grps.Name = "TextBox_hdf5_grps"
        TextBox_hdf5_grps.Size = New Size(339, 23)
        TextBox_hdf5_grps.TabIndex = 100
        TextBox_hdf5_grps.TextAlign = HorizontalAlignment.Center
        ' 
        ' Label12
        ' 
        Label12.AutoSize = True
        Label12.Font = New Font("Calibri", 8.0F)
        Label12.Location = New Point(11, 477)
        Label12.Margin = New Padding(5, 0, 5, 0)
        Label12.Name = "Label12"
        Label12.Size = New Size(98, 13)
        Label12.TabIndex = 101
        Label12.Text = "hdf5 data available"
        ' 
        ' Form_Traupixe_H5_2024
        ' 
        AutoScaleDimensions = New SizeF(96F, 96F)
        AutoScaleMode = AutoScaleMode.Dpi
        AutoScroll = True
        AutoSize = True
        ClientSize = New Size(1221, 571)
        Controls.Add(ComboBox_Type_Calc)
        Controls.Add(Label12)
        Controls.Add(TextBox_hdf5_grps)
        Controls.Add(Label4)
        Controls.Add(Label3)
        Controls.Add(Label_QFile)
        Controls.Add(Button1)
        Controls.Add(TextProcessIf)
        Controls.Add(Button_Run2)
        Controls.Add(Label1)
        Controls.Add(Panel2)
        Controls.Add(Panel1)
        Controls.Add(TxtBox_HDF5_File)
        Controls.Add(ListBox_HDF5)
        Controls.Add(Label11)
        Controls.Add(Box_txtFiltre)
        Controls.Add(Button8)
        Controls.Add(Button7)
        Controls.Add(Button6)
        Controls.Add(TabControl1)
        Controls.Add(StatusStrip1)
        Controls.Add(ListFileInit)
        Controls.Add(LabelNew)
        Controls.Add(LabelAppend)
        Controls.Add(Same_Z)
        Controls.Add(TextXLS)
        Controls.Add(Text_Status)
        Controls.Add(ProgressBar1)
        Controls.Add(Label2)
        Controls.Add(BtRefresh)
        Controls.Add(LvFiles)
        Controls.Add(trvFolders)
        Controls.Add(ComboBoxDrive)
        Controls.Add(MenuStrip1)
        Font = New Font("Calibri", 8F)
        FormBorderStyle = FormBorderStyle.FixedSingle
        MainMenuStrip = MenuStrip1
        Margin = New Padding(5, 4, 5, 4)
        MaximizeBox = False
        Name = "Form_Traupixe_H5_2024"
        RightToLeftLayout = True
        StartPosition = FormStartPosition.Manual
        Text = "TrauPIXE 25.14 - C2RMF"
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        StatusStrip1.ResumeLayout(False)
        StatusStrip1.PerformLayout()
        TabControl1.ResumeLayout(False)
        Calcul.ResumeLayout(False)
        Calcul.PerformLayout()
        Tab_Adjust.ResumeLayout(False)
        Tab_Adjust.PerformLayout()
        TabPage1.ResumeLayout(False)
        Panel1.ResumeLayout(False)
        Panel1.PerformLayout()
        Panel2.ResumeLayout(False)
        Panel2.PerformLayout()
        ResumeLayout(False)
        PerformLayout()

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
        '  Me.Height = Screen.PrimaryScreen.Bounds.Height * 0.6
        '  Me.Width = Screen.PrimaryScreen.Bounds.Width * 0.6
        Me.BringToFront()
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

        Me.ListBox_HDF5.Visible = False
        Me.Label4.Visible = False
        Me.Label3.Visible = False
        Me.Label12.Visible = False
        Me.TxtBox_HDF5_File.Visible = False
        Me.TextBox_hdf5_grps.Visible = False

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
    Friend WithEvents Button1 As Button
    Friend WithEvents Label_QFile As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBox_hdf5_grps As TextBox
    Friend WithEvents Label12 As Label
End Class
