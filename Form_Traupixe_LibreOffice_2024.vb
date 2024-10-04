Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports ClosedXML.Excel
Imports ClosedXML.Excel.Ranges
Imports DocumentFormat.OpenXml.Drawing



Public Class Form_Traupixe_H5_2024

    Private Const ERROR_FILE_NOT_FOUND = 2&
    Private Const ERROR_PATH_NOT_FOUND = 3&
    Private Const ERROR_BAD_FORMAT = 11&
    Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
    Private Const SE_ERR_ASSOCINCOMPLETE = 27
    Private Const SE_ERR_DDEBUSY = 30
    Private Const SE_ERR_DDEFAIL = 29
    Private Const SE_ERR_DDETIMEOUT = 28
    Private Const SE_ERR_DLLNOTFOUND = 32
    Private Const SE_ERR_FNF = 2                ' file not found
    Private Const SE_ERR_NOASSOC = 31
    Private Const SE_ERR_PNF = 3                ' path not found
    Private Const SE_ERR_OOM = 8                ' out of memory
    Private Const SE_ERR_SHARE = 26



    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hWnd As Long, ByVal lpOperation As String,
  ByVal lpFile As String, ByVal lpParameters As String,
  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

    Public Declare Sub FITtoPNG_XPV8_visible1 Lib "c:\windows\system32\FITtoPNG_XPV8_V.dll" _
    Alias "FITtoPNG_XPV8_V" (ByVal cheminSource As String, ByVal CheminDest As String, ByVal nom As String, ByVal pause As Long)

    Public Declare Sub FiTtoPNG_LV16 Lib ".\FiTtoPNG_2023_DLL.dll" _
        Alias "FiTtoPNG_DLL_LV2016_2023" (ByVal Pause As Integer, ByVal CheminSource As String, ByVal CheminDest As String, ByVal Nom As String, ByVal Folder As String)

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetVersion Lib "kernel32" () As Long

    'Dim AppExcel As New Excel.Application
    'Dim worksheet As Excel.Worksheet
    'Dim workbook As Excel.Workbook
    ' Dim Fs_Log As Object
    Dim skip_Pb_mtx As Boolean
    Dim nb_process_custom As Boolean
    Dim Global_Nb_Swap As Integer
    Dim K_Name_HED_Mat As String
    Dim L_Name_HED_Mat As String
    Dim M_Name_HED_Mat As String
    Dim K_Path_HED_Mat As String
    Dim L_Path_HED_Mat As String
    Dim M_Path_HED_Mat As String
    Dim K_HED_Mat As Boolean
    Dim L_HED_Mat As Boolean
    Dim M_HED_Mat As Boolean
    Dim Use_HED_Mat As Boolean
    Dim Use_ext_charge_Mat As Boolean
    Dim Use_ext_charge_Trc(10) As Boolean
    Dim num_column_charge_csv_MAT As Integer
    Dim num_column_charge_csv_TRC(10) As Integer
    Dim limite_conc_red_ok As Integer

    Dim K_Name_HED_Trc(10) As String
    Dim L_Name_HED_Trc(10) As String
    Dim M_Name_HED_Trc(10) As String
    Dim K_Path_HED_Trc(10) As String
    Dim L_Path_HED_Trc(10) As String
    Dim M_Path_HED_Trc(10) As String
    Dim K_HED_Trc(10) As Boolean
    Dim L_HED_Trc(10) As Boolean
    Dim M_HED_Trc(10) As Boolean
    Dim Use_HED_Trc(10) As Boolean
    Dim Fatal_Error As Boolean
    Dim Error_Matrix(10) As Boolean
    Dim Error_Trace(10, 10) As Boolean
    Dim Process_If_Z As String
    Dim Process_If_Sign As String
    Dim Process_If_Value As String
    Dim Process_Abort(10) As Boolean
    Dim Charge_Exp(1, 1) As String
    Dim Det_name_with_charge(10) As String



    Dim gamma_filename(1000) As String
    Dim tab_gamma_external_value_ok(1000) As Boolean
    Dim gamma_conc(1000, 100) As Integer
    Dim gamma_conc_init(1000, 100) As Integer
    Dim gamma_conc_oxide(1000, 100) As Integer
    Dim sum_gamma_oxide(100) As Integer
    Dim sum_gamma_conc(100) As Integer
    Dim gamma_ok As Boolean
    Dim ext_tech(100) As String
    Dim nb_gamma_and_pixe As Integer

    Dim gamma_mode As Boolean
    Dim info_gamma_name(100) As String
    Dim info_gamma_z(100) As String
    Dim tab_rapport_oxide_gamma(100) As Single
    Dim nb_gamma As Integer
    Dim tab_select_file_indices() As Integer
    Public Structure MyMat
        Public Z As String
        Public conc As Double
        Public layer As String
        Public name As String
    End Structure


    Public atomic_info_Z(125) As String
    Public atomic_info_name(125) As String
    Public atomic_info_mass(125) As Single

    Dim Gupix_Path = "c:\gupixwin\gupix"
    Dim Ext_Trc0 As String
    Dim Ext_Trc1 As String
    Dim Ext_Trc2 As String
    Dim Ext_Trc3 As String
    Dim Ext_Trc4 As String
    Dim Ext_Trc5 As String
    Dim Ext_Trc6 As String
    Dim Ext_Trc7 As String
    Dim Ext_Trc8 As String

    Dim Ext_file_Mat As String
    Dim Ext_det_mat As String
    Dim Ext_det0 As String
    Dim Ext_det1 As String
    Dim Ext_det2 As String
    Dim Ext_det3 As String
    Dim Ext_det4 As String
    Dim Ext_det5 As String
    Dim Ext_det6 As String
    Dim Ext_det7 As String
    Dim Ext_det8 As String

    Dim Lst_Files_Det0() As String
    Dim Lst_Files_Det1() As String
    Dim Lst_Files_Det2() As String
    Dim Lst_Files_Det3() As String
    Dim Lst_Files_Det4() As String
    Dim Lst_Files_Det5() As String
    Dim Lst_Files_Det6() As String
    Dim Lst_Files_Det7() As String
    Dim Lst_Files_Det8() As String



    '  ######## HDF5 
    Dim Ref_DataSet_ToRead(10) As String
    Dim Chemin_hdf5 As String
    Dim AllSpectres_hdf5(10, 20, 2048) As String
    Dim Attrib_Spectrum(10, 10, 10) As String
    Dim hdf5_mode As Boolean
    Dim glob_gamma_mode As Boolean
    Dim path_gamma As String


    Dim Nom_Projet As String
    Dim NomOrdi As String
    Dim Dir_EDF As String
    Dim Nb_Spectres_EDF As Integer
    Dim Nb_Canaux As Integer
    Dim Myseek_Mat As Integer

    Dim Tab_Name_File_1() As String
    Dim Tab_Comment_1() As String
    Dim Tab_IsPonctuel_1() As Boolean
    Dim Info_Mat_Raie_1(,) As Boolean
    Dim Val_Mat_Height_1(,) As Integer
    Dim Val_Mat_Area_1(,) As Integer
    Dim Val_Mat_Conc_1(,) As Integer
    Dim Val_Mat_Oxyde_1(,) As Integer
    Dim Val_Mat_LOD_1(,) As Integer
    Dim Val_Mat_Fit_Error_1(,) As Double
    Dim Val_Mat_Mtx_1(,) As String
    Dim Val_Mat_Best_1(,) As Integer
    Dim Val_Mat_Best_ConcOK_1(,) As Double
    Dim Val_Mat_Best_Conc_1(,) As Double
    Dim Val_Mat_Best_Yes_1(,)
    Dim Val_Mat_Best_RED_1(,)
    Dim Val_Mat_Best_ConcOk_RED_1(,)
    Dim Val_Mat_Best_StrConc_1(,)
    Dim Val_Mat_Best_Yes_RED_1(,)
    Dim Val_Inv_Mtx(,) As String
    Dim NomDet_Mat_1 As String
    Dim NomDet_Trc_1() As String
    Dim Val_Mat_Total_Error_1(,) As Double
    Dim Val_Mat_Final_Error_1(,) As Double
    Dim Val_Mat_Y_N_Q_1(,) As String
    Dim Val_Conc_S_RED_ppm_1(,) As String
    Dim Val_Conc_S_RED100_1(,) As String
    Dim Val_Error_S_1(,) As String
    Dim Val_Conc_And_Error_1(,) As String
    Dim Val_Conc_And_Error100_1(,) As String
    Dim Val_Conc_S_100_1(,) As String
    Dim Val_Conc_S_ppm_1(,) As String
    Dim Val_Choix_S_1(,) As String
    Dim Val_YNQ_Final_1(,) As String



    Dim Val_Trc_Height_1(,) As Integer
    Dim Val_Trc_Conc_1(,) As Integer
    Dim Val_Trc_Oxyde_1(,) As Integer
    Dim Val_Trc_LOD_1(,) As Integer
    Dim Val_Trc_Area_1(,) As Integer
    Dim Val_Trc_Fit_Error_1(,) As Double
    Dim Val_Trc_Total_Error_1(,) As Double
    Dim Val_Trc_WithPivot_Error_1(,) As Double
    Dim Val_Trc_Y_N_Q_1(,) As String


    Dim USACulture = New CultureInfo("en-US")
    Dim Tab_IsPonctuel() As Boolean
    Dim Tab_Comment() As String
    Dim i As Integer
    Dim i2 As Integer
    Dim PathData As String
    Dim Ext_Mat As String '= "*.x0"
    Dim Ext_Trc As String '= "*.x10"
    Dim Ext_Par_Mat As String '= "*BE0*.par"
    Dim Ext_Par_Trc As String '= "*HE10*.par"
    Dim NomDet_Mat As String
    Dim NomDet_Trc() As String
    Dim Myinit As Boolean = False
    Dim Select_Par_files As Integer
    Dim Offset_Excel As Integer
    Dim Ratio As String
    Dim Pivot1(20) As String
    Dim OnlyTrace As Boolean
    Dim MyDepth As Boolean
    Dim Extract_Trace As Boolean
    Dim Calcul_With_Trc As Boolean
    Dim Calcul_Without_Pivot As Boolean
    Dim Nb_Elem_Unique As Integer
    Dim Nb_Elem_Unique_sans_external As Integer
    Dim Lect_Depth As Boolean

    Dim Chemin_Rapport As String
    Dim Chemin_Processed_Data As String
    Dim hConsole As Long
    Dim Chemin_Data As String
    Dim Fichier_Matrix() As String, Fichier_Trace() As String, Fichier_Trace0() As String, Fichier_Trace1() As String
    Dim Fichier_Trace2() As String, Fichier_Trace3() As String, Fichier_Trace4() As String
    Dim Fichier_Trace5() As String, Fichier_Trace6() As String, Fichier_Trace7() As String
    Dim Fichier_Trace8() As String

    Dim Nb_Elements_Mat As Integer
    Dim Nb_Elements_Trc() As Integer
    Dim Nb_Oxyde_Mat As Integer, Nb_Oxyde_X2 As Integer
    Dim Nb_Ligne_Oxyde_X1 As Integer, Nb_Ligne_Oxyde_X2 As Integer
    Dim Nb_Oxyde_Trc() As Integer
    Dim Matrix As Mat
    Dim refresh_File As Boolean
    Dim Tab_File_Par_Trc(10) As String
    Dim Spectrum_Trc(,) As String
    Dim SwapDrive As String
    Dim Dir_Calc As String
    Dim Tab_Pivot(10, 10) As Integer
    Dim Tab_Val_Mat(50) As Val_Elem
    Dim Tab_Val_Trc(1, 1) As Val_Elem
    Dim Tab_Info_Mat As Info_Elem
    Dim Tab_Info_Trc() As Info_Elem
    Dim Info_Trc_Raie(,) As Boolean
    Dim Info_Mat_Raie(,) As Boolean
    Dim Tab_Matrix() As Mat
    Dim Tab_Oxyde_X0() As Val_Oxyde
    Dim Tab_Val_Oxyde_Mat() As Val_Oxyde
    Dim Tab_Val_Oxyde_Trc(,) As Val_Oxyde
    Dim Tab_Info_Oxyde_Trc() As Info_Oxyde
    Dim Info_Oxyde_Mat As Info_Oxyde
    Dim Indice_Pivot_Mat(10) As Integer
    Dim Indice_Pivot_trc(10, 10) As Integer
    Dim Tab_Trc_as_Oxy() As Integer
    Dim Tab_Mat_Error() As Boolean
    Dim Nom_Excel_Trx_O As String
    Dim Nb_Elements_Mtx_inv As Integer
    Dim Nb_Trc_Total As Integer
    Dim Nb_Elements_100_Mat As Integer
    Dim Nb_Elements_100_Trc(100) As Integer
    Dim Tab_Name_File() As String
    Dim Adjust_Filter_B As Boolean
    Dim Z_Elem_Inv As String
    Dim Conc_Invisible As Double

    Dim StrFiltres_X1 As String, StrFiltres_Trc(10) As String
    Public MyChargeStd As Single
    Dim Fact_Correctif As Single = 1
    Dim Pivot As Integer
    Dim Str_Version_Os As String
    Dim IntNb_File As Integer
    Dim Ext_Images_Mat As String
    Dim Ext_Images_Trc As String
    Dim Piv As String
    Dim Txt_Fichier_PAR_Mat_Filter() As String
    Dim Txt_Fichier_PAR_Trc_Filter() As String
    Dim Txt_Fichier_PAR_Mat_HED() As String
    Dim Txt_Fichier_PAR_Trc_HED(,) As String

    Dim Txt_Fichier_PAR_Mat As String
    Dim Txt_Fichier_PAR_Trc As String
    Dim Txt_Fichier_PAR_Trc_P1 As String
    Dim Txt_Fichier_PAR_Trc_P2 As String
    Dim Nb_Ligne_Trc As Integer
    Dim Val_Charge_Trc() As Double
    Dim Val_Charge_Trc_Init As Double
    Dim Val_Charge_Trc_Mean As Double

    Dim Nb_Etoile As Integer
    Dim Chemin_GupixWin As String
    Dim Chemin_GupixWin_Multi() As String
    Dim Chemin_Temp As String
    Dim Nb_Process As Integer
    Dim Nb_Process1 As Integer
    Dim First_Init As Boolean
    Dim First_Init_Trc(10) As Boolean
    Dim Tab_Num_Trc(10) As Integer
    Dim Nb_Trc As Integer
    Dim File_Trc_Ext As Integer

    'Dim xlApp As Excel.Application
    Dim xlBook As XLWorkbook
    Dim xlApp As Object
    'Dim xlBook As Object

    Dim newexcel As Object
    'Dim xlBook As Object
    Dim xlSheet_Info = New XLWorkbook() 'Worksheet
    Dim xlSheet_ExpData = New XLWorkbook()
    Dim xlSheet_LOD = New XLWorkbook 'As Excel.Worksheet
    Dim xlSheet_Area = New XLWorkbook 'As Excel.Worksheet
    Dim xlSheet_Conc = New XLWorkbook 'As Excel.Worksheet
    Dim xlSheet_Oxyde = New XLWorkbook 'As Excel.Worksheet
    'Dim xlSheet_Stat_Err As Excel.Worksheet
    Dim xlSheet_Height = New XLWorkbook
    Dim xlSheet_Fit_Err = New XLWorkbook
    Dim xlSheet_Masse = New XLWorkbook
    Dim xlSheet_Depth = New XLWorkbook
    Dim xlSheet_Mtx = New XLWorkbook
    ' Dim xlSheet_Conc_100 As Object
    Dim xlSheet_Final_Conc_100 = New XLWorkbook
    ' Dim xlSheet_Conc_ppm As Object
    Dim xlSheet_S_Conc_100 = New XLWorkbook
    Dim xlSheet_S_Conc_ppm = New XLWorkbook
    Dim xlSheet_Choix_S = New XLWorkbook

    Dim xlSheet_Final_Conc_ppm = New XLWorkbook
    Dim xlSheet_ConcAndError100 = New XLWorkbook
    Dim xlSheet_ConcAndErrorPPM = New XLWorkbook
    Dim xlSheet_S_Conc_ppm_RED = New XLWorkbook
    Dim xlSheet_S_Conc_100_RED = New XLWorkbook
    Dim xlSheet_Conc_Error = New XLWorkbook
    Dim xlSheet_Total_Error = New XLWorkbook
    Dim xlSheet_Conc_Error_V2 = New XLWorkbook
    Dim xlSheet_Error = New XLWorkbook
    Dim xlSheet_S_Conc_Error_ppm = New XLWorkbook
    Dim xlSheet_S_Conc_Error_100 = New XLWorkbook


    Dim thread As System.Threading.Thread

    'Dim Thread_Best_Value(100) As System.Threading.Thread
    Dim Thread_Excel_Format_Data As System.Threading.Thread
    Dim Thread_Excel_Format_Style As System.Threading.Thread
    Dim Thread_Excel_Format_Italic As System.Threading.Thread
    Dim thread_tab_Calcul_Ecriture_Charge(100) As System.Threading.Thread
    'Dim thread_tab_Ecriture_Charge(100) As System.Threading.Thread
    Dim thread_tab_Element(100) As System.Threading.Thread
    Dim thread_tab_Just_Element(100) As System.Threading.Thread
    Dim thread_tab_oxyde(100) As System.Threading.Thread
    Dim thread_tab_Recup_Filter(100) As System.Threading.Thread
    Dim thread_tab_Data_Excel(100) As System.Threading.Thread
    Dim thread_tab_Calcul_Best_Conc(100) As System.Threading.Thread
    Dim thread_tab_FitToPNG(100) As System.Threading.Thread
    Dim thread_tab_FitToPNG_TRC(100) As System.Threading.Thread
    Dim thread_tab_best_conc(100) As System.Threading.Thread
    Dim thread2 As System.Threading.Thread


    Dim Tab_Info_Mat_OK(50) As Info_Elem
    Dim Tab_Info_Trc_OK(50, 50) As Info_Elem
    Dim Tab_Last_Entete(100) As String
    Dim Tab_Entete_Mat(150) As String
    Dim Tab_Entete_Trc(50, 50) As String
    Dim Tab_Entete_100(50) As String
    Dim Tab_Z_100(50) As Integer
    Dim Tab_Entete_Inv(20) As String

    Dim Z_Elem_100_Mat(100) As Integer
    Dim Z_Elem_100_Trc(50, 100) As Integer

    Dim Indice_Elem_100_Mat(100) As Integer
    Dim Indice_Elem_100_Trc(50, 100) As Integer

    Dim Info_Experience_Mat() As Info_Exp
    Dim Info_Experience_Trc(,) As Info_Exp
    Dim Val_Mat_Mtx(,) As String
    Dim Val_Mat_Height(,) As Integer
    Dim Val_Mat_Conc(,) As Integer
    Dim Val_Mat_Oxyde(,) As Integer
    Dim Val_Mat_LOD(,) As Integer
    Dim Val_Mat_Area(,) As Integer
    Dim Val_Mat_Stat_Error(,) As Double
    Dim Val_Mat_Fit_Error(,) As Double
    Dim Val_Mat_Total_Error(,) As Double
    Dim Val_Mat_Final_Error(,) As Double
    Dim Val_Mat_Y_N_Q(,) As String

    Dim Val_Trc_Height(,) As Integer
    Dim Val_Trc_Conc(,) As Integer
    Dim Val_Trc_Oxyde(,) As Integer
    Dim Val_Trc_LOD(,) As Integer
    Dim Val_Trc_Area(,) As Integer
    Dim Val_Trc_Stat_Error(,) As Double
    Dim Val_Trc_Fit_Error(,) As Double
    Dim Val_Trc_Total_Error(,) As Double
    Dim Val_Trc_WithPivot_Error(,) As Double
    Dim Val_Trc_Pivot_Error(,) As Double

    Dim Val_Trc_Y_N_Q(,) As String
    Dim Val_Trc_Error_Pivot(,)

    Dim Val_Mat_Best(,) As Integer
    Dim Val_Mat_Best_Yes(,) As Double
    Dim Val_Mat_Best_Yes_RED(,) As Double


    Dim Val_Trc_Best_Yes(,) As Double
    Dim Val_Trc_Best_Yes_RED(,) As Double
    Dim Val_Mat_Best_Conc(,) As Integer
    Dim Str_Mat_Conc_100(,) As String
    Dim Val_Mat_Conc_S100(,) As String
    Dim Val_Mat_Conc_RED(,) As String
    Dim Val_Choix_S(,) As String
    Dim Val_Conc_S_100(,) As String
    Dim Val_Conc_S_ppm(,) As String
    Dim Val_Conc_S_RED_ppm(,) As String
    Dim Val_Conc_S_RED_100(,) As String
    Dim Val_Conc_S_RED100(,) As String
    Dim Val_Error_S(,) As String


    Dim Val_Conc_ppm_Final(,) As String
    Dim Val_YNQ_Final(,) As String
    Dim Val_Mat_Conc_ppm(,) As String

    Dim Val_Mat_Best_RED(,) As Integer
    Dim Val_Mat_Best_ConcOk_RED(,) As String
    Dim Val_Mat_Best_StrConc(,) As String
    Dim Val_Mat_Conc_Error(,) As String
    Dim Val_Mat_Conc_Error_V2(,) As String


    Dim Val_Trc_Best_Conc(,) As Integer
    Dim Val_Trc_Conc100(,) As String
    Dim Val_Trc_Conc_ppm(,) As String
    Dim Val_Trc_Conc_RED(,) As String
    Dim Val_Trc_Best_Conc_Final(,) As String

    Dim Val_Trc_Best_ConcOk_RED(,) As String
    Dim Val_Trc_Best_StrConc(,) As String
    Dim Val_Trc_Conc_Error(,) As String
    Dim Val_Trc_Conc_Error_V2(,) As String
    ' Dim Val_Trc_S_Error(,) As String

    Dim Val_Conc_And_Error(,) As String
    Dim Val_Conc_And_Error100(,) As String

    Dim SwapDrive_Name As String
    Dim SwapDrive1 As String
    Dim SwapDrive2 As String
    Dim SwapDrive3 As String
    Dim SwapDrive4 As String
    Dim Tab_Swap(10) As String



    Private Structure Val_Elem
        Public Val_Mtx() As String
        Public Conc() As Double
        Public Area() As Double
        Public Trans() As Double
        Public Stat_Err() As Double
        Public Fit_Err() As Double
        Public LOD() As Long
        Public YN() As String
        Public Depth() As Single
        Public Charge As String

        Public Count_Rate As String
        Public Current As String
        Public Chi2 As String
        Public Res As String
        Public Selected_Pivot() As Integer

        Public TotalError() As String
        Public StrConc() As String
        Public Yes() As Double
        Public Yes_RED() As Double
        Public Best() As Double
        Public Best_RED() As Double

        Public Conc100() As String
        Public ConcPPM() As String
        Public Error100() As String
        Public ErrorPPM() As String
        Public ConcAndError100() As String
        Public ConcAndErrorPPM() As String

        Public ConcOK() As String
        Public ConcOK_RED() As String
        Public Nb_Dig() As Integer
        Public ConcOK_Final() As String
        Public ConcOK_Final_V2() As String
        Public Conc_Error() As String
        Public Conc_Error_V2() As String


    End Structure

    Private Structure Info_Exp
        Public Charge As String
        Public Count_Rate As String
        Public Current As String
        Public Chi2 As String
        Public Res As String
        Public Selected_Pivot As String
        Public New_charge As String
        Public Filters As String

    End Structure




    Private Structure Info_Elem
        Public Z() As Integer
        Public Elem() As String
        Public Raie() As String
        Public Inv() As String
        Public Nom_V2() As String
    End Structure


    Private Structure Info_Oxyde
        Public Z() As String
        Public nom() As String
    End Structure


    Private Structure Val_Oxyde
        Public Conc_Oxy() As Double
    End Structure

    Private Structure Mat
        Public Z() As Integer
        Public Raie() As String
    End Structure

    Structure Struct_Parametres_Thread
        Public voie As Integer
        Public Fact_Correct As Single
        Public Num_Proc As Integer
        Public Num_File As Integer
        Public Num_Trc As Integer
        Public Nb_Calcul As Integer
        Public Num_Data As Integer
        Public Val_Charge_Trc As Double
        Public Offset_Trc As Integer
        Public File_Name As String
    End Structure
    Structure parametres_best_conc_Thread
        Dim indx_file As Integer
        Dim num_process As Integer
        Dim nb_trace As Integer
        'tab_select_file_indices(l + ((Nb_Process1) * (J - 1))), l, Nb_Trc
    End Structure
    Private Sub FolderBrowserDialog1_HelpRequest(sender As Object, e As EventArgs)

    End Sub

    Private Sub File4_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub



    Private Sub ComboBoxDrive_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxDrive.SelectedIndexChanged
        Dim Drive_Short As String
        trvFolders.Nodes.Clear()
        Dim Tnode As TreeNode = trvFolders.Nodes.Add("(Drive )" + ComboBoxDrive.SelectedItem) 'Add Main Node
        Drive_Short = Mid(ComboBoxDrive.SelectedItem, 1, 3)
        AddAllFolders(Tnode, Drive_Short)
        trvFolders.Select()

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs)

    End Sub


    Private Sub AddAllFolders(ByVal TNode As TreeNode, ByVal FolderPath As String)

        Dim TA As Array
        '  

        Try
            TA = Directory.GetDirectories(FolderPath, "*")
            Array.Sort(TA)
            For Each FolderNode As String In TA 'Directory.GetDirectories(FolderPath, "*") 'Load All Sub Folders 

                Dim SubFolderNode As TreeNode = TNode.Nodes.Add(FolderNode.Substring(FolderNode.LastIndexOf("\"c) + 1)) 'Add Each Sub Folder Name

                SubFolderNode.Tag = FolderNode 'Set Tag For Each Sub Folder

                SubFolderNode.Nodes.Add("Loading...")

            Next

        Catch ex As Exception

            MessageBox.Show(ex.Message) 'Something Went Wrong

        End Try

    End Sub


    Private Sub trvFolders_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvFolders.AfterSelect

        'Clear Existing Items
        trvFolders.SelectedNode.Expand()
        If trvFolders.SelectedNode.Nodes.Count = 1 AndAlso trvFolders.SelectedNode.Nodes(0).Text = "Loading..." Then
            trvFolders.SelectedNode.Nodes.Clear() 'Reset
            AddAllFolders(trvFolders.SelectedNode, CStr(trvFolders.SelectedNode.Tag))
        End If
        TxtBox_HDF5_File.Text = ""
        Myinit = True
        hdf5_mode = False
        List_Par_Files_Trc()
        Maj_Files_Mat()
        Ext_Par_Trc = "*.par"
        Maj_Par_Files_Trc(Par_det0, "det0")
        Maj_Par_Files_Mat()
        Maj_HDF5_Files()
        Chemin_Data = CStr(trvFolders.SelectedNode.Tag) 'trvFolders.SelectedNode

        If Chemin_Data <> Nothing And Len(Chemin_Data) > 4 Then
            If trvFolders.SelectedNode.Parent.Tag <> Nothing Then
                Chemin_Processed_Data = trvFolders.SelectedNode.Parent.Tag & "\Processed_data\" ''Chemin_Data = processed_data  ''''' A FAIRE ''''
            Else
                Chemin_Processed_Data = Chemin_Data & "\Processed_data\" ''Chemin_Data = processed_data  ''''' A FAIRE ''''
            End If

        Else
            Chemin_Processed_Data = Chemin_Data
        End If
        Load_config_exp_init()
        Load_charge_exp_csv()
        Det_use_charge()
    End Sub

    Private Sub BtRefresh_Click(sender As Object, e As EventArgs) Handles BtRefresh.Click
        List_Par_Files_Trc()
        If hdf5_mode = True Then
            List_HDF5_group(Chemin_hdf5)
        Else
            Maj_Files_Mat()
        End If

    End Sub

    Sub Maj_Files_Mat()

        Dim FileExtension As String 'Stores File Extension
        Dim SubItemIndex As Integer 'Sub Item Counter
        LvFiles.Items.Clear()
        'Par_Mat.Text = ""

        Dim folder As String
        Dim ext_trc_is_ext As Boolean = False
        Dim pos1 As Integer
        Dim str_filter As String
        IntNb_File = 0

        Try
            folder = CStr(trvFolders.SelectedNode.Tag) 'Folder Name
        Catch ex As Exception
            folder = "c:\"
        End Try

        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

            Try
                ''''''''''''''''********************************************************************************************** REGARDE SI DEF Fichier est une extention de fichier ou un mot clef (INFN)
                For Each file As String In Directory.GetFiles(folder, "*" & Ext_Mat & "*") 'Get Files In Folder
                    FileExtension = System.IO.Path.GetExtension(file) 'Get File Extension(s)
                    FileExtension = Strings.Right(FileExtension, Len(FileExtension) - 1)
                    pos1 = InStr(FileExtension, Ext_Mat, CompareMethod.Text)
                    If pos1 > 0 Then
                        ext_trc_is_ext = True
                        Exit For
                    End If

                Next
            Catch ex As Exception

            End Try
        End If

        If ext_trc_is_ext Then ' Cas ou le mot clef est l'extention du fichier 'ex: X1
            If Box_txtFiltre.Text = "*" Or Box_txtFiltre.Text = "" Then
                str_filter = "*." & Ext_Mat
            Else
                str_filter = "*" & Box_txtFiltre.Text & "*." & Ext_Mat
            End If
        Else
            str_filter = "*" & Ext_Mat & "*"   'Cas ou le mot clef n'est pas l'extention mais un vrai mot clef (ex: ROMA)
            Box_txtFiltre.Text = "*"

        End If

        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

            Try
                ''''''''''''''''********************************************************************************************** MATRICE DATA FILES
                For Each file As String In Directory.GetFiles(folder, str_filter) 'Get Files In Folder
                    FileExtension = System.IO.Path.GetExtension(file) 'Get File Extension(s)
                    If LCase(FileExtension) <> ".par" And LCase(FileExtension) <> ".ini" Then 'And LCase(FileExtension) <> ".xls" And LCase(FileExtension) <> ".xlsx" Then
                        LvFiles.Items.Add(System.IO.Path.GetFileNameWithoutExtension(file))
                        Ext_det_mat = FileExtension
                        SubItemIndex += 1

                    End If
                Next

            Catch ex As Exception 'Something Went Wrong


                MessageBox.Show(ex.Message)
            Finally
                If SubItemIndex = 0 Then
                    LvFiles.Items.Add("No files ...")
                Else
                    ReDim Lst_Files_Det0(SubItemIndex)
                    ReDim Lst_Files_Det1(SubItemIndex)
                    ReDim Lst_Files_Det2(SubItemIndex)
                    ReDim Lst_Files_Det3(SubItemIndex)
                    ReDim Lst_Files_Det4(SubItemIndex)
                    ReDim Lst_Files_Det5(SubItemIndex)
                    ReDim Lst_Files_Det6(SubItemIndex)
                    ReDim Lst_Files_Det7(SubItemIndex)
                    ReDim Lst_Files_Det8(SubItemIndex)
                End If


            End Try
            Text_gamma.Visible = False
            chk_external_ok.Visible = False
            gamma_ok = False
            nb_gamma = 0
            glob_gamma_mode = False

            If chk_external_ok.Checked = True Then
                For Each file As String In Directory.GetFiles(folder, "*.csv") 'Get Files In Folder
                    FileExtension = System.IO.Path.GetExtension(file) 'Get File Extension(s)
                    If LCase(FileExtension) = ".csv" Then 'Or LCase(FileExtension) = ".xlsx" Then
                        If LCase(System.IO.Path.GetFileNameWithoutExtension(file)) = "external-conc" Then
                            Text_gamma.Text = System.IO.Path.GetFileName(file) & " found"
                            Text_gamma.Visible = True
                            path_gamma = file
                            chk_external_ok.Visible = True

                            read_gamma_name_csv()

                        End If
                    End If

                Next

            Else
                gamma_mode = False
                gamma_ok = False
                nb_gamma = 0

                Text_gamma.Text = "No external elements for this processing"
                ReDim info_gamma_z(1)
                ReDim info_gamma_name(1)
                ReDim ext_tech(1)
            End If

        End If


    End Sub


    Sub Maj_Files_Trc(Det As String, ext_trc As String)
        Dim FileExtension As String 'Stores File Extension
        Dim SubItemIndex As Integer 'Sub Item Counter
        Dim pos1
        Dim ext_trc_is_ext As Boolean
        Dim str_filter As String

        Dim folder As String
        ext_trc_is_ext = False

        Try
            folder = CStr(trvFolders.SelectedNode.Tag) 'Folder Name
        Catch ex As Exception
            folder = "c:\"
        End Try

        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

            Try
                ''''''''''''''''********************************************************************************************** MATRICE DATA FILES
                For Each file As String In Directory.GetFiles(folder, "*" & ext_trc & "*") 'Get Files In Folder
                    FileExtension = System.IO.Path.GetExtension(file) 'Get File Extension(s)
                    FileExtension = Strings.Right(FileExtension, Len(FileExtension) - 1)
                    pos1 = InStr(FileExtension, ext_trc, CompareMethod.Text)
                    If pos1 > 0 Then
                        ext_trc_is_ext = True
                        Exit For
                    End If

                Next
            Catch ex As Exception

            End Try
        End If


        If ext_trc_is_ext Then ' Cas ou le mot clef est l'extention du fichier 'ex: X1
            If Box_txtFiltre.Text = "*" Or Box_txtFiltre.Text = "" Then
                str_filter = "*." & ext_trc
            Else
                str_filter = "*" & Box_txtFiltre.Text & "*." & ext_trc
            End If


        Else
            str_filter = "*" & ext_trc & "*"   'Cas ou le mot clef n'est pas l'extention mais un vrai mot clef (ex: ROMA)
            Box_txtFiltre.Text = "*"
        End If


        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

            Try
                ''''''''''''''''********************************************************************************************** MATRICE DATA FILES
                For Each file As String In Directory.GetFiles(folder, str_filter) 'Get Files In Folder
                    FileExtension = System.IO.Path.GetExtension(file) 'Get File Extension(s)

                    If LCase(FileExtension) <> ".par" And LCase(FileExtension) <> ".ini" Then
                        Select Case Det
                            Case "det0"
                                Lst_Files_Det0(SubItemIndex) = System.IO.Path.GetFileName(file)
                            Case "det1"
                                Lst_Files_Det1(SubItemIndex) = System.IO.Path.GetFileName(file)
                            Case "det2"
                                Lst_Files_Det2(SubItemIndex) = System.IO.Path.GetFileName(file)
                            Case "det3"
                                Lst_Files_Det3(SubItemIndex) = System.IO.Path.GetFileName(file)
                            Case "det4"
                                Lst_Files_Det4(SubItemIndex) = System.IO.Path.GetFileName(file)
                            Case "det5"
                                Lst_Files_Det5(SubItemIndex) = System.IO.Path.GetFileName(file)
                            Case "det6"
                                Lst_Files_Det6(SubItemIndex) = System.IO.Path.GetFileName(file)
                            Case "det7"
                                Lst_Files_Det7(SubItemIndex) = System.IO.Path.GetFileName(file)
                            Case "det8"
                                Lst_Files_Det8(SubItemIndex) = System.IO.Path.GetFileName(file)

                            Case Else
                                Exit Select
                        End Select

                        SubItemIndex += 1
                    End If

                Next

            Catch ex As Exception 'Something Went Wrong

                MessageBox.Show(ex.Message)
            Finally
                If SubItemIndex = 0 Then LvFiles.Items.Add("No files ...")
            End Try
        End If


    End Sub

    'Sub Maj_Files_Trc(Det As String, ext_trc As String) '########## BEFORE NOV 2023

    '    Dim FileExtension As String 'Stores File Extension
    '    Dim SubItemIndex As Integer 'Sub Item Counter
    '    ' LvFiles.Items.Clear()
    '    'Par_Mat.Text = ""

    '    Dim folder As String
    '    Try
    '        folder = CStr(trvFolders.SelectedNode.Tag) 'Folder Name
    '    Catch ex As Exception
    '        folder = "c:\"
    '    End Try


    '    If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

    '        Try
    '            ''''''''''''''''********************************************************************************************** MATRICE DATA FILES
    '            For Each file As String In Directory.GetFiles(folder, "*" & ext_trc & "*") 'Get Files In Folder
    '                FileExtension = Path.GetExtension(file) 'Get File Extension(s)

    '                If LCase(FileExtension) <> ".par" Then
    '                    Select Case Det
    '                        'case 0
    '                        '    par_mat.text = lstpar_trc.selecteditem
    '                        Case "det0"
    '                            Lst_Files_Det0(SubItemIndex) = Path.GetFileName(file)
    '                        Case "det1"
    '                            Lst_Files_Det1(SubItemIndex) = Path.GetFileName(file)
    '                        Case "det2"
    '                            Lst_Files_Det2(SubItemIndex) = Path.GetFileName(file)
    '                        Case "det3"
    '                            Lst_Files_Det3(SubItemIndex) = Path.GetFileName(file)
    '                        Case "det4"
    '                            Lst_Files_Det4(SubItemIndex) = Path.GetFileName(file)
    '                        Case "det5"
    '                            Lst_Files_Det5(SubItemIndex) = Path.GetFileName(file)
    '                        Case "det6"
    '                            Lst_Files_Det6(SubItemIndex) = Path.GetFileName(file)
    '                        Case "det7"
    '                            Lst_Files_Det7(SubItemIndex) = Path.GetFileName(file)
    '                        Case "det8"
    '                            Lst_Files_Det8(SubItemIndex) = Path.GetFileName(file)

    '                        Case Else
    '                            Exit Select
    '                    End Select


    '                    'LvFiles.Items.Add(Path.GetFileNameWithoutExtension(file))
    '                    SubItemIndex += 1
    '                End If

    '            Next

    '        Catch ex As Exception 'Something Went Wrong

    '            MessageBox.Show(ex.Message)
    '        Finally
    '            If SubItemIndex = 0 Then LvFiles.Items.Add("No files ...")
    '        End Try
    '    End If


    'End Sub

    Sub List_Par_Files_Trc()

        Dim FileExtension As String 'Stores File Extension
        Dim SubItemIndex As Integer 'Sub Item Counter
        Dim i As Integer
        Dim Name As String

        LstPar_Trc.Items.Clear()
        'Adjust_Filter.Enabled = False
        Check_det0.Enabled = False
        Check_det1.Enabled = False
        Check_det2.Enabled = False
        Check_det3.Enabled = False
        Check_det4.Enabled = False
        Check_det5.Enabled = False
        Check_det6.Enabled = False
        Check_det7.Enabled = False
        Check_det8.Enabled = False

        Check_det0.Checked = False
        Check_det1.Checked = False
        Check_det2.Checked = False
        Check_det3.Checked = False
        Check_det4.Checked = False
        Check_det5.Checked = False
        Check_det6.Checked = False
        Check_det7.Checked = False
        Check_det8.Checked = False
        Par_det0.Text = ""


        Check_det0.BackColor = System.Drawing.Color.LightGray
        Check_det1.BackColor = System.Drawing.Color.LightGray
        Check_det2.BackColor = System.Drawing.Color.LightGray
        Check_det3.BackColor = System.Drawing.Color.LightGray
        Check_det4.BackColor = System.Drawing.Color.LightGray
        Check_det5.BackColor = System.Drawing.Color.LightGray
        Check_det6.BackColor = System.Drawing.Color.LightGray
        Check_det7.BackColor = System.Drawing.Color.LightGray
        Check_det8.BackColor = System.Drawing.Color.LightGray





        FileExtension = "*.par"
        Dim folder As String
        Try
            folder = CStr(trvFolders.SelectedNode.Tag) 'Folder Name
        Catch ex As Exception
            folder = "c:\"
        End Try


        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then
            Try
                For i = 0 To 8
                    Select Case i
                        Case 0
                            If Ext_Trc0 <> "" And Ext_Trc0 <> Ext_Mat Then
                                FileExtension = "*" & Ext_Trc0 & ".par" '"*HE1.par"
                            Else
                                FileExtension = "NoTrcDefined"
                            End If

                        Case 1
                            If Ext_Trc1 <> "" And Ext_Trc1 <> Ext_Mat Then
                                FileExtension = "*" & Ext_Trc1 & ".par" '"*HE1.par"
                            Else
                                FileExtension = "NoTrcDefined"
                            End If

                        Case 2
                            If Ext_Trc2 <> "" And Ext_Trc2 <> Ext_Mat Then
                                FileExtension = "*" & Ext_Trc2 & ".par" '"*HE1.par"
                            Else
                                FileExtension = "NoTrcDefined"
                            End If
                            '"*HE2*.par"
                        Case 3
                            If Ext_Trc3 <> "" And Ext_Trc3 <> Ext_Mat Then
                                FileExtension = "*" & Ext_Trc3 & ".par" '"*HE1.par"
                            Else
                                FileExtension = "NoTrcDefined"
                            End If
                        Case 4
                            If Ext_Trc4 <> "" And Ext_Trc4 <> Ext_Mat Then
                                FileExtension = "*" & Ext_Trc4 & ".par" '"*HE1.par"
                            Else
                                FileExtension = "NoTrcDefined"
                            End If
                        Case 5
                            If Ext_Trc5 <> "" And Ext_Trc5 <> Ext_Mat Then
                                FileExtension = "*" & Ext_Trc5 & ".par" '"*HE1.par"
                            Else
                                FileExtension = "NoTrcDefined"
                            End If
                        Case 6
                            If Ext_Trc6 <> "" And Ext_Trc6 <> Ext_Mat Then
                                FileExtension = "*" & Ext_Trc6 & ".par" '"*HE1.par"
                            Else
                                FileExtension = "NoTrcDefined"
                            End If
                        Case 7
                            If Ext_Trc7 <> "" And Ext_Trc7 <> Ext_Mat Then
                                FileExtension = "*" & Ext_Trc7 & ".par" '"*HE1.par"
                            Else
                                FileExtension = "NoTrcDefined"
                            End If
                        Case 8
                            If Ext_Trc8 <> "" And Ext_Trc8 <> Ext_Mat Then
                                FileExtension = "*" & Ext_Trc8 & ".par" '"*HE1.par"
                            Else
                                FileExtension = "NoTrcDefined"
                            End If

                    End Select


                    For Each file As String In Directory.GetFiles(folder, FileExtension) 'Get Files In Folder
                        Name = System.IO.Path.GetFileName(file) 'Get File Extension(s)
                        LstPar_Trc.Items.Add(System.IO.Path.GetFileName(file))
                        SubItemIndex += 1
                    Next

                    Select Case i
                        Case 0
                            If SubItemIndex <> 0 Then
                                Check_det0.Checked = False
                                Check_det0.Enabled = True
                                Check_det0.BackColor = System.Drawing.Color.DarkGray

                            End If

                        Case 1
                            If SubItemIndex <> 0 Then
                                Check_det1.Checked = False
                                Check_det1.Enabled = True
                                Check_det1.BackColor = System.Drawing.Color.DarkGray

                            End If
                        Case 2
                            If SubItemIndex <> 0 Then
                                Check_det2.Checked = False
                                Check_det2.Enabled = True
                                Check_det2.BackColor = System.Drawing.Color.DarkGray

                            End If
                        Case 3
                            If SubItemIndex <> 0 Then
                                Check_det3.Checked = False
                                Check_det3.Enabled = True
                                Check_det3.BackColor = System.Drawing.Color.DarkGray

                            End If

                        Case 4
                            If SubItemIndex <> 0 Then
                                Check_det4.Checked = False
                                Check_det4.Enabled = True
                                Check_det4.BackColor = System.Drawing.Color.DarkGray

                            End If
                        Case 5
                            If SubItemIndex <> 0 Then
                                Check_det5.Checked = False
                                Check_det5.Enabled = True
                                Check_det5.BackColor = System.Drawing.Color.DarkGray

                            End If
                        Case 6
                            If SubItemIndex <> 0 Then
                                Check_det6.Checked = False
                                Check_det6.Enabled = True
                                Check_det6.Visible = True
                                Check_det6.BackColor = System.Drawing.Color.DarkGray

                            End If
                        Case 7
                            If SubItemIndex <> 0 Then
                                Check_det7.Checked = False
                                Check_det7.Enabled = True
                                Check_det7.BackColor = System.Drawing.Color.DarkGray

                            End If
                        Case 8
                            If SubItemIndex <> 0 Then
                                Check_det8.Checked = False
                                Check_det8.Enabled = True
                                Check_det8.BackColor = System.Drawing.Color.DarkGray

                            End If

                    End Select
                    SubItemIndex = 0
                Next i
            Catch ex As Exception 'Something Went Wrong
                MessageBox.Show(ex.Message)
            Finally
                If SubItemIndex = 0 Then LstPar_Trc.Items.Add("No files ...")
            End Try


            If CbDetMat.Text = Check_det0.Text Then
                Check_det0.Enabled = False
                Check_det0.Checked = False
            End If

            If CbDetMat.Text = Check_det1.Text Then
                Check_det1.Enabled = False
                Check_det1.Checked = False
            End If

            If CbDetMat.Text = Check_det2.Text Then
                Check_det1.Enabled = False
                Check_det1.Checked = False
            End If

            If CbDetMat.Text = Check_det3.Text Then
                Check_det3.Enabled = False
                Check_det3.Checked = False
            End If

            If CbDetMat.Text = Check_det4.Text Then
                Check_det4.Enabled = False
                Check_det4.Checked = False
            End If

            If CbDetMat.Text = Check_det5.Text Then
                Check_det5.Enabled = False
                Check_det5.Checked = False
            End If

            If CbDetMat.Text = Check_det6.Text Then
                Check_det6.Enabled = False
                Check_det6.Checked = False
            End If

            If CbDetMat.Text = Check_det7.Text Then
                Check_det7.Enabled = False
                Check_det7.Checked = False
            End If

            If CbDetMat.Text = Check_det8.Text Then
                Check_det8.Enabled = False
                Check_det8.Checked = False
            End If

        End If


    End Sub

    Sub Maj_HDF5_Files()

        Dim FileExtension As String 'Stores File Extension
        Dim SubItemIndex As Integer 'Sub Item Counter
        ListBox_HDF5.Items.Clear()


        Dim folder As String = CStr(trvFolders.SelectedNode.Tag) 'Folder Name

        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

            Try

                For Each file As String In Directory.GetFiles(folder, "*.h*5") 'Get Files In Folder
                    FileExtension = System.IO.Path.GetExtension(file) 'Get File Extension(s)
                    ListBox_HDF5.Items.Add(System.IO.Path.GetFileName(file))
                    SubItemIndex += 1
                Next
                If SubItemIndex = 1 Then
                    'TxtBox_HDF5_File.Text = ListBox_HDF5.Items(0)
                    'List_HDF5_group(Chemin_Data + "\" + TxtBox_HDF5_File.Text)
                End If
            Catch ex As Exception 'Something Went Wrong
                MessageBox.Show(ex.Message)
            Finally
                If SubItemIndex = 0 Then ListBox_HDF5.Items.Add("No files ...")
            End Try

        End If


    End Sub


    Sub Maj_Par_Files_Mat()

        Dim FileExtension As String 'Stores File Extension
        Dim SubItemIndex As Integer 'Sub Item Counter
        LstPar_Mat.Items.Clear()

        Dim folder As String
        folder = CStr(trvFolders.SelectedNode.Tag) 'Folder Name

        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

            Try

                For Each file As String In Directory.GetFiles(folder, Ext_Par_Mat) 'Get Files In Folder
                    FileExtension = System.IO.Path.GetExtension(file) 'Get File Extension(s)
                    LstPar_Mat.Items.Add(System.IO.Path.GetFileName(file))
                    SubItemIndex += 1
                Next
                If SubItemIndex = 1 Then
                    Par_Mat.Text = LstPar_Mat.Items(0)
                    ComboBox_Type_F.Items.Insert(0, Me.CbDetMat.Text)
                    ComboBox_Type_F.SelectedIndex = 0
                Else
                    Par_Mat.Text = ""
                    ' hdf5_mode = False
                End If

            Catch ex As Exception 'Something Went Wrong
                MessageBox.Show(ex.Message)
            Finally
                If SubItemIndex = 0 Then LstPar_Mat.Items.Add("No files ...")
            End Try

        End If


    End Sub


    Sub Maj_Par_Files_Trc(Trc_Select As Object, Det As String)

        Dim FileExtension As String 'Stores File Extension
        Dim SubItemIndex As Integer 'Sub Item Counter
        LstPar_Trc.Items.Clear()
        FileExtension = ""
        Dim folder As String
        Try
            folder = CStr(trvFolders.SelectedNode.Tag) 'Folder Name
        Catch ex As Exception
            folder = "c:\"
        End Try

        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

            Try

                For Each file As String In Directory.GetFiles(folder, Ext_Par_Trc) 'Get Files In Folder
                    FileExtension = System.IO.Path.GetExtension(file) 'Get File Extension(s)
                    LstPar_Trc.Items.Add(System.IO.Path.GetFileName(file))
                    SubItemIndex += 1
                Next
                If SubItemIndex = 1 Then Trc_Select.text = LstPar_Trc.Items(0)
            Catch ex As Exception 'Something Went Wrong
                MessageBox.Show(ex.Message)
            Finally
                If SubItemIndex = 0 Then LstPar_Trc.Items.Add("No files ...")
            End Try

            If FileExtension <> "" Then
                Select Case Det
                    'case 0
                    '    par_mat.text = lstpar_trc.selecteditem
                    Case "det0"
                        Ext_det0 = FileExtension
                    Case "det1"
                        Ext_det1 = FileExtension
                    Case "det2"
                        Ext_det2 = FileExtension
                    Case "det3"
                        Ext_det3 = FileExtension
                    Case "det4"
                        Ext_det4 = FileExtension
                    Case "det5"
                        Ext_det5 = FileExtension
                    Case "det6"
                        Ext_det6 = FileExtension
                    Case "det7"
                        Ext_det7 = FileExtension
                    Case "det8"
                        Ext_det8 = FileExtension

                    Case Else
                        Exit Select
                End Select
            End If
        End If


    End Sub

    Private Sub CbDetMat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CbDetMat.SelectedIndexChanged

        Select_Par_files = 0
        Par_Mat.Text = ""


        If ComboBox_Type_Calc.Text = "" Then ComboBox_Type_Calc.Text = "Ponctual" ''Init
        If CbDetMat.Text = Nothing Then
            Ext_Mat = "*"
            Ext_Par_Mat = "*.par"
        Else
            Ext_Mat = CbDetMat.Text
            Ext_Par_Mat = "*" & CbDetMat.Text & "*.par"
            If Myinit = True Then
                Try
                    ComboBox_Type_F.Items.RemoveAt(0)
                Catch ex As Exception

                End Try

                ComboBox_Type_F.Items.Insert(0, Ext_Mat)
                ComboBox_Type_F.SelectedIndex = 0
                'Adjust_Filter.Enabled = True
            End If
        End If


        If Myinit = True And Ext_Mat <> "" Then
            Maj_Par_Files_Mat()
            If hdf5_mode = True Then
                List_HDF5_group(Chemin_hdf5)
            Else
                Maj_Files_Mat()
            End If
            List_Par_Files_Trc()
        ElseIf Ext_Mat <> "*" Then
            List_Par_Files_Trc()
            LstPar_Mat.Items.Clear()
            List_Par_Files_Trc()
        End If

        If Myinit = False Then

        End If

    End Sub

    Private Sub LstPar_Trc_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles LstPar_Trc.MouseDoubleClick

        Select Case Select_Par_files
            Case 0
                Par_det0.Text = LstPar_Trc.SelectedItem
                Pivot_det0.Focus()
            Case 1
                Par_det1.Text = LstPar_Trc.SelectedItem
                Pivot_det1.Focus()
            Case 2
                Par_det2.Text = LstPar_Trc.SelectedItem
                Pivot_det2.Focus()
            Case 3
                Par_det3.Text = LstPar_Trc.SelectedItem
                Pivot_det3.Focus()
            Case 4
                Par_det4.Text = LstPar_Trc.SelectedItem
                Pivot_det4.Focus()
            Case 5
                Par_det5.Text = LstPar_Trc.SelectedItem
                Pivot_det5.Focus()
            Case 6
                Par_det6.Text = LstPar_Trc.SelectedItem
                Pivot_det6.Focus()
            Case 7
                Par_det7.Text = LstPar_Trc.SelectedItem
                Pivot_det7.Focus()
            Case 8
                Par_det8.Text = LstPar_Trc.SelectedItem
                Pivot_det8.Focus()
            Case Else
                Exit Select
        End Select

    End Sub

    Sub Load_config_exp_init()
        Dim str = ""
        Dim SubItemIndex As Integer
        Dim Ind1 As Integer
        Dim SplitText() As String
        Dim folder As String = CStr(trvFolders.SelectedNode.Tag)
        limite_conc_red_ok = 50000 'utlisé dans "calcul_best_conc" si conc. < 5% on prend la valeur RED

        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

            Try
                ''''''''''''''''********************************************************************************************** MATRICE DATA FILES

                For Each file As String In Directory.GetFiles(folder, "config-exp.ini") 'Get Files In Folder
                    SubItemIndex += 1
                Next

            Catch ex As Exception 'Something Went Wrong
                MessageBox.Show(ex.Message)
            Finally

                If SubItemIndex > 0 Then
                    Ind1 = 0
                    str = File.ReadAllText(Chemin_Data & "\config-exp.ini") 'str
                    SplitText = Split(str, vbCrLf)

                    For Each Str1 As String In SplitText
                        Select Case UCase(Str1)
                            Case "[TRACE-OXIDE]"
                                Text_Lst_Ox_Trc.Text = SplitText(Ind1 + 1)
                        End Select
                        Ind1 += 1
                    Next

                    str = Text_Lst_Ox_Trc.Text
                    Check_Trc_As_Oxy.Checked = True
                    If UCase(str) = "ALL TRACE AS OXIDE" Then
                        Ck_AllAsOxy.Checked = True
                        Check_Trc_As_Oxy.Checked = True
                        Check_Trc_As_Oxy.Enabled = False
                    ElseIf UCase(str) = "NO OXIDE" Then
                        Ck_AllAsOxy.Checked = False
                        Check_Trc_As_Oxy.Checked = False
                    Else

                        Check_Trc_As_Oxy.Checked = True
                    End If

                Else

                End If

            End Try
        End If

    End Sub

    Private Sub Adjust_Filter_Click(sender As Object, e As EventArgs) Handles Adjust_Filter.Click
        Dim ext As String
        Dim Nb_file As Integer
        Dim i As Integer
        Dim t As Integer
        Dim p As Integer
        Dim Filter_T As Single
        Dim Filter_From As Single
        Dim Filter_To As Single
        Dim Filter_Step As Single
        Dim pos_coma As Integer
        Dim Text_Pivot As String

        Fatal_Error = False
        ext = Mid(Ext_Mat, 2)
        Creer_tab_trc() ' Calcul Nb_Trc
        ReDim NomDet_Trc(Nb_Trc - 1)
        Nb_file = 0
        Adjust_Filter_B = True
        Ck_AllAsOxy.Checked = True
        Filter_From = Single.Parse(TextF_From.Text, USACulture) 'CSng(TextF_From.Text)
        Filter_To = Single.Parse(TextF_To.Text, USACulture) 'CSng(TextF_To.Text)
        Filter_Step = Single.Parse(TextF_Step.Text, USACulture) 'CSng()
        Nb_Process = ((Filter_To - Filter_From) / Filter_Step) + 1
        'Nb_Process = 10
        ReDim Error_Matrix(Nb_Process)
        ReDim Fichier_Matrix(Nb_Process) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace(Nb_Process)
        ReDim Fichier_Trace0(Nb_Process)
        ReDim Fichier_Trace1(Nb_Process) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace2(Nb_Process) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace3(Nb_Process) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace4(Nb_Process) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace5(Nb_Process) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace6(Nb_Process) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace7(Nb_Process) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace8(Nb_Process) 'LvFiles.SelectedItems.Count)
        ReDim Val_Charge_Trc(Nb_Process)
        ReDim Tab_Info_Mat.Z(50)
        ReDim Tab_Info_Mat.Elem(50)
        ReDim Tab_Info_Mat.Raie(50)
        ReDim Tab_Info_Mat.Inv(50)
        ' Dim Txt_Fichier_PAR_Mat_Filter
        ReDim Tab_Info_Mat.Nom_V2(50)
        ReDim Tab_Comment(Nb_Process)


        ext = Strings.Right(Ext_Mat, Len(Ext_Mat) - 1)

        For i = 0 To LvFiles.SelectedItems.Count - 1
            'Fichier_Matrix(Nb_file) = LvFiles.SelectedItems(I).Text + Ext_file_Mat '######## BEOFRE NOV 2023
            Fichier_Matrix(Nb_file) = LvFiles.SelectedItems(i).Text + Ext_det_mat
            If Check_det0.Checked = True Then Fichier_Trace0(Nb_file) = Lst_Files_Det0(LvFiles.SelectedIndices(i))
            If Check_det1.Checked = True Then Fichier_Trace1(Nb_file) = Lst_Files_Det1(LvFiles.SelectedIndices(i))
            If Check_det2.Checked = True Then Fichier_Trace2(Nb_file) = Lst_Files_Det2(LvFiles.SelectedIndices(i))
            If Check_det3.Checked = True Then Fichier_Trace3(Nb_file) = Lst_Files_Det3(LvFiles.SelectedIndices(i))
            If Check_det4.Checked = True Then Fichier_Trace4(Nb_file) = Lst_Files_Det4(LvFiles.SelectedIndices(i))
            If Check_det5.Checked = True Then Fichier_Trace5(Nb_file) = Lst_Files_Det5(LvFiles.SelectedIndices(i))
            If Check_det6.Checked = True Then Fichier_Trace6(Nb_file) = Lst_Files_Det6(LvFiles.SelectedIndices(i))
            If Check_det7.Checked = True Then Fichier_Trace7(Nb_file) = Lst_Files_Det7(LvFiles.SelectedIndices(i))
            If Check_det8.Checked = True Then Fichier_Trace8(Nb_file) = Lst_Files_Det8(LvFiles.SelectedIndices(i)) 'LvFiles.SelectedItems(I).Text + ".x13"

            Nb_file = Nb_file + 1
        Next i


        ReDim Txt_Fichier_PAR_Mat_Filter(Nb_Process)
        ReDim Txt_Fichier_PAR_Trc_Filter(Nb_Process)
        ReDim Chemin_GupixWin_Multi(Nb_Process)
        'Chemin_GupixWin = "c\gupixwin\gupix"
        ' Dim myDocuPath = Environ$("USERPROFILE") & "\Documents"
        Dim myDocuPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        'Chemin_GupixWin = myDocuPath & "\gupixwin\gupix"
        'Chemin_GupixWin = "c\gupixwin\gupix"
        List_Create_Swap()

        Text_Pivot = ""
        If Nb_Trc > 0 Then
            Calcul_With_Trc = True
        Else
            'ComboBox_Type_F.SelectedIndex = 0
        End If

        If Par_Mat.Text <> "" And Nb_file > 0 Then
            If Par_Mat.Text = "" Then OnlyTrace = True

            For t = 0 To Nb_Trc - 1
                Select Case Tab_Num_Trc(t)
                    Case 0
                        If Par_det0.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det0.Text
                        NomDet_Trc(t) = Check_det0.Text '"BE0"
                    Case 1
                        If Par_det1.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det1.Text
                        NomDet_Trc(t) = Check_det1.Text '"HE1"
                    Case 2
                        If Par_det2.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det2.Text
                        NomDet_Trc(t) = Check_det2.Text '"HE2"
                    Case 3
                        If Par_det3.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det3.Text
                        NomDet_Trc(t) = Check_det3.Text '"HE3"
                    Case 4
                        If Par_det4.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det4.Text
                        NomDet_Trc(t) = Check_det4.Text '"HE4"
                    Case 5 ' 1+2
                        If Par_det5.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det5.Text
                        NomDet_Trc(t) = Check_det5.Text '"HE11"
                    Case 6 ' All
                        If Par_det6.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det6.Text
                        NomDet_Trc(t) = Check_det6.Text '"HE10"
                    Case 7 '2+3
                        If Par_det7.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det7.Text
                        NomDet_Trc(t) = Check_det7.Text '"HE12"
                    Case 8
                        If Par_det8.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det8.Text
                        NomDet_Trc(t) = Check_det8.Text '"HE13"

                End Select

                Dim J = Tab_Num_Trc(t)
                pos_coma = 0

                If Text_Pivot <> "" Then
                    Dim Nb_pivot = 0
                    Do
                        pos_coma = InStr(pos_coma + 1, Text_Pivot, ",", vbTextCompare)
                        If pos_coma > 0 Then
                            Nb_pivot = Nb_pivot + 1
                            Tab_Pivot(t, Nb_pivot) = CInt(Mid(Text_Pivot, pos_coma + 1, 2)) 'Len(Pivot1(J).Text) - Pos_Coma))
                        Else
                            Tab_Pivot(t, 0) = CInt(Mid(Text_Pivot, 1, 2))
                            Nb_pivot = Nb_pivot + 1
                        End If
                    Loop While pos_coma > 0
                End If

            Next t
        End If


        If Par_Mat.Text <> "" Then
            NomDet_Mat = CStr(CbDetMat.SelectedIndex)
        Else
            Exit Sub
        End If

        If glob_gamma_mode = True Then
            'Read_gamma_xls()
            Load_gamma_csv()
        End If

        Filter_T = Filter_From
        If ComboBox_Type_F.Text = CbDetMat.Text Then

            For p = 0 To Nb_Process - 1
                Filter_Create_Par_Mat(TextF_Z.Text, Filter_T, p)
                Tab_Comment(p) = Strings.Replace(CStr(Filter_T), ",", ".")
                Filter_T = Math.Round(Filter_T + Filter_Step, 2)
            Next p
        Else
            For t = 0 To Nb_Trc - 1
                If ComboBox_Type_F.Text = NomDet_Trc(t) Then
                    For p = 0 To Nb_Process - 1
                        Filter_Create_Par_Trc(TextF_Z.Text, Filter_T, p, t) ' Crée N fois le même fichier Mat Par avec Z filtre = 1000 pour ne pas chercher un vrai Z Filter
                        Tab_Comment(p) = CStr(Filter_T)
                        Filter_T = Math.Round(Filter_T + Filter_Step, 2) 'CSng(TextF_Step.Text)
                        '     Filter_Create_Par_Trc("1000", Filter_T, i) ' Crée N fois le même fichier Mat Par
                    Next p
                End If
            Next
        End If

        TextF_Z.Enabled = False
        TextF_From.Enabled = False
        TextF_To.Enabled = False
        TextF_Step.Enabled = False
        Adjust_Filter.Enabled = False
        TabControl1.Enabled = False

        ProgressBar1.Maximum = Nb_Process + (Nb_Process * Nb_Trc * 3) + Nb_Process * 3 'Nb_Process * ((Nb_Trc * 2) + 1) * Nb_file
        ProgressBar1.Value = 0

        If Nb_file > 0 Then Main_Process_Adjust_Filter(Nb_file, Nb_Process)

        TextF_Z.Enabled = True
        TextF_From.Enabled = True
        TextF_To.Enabled = True
        TextF_Step.Enabled = True
        Adjust_Filter.Enabled = True
        TabControl1.Enabled = True
        ToolStripStatusLabel1.Text = "Calcul Finished"
        ' ResetAllTab()

    End Sub

    'Comm

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button2.Enabled = False
        Tps_Calc.Visible = True
        Run_Calcul_Ponctual()
        Button2.Enabled = True
        Tps_Calc.Visible = False
    End Sub

    Private Sub Button_Run2_Click(sender As Object, e As EventArgs) Handles Button_Run2.Click
        Button2.Enabled = False
        Tps_Calc.Visible = True
        Run_Calcul_Ponctual()
        Button2.Enabled = True
        Tps_Calc.Visible = False
    End Sub

    Private Sub Run_Calcul_Ponctual()
        Dim Nb_file As Integer
        Dim I As Integer
        IntNb_File = LvFiles.SelectedItems.Count


        If IntNb_File = 0 Then
            Try
                For I = 0 To LvFiles.Items.Count - 1
                    LvFiles.Items(I).Selected = True 'change to false to uncheck them
                Next I
                IntNb_File = LvFiles.SelectedItems.Count
            Catch ex As Exception
                MsgBox("Selected point to process", MsgBoxStyle.MsgBoxHelp, "Error")
                ToolStripStatusLabel1.Text = "Please select spectra first.."
                Exit Sub
            End Try

            'can be like this too

        End If

        ReDim Fichier_Matrix(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace0(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace1(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace2(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace3(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace4(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace5(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace6(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace7(LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace8(LvFiles.SelectedItems.Count)
        ReDim Tab_Info_Mat.Z(50)
        ReDim Tab_Info_Mat.Elem(50)
        ReDim Tab_Info_Mat.Raie(50)
        ReDim Tab_Info_Mat.Inv(50)
        ReDim Tab_Info_Mat.Nom_V2(50)


        Dim ext As String
        Dim Nb_Rep As Integer
        Dim pos_coma As Integer
        Dim Text_Pivot As String
        Dim Traupixe_Init As Boolean
        Dim File_Text1
        Dim SplitText() As String
        Dim Ret1 As Boolean

        Fatal_Error = False
        First_Init = False ' Permet lecture Absorbers 
        ReDim First_Init_Trc(10) 'Permet lecture Absorbers, remets tout False 

        TextXLS.Text = ""
        Creer_tab_trc()
        Text_Pivot = ""
        Chemin_Rapport = ""
        Adjust_Filter_B = False
        Calcul_With_Trc = False
        OnlyTrace = False
        xlApp = Nothing
        If Nb_Trc > 0 Then Calcul_With_Trc = True
        ReDim NomDet_Trc(Nb_Trc - 1)
        IntNb_File = LvFiles.SelectedItems.Count
        ToolStripStatusLabel1.Text = "Start Process"

        hdf5_mode = True 'Mode hdf5 par default


        Det_use_charge()

        If glob_gamma_mode = True Then
            'Read_gamma_xls()
            Load_gamma_csv()
        End If

        If TxtBox_HDF5_File.Text = "" Then
            hdf5_mode = False
        End If

        If Par_Mat.Text = "" Then OnlyTrace = True

        If IntNb_File > 0 Then

            For I = 0 To Nb_Trc - 1

                Select Case Tab_Num_Trc(I)
                    Case 0
                        If Par_det0.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det0.Text
                        NomDet_Trc(I) = Check_det0.Text '"BE0"
                    Case 1
                        If Par_det1.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det1.Text
                        NomDet_Trc(I) = Check_det1.Text '"HE1"
                    Case 2
                        If Par_det2.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det2.Text
                        NomDet_Trc(I) = Check_det2.Text '"HE2"
                    Case 3
                        If Par_det3.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det3.Text
                        NomDet_Trc(I) = Check_det3.Text '"HE3"
                    Case 4
                        If Par_det4.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det4.Text
                        NomDet_Trc(I) = Check_det4.Text '"HE4"
                    Case 5 ' 1+2
                        If Par_det5.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det5.Text
                        NomDet_Trc(I) = Check_det5.Text '"HE11"
                    Case 6 ' All
                        If Par_det6.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det6.Text
                        NomDet_Trc(I) = Check_det6.Text '"HE10"
                    Case 7 '2+3
                        If Par_det7.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det7.Text
                        NomDet_Trc(I) = Check_det7.Text '"HE12"
                    Case 8
                        If Par_det8.Text = "" Then Exit Sub
                        Text_Pivot = Pivot_det8.Text
                        NomDet_Trc(I) = Check_det8.Text '"HE13"

                End Select

                Dim J = Tab_Num_Trc(I)
                pos_coma = 0

                If Text_Pivot <> "" And Text_Pivot <> "Q-File" Then
                    Dim Nb_pivot = 0
                    Do
                        pos_coma = InStr(pos_coma + 1, Text_Pivot, ",", vbTextCompare)
                        If pos_coma > 0 Then
                            Nb_pivot = Nb_pivot + 1
                            Tab_Pivot(I, Nb_pivot) = CInt(Mid(Text_Pivot, pos_coma + 1, 2)) 'Len(Pivot1(J).Text) - Pos_Coma))
                        Else
                            Try
                                Tab_Pivot(I, 0) = CInt(Mid(Text_Pivot, 1, 2))
                                Nb_pivot = Nb_pivot + 1

                            Catch ex As Exception
                                MsgBox("Pivot definition error must be Z number (e.g. ""26"") ", MsgBoxStyle.MsgBoxHelp, "Error")
                                ToolStripStatusLabel1.Text = "Pivot definition error must be Z number (e.g. ""26"") "
                                Exit Sub
                            End Try

                        End If
                    Loop While pos_coma > 0
                End If

            Next I
        End If

        If Par_Mat.Text <> "" Then
            NomDet_Mat = CStr(CbDetMat.SelectedIndex)
        ElseIf Nb_Trc > 0 Then
            OnlyTrace = True
        Else
            MsgBox("No paramater file selected !", MsgBoxStyle.MsgBoxHelp, "Error")
            ToolStripStatusLabel1.Text = "No paramater file selected !"
            Exit Sub
        End If
        ToolStripStatusLabel1.Text = "Update/Create Config-exp.ini"
        Traupixe_Init = File.Exists(Chemin_Data & "\config-exp.ini")

        If Text_Lst_Ox_Trc.Text = "" Or Text_Lst_Ox_Trc.Text = "No oxide" Then Check_Trc_As_Oxy.Checked = False

        'If Check_Trc_As_Oxy.Checked = True Or Ck_AllAsOxy.Checked = True Then
        If Traupixe_Init = False Then
            File.WriteAllText(Chemin_Data & "\config-exp.ini", "[Trace-oxide]" & vbCrLf & Text_Lst_Ox_Trc.Text)
        Else
            File_Text1 = File.ReadAllText(Chemin_Data & "\config-exp.ini")
        End If

        If Text_Lst_Ox_Trc.Text <> "" Then
            File.WriteAllText(Chemin_Data & "\config-exp.ini", "[Trace-oxide]" & vbCrLf & Text_Lst_Ox_Trc.Text)
        Else
            File.WriteAllText(Chemin_Data & "\config-exp.ini", "[Trace-oxide]" & vbCrLf & "NO OXIDE")
        End If

        '########################################## TRACE AS oxide IN EXCEL SHEET 100% ,ppm , S_100 et S_ppm
        If Nb_Trc > 0 Then
            If Check_Trc_As_Oxy.Checked = True And Ck_AllAsOxy.Checked = True Then
                Nom_Excel_Trx_O = "_All-Trc-as-oxide" '& Strings.Replace(Text_Lst_Ox_Trc.Text, ",", "_")
                Lecture_Fichier_Par_Trc(2, 0)
            ElseIf UCase(Text_Lst_Ox_Trc.Text) <> "NO OXIDE" Then

                Nom_Excel_Trx_O = "_Elem-Ox_" & Strings.Replace(Text_Lst_Ox_Trc.Text, ",", "_")
                SplitText = Split(Text_Lst_Ox_Trc.Text, ",")
                ReDim Tab_Trc_as_Oxy(UBound(SplitText))

                For I = 0 To UBound(SplitText)
                    Try
                        Tab_Trc_as_Oxy(I) = CInt(SplitText(I))
                    Catch ex As Exception

                    End Try

                Next
            End If
        Else
            Nom_Excel_Trx_O = ""
        End If
        Nb_file = 0


        If Ext_Mat = "" Then Ext_Mat = "*.x0" ' Case Only Trace a=, on doir rajouter une ext pour autre fonction

        ext = Strings.Right(Ext_Mat, Len(Ext_Mat) - 1)
        ReDim tab_select_file_indices(LvFiles.SelectedItems.Count)
        For I = 0 To LvFiles.SelectedItems.Count - 1
            'Fichier_Matrix(Nb_file) = LvFiles.SelectedItems(I).Text + Ext_file_Mat '######## BEOFRE NOV 2023
            Fichier_Matrix(Nb_file) = LvFiles.SelectedItems(I).Text + Ext_det_mat
            If Check_det0.Checked = True Then Fichier_Trace0(Nb_file) = Lst_Files_Det0(LvFiles.SelectedIndices(I))
            If Check_det1.Checked = True Then Fichier_Trace1(Nb_file) = Lst_Files_Det1(LvFiles.SelectedIndices(I))
            If Check_det2.Checked = True Then Fichier_Trace2(Nb_file) = Lst_Files_Det2(LvFiles.SelectedIndices(I))
            If Check_det3.Checked = True Then Fichier_Trace3(Nb_file) = Lst_Files_Det3(LvFiles.SelectedIndices(I))
            If Check_det4.Checked = True Then Fichier_Trace4(Nb_file) = Lst_Files_Det4(LvFiles.SelectedIndices(I))
            If Check_det5.Checked = True Then Fichier_Trace5(Nb_file) = Lst_Files_Det5(LvFiles.SelectedIndices(I))
            If Check_det6.Checked = True Then Fichier_Trace6(Nb_file) = Lst_Files_Det6(LvFiles.SelectedIndices(I))
            If Check_det7.Checked = True Then Fichier_Trace7(Nb_file) = Lst_Files_Det7(LvFiles.SelectedIndices(I))
            If Check_det8.Checked = True Then Fichier_Trace8(Nb_file) = Lst_Files_Det8(LvFiles.SelectedIndices(I)) 'LvFiles.SelectedItems(I).Text + ".x13"
            tab_select_file_indices(I) = LvFiles.SelectedIndices(I)
            Nb_file = Nb_file + 1
        Next I
        ReDim Preserve tab_select_file_indices(Nb_file)


        If Check_det0.Checked = False Then ReDim Fichier_Trace0(0)
        If Check_det1.Checked = False Then ReDim Fichier_Trace1(0)
        If Check_det2.Checked = False Then ReDim Fichier_Trace2(0)
        If Check_det3.Checked = False Then ReDim Fichier_Trace3(0)
        If Check_det4.Checked = False Then ReDim Fichier_Trace4(0)
        If Check_det5.Checked = False Then ReDim Fichier_Trace5(0)
        If Check_det6.Checked = False Then ReDim Fichier_Trace6(0)
        If Check_det7.Checked = False Then ReDim Fichier_Trace7(0)
        If Check_det8.Checked = False Then ReDim Fichier_Trace8(0)


        If Nb_file = 0 Then Exit Sub

        If LvFiles.SelectedItems.Count > Environment.ProcessorCount + 1 Then
            Nb_Process = Environment.ProcessorCount  '25
            Nb_Proc.Text = CStr(Nb_Process)
        Else
            Nb_Process = LvFiles.SelectedItems.Count
            Nb_Proc.Text = CStr(Nb_Process)
        End If

        ReDim Val_Charge_Trc(Nb_Process)
        ReDim Chemin_GupixWin_Multi(Nb_file) 'As String

        '###########   Mat + (Trc) + Excel*4
        ProgressBar1.Maximum = Nb_file + (Nb_file * Nb_Trc * 3) + Nb_file * 3
        ProgressBar1.Value = 0

        'Chemin_GupixWin = "C\gupixwin\gupix"
        'Dim userProfile = %userprofile% '.ExpandEnvironmentStrings("")
        'Dim myDocuPath = Environ$("USERPROFILE") & "\My Documents"
        Dim ComputerName As String
        ComputerName = System.Net.Dns.GetHostName

        If ComputerName = "server-aglae" Then
            Dim myDocuPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Chemin_GupixWin = myDocuPath & "\gupixwin\gupix"
        Else
            ' Chemin_GupixWin = "c\gupixwin\gupix"
            '########### Lue dans fichier de config 
        End If

        LvFiles.Enabled = False
        If Nb_file < Nb_Process Then
            Nb_Process = Nb_file
        End If
        Nb_Rep = Nb_Process

        'Nb_Swap = List_Create_Swap()
        IntNb_File = Nb_file
        'If (Depouillement(IntNb_File) = True) Then
        If Fatal_Error = True Then
            MsgBox("Fatal Error, please check spectra file And parameter file are present into the data folder", MsgBoxStyle.Information, "Fatal Error")
            Exit Sub
        End If
        Ret1 = Main_Process_Ponctual(IntNb_File)

        If Ret1 = True Then
            '  Excel_Save(-1)
            Excel_Close()
            LvFiles.Enabled = True
            ''''''''''''''''''''''''''''''MsgBox("Excel file saved in folder   " + vbCrLf + trvFolders.SelectedNode.Tag + "\", vbInformation)
        ElseIf Fatal_Error = True Then
            Exit Sub
        Else

            MsgBox("Please,you must close the Excel file", MsgBoxStyle.Information, "Error Excel file already open")
            Exit Sub
        End If '

        For I = 0 To Nb_Rep - 1
            Try
                Kill(Chemin_GupixWin_Multi(I) & "\*.*")
                RmDir(Chemin_GupixWin_Multi(I)) '& "\gupix"
                IntNb_File = Nb_file
            Catch ex As Exception

            End Try

        Next I

        For I = 0 To Global_Nb_Swap - 1
            Try
                RmDir(Tab_Swap(I)) '& "\gupix"
            Catch ex As Exception

            End Try

        Next

        Button2.Enabled = True
    End Sub


    Public Function Main_Process_Adjust_Filter(nb_file As Integer, Nb_Process As Integer) As Boolean
        Dim i As Integer
        Dim fs
        Dim Num_Fichier As Integer
        Dim J As Integer
        Dim nb_loop As Integer
        Dim start_calc
        Dim Thread_Data_Excel_Alive(100) As Boolean
        Dim Tab_Inc_Done(100) As Boolean
        Dim Parametres_All_Thread As Struct_Parametres_Thread
        Dim nb_process_main As Integer
        Dim nb_elem As Integer


        Nb_Oxyde_X2 = 0
        Nb_Oxyde_Mat = 0
        IntNb_File = nb_file
        ProgressBar1.Visible = True
        fs = CreateObject("Scripting.FileSystemObject")

        ReDim Tab_Matrix(Nb_Process)
        ReDim Tab_Val_Oxyde_Mat(Nb_Process)
        ReDim Tab_Val_Oxyde_Trc(Nb_Trc, Nb_Process)
        ReDim Tab_Val_Trc(Nb_Trc, Nb_Process)
        ReDim Tab_Val_Mat(Nb_Process)


        For p = 0 To nb_file
            ReDim Tab_Val_Mat(p).Val_Mtx(50)
            ReDim Tab_Val_Mat(p).Depth(50)
            ReDim Tab_Val_Mat(p).Selected_Pivot(50)
            ReDim Tab_Val_Mat(p).ConcOK(50)
        Next

        ReDim Info_Oxyde_Mat.nom(50)
        ReDim Info_Oxyde_Mat.Z(50)


        If Nb_Trc > 0 Then
            ReDim Tab_Info_Oxyde_Trc(Nb_Trc)
            ReDim Tab_Info_Trc(Nb_Trc)
            ReDim Nb_Elements_Trc(Nb_Trc)

            For p = 0 To Nb_Trc
                ReDim Tab_Info_Oxyde_Trc(p).nom(50)
                ReDim Tab_Info_Oxyde_Trc(p).Z(50)
                ReDim Tab_Info_Trc(p).Elem(50)
                ReDim Tab_Info_Trc(p).Inv(50)
                ReDim Tab_Info_Trc(p).Raie(50)
                ReDim Tab_Info_Trc(p).Z(50)
                ReDim Tab_Info_Trc(p).Nom_V2(50)

                For q = 0 To Nb_Process
                    ReDim Tab_Val_Trc(p, q).Trans(50)
                    ReDim Tab_Val_Trc(p, q).Depth(50)
                Next

            Next

        End If

        Lecture_Fichier_Par_Mat()

        Progress.Text = Trim$(Str$(0)) & "/" & Trim$(Str$(nb_file))
        start_calc = My.Computer.Clock.TickCount
        If Fatal_Error = True Then
            MsgBox("Fatal Error, please check Matrix parameter file Is present into the data folder", MsgBoxStyle.Information, "Fatal error readind parameter Matrix file")
            Exit Function
        End If

        '###########################calcul temps restant

        Num_Fichier = 0 '((J - 1) * Nb_Process1)
        Chemin_Rapport = ""
        For J = 1 To nb_file  '4############################# Main BOUCLE
            ReDim Tab_Name_File(Nb_Process - 1)
            'ReDim Tab_Comment(Nb_Process - 1)
            ReDim Tab_IsPonctuel(Nb_Process - 1)
            ReDim Info_Experience_Mat(Nb_Process - 1)

            ReDim Info_Mat_Raie(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Height(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Area(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Conc(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Oxyde(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_LOD(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Stat_Error(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Fit_Error(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Total_Error(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Final_Error(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Y_N_Q(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Str_Mat_Conc_100(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Conc_ppm(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Conc_RED(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Best_Yes(Nb_Process - 1, Nb_Elements_Mat - 1)

            ReDim Val_Mat_Best_Yes_RED(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim tab_select_file_indices(Nb_Process - 1)

            Num_Fichier = J - 1
            '######################################################################## SEQUENCE MATRICE ######################################
            'If ComboBox_Type_F.Text = "Trace" Then
            For l = 0 To Nb_Process - 1
                Fichier_Matrix(l) = Fichier_Matrix(0)
                tab_select_file_indices(l) = LvFiles.SelectedIndices(0)
            Next l

            If ComboBox_Type_F.Text = CbDetMat.Text Then
                Sequence_Matrix_Multi_Thread(0, 0, Nb_Process)
            Else
                nb_process_main = Nb_Process
                Nb_Process = 1
                Sequence_Matrix_Multi_Thread(0, 0, 1) ' 1 Process
                Nb_Process = nb_process_main
                For i = 1 To Nb_Process - 1
                    Info_Experience_Mat(i) = Info_Experience_Mat(0)
                    Tab_Name_File(i) = Tab_Name_File(0)

                    My.Computer.FileSystem.CopyFile((Chemin_GupixWin_Multi(0) & "\PixMTX.out"), (Chemin_GupixWin_Multi(i) & "\PixMTX.out"), True)
                    'My.Computer.FileSystem.CopyFile((Chemin_GupixWin_Multi(0) & "\.out"), (Chemin_GupixWin_Multi(i) & "\PixMTX.out"), True)

                    For nb_elem = 0 To Nb_Elements_Mat - 1
                        Val_Mat_Conc(i, nb_elem) = Val_Mat_Conc(0, nb_elem)
                        Val_Mat_Oxyde(i, nb_elem) = Val_Mat_Oxyde(0, nb_elem)
                        Val_Mat_LOD(i, nb_elem) = Val_Mat_LOD(0, nb_elem)
                        Val_Mat_Fit_Error(i, nb_elem) = Val_Mat_Fit_Error(0, nb_elem)
                        Val_Mat_Total_Error(i, nb_elem) = Val_Mat_Total_Error(0, nb_elem)
                        Val_Mat_Final_Error(i, nb_elem) = Val_Mat_Final_Error(0, nb_elem)
                        Val_Mat_Y_N_Q(i, nb_elem) = Val_Mat_Y_N_Q(0, nb_elem)
                        Str_Mat_Conc_100(i, nb_elem) = Str_Mat_Conc_100(0, nb_elem)
                        Val_Mat_Conc_ppm(i, nb_elem) = Val_Mat_Conc_ppm(0, nb_elem)
                        Val_Mat_Conc_RED(i, nb_elem) = Val_Mat_Conc_RED(0, nb_elem)
                        Val_Mat_Best_Yes(i, nb_elem) = Val_Mat_Best_Yes(0, nb_elem)
                    Next


                Next i
            End If



            ToolStripStatusLabel1.Text = "Matrix Sequence Finish"


            Nb_Trc_Total = 0

            If Nb_Trc > 0 Then

                For i = 0 To Nb_Trc - 1
                    copy_par_file_trc(i)
                    Lecture_par_trc_HED_NbElem(i) ' RECUPERE LES chemins fichiers HED ?
                Next i

                For i = 0 To Nb_Trc - 1
                    Nb_Trc_Total = Nb_Trc_Total + Nb_Elements_Trc(i)
                Next i

                Nb_Trc_Total = Nb_Trc_Total - 1
                ReDim Info_Experience_Trc(Nb_Process - 1, Nb_Trc)
                ReDim First_Init_Trc(10) 'Permet lecture Absorbers

                ReDim Info_Trc_Raie(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Height(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Area(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Conc(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Oxyde(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_LOD(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Stat_Error(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Fit_Error(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_WithPivot_Error(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Total_Error(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Y_N_Q(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Error_Pivot(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Conc100(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Conc_ppm(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Best_Yes(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Best_Yes_RED(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Pivot_Error(Nb_Process - 1, Nb_Trc)
                ReDim Error_Trace(Nb_Process - 1, Nb_Trc)

            End If

            If Fatal_Error = True Then Exit Function


            If Nb_Trc > 0 Then

                For i = 0 To Nb_Trc - 1
                    Sequence_Trace_Multi_Thread(Num_Fichier, nb_file, i)
                Next i

            End If


            If Chemin_Rapport = "" Then 'Then
                xlBook = Excel_Open(0, Chemin_Rapport)
            End If
            '  Excel_Save(J)
            'xlApp.visible = False
            On Error Resume Next

            Entete_100()

            Dim para_best_C As parametres_best_conc_Thread
            For l = 0 To Nb_Process - 1
                '''''Calcul_Final_Best_Conc_New_Thread(l + ((Nb_Process) * (J - 1)), l, Nb_Trc) ', nb_data_read
                If OnlyTrace = False Then
                    para_best_C.indx_file = tab_select_file_indices(l + ((Nb_Process1) * (J - 1)))
                    para_best_C.num_process = l
                    para_best_C.nb_trace = Nb_Trc
                    'Calcul_Final_Best_Conc_New_Thread(tab_select_file_indices(l + ((Nb_Process1) * (J - 1))), l, Nb_Trc) ', nb_data_read
                    Calcul_Final_Best_Conc_New_Thread(para_best_C) ', nb_data_read
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Else
                    Calcul_Final_Best_Only_Trace_Conc_New_Thread(l + ((Nb_Process1) * (J - 1)), l, Nb_Trc) ', nb_data_read
                    ProgressBar1.Value = ProgressBar1.Value + 1
                End If

                If Chk_RoundValue.Checked = True Then
                    Arrondi_Mat_Elem(l)
                    Arrondi_Mat_Oxyde(l)
                    Arrondi_Trc_Elem(l)
                    Arrondi_Trc_Oxyde(l)
                End If

            Next l

            Parametres_All_Thread.Num_File = Nb_Process1 * (J - 1)
            Parametres_All_Thread.Nb_Calcul = Nb_Process
            Tab_IsPonctuel_1 = Tab_IsPonctuel
            Tab_Name_File_1 = Tab_Name_File
            Tab_Comment_1 = Tab_Comment
            Info_Mat_Raie_1 = Info_Mat_Raie

            Val_Mat_Area_1 = Val_Mat_Area
            Val_Mat_Conc_1 = Val_Mat_Conc
            Val_Mat_LOD_1 = Val_Mat_LOD
            Val_Mat_Height_1 = Val_Mat_Height
            Val_Mat_Fit_Error_1 = Val_Mat_Fit_Error
            Val_Mat_Total_Error_1 = Val_Mat_Total_Error
            Val_Mat_Final_Error_1 = Val_Mat_Final_Error
            Val_Mat_Y_N_Q_1 = Val_Mat_Y_N_Q
            Val_Conc_S_100_1 = Val_Conc_S_100
            Val_Conc_S_ppm_1 = Val_Conc_S_ppm
            Val_Conc_S_RED_ppm_1 = Val_Conc_S_RED_ppm
            Val_Conc_S_RED100_1 = Val_Conc_S_RED100
            Val_Error_S_1 = Val_Error_S
            Val_Conc_And_Error_1 = Val_Conc_And_Error
            Val_Conc_And_Error100_1 = Val_Conc_And_Error100
            Val_Choix_S_1 = Val_Choix_S
            Val_Mat_Mtx_1 = Val_Mat_Mtx
            NomDet_Trc_1 = NomDet_Trc

            Val_Trc_Y_N_Q_1 = Val_Trc_Y_N_Q
            Val_Trc_Conc_1 = Val_Trc_Conc
            Val_Trc_LOD_1 = Val_Trc_LOD
            Val_Trc_Area_1 = Val_Trc_Area
            Val_Trc_Fit_Error_1 = Val_Trc_Fit_Error
            Val_Trc_WithPivot_Error_1 = Val_Trc_WithPivot_Error
            Val_Trc_Height_1 = Val_Trc_Height
            Val_YNQ_Final_1 = Val_YNQ_Final

            If mnuOxydeOUI.Checked = True Then
                Val_Mat_Oxyde_1 = Val_Mat_Oxyde
                Val_Trc_Oxyde_1 = Val_Trc_Oxyde
            End If
            Ecrire_Entete_Excel(0)
            Excel_Write_Filename_Comment(0, Nb_Process)
            Function_Data_To_LibreOffice(Parametres_All_Thread)
            ' Parametres_All_Thread.
            ' For l = 0 To Nb_Process - 1
            Function_Excel_Format_Style(Parametres_All_Thread)
            'Function_Excel_Format_Italic(Parametres_All_Thread)
            ' Next
            Excel_Legend(Nb_Process, True, False)
            Excel_Save(J)


        Next J


    End Function



    Public Function Main_Process_Ponctual(Nb_file As Integer) As Boolean
        Dim i As Integer
        Dim Num_Fichier As Integer
        Dim J, Reste
        Dim l
        Dim nb_loop As Integer
        Dim start_calc
        Dim Tps_1_Loop As Integer, Second As Integer, Minutes As Integer, MyH
        Dim Tps_Total As Integer
        Dim Thread_Data_Excel_Alive(100) As Boolean
        Dim Tab_Inc_Done(100) As Boolean
        Dim legend As Boolean
        Dim MyErase As Boolean
        Dim Parametres_All_Thread As Struct_Parametres_Thread
        Dim Tab_Thread_Data_Excel_Alive As Boolean
        Dim Tab_Thread_Format_Style_Excel_Alive As Boolean
        Dim Tab_Thread_Format_Italic_Excel_Alive As Boolean
        Dim Tab_Thread_best_conc_Alive(100) As Boolean
        Dim Set_Finish(Nb_Process) As Boolean
        Dim Format_Range_1 As Boolean

        Format_Range_1 = False
        Nb_Oxyde_X2 = 0
        Nb_Oxyde_Mat = 0
        Nb_Elements_Mat = 0
        ProgressBar1.Visible = True


        ReDim Tab_Matrix(Nb_file)
        ReDim Tab_Val_Oxyde_Mat(Nb_file)
        ReDim Tab_Val_Oxyde_Trc(Nb_Trc, Nb_file)
        ReDim Tab_Val_Trc(Nb_Trc, Nb_file)
        ReDim Tab_Val_Mat(Nb_file)

        For p = 0 To Nb_file
            ReDim Tab_Val_Mat(p).Val_Mtx(50)
            ReDim Tab_Val_Mat(p).Trans(50)
            ReDim Tab_Val_Mat(p).Depth(50)
            ReDim Tab_Val_Mat(p).Selected_Pivot(50)
            ReDim Tab_Val_Mat(p).ConcOK(50)
        Next

        ReDim Info_Oxyde_Mat.nom(50)
        ReDim Info_Oxyde_Mat.Z(50)


        If Nb_Trc > 0 Then
            ReDim Tab_Info_Oxyde_Trc(Nb_Trc)
            ReDim Tab_Info_Trc(Nb_Trc)
            ReDim Nb_Elements_Trc(Nb_Trc)

            For p = 0 To Nb_Trc
                ReDim Tab_Info_Oxyde_Trc(p).nom(50)
                ReDim Tab_Info_Oxyde_Trc(p).Z(50)
                ReDim Tab_Info_Trc(p).Elem(50)
                ReDim Tab_Info_Trc(p).Inv(50)
                ReDim Tab_Info_Trc(p).Raie(50)
                ReDim Tab_Info_Trc(p).Z(50)
                ReDim Tab_Info_Trc(p).Nom_V2(50)

                For q = 0 To Nb_Process
                    ReDim Tab_Val_Trc(p, q).Trans(50)
                    ReDim Tab_Val_Trc(p, q).Depth(50)
                Next
            Next
        End If


        Nb_Process = CInt(Nb_Proc.Text)
        If OnlyTrace = False Then Lecture_Fichier_Par_Mat()
        List_Create_Swap()
        ' Nb_Elements_Mat = Nb_Elements_Mat - 1

        Reste = Nb_file Mod Nb_Process
        nb_loop = Math.Truncate(Nb_file / Nb_Process)
        If Reste > 0 Then
            nb_loop = nb_loop + 1
        End If

        Progress.Text = Trim$(Str$(0)) & "/" & Trim$(Str$(IntNb_File))
        start_calc = My.Computer.Clock.TickCount

        ReDim Tab_IsPonctuel_1(Nb_Process - 1)

        'TEST SI LE FICHIER EST OUVERT SOUS EXCEL / LIBREOFFICE

        xlBook = Excel_Open(0, "")


        Try
            If xlBook.Author = Nothing Then
                Main_Process_Ponctual = False
                Exit Function
            End If
            Chemin_Rapport = ""

        Catch ex As Exception
            Main_Process_Ponctual = False
            Exit Function
        End Try


        '################### READ ALL SPECTRA DATASET AND ATTRIB IN HDF5 FILE HERE

        If hdf5_mode = True Then hdf5_Read_Dataset_Attrib()
        For i = 0 To Nb_Trc - 1
            copy_par_file_trc(i)
            Lecture_par_trc_HED_NbElem(i)
        Next

        For J = 1 To nb_loop  '4############################# Main BOUCLE

            '###########################calcul temps restant

            If J > 1 Then
                Second = Tps_1_Loop / 1000 * (nb_loop - (J - 1))
                Minutes = Int((Second Mod 3600) / 60)
                Second = Second - (Minutes * 60)

                MyH = IIf(1 < 10, String.Format(0, "0#"), LTrim(Str(0))) & ":" & String.Format(Minutes, "00") & ":" & String.Format(Second, "0#")
                MyH = IIf(1 < 10, String.Format(0, "0#"), LTrim(Str(0))) & ":" & String.Format("{0:0#}", Minutes) & ":" & String.Format("{0:0#}", Second)
                Tps_Calc.Text = MyH
            End If


            If J > 1 And J = nb_loop Then '######### CALCUL DE NOMBRE DE PROCESS
                If Reste > 0 Then
                    Nb_Process1 = Nb_Process
                    Nb_Process = Reste
                Else
                    Nb_Process1 = Nb_Process
                    'Nb_Process = P2
                End If
            Else
                Nb_Process1 = Nb_Process
            End If

            Num_Fichier = ((J - 1) * Nb_Process1)

            ReDim Tab_Name_File(Nb_Process - 1)
            ReDim Tab_Comment(Nb_Process - 1)
            ReDim Tab_IsPonctuel(Nb_Process - 1)
            ReDim Info_Experience_Mat(Nb_Process - 1)
            First_Init = False
            ReDim Process_Abort(Nb_Process - 1)
            'For i = 0 To Nb_Process - 1
            '    ReDim Info_Experience_Mat(i).Selected_Pivot(Nb_Trc)
            'Next i
            ReDim Info_Mat_Raie(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Height(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Area(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Conc(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Oxyde(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_LOD(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Stat_Error(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Fit_Error(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Total_Error(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Final_Error(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Y_N_Q(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Str_Mat_Conc_100(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Conc_ppm(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Conc_RED(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Best_Yes(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Val_Mat_Best_Yes_RED(Nb_Process - 1, Nb_Elements_Mat - 1)
            ReDim Txt_Fichier_PAR_Mat_Filter(Nb_Process - 1)



            ReDim Txt_Fichier_PAR_Mat_HED(Nb_Process - 1)
            ReDim Txt_Fichier_PAR_Trc_HED(Nb_Process, Nb_Trc)


            ' Sequence_Matrix_Multi(Num_Fichier, IntNb_File)
            If OnlyTrace = False Then
                If Use_HED_Mat = True Then
                    Create_Fichier_Par_Mat_HED()
                End If
                Sequence_Matrix_Multi_Thread(Num_Fichier, IntNb_File, Nb_Process)
            End If

            If Fatal_Error = True Then Exit Function

            Nb_Trc_Total = 0

            If Nb_Trc > 0 Then

                'For i = 0 To Nb_Trc - 1
                '    Lecture_Fichier_Par_Trc(4, i) ' RECUPERE chemin fichiers HED
                'Next i

                For i = 0 To Nb_Trc - 1
                    Nb_Trc_Total = Nb_Trc_Total + Nb_Elements_Trc(i)
                Next i
                Nb_Trc_Total = Nb_Trc_Total - 1

                ReDim Info_Experience_Trc(Nb_Process - 1, Nb_Trc)
                ReDim Info_Trc_Raie(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Height(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Area(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Conc(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Oxyde(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_LOD(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Stat_Error(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Fit_Error(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_WithPivot_Error(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Total_Error(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Y_N_Q(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Error_Pivot(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Conc100(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Conc_ppm(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Best_Yes(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Best_Yes_RED(Nb_Process - 1, Nb_Trc_Total)
                ReDim Val_Trc_Pivot_Error(Nb_Process - 1, Nb_Trc)
                ReDim Error_Trace(Nb_Process - 1, Nb_Trc)

            End If


            For i = 0 To Nb_Trc - 1
                Sequence_Trace_Multi_Thread(Num_Fichier, IntNb_File, i)
            Next i

            If Fatal_Error = True Then Exit Function

            Progress.Text = Trim$(Str$(Num_Fichier + Nb_Process)) & "/" & Trim$(Str$(IntNb_File))

            System.Threading.Thread.Sleep(100) : Application.DoEvents()


            'On Error GoTo 1


            If Chemin_Rapport = "" Then 'Then
                'Excel_Open(i)
                xlBook = Excel_Open(0, Chemin_Rapport)
                Excel_Save(J) ' Save as.....
            Else
                Excel_Save(-1) ' Save..
            End If

            Entete_100()
            ReDim thread_tab_best_conc(Nb_Process)
            If Fatal_Error = True Then Exit Function
            Dim para_best_C As parametres_best_conc_Thread
            For l = 0 To Nb_Process - 1
                thread_tab_best_conc(l) = New System.Threading.Thread(AddressOf Calcul_Final_Best_Conc_New_Thread)
                If OnlyTrace = False Then

                    If Error_Matrix(l) = False Then
                        para_best_C.indx_file = tab_select_file_indices(l + ((Nb_Process1) * (J - 1)))
                        para_best_C.num_process = l
                        para_best_C.nb_trace = Nb_Trc
                        Tab_Thread_best_conc_Alive(l) = False
                        thread_tab_best_conc(l).Start(para_best_C)

                        ' Calcul_Final_Best_Conc_New_Thread(para_best_C) ', nb_data_read

                    Else
                        Calcul_Final_Best_Only_Trace_Conc_New_Thread(l + ((Nb_Process1) * (J - 1)), l, Nb_Trc) ', nb_data_read
                    End If
                End If

            Next l

            Dim best_done = 0
            Dim inc(100) As Boolean
            Do
                For l = 0 To Nb_Process - 1
                    Tab_Thread_best_conc_Alive(l) = thread_tab_best_conc(l).IsAlive
                    System.Threading.Thread.Sleep(50) : Application.DoEvents()
                    If Tab_Thread_best_conc_Alive(l) = False And inc(l) = False Then
                        best_done += 1
                        inc(l) = True
                        ProgressBar1.Value = ProgressBar1.Value + 1
                    End If
                Next l

            Loop While best_done <> Nb_Process


            For l = 0 To Nb_Process - 1
                Ecrire_Entete_Excel(l + ((Nb_Process1) * (J - 1)))

                If Chk_RoundValue.Checked = True Then
                    Arrondi_Mat_Elem(l)
                    Arrondi_Trc_Elem(l)

                    If mnuOxydeOUI.Checked = True Then
                        Arrondi_Mat_Oxyde(l)
                        Arrondi_Trc_Oxyde(l)
                    End If

                End If
                Arrondi_Mat_Error(l)
                Arrondi_Trc_Error(l)

                Excel_Create_Sheet_conc_and_error(l)

            Next l

            Excel_Write_Filename_Comment((Nb_Process1) * (J - 1), Nb_Process)
            If J = nb_loop Then legend = True
            MyErase = False
            If J = 1 And Offset_Excel > 0 Then MyErase = True
            Excel_Legend(Nb_file, legend, MyErase)
            'CLOSEDXML      Excel_Legend(Nb_file, legend, MyErase)
            Excel_Save(-1)

            If Format_Range_1 = True Then
                Do
                    System.Threading.Thread.Sleep(50) : Application.DoEvents()
                Loop While Tab_Thread_Data_Excel_Alive = True Or Tab_Thread_Format_Italic_Excel_Alive = True Or Tab_Thread_Format_Style_Excel_Alive = True

            End If

            Parametres_All_Thread.Num_File = Nb_Process1 * (J - 1)
            Parametres_All_Thread.Nb_Calcul = Nb_Process
            Tab_IsPonctuel_1 = Tab_IsPonctuel
            Tab_Name_File_1 = Tab_Name_File
            Tab_Comment_1 = Tab_Comment
            Info_Mat_Raie_1 = Info_Mat_Raie

            Val_Mat_Area_1 = Val_Mat_Area
            Val_Mat_Conc_1 = Val_Mat_Conc
            Val_Mat_LOD_1 = Val_Mat_LOD
            Val_Mat_Height_1 = Val_Mat_Height

            Val_Mat_Fit_Error_1 = Val_Mat_Fit_Error
            Val_Mat_Total_Error_1 = Val_Mat_Total_Error
            Val_Mat_Final_Error_1 = Val_Mat_Final_Error
            Val_Mat_Y_N_Q_1 = Val_Mat_Y_N_Q
            Val_Conc_S_100_1 = Val_Conc_S_100
            Val_Conc_S_ppm_1 = Val_Conc_S_ppm
            Val_Conc_S_RED_ppm_1 = Val_Conc_S_RED_ppm
            Val_Conc_S_RED100_1 = Val_Conc_S_RED100
            Val_Error_S_1 = Val_Error_S
            Val_Conc_And_Error_1 = Val_Conc_And_Error
            Val_Conc_And_Error100_1 = Val_Conc_And_Error100
            Val_Choix_S_1 = Val_Choix_S
            Val_Mat_Mtx_1 = Val_Mat_Mtx
            NomDet_Trc_1 = NomDet_Trc

            Val_Trc_Y_N_Q_1 = Val_Trc_Y_N_Q
            Val_Trc_Conc_1 = Val_Trc_Conc
            Val_Trc_LOD_1 = Val_Trc_LOD
            Val_Trc_Area_1 = Val_Trc_Area
            Val_Trc_Fit_Error_1 = Val_Trc_Fit_Error
            Val_Trc_WithPivot_Error_1 = Val_Trc_WithPivot_Error
            Val_Trc_Height_1 = Val_Trc_Height
            Val_YNQ_Final_1 = Val_YNQ_Final

            If mnuOxydeOUI.Checked = True Then
                Val_Mat_Oxyde_1 = Val_Mat_Oxyde
                Val_Trc_Oxyde_1 = Val_Trc_Oxyde
            End If


            Function_Data_To_LibreOffice(Parametres_All_Thread)
            Function_Excel_Format_Style(Parametres_All_Thread)
            Function_Excel_Format_Italic(Parametres_All_Thread)
            Excel_Save(-1)
            Format_Range_1 = True

            If J = 1 Then
                Tps_1_Loop = My.Computer.Clock.TickCount - start_calc
                Tps_Total = (CInt(Tps_1_Loop) * nb_loop) / 1000
            End If
        Next J


        Excel_Save(-1)

        'CLOSEDXML xlApp.visible = True

        'Fermer_Gupix (hConsole)
        Second = (Tps_1_Loop / 1000) * nb_loop
        Minutes = Int((Second Mod 3600) / 60)
        Second = Second - (Minutes * 60)
        ' xlApp.visible = True
        '''' xlApp.WindowState = Excel.XlWindowState.xlNormal
        MyH = IIf(1 < 10, String.Format(0, "0#"), LTrim(Str(0))) & ":" & String.Format(Minutes, "0#") & ":" & String.Format(Second, "0#")

        Tps_Calc.Text = "??:??:??" ' MyH
        '  Fs_Log.Close()

        Main_Process_Ponctual = True
        '    Quitter.Visible = False
        MyPause.Visible = False
        ProgressBar1.Visible = False


        ToolStripStatusLabel1.Text = "Calcul Finished"

    End Function


    '***********************************************************************************************
    '****************************** SEQUENCE MATRICE************************************************
    '***********************************************************************************************


    Public Sub Th_FitToPNG(Parametres As Struct_Parametres_Thread)
        Dim LNum_File As Integer
        Dim Num_Proc As Integer
        Dim Num_voie As Integer
        Dim File_Name As String
        Dim Chemin_source As String
        LNum_File = Parametres.Num_File
        Num_Proc = Parametres.Num_Proc
        Num_voie = Parametres.voie
        File_Name = Parametres.File_Name
        Dim Pixtable_ok As Boolean


        Do
            Pixtable_ok = File.Exists(Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE.OUT")
            System.Threading.Thread.Sleep(20)
        Loop While Pixtable_ok = False

        Dim information = My.Computer.FileSystem.GetFileInfo(Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE.OUT")

        If information.Length > 10000 Then
            Application.DoEvents() : System.Threading.Thread.Sleep(50)
        Else
            Application.DoEvents() : System.Threading.Thread.Sleep(500)
        End If

        Try
            File.Delete(Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE" & CStr(Num_Proc + LNum_File) & "_" & CStr(Num_voie) & ".OUT")
        Catch ex As Exception

        End Try

        Try
            My.Computer.FileSystem.CopyFile(Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE.OUT", Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE" & CStr(Num_Proc + LNum_File) & "_" & CStr(Num_voie) & ".OUT", True)
            System.Threading.Thread.Sleep(40)
        Catch ex As Exception
            System.Threading.Thread.Sleep(700)
            'File.Delete(Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE" & CStr(Num_Proc + LNum_File) & "_" & CStr(Num_voie) & ".OUT")
            My.Computer.FileSystem.CopyFile(Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE.OUT", Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE" & CStr(Num_Proc + LNum_File) & "_" & CStr(Num_voie) & ".OUT", True)
        End Try

        Try
            ' Application.DoEvents()
            ' FITtoPNG_XPV8_Hide1(Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE" & CStr(Num_Proc + LNum_File) & "_" & CStr(Num_voie) & ".OUT", Chemin_Processed_Data, File_Name, 0) 'CLng(Sec_Delay.Text))
            '2023''''''''''''''''''''''  FITto
            Chemin_source = Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE" & CStr(Num_Proc + LNum_File) & "_" & CStr(Num_voie) & ".OUT" 'PNG_2023_Hide(File_Name, Chemin_Processed_Data, Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE" & CStr(Num_Proc + LNum_File) & "_" & CStr(Num_voie) & ".OUT", "Toto") 'CLng(Sec_Delay.Text))
            FiTtoPNG_LV16(0, Chemin_source, Chemin_Processed_Data, File_Name, Nom_Projet)
        Catch ex As Exception

        End Try

        Try
            File.Delete(Chemin_GupixWin_Multi(Num_Proc) & "\PIXTABLE" & CStr(Num_Proc + LNum_File) & "_" & CStr(Num_voie) & ".OUT")
        Catch ex As Exception

        End Try
    End Sub


    Public Sub Function_Data_To_LibreOffice(Parametres As Struct_Parametres_Thread)
        Dim LNum_File As Integer
        Dim LNb_process As Integer
        LNum_File = Parametres.Num_File
        LNb_process = Parametres.Nb_Calcul
        Excel_Data_ClosedXml(LNum_File, LNb_process, xlBook)

    End Sub

    'CLOSEDXML 
    Public Sub Function_Excel_Format_Style(Parametres As Struct_Parametres_Thread)
        Dim LNum_File As Integer
        Dim LNb_process As Integer
        Dim Nom_Excel As String
        LNum_File = Parametres.Num_File
        LNb_process = Parametres.Nb_Calcul
        Nom_Excel = ""
        Th_Excel_Format_Range_Special(LNum_File, LNb_process, Nom_Excel)
    End Sub


    Public Function Insert_Charge_in_Par_File(TextPar As String, Num_File As Integer, Num_column As Integer) As String
        Dim Split1() As String
        Dim Return_str As String
        Dim Pos_str_charge As Integer
        Dim num_line As Integer
        Return_str = ""
        Split1 = Strings.Split(TextPar, vbCrLf)
        For Each Str1 In Split1
            Pos_str_charge = Strings.InStr(1, Str1, "(Beam charge ", vbTextCompare)
            If Pos_str_charge > 0 Then Exit For
            num_line += 1
        Next

        Split1(num_line) = Charge_Exp(Num_column, tab_select_file_indices(Num_File)) & vbTab & vbTab & "(Beam charge in units consistent with H-value)"

        For Each Str1 In Split1
            Return_str = Return_str & Str1 & vbCrLf
        Next

        Return Return_str

        For i = 0 To 10

        Next

    End Function

    '***********************************************************************************************
    '****************************** SEQUENCE MATRICE************************************************
    '***********************************************************************************************

    Public Sub Sequence_Matrix_Multi_Thread(Num_Fichier As Integer, Num_Fin As Integer, nb_process_loc As Integer)
        Dim canaux As Integer
        Dim Charge As Single
        Dim i As Integer
        Dim j As Integer
        Dim T As Integer
        Dim Str_Entete As String
        Dim data As String
        Dim Result As String
        Dim ret
        Dim Tab_PixWin_END(1000) As Boolean
        Dim Tab_Data_Lue(1000) As Boolean

        Dim MyYear As String
        Dim Second_Temp As Byte, Minute_Temp As Byte, Heure_Temp As Long
        Dim MyHeures As String, Second As String, MyDate As String
        Dim pos As Integer, pos1 As Integer, Pos0 As Byte
        Dim tps_cps As String
        Dim commentaire As String
        Dim P As Integer
        Dim nb_data_read As Integer
        Dim Nb_Calcul As Integer
        Dim get_tps As Boolean
        Dim start1
        Dim Lect_Depth As Boolean
        Dim Parametres_All_Thread As Struct_Parametres_Thread
        Dim Tab_Thread_Lit_Element_Alive(100) As Boolean
        Dim Tab_Thread_best_conc_Alive(100) As Boolean
        Dim Tab_Inc_Done(100) As Boolean
        Dim Tab_Thread_Oxyde_Alive(100) As Boolean
        Dim Tab_Inc_Oxyde_Done(100) As Boolean
        Dim Tab_Thread_Recup_Filter_Alive(100) As Boolean
        Dim Tab_Inc_Recup_Filter_Done(100) As Boolean
        Dim Tab_Thread_FitToPNG_Alive(100) As Boolean
        Dim SplitText() As String
        Dim SplitText1() As String
        Dim SR As StreamReader
        Dim Somme As String
        Dim gupix_end As String

        Lect_Depth = True


        ToolStripStatusLabel1.Text = "Matrix processing"
        '  Fs_Log.writeline("Calculate MATRIX" & CStr(Num_Fichier))
        data = ""
        For v = 0 To 10


        Next
        For i = 0 To nb_process_loc - 1

            '********************************************************* Ecriture ENTETE DATE COMMENTAIRE SUR GUPIXWIN.PAR *******************

            If hdf5_mode = True Then '######################################### HDF5 #### 05/2023
                MyYear = Attrib_Spectrum(Num_Fichier + i, 0, 4)
                MyDate = "'" & Attrib_Spectrum(Num_Fichier + i, 0, 4) & "-" & Attrib_Spectrum(Num_Fichier + i, 0, 3) & "-01'"

                Second = Attrib_Spectrum(Num_Fichier + i, 0, 2)
                Heure_Temp = Int(Second / 3600)
                Minute_Temp = Int((Second Mod 3600) / 60)
                Second_Temp = Second Mod 3600 - Minute_Temp * 60
                MyHeures = IIf(Heure_Temp < 10, String.Format(Heure_Temp, "0#"), LTrim(Str(Heure_Temp))) & ":" & String.Format(Minute_Temp, "0#") & ":" & String.Format(Second_Temp, "0#")
                tps_cps = Attrib_Spectrum(Num_Fichier + i, 0, 1) 'ACQ TIME
                commentaire = " " & Attrib_Spectrum(Num_Fichier + i, 0, 5)
                Tab_Comment(i) = commentaire

                data = "2048 1" & vbCrLf
                data = data & MyDate & " '" & MyHeures & "' " & tps_cps & "    " & Attrib_Spectrum(Num_Fichier + i, 0, 2) & " ' ' " & vbCrLf

                For canaux = 0 To 2047
                    data = data & AllSpectres_hdf5(Num_Fichier + i, 0, canaux) & vbCrLf ' Nb_file, Num_det, canaux
                Next

            Else '############################## FICHIER ASCII AVANT HDF5 #####################

                If Adjust_Filter_B = True Then
                    SR = File.OpenText(Chemin_Data + "\" + Fichier_Matrix(0))
                Else
                    Try
                        SR = File.OpenText(Chemin_Data + "\" + Fichier_Matrix(i + Num_Fichier))
                    Catch ex As Exception
                        MsgBox("Fatal Error, check matrix spectra file '" & Fichier_Matrix(i + Num_Fichier) & "'present into the data folder or error inside the file", MsgBoxStyle.Information, "Error reading Matrix file")
                        Fatal_Error = True
                        Exit Sub
                    End Try

                End If

                data = SR.ReadLine
                If Len(data) < 6 Then
                    data = "2048  0" & vbCrLf
                End If

                If data = Nothing Then
                    MsgBox("Fatal Error, check matrix spectra file'" & Fichier_Matrix(i + Num_Fichier) & "' file is empty", MsgBoxStyle.Information, "Error reading Matrix file")
                    Fatal_Error = True
                    Exit Sub
                End If

                Dim entete_tmp = Mid(data, 5, 2)
                Dim nbcanaux = Mid(data, 1, 4)
                'data = Strings.StrReverse(data)
                If entete_tmp = " 1" Then
                    'data = Strings.StrReverse(data)
                    Str_Entete = SR.ReadLine
                    SplitText = Split(Str_Entete, " ")
                    MyYear = SplitText(0) 'Mid(Str_Entete, 1, 4)
                    MyDate = "'" & MyYear & "-" & Mid(Str_Entete, 6, 2) & "-01'"
                    Second = SplitText(2) 'Mid(Str_Entete, 9, 6)
                    Heure_Temp = Int(Second / 3600)
                    Minute_Temp = Int((Second Mod 3600) / 60)
                    Second_Temp = Second Mod 3600 - Minute_Temp * 60
                    MyHeures = IIf(Heure_Temp < 10, String.Format(Heure_Temp, "0#"), LTrim(Str(Heure_Temp))) & ":" & String.Format(Minute_Temp, "0#") & ":" & String.Format(Second_Temp, "0#")

                    Somme = SplitText(4)
                    pos = 0

                    For P = 0 To 3 'Extrait le real time et integrale du spectre
                        pos = InStr(pos + 1, Str_Entete, " ", vbTextCompare)
                        If P = 2 Then
                            pos1 = pos
                            Pos0 = pos
                        End If
                        pos1 = InStr(pos1 + 1, Str_Entete, " ", vbTextCompare)
                    Next P

                    tps_cps = SplitText(3)
                    SplitText = Split(Str_Entete, "'")

                    'commentaire = " '" & Mid(Str_Entete, pos1, Len(Str_Entete) + 1 - pos1) & "'"

                    If Adjust_Filter_B = False Then 'Empêche remplacement epaisseur Filtre déja dans Tab_Comment() si Adjust Filter fonction

                        Try
                            commentaire = SplitText(1)
                            SplitText1 = Split(commentaire, ",")

                            If UBound(SplitText1) > 8 Then ' La ligne commentaire est de format " DrN,1000,1000,50,50,Proton,3000keV,......"
                                Tab_Comment(i) = SplitText1(0)
                                commentaire = SplitText1(0)
                            Else
                                Tab_Comment(i) = commentaire
                            End If
                        Catch ex As Exception
                            commentaire = "---"
                            Tab_Comment(i) = commentaire
                        End Try

                    End If

                    If Mid(commentaire, 2, 2) <> " '" Then
                        commentaire = " '" & commentaire & "'"
                    Else
                        commentaire = " " & commentaire
                    End If

                    If nbcanaux <> 0 Then ' BUG sept 2024 error spectre double dans fichier !!
                        data = data & vbCrLf & MyDate & " '" & MyHeures & "' " + tps_cps + " " + Somme + "  '  '" + vbCrLf
                        For j = 0 To nbcanaux - 1
                            data += SR.ReadLine() + vbCrLf
                        Next j
                    Else
                        data = data & vbCrLf & MyDate & " '" & MyHeures & "' " + tps_cps + " " + Somme + "  '  '" + vbCrLf + SR.ReadToEnd()
                    End If

                    '& SR.ReadToEnd()
                    Dim titi = Len(data)
                Else
                    If nbcanaux <> 0 Then
                        For j = 0 To nbcanaux - 1
                            data += SR.ReadLine() + vbCrLf
                        Next j
                    Else
                        data = data & vbCrLf & SR.ReadToEnd()
                    End If
                    'Example      '2007-2-12' '15:19:50' 572  1709185 'COMMENTAIRE'
                End If
                SR.Close()

            End If
            If Len(data) < 400 Then
                MsgBox("Fatal Error, check matrix spectra file'" & Fichier_Matrix(i + Num_Fichier) & "' , data are less than 100 values", MsgBoxStyle.Information, "Error reading Matrix file")
                Fatal_Error = True
                Exit Sub
            End If

            Kill(Chemin_GupixWin_Multi(i) & "\gupixwin.par")
            Kill(Chemin_GupixWin_Multi(i) & "\*.out")
            Kill(Chemin_GupixWin_Multi(i) & "\*.csv")




            If Adjust_Filter_B = True And ComboBox_Type_F.Text = CbDetMat.Text Then
                If Use_ext_charge_Mat = True Then Txt_Fichier_PAR_Mat_Filter(i) = Insert_Charge_in_Par_File(Txt_Fichier_PAR_Mat_Filter(i), i + Num_Fichier, num_column_charge_csv_MAT) '0 = Matrix 
                File.WriteAllText(Chemin_GupixWin_Multi(i) & "\gupixwin.par", Txt_Fichier_PAR_Mat_Filter(i) & vbCrLf & data & vbCrLf & "--------------------------------------------------" & vbCrLf)

            ElseIf Use_HED_Mat = True Then
                If Use_ext_charge_Mat = True Then Txt_Fichier_PAR_Mat_HED(i) = Insert_Charge_in_Par_File(Txt_Fichier_PAR_Mat_HED(i), i + Num_Fichier, num_column_charge_csv_MAT) '0 = Matrix 
                File.WriteAllText(Chemin_GupixWin_Multi(i) & "\gupixwin.par", Txt_Fichier_PAR_Mat_HED(i) & vbCrLf & data & vbCrLf & "--------------------------------------------------" & vbCrLf)
            Else
                If Use_ext_charge_Mat = True Then Txt_Fichier_PAR_Mat = Insert_Charge_in_Par_File(Txt_Fichier_PAR_Mat, i + Num_Fichier, num_column_charge_csv_MAT) '0 = Matrix 
                File.WriteAllText(Chemin_GupixWin_Multi(i) & "\gupixwin.par", Txt_Fichier_PAR_Mat & vbCrLf & data & vbCrLf & "--------------------------------------------------" & vbCrLf)
            End If

            Try
                Kill(Chemin_GupixWin_Multi(i) & "\pixwin.end")
            Catch ex As Exception

            End Try

        Next i

        '************************************* EXECUTION DE PIXWIN.EXE ********************
        For i = 0 To nb_process_loc - 1

            If hdf5_mode = True Then
                If Attrib_Spectrum(Num_Fichier + i, 0, 7) > 0 Then ret = Shell(Chemin_GupixWin_Multi(i) & "\pixwin.bat", vbHide)
            Else
                ret = Shell(Chemin_GupixWin_Multi(i) & "\pixwin.bat", vbHide)
            End If



            Application.DoEvents() : System.Threading.Thread.Sleep(20)
            thread_tab_Recup_Filter(i) = New System.Threading.Thread(AddressOf Recup_Cps_Charge_Filters_Mat_Thread)
            thread_tab_Element(i) = New System.Threading.Thread(AddressOf Lit_Element_Multi_Thread)
            thread_tab_FitToPNG(i) = New System.Threading.Thread(AddressOf Th_FitToPNG)

            If mnuOxydeOUI.Checked = True Then
                thread_tab_oxyde(i) = New System.Threading.Thread(AddressOf Lit_Oxyde_Multi_Thread)
                Tab_Inc_Oxyde_Done(i) = False
                Tab_Thread_Oxyde_Alive(i) = True
            Else

            End If

            Tab_Inc_Recup_Filter_Done(i) = False
            Tab_Thread_Recup_Filter_Alive(i) = True

            Tab_Inc_Done(i) = False
            Tab_Thread_Lit_Element_Alive(i) = True
            Tab_Thread_FitToPNG_Alive(i) = False
            Error_Matrix(i) = False

        Next 'I

        start1 = DateAndTime.Timer

        Do

            For i = 0 To nb_process_loc - 1

                If hdf5_mode = True Then
                    If Attrib_Spectrum(Num_Fichier + i, 0, 7) = 0 And Tab_Data_Lue(i) = False Then ' CAS SPECTRE = 0 
                        Tab_Data_Lue(i) = True
                        Nb_Calcul = Nb_Calcul + 1
                    End If
                End If

                Tab_PixWin_END(i) = File.Exists(Chemin_GupixWin_Multi(i) & "\pixwin.end")

                If Tab_PixWin_END(i) = False Then ToolStripStatusLabel1.Text = "Matrix processing file " & Fichier_Matrix(Num_Fichier + i)
                If Tab_Thread_FitToPNG_Alive(i) = True Then ToolStripStatusLabel1.Text = "FitToPNG " & Fichier_Matrix(Num_Fichier + i)

                If Tab_PixWin_END(i) = True And Tab_Data_Lue(i) = False And Tab_Thread_Lit_Element_Alive(i) = False And Tab_Thread_Oxyde_Alive(i) = False Then
                    SR = File.OpenText(Chemin_GupixWin_Multi(i) & "\pixwin.end")
                    gupix_end = SR.ReadLine
                    SR.Close()
                    If CInt(gupix_end) = 0 Then
                        My.Application.DoEvents() : System.Threading.Thread.Sleep(20)
                        Nb_Calcul = Nb_Calcul + 1
                        Tab_Data_Lue(i) = True
                        Parametres_All_Thread.voie = 1
                        Parametres_All_Thread.Num_File = Num_Fichier
                        Parametres_All_Thread.Num_Proc = i
                        Parametres_All_Thread.Num_Trc = 0
                        Parametres_All_Thread.Fact_Correct = 1

                        If Adjust_Filter_B = True Then
                            Parametres_All_Thread.File_Name = Fichier_Matrix(0)
                        Else
                            Parametres_All_Thread.File_Name = Fichier_Matrix(Num_Fichier + i)
                        End If

                        thread_tab_Element(i).Start(Parametres_All_Thread)
                        My.Application.DoEvents() : System.Threading.Thread.Sleep(50)
                        thread_tab_Recup_Filter(i).Start(Parametres_All_Thread)



                        If mnuOxydeOUI.Checked = True Then
                            thread_tab_oxyde(i).Start(Parametres_All_Thread)
                        Else
                            Tab_Thread_Oxyde_Alive(i) = False
                        End If

                        My.Application.DoEvents() : System.Threading.Thread.Sleep(10)

                        ' If Lect_Depth = True Then Lect_Depth = Lit_Depth(1, I)
                        ProgressBar1.Value = ProgressBar1.Value + 1

                        If Adjust_Filter_B = False And Tab_Thread_FitToPNG_Alive(i) = False Then
                            'Dim fs As Object
                            'fs = CreateObject("Scripting.FileSystemObject")
                            ToolStripStatusLabel1.Text = "Plot Fit To image PNG " & CStr(i)
                            Application.DoEvents() : System.Threading.Thread.Sleep(500)
                            Dim Pixtable_ok As Boolean
                            Do
                                Pixtable_ok = File.Exists(Chemin_GupixWin_Multi(i) & "\PIXTABLE.OUT")
                                System.Threading.Thread.Sleep(20)
                            Loop While Pixtable_ok = False

                            thread_tab_FitToPNG(i).Start(Parametres_All_Thread)

                        End If

                    Else 'ERROR WITH GUPIX PROCESS
                        Nb_Calcul = Nb_Calcul + 1
                        Tab_Data_Lue(i) = True
                        Error_Matrix(i) = True
                    End If

                End If

                My.Application.DoEvents() : System.Threading.Thread.Sleep(20)
                Tab_Thread_Lit_Element_Alive(i) = thread_tab_Element(i).IsAlive
                If mnuOxydeOUI.Checked = True Then Tab_Thread_Oxyde_Alive(i) = thread_tab_oxyde(i).IsAlive
                Tab_Thread_Recup_Filter_Alive(i) = thread_tab_Recup_Filter(i).IsAlive
                Tab_Thread_FitToPNG_Alive(i) = thread_tab_FitToPNG(i).IsAlive

                If Tab_Thread_Lit_Element_Alive(i) = False And Tab_Thread_Oxyde_Alive(i) = False And Tab_Thread_Recup_Filter_Alive(i) = False And Tab_Data_Lue(i) = True And Tab_Inc_Done(i) = False And Tab_Thread_Oxyde_Alive(i) = False And Tab_Thread_FitToPNG_Alive(i) = False Then
                    nb_data_read = nb_data_read + 1
                    Tab_Inc_Done(i) = True

                    If Error_Matrix(i) = False Then
                        If Calcul_With_Trc = False Then
                            ReDim Val_Mat_Mtx(nb_process_loc - 1, CInt(Nb_Elements_Mat))
                            ReDim Val_Inv_Mtx(nb_process_loc - 1, 50)
                            If nb_gamma > 0 And tab_gamma_external_value_ok(Num_Fichier + i) Then
                                Insert_Matrix_gamma(tab_select_file_indices(Num_Fichier + i), i)
                            Else
                                Insert_Matrix(Num_Fichier + i, i)
                            End If
                        End If
                        ' ProgressBar1.Value = ProgressBar1.Value + 1
                    End If
                End If

            Next i

            My.Application.DoEvents() : System.Threading.Thread.Sleep(20)

            If (nb_process_loc - nb_data_read) < nb_process_loc - (nb_process_loc - 1) And get_tps = False Then
                start1 = DateAndTime.Timer '############ Initilisate le cpt Temps pour le dernier Process
                get_tps = True
            ElseIf (nb_process_loc - nb_data_read) >= nb_process_loc Then
                start1 = DateAndTime.Timer
            End If
            'Start1 = Timer
            If DateAndTime.Timer - start1 > nb_process_loc * 4 Then ' Temps d'attente pour le dernier calcul
                Exit Do
            End If
        Loop While nb_data_read <> nb_process_loc


        Result = 0
        'My.Application.DoEvents() :System.Threading.Thread.Sleep(10)
        Select Case Result

            Case 0

            Case 1 To 100
                MsgBox("Erreur", vbCritical, "Erreur Execution Pixwin")
        End Select



        Ratio = String.Format((Val(Charge) / MyChargeStd), "0#.#####")  '


        '"""Recup OXYDE
        'If mnuOxydeOUI.Checked = True Then Lit_Oxyde tab_oxyde, 1
        'Lit_Depth (1)
        For i = 0 To nb_process_loc - 1
            thread_tab_Element(i) = Nothing
            thread_tab_oxyde(i) = Nothing
        Next i

    End Sub
    '***********************************************************************************************
    '****************************** SEQUENCE TRC ****************************************************
    '***********************************************************************************************

    Private Sub Sequence_Trace_Multi_Thread(Num_Fichier As Integer, Int_Fichier As Integer, Num_Trc As Integer)

        Dim MyFactCorrect As Single
        Dim Charge As Single
        Dim t As Integer
        Dim i As Integer
        Dim j As Integer
        Dim data As String
        Dim MyYear As String
        Dim Second_Temp As Byte
        Dim Minute_Temp As Byte
        Dim Heure_Temp As Long
        Dim MyHeures As String
        Dim Second As String
        Dim MyDate As String
        Dim tps_cps As String
        Dim commentaire As String

        Dim Lect_Depth As Boolean
        Dim Tab_PixWin_END(100) As Boolean
        Dim Tab_Pivot_Done(100) As Boolean
        Dim Tab_Data_Lue(100) As Boolean
        Dim nb_data_read As Integer
        Dim Tmp_Fichier_PAR_Trc As String
        Dim Nb_Calcul As Integer
        Dim File_Trc As String
        Dim Num_File_final As Integer

        Dim Parametres_All_Thread As Struct_Parametres_Thread
        Dim Tab_Thread_Lit_Element_Alive(100) As Boolean
        Dim Tab_Thread_Lit_Just_Element_Alive(100) As Boolean
        Dim Tab_Inc_Done(100) As Boolean
        Dim Tab_Fit_Done(100) As Boolean
        Dim Tab_Thread_Oxyde_Alive(100) As Boolean
        Dim Tab_Inc_Oxyde_Done(100) As Boolean
        Dim Tab_Thread_Recup_Filter_Alive(100) As Boolean
        Dim Tab_Inc_Recup_Filter_Done(100) As Boolean
        Dim Tab_Thread_Ecriture_Charge_Alive(100) As Boolean
        Dim Tab_Thread_Calcul_Ecriture_Charge_Alive(100) As Boolean
        Dim Tab_Run_Done(100) As Boolean
        Dim Tab_Thread_FitToPNG_Alive(100) As Boolean
        Dim SR As StreamReader
        Charge = 1
        t = 0
        Lect_Depth = True

        Txt_Fichier_PAR_Trc = ""
        If OnlyTrace = False Then Lecture_Fichier_Par_Trc(1, Num_Trc)
        Pivot = Val(Pivot1(Tab_Num_Trc(Num_Trc)))
        Tmp_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc
        ReDim Val_Mat_Mtx(Nb_Process - 1, CInt(Nb_Elements_Mat))
        ReDim Val_Inv_Mtx(Nb_Process - 1, 50)
        Indice_Pivot_trc(Num_Trc, 0) = -1
        data = ""

        For i = 0 To Nb_Process - 1

            If Error_Matrix(i) = False Then
                Txt_Fichier_PAR_Trc = Tmp_Fichier_PAR_Trc

                If mnuOxydeOUI.Checked = True And OnlyTrace = False Then

                    If nb_gamma > 0 And tab_gamma_external_value_ok(Num_Fichier + i) = True Then
                        Insert_Matrix_gamma(tab_select_file_indices(Num_Fichier + i), i)
                    Else
                        Insert_Matrix(Num_Fichier + i, i)
                    End If
                Else
                    'Insert la matrice sous forme elementaire + Element Invisible a la fin
                    If nb_gamma > 0 And tab_gamma_external_value_ok(Num_Fichier + i) = True Then
                        Insert_Matrix_gamma(Num_Fichier + i, i)
                    Else
                        Insert_Matrix(Num_Fichier + i, i)
                    End If
                End If

                If OnlyTrace = False Then Lecture_Fichier_Par_Trc(2, Num_Trc) ' Relecture du fichier par jusqu'a données
                If OnlyTrace = True Then
                    Txt_Fichier_PAR_Trc = ""
                    Lecture_Fichier_Par_Trc(3, Num_Trc) ' LECTURE JUSQU'A SPECTRA DATA
                End If

                If Adjust_Filter_B = True Then
                    Num_File_final = 0
                Else
                    Num_File_final = i + Num_Fichier
                End If

                Select Case Tab_Num_Trc(Num_Trc)
                    Case 0
                        File_Trc = Fichier_Trace0(Num_File_final)
                    Case 1
                        File_Trc = Fichier_Trace1(Num_File_final)
                    Case 2
                        File_Trc = Fichier_Trace2(Num_File_final)
                    Case 3
                        File_Trc = Fichier_Trace3(Num_File_final)
                    Case 4
                        File_Trc = Fichier_Trace4(Num_File_final)
                    Case 5 ' 1+2
                        File_Trc = Fichier_Trace5(Num_File_final)
                    Case 6 ' All
                        File_Trc = Fichier_Trace6(Num_File_final)
                    Case 7 '2+3
                        File_Trc = Fichier_Trace7(Num_File_final)
                    Case 8
                        File_Trc = Fichier_Trace8(Num_File_final)
                    Case Else
                        File_Trc = Fichier_Trace0(Num_File_final)
                End Select

                ToolStripStatusLabel1.Text = "Trace Calculation " & File_Trc
                My.Application.DoEvents()


                If hdf5_mode = True Then

                    '********************************************************* Ecriture ENTETE DATE COMMENTAIRE SUR GUPIXWIN.PAR *******************
                    MyYear = Attrib_Spectrum(Num_Fichier, Num_Trc + 1, 4)
                    MyDate = "'" & Attrib_Spectrum(Num_Fichier, Num_Trc + 1, 4) & "-" & Attrib_Spectrum(Num_Fichier, Num_Trc, 3) & "-01'"
                    Second = Attrib_Spectrum(Num_Fichier, Num_Trc + 1, 2)
                    Heure_Temp = Int(Second / 3600)
                    Minute_Temp = Int((Second Mod 3600) / 60)
                    Second_Temp = Second Mod 3600 - Minute_Temp * 60
                    MyHeures = IIf(Heure_Temp < 10, String.Format(Heure_Temp, "0#"), LTrim(CStr(Heure_Temp))) & ":" & String.Format(Minute_Temp, "0#") & ":" & String.Format(Second_Temp, "0#")
                    tps_cps = Attrib_Spectrum(Num_Fichier, Num_Trc + 1, 1) 'ACQ TIME
                    commentaire = " " & Attrib_Spectrum(Num_Fichier, Num_Trc + 1, 5)

                    data = "2048 1" & vbCrLf
                    data = data & MyDate & " '" & MyHeures & "' " & tps_cps & "    " & Attrib_Spectrum(Num_Fichier, 0, 2) & " ' ' " & vbCrLf

                    For canaux = 0 To 2047
                        data = data & AllSpectres_hdf5(Num_Fichier, Num_Trc + 1, canaux) & vbCrLf ' Nb_file, Num_det, canaux
                    Next
                Else

                    Try
                        SR = File.OpenText(Chemin_Data + "\" + File_Trc)
                    Catch ex As Exception
                        MsgBox("Fatal Error, check spectra file for trace N°" & Tab_Num_Trc(Num_Trc) & " or " & Ref_DataSet_ToRead(Num_Trc) & " are present into the data folder ", MsgBoxStyle.Information, "Error reading file trace")
                        Error_Trace(i, Num_Trc) = True 'Fatal_Error = True
                    End Try

                    If Error_Trace(i, Num_Trc) = False Then
                        data = SR.ReadLine

                        If Len(data) < 6 Then ' BUg pas d'entete avec le nb de canaux on le force à 2048 !
                            data = "2048  0" & vbCrLf
                        End If

                        Dim entete_tmp = Mid(data, 5, 2)
                        Dim nbcanaux = Mid(data, 1, 4)

                        If entete_tmp = " 1" Then 'Strings.RSet(data, 1) = "1" Or Strings.RSet(data, 2) = "1 " Then
                            Dim Str_Entete = SR.ReadLine
                            MyYear = Mid(Str_Entete, 1, 4)
                            MyDate = "'" & MyYear & "-" & Mid(Str_Entete, 6, 2) & "-01'"
                            Second = Mid(Str_Entete, 9, 6)
                            Heure_Temp = Int(Second / 3600)
                            Minute_Temp = Int((Second Mod 3600) / 60)
                            Second_Temp = Second Mod 3600 - Minute_Temp * 60
                            MyHeures = IIf(Heure_Temp < 10, String.Format(Heure_Temp, "0#"), LTrim(CStr(Heure_Temp))) & ":" & String.Format(Minute_Temp, "0#") & ":" & String.Format(Second_Temp, "0#")

                            Dim pos = 0
                            For P = 0 To 2
                                pos = InStr(pos + 1, Str_Entete, " ", vbTextCompare)
                            Next P
                            Dim pos1 = pos

                            For P = 0 To 1
                                pos1 = InStr(pos1 + 1, Str_Entete, " ", vbTextCompare)
                            Next P

                            tps_cps = Mid(Str_Entete, pos, pos1 - pos)
                            commentaire = " '" & Mid(Str_Entete, pos1, Len(Str_Entete) + 1 - pos1) & "'"

                            If nbcanaux <> 0 Then
                                data = data & vbCrLf & MyDate & " '" & MyHeures & "' " & tps_cps & " '  '" & vbCrLf

                                For j = 0 To nbcanaux - 1
                                    data += SR.ReadLine() + vbCrLf
                                Next j

                            Else
                                data = data & vbCrLf & MyDate & " '" & MyHeures & "' " & tps_cps & " '  '" & vbCrLf & SR.ReadToEnd
                            End If

                            'data = data & vbCrLf & MyDate & " '" & MyHeures & "' " & tps_cps & " '  '" & vbCrLf & SR.ReadToEnd

                        Else
                            data = data & vbCrLf & SR.ReadToEnd
                        End If
                        'Example      '2007-2-12' '15:19:50' 572  1709185 'COMMENTAIRE'
                        SR.Close()
                    End If


                End If

                If Error_Trace(i, Num_Trc) = False Then

                    If Adjust_Filter_B = True And ComboBox_Type_F.Text = NomDet_Trc(Num_Trc) Then ' ADJUST FILTER
                        Kill(Chemin_GupixWin_Multi(i) & "\gupixwin.par")
                        File.WriteAllText(Chemin_GupixWin_Multi(i) & "\gupixwin.par", Txt_Fichier_PAR_Trc_Filter(Num_Fichier + i) & vbCrLf & data & vbCrLf & "--------------------------------------------------" & vbCrLf)
                        Pivot = Val(Pivot1(Tab_Num_Trc(Num_Trc)))
                    ElseIf Use_HED_Trc(Num_Trc) = True Then
                        If Use_ext_charge_Trc(Tab_Num_Trc(Num_Trc)) = True Then Txt_Fichier_PAR_Trc_HED(i, Num_Trc) = Insert_Charge_in_Par_File(Txt_Fichier_PAR_Trc_HED(i, Num_Trc), Num_Fichier + i, num_column_charge_csv_TRC(i))
                        File.WriteAllText(Chemin_GupixWin_Multi(i) & "\gupixwin.par", Txt_Fichier_PAR_Trc_HED(i, Num_Trc) & vbCrLf & data & vbCrLf & "--------------------------------------------------" & vbCrLf)
                    Else
                        Kill(Chemin_GupixWin_Multi(i) & "\gupixwin.par")
                        If Use_ext_charge_Trc(Tab_Num_Trc(Num_Trc)) = True Then Txt_Fichier_PAR_Trc = Insert_Charge_in_Par_File(Txt_Fichier_PAR_Trc, Num_Fichier + i, num_column_charge_csv_TRC(i))
                        'Tab_Num_Trc(Num_Trc) indique le num de det. trace coché correspondant à det1,det2,det3,det4,....
                        File.WriteAllText(Chemin_GupixWin_Multi(i) & "\gupixwin.par", Txt_Fichier_PAR_Trc & vbCrLf & data & vbCrLf & "--------------------------------------------------" & vbCrLf)
                    End If

                    Try
                        Kill(Chemin_GupixWin_Multi(i) & "\pixwin.end") ' DELETE pixwin.end
                    Catch ex As Exception

                    End Try
                End If

            End If
        Next i

        If OnlyTrace = False Then
            For i = 0 To Nb_Process - 1

                If Error_Matrix(i) = False And Error_Trace(i, Num_Trc) = False Then
                    Shell(Chemin_GupixWin_Multi(i) & "\pixwin.bat", AppWinStyle.Hide, False) ' Chemin_GupixWin_Multi(I), , Me.hWnd, I)
                Else
                    MsgBox("Matrix error: '" + CStr(Error_Matrix(i)) & " or Trace error: " & CStr(Error_Trace(i, Num_Trc)), MsgBoxStyle.Information, "Error reading matrix parameter")
                End If

                thread_tab_Recup_Filter(i) = New System.Threading.Thread(AddressOf Recup_Cps_Charge_Filters_Trc_Thread)
                thread_tab_Just_Element(i) = New System.Threading.Thread(AddressOf Lit_Element_Multi_Simple_Thread)
                If OnlyTrace = False Then thread_tab_Calcul_Ecriture_Charge(i) = New System.Threading.Thread(AddressOf Calcul_Ecriture_Charge_Thread)

                Tab_Inc_Recup_Filter_Done(i) = False
                Tab_Thread_Recup_Filter_Alive(i) = True
                Tab_Inc_Done(i) = False
                Tab_Fit_Done(i) = False
                Tab_Thread_Lit_Element_Alive(i) = True
                System.Threading.Thread.Sleep(20)
            Next i

            File_Trc = ""
            Do

                For i = 0 To Nb_Process - 1

                    If Error_Matrix(i) = False And Error_Trace(i, Num_Trc) = False Then
                        Select Case Tab_Num_Trc(Num_Trc)
                            Case 0
                                File_Trc = Fichier_Trace0(i + Num_Fichier)
                            Case 1
                                File_Trc = Fichier_Trace1(i + Num_Fichier)
                            Case 2
                                File_Trc = Fichier_Trace2(i + Num_Fichier)
                            Case 3
                                File_Trc = Fichier_Trace3(i + Num_Fichier)
                            Case 4
                                File_Trc = Fichier_Trace4(i + Num_Fichier)
                            Case 5 ' 1+2
                                File_Trc = Fichier_Trace5(i + Num_Fichier)
                            Case 6 ' All
                                File_Trc = Fichier_Trace6(i + Num_Fichier)
                            Case 7 '2+3
                                File_Trc = Fichier_Trace7(i + Num_Fichier)
                            Case 8
                                File_Trc = Fichier_Trace8(i + Num_Fichier)

                        End Select


                        My.Application.DoEvents() : System.Threading.Thread.Sleep(20)
                        Tab_PixWin_END(i) = File.Exists(Chemin_GupixWin_Multi(i) & "\pixwin.end")
                        If Tab_PixWin_END(i) = False Then ToolStripStatusLabel1.Text = "First Trace processing " & File_Trc
                        Parametres_All_Thread.voie = 2
                        Parametres_All_Thread.Num_File = Num_Fichier
                        Parametres_All_Thread.Num_Proc = i
                        Parametres_All_Thread.Num_Trc = Num_Trc
                        Parametres_All_Thread.Fact_Correct = 1
                        Parametres_All_Thread.Offset_Trc = 0

                        If Num_Trc > 0 Then
                            For t = 0 To Num_Trc - 1
                                Parametres_All_Thread.Offset_Trc = Parametres_All_Thread.Offset_Trc + Nb_Elements_Trc(t)
                            Next
                        Else
                            'Parametres_All_Thread.Offset_Trc = Parametres_All_Thread.Offset_Trc + Nb_Elements_Trc(K)
                        End If

                        If Tab_PixWin_END(i) = True And Tab_Data_Lue(i) = False Then
                            System.Threading.Thread.Sleep(20)
                            Nb_Calcul = Nb_Calcul + 1
                            MyFactCorrect = Ratio ' Rapport(num_fichier)

                            '''''''''If Nb_Calcul = 1 Then Recup_Cps_Charge_Filters_Trc(Num_Fichier, I, Num_Trc) 'Lecture nb_elements_X2
                            thread_tab_Just_Element(i).Start(Parametres_All_Thread)
                            Tab_Data_Lue(i) = True
                        End If

                        Tab_Thread_Lit_Just_Element_Alive(i) = thread_tab_Just_Element(i).IsAlive
                        My.Application.DoEvents() : System.Threading.Thread.Sleep(5)

                        If Tab_PixWin_END(i) = True And Tab_Thread_Lit_Just_Element_Alive(i) = False And Tab_Data_Lue(i) = True And Tab_Inc_Done(i) = False And Tab_Pivot_Done(i) = False Then
                            'Tab_Pivot_Done(I) = True
                            Tab_Data_Lue(i) = True
                            nb_data_read = nb_data_read + 1
                            ProgressBar1.Value = ProgressBar1.Value + 1
                            Tab_Inc_Done(i) = True
                            Parametres_All_Thread.Num_Proc = i
                            Parametres_All_Thread.Num_File = Num_Fichier
                            Parametres_All_Thread.Num_Trc = Num_Trc
                            Parametres_All_Thread.Num_Data = 0
                            Parametres_All_Thread.Nb_Calcul = nb_data_read

                            ' Ne fais pas le calcul de charge
                            If Use_ext_charge_Trc(Tab_Num_Trc(Num_Trc)) = False Then
                                thread_tab_Calcul_Ecriture_Charge(i).Start(Parametres_All_Thread)
                                Tab_Thread_Calcul_Ecriture_Charge_Alive(i) = True
                                Try
                                    Kill(Chemin_GupixWin_Multi(i) & "\pixwin.end") ' DELETE PIXWIN.END si pas de val Q dans fichier CSV (06/2024)
                                Catch ex As Exception

                                End Try

                            Else
                                Indice_Pivot_Mat(0) = -1 ' Pas de Pivot
                            End If

                        ElseIf (Error_Matrix(i) = True Or Error_Trace(i, Num_Trc) = True) And Tab_Data_Lue(i) = False Then  'ERROR MATRIX
                            Tab_Data_Lue(i) = True
                            nb_data_read = nb_data_read + 1
                            Try
                                ProgressBar1.Value = ProgressBar1.Value + 1
                            Catch ex As Exception

                            End Try
                            Tab_Inc_Done(i) = True
                        End If
                    End If

                Next i

                If nb_data_read > Nb_Process + 1 Then nb_data_read = Nb_Process 'Sort de la boucle while

            Loop While nb_data_read <> Nb_Process

        End If

        nb_data_read = 0

        For i = 0 To Nb_Process - 1
            Tab_Data_Lue(i) = False
            Tab_Run_Done(i) = False
            Tab_Inc_Done(i) = False
            Tab_Fit_Done(i) = False
            Tab_Thread_FitToPNG_Alive(i) = False
            thread_tab_Just_Element(i) = Nothing
            thread_tab_oxyde(i) = Nothing
            thread_tab_Element(i) = New System.Threading.Thread(AddressOf Lit_Element_Multi_Thread)
            thread_tab_oxyde(i) = New System.Threading.Thread(AddressOf Lit_Oxyde_Multi_Thread)
            thread_tab_FitToPNG_TRC(i) = New System.Threading.Thread(AddressOf Th_FitToPNG)
            thread_tab_Recup_Filter(i) = New System.Threading.Thread(AddressOf Recup_Cps_Charge_Filters_Trc_Thread)

        Next



        Do ' SECOND CALCUL TRACE

            'If Tab_PixWin_END(I) = True And Tab_Pivot_Done(I) = True And Tab_Data_Lue(I) = False Then


            For i = 0 To Nb_Process - 1
                If Error_Matrix(i) = False And Error_Trace(i, Num_Trc) = False Then
                    Parametres_All_Thread.Num_Proc = i
                    Parametres_All_Thread.Num_File = Num_Fichier
                    Parametres_All_Thread.Num_Trc = Num_Trc
                    Parametres_All_Thread.Num_Data = 0
                    Parametres_All_Thread.Nb_Calcul = nb_data_read
                    Parametres_All_Thread.voie = 2

                    If OnlyTrace = True Then
                        Parametres_All_Thread.Offset_Trc = 0
                        If Num_Trc > 0 Then
                            For K = 0 To Num_Trc - 1
                                Parametres_All_Thread.Offset_Trc = Parametres_All_Thread.Offset_Trc + Nb_Elements_Trc(K)
                            Next
                        End If
                    End If


                    If OnlyTrace = False Then
                        Tab_Thread_Calcul_Ecriture_Charge_Alive(i) = thread_tab_Calcul_Ecriture_Charge(i).IsAlive
                    Else
                        Tab_Thread_Calcul_Ecriture_Charge_Alive(i) = False
                    End If


                    If Tab_Thread_Calcul_Ecriture_Charge_Alive(i) = False And Tab_Run_Done(i) = False Then

                        Select Case Tab_Num_Trc(Num_Trc)
                            Case 0
                                File_Trc = Fichier_Trace0(i + Num_Fichier)
                            Case 1
                                File_Trc = Fichier_Trace1(i + Num_Fichier)
                            Case 2
                                File_Trc = Fichier_Trace2(i + Num_Fichier)
                            Case 3
                                File_Trc = Fichier_Trace3(i + Num_Fichier)
                            Case 4
                                File_Trc = Fichier_Trace4(i + Num_Fichier)
                            Case 5 ' 1+2
                                File_Trc = Fichier_Trace5(i + Num_Fichier)
                            Case 6 ' All
                                File_Trc = Fichier_Trace6(i + Num_Fichier)
                            Case 7 '2+3
                                File_Trc = Fichier_Trace7(i + Num_Fichier)
                            Case 8
                                File_Trc = Fichier_Trace8(i + Num_Fichier)
                            Case Else
                                File_Trc = "Fichier n° " & CStr(i + Num_Fichier)

                        End Select

                        Parametres_All_Thread.File_Name = File_Trc

                        If Use_ext_charge_Trc(Tab_Num_Trc(Num_Trc)) = False Then
                            ToolStripStatusLabel1.Text = "Second Trace processing (2)...." & File_Trc ' CStr(I + Num_Fichier)
                            Shell(Chemin_GupixWin_Multi(i) & "\pixwin.bat", AppWinStyle.Hide, False) ' Chemin_GupixWin_Multi(I), , Me.hWnd, I)
                        End If
                        Tab_Run_Done(i) = True
                    End If

                    Tab_PixWin_END(i) = File.Exists(Chemin_GupixWin_Multi(i) & "\pixwin.end")

                    If Tab_Data_Lue(i) = False And Tab_PixWin_END(i) = True And Tab_Run_Done(i) = True Then

                        thread_tab_Recup_Filter(i).Start(Parametres_All_Thread)
                        thread_tab_Element(i).Start(Parametres_All_Thread)

                        '"""Recup OXYDE
                        If mnuOxydeOUI.Checked = True Then

                            thread_tab_oxyde(i).Start(Parametres_All_Thread)
                        Else
                            Nb_Oxyde_X2 = 0
                        End If

                        Select Case Tab_Num_Trc(Num_Trc)
                            Case 0
                                File_Trc = Fichier_Trace0(i + Num_Fichier)
                            Case 1
                                File_Trc = Fichier_Trace1(i + Num_Fichier)
                            Case 2
                                File_Trc = Fichier_Trace2(i + Num_Fichier)
                            Case 3
                                File_Trc = Fichier_Trace3(i + Num_Fichier)
                            Case 4
                                File_Trc = Fichier_Trace4(i + Num_Fichier)
                            Case 5 ' 1+2
                                File_Trc = Fichier_Trace5(i + Num_Fichier)
                            Case 6 ' All
                                File_Trc = Fichier_Trace6(i + Num_Fichier)
                            Case 7 '2+3
                                File_Trc = Fichier_Trace7(i + Num_Fichier)
                            Case 8
                                File_Trc = Fichier_Trace8(i + Num_Fichier)
                            Case Else
                                File_Trc = "Fichier n° " & CStr(i + Num_Fichier)
                        End Select

                        'FitToPNG(File_Trc, I)
                        If Lect_Depth = True Then
                            '           Lect_Depth = Lit_Depth(2, I)
                        End If


                        If Adjust_Filter_B = False And Tab_Fit_Done(i) = False Then 'Tab_Thread_FitToPNG_Alive(I) = False Then 'FitToPNG(File_Trc, I)
                            Parametres_All_Thread.File_Name = File_Trc
                            thread_tab_FitToPNG_TRC(i).Start(Parametres_All_Thread)
                            Tab_Fit_Done(i) = True
                        End If

                        Tab_Data_Lue(i) = True
                        ProgressBar1.Value = ProgressBar1.Value + 1

                    End If

                    My.Application.DoEvents() : System.Threading.Thread.Sleep(20)
                    Tab_Thread_Lit_Element_Alive(i) = thread_tab_Element(i).IsAlive
                    Tab_Thread_Oxyde_Alive(i) = thread_tab_oxyde(i).IsAlive
                    Tab_Thread_Recup_Filter_Alive(i) = thread_tab_Recup_Filter(i).IsAlive

                    If Adjust_Filter_B = False Then
                        Tab_Thread_FitToPNG_Alive(i) = thread_tab_FitToPNG_TRC(i).IsAlive
                    Else
                        Tab_Thread_FitToPNG_Alive(i) = False
                    End If


                    If Tab_Thread_FitToPNG_Alive(i) = True Then ToolStripStatusLabel1.Text = "FitToPNG " & File_Trc

                    If Tab_Thread_Lit_Element_Alive(i) = False And Tab_Thread_Oxyde_Alive(i) = False And Tab_Thread_Recup_Filter_Alive(i) = False And Tab_Run_Done(i) = True And Tab_Data_Lue(i) = True And Tab_Inc_Done(i) = False And Tab_Thread_FitToPNG_Alive(i) = False Then
                        nb_data_read = nb_data_read + 1
                        Tab_Inc_Done(i) = True
                        ' Arrondi_Trc_Oxyde(i, Num_Trc, Parametres_All_Thread.Offset_Trc)

                    End If 'ALIVE = FASLE
                Else
                    nb_data_read = nb_data_read + 1
                    Tab_Inc_Done(i) = True
                End If

            Next i

        Loop While nb_data_read <> Nb_Process

    End Sub

    '***********************************************************************************************
    '****************************** SEQUENCE TRC ****************************************************
    '***********************************************************************************************


    '###################################### COUNTS / SEC.  CHARGE FILTERS TRACE   ###################################
    Sub Recup_Cps_Charge_Filters_Trc_Thread(Parametres As Struct_Parametres_Thread)
        Dim Pos_Cps As Integer

        Dim Ligne_Lu As String
        Dim i As Integer
        Dim New_Num_File As Integer
        Pos_Cps = 0
        Dim FileExist As Boolean
        Dim Voie As Integer
        Dim Num_File As Integer
        Dim Num_Proc As Integer
        Dim Num_Trc As Integer
        Dim SR As StreamReader

        Voie = Parametres.voie
        Num_Proc = Parametres.Num_Proc
        Num_File = Parametres.Num_File
        Num_Trc = Parametres.Num_Trc
        New_Num_File = Num_File

        If Num_File >= Nb_Process1 Then
            New_Num_File = Num_File - Nb_Process1
        Else
            New_Num_File = Num_File
        End If

        For i = 0 To 10
            FileExist = File.Exists(Chemin_GupixWin_Multi(Num_Proc) & "\pixstats.out")
            If FileExist = True Then Exit For
            My.Application.DoEvents() : System.Threading.Thread.Sleep(50)
        Next i

        My.Application.DoEvents() : System.Threading.Thread.Sleep(50)

        Try
            SR = File.OpenText((Chemin_GupixWin_Multi(Num_Proc) & "\pixstats.out"))
        Catch ex As Exception
            Sleep(500)
            SR = File.OpenText((Chemin_GupixWin_Multi(Num_Proc) & "\pixstats.out"))
        End Try

        For i = 0 To 9
            SR.ReadLine()
        Next i

        Ligne_Lu = SR.ReadLine()
        Pos_Cps = InStr(1, Ligne_Lu, " rate(cps):", vbTextCompare)
        Info_Experience_Trc(Num_Proc, Num_Trc).Count_Rate = Mid(Ligne_Lu, Pos_Cps + 11, 8)
        SR.ReadLine()

        Ligne_Lu = SR.ReadLine()
        Pos_Cps = InStr(1, Ligne_Lu, " Charge: ", vbTextCompare)


        Do
            Ligne_Lu = SR.ReadLine()
            Pos_Cps = InStr(1, Ligne_Lu, " Looking for", vbTextCompare)
        Loop While Pos_Cps = 0

        If mnuOxydeOUI.Checked = True Then Nb_Oxyde_X2 = Nb_Elements_Trc(Num_Trc)

        If First_Init_Trc(Num_Trc) = False Then
            First_Init_Trc(Num_Trc) = True
            Do
                Ligne_Lu = SR.ReadLine()
                Pos_Cps = InStr(1, Ligne_Lu, " Filters", vbTextCompare)
                If Pos_Cps = 0 Then Pos_Cps = InStr(1, Ligne_Lu, " Absorbers", vbTextCompare)
            Loop While Pos_Cps = 0

            ' StrFiltres_Trc(Num_Trc) = SR.ReadLine()
            Info_Experience_Trc(Num_Proc, Num_Trc).Filters = SR.ReadLine() 'Mid(Ligne_Lu, Pos_Cps + 11, 8)
            Info_Experience_Trc(0, Num_Trc).Filters = Info_Experience_Trc(Num_Proc, Num_Trc).Filters
        Else
            Info_Experience_Trc(Num_Proc, Num_Trc).Filters = Info_Experience_Trc(0, Num_Trc).Filters 'Mid(Ligne_Lu, Pos_Cps + 11, 8)
        End If

        SR.Close()


    End Sub



    Sub Filter_Create_Par_Mat(Z As String, Filter_Thickness As Single, Num_Process As Integer)
        Dim i As Integer
        Dim Str As String
        Dim Pos_Space As Integer
        Dim Nb_Elem_local As Integer
        Dim Nb_Elem_Inv As Integer
        Dim Str_Find As Integer
        Dim Nb_Filter As Integer
        Dim Tab_Filter() As String
        Dim Tab_Filter_2(,) As String
        Dim Str2_Nb_Filter As String
        Dim Nb_Filter_Total As Integer
        Dim Z_Filter As Integer
        Dim SR As StreamReader
        Dim local_str As String

        Txt_Fichier_PAR_Mat = ""
        SR = File.OpenText(Chemin_Data + "\" + Par_Mat.Text)

        Do
            Str = SR.ReadLine
            Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf
            Str_Find = InStr(1, Str, "number of x-ray absorbing filters", vbTextCompare)
        Loop While Str_Find < 1
        Nb_Filter = CInt(Strings.Left(Str, 2))

        ReDim Tab_Filter(Nb_Filter - 1)
        ReDim Tab_Filter_2(Nb_Filter - 1, 2)
        Str2_Nb_Filter = "(mm if air), then % hole area)"

        Nb_Filter_Total = 0

        For i = 0 To Nb_Filter - 1
            Tab_Filter(i) = SR.ReadLine
            Nb_Filter_Total = Nb_Filter_Total + 1
            Str2_Nb_Filter = SR.ReadLine
        Next i

        For i = 0 To Nb_Filter_Total - 1
            Pos_Space = InStr(1, Tab_Filter(i), " ", vbTextCompare)
            Z_Filter = Strings.Left(Tab_Filter(i), Pos_Space - 1)
            local_str = Strings.Replace(CStr(Filter_Thickness), ",", ".")
            If Z_Filter = Z Then
                Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) & Z & " " & local_str & " 0" & vbTab & "(One line for each filter, giving Z, then thickness in microns" & vbCrLf
                Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) & vbTab & vbTab & Str2_Nb_Filter & vbCrLf
            Else
                Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) & Tab_Filter(i) & vbCrLf
                Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) & vbTab & vbTab & Str2_Nb_Filter & vbCrLf
            End If
        Next i

        Do
            Str = SR.ReadLine
            Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf
        Loop While Str <> "----------Define fit elements----------"

        Str = SR.ReadLine
        Nb_Elem_local = CInt(Strings.Left(Str, 2))
        Nb_Elements_Mat = Nb_Elem_local
        Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf


        Do
            Str = SR.ReadLine
            Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf
        Loop While Str <> "----------Define invisible elements----------"

        Str = SR.ReadLine
        Nb_Elem_Inv = CInt(Strings.Left(Str, 2))
        Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf

        Do
            Str = SR.ReadLine
            Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str
            If Str <> "----------Spectral Data-----------" Then Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + vbCrLf

        Loop While Str <> "----------Spectral Data-----------"
        SR.Close()

    End Sub

    Sub HED_Create_Par_Mat(Z As String, Filter_Thickness As Single, Num_Process As Integer)
        Dim i As Integer
        Dim Str As String
        Dim Pos_Space As Integer
        Dim Nb_Elem_local As Integer
        Dim Nb_Elem_Inv As Integer
        Dim Str_Find As Integer
        Dim Nb_Filter As Integer
        Dim Tab_Filter() As String
        Dim Tab_Filter_2(,) As String
        Dim Str2_Nb_Filter As String
        Dim Nb_Filter_Total As Integer
        Dim Z_Filter As Integer
        Dim SR As StreamReader
        Dim local_str As String

        Txt_Fichier_PAR_Mat = ""
        SR = File.OpenText(Chemin_Data + "\" + Par_Mat.Text)

        Do
            Str = SR.ReadLine
            Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf
            Str_Find = InStr(1, Str, "number of x-ray absorbing filters", vbTextCompare)
        Loop While Str_Find < 1
        Nb_Filter = CInt(Strings.Left(Str, 2))

        ReDim Tab_Filter(Nb_Filter - 1)
        ReDim Tab_Filter_2(Nb_Filter - 1, 2)
        Str2_Nb_Filter = "(mm if air), then % hole area)"

        Nb_Filter_Total = 0

        For i = 0 To Nb_Filter - 1
            Tab_Filter(i) = SR.ReadLine
            Nb_Filter_Total = Nb_Filter_Total + 1
            Str2_Nb_Filter = SR.ReadLine

        Next i

        For i = 0 To Nb_Filter_Total - 1
            Pos_Space = InStr(1, Tab_Filter(i), " ", vbTextCompare)
            Z_Filter = Strings.Left(Tab_Filter(i), Pos_Space - 1)
            local_str = Strings.Replace(CStr(Filter_Thickness), ",", ".")
            If Z_Filter = Z Then
                Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) & Z & " " & local_str & " 0" & vbTab & "(One line for each filter, giving Z, then thickness in microns" & vbCrLf
                Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) & vbTab & vbTab & Str2_Nb_Filter & vbCrLf
            Else
                Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) & Tab_Filter(i) & vbCrLf
                Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) & vbTab & vbTab & Str2_Nb_Filter & vbCrLf
            End If
        Next i

        Do
            Str = SR.ReadLine
            Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf
        Loop While Str <> "----------Define fit elements----------"

        Str = SR.ReadLine
        Nb_Elem_local = CInt(Strings.Left(Str, 2))
        Nb_Elements_Mat = Nb_Elem_local
        Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf


        Do
            Str = SR.ReadLine
            Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf
        Loop While Str <> "----------Define invisible elements----------"

        Str = SR.ReadLine
        Nb_Elem_Inv = CInt(Strings.Left(Str, 2))
        Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str + vbCrLf

        Do
            Str = SR.ReadLine
            Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + Str
            If Str <> "----------Spectral Data-----------" Then Txt_Fichier_PAR_Mat_Filter(Num_Process) = Txt_Fichier_PAR_Mat_Filter(Num_Process) + vbCrLf

        Loop While Str <> "----------Spectral Data-----------"
        SR.Close()

    End Sub


    Sub Filter_Create_Par_Trc(Z As String, Filter_Thickness As Single, Num_Process As Integer, Num_trc As Integer)
        Dim i As Integer
        Dim Str As String
        Dim Pos_Space As Integer
        Dim Str_Find As Integer
        Dim Nb_Filter As Integer
        Dim Tab_Filter() As String
        Dim Tab_Filter_2(,) As String
        Dim Str2_Nb_Filter As String
        Dim Nb_Filter_Total As Integer
        Dim Z_Filter As Integer
        Dim local_str As String
        Dim SR As StreamReader

        SR = File.OpenText(Chemin_Data + "\" + Tab_File_Par_Trc(Num_trc))

        Do
            Str = SR.ReadLine
            Txt_Fichier_PAR_Trc_Filter(Num_Process) = Txt_Fichier_PAR_Trc_Filter(Num_Process) + Str + vbCrLf
            Str_Find = InStr(1, Str, "number of x-ray absorbing filters", vbTextCompare)
        Loop While Str_Find < 1
        Nb_Filter = CInt(Strings.Left(Str, 2))

        ReDim Tab_Filter(Nb_Filter - 1)
        ReDim Tab_Filter_2(Nb_Filter - 1, 2)
        Str2_Nb_Filter = "(mm if air), then % hole area)"

        Nb_Filter_Total = 0

        For i = 0 To Nb_Filter - 1
            Tab_Filter(i) = SR.ReadLine
            Nb_Filter_Total = Nb_Filter_Total + 1
            Str2_Nb_Filter = SR.ReadLine
        Next i



        For i = 0 To Nb_Filter_Total - 1
            Pos_Space = InStr(1, Tab_Filter(i), " ", vbTextCompare)
            Z_Filter = Strings.Left(Tab_Filter(i), Pos_Space - 1)
            local_str = Strings.Replace(CStr(Filter_Thickness), ",", ".")
            If Z_Filter = Z Then
                Txt_Fichier_PAR_Trc_Filter(Num_Process) = Txt_Fichier_PAR_Trc_Filter(Num_Process) & Z & " " & local_str & " 0" & vbCrLf
                Txt_Fichier_PAR_Trc_Filter(Num_Process) = Txt_Fichier_PAR_Trc_Filter(Num_Process) & vbTab & vbTab & Str2_Nb_Filter & vbCrLf
            Else
                Txt_Fichier_PAR_Trc_Filter(Num_Process) = Txt_Fichier_PAR_Trc_Filter(Num_Process) & Tab_Filter(i) & vbCrLf
                Txt_Fichier_PAR_Trc_Filter(Num_Process) = Txt_Fichier_PAR_Trc_Filter(Num_Process) & vbTab & vbTab & Str2_Nb_Filter & vbCrLf
            End If

        Next i


        Do
            Str = SR.ReadLine 'Str = SplitText(Num_line)
            Txt_Fichier_PAR_Trc_Filter(Num_Process) = Txt_Fichier_PAR_Trc_Filter(Num_Process) + Str
            If Str <> "----------Spectral Data-----------" Then Txt_Fichier_PAR_Trc_Filter(Num_Process) = Txt_Fichier_PAR_Trc_Filter(Num_Process) + vbCrLf

        Loop While Str <> "----------Spectral Data-----------"
        SR.Close()

    End Sub


    Sub Lecture_Fichier_Par_Mat()
        Dim i As Integer
        Dim Str As String
        Dim Pos_Coma As Integer
        Dim Pos_Depth As Integer
        Dim Pos_Coma1 As Integer
        Dim Pos_Coma2 As Integer
        Dim Nb_Elem_local As Integer
        Dim Nb_Elem_Inv As Integer
        Dim NL_Or_Filter As String
        Dim SR As StreamReader
        Dim All_line
        Dim indx_1 As Integer
        Dim pos_HED As Integer
        Dim str_HED As String
        Dim info_HED() As String
        Dim split_HED_K() As String
        Dim split_HED_L() As String
        Dim split_HED_M() As String



        ToolStripStatusLabel1.Text = "Read Matrix PAR file"
        Txt_Fichier_PAR_Mat = ""

        Try
            My.Computer.FileSystem.CopyFile(Chemin_Data + "\" + Par_Mat.Text, Chemin_Processed_Data & Par_Mat.Text, True)
        Catch ex As Exception
            MsgBox("Fatal Error, please check Matrix spectra file '" + Par_Mat.Text & "' is present into the data folder", MsgBoxStyle.Information, "Error reading matrix parameter")
            Fatal_Error = True
        End Try


        All_line = IO.File.ReadAllLines(Chemin_Data + "\" + Par_Mat.Text)

        indx_1 = Array.IndexOf(All_line, "----------Define fit elements----------")
        Pos_Depth = Array.IndexOf(All_line, "penetration depths ")

        '############   Search for file HED 
        pos_HED = Array.IndexOf(All_line, vbTab & "(1st entry is constant H-value or -1 if H is energy-dependent.")
        str_HED = All_line(pos_HED - 1)
        info_HED = Split(str_HED, " ")
        K_HED_Mat = False
        L_HED_Mat = False
        M_HED_Mat = False
        Use_HED_Mat = False
        If info_HED(0) = "-1" Then 'HED used
            Use_HED_Mat = True
            K_Path_HED_Mat = Strings.Replace(info_HED(1), """", "") 'take out character " from path
            If K_Path_HED_Mat <> "" Then
                split_HED_K = Split(K_Path_HED_Mat, "\")
                K_Name_HED_Mat = split_HED_K(UBound(split_HED_K))
                K_HED_Mat = True
                K_Path_HED_Mat = Chemin_Data & "\" & K_Name_HED_Mat
            End If

            If UBound(info_HED) >= 2 And info_HED(2) <> "" Then
                L_Path_HED_Mat = Strings.Replace(info_HED(2), """", "") 'take out character " from path
                split_HED_L = Split(L_Path_HED_Mat, "\")
                L_Name_HED_Mat = split_HED_L(UBound(split_HED_L))
                L_HED_Mat = True
                L_Path_HED_Mat = Chemin_Data & "\" & L_Name_HED_Mat
            End If

            If UBound(info_HED) >= 3 And info_HED(3) <> "" Then
                M_Path_HED_Mat = Strings.Replace(info_HED(3), """", "") 'take out character " from path
                split_HED_M = Split(M_Path_HED_Mat, "\")
                M_Name_HED_Mat = split_HED_M(UBound(split_HED_M))
                M_HED_Mat = True
                M_Path_HED_Mat = Chemin_Data & "\" & M_Name_HED_Mat
            End If

        End If

        If Pos_Depth > 0 Then Lect_Depth = True

        For i = 0 To indx_1  'read until "Define fit elements"
            Txt_Fichier_PAR_Mat = Txt_Fichier_PAR_Mat + All_line(i) + vbCrLf
        Next

        Nb_Elem_local = CInt(Strings.Left(All_line(indx_1 + 1), 2))
        Nb_Elements_Mat = Nb_Elem_local

        For i = indx_1 + 1 To indx_1 + 7
            Txt_Fichier_PAR_Mat = Txt_Fichier_PAR_Mat + All_line(i) + vbCrLf
        Next

        For i = indx_1 + 8 To indx_1 + 8 + Nb_Elem_local - 1
            Str = All_line(i) 'SR.ReadLine()
            Txt_Fichier_PAR_Mat = Txt_Fichier_PAR_Mat + Str + vbCrLf
            Pos_Coma1 = 0
            Pos_Coma2 = 0
            Pos_Coma = InStr(1, Str, ",", vbTextCompare)
            If Pos_Coma > 0 Then Pos_Coma1 = InStr(Pos_Coma + 1, Str, ",", vbTextCompare)
            If Pos_Coma1 > 0 Then Pos_Coma2 = InStr(Pos_Coma1 + 1, Str, ",", vbTextCompare)
            If Pos_Coma > 1 Then Nb_Elements_Mat = Nb_Elements_Mat + 1
            If Pos_Coma1 > 1 Then Nb_Elements_Mat = Nb_Elements_Mat + 1
            If Pos_Coma2 > 1 Then Nb_Elements_Mat = Nb_Elements_Mat + 1
            NL_Or_Filter = Mid(Str, 4, 2)
            If NL_Or_Filter = "8 " Then 'Si element est définie comme FILTER l'élément ne donne pas de résulat
                Nb_Elements_Mat = Nb_Elements_Mat - 1
                Nb_Elem_local = Nb_Elem_local - 1
            End If
        Next


        indx_1 = indx_1 + 8 + Nb_Elem_local
        Do
            Str = All_line(indx_1) 'SR.ReadLine()
            Txt_Fichier_PAR_Mat = Txt_Fichier_PAR_Mat + Str + vbCrLf
            My.Application.DoEvents()
            indx_1 += 1
        Loop While Str <> "----------Define invisible elements----------"

        Str = All_line(indx_1) 'SR.ReadLine()
        indx_1 += 1
        Nb_Elem_Inv = CInt(Strings.Left(Str, 2))
        Txt_Fichier_PAR_Mat = Txt_Fichier_PAR_Mat + Str + vbCrLf

        For i = indx_1 To indx_1 + 6 'Passe les lignes de commentaires
            Str = All_line(indx_1) 'SR.ReadLine()
            Txt_Fichier_PAR_Mat = Txt_Fichier_PAR_Mat + Str + vbCrLf
            indx_1 += 1
        Next
        Str = All_line(indx_1) 'SR.ReadLine()
        indx_1 += 1
        Txt_Fichier_PAR_Mat = Txt_Fichier_PAR_Mat + Str + vbCrLf

        If Nb_Elem_Inv <> -1 Then
            Z_Elem_Inv = CInt(Strings.Left(Str, 2))
        End If

        If Nb_Elem_Inv = -1 Or CInt(Z_Elem_Inv) <> 8 And mnuOxydeOUI.Checked = True Then
            mnuOxydeNON_Click()
        End If

        Do
            Str = All_line(indx_1) 'SR.ReadLine()
            Txt_Fichier_PAR_Mat = Txt_Fichier_PAR_Mat + Str
            If Str <> "----------Spectral Data-----------" Then Txt_Fichier_PAR_Mat = Txt_Fichier_PAR_Mat + vbCrLf
            indx_1 += 1
        Loop While Str <> "----------Spectral Data-----------"


    End Sub

    Sub Create_Fichier_Par_Mat_HED()
        Dim i As Integer
        Dim Str As String
        Dim All_line
        Dim indx_1 As Integer
        Dim str_HED As String
        Dim Local_Text As String

        Local_Text = ""
        ToolStripStatusLabel1.Text = "Create Matrix PAR file with HED"
        All_line = IO.File.ReadAllLines(Chemin_Data & "\" + Par_Mat.Text)
        indx_1 = Array.IndexOf(All_line, vbTab & "(1st entry is constant H-value or -1 if H is energy-dependent.")

        For i = 0 To indx_1 - 2 'read until "Define fit elements"
            Local_Text = Local_Text & All_line(i) & vbCrLf
        Next
        str_HED = "-1 "

        For j = 0 To Nb_Process - 1
            str_HED = "-1 "
            If K_HED_Mat = True Then
                str_HED = str_HED & """" & Chemin_GupixWin_Multi(j) & "\" & K_Name_HED_Mat & """ "
                Try
                    My.Computer.FileSystem.CopyFile(K_Path_HED_Mat, Chemin_GupixWin_Multi(j) & "\" & K_Name_HED_Mat, True)
                Catch ex As Exception
                    My.Computer.FileSystem.CopyFile(K_Path_HED_Mat, Chemin_GupixWin_Multi(j) & "\" & K_Name_HED_Mat, True)
                End Try
            Else
                str_HED = str_HED + " "" "
            End If

            If L_HED_Mat = True Then
                str_HED = str_HED & """" & Chemin_GupixWin_Multi(j) & "\" & L_Name_HED_Mat & """ "
                Try
                    My.Computer.FileSystem.CopyFile(L_Path_HED_Mat, Chemin_GupixWin_Multi(j) & "\" & L_Name_HED_Mat, True)
                Catch ex As Exception
                    My.Computer.FileSystem.CopyFile(L_Path_HED_Mat, Chemin_GupixWin_Multi(j) & "\" & L_Name_HED_Mat, True)
                End Try
            Else
                str_HED = str_HED + " "" "
            End If

            If M_HED_Mat = True Then
                str_HED = str_HED & """" & Chemin_GupixWin_Multi(j) & "\" & M_Name_HED_Mat & """ "
                Try
                    My.Computer.FileSystem.CopyFile(M_Path_HED_Mat, Chemin_GupixWin_Multi(j) & "\" & M_Name_HED_Mat, True)
                Catch ex As Exception
                    My.Computer.FileSystem.CopyFile(M_Path_HED_Mat, Chemin_GupixWin_Multi(j) & "\" & M_Name_HED_Mat, True)
                End Try
            Else
                str_HED = str_HED + " "" "
            End If
            Txt_Fichier_PAR_Mat_HED(j) = Local_Text & str_HED & vbCrLf
        Next

        Local_Text = ""

        Do
            Str = All_line(indx_1) 'SR.ReadLine()
            Local_Text = Local_Text + Str
            If Str <> "----------Spectral Data-----------" Then Local_Text = Local_Text & vbCrLf
            indx_1 += 1
        Loop While Str <> "----------Spectral Data-----------"

        For j = 0 To Nb_Process - 1
            Txt_Fichier_PAR_Mat_HED(j) = Txt_Fichier_PAR_Mat_HED(j) & Local_Text
        Next



    End Sub

    Sub Create_Fichier_Par_Trc_HED(num_trc As Integer)
        Dim i As Integer
        Dim Str As String
        Dim All_line
        Dim indx_1 As Integer
        Dim str_HED As String
        Dim Local_Text As String

        Local_Text = ""
        ToolStripStatusLabel1.Text = "Create trace PAR file with HED"
        All_line = IO.File.ReadAllLines(Chemin_Data & "\" + Tab_File_Par_Trc(num_trc))
        indx_1 = Array.IndexOf(All_line, vbTab & "(1st entry is constant H-value or -1 if H is energy-dependent.")

        For i = 0 To indx_1 - 2 'read until "Define fit elements"
            Local_Text = Local_Text & All_line(i) & vbCrLf
        Next
        str_HED = "-1 "

        For j = 0 To Nb_Process - 1
            If K_HED_Trc(j) = True Then
                str_HED = str_HED & """" & Chemin_GupixWin_Multi(j) & "\" & K_Name_HED_Trc(j) & """ "
                Try
                    My.Computer.FileSystem.CopyFile(K_Path_HED_Trc(j), Chemin_GupixWin_Multi(j) & "\" & K_Name_HED_Trc(j), True)
                Catch ex As Exception
                    My.Computer.FileSystem.CopyFile(K_Path_HED_Trc(j), Chemin_GupixWin_Multi(j) & "\" & K_Name_HED_Trc(j), True)
                End Try
            Else
                str_HED = str_HED + " "" "
            End If

            If L_HED_Trc(j) = True Then
                str_HED = str_HED & """" & Chemin_GupixWin_Multi(j) & "\" & L_Name_HED_Trc(j) & """ "
                Try
                    My.Computer.FileSystem.CopyFile(L_Path_HED_Trc(j), Chemin_GupixWin_Multi(j) & "\" & L_Name_HED_Trc(j), True)
                Catch ex As Exception
                    My.Computer.FileSystem.CopyFile(L_Path_HED_Trc(j), Chemin_GupixWin_Multi(j) & "\" & L_Name_HED_Trc(j), True)
                End Try
            Else
                str_HED = str_HED + " "" "
            End If

            If K_HED_Trc(j) = True Then
                str_HED = str_HED & """" & Chemin_GupixWin_Multi(j) & "\" & M_Name_HED_Trc(j) & """ "
                Try
                    My.Computer.FileSystem.CopyFile(M_Path_HED_Trc(j), Chemin_GupixWin_Multi(j) & "\" & M_Name_HED_Trc(j), True)
                Catch ex As Exception
                    My.Computer.FileSystem.CopyFile(M_Path_HED_Trc(j), Chemin_GupixWin_Multi(j) & "\" & M_Name_HED_Trc(j), True)
                End Try
            Else
                str_HED = str_HED + " "" "
            End If
            Txt_Fichier_PAR_Trc_HED(j, num_trc) = Local_Text & str_HED & vbCrLf
        Next

        Local_Text = ""

        Do
            Str = All_line(indx_1) 'SR.ReadLine()
            Local_Text = Local_Text + Str
            If Str <> "----------Spectral Data-----------" Then Local_Text = Local_Text & vbCrLf
            indx_1 += 1
        Loop While Str <> "----------Spectral Data-----------"

        For j = 0 To Nb_Process - 1
            Txt_Fichier_PAR_Mat_HED(j) = Txt_Fichier_PAR_Mat_HED(j) & Local_Text
        Next

    End Sub


    Sub Creer_tab_trc()

        LstPar_Trc.Refresh()
        LstPar_Mat.Refresh()

        Nb_Trc = 0

        If Check_det0.Checked = True And Par_det0.Text <> "" Then
            Tab_Num_Trc(Nb_Trc) = 0
            Tab_File_Par_Trc(Nb_Trc) = Par_det0.Text
            Ref_DataSet_ToRead(Nb_Trc) = Ext_Trc0 'Par_det0.Text 'Ext_Trc0
            Nb_Trc = Nb_Trc + 1

        Else
            Tab_File_Par_Trc(Nb_Trc) = ""
            Par_det0.Text = ""
        End If

        If Check_det1.Checked = True And Par_det1.Text <> "" Then
            Tab_Num_Trc(Nb_Trc) = 1
            Tab_File_Par_Trc(Nb_Trc) = Par_det1.Text
            Ref_DataSet_ToRead(Nb_Trc) = Ext_Trc1 'Par_det1.Text 'Ext_Trc1 '"x1"
            Nb_Trc = Nb_Trc + 1
        Else
            Tab_File_Par_Trc(Nb_Trc) = ""
            Par_det1.Text = ""
        End If

        If Check_det2.Checked = True And Par_det2.Text <> "" Then
            Tab_Num_Trc(Nb_Trc) = 2
            Tab_File_Par_Trc(Nb_Trc) = Par_det2.Text
            Ref_DataSet_ToRead(Nb_Trc) = Ext_Trc2
            Nb_Trc = Nb_Trc + 1 '"x2"
        Else
            Tab_File_Par_Trc(Nb_Trc) = ""
            Par_det2.Text = ""
        End If

        If Check_det3.Checked = True And Par_det3.Text <> "" Then
            Tab_Num_Trc(Nb_Trc) = 3
            Tab_File_Par_Trc(Nb_Trc) = Par_det3.Text
            Ref_DataSet_ToRead(Nb_Trc) = Ext_Trc3 '"x3"
            Nb_Trc = Nb_Trc + 1

        Else
            Tab_File_Par_Trc(Nb_Trc) = ""
            Par_det3.Text = ""
        End If

        If Check_det4.Checked = True And Par_det4.Text <> "" Then
            Tab_Num_Trc(Nb_Trc) = 4
            Tab_File_Par_Trc(Nb_Trc) = Par_det4.Text
            Ref_DataSet_ToRead(Nb_Trc) = Ext_Trc4 ' "x4"
            Nb_Trc = Nb_Trc + 1
        Else
            Tab_File_Par_Trc(Nb_Trc) = ""
            Par_det4.Text = ""
        End If

        If Check_det5.Checked = True And Par_det5.Text <> "" Then
            Tab_Num_Trc(Nb_Trc) = 5
            Tab_File_Par_Trc(Nb_Trc) = Par_det5.Text
            Ref_DataSet_ToRead(Nb_Trc) = Ext_Trc5 '"x10"
            Nb_Trc = Nb_Trc + 1
        Else
            Tab_File_Par_Trc(Nb_Trc) = ""
            Par_det5.Text = ""
        End If

        If Check_det6.Checked = True And Par_det6.Text <> "" Then
            Tab_Num_Trc(Nb_Trc) = 6
            Tab_File_Par_Trc(Nb_Trc) = Par_det6.Text
            Ref_DataSet_ToRead(Nb_Trc) = Ext_Trc6 '"x11"
            Nb_Trc = Nb_Trc + 1
        Else
            Tab_File_Par_Trc(Nb_Trc) = ""
            Par_det6.Text = ""
        End If

        If Check_det7.Checked = True And Par_det7.Text <> "" Then
            Tab_Num_Trc(Nb_Trc) = 7
            Tab_File_Par_Trc(Nb_Trc) = Par_det7.Text
            Ref_DataSet_ToRead(Nb_Trc) = Ext_Trc7 '7"x12"
            Nb_Trc = Nb_Trc + 1
        Else
            Tab_File_Par_Trc(Nb_Trc) = ""
            Par_det7.Text = ""
        End If

        If Check_det8.Checked = True And Par_det8.Text <> "" Then
            Tab_Num_Trc(Nb_Trc) = 8
            Tab_File_Par_Trc(Nb_Trc) = Par_det8.Text
            Ref_DataSet_ToRead(Nb_Trc) = Ext_Trc8 '"x13"
            Nb_Trc = Nb_Trc + 1
        Else
            Tab_File_Par_Trc(Nb_Trc) = ""
            Par_det8.Text = ""
        End If

    End Sub

    Sub Lecture_Fichier_Par_Trc_OLD(partie As Integer, Num_Trc As Integer)

        Dim i As Integer
        Dim Str As String
        Dim MyInStr As Integer
        Dim Nb_Elem_local As Integer
        Dim Pos_Coma As Integer
        Dim Pos_Coma1 As Integer
        Dim Pos_Coma2 As Integer
        Dim NL_Or_Filter As String
        Dim SR As StreamReader
        Dim All_line
        Dim indx_1 As Integer
        Dim Pos_Charge As Integer

        ToolStripStatusLabel1.Text = "Read Trace PAR File " & CStr(Num_Trc) & ",Part:" & CStr(partie)

        Try
            My.Computer.FileSystem.CopyFile(Chemin_Data + "\" + Tab_File_Par_Trc(Num_Trc), Chemin_Processed_Data & Tab_File_Par_Trc(Num_Trc), True)
        Catch ex As Exception

        End Try

        SR = File.OpenText(Chemin_Data + "\" + Tab_File_Par_Trc(Num_Trc))
        Nb_Ligne_Trc = 0
        All_line = IO.File.ReadAllLines(Chemin_Data + "\" + Tab_File_Par_Trc(Num_Trc))

        Select Case partie

            Case 1
                indx_1 = 0
                Do
                    Str = All_line(indx_1)
                    MyInStr = InStr(1, Str, "Beam Charge", vbTextCompare) 'Lit Beam Charge initial
                    Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc & Str & vbCrLf
                    Nb_Ligne_Trc = Nb_Ligne_Trc + 1
                    If MyInStr > 0 Then Val_Charge_Trc_Init = Val(Str)
                    My.Application.DoEvents()
                    indx_1 += 1
                Loop While Str <> "----------Define matrix elements----------"

            Case 2


                Do
                    Str = SR.ReadLine()
                    My.Application.DoEvents()
                Loop While Str <> "----------Define trace elements----------"

                Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + Str + vbCrLf
                Str = SR.ReadLine()

                Do
                    Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + Str
                    If Str <> "----------Spectral Data-----------" Then Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + vbCrLf
                    Str = SR.ReadLine()

                    My.Application.DoEvents()
                Loop While Str <> "----------Spectral Data-----------"
                Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + "----------Spectral Data-----------" '+ vbCrLf


            Case 3
                Str = SR.ReadLine()
                Do
                    Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + Str
                    If Str <> "----------Spectral Data-----------" Then Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + vbCrLf
                    Str = SR.ReadLine()
                    My.Application.DoEvents()
                Loop While Str <> "----------Spectral Data-----------"
                Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + "----------Spectral Data-----------" '+ vbCrLf

            Case 4

                indx_1 = Array.IndexOf(All_line, "----------Define trace elements----------")

                Str = All_line(indx_1 + 1) 'SR.ReadLine()
                Nb_Elements_Trc(Num_Trc) = CInt(Strings.Left(Str, 2))
                Nb_Elem_local = Nb_Elements_Trc(Num_Trc)

                indx_1 += 12

                For i = 0 To Nb_Elem_local - 1
                    Str = All_line(indx_1) 'SR.ReadLine()
                    indx_1 += 1
                    Pos_Coma1 = 0
                    Pos_Coma2 = 0
                    Pos_Coma = InStr(1, Str, ",", vbTextCompare)
                    If Pos_Coma > 0 Then Pos_Coma1 = InStr(Pos_Coma + 1, Str, ",", vbTextCompare)
                    If Pos_Coma1 > 0 Then Pos_Coma2 = InStr(Pos_Coma1 + 1, Str, ",", vbTextCompare)
                    If Pos_Coma > 1 Then Nb_Elements_Trc(Num_Trc) = Nb_Elements_Trc(Num_Trc) + 1
                    If Pos_Coma1 > 1 Then Nb_Elements_Trc(Num_Trc) = Nb_Elements_Trc(Num_Trc) + 1
                    If Pos_Coma2 > 1 Then Nb_Elements_Trc(Num_Trc) = Nb_Elements_Trc(Num_Trc) + 1
                    NL_Or_Filter = Mid(Str, 4, 2)
                    If NL_Or_Filter = "8 " Then 'Si element est définie comme FILTER l'élément ne donne pas de résulat
                        Nb_Elements_Trc(Num_Trc) = Nb_Elements_Trc(Num_Trc) - 1
                        Nb_Elem_local = Nb_Elem_local - 1
                    End If


                Next
        End Select

        SR.Close()
    End Sub

    Sub copy_par_file_trc(Num_trc As Integer)

        Try
            My.Computer.FileSystem.CopyFile(Chemin_Data + "\" + Tab_File_Par_Trc(Num_trc), Chemin_Processed_Data & Tab_File_Par_Trc(Num_trc), True)
        Catch ex As Exception
            MsgBox("Fatal Error, please check trace parameter file '" + Tab_File_Par_Trc(Num_trc) & "'is present into the data folder.", MsgBoxStyle.Information, "Error reading trace parameter file")
            Fatal_Error = True
            Exit Sub
        End Try

    End Sub

    Sub Lecture_par_trc_HED_NbElem(Num_trc As Integer)
        Dim pos_HED As Integer
        Dim str_HED As String
        Dim info_HED() As String
        Dim split_HED_K() As String
        Dim split_HED_L() As String
        Dim split_HED_M() As String
        Dim SR As StreamReader
        Dim All_line
        Dim indx_1 As Integer
        Dim Str As String
        Dim Nb_Elem_local As Integer
        Dim Pos_Coma As Integer
        Dim Pos_Coma1 As Integer
        Dim Pos_Coma2 As Integer
        Dim NL_Or_Filter As String
        Dim nb_parasite As Integer


        SR = File.OpenText(Chemin_Data + "\" + Tab_File_Par_Trc(Num_trc))
        Nb_Ligne_Trc = 0

        All_line = IO.File.ReadAllLines(Chemin_Data + "\" + Tab_File_Par_Trc(Num_trc))

        '############   Search for file HED 
        pos_HED = Array.IndexOf(All_line, vbTab & "(1st entry is constant H-value or -1 if H is energy-dependent.")
        str_HED = All_line(pos_HED - 1)
        info_HED = Split(str_HED, " ")
        K_HED_Trc(Num_trc) = False
        L_HED_Trc(Num_trc) = False
        M_HED_Trc(Num_trc) = False
        Use_HED_Trc(Num_trc) = False



        If info_HED(0) = "-1" Then 'HED used
            Use_HED_Trc(Num_trc) = True
            K_Path_HED_Trc(Num_trc) = Strings.Replace(info_HED(1), """", "") 'take out character " from path
            If K_Path_HED_Trc(Num_trc) <> "" Then
                split_HED_K = Split(K_Path_HED_Trc(Num_trc), "\")
                K_Name_HED_Trc(Num_trc) = split_HED_K(3)
                K_Name_HED_Trc(Num_trc) = Strings.Left(K_Name_HED_Trc(Num_trc), Len(K_Name_HED_Trc(Num_trc)) - 1)
                K_HED_Trc(Num_trc) = True
            End If

            If UBound(info_HED) >= 2 And info_HED(2) <> "" Then
                L_Path_HED_Trc(Num_trc) = Strings.Replace(info_HED(2), """", "") 'take out character " from path
                split_HED_L = Split(L_Path_HED_Trc(Num_trc), "\")
                L_Name_HED_Trc(Num_trc) = split_HED_L(3)
                L_Name_HED_Trc(Num_trc) = Strings.Left(L_Name_HED_Trc(Num_trc), Len(L_Name_HED_Trc(Num_trc)) - 1)
                L_HED_Trc(Num_trc) = True
            End If

            If UBound(info_HED) >= 3 And info_HED(3) <> "" Then
                M_Path_HED_Trc(Num_trc) = Strings.Replace(info_HED(3), """", "") 'take out character " from path
                split_HED_M = Split(M_Path_HED_Trc(Num_trc), "\")
                M_Name_HED_Trc(Num_trc) = split_HED_M(3)
                M_Name_HED_Trc(Num_trc) = Strings.Left(M_Name_HED_Trc(Num_trc), Len(M_Name_HED_Trc(Num_trc)) - 1)
                M_HED_Trc(Num_trc) = True
            End If

        End If


        indx_1 = Array.IndexOf(All_line, "----------Define trace elements----------")
        Str = All_line(indx_1 + 1) 'SR.ReadLine()
        Nb_Elements_Trc(Num_trc) = CInt(Strings.Left(Str, 2))
        Nb_Elem_local = Nb_Elements_Trc(Num_trc)
        indx_1 += 12 ' Skip 12 line
        nb_parasite = 1
        For i = 0 To Nb_Elem_local - 1
            Str = All_line(indx_1) 'SR.ReadLine()
            indx_1 += 1
            Pos_Coma1 = 0
            Pos_Coma2 = 0
            nb_parasite = 1
            Pos_Coma = InStr(1, Str, ",", vbTextCompare)
            If Pos_Coma > 0 Then Pos_Coma1 = InStr(Pos_Coma + 1, Str, ",", vbTextCompare)
            If Pos_Coma1 > 0 Then Pos_Coma2 = InStr(Pos_Coma1 + 1, Str, ",", vbTextCompare)
            If Pos_Coma > 1 Then
                Nb_Elements_Trc(Num_trc) = Nb_Elements_Trc(Num_trc) + 1
                nb_parasite += 1
            End If

            If Pos_Coma1 > 1 Then
                Nb_Elements_Trc(Num_trc) = Nb_Elements_Trc(Num_trc) + 1
                nb_parasite += 1
            End If
            If Pos_Coma2 > 1 Then
                Nb_Elements_Trc(Num_trc) = Nb_Elements_Trc(Num_trc) + 1
                nb_parasite += 1
            End If
            NL_Or_Filter = Mid(Str, 4, 2)
            If NL_Or_Filter = "8 " Then 'Si element est définie comme FILTER l'élément ne donne pas de résulat
                Nb_Elements_Trc(Num_trc) = Nb_Elements_Trc(Num_trc) - nb_parasite
                Nb_Elem_local = Nb_Elem_local - nb_parasite
            End If


        Next

    End Sub

    Sub Lecture_Fichier_Par_Trc(partie As Integer, Num_Trc As Integer)

        Dim Str As String
        Dim MyInStr As Integer
        Dim SR As StreamReader
        Dim All_line
        Dim indx_1 As Integer




        ToolStripStatusLabel1.Text = "Read Trace PAR File " & CStr(Num_Trc) & ",Part:" & CStr(partie)

        SR = File.OpenText(Chemin_Data + "\" + Tab_File_Par_Trc(Num_Trc))
        Nb_Ligne_Trc = 0

        All_line = IO.File.ReadAllLines(Chemin_Data + "\" + Tab_File_Par_Trc(Num_Trc))


        Select Case partie

            Case 1
                indx_1 = 0
                Do
                    Str = All_line(indx_1)
                    MyInStr = InStr(1, Str, "Beam Charge", vbTextCompare) 'Lit Beam Charge initial
                    Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc & Str & vbCrLf
                    Nb_Ligne_Trc = Nb_Ligne_Trc + 1
                    If MyInStr > 0 Then Val_Charge_Trc_Init = Val(Str)
                    My.Application.DoEvents()
                    indx_1 += 1
                Loop While Str <> "----------Define matrix elements----------"

            Case 2

                Do
                    Str = SR.ReadLine()
                    My.Application.DoEvents()
                Loop While Str <> "----------Define trace elements----------"

                Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + Str + vbCrLf
                Str = SR.ReadLine()

                Do
                    Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + Str
                    If Str <> "----------Spectral Data-----------" Then Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + vbCrLf
                    Str = SR.ReadLine()

                    My.Application.DoEvents()
                Loop While Str <> "----------Spectral Data-----------"
                Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + "----------Spectral Data-----------" '+ vbCrLf


            Case 3
                Str = SR.ReadLine()
                Do
                    Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + Str
                    If Str <> "----------Spectral Data-----------" Then Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + vbCrLf
                    Str = SR.ReadLine()
                    My.Application.DoEvents()
                Loop While Str <> "----------Spectral Data-----------"
                Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + "----------Spectral Data-----------" '+ vbCrLf

        End Select

        SR.Close()
    End Sub




    Private Sub Calcul_Ecriture_Charge_Thread(Parametres As Struct_Parametres_Thread) ' Mycharge As Single, Num_Fich As Integer, Nb_Calcul As Integer, Num_Trc As Integer, num_data As Integer)  ', Num_Pix As Integer)
        Dim Conc_Mat, Conc_Trc As Double
        Dim Total_Error_Mat(10) As Double
        Dim Total_Error_Trc(10) As Double
        Dim Indice_Z As Integer
        Dim i As Integer
        Dim J, K As Integer
        Dim Tab3(3) As Single

        Dim Nb_Temp_Pivot As Integer
        Dim Pivot_Mat As Integer
        Dim Pivot_Trc As Integer
        Dim R As Byte
        Dim Num_File As Integer
        Dim Num_Proc As Integer
        Dim Num_Trc As Integer
        Dim num_data As Integer
        Dim Nb_Calcul As Integer
        Dim P As Integer
        Dim Str As String
        Dim Local_Txt_Fichier_PAR_Trc As String
        Dim offset_trc As Integer
        Dim Error_1 As Single
        Dim Somme_1 As Single
        Dim Somme_2 As Single
        Dim SR As StreamReader
        Dim posStar As Integer
        Dim PivotInMat As Boolean
        Dim PivotInTrc As Boolean


        Num_Proc = Parametres.Num_Proc
        Num_File = Parametres.Num_File
        Num_Trc = Parametres.Num_Trc
        num_data = Parametres.Num_Data
        Nb_Calcul = Parametres.Nb_Calcul

        For i = 0 To Num_Trc - 1
            offset_trc = offset_trc + Nb_Elements_Trc(i)
        Next

        For J = 0 To 10
            R = Tab_Pivot(Num_Trc, J)
            If R > 0 Then
                Nb_Temp_Pivot = Nb_Temp_Pivot + 1
            Else
                Exit For
            End If
        Next J

        If Indice_Pivot_trc(Num_Trc, 0) = -1 Then ' RECHERCHE DE L'INDICE DU PIVOT DANS LE TAB CONCENTRATION
            J = 0
            For K = 0 To Nb_Temp_Pivot - 1
                Pivot = Tab_Pivot(Num_Trc, K)
                J = 0
                PivotInMat = False
                Do
                    Indice_Z = Tab_Info_Mat.Z(J)
                    posStar = Strings.InStr(1, Tab_Info_Mat.Raie(J), "*", vbTextCompare)
                    If (Indice_Z = Pivot) And posStar = 0 Then
                        PivotInmat = True
                        Exit Do
                    End If
                    J = J + 1
                Loop While (J < Nb_Elements_Mat)

                If PivotInMat = True Then ' Le Z a été trouvé dans la liste MAT ?
                    Indice_Pivot_Mat(K) = J
                Else
                    Indice_Pivot_Mat(K) = -1
                End If

                i = 0
                PivotInTrc = False
                Do
                    Indice_Z = Tab_Info_Trc(Num_Trc).Z(i)
                    posStar = Strings.InStr(1, Tab_Info_Trc(Num_Trc).Raie(i), "*", vbTextCompare)
                    If (Indice_Z = Pivot) And posStar = 0 Then
                        PivotInTrc = True
                        Exit Do
                    End If
                    i = i + 1
                Loop While (i < Nb_Elements_Trc(Num_Trc))

                If PivotInTrc = True Then ' Le Z a été trouvé dans la liste TRC ?
                    Indice_Pivot_trc(Num_Trc, K) = i
                Else
                    Indice_Pivot_trc(Num_Trc, K) = -1
                End If

            Next K
        End If

        For K = 0 To Nb_Temp_Pivot - 1
            If Indice_Pivot_Mat(K) <> -1 And Indice_Pivot_trc(Num_Trc, K) <> -1 Then

                If Val_Mat_Total_Error(Num_Proc, Indice_Pivot_Mat(K)) = 0 Then '### error = 0 means element not present
                    Total_Error_Mat(K) = 999999
                Else
                    Total_Error_Mat(K) = Val_Mat_Total_Error(Num_Proc, Indice_Pivot_Mat(K))
                End If
                If Val_Trc_Total_Error(Num_Proc, offset_trc + Indice_Pivot_trc(Num_Trc, K)) = 0 Then
                    Total_Error_Trc(K) = 999999
                Else
                    Total_Error_Trc(K) = Val_Trc_Total_Error(Num_Proc, offset_trc + Indice_Pivot_trc(Num_Trc, K))
                End If
            Else
                Total_Error_Mat(K) = 9999999
                Total_Error_Trc(K) = 9999999
            End If


        Next K

        Pivot_Mat = Indice_Pivot_Mat(0)
        Pivot_Trc = Indice_Pivot_trc(Num_Trc, 0)
        Indice_Z = 0

        For K = 0 To Nb_Temp_Pivot - 2
            Error_1 = Val_Mat_Total_Error(Num_Proc, K)
            Somme_1 = Math.Sqrt(Total_Error_Mat(Indice_Z) ^ 2 + Total_Error_Trc(Indice_Z) ^ 2)
            Somme_2 = Math.Sqrt(Total_Error_Mat(K + 1) ^ 2 + Total_Error_Trc(K + 1) ^ 2)

            If Somme_2 < Somme_1 Then 'And Total_Error_Trc(K + 1) < Total_Error_Trc(Indice_Z) Then
                Pivot_Mat = Indice_Pivot_Mat(K + 1)
                Pivot_Trc = Indice_Pivot_trc(Num_Trc, K + 1)
                Indice_Z = K + 1
            End If
            'Conc_Mat = Spectrum_Mat(0)
        Next K


        If Pivot_Trc <> -1 Then ' TROUVER UN PIVOT
            Conc_Mat = Val_Mat_Conc(Num_Proc, Pivot_Mat)
            Conc_Trc = Val_Trc_Conc(Num_Proc, Pivot_Trc + offset_trc)
            Info_Experience_Trc(Num_Proc, Num_Trc).Selected_Pivot = Tab_Info_Trc(Num_Trc).Z(Pivot_Trc)

            If Conc_Mat > 0 And Conc_Trc > 0 Then
                Val_Charge_Trc(Num_Proc) = (1 * (Conc_Trc) / (Conc_Mat)) * Val_Charge_Trc_Init  'CALCUL DE LA CHARGE
                Info_Experience_Trc(Num_Proc, Num_Trc).New_charge = CStr(Strings.Replace(Math.Round(Val_Charge_Trc(Num_Proc), 8), ",", ".", , , vbTextCompare)) 'Val_Charge_Trc(Num_Proc)

            Else
                Val_Charge_Trc(Num_Proc) = Val_Charge_Trc_Init 'MET LA CHARGE A LA VALEUR INITIALE
            End If

            If MyChargeStd <> 0 Then
                Val_Charge_Trc(Num_Proc) = MyChargeStd
            End If
        Else
            Val_Charge_Trc(Num_Proc) = Val_Charge_Trc_Init 'MET LA CHARGE A LA VALEUR INITIALE
        End If

        ToolStripStatusLabel1.Text = "Write Q in Trace PAR file"
        Txt_Fichier_PAR_Trc = ""
        SR = File.OpenText(Chemin_GupixWin_Multi(Num_Proc) & "\gupixwin.par")

        Local_Txt_Fichier_PAR_Trc = ""

        For P = 0 To 11
            Str = SR.ReadLine()
            Local_Txt_Fichier_PAR_Trc = Local_Txt_Fichier_PAR_Trc + Str + vbCrLf
        Next P

        Local_Txt_Fichier_PAR_Trc = Local_Txt_Fichier_PAR_Trc + CStr(Strings.Replace(Val_Charge_Trc(Num_Proc), ",", ".", , , vbTextCompare)) + "        (Beam charge in units consistent with H-value)" + vbCrLf

        Str = SR.ReadLine()
        Do While SR.EndOfStream <> True
            Str = SR.ReadLine()
            Local_Txt_Fichier_PAR_Trc = Local_Txt_Fichier_PAR_Trc + Str + vbCrLf
        Loop

        SR.Close()

        Kill(Chemin_GupixWin_Multi(Num_Proc) & "\gupixwin.par")
        File.WriteAllText(Chemin_GupixWin_Multi(Num_Proc) & "\gupixwin.par", Local_Txt_Fichier_PAR_Trc & vbCrLf)

    End Sub


    Private Sub Calcul_Charge_Multi_Pivot(Mycharge As Single, Num_file As Integer, Nb_Calcul As Integer, Num_Trc As Integer, num_data As Integer)  ', Num_Pix As Integer)
        Dim Conc_Mat, Conc_Trc As Single
        Dim Area_Mat(10) As Double
        Dim Area_Trc(10) As Double
        Dim Indice_Z As Integer
        Dim i, J, K As Integer
        Dim Nb_Temp_Pivot As Integer
        Dim Pivot_Mat As Integer
        Dim Pivot_Trc As Integer
        Dim R As Byte

        ToolStripStatusLabel1.Text = "Calculate Trace charge "


        For J = 0 To 10
            R = Tab_Pivot(Num_Trc, J)

            If R > 0 Then
                Nb_Temp_Pivot = Nb_Temp_Pivot + 1
            Else
                Exit For
            End If

        Next J

        If Nb_Calcul = 1 And num_data = 0 Then ' RECHERCHE DE L'INDICE DU PIVOT DANS LE TAB CONCENTRATION

            J = 0

            For K = 0 To Nb_Temp_Pivot - 1
                Pivot = Tab_Pivot(Num_Trc, K)
                J = 0

                Do
                    Indice_Z = Tab_Info_Mat.Z(J)
                    If (Indice_Z = Pivot) Then Exit Do
                    J = J + 1
                Loop While (J < Nb_Elements_Mat)
                Indice_Pivot_Mat(K) = J
                i = 0
                Do
                    Indice_Z = Tab_Info_Trc(Num_Trc).Z(i)
                    If (Indice_Z = Pivot) Then
                        Exit Do
                    End If
                    i = i + 1
                Loop While (i < Nb_Elements_Trc(Num_Trc))
                Indice_Pivot_trc(Num_Trc, K) = i
            Next K
        End If


        For K = 0 To Nb_Temp_Pivot - 1
            Area_Mat(K) = Tab_Val_Mat(Num_file).Area(Indice_Pivot_Mat(K))
            Area_Trc(K) = Tab_Val_Trc(Num_Trc, Num_file).Area(Indice_Pivot_trc(Num_Trc, K))
        Next K

        Pivot_Mat = Indice_Pivot_Mat(0)
        Pivot_Trc = Indice_Pivot_trc(Num_Trc, 0)
        Indice_Z = 0

        For K = 0 To Nb_Temp_Pivot - 1
            If Area_Mat(K + 1) > Area_Mat(Indice_Z) And Area_Trc(K + 1) > Area_Trc(Indice_Z) Then
                Pivot_Mat = Indice_Pivot_Mat(K + 1)
                Pivot_Trc = Indice_Pivot_trc(Num_Trc, K + 1)
                Indice_Z = K + 1
            End If
            'Conc_Mat = Spectrum_Mat(0)

        Next K

        Conc_Mat = Tab_Val_Mat(Num_file - num_data).Conc(Pivot_Mat)
        Conc_Trc = Tab_Val_Trc(Num_Trc, Num_file - num_data).Conc(Pivot_Trc)
        Tab_Val_Mat(Num_file).Selected_Pivot(Num_Trc) = Tab_Info_Mat.Z(Pivot_Mat)

        If Conc_Mat > 0 And Conc_Trc > 0 Then
            Val_Charge_Trc(Num_file + Nb_Calcul) = (1 * (Conc_Trc) / (Conc_Mat)) * Val_Charge_Trc_Init 'CALCUL DE LA CHARGE
        Else
            Val_Charge_Trc(Num_file + Nb_Calcul) = Val_Charge_Trc_Init 'MET LA CHARGE A LA VALEUR INITIALE
        End If

        If MyChargeStd <> 0 Then
            Val_Charge_Trc(Num_file + Nb_Calcul) = MyChargeStd
        End If
    End Sub

    '***********************************************************************************************
    '**************************** LECTURE IN FILE ::: Z , CONC , AIRE , LOD ************************
    '***********************************************************************************************


    Sub Lit_Element_Multi_Thread(Parametres As Struct_Parametres_Thread)
        Dim i As Integer
        Dim Tab_Elements(100) As String
        Dim Element As Integer
        Dim fso2 As Object
        Dim Str As String
        Dim Pos_Text As Integer
        Dim PosEtoile As Integer
        Dim voie As Integer
        Dim Fact_Correct As Single
        Dim Num_Proc As Integer
        Dim Num_File As Integer
        Dim Num_Trc As Integer
        Dim offset_trc As Integer
        Dim SplitText() As String
        Dim lastOk As Integer
        Dim StatOut As Boolean
        Dim Taille_File As Integer
        Dim Conc_Arrondi As Integer
        Dim Fit_err As Single
        Dim Stat_err As Single
        Dim Total_error As Single
        Dim SR As StreamReader
        Dim Pos_e As Integer


        voie = Parametres.voie
        Num_Proc = Parametres.Num_Proc
        Num_File = Parametres.Num_File
        Num_Trc = Parametres.Num_Trc
        Fact_Correct = Parametres.Fact_Correct
        offset_trc = Parametres.Offset_Trc

        fso2 = CreateObject("Scripting.FileSystemObject")

        Do
            StatOut = File.Exists(Chemin_GupixWin_Multi(Num_Proc) & "\PIXCONC.OUT")
            Application.DoEvents() : System.Threading.Thread.Sleep(20)
            If StatOut = True Then Taille_File = FileLen(Chemin_GupixWin_Multi(Num_Proc) & "\PIXCONC.OUT")
            i = i + 1
            If StatOut = True And Taille_File > 0 Then Exit Do
        Loop While StatOut = False And i < 50
        Application.DoEvents() : System.Threading.Thread.Sleep(100)

        SR = File.OpenText((Chemin_GupixWin_Multi(Num_Proc) & "\PIXCONC.OUT"))
        ToolStripStatusLabel1.Text = "Read elemental concentration process n°" & CStr(Num_Proc)
        Element = 0
        SR.ReadLine()
        Str = SR.ReadLine()

        ' ToolStripStatusLabel1.Text = "Reading Current"
        Pos_Text = InStr(2, Str, "nA:", vbTextCompare)
        If voie = 1 Then Info_Experience_Mat(Num_Proc).Current = Mid(Str, Pos_Text + 3, 8) ' Tab_Val_Mat(Num_Proc + Num_File).Current = Mid(Str, Pos_Text + 3, 8)
        If voie = 2 Then Info_Experience_Trc(Num_Proc, Num_Trc).Current = Mid(Str, Pos_Text + 3, 8) 'Tab_Val_Trc(Num_Trc, Num_Proc + Num_File).Current = Mid(Str, Pos_Text + 3, 8)


        Str = SR.ReadLine()
        '  ToolStripStatusLabel1.Text = "Reading Resolution Détecteur"
        Pos_Text = InStr(1, Str, "DetRes(eV):", vbTextCompare)

        If voie = 1 Then Info_Experience_Mat(Num_Proc).Res = Mid(Str, Pos_Text + 11, 5) ' Tab_Val_Mat(Num_Proc + Num_File).Res = Mid(Str, Pos_Text + 11, 5)
        If voie = 2 Then Info_Experience_Trc(Num_Proc, Num_Trc).Res = Mid(Str, Pos_Text + 11, 5) ' Tab_Val_Trc(Num_Trc, Num_Proc + Num_File).Res = Mid(Str, Pos_Text + 11, 5)


        Str = SR.ReadLine()
        '   ToolStripStatusLabel1.Text = "Reading Chi**2"
        Pos_Text = InStr(1, Str, "Chi**2: ", vbTextCompare)
        If voie = 1 Then Info_Experience_Mat(Num_Proc).Chi2 = Mid(Str, Pos_Text + 7, 9) 'Tab_Val_Mat(Num_Proc + Num_File).Chi2 = Mid(Str, Pos_Text + 7, 9)
        If voie = 2 Then Info_Experience_Trc(Num_Proc, Num_Trc).Chi2 = Mid(Str, Pos_Text + 7, 9) ' Tab_Val_Trc(Num_Trc, Num_Proc + Num_File).Chi2 = Mid(Str, Pos_Text + 7, 9)


        SR.ReadLine()
        Str = SR.ReadToEnd
        SplitText = Split(Str, vbCrLf)
        lastOk = 0

        For k = 0 To UBound(SplitText) ' CLEAN SPLIT ARRAY
            If SplitText(k) <> "" Then
                Tab_Elements(lastOk) = Trim(SplitText(k))
                lastOk = lastOk + 1
            End If
        Next k

        Nb_Etoile = 0

        If voie = 1 Then

            For i = 0 To Nb_Elements_Mat - 1

                If Num_File = 0 Then
                    If i = 0 Then ReDim Tab_Info_Mat.Z(Nb_Elements_Mat - 1)
                    Tab_Info_Mat.Z(i) = Val(Mid(Tab_Elements(i), 1, 2)) 'Z de l'elem
                    Tab_Info_Mat.Elem(i) = Mid(Tab_Elements(i), 5, 2)   ' Nom de l'elem

                    If Mid(Tab_Elements(i), 8, 2) <> "K " Then ' Raie utilisé de l'elem
                        Tab_Info_Mat.Raie(i) = " #" & Mid(Tab_Elements(i), 7, 3)
                    Else
                        Tab_Info_Mat.Raie(i) = ""
                    End If

                    PosEtoile = InStr(1, Tab_Elements(i), "*", vbTextCompare)
                    If PosEtoile > 0 Then ' Trouve eleme non inclu dans le bouclage a 100%
                        Tab_Info_Mat.Raie(i) = " #" & Mid(Tab_Elements(i), 7, 3) & "*"
                        Nb_Etoile = Nb_Etoile + 1
                        Info_Mat_Raie(Num_Proc, i) = True
                    End If
                End If

                Val_Mat_Height(Num_Proc, i) = Val(Mid(Tab_Elements(i), 12, 8))
                Val_Mat_Area(Num_Proc, i) = Val(Mid(Tab_Elements(i), 20, 8))
                Val_Mat_Stat_Error(Num_Proc, i) = Double.Parse(Mid(Tab_Elements(i), 38, 8), USACulture) 'Val(Mid(Tab_Elements(i), 38, 8))
                Val_Mat_Fit_Error(Num_Proc, i) = Double.Parse(Mid(Tab_Elements(i), 48, 7), USACulture) 'Val(Mid(Tab_Elements(i), 48, 10))
                Stat_err = Val_Mat_Stat_Error(Num_Proc, i) ' Val(Mid(Tab_Elements(i), 48, 10))
                Fit_err = Val_Mat_Fit_Error(Num_Proc, i)

                ''''BEFORE 12/01/2022
                'Total_error = Math.Sqrt((Fit_err ^ 2) + (Stat_err ^ 2))
                Total_error = Fit_err

                Val_Mat_Total_Error(Num_Proc, i) = Total_error
                Try
                    Val_Mat_Conc(Num_Proc, i) = Val(Mid(Tab_Elements(i), 28, 10))
                Catch ex As Exception
                    Pos_e = -1
                    Pos_e = InStr(1, Mid(Tab_Elements(i), 28, 10), "e", vbTextCompare)
                    If Pos_e > 0 Then 'exponetiel in conc. error 
                        Val_Mat_Conc(Num_Proc, i) = 99999999
                    Else
                        Val_Mat_Conc(Num_Proc, i) = Integer.Parse((Mid(Tab_Elements(i), 28, 10)), USACulture) 'Double.Parse(Mid(Tab_Elements(i), 38, 8), USACulture)
                    End If

                End Try

                Val_Mat_LOD(Num_Proc, i) = Val(Mid(Tab_Elements(i), 56, 10))
                Val_Mat_Y_N_Q(Num_Proc, i) = Mid(Tab_Elements(i), 67, 1)
            Next i

        Else


            For i = 0 To Nb_Elements_Trc(Num_Trc) - 1

                If Num_File = 0 Then
                    Tab_Info_Trc(Num_Trc).Z(i) = Val(Mid(Tab_Elements(i), 1, 2)) 'Z de l'elem
                    Tab_Info_Trc(Num_Trc).Elem(i) = Mid(Tab_Elements(i), 5, 2)   ' Nom de l'elem

                    If Mid(Tab_Elements(i), 8, 2) <> "K " Then ' Raie utilisé de l'elem
                        Tab_Info_Trc(Num_Trc).Raie(i) = " #" & Mid(Tab_Elements(i), 7, 3)
                    Else
                        Tab_Info_Trc(Num_Trc).Raie(i) = ""
                    End If

                    PosEtoile = InStr(1, Tab_Elements(i), "*", vbTextCompare)
                    If PosEtoile > 0 And PosEtoile < 20 Then
                        Tab_Info_Trc(Num_Trc).Raie(i) = "*" & Tab_Info_Trc(Num_Trc).Raie(i) & "*"
                        Info_Trc_Raie(Num_Proc, offset_trc + i) = True
                    End If
                End If

                Val_Trc_Y_N_Q(Num_Proc, offset_trc + i) = Mid(Tab_Elements(i), 67, 1)
                Val_Trc_Height(Num_Proc, offset_trc + i) = Val(Mid(Tab_Elements(i), 12, 10))
                Val_Trc_Area(Num_Proc, offset_trc + i) = Val(Mid(Tab_Elements(i), 20, 8))
                Try
                    Conc_Arrondi = Val(Mid(Tab_Elements(i), 28, 10))
                Catch ex As Exception
                    Pos_e = -1
                    Pos_e = InStr(1, Mid(Tab_Elements(i), 28, 10), "e", vbTextCompare)
                    If Pos_e > 0 Then 'exponetiel in conc. error 
                        Conc_Arrondi = 99999999
                    Else ' il y a une virgule conc < 1ppm ?
                        Conc_Arrondi = Integer.Parse((Mid(Tab_Elements(i), 28, 10)), USACulture) 'Double.Parse(Mid(Tab_Elements(i), 38, 8), USACulture)
                    End If
                End Try

                If Val(Mid(Tab_Elements(i), 28, 10)) < 1 And Val(Mid(Tab_Elements(i), 28, 10)) > 0 Then Conc_Arrondi = 1

                Val_Trc_Stat_Error(Num_Proc, offset_trc + i) = Double.Parse(Mid(Tab_Elements(i), 38, 8), USACulture) 'Val(Mid(Tab_Elements(i), 38, 8))
                Val_Trc_Fit_Error(Num_Proc, offset_trc + i) = Double.Parse(Mid(Tab_Elements(i), 48, 7), USACulture) 'Val(Mid(Tab_Elements(i), 48, 7))

                Fit_err = Val_Trc_Fit_Error(Num_Proc, offset_trc + i)
                Stat_err = Val_Trc_Stat_Error(Num_Proc, offset_trc + i)

                ''''BEFORE 12/01/2022
                'Total_error = Math.Sqrt((Fit_err ^ 2) + (Stat_err ^ 2))
                Total_error = Fit_err
                Val_Trc_Total_Error(Num_Proc, i) = Total_error
                Val_Trc_Conc(Num_Proc, offset_trc + i) = Conc_Arrondi ' Val_Trc_Total_Error(Num_Proc, i) 'Conc_Arrondi
                Val_Trc_LOD(Num_Proc, offset_trc + i) = Val(Mid(Tab_Elements(i), 56, 10))
                If Val(Mid(Tab_Elements(i), 56, 10)) < 1 And Val(Mid(Tab_Elements(i), 56, 10)) > 0 Then Val_Trc_LOD(Num_Proc, offset_trc + i) = 1

            Next i
        End If

        SR.Close()

    End Sub

    Sub Lit_Element_Multi_Simple_Thread(Parametres As Struct_Parametres_Thread)
        Dim i As Integer
        Dim Tab_Elements(100) As String

        Dim Element As Integer

        Dim Str As String
        Dim New_Num_File As Integer
        Dim voie As Integer
        Dim Fact_Correct As Single
        Dim Num_Proc As Integer
        Dim Num_File As Integer
        Dim Num_Trc As Integer
        Dim offset_trc As Integer
        Dim SplitText() As String
        Dim LastOk As Integer
        Dim StatOut As Boolean
        Dim Fit_err As Single
        Dim Stat_err As Single
        Dim SR As StreamReader

        voie = Parametres.voie
        Num_Proc = Parametres.Num_Proc
        Num_File = Parametres.Num_File
        Num_Trc = Parametres.Num_Trc
        Fact_Correct = Parametres.Fact_Correct
        offset_trc = Parametres.Offset_Trc

        If Num_Proc >= Nb_Process1 Then
            New_Num_File = Num_Proc - Nb_Process1
        Else
            New_Num_File = Num_Proc
        End If

        Do
            StatOut = File.Exists(Chemin_GupixWin_Multi(Num_File) & "\PIXCONC.OUT")
            Application.DoEvents() : System.Threading.Thread.Sleep(20)
            i = i + 1
            If StatOut = True Then Exit Do
        Loop While StatOut = False And i < 50
        Application.DoEvents() : System.Threading.Thread.Sleep(50)


        SR = File.OpenText((Chemin_GupixWin_Multi(Num_Proc) & "\PIXCONC.OUT"))

        ToolStripStatusLabel1.Text = "Read Trace concentration process n°" & CStr(Num_Proc)
        Element = 0

        SR.ReadLine()
        SR.ReadLine()
        SR.ReadLine()
        SR.ReadLine()
        SR.ReadLine()

        'ToolStripStatusLabel1.Text = "Reading Data First" & CStr(Num_Proc)

        Str = SR.ReadToEnd()
        SplitText = Split(Str, vbCrLf)
        LastOk = 0

        For k = 0 To UBound(SplitText) ' CLEAN SPLIT ARRAY
            If SplitText(k) <> "" Then
                Tab_Elements(LastOk) = Trim(SplitText(k))
                LastOk = LastOk + 1
            End If
        Next k


        If voie = 1 Then
            For i = 0 To Nb_Elements_Mat - 1
                Val_Mat_Conc(Num_Proc, i) = Double.Parse(Mid(Tab_Elements(i), 28, 10), USACulture) 'Val(Mid(Tab_Elements(i), 28, 10))
                Val_Mat_Area(Num_Proc, i) = Val(Mid(Tab_Elements(i), 20, 8))
                Tab_Info_Mat.Z(i) = Val(Mid(Tab_Elements(i), 1, 2))
                Fit_err = Double.Parse(Mid(Tab_Elements(i), 48, 10), USACulture) 'Val(Mid(Tab_Elements(i), 48, 10))
                Stat_err = Double.Parse(Mid(Tab_Elements(i), 38, 8), USACulture) 'Val(Mid(Tab_Elements(i), 38, 8))
                ''''BEFORE 12/01/2022
                'Val_Mat_Total_Error(Num_Proc, i) = Math.Sqrt((Fit_err ^ 2) + (Stat_err ^ 2))
                Val_Mat_Total_Error(Num_Proc, i) = Fit_err
            Next i
        Else

            For i = 0 To Nb_Elements_Trc(Num_Trc) - 1
                Try
                    Val_Trc_Conc(Num_Proc, offset_trc + i) = Val(Mid(Tab_Elements(i), 28, 10))
                Catch ex As Exception
                    Val_Trc_Conc(Num_Proc, offset_trc + i) = 0
                End Try

                Val_Trc_Area(Num_Proc, offset_trc + i) = Val(Mid(Tab_Elements(i), 20, 8))
                Tab_Info_Trc(Num_Trc).Z(i) = CInt(Mid(Tab_Elements(i), 1, 2))

                Val_Trc_Stat_Error(Num_Proc, offset_trc + i) = Double.Parse(Mid(Tab_Elements(i), 38, 8), USACulture) 'Val(Mid(Tab_Elements(i), 38, 8))
                Val_Trc_Fit_Error(Num_Proc, offset_trc + i) = Double.Parse(Mid(Tab_Elements(i), 48, 7), USACulture) 'Val(Mid(Tab_Elements(i), 48, 7))

                Stat_err = Val_Trc_Stat_Error(Num_Proc, offset_trc + i)
                Fit_err = Val_Trc_Fit_Error(Num_Proc, offset_trc + i)
                ''''BEFORE 12/01/2022
                'Val_Trc_Total_Error(Num_Proc, i) = Math.Sqrt((Fit_err ^ 2) + (Stat_err ^ 2))
                Val_Trc_Total_Error(Num_Proc, offset_trc + i) = Fit_err


            Next i
        End If

        SR.Close()
    End Sub


    Private Sub Lit_Oxyde_Multi_Thread(Parametres As Struct_Parametres_Thread)
        Dim j, i As Integer
        Dim PosEtoile As Integer
        Dim Oxyde As Integer
        Dim tab_oxyde(100) As String
        Dim Str As String
        Dim Str1 As String
        Dim Str2 As String
        Dim Oxyde_Ok As Boolean
        Dim New_Num_File As Integer
        Dim N1 As Integer
        Dim PosEsp As Integer
        Dim Num_Proc As Integer
        Dim Num_File As Integer
        Dim Num_Trc As Integer
        Dim Num_Oxy_Local As Integer
        Dim Voie As Integer
        Dim StatOut As Boolean
        Dim Offset_Trc As Integer
        Dim Taille_File As Integer
        Dim Conc_Arrondi As Integer
        Dim Total_Error As Single
        Dim SR As StreamReader
        Dim SplitText() As String

        Voie = Parametres.voie
        Num_Proc = Parametres.Num_Proc
        Num_File = Parametres.Num_File
        Num_Trc = Parametres.Num_Trc
        Offset_Trc = Parametres.Offset_Trc

        If Voie = 1 Then
            ToolStripStatusLabel1.Text = " Read Matrice oxide conc. process n°" & CStr(Num_Proc)
        Else
            ToolStripStatusLabel1.Text = " Read Trace oxide conc. process n°" & CStr(Num_Proc)
        End If

        Do
            StatOut = File.Exists(Chemin_GupixWin_Multi(Num_Proc) & "\PIXOXIDE.OUT")
            Application.DoEvents() : System.Threading.Thread.Sleep(20)
            i = i + 1
            If StatOut = True Then Exit Do 'And Taille_File > 0 Then Exit Do
        Loop While StatOut = False And i < 200
        Application.DoEvents() : System.Threading.Thread.Sleep(50)

        SR = File.OpenText((Chemin_GupixWin_Multi(Num_Proc) & "\PIXOXIDE.OUT"))

        Nb_Ligne_Oxyde_X1 = 0
        Nb_Ligne_Oxyde_X2 = 0
        Oxyde = 0
        Num_Oxy_Local = 0

        If Voie = 1 Then            '#####################  VOIE 1
            For i = 0 To 7
                Str = SR.ReadLine()
                N1 = InStr(1, Str, "c[elmt]", vbTextCompare)
                If N1 > 2 Then Exit For 'SR.skipline
            Next i
            SR.ReadLine()

            Do While SR.EndOfStream <> True
                tab_oxyde(Num_Oxy_Local) = SR.ReadLine()
                Num_Oxy_Local = Num_Oxy_Local + 1
            Loop

            SR.Close()
        Else                            '##################### VOIE 2
            For i = 0 To 5
                Str = SR.ReadLine()
                N1 = InStr(1, Str, "c[elmt]", vbTextCompare)
                If N1 > 2 Then Exit For 'file_B2.skipline
            Next i
            SR.ReadLine()

            Do While SR.EndOfStream <> True
                tab_oxyde(Num_Oxy_Local) = SR.ReadLine()
                Num_Oxy_Local = Num_Oxy_Local + 1
            Loop

        End If

        SR.Close()

        j = 0
        On Error GoTo suite

        If Voie = 1 Then
            For i = 0 To Nb_Elements_Mat - 1 ' "-3" car on ignore la dernière ligne
                If Num_File = 0 Then
                    Info_Oxyde_Mat.Z(j) = Mid(tab_oxyde(i), 5, 3)
                    PosEsp = InStr(1, Mid(tab_oxyde(i), 12, 6), " ", vbTextCompare)
                    Info_Oxyde_Mat.nom(j) = Mid(tab_oxyde(i), 12, PosEsp - 1) '& Tab_Info_Mat.Raie(i)
                End If
                Val_Mat_Oxyde(Num_Proc, i) = CInt(Strings.Mid(tab_oxyde(i), 41, 10))

                j = j + 1
            Next i
        Else 'TRACE
            For i = 0 To Nb_Elements_Trc(Num_Trc) - 1 ' Nb_Elements_Trc(Num_Trc) - 1  '4 ligne en trop
                If Num_File = 0 Then
                    Tab_Info_Oxyde_Trc(Num_Trc).Z(j) = Mid(tab_oxyde(i), 5, 3)
                    PosEsp = InStr(1, Mid(tab_oxyde(i), 12, 6), " ", vbTextCompare)
                    Tab_Info_Oxyde_Trc(Num_Trc).nom(j) = Mid(tab_oxyde(i), 12, PosEsp - 1)
                    PosEtoile = InStr(1, tab_oxyde(i), "*", vbTextCompare)
                    If PosEtoile > 0 Then Tab_Info_Oxyde_Trc(Num_Trc).nom(j) = "*" & Tab_Info_Oxyde_Trc(Num_Trc).nom(j) & "*"
                End If
                Val_Trc_Oxyde(Num_Proc, j + Offset_Trc) = CInt(Mid(tab_oxyde(i), 41, 10))
                j = j + 1
            Next i
        End If

suite:


    End Sub

    Sub Recup_Cps_Charge_Filters_Mat_Thread(Parametres As Struct_Parametres_Thread)

        Dim Pos_Cps As Integer
        Dim Ligne_Lu As String
        Dim i As Integer
        Dim New_Num_File As Integer
        Dim Num_Trc As Integer
        Dim Voie As Integer
        Dim Num_File As Integer
        Dim Num_Proc As Integer
        Dim StatOut As Boolean
        Dim Taille_File As Integer
        Dim SR As StreamReader

        Voie = Parametres.voie
        Num_Proc = Parametres.Num_Proc
        Num_File = Parametres.Num_File
        Num_Trc = Parametres.Num_Trc
        Pos_Cps = 0

        If Num_File >= Nb_Process1 Then
            New_Num_File = Num_File - Nb_Process1
        Else
            New_Num_File = Num_File
        End If


        Do
            StatOut = File.Exists(Chemin_GupixWin_Multi(Num_Proc) & "\pixstats.out")
            Application.DoEvents() : System.Threading.Thread.Sleep(20)
            If StatOut = True Then Taille_File = FileLen(Chemin_GupixWin_Multi(Num_Proc) & "\pixstats.out")
            i = i + 1
            If StatOut = True And Taille_File > 0 Then Exit Do
        Loop While StatOut = False And i < 50

        Application.DoEvents() : System.Threading.Thread.Sleep(50)

        Try
            SR = File.OpenText((Chemin_GupixWin_Multi(Num_Proc) & "\pixstats.out"))
        Catch ex As Exception
            MsgBox("Error reading " & Chemin_GupixWin_Multi(Num_Proc) & "\pixstats.out", MsgBoxStyle.MsgBoxHelp, "Error")
            Exit Sub
        End Try

        ToolStripStatusLabel1.Text = "Read experimental info."

        For i = 0 To 9
            Ligne_Lu = SR.ReadLine()
        Next i

        Ligne_Lu = SR.ReadLine()
        Pos_Cps = InStr(1, Ligne_Lu, " rate(cps):", vbTextCompare)
        Info_Experience_Mat(Num_Proc).Count_Rate = Mid(Ligne_Lu, Pos_Cps + 11, 8)
        SR.ReadLine() ' SKIP a Line
        Ligne_Lu = SR.ReadLine()
        Pos_Cps = InStr(1, Ligne_Lu, " Charge: ", vbTextCompare)
        Info_Experience_Mat(Num_Proc).Charge = Mid(Ligne_Lu, Pos_Cps + 9, 8)

        If Num_Proc + Num_File = 0 Or First_Init = False Or Adjust_Filter_B = True Then
            First_Init = True 'Ne passera plus ici
            Do
                Ligne_Lu = SR.ReadLine()
                Pos_Cps = InStr(1, Ligne_Lu, " Looking for", vbTextCompare)
            Loop While Pos_Cps = 0
            If mnuOxydeOUI.Checked = True Then Nb_Oxyde_Mat = Nb_Elements_Mat

            Do
                Ligne_Lu = SR.ReadLine()
                Pos_Cps = InStr(1, Ligne_Lu, " Filters", vbTextCompare)
                If Pos_Cps = 0 Then Pos_Cps = InStr(1, Ligne_Lu, " Absorbers", vbTextCompare)
            Loop While Pos_Cps = 0
            Info_Experience_Mat(Num_Proc).Filters = SR.ReadLine()
            Info_Experience_Mat(0).Filters = Info_Experience_Mat(Num_Proc).Filters
        Else
            Info_Experience_Mat(Num_Proc).Filters = Info_Experience_Mat(0).Filters
        End If

        '  RECHERCHE JUSQU 'A LA FIN
        Do While Not SR.EndOfStream
            Ligne_Lu = SR.ReadLine()
            Pos_Cps = InStr(1, Ligne_Lu, " multiplied by", vbTextCompare)
            If Pos_Cps > 0 Then Exit Do
        Loop

        If Pos_Cps > 0 Then
            Info_Experience_Mat(Num_Proc).New_charge = Strings.Replace(CStr(Math.Round(Double.Parse(Mid(Ligne_Lu, Pos_Cps + 15, 11), USACulture), 8)), ",", ".", vbTextCompare) 'Mid(Ligne_Lu, Pos_Cps + 15, 11) 'CStr(Double.Parse(Mid(Ligne_Lu, Pos_Cps + 15, 11), USACulture)) ' Mid(Ligne_Lu, Pos_Cps + 15, 11)
        Else
            'Pas de multiply by donc . Charge = Valeur initial
            Info_Experience_Mat(Num_Proc).New_charge = 1
        End If

        SR.Close()

    End Sub



    Sub Insert_Matrix(Num_File As Integer, Num_Proc As Integer) '####################################### INSERT MATRICE #######################
        Dim i As Integer
        Dim Nb_El_Mat As Integer
        Dim Str As String

        Dim pos_space As Integer
        Dim SplitText() As String
        Dim LastOK As Integer
        Dim k As Integer
        Dim Elem_Type As String
        Dim Text_Mat As String
        Dim Nb_elem_Inv As Integer
        Dim SR As StreamReader

        ToolStripStatusLabel1.Text = "Insert Matrix in trace PAR file"
        Text_Mat = ""

        Try
            SR = File.OpenText((Chemin_GupixWin_Multi(Num_Proc) & "\PixMTX.out"))
        Catch ex As Exception
            MsgBox("Error reading " & Chemin_GupixWin_Multi(Num_Proc) & "\PixMTX.out", MsgBoxStyle.MsgBoxHelp, "Error")
            Exit Sub
        End Try
        Str = SR.ReadLine()
        Nb_El_Mat = 0
        i = 0
        Nb_Elements_Mtx_inv = 0

        Do While SR.EndOfStream <> True

            Str = SR.ReadLine

            pos_space = InStr(1, Str, " ", vbTextCompare)
            SplitText = Split(Str, " ")
            LastOK = 0

            For k = 0 To UBound(SplitText) ' CLEAN SPLIT ARRAY
                If SplitText(k) <> "" Then
                    SplitText(LastOK) = Trim(SplitText(k))
                    LastOK = LastOK + 1
                End If
            Next k
            LastOK = LastOK - 1
            ReDim Preserve SplitText(LastOK)

            If Num_Proc = 0 Then 'LECTURE/ECRITURE ENTETE
                Tab_Info_Mat.Inv(i) = SplitText(0)
            End If


            For j = 0 To Nb_Elements_Mat - 1
                If SplitText(0) < 11 Then
                    Tab_Entete_Inv(Nb_elem_Inv) = SplitText(0)
                    Conc_Invisible = Double.Parse(SplitText(1), USACulture) * 1
                    Val_Inv_Mtx(Num_Proc, Nb_elem_Inv) = SplitText(1) 'Conc_Invisible
                    Nb_El_Mat = Nb_El_Mat + 1
                    Nb_elem_Inv = Nb_elem_Inv + 1
                    Nb_Elements_Mtx_inv = Nb_elem_Inv
                    Text_Mat = Text_Mat + Str + vbCrLf
                    Exit For
                End If

                If SplitText(0) = CStr(Tab_Info_Mat.Z(j)) And SplitText(1) <> "0.0000" Then
                    Val_Mat_Mtx(Num_Proc, j) = SplitText(1)
                    Nb_El_Mat = Nb_El_Mat + 1
                    Text_Mat = Text_Mat + Str + vbCrLf
                    Exit For
                End If

            Next
            i = i + 1

        Loop

        If mnuOxydeOUI.Checked = True Then
            Elem_Type = " 1"
        Else
            Elem_Type = " 0"
        End If


        Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + CStr(Nb_El_Mat) + Elem_Type + vbTab + "(Number of matrix elements for thick/intermediate/layered target followed" + vbCrLf _
            + vbTab + vbTab + "by weight type.  Number of elements shows -1 if n/a; weight type 0 for" + vbCrLf _
            + vbTab + vbTab + "weight, 1 for oxide.)" + vbCrLf _
        + vbTab + "(1 line per matrix element: each line has matrix element atomic" + vbCrLf _
        + vbTab + "number then fractional concentration, then layer number or NL if" + vbCrLf _
        + vbTab + "not layered: if 1st entry was -1, there are no more lines.)" + vbCrLf
        Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + Text_Mat '+ vbCrLf ' + " NL " + Mid(tab_Info_mat.Raie(i), 1, 2) + vbCrLf
        System.Threading.Thread.Sleep(10)
        SR.Close()

    End Sub

    Sub Insert_Matrix_gamma(Num_File As Integer, Num_Proc As Integer) '####################################### INSERT MATRICE #######################
        Dim i As Integer
        Dim Nb_El_Mat As Integer
        Dim MyStr As String
        Dim pos_space As Integer
        Dim SplitText() As String
        Dim LastOK As Integer
        Dim k As Integer
        Dim Elem_Type As String
        Dim Text_Mat As String
        Dim Nb_elem_Inv_local As Integer
        Dim SR As StreamReader
        Dim conc_pixe_tmp As Double
        Dim conc_pixe_oxide_tmp As Double
        Dim num_gamma As Integer
        Dim tmp_struct(50) As MyMat
        Dim last_sum As Double
        Dim AStr As String
        Dim normalize_factor As Double
        Dim Sum_pixe_oxide_correction As Double 'Fraction d'oxygene a enlever de l'O total
        Dim nb_el_only_gamma As Integer
        Dim el_only_gamma As Boolean
        Dim pos1 As Integer
        Dim indx_G As Integer


        ToolStripStatusLabel1.Text = "Insert Matrix in trace PAR file"
        Text_Mat = ""

        Try
            SR = File.OpenText((Chemin_GupixWin_Multi(Num_Proc) & "\PixMTX.out"))
        Catch ex As Exception
            MsgBox("Error reading " & Chemin_GupixWin_Multi(Num_Proc) & "\PixMTX.out", MsgBoxStyle.MsgBoxHelp, "Error")
            Exit Sub
        End Try

        MyStr = SR.ReadLine()
        Nb_El_Mat = 0
        nb_el_only_gamma = 0
        i = 0
        Nb_Elements_Mtx_inv = 0
        num_gamma = 0
        Nb_elem_Inv_local = 0
        Dim tmp_Nb_Elements_Mat = CInt(Strings.Left(MyStr, 2))
        ReDim tmp_struct(tmp_Nb_Elements_Mat)

        Do While SR.EndOfStream <> True

            MyStr = SR.ReadLine

            pos_space = InStr(1, MyStr, " ", vbTextCompare)
            SplitText = Split(MyStr, " ")
            LastOK = 0

            For k = 0 To UBound(SplitText) ' CLEAN SPLIT ARRAY
                If SplitText(k) <> "" Then
                    SplitText(LastOK) = Trim(SplitText(k))
                    LastOK = LastOK + 1
                End If
            Next k
            LastOK = LastOK - 1
            ReDim Preserve SplitText(LastOK)
            tmp_struct(i).Z = SplitText(0)
            tmp_struct(i).conc = Double.Parse(SplitText(1), USACulture)
            tmp_struct(i).layer = SplitText(2)
            tmp_struct(i).name = SplitText(3)

            If Num_Proc = 0 Then 'LECTURE/ECRITURE ENTETE
                Tab_Info_Mat.Inv(i) = SplitText(0)
            End If

            '################################################### LECTURE DATA PIXMTX.OUT

            If SplitText(0) = 8 Then 'Ex : Oxygene Z=8
                If Num_Proc = 0 Then Tab_Entete_Inv(Nb_elem_Inv_local) = SplitText(0)
                conc_pixe_oxide_tmp = conc_pixe_oxide_tmp + Double.Parse(SplitText(1), USACulture) * 1
                Nb_El_Mat += 1
            End If

            indx_G = -1
            indx_G = Array.IndexOf(info_gamma_z, SplitText(0))

            If SplitText(0) < 11 And SplitText(0) <> 8 Then 'Ex : OTHER INV ELEMENT < 11 but not 8
                If Num_Proc = 0 Then Tab_Entete_Inv(Nb_elem_Inv_local) = SplitText(0)
                conc_pixe_tmp = conc_pixe_tmp + Double.Parse(SplitText(1), USACulture) * 1 'Add to conc PIXE total
                Nb_El_Mat += 1

            ElseIf indx_G <> -1 Then 'Elem. Present en PIXE ET Gamma Ex: Z= 11 par Gamma, soustrait la fraction oxyde a O lut dans PIXMTX.out (Sum_pixe_oxide_correction)

                If CInt(gamma_conc_init(Num_File, indx_G)) > 0 Then
                    Sum_pixe_oxide_correction += (Double.Parse(SplitText(1), USACulture)) * tab_rapport_oxide_gamma(indx_G)

                    Try
                        If info_gamma_z(num_gamma + 1) <> "" Then num_gamma += 1
                    Catch ex As Exception

                    End Try

                    Nb_El_Mat += 1

                ElseIf CInt(gamma_conc_init(Num_File, indx_G)) = 0 Then 'O en valeur Gamma on prend la valeur PIXE 

                    conc_pixe_tmp = conc_pixe_tmp + Double.Parse(SplitText(1), USACulture)
                    Nb_El_Mat += 1
                End If

            ElseIf SplitText(0) <> info_gamma_z(num_gamma) And SplitText(0) <> "8" Then 'Cas classic
                conc_pixe_tmp += Double.Parse(SplitText(1), USACulture)
                Nb_El_Mat += 1
            End If

            i = i + 1
        Loop 'END of FILE

        ReDim Preserve tmp_struct(i)
        Dim sum_t_gamma As Double
        Dim sum_t_pixe As Double


        num_gamma = 0
        If nb_gamma > 0 Then
            sum_t_gamma = (sum_gamma_conc(Num_File) + sum_gamma_oxide(Num_File)) / 1000000
            sum_t_pixe = conc_pixe_tmp + (conc_pixe_oxide_tmp - Sum_pixe_oxide_correction)
            'normalize_factor = (1 - ((sum_gamma_conc(Num_File) + sum_gamma_oxide(Num_File)) / 1000000)) / (conc_pixe_tmp + (conc_pixe_oxide_tmp - Sum_pixe_oxide_correction))
            normalize_factor = (1 - sum_t_gamma) / sum_t_pixe
        Else
            normalize_factor = 1
        End If

        For j = 0 To UBound(tmp_struct) - 1

            If tmp_struct(j).Z <> info_gamma_z(num_gamma) Then 'PAS un elemnt fourni par Gamma

                Select Case CInt(tmp_struct(j).Z) ' Pas dans Gamma
                    Case 8
                        tmp_struct(j).conc = Math.Round((tmp_struct(j).conc * normalize_factor) + (sum_gamma_oxide(Num_File) / 1000000), 4, MidpointRounding.AwayFromZero) 'Gamma
                        Val_Inv_Mtx(Num_Proc, Nb_elem_Inv_local) = CStr(tmp_struct(j).conc)
                        Nb_elem_Inv_local = Nb_elem_Inv_local + 1
                        Nb_Elements_Mtx_inv = Nb_elem_Inv_local
                    Case 1 To 7
                        tmp_struct(j).conc = Math.Round(tmp_struct(j).conc * normalize_factor, 4, MidpointRounding.AwayFromZero) 'Normalize Autre que Gamma
                        Val_Inv_Mtx(Num_Proc, Nb_elem_Inv_local) = CStr(tmp_struct(j).conc)
                        Nb_elem_Inv_local = Nb_elem_Inv_local + 1
                        Nb_Elements_Mtx_inv = Nb_elem_Inv_local
                        last_sum += tmp_struct(j).conc
                    Case Else
                        tmp_struct(j).conc = Math.Round(tmp_struct(j).conc * normalize_factor, 4, MidpointRounding.AwayFromZero) 'Normalize Autre que Gamma
                End Select

                last_sum = last_sum + tmp_struct(j).conc
            Else '############# GAMMA et MTX

                Select Case CInt(tmp_struct(j).Z)
                    Case 8
                        tmp_struct(j).conc = Math.Round((tmp_struct(j).conc * normalize_factor) + sum_gamma_oxide(Num_File) / 1000000, 4, MidpointRounding.AwayFromZero) 'Gamma
                        Val_Inv_Mtx(Num_Proc, Nb_elem_Inv_local) = CStr(tmp_struct(j).conc)
                        Nb_elem_Inv_local = Nb_elem_Inv_local + 1
                        Nb_Elements_Mtx_inv = Nb_elem_Inv_local
                        last_sum += tmp_struct(j).conc
                    Case 1 To 7
                        tmp_struct(j).conc = Math.Round((gamma_conc(Num_File, num_gamma) / 1000000), 4, MidpointRounding.AwayFromZero) 'Gamma
                        Val_Inv_Mtx(Num_Proc, Nb_elem_Inv_local) = CStr(tmp_struct(j).conc)
                        Nb_elem_Inv_local = Nb_elem_Inv_local + 1
                        Nb_Elements_Mtx_inv = Nb_elem_Inv_local
                        Nb_El_Mat += 1

                    Case Else
                        tmp_struct(j).conc = Math.Round((gamma_conc(Num_File, num_gamma) / 1000000), 4, MidpointRounding.AwayFromZero) 'Gamma

                End Select
                last_sum += tmp_struct(j).conc
            End If
        Next

        'Construction Fichier PixMTX.out

        last_sum = 0

        For i = 0 To nb_gamma - 1
            el_only_gamma = True
            For j = 0 To UBound(tmp_struct) - 1

                If info_gamma_z(i) = tmp_struct(j).Z Then
                    el_only_gamma = False
                    ' If info_gamma(num_gamma + 1, 0) <> "" Then num_gamma += 1
                End If
            Next j

            If el_only_gamma = True Then

                If Num_Proc = 0 Then Tab_Entete_Inv(Nb_elem_Inv_local) = info_gamma_z(i) 'Ex Lithium Z = 3

                nb_el_only_gamma += 1
                If gamma_conc(Num_File, i) > 0 Then
                    Nb_El_Mat += 1
                    AStr = Strings.Format(Math.Round(gamma_conc(Num_File, i) / 1000000, 4), "0.0000")
                    last_sum += Math.Round(gamma_conc(Num_File, i) / 1000000, 4)
                    pos1 = Strings.InStr(info_gamma_name(i), "O", vbTextCompare) 'Recherche l'element O dans le nom pour récuperer nom elementaire
                    Text_Mat = Text_Mat & info_gamma_z(i) & " " & Strings.Replace(AStr, ",", ".") & "  NL" & "  "

                    Select Case info_gamma_z(i)

                        Case 1
                            Text_Mat = Text_Mat & "H" & vbCrLf
                        Case 2
                            Text_Mat = Text_Mat & "He" & vbCrLf
                        Case 3
                            Text_Mat = Text_Mat & "Li" & vbCrLf
                        Case 4
                            Text_Mat = Text_Mat & "Be" & vbCrLf
                        Case 5
                            Text_Mat = Text_Mat & "B" & vbCrLf
                        Case 6
                            Text_Mat = Text_Mat & "C" & vbCrLf
                        Case 7
                            Text_Mat = Text_Mat & "N" & vbCrLf
                        Case 9
                            Text_Mat = Text_Mat & "F" & vbCrLf

                        Case Else
                            If pos1 > 0 Then
                                Text_Mat = Text_Mat & Strings.Left(info_gamma_name(i), pos1 - 1) & vbCrLf
                            Else
                                Text_Mat = Text_Mat & Strings.Left(info_gamma_name(i), 2) & vbCrLf
                            End If

                    End Select
                    el_only_gamma = True

                End If
            End If

        Next i

        For j = 0 To UBound(tmp_struct) - 1
            last_sum += tmp_struct(j).conc
            If j = UBound(tmp_struct) - 1 Then
                tmp_struct(j).conc = (1 - last_sum) + tmp_struct(j).conc
            End If
            AStr = Strings.Format(tmp_struct(j).conc, "0.0000")
            Text_Mat = Text_Mat & tmp_struct(j).Z & " " & Strings.Replace(AStr, ",", ".") & "  " & tmp_struct(j).layer & "  " & tmp_struct(j).name & vbCrLf
        Next

        Elem_Type = " 0"

        Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + CStr(Nb_El_Mat) + Elem_Type + vbTab + "(Number of matrix elements for thick/intermediate/layered target followed" + vbCrLf _
            + vbTab + vbTab + "by weight type.  Number of elements shows -1 if n/a; weight type 0 for" + vbCrLf _
            + vbTab + vbTab + "weight, 1 for oxide.)" + vbCrLf _
        + vbTab + "(1 line per matrix element: each line has matrix element atomic" + vbCrLf _
        + vbTab + "number then fractional concentration, then layer number or NL if" + vbCrLf _
        + vbTab + "not layered: if 1st entry was -1, there are no more lines.)" + vbCrLf
        Txt_Fichier_PAR_Trc = Txt_Fichier_PAR_Trc + Text_Mat '+ vbCrLf ' + " NL " + Mid(tab_Info_mat.Raie(i), 1, 2) + vbCrLf
        System.Threading.Thread.Sleep(10)
        SR.Close()

    End Sub



    Public Function Lit_Depth(voie As Object, Num_File As Integer) As Boolean
        Dim i As Integer
        Dim fso2
        Dim fil2
        Dim file_B2
        Dim Str As String
        Dim Pos_depth As Integer
        If voie = 1 Then MyDepth = False
        fso2 = CreateObject("Scripting.FileSystemObject")
        fil2 = fso2.GetFile(Chemin_GupixWin & "\Pixstats.OUT")
        file_B2 = fil2.OpenAsTextStream(1)
        ToolStripStatusLabel1.Text = "Import penetration Depth"
        file_B2 = fil2.OpenAsTextStream(1)

        Do
            Str = file_B2.readline
            Pos_depth = InStr(1, Str, "Target depth", vbTextCompare)
        Loop While file_B2.AtEndOfStream <> True And Pos_depth = 0

        If Pos_depth = 0 Then
            Lit_Depth = False
            Exit Function
        End If
        Lit_Depth = True
        MyDepth = True

        Str = file_B2.readline
        Str = file_B2.readline


        Do
            If voie = 1 Then
                Tab_Val_Mat(Num_File).Depth(i) = Math.Round(Val(Mid(Strings.Replace(Str, ".", ","), 14, 10)))
            Else
                Tab_Val_Trc(1, Num_File).Depth(i) = Math.Round(Val(Mid(Strings.Replace(Str, ".", ","), 14, 10))) '##################### INDEX A FAIRE 
            End If
            i = i + 1
            Str = file_B2.readline
            Pos_depth = InStr(1, Str, "--------", vbTextCompare)
        Loop While file_B2.AtEndOfStream <> True And Pos_depth = 0

    End Function

    Function return_Nom_Pivot(Txt_par As String) As String
        Dim Nom_local As String

        Select Case Txt_par

            Case "13"
                Nom_local = "-Al"
            Case "14"
                Nom_local = "-Si"
            Case "20"
                Nom_local = "-Ca"
            Case "25"
                Nom_local = "-Mn"
            Case "26"
                Nom_local = "-Fe"
            Case "29"
                Nom_local = "-Cu"
            Case "30"
                Nom_local = "-Zn"
            Case "47"
                Nom_local = "-Ag"
            Case "50"
                Nom_local = "-Sn"
            Case "79"
                Nom_local = "-Au"
            Case "82"
                Nom_local = "-Pb"
            Case ""
                Nom_local = ""
            Case Else
                Nom_local = "-" & Txt_par
        End Select
        Return Nom_local
    End Function


    '***********************************************************************************************
    '************************************* FONCTION EXCEL  *****************************************
    '***********************************************************************************************
    Public Function Excel_Open(Num_Fichier As Integer, Nom_File As String) As XLWorkbook
        Dim Comm1 As String
        Dim Comm2 As String
        Dim Comm3 As String
        Dim i As Integer
        Dim MyInd As Integer


        Dim Nom_Rapport As String
        Dim Offset_Conc100 As Integer
        Dim Nom_Pivot_Local As String
        Dim SplitText() As String
        Dim xlBook1 As XLWorkbook

        ToolStripStatusLabel1.Text = "Open Excel"
        Piv = ""
        Comm3 = ""
        Comm2 = ""
        'For J = 0 To Nb_Trc - 1
        If Nom_File = "" Then

            For i = 0 To 9
                Select Case i
                    Case 0
                        If Pivot_det0.Text <> "" Then
                            Nom_Pivot_Local = return_Nom_Pivot(Pivot_det0.Text)
                            Piv = Piv & "-" & Check_det0.Text & Nom_Pivot_Local
                            MyInd = MyInd + 1
                            Comm2 = Comm2 & Check_det0.Text
                        End If
                    Case 1
                        If Pivot_det1.Text <> "" Then
                            Nom_Pivot_Local = return_Nom_Pivot(Pivot_det1.Text)
                            Piv = Piv & "-" & Check_det1.Text & Nom_Pivot_Local
                            MyInd = MyInd + 1
                            Comm2 = Comm2 & Check_det1.Text
                        End If

                    Case 2
                        Nom_Pivot_Local = return_Nom_Pivot(Pivot_det2.Text)
                        If Nom_Pivot_Local <> "" Then
                            Piv = Piv & "-" & Check_det2.Text & Nom_Pivot_Local
                            MyInd = MyInd + 1
                            Comm2 = Comm2 & Check_det2.Text
                        End If
                    Case 3
                        Nom_Pivot_Local = return_Nom_Pivot(Pivot_det3.Text)
                        If Nom_Pivot_Local <> "" Then
                            Piv = Piv & "-" & Check_det3.Text & Nom_Pivot_Local
                            MyInd = MyInd + 1
                            Comm2 = Comm2 & Check_det3.Text
                        End If
                    Case 4
                        Nom_Pivot_Local = return_Nom_Pivot(Pivot_det4.Text)
                        If Nom_Pivot_Local <> "" Then
                            Piv = Piv & "-" & Check_det4.Text & Nom_Pivot_Local
                            MyInd = MyInd + 1
                            Comm2 = Comm2 & Check_det4.Text
                        End If
                    Case 5
                        Nom_Pivot_Local = return_Nom_Pivot(Pivot_det5.Text)
                        If Nom_Pivot_Local <> "" Then
                            Piv = Piv & "-" & Check_det5.Text & Nom_Pivot_Local
                            MyInd = MyInd + 1
                            Comm2 = Comm2 & Check_det5.Text
                        End If
                    Case 6
                        Nom_Pivot_Local = return_Nom_Pivot(Pivot_det6.Text)
                        If Nom_Pivot_Local <> "" Then
                            Piv = Piv & "-" & Check_det6.Text & Nom_Pivot_Local
                            MyInd = MyInd + 1
                            Comm2 = Comm2 & Check_det6.Text
                        End If
                    Case 7
                        Nom_Pivot_Local = return_Nom_Pivot(Pivot_det7.Text)
                        If Nom_Pivot_Local <> "" Then
                            Piv = Piv & "-" & Check_det7.Text & Nom_Pivot_Local
                            MyInd = MyInd + 1
                            Comm2 = Comm2 & Check_det7.Text
                        End If
                    Case 8
                        Nom_Pivot_Local = return_Nom_Pivot(Pivot_det8.Text)
                        If Nom_Pivot_Local <> "" Then
                            Piv = Piv & "-" & Check_det8.Text & Nom_Pivot_Local
                            MyInd = MyInd + 1
                            Comm2 = Comm2 & Check_det8.Text
                        End If

                    Case Else
                        Piv = Piv
                End Select
            Next i


            If hdf5_mode = True Then
                SplitText = Split(TxtBox_HDF5_File.Text, "_")
                ' ///processed_data
                Nom_Projet = "PRJ"

                Try
                    Nom_Projet = SplitText(2)
                Catch ex As Exception
                    Nom_Projet = "PRJ"
                End Try
            Else
                SplitText = Split(Fichier_Matrix(0), "_")

                Try
                    Nom_Projet = SplitText(3)
                Catch ex As Exception
                    Nom_Projet = "PRJ"
                End Try
            End If

            If Adjust_Filter_B = True Then
                Nom_Rapport = Nom_Projet & "_Filter"
            Else
                Nom_Rapport = Nom_Projet
            End If


            If OnlyTrace = False And Calcul_With_Trc = True Then
                Comm1 = "_MAT-" & CbDetMat.Text
                Comm2 = "_TRC" '&
            ElseIf OnlyTrace = True Then
                'Comm1 = ""
                Comm1 = "_TRC"
                Comm2 = "_" & "Comm2" & "_"
            Else
                Comm1 = "_MAT-" & CbDetMat.Text
                Comm2 = ""
            End If

            If Check_Trc_As_Oxy.Checked = True And Nom_Excel_Trx_O = "" Then
                Comm2 = Comm2 & "_O" '&
            End If

            If Nom_Excel_Trx_O <> "" Then Comm3 = Nom_Excel_Trx_O

            If Adjust_Filter_B = True Then
                TextXLS.Text = "Filter-Thickness-adjustment_" & Nom_Rapport & "_" & ComboBox_Type_F.Text & "_Z-var-" & TextF_Z.Text & ".xlsx" '& Comm1 & Piv & ".xls"
            Else
                TextXLS.Text = "TRAUPIXE-" & Nom_Rapport & Comm1 & Comm2 & Piv & Comm3 & ".xlsx" '& Comm1 & Piv & ".xls"
            End If
            Nom_File = TextXLS.Text

        End If

        'Chemin_Rapport = Chemin_Data & "\" & Nom_File
        Chemin_Rapport = Chemin_Processed_Data & Nom_File
        Offset_Excel = 0
        Offset_Conc100 = 0

        Try
            xlBook1 = New XLWorkbook(Chemin_Rapport)
        Catch ex As Exception
            Dim MRet = Excel_Save_Test(1)
            If MRet = 1 Then
                Excel_Open = Nothing
                Exit Function
            End If
            GoTo creer_new
        End Try
        xlSheet_ExpData = xlBook1.Worksheet("Exp. data")
        xlSheet_Conc = xlBook1.Worksheet("Elemental Conc.")
        If mnuOxydeOUI.Checked = True Then
            xlSheet_Oxyde = xlBook1.Worksheet("Oxide Conc.")
        End If
        xlSheet_LOD = xlBook1.Worksheet("LOD")
        xlSheet_Area = xlBook1.Worksheet("Area")
        If MyDepth = True Then xlSheet_Depth = xlBook1.Worksheet("Depth")
        xlSheet_Height = xlBook1.Worksheet("Peak Height")

        xlSheet_Info = xlBook1.Worksheet("Informations")
        xlSheet_Fit_Err = xlBook1.Worksheet("Fit-Error")
        xlSheet_Mtx = xlBook1.Worksheet("Matrix")

        Try
            xlSheet_Total_Error = xlBook1.Worksheet("Total Unc")
        Catch ex As Exception
            xlSheet_Total_Error = xlBook1.Worksheets.Add("Total Unc") 'After:=xlBook1.Worksheet(xlBook1.Worksheet.Count))
            xlSheet_Total_Error.Position = 2
        End Try

        Try
            xlSheet_S_Conc_Error_ppm = xlBook1.Worksheet("S_Conc. & Unc ppm")
        Catch ex As Exception
            xlSheet_S_Conc_Error_ppm = xlBook1.Worksheets.Add("S_Conc. & Unc ppm") 'After:=xlBook1.Worksheet(xlBook1.Worksheet.Count))
        End Try

        Try
            xlSheet_S_Conc_Error_100 = xlBook1.Worksheet("S_Conc. & Unc %")
        Catch ex As Exception
            xlSheet_S_Conc_Error_100 = xlBook1.Worksheets.Add("S_Conc. & Unc %") 'After:=xlBook1.Worksheet(xlBook1.Worksheet.Count))
        End Try

        Try
            xlSheet_S_Conc_ppm_RED = xlBook1.Worksheet("S_Conc. ppm (RED)")
        Catch ex As Exception
            xlSheet_S_Conc_ppm_RED = xlBook1.Worksheets.Add("S_Conc. ppm (RED)")

        End Try

        Try
            xlSheet_S_Conc_100_RED = xlBook1.Worksheet("S_Conc. % (RED)")
        Catch ex As Exception
            xlSheet_S_Conc_100_RED = xlBook1.Worksheets.Add("S_Conc. % (RED)")

        End Try

        Try
            xlSheet_S_Conc_100 = xlBook1.Worksheet("S_Conc. %")
        Catch ex As Exception
            xlSheet_S_Conc_100 = xlBook1.Worksheet("S_Conc. %")
        End Try

        Try
            xlSheet_S_Conc_ppm = xlBook1.Worksheet("S_Conc. ppm")
        Catch ex As Exception
            xlSheet_S_Conc_ppm = xlBook1.Worksheet("S_Conc. ppm")
        End Try

        Try
            xlSheet_Choix_S = xlBook1.Worksheet("S_Best Det.")
        Catch ex As Exception
            xlSheet_Choix_S = xlBook1.Worksheet("S_Conc. ppm")
        End Try


        Try
            xlSheet_Total_Error = xlBook1.Worksheet("Total Unc")
        Catch ex As Exception
            xlSheet_Total_Error = xlBook1.Worksheets.Add("Total Unc")

        End Try

        Dim Row
        Dim empty As Boolean
        Do
            Row = xlSheet_Conc.Cell(Offset_Excel + 3, 1)
            empty = Row.IsEmpty()
            'Temp1 = xlSheet_Conc.Cell(Offset_Excel + 3, 1).Value
            Offset_Excel = Offset_Excel + 1 'xlSheet_Conc.LastRowUsed().RowNumber() + 1
        Loop While empty = False
        Offset_Excel = Offset_Excel - 1

        If Num_Fichier = 0 Then
            ToolStripStatusLabel1.Text = "Data append to previous Excel file"
            LabelAppend.Visible = True
        End If

        GoTo OpenWorkbook_OK


creer_new:


        ToolStripStatusLabel1.Text = "Create new Excel file"
        LabelNew.Visible = True
        Offset_Excel = 0
        xlBook1 = New XLWorkbook()
        xlSheet_ExpData = xlBook1.Worksheets.Add("Exp. data")
        xlSheet_Conc = xlBook1.Worksheets.Add("Elemental Conc.")

        If mnuOxydeOUI.Checked = True Then
            xlSheet_Oxyde = xlBook1.Worksheets.Add("Oxide Conc.")
        End If

        xlSheet_LOD = xlBook1.Worksheets.Add("LOD")
        xlSheet_S_Conc_Error_ppm = xlBook1.Worksheets.Add("S_Conc. & Unc ppm")
        xlSheet_S_Conc_Error_100 = xlBook1.Worksheets.Add("S_Conc. & Unc %")
        xlSheet_S_Conc_ppm = xlBook1.Worksheets.Add("S_Conc. ppm")
        xlSheet_S_Conc_100 = xlBook1.Worksheets.Add("S_Conc. %")
        xlSheet_S_Conc_ppm_RED = xlBook1.Worksheets.Add("S_Conc. ppm (RED)")
        xlSheet_S_Conc_100_RED = xlBook1.Worksheets.Add("S_Conc. % (RED)")
        xlSheet_Total_Error = xlBook1.Worksheets.Add("Total Unc")
        xlSheet_Fit_Err = xlBook1.Worksheets.Add("Fit-Error")
        xlSheet_Height = xlBook1.Worksheets.Add("Peak Height")
        xlSheet_Area = xlBook1.Worksheets.Add("Area")
        xlSheet_Mtx = xlBook1.Worksheets.Add("Matrix")
        xlSheet_Info = xlBook1.Worksheets.Add("Informations")
        xlSheet_Choix_S = xlBook1.Worksheets.Add("S_Best Det.")


OpenWorkbook_OK:

        xlSheet_Conc.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        If mnuOxydeOUI.Checked = True Then
            xlSheet_Oxyde.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        End If
        xlSheet_LOD.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_Mtx.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_Fit_Err.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_Height.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_Area.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_ExpData.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_Info.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_Choix_S.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_Total_Error.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_S_Conc_ppm_RED.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_S_Conc_100_RED.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_S_Conc_ppm.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_S_Conc_100.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_S_Conc_Error_ppm.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        xlSheet_S_Conc_Error_100.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center

        Return xlBook1
    End Function




    Public Sub Excel_Close()
        xlBook = Nothing
        xlSheet_Area = Nothing
        xlSheet_Conc = Nothing
        xlSheet_LOD = Nothing
        xlSheet_ExpData = Nothing
        If mnuOxydeOUI.Checked = True Then xlSheet_Oxyde = Nothing
        If MyDepth = True Then xlSheet_Depth = Nothing
        'xlSheet_Stat_Err = Nothing
        xlSheet_Fit_Err = Nothing
        xlSheet_Height = Nothing
        xlSheet_Info = Nothing

    End Sub


    '#####################################   SAVE EXCEL   #################################################
    Public Function Excel_Save(Num_Fichier As Integer) As Integer
        Dim num_rapport, N As Integer
        Dim toto As Integer
        ' Dim T As Integer
        Dim Tst1 As String
        Tst1 = ""

        N = 1
        ToolStripStatusLabel1.Text = "Save Excel"
        On Error Resume Next

        '  Do
        num_rapport = Str(N)
        Err.Number = 1



        If Num_Fichier = 1 And Offset_Excel = 0 Then
            xlBook.SaveAs(Chemin_Rapport) ', xlExcel8) ', xlExcel7) 'genere Err.number
            ToolStripStatusLabel1.Text = "Excel Save as..."

        Else
            xlBook.Save()
            ToolStripStatusLabel1.Text = "Excel Save"


        End If

        Excel_Save = 0

    End Function

    Public Function Excel_Save_Test(Num_Fichier As Integer) As Integer
        Dim num_rapport, N As Integer
        Dim toto As Integer
        Dim Tst1 As String
        Tst1 = ""

        N = 1
        ToolStripStatusLabel1.Text = "Save Excel"
        On Error Resume Next
        If xlSheet_Conc = Nothing Then
            Excel_Save_Test = 0
            Exit Function
        End If
        '  Do
        num_rapport = Str(N)
        Err.Number = 1


        If Num_Fichier = 1 And Offset_Excel = 0 Then
            Tst1 = "Toto1"
        Else
            xlSheet_Conc.Cell(1, 1).Value = "Toto1"
            xlBook.Save()
            Tst1 = xlSheet_Conc.Cell(1, 1).Value

        End If

        If Tst1 = "Toto1" Then
            Excel_Save_Test = 0
            ToolStripStatusLabel1.Text = "OK Excel saved ..."
            xlSheet_Conc.Cell(1, 1).Value = ""
        Else
            Excel_Save_Test = 1
            ToolStripStatusLabel1.Text = "Error file Excel open"
        End If

    End Function


    Sub Entete_100()
        Dim i As Integer
        Dim t As Integer
        Dim k As Integer
        Dim Z As Integer
        Dim Z1 As Integer
        Dim Z_Done As Boolean
        Dim Num1 As Integer
        Dim Elem_Mat_as_Oxyde As Boolean
        Dim loc_All_Z_Trc(200) As Integer

        Elem_Mat_as_Oxyde = False

        ReDim Tab_Entete_100(50)
        ReDim Tab_Z_100(50)
        Dim indx1 As Integer
        Dim loc_Z_G As Integer


        For Z = 11 To 100
            Z_Done = False

            If Z = 26 Then
                Z = 26
            End If

            For i = 0 To Nb_Elements_Mat - 1
                Z1 = Tab_Info_Mat.Z(i)

                If Z1 = Z Then
                    If mnuOxydeOUI.Checked = True Then 'MODE OXYDE ?
                        Elem_Mat_as_Oxyde = True

                        For t = 0 To Nb_Trc - 1 'Recherche si Z est en TRACE donc en elementaire par default
                            For j = 0 To Nb_Elements_Trc(t) - 1 'Nb_Oxyde_Trc(T) - 1
                                If Tab_Info_Trc(t).Z(j) = Z Then 'Z présent en TRACE 
                                    Elem_Mat_as_Oxyde = False 'donc Conc. MAT et TRC en élementaire
                                    If Check_Trc_As_Oxy.Checked = True And Ck_AllAsOxy.Checked = False Then

                                        For k = 0 To UBound(Tab_Trc_as_Oxy) ' Z Demandé en OXYDE OU ELEM ?
                                            If Z = Tab_Trc_as_Oxy(k) Then Elem_Mat_as_Oxyde = True
                                        Next k

                                    ElseIf Ck_AllAsOxy.Checked = True Then
                                        Elem_Mat_as_Oxyde = True
                                    End If
                                End If

                            Next j
                        Next t
                    End If

                    If mnuOxydeOUI.Checked = True And Elem_Mat_as_Oxyde = True Then
                        Tab_Entete_100(Num1) = Info_Oxyde_Mat.nom(i)
                    Else
                        Tab_Entete_100(Num1) = Tab_Info_Mat.Elem(i)
                    End If
                    Tab_Z_100(Num1) = Z
                    Num1 = Num1 + 1
                    Z_Done = True
                    Exit For
                End If
            Next i

            If Z_Done = False Then
                For t = 0 To Nb_Trc - 1
                    If Z_Done = True Then Exit For
                    For i = 0 To Nb_Elements_Trc(t) - 1
                        Z1 = Tab_Info_Trc(t).Z(i)

                        If Z1 = Z And Z_Done = False Then

                            If Check_Trc_As_Oxy.Checked = True And Ck_AllAsOxy.Checked = False Then
                                For j = 0 To UBound(Tab_Trc_as_Oxy)

                                    If Tab_Trc_as_Oxy(j) = Z Then
                                        Tab_Entete_100(Num1) = Tab_Info_Oxyde_Trc(t).nom(i) 'OXYDE
                                        ' Tab_Z_100(Num1) = Z
                                        ' Num1 = Num1 + 1
                                        Exit For
                                    Else
                                        Tab_Entete_100(Num1) = Tab_Info_Trc(t).Elem(i) ' NOM ELEMENTAIRE
                                        'Exit For
                                    End If
                                Next j
                            ElseIf Ck_AllAsOxy.Checked = True Then
                                Tab_Entete_100(Num1) = Tab_Info_Oxyde_Trc(t).nom(i) 'OXYDE
                                'Exit For
                            Else

                                Tab_Entete_100(Num1) = Tab_Info_Trc(t).Elem(i)
                            End If

                            Tab_Z_100(Num1) = Z
                            Num1 = Num1 + 1
                            Z_Done = True
                            Exit For
                        End If
                    Next
                Next t
            End If
        Next Z

        Nb_Elem_Unique = Num1
        Nb_Elem_Unique_sans_external = Num1



        '############################## CHECK IF Z est ONLY GAMMA pour ajouter à l'entete excel
        'indx_mat = Array.IndexOf(Tab_Z_100, Z_gamma) 'Recheche Z_gamma dans Z Matrice
        'If indx_mat <> -1 Then
        '    el_only_gamma = False
        '    '    nb_gamma_and_pixe += 1
        'End If
        Dim nb As Integer
        nb = 0

        For i = 0 To nb_gamma - 1
            loc_Z_G = info_gamma_z(i)
            indx1 = Array.IndexOf(Tab_Z_100, loc_Z_G)

            If indx1 = -1 Then                      'Only Gamma
                Tab_Entete_100(Nb_Elem_Unique) = info_gamma_name(i) & "-" & ext_tech(i)
                Nb_Elem_Unique += 1
            Else                    'Replace original Name with Gamma Name + Tech
                Tab_Entete_100(indx1) = info_gamma_name(i) & "-" & ext_tech(i)
            End If

        Next i

        ReDim Preserve Tab_Entete_100(Nb_Elem_Unique - 1)

        ReDim Val_Conc_S_ppm(Nb_Process - 1, Nb_Elem_Unique - 1)
        ReDim Val_Conc_S_100(Nb_Process - 1, Nb_Elem_Unique - 1)
        ReDim Val_Conc_S_RED_ppm(Nb_Process - 1, Nb_Elem_Unique - 1)
        ReDim Val_Conc_S_RED100(Nb_Process - 1, Nb_Elem_Unique - 1)
        ReDim Val_Choix_S(Nb_Process - 1, Nb_Elem_Unique - 1)
        ReDim Val_YNQ_Final(Nb_Process - 1, Nb_Elem_Unique - 1)
        ReDim Val_Error_S(Nb_Process - 1, Nb_Elem_Unique - 1)
        ReDim Val_Conc_And_Error(Nb_Process - 1, Nb_Elem_Unique * 2 - 1)
        ReDim Val_Conc_And_Error100(Nb_Process - 1, Nb_Elem_Unique * 2 - 1)
    End Sub



    '########################################################################################################
    '#####################################   ECRITURE ENTETE EXCEL   #######################################
    '########################################################################################################


    ''' CLOSED XML
    Public Sub Ecrire_Entete_Excel(Num_Fichier As Integer)
        Dim i As Integer
        Dim t As Integer
        Dim H As Integer
        Dim Offset_Trc As Integer
        Dim Num_ligne As Integer
        ToolStripStatusLabel1.Text = "Write Matrix header in excel"
        Dim NomDet As String

        Num_ligne = 2

        NomDet = "BE"
        If Num_Fichier = 0 And Offset_Excel = 0 Then
            If Par_Mat.Text <> "" Then
                NomDet = CbDetMat.Text
                If Tab_Num_Trc(t) = 0 Then NomDet = "BE"

                With xlSheet_Conc.Cell(1, 3)
                    .value = CbDetMat.Text
                    '     .HorizontalAlignment = xlCenter
                End With

                With xlSheet_Total_Error.Cell(1, 3)
                    .value = CbDetMat.Text
                    '    .HorizontalAlignment = xlCenter
                End With
            End If

            For t = 0 To Nb_Trc - 1
                If t > 0 Then Offset_Trc = Offset_Trc + Nb_Elements_Trc(t - 1)

                With xlSheet_Conc.Cell(1, Nb_Elements_Mat + Offset_Trc + 3)
                    .value = NomDet_Trc(t)

                End With

                With xlSheet_Area.Cell(1, Nb_Elements_Mat + Offset_Trc + 3)
                    .value = NomDet_Trc(t)
                    '  .HorizontalAlignment = xlCenter
                End With

                With xlSheet_LOD.Cell(1, Nb_Elements_Mat + Offset_Trc + 3)
                    .value = NomDet_Trc(t)
                    ' .HorizontalAlignment = xlCenter
                End With

                With xlSheet_Height.Cell(1, Nb_Elements_Mat + Offset_Trc + 3)
                    .value = NomDet_Trc(t)
                    '.HorizontalAlignment = xlCenter
                End With

                With xlSheet_Total_Error.Cell(1, Nb_Elements_Mat + Offset_Trc + 3)
                    .value = NomDet_Trc(t)
                    '.HorizontalAlignment = xlCenter
                End With


                With xlSheet_Fit_Err.Cell(1, Nb_Elements_Mat + Offset_Trc + 3)
                    .value = NomDet_Trc(t)
                    '  .HorizontalAlignment = xlCenter
                End With

            Next t

            For i = 0 To Nb_Elements_Mat - 1  ' ########################################### MATRICE
                Application.DoEvents() ': Sleep 5
                With xlSheet_Mtx.Cell(Num_ligne, i + 3)
                    .value = Tab_Info_Mat.Z(i)
                End With
            Next i

            For i = 0 To Nb_Elements_Mtx_inv - 1
                With xlSheet_Mtx.Cell(Num_ligne, Nb_Elements_Mat + 3 + i)
                    .value = Tab_Entete_Inv(i)
                End With
            Next i

            For i = 0 To Nb_Elements_Mat - 1  ' ########################################### MATRICE
                Application.DoEvents() ': Sleep 5

                With xlSheet_Area.Cell(Num_ligne, i + 3)
                    .value = Tab_Info_Mat.Elem(i) + Tab_Info_Mat.Raie(i)
                End With

                With xlSheet_Conc.Cell(Num_ligne, i + 3)
                    .value = Tab_Info_Mat.Elem(i) + Tab_Info_Mat.Raie(i) ' Str(Tab_Info_Mat.Raie(i))
                End With

                With xlSheet_LOD.Cell(Num_ligne, i + 3)
                    .value = Tab_Info_Mat.Elem(i) + Tab_Info_Mat.Raie(i)
                End With

                If MyDepth = True Then
                    With xlSheet_Depth.Cell(Num_ligne, i + 3)
                        .value = Tab_Info_Mat.Elem(i) + Tab_Info_Mat.Raie(i)
                    End With
                End If

                With xlSheet_Height.Cell(Num_ligne, i + 3)
                    .value = Tab_Info_Mat.Elem(i) + Tab_Info_Mat.Raie(i)
                End With

                With xlSheet_Fit_Err.Cell(Num_ligne, i + 3)
                    .value = Tab_Info_Mat.Elem(i) + Tab_Info_Mat.Raie(i)
                End With


                With xlSheet_Total_Error.Cell(Num_ligne, i + 3)
                    .value = Tab_Info_Mat.Elem(i) + Tab_Info_Mat.Raie(i)
                End With
            Next i


            xlSheet_ExpData.Cell(1, 3).value = ("Mat * by ")
            xlSheet_ExpData.Cell(1, 4 + Nb_Trc).value = ("Av.(nA) " & CbDetMat.Text)
            xlSheet_ExpData.Cell(1, 5 + (Nb_Trc * 2)).value = ("Chi2 (" & CbDetMat.Text) & ")"
            xlSheet_ExpData.Cell(1, 6 + (Nb_Trc * 3)).value = ("Res. (" & CbDetMat.Text) & ")"
            xlSheet_ExpData.Cell(1, 7 + (Nb_Trc * 4)).value = ("Cnt/sec. " & CbDetMat.Text) & ")"
            xlSheet_ExpData.Cell(1, 8 + (Nb_Trc * 6)).value = ("Filters (" & CbDetMat.Text) & ")"
            xlSheet_ExpData.Cell(1, 9 + (Nb_Trc * 7)).value = ("Par file (" & CbDetMat.Text) & ")"


            For t = 0 To Nb_Trc - 1
                xlSheet_ExpData.Cell(1, 4 + t).value = ("Q (" & NomDet_Trc(t)) & ")"
                xlSheet_ExpData.Cell(1, 5 + t + Nb_Trc).value = ("Av.(nA) " & NomDet_Trc(t)) & ")"
                xlSheet_ExpData.Cell(1, 6 + t + (Nb_Trc * 2)).value = ("Chi2 (" & NomDet_Trc(t)) & ")"
                xlSheet_ExpData.Cell(1, 7 + t + (Nb_Trc * 3)).value = ("Res. (" & NomDet_Trc(t)) & ")"
                xlSheet_ExpData.Cell(1, 8 + t + (Nb_Trc * 4)).value = ("Cnt/sec. (" & NomDet_Trc(t)) & ")"
                xlSheet_ExpData.Cell(1, 8 + t + (Nb_Trc * 5)).value = ("Z Pivot (" & NomDet_Trc(t)) & ")"
                xlSheet_ExpData.Cell(1, 9 + t + (Nb_Trc * 6)).value = ("Filters (" & NomDet_Trc(t)) & ")"
                xlSheet_ExpData.Cell(1, 10 + t + (Nb_Trc * 7)).value = (".par file (" & NomDet_Trc(t)) & ")"
            Next t


            '************************************ECRITURE ELEMENT TRACE ******************************
            ToolStripStatusLabel1.Text = "Write Trace header in Excel"
            Offset_Trc = 0
            For t = 0 To Nb_Trc - 1

                If t > 0 Then Offset_Trc = Offset_Trc + Nb_Elements_Trc(t - 1)

                For i = 0 To Nb_Elements_Trc(t) - 1

                    Application.DoEvents() ': Sleep 5

                    With xlSheet_Area.Cell(Num_ligne, Nb_Elements_Mat + i + Offset_Trc + 3)
                        .value = Tab_Info_Trc(t).Elem(i) + Tab_Info_Trc(t).Raie(i)
                        .Style.Font.Bold = True
                        ' .HorizontalAlignment = xlCenter
                    End With

                    With xlSheet_Conc.Cell(Num_ligne, Nb_Elements_Mat + i + Offset_Trc + 3)
                        .value = Tab_Info_Trc(t).Elem(i) + Tab_Info_Trc(t).Raie(i)
                        .Style.Font.Bold = True
                        ' .HorizontalAlignment = xlCenter
                    End With

                    With xlSheet_LOD.Cell(Num_ligne, Nb_Elements_Mat + i + Offset_Trc + 3)
                        .value = Tab_Info_Trc(t).Elem(i) + Tab_Info_Trc(t).Raie(i)
                        .Style.Font.Bold = True
                        ' .HorizontalAlignment = xlCenter
                    End With

                    If MyDepth = True Then
                        With xlSheet_Depth.Cell(Num_ligne, Nb_Elements_Mat + i + Offset_Trc + 3)
                            .value = Tab_Info_Trc(t).Elem(i) + Tab_Info_Trc(t).Raie(i)
                            .Style.Font.Bold = True
                        End With
                    End If

                    With xlSheet_Fit_Err.Cell(Num_ligne, Nb_Elements_Mat + i + Offset_Trc + 3)
                        .value = Tab_Info_Trc(t).Elem(i) + Tab_Info_Trc(t).Raie(i)
                        .Style.Font.Bold = True
                    End With

                    With xlSheet_Total_Error.Cell(Num_ligne, Nb_Elements_Mat + i + Offset_Trc + 3)
                        .value = Tab_Info_Trc(t).Elem(i) + Tab_Info_Trc(t).Raie(i)
                        .Style.Font.Bold = True
                        ' .HorizontalAlignment = xlCenter
                    End With


                    With xlSheet_Height.Cell(Num_ligne, Nb_Elements_Mat + i + Offset_Trc + 3)
                        .value = Tab_Info_Trc(t).Elem(i) + Tab_Info_Trc(t).Raie(i)
                        .Style.Font.Bold = True
                        ' .HorizontalAlignment = xlCenter
                    End With
                    'Offset_Trc = Offset_Trc + 1
                Next i
            Next t

            If mnuOxydeOUI.Checked = True Then

                For i = 0 To Nb_Oxyde_Mat - 1
                    With xlSheet_Oxyde.Cell(Num_ligne, i + 3)
                        .value = Info_Oxyde_Mat.nom(i) & Tab_Info_Mat.Raie(i)
                        ' .HorizontalAlignment = xlCenter
                    End With


                Next i

                Offset_Trc = 0

                For t = 0 To Nb_Trc - 1
                    If t > 0 Then Offset_Trc = Offset_Trc + Nb_Elements_Trc(t - 1)

                    With xlSheet_Oxyde.Cell(1, Nb_Oxyde_Mat + Offset_Trc + 3)
                        .value = NomDet_Trc(t)
                    End With

                    For i = 0 To Nb_Elements_Trc(t) - 1

                        With xlSheet_Oxyde.Cell(Num_ligne, Nb_Oxyde_Mat + i + Offset_Trc + 3)
                            .value = Tab_Info_Oxyde_Trc(t).nom(i) & Tab_Info_Trc(t).Raie(i) ' Tab_Info_X2.Raie(i)
                            .Style.Font.Bold = True
                            ' .HorizontalAlignment = xlCenter
                        End With



                    Next i
                Next t
            End If


            For i = 0 To Nb_Elem_Unique - 1

                With xlSheet_S_Conc_100.Cell(Num_ligne, i + 3)
                    .value = Tab_Entete_100(i)
                    ' .HorizontalAlignment = xlCenter
                End With

                With xlSheet_S_Conc_ppm.Cell(Num_ligne, i + 3)
                    .value = Tab_Entete_100(i)
                    ' .HorizontalAlignment = xlCenter
                End With

                With xlSheet_Choix_S.Cell(Num_ligne, i + 3)
                    .value = Tab_Entete_100(i)
                    ' .HorizontalAlignment = xlCenter
                End With

                With xlSheet_S_Conc_ppm_RED.Cell(Num_ligne, i + 3)
                    .value = Tab_Entete_100(i)
                    ' .HorizontalAlignment = xlCenter
                End With


                With xlSheet_S_Conc_100_RED.Cell(Num_ligne, i + 3)
                    .value = Tab_Entete_100(i)
                    ' .HorizontalAlignment = xlCenter
                End With


            Next

            For i = 0 To Nb_Elem_Unique - 1

                With xlSheet_S_Conc_Error_ppm.Cell(Num_ligne, (i * 2) + 3)
                    .value = Tab_Entete_100(i)
                    .Style.Font.Bold = True
                End With

                With xlSheet_S_Conc_Error_ppm.Column((i * 2) + 3)
                    .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
                    .Width = 7
                End With

                With xlSheet_S_Conc_Error_ppm.Cell(Num_ligne, (i * 2) + 4)
                    .value = "Unc%"
                End With

                With xlSheet_S_Conc_Error_ppm.Column((i * 2) + 4)
                    .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
                    .Width = 7
                End With


                With xlSheet_S_Conc_Error_100.Cell(Num_ligne, (i * 2) + 3)
                    .value = Tab_Entete_100(i)
                    .Style.Font.Bold = True
                End With

                With xlSheet_S_Conc_Error_100.Column((i * 2) + 3)
                    .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
                    .Width = 7
                End With


                With xlSheet_S_Conc_Error_100.Cell(Num_ligne, (i * 2) + 4)
                    .value = "Unc%"
                End With

                With xlSheet_S_Conc_Error_100.Column((i * 2) + 4)
                    .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
                    .Width = 7
                End With


            Next

        End If

    End Sub

    'CLOSEDXML 

    Public Sub Excel_Write_Filename_Comment(Num_Fichier As Integer, Nb_Proc As Integer)
        'CLOSEDXML Dim Rng As Excel.Range
        Dim StartRow, StartCol
        Dim i As Integer
        Dim Pos_Ext As Integer
        Dim Filter_String As String
        ToolStripStatusLabel1.Text = "Write filename & comment in Excel"
        Application.DoEvents()
        Dim Tab_Ref_Woorksheet() As Object

        If mnuOxydeOUI.Checked = True Then
            ReDim Tab_Ref_Woorksheet(15)
        Else
            ReDim Tab_Ref_Woorksheet(14)
        End If

        Dim multiList_Str As New List(Of List(Of String))
        For i = 0 To Nb_Proc - 1
            multiList_Str.Add(New List(Of String))
        Next
        ' Adjust_Filter_B = True

        For i = 0 To Nb_Process - 1
            Pos_Ext = InStr(1, Fichier_Matrix(i + Num_Fichier), ".", vbTextCompare)
            Tab_Name_File(i) = Mid(Fichier_Matrix(i + Num_Fichier), 1, Pos_Ext - 1)
        Next

        If Adjust_Filter_B = True Then
            StartRow = (Num_Fichier * Nb_Proc) + 3 + Offset_Excel
        Else
            StartRow = Num_Fichier + 3 + Offset_Excel
        End If

        StartCol = 1
        Tab_Ref_Woorksheet(0) = xlSheet_Conc
        Tab_Ref_Woorksheet(1) = xlSheet_LOD
        Tab_Ref_Woorksheet(2) = xlSheet_Area
        Tab_Ref_Woorksheet(3) = xlSheet_Fit_Err
        Tab_Ref_Woorksheet(4) = xlSheet_Total_Error
        Tab_Ref_Woorksheet(5) = xlSheet_Height
        Tab_Ref_Woorksheet(6) = xlSheet_Mtx
        Tab_Ref_Woorksheet(7) = xlSheet_ExpData
        Tab_Ref_Woorksheet(8) = xlSheet_S_Conc_ppm_RED
        Tab_Ref_Woorksheet(9) = xlSheet_S_Conc_100
        Tab_Ref_Woorksheet(10) = xlSheet_S_Conc_ppm
        Tab_Ref_Woorksheet(11) = xlSheet_Choix_S
        Tab_Ref_Woorksheet(12) = xlSheet_S_Conc_100_RED
        Tab_Ref_Woorksheet(13) = xlSheet_S_Conc_Error_ppm
        Tab_Ref_Woorksheet(14) = xlSheet_S_Conc_Error_100

        If mnuOxydeOUI.Checked = True Then
            Tab_Ref_Woorksheet(15) = xlSheet_Oxyde
        End If


        For Each ws In Tab_Ref_Woorksheet
            ws.Cell(StartRow, 1).InsertData(Tab_Name_File)
            ws.Column(1).Width = 30
            ws.Cell(StartRow, 2).InsertData(Tab_Comment)
            ws.Column(2).Width = 20
        Next



    End Sub

    Sub Excel_Legend(Nb_file As Integer, legend As Boolean, MyErase As Boolean)
        Dim i As Integer
        Dim j As Integer
        Dim t As Integer
        Dim MyColor
        Dim Tab_Ref_Woorksheet(7) As Object
        Dim Tab_Ref_Woorksheet_Error(2) As Object

        Tab_Ref_Woorksheet(0) = xlSheet_S_Conc_ppm_RED
        Tab_Ref_Woorksheet(1) = xlSheet_S_Conc_100
        Tab_Ref_Woorksheet(2) = xlSheet_S_Conc_ppm
        Tab_Ref_Woorksheet(3) = xlSheet_Choix_S
        Tab_Ref_Woorksheet(4) = xlSheet_S_Conc_100_RED
        Tab_Ref_Woorksheet(5) = xlSheet_S_Conc_Error_ppm
        Tab_Ref_Woorksheet(6) = xlSheet_S_Conc_Error_100
        Tab_Ref_Woorksheet(7) = xlSheet_S_Conc_ppm_RED

        Tab_Ref_Woorksheet_Error(0) = xlSheet_S_Conc_Error_ppm
        Tab_Ref_Woorksheet_Error(1) = xlSheet_S_Conc_Error_100
        Tab_Ref_Woorksheet_Error(2) = xlSheet_Total_Error

        If MyErase = True Then
            '##################### EFFACE LEGENDE PRECEDENTE

            For j = 0 To 3

                For Each ws In Tab_Ref_Woorksheet
                    ws.Cell(5 + Offset_Excel + j, 6 + i).Value = ""
                    ws.Cell(5 + Offset_Excel + j, 6 + i).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 255)
                Next

                For Each ws In Tab_Ref_Woorksheet_Error
                    ws.Cell(5 + Offset_Excel + j, 6 + i).Value = ""
                    ws.Cell(5 + Offset_Excel + j, 6 + i).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 255)
                Next
            Next j

            For t = 0 To Nb_Trc - 1

                For Each ws In Tab_Ref_Woorksheet
                    ws.Cell(7 + Offset_Excel, 7 + t).Value = ""
                    ws.Cell(7 + Offset_Excel, 7 + t).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 255) 'XLColor.FromArgb(255, 255, 255) ' Mat sans fond
                Next
            Next
        End If
        i = 0

        If legend = True Then

            For Each ws In Tab_Ref_Woorksheet
                ws.Cell(Nb_file + 6 + Offset_Excel, 6).Value = "Color code for Detector"
                ws.Cell(Nb_file + 6 + Offset_Excel, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 255)
            Next

            For Each ws In Tab_Ref_Woorksheet
                ws.Cell(Nb_file + 7 + Offset_Excel, 6).Value = CbDetMat.Text
                ws.Cell(Nb_file + 7 + Offset_Excel, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 255) ' Mat sans fond
            Next


            For Each ws In Tab_Ref_Woorksheet_Error
                ws.Cell(Nb_file + 8 + Offset_Excel, 6).Value = "--> Matrix Error(%) = Fit_Error " & " --> Trace Error(%) = sqrt(Pivot_Error^2 + Fit_Error ^2) (Section 13C Gupix manuel)"

            Next

            For t = 0 To Nb_Trc - 1

                Select Case t
                    Case 0 '"X0"
                        MyColor = XLColor.FromArgb(230, 230, 230)
                    Case 1
                        MyColor = XLColor.FromArgb(200, 200, 200) ' GRIS CLAIRE
                    Case 2
                        MyColor = XLColor.FromArgb(170, 170, 170) ' ORANGE CLAIRE
                    Case 3
                        MyColor = XLColor.FromArgb(145, 145, 145)  ' JAUNE CLAIRE 
                    Case 4
                        MyColor = XLColor.FromArgb(200, 255, 200) ' VERT CLAIRE
                    Case 5
                        MyColor = XLColor.FromArgb(255, 200, 110) 'ORANGE FONCE
                    Case 6
                        MyColor = XLColor.FromArgb(145, 255, 145) ' JAUNE CLAIRE
                    Case 7
                        MyColor = XLColor.FromArgb(180, 180, 225) ' MAUVE FONCE CLAIRE
                    Case 8
                        MyColor = XLColor.FromArgb(255, 255, 200) ' JAUNE CLAIRE
                    Case Else
                        MyColor = XLColor.FromArgb(200, 200, 200) ' GRIS CLAIRE
                End Select

                For Each ws In Tab_Ref_Woorksheet
                    ws.Cell(Nb_file + 7 + Offset_Excel, 7 + t).Value = NomDet_Trc(t)
                    ws.Cell(Nb_file + 7 + Offset_Excel, 7 + t).Style.Fill.BackgroundColor = MyColor 'XLColor.FromArgb(255, 255, 255) ' Mat sans fond
                Next

                i = i + 1

            Next t
        End If

    End Sub

    ''########################################################################################################
    ''#####################################   ECRITURE DATA EXCEL   #########################################
    ''########################################################################################################

    Public Sub Excel_Data_ClosedXml(Num_Fichier As Integer, Nb_Proc As Integer, xlBook As IXLWorkbook)
        Dim Str_Rectif As String
        Dim Str_Fact, Str_Std As String
        Dim Tab_Color(100) As Integer
        Dim Tab_Font(100) As String
        Dim Tab_Italic(100) As String
        Dim StartRow, StartCol
        Dim SizeTab As Integer
        Dim SizeTab_Trc As Integer
        Dim i As Integer


        Dim multiList_Int As New List(Of List(Of Integer))
        For i = 0 To Nb_Proc - 1
            multiList_Int.Add(New List(Of Integer))
        Next i

        Dim multiList_Dbl As New List(Of List(Of Double))
        For i = 0 To Nb_Proc - 1
            multiList_Dbl.Add(New List(Of Double))
        Next
        Dim multiList As New List(Of List(Of String))
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        'multiList.Add(New List(Of String))
        Str_Fact = 1
        Str_Std = Str(MyChargeStd)
        Str_Rectif = "=" + Str_Fact + "/" + Str_Std

        ToolStripStatusLabel1.Text = "Write Matrix data in Excel"
        Application.DoEvents()

        StartRow = Num_Fichier + 3 + Offset_Excel
        StartCol = 3
        SizeTab = UBound(Val_Mat_Conc_1, 2)
        SizeTab_Trc = 0
        If Calcul_With_Trc = True Then SizeTab_Trc = UBound(Val_Trc_Conc_1, 2)
        ' CONC
        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList_Int(j).Add(Val_Mat_Conc_1(j, i))
            Next
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To SizeTab_Trc
                    multiList_Int(j).Add(Val_Trc_Conc_1(j, k))
                Next
            End If
        Next j
        xlSheet_Conc.Cell(StartRow, StartCol).InsertData(multiList_Int)
        multiList_Int.Clear()
        For i = 0 To Nb_Proc - 1
            multiList_Int.Add(New List(Of Integer))
        Next i

        'AREA
        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList_Int(j).Add(Val_Mat_Area_1(j, i))
            Next
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To SizeTab_Trc
                    multiList_Int(j).Add(Val_Trc_Area_1(j, k))
                Next
            End If
        Next j
        xlSheet_Area.Cell(StartRow, StartCol).InsertData(multiList_Int)
        multiList_Int.Clear()
        For i = 0 To Nb_Proc - 1
            multiList_Int.Add(New List(Of Integer))
        Next i


        'LOD
        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList_Int(j).Add(Val_Mat_LOD_1(j, i))
            Next
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To SizeTab_Trc
                    multiList_Int(j).Add(Val_Trc_LOD_1(j, k))
                Next
            End If
        Next j
        xlSheet_LOD.Cell(StartRow, StartCol).InsertData(multiList_Int)
        multiList_Int.Clear()
        For i = 0 To Nb_Proc - 1
            multiList_Int.Add(New List(Of Integer))
        Next i


        'xlSheet_Fit_Err
        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList_Dbl(j).Add(Val_Mat_Fit_Error_1(j, i))
            Next
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To SizeTab_Trc
                    multiList_Dbl(j).Add(Val_Trc_Fit_Error_1(j, k))
                Next
            End If
        Next j
        xlSheet_Fit_Err.Cell(StartRow, StartCol).InsertData(multiList_Dbl)
        multiList_Dbl.Clear()
        For i = 0 To Nb_Proc - 1
            multiList_Dbl.Add(New List(Of Double))
        Next

        'xlSheet_Total_Error
        For j = 0 To Nb_Proc - 1

            For i = 0 To SizeTab
                multiList_Dbl(j).Add(Val_Mat_Total_Error_1(j, i))
            Next
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To SizeTab_Trc
                    multiList_Dbl(j).Add(Val_Trc_WithPivot_Error_1(j, k))
                Next
            End If
        Next j
        xlSheet_Total_Error.Cell(StartRow, StartCol).InsertData(multiList_Dbl)
        multiList_Dbl.Clear()
        For i = 0 To Nb_Proc - 1
            multiList_Dbl.Add(New List(Of Double))
        Next

        'xlSheet_Height
        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList(j).Add(Val_Mat_Height_1(j, i))
            Next
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To SizeTab_Trc
                    multiList(j).Add(Val_Trc_Height_1(j, k))
                Next
            End If
        Next j
        xlSheet_Height.Cell(StartRow, StartCol).InsertData(multiList)

        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        'Exp Data_Charge
        Try
            For j = 0 To Nb_Proc - 1
                multiList(j).Add(Info_Experience_Mat(j).New_charge)
                If Calcul_With_Trc = True Then ' WITH TRACE
                    For k = 0 To Nb_Trc - 1
                        multiList(j).Add(Info_Experience_Trc(j, k).New_charge)
                    Next
                End If
            Next j
            xlSheet_ExpData.Cell(StartRow, StartCol).InsertData(multiList)
            multiList.Clear()

            For i = 0 To Nb_Proc - 1
                multiList.Add(New List(Of String))
            Next
        Catch ex As Exception

        End Try

        'Exp Data_Current
        For j = 0 To Nb_Proc - 1
            multiList(j).Add(Info_Experience_Mat(j).Current)
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To Nb_Trc - 1
                    multiList(j).Add(Info_Experience_Trc(j, k).Current)
                Next
            End If
        Next j
        xlSheet_ExpData.Cell(StartRow, StartCol + Nb_Trc + 1).InsertData(multiList)
        multiList.Clear()

        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        'Exp Data_Chi2
        For j = 0 To Nb_Proc - 1
            multiList(j).Add(Info_Experience_Mat(j).Chi2)
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To Nb_Trc - 1
                    multiList(j).Add(Info_Experience_Trc(j, k).Chi2)
                Next
            End If
        Next j
        xlSheet_ExpData.Cell(StartRow, StartCol + (Nb_Trc * 2) + 2).InsertData(multiList)
        multiList.Clear()

        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next


        'Exp Data_Resolution
        For j = 0 To Nb_Proc - 1
            multiList(j).Add(Info_Experience_Mat(j).Res)
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To Nb_Trc - 1
                    multiList(j).Add(Info_Experience_Trc(j, k).Res)
                Next
            End If
        Next j
        xlSheet_ExpData.Cell(StartRow, StartCol + (Nb_Trc * 3) + 3).InsertData(multiList)
        multiList.Clear()

        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        'Exp Data_CountRate
        For j = 0 To Nb_Proc - 1
            multiList(j).Add(Info_Experience_Mat(j).Count_Rate)
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To Nb_Trc - 1
                    multiList(j).Add(Info_Experience_Trc(j, k).Count_Rate)
                Next
            End If
        Next j
        xlSheet_ExpData.Cell(StartRow, StartCol + (Nb_Trc * 4) + 4).InsertData(multiList)
        multiList.Clear()

        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        'Exp Data_Pivot
        For j = 0 To Nb_Proc - 1
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To Nb_Trc - 1
                    multiList(j).Add(Info_Experience_Trc(j, k).Selected_Pivot)
                Next
            End If
        Next j
        xlSheet_ExpData.Cell(StartRow, StartCol + (Nb_Trc * 5) + 5).InsertData(multiList)
        multiList.Clear()

        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        'Exp Data_Filters
        For j = 0 To Nb_Proc - 1
            multiList(j).Add(Info_Experience_Mat(j).Filters)
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To Nb_Trc - 1
                    multiList(j).Add(Info_Experience_Trc(j, k).Filters)
                Next
            End If
        Next j
        xlSheet_ExpData.Cell(StartRow, StartCol + (Nb_Trc * 6) + 5).InsertData(multiList)
        multiList.Clear()

        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        'Exp Data _Parameter Files
        For j = 0 To Nb_Proc - 1
            multiList(j).Add(Par_Mat.Text)
            If Calcul_With_Trc = True Then ' WITH TRACE
                For k = 0 To Nb_Trc - 1
                    multiList(j).Add(Tab_File_Par_Trc(k))
                Next
            End If
        Next j
        xlSheet_ExpData.Cell(StartRow, StartCol + (Nb_Trc * 7) + 6).InsertData(multiList)
        multiList.Clear()

        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next


        'xlSheet_Mtx

        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList(j).Add(Val_Mat_Mtx_1(j, i))
            Next
        Next j
        xlSheet_Mtx.Cell(StartRow, StartCol).InsertData(multiList)
        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next


        For j = 0 To Nb_Proc - 1
            For i = 0 To Nb_Elements_Mtx_inv - 1
                multiList(j).Add(Val_Inv_Mtx(j, i))
            Next
        Next j
        xlSheet_Mtx.Cell(StartRow, StartCol + SizeTab).InsertData(multiList)
        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        'OXYDE
        If mnuOxydeOUI.Checked = True Then
            For j = 0 To Nb_Proc - 1
                For i = 0 To SizeTab
                    multiList_Int(j).Add(Val_Mat_Oxyde_1(j, i))
                Next
                If Calcul_With_Trc = True Then
                    For k = 0 To SizeTab_Trc
                        multiList_Int(j).Add(Val_Trc_Oxyde_1(j, k))
                    Next
                End If
            Next j
            xlSheet_Oxyde.Cell(StartRow, StartCol).InsertData(multiList_Int)
        End If
        multiList_Int.Clear()
        For i = 0 To Nb_Proc - 1
            multiList_Int.Add(New List(Of Integer))
        Next i

        ''############################################################# SPECIAL
        SizeTab = UBound(Val_Conc_S_100_1, 2)
        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList(j).Add(Val_Conc_S_RED_ppm_1(j, i))
            Next
        Next j
        xlSheet_S_Conc_ppm.Cell(StartRow, StartCol).InsertData(multiList)
        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList(j).Add(Val_Conc_S_RED100_1(j, i))
            Next
        Next j
        xlSheet_S_Conc_100.Cell(StartRow, StartCol).InsertData(multiList)
        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList(j).Add(Val_Conc_S_100_1(j, i))
            Next
        Next j
        xlSheet_S_Conc_100_RED.Cell(StartRow, StartCol).InsertData(multiList)

        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList(j).Add(Val_Conc_S_ppm_1(j, i))
            Next
        Next j
        xlSheet_S_Conc_ppm_RED.Cell(StartRow, StartCol).InsertData(multiList)
        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList(j).Add(Val_Choix_S_1(j, i))
            Next
        Next j
        xlSheet_Choix_S.Cell(StartRow, StartCol).InsertData(multiList)
        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next
        SizeTab = UBound(Val_Conc_And_Error_1, 2)
        'xlSheet_S_Conc_Error_ppm
        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList(j).Add(Val_Conc_And_Error_1(j, i))
            Next
        Next j
        xlSheet_S_Conc_Error_ppm.Cell(StartRow, StartCol).InsertData(multiList)
        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next

        'xlSheet_S_Conc_Error_100
        For j = 0 To Nb_Proc - 1
            For i = 0 To SizeTab
                multiList(j).Add(Val_Conc_And_Error100_1(j, i))
            Next
        Next j
        xlSheet_S_Conc_Error_100.Cell(StartRow, StartCol).InsertData(multiList)
        multiList.Clear()
        For i = 0 To Nb_Proc - 1
            multiList.Add(New List(Of String))
        Next


    End Sub


    Public Sub Function_Excel_Format_Italic(Parametres As Struct_Parametres_Thread)
        Dim LNum_File As Integer
        Dim LNb_process As Integer
        Dim Column As String
        Dim Rng_I As String

        Dim i As Integer
        LNum_File = Parametres.Num_File
        LNb_process = Parametres.Nb_Calcul
        Rng_I = ""
        If Error_Matrix(LNb_process) = True Then Exit Sub
        For i = 0 To Nb_Elem_Unique * 2 - 1
            Column = xlSheet_Conc.Cell(1, i * 2 + 4).Address.ColumnLetter
            If Rng_I = "" Then
                Rng_I = Column
            Else
                Rng_I = Rng_I & "," & Column
            End If
        Next i

        Try
            xlSheet_S_Conc_Error_ppm.Columns(Rng_I).Style.Font.Italic = True
        Catch ex As Exception
            System.Threading.Thread.Sleep(100)
            xlSheet_S_Conc_Error_ppm.Columns(Rng_I).Style.Font.Italic = True
        End Try

        Try
            xlSheet_S_Conc_Error_100.Columns(Rng_I).Style.Font.Italic = True
        Catch ex As Exception
            System.Threading.Thread.Sleep(100)
            xlSheet_S_Conc_Error_100.Columns(Rng_I).Style.Font.Italic = True
        End Try

    End Sub


    Sub Th_Excel_Format_Range_Special(Num_Fichier As Integer, Nb_Process As Integer, Nom_Excel As String)
        Dim i As Integer
        Dim p As Integer
        Dim t As Integer
        Dim Offset_Trc As Integer
        Dim Rng_Q As String
        Dim Rng_N As String
        Dim Rng_S As String
        Dim Rng_S_Err As String
        Dim Column As String
        Dim Row As String
        Dim First_Cell_N As Boolean
        Dim First_Cell_S As Boolean
        Dim First_Cell_Q As Boolean
        Dim Nb_Sheets As Integer
        Dim Num_File_Init As Integer
        Dim Num_Proc As Integer
        Num_File_Init = Num_Fichier
        Dim toto As Boolean
        Application.DoEvents()


        ' ################################################
        ' ################################################ CODE A FINIR  -> PAS ENTETE GUPIX 
        toto = False
        If toto = True Then

            First_Cell_N = True
            First_Cell_Q = True
            Nb_Sheets = 13 '8
            Num_Proc = 0

        End If
        ' ################################################
        ' ################################################ CODE A FINIR  -> PAS ENTETE GUPIX 

        '    '###########################################################################  SET FORMAT "?" ##########################################
        '    'If First_Cell_Q <> True Then '## Au moins une valeur ?
        '    '    For j = 0 To Nb_Sheets
        '    '        Rng_Q(j).Font.Color = RGB(40, 100, 170)
        '    '        Rng_Q(j).Font.Bold = True
        '    '        Rng_Q(j).Font.Italic = True
        '    '    Next
        '    'End If '

        First_Cell_N = True
        First_Cell_Q = True
        Num_Proc = 0

        Rng_Q = ""
        Rng_N = ""



        For p = 0 To Nb_Process - 1
            ToolStripStatusLabel1.Text = "Set cells font color, ? = RED, N = Blue process n°" & CStr(p)

            For i = 0 To Nb_Elements_Mat - 1 '############################################## ?/N MATRIX ##########################################

                If Val_Mat_Y_N_Q_1(Num_Proc, i) = "?" Then
                    Column = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.ColumnLetter
                    Row = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.RowNumber
                    If First_Cell_Q = False Then 'Union on ajoute au range
                        Rng_Q = Rng_Q & "," & Column & Row
                    Else
                        Rng_Q = Column & Row
                        First_Cell_Q = False
                    End If
                End If

                If Val_Mat_Y_N_Q_1(Num_Proc, i) = "N" Then
                    Column = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.ColumnLetter
                    Row = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.RowNumber

                    If First_Cell_N = False Then
                        Rng_N = Rng_N & "," & Column & Row
                    Else
                        Rng_N = Column & Row
                        First_Cell_N = False
                    End If
                End If
            Next i


            For t = 0 To Nb_Trc - 1 '############################################## ?/N TRACE ##########################################
                If t > 0 Then
                    Offset_Trc = Offset_Trc + Nb_Elements_Trc(t - 1)
                Else
                    Offset_Trc = 0
                End If

                For i = 0 To Nb_Elements_Trc(t) - 1
                    Application.DoEvents()
                    If Val_Trc_Y_N_Q_1(Num_Proc, i + Offset_Trc) = "?" Then
                        Column = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i + Nb_Elements_Mat + Offset_Trc + 3).Address.ColumnLetter
                        Row = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i + Nb_Elements_Mat + Offset_Trc + 3).Address.RowNumber

                        If First_Cell_Q = False Then
                            Rng_Q = Rng_Q & "," & Column & Row
                        Else
                            Rng_Q = Column & Row
                            First_Cell_Q = False
                        End If
                    End If

                    If Val_Trc_Y_N_Q_1(Num_Proc, i + Offset_Trc) = "N" Then
                        Column = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i + Nb_Elements_Mat + Offset_Trc + 3).Address.ColumnLetter
                        Row = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i + Nb_Elements_Mat + Offset_Trc + 3).Address.RowNumber

                        If First_Cell_N = False Then
                            Rng_N = Rng_N & "," & Column & Row
                        Else
                            Rng_N = Column & Row
                            First_Cell_N = False
                        End If
                    End If
                Next i
            Next t

            Num_Fichier = Num_Fichier + 1
            Num_Proc = Num_Proc + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
        Next p
        Application.DoEvents()
        '###########################################################################  SET FORMAT "?" ##########################################
        If First_Cell_Q <> True Then '## Au moins une valeur ?
            xlSheet_Conc.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red 'XLColor.FromArgb(255, 0, 0) '(255, 0, 0)
            xlSheet_Conc.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_Area.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red 'XLColor.FromArgb(255, 0, 0) '(255, 0, 0)
            xlSheet_Area.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_LOD.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red 'XLColor.FromArgb(255, 0, 0) '(255, 0, 0)
            xlSheet_LOD.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_Total_Error.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_Total_Error.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_Fit_Err.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_Fit_Err.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_Height.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_Height.Ranges(Rng_Q).Style.Font.Bold = True


            If mnuOxydeOUI.Checked = True Then
                xlSheet_Oxyde.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red 'XLColor.FromArgb(255, 0, 0) '(255, 0, 0)
                xlSheet_Oxyde.Ranges(Rng_Q).Style.Font.Bold = True
            End If

        End If

        '##########################################################################   SET FORMAT "N" ##########################################
        If First_Cell_N <> True Then '## Au moins une valeur ?
            xlSheet_Conc.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240) '(255, 0, 0)
            xlSheet_Conc.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_Area.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240) '(255, 0, 0)
            xlSheet_Area.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_LOD.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240) '(255, 0, 0)
            xlSheet_LOD.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_Total_Error.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240) '(255, 0, 0)
            xlSheet_Total_Error.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_Fit_Err.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240) '(255, 0, 0)
            xlSheet_Fit_Err.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_Height.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240) '(255, 0, 0)
            xlSheet_Height.Ranges(Rng_N).Style.Font.Italic = True

            If mnuOxydeOUI.Checked = True Then
                xlSheet_Oxyde.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240)
                xlSheet_Oxyde.Ranges(Rng_N).Style.Font.Italic = True
            End If

        End If


        '#################################################################### FORMAT N Feuille "S" ##################################
        Num_Proc = 0
        Num_Fichier = Num_File_Init
        First_Cell_S = True
        First_Cell_Q = True
        First_Cell_N = True
        Rng_Q = ""
        Rng_N = ""
        Rng_S = ""

        For p = 0 To Nb_Process - 1
            ToolStripStatusLabel1.Text = "Set 'S_' font color, ? = RED, N = Blue, process n°" & CStr(p)
            Application.DoEvents()
            For i = 0 To Nb_Elem_Unique - 1
                If Val_YNQ_Final(Num_Proc, i) = "?" Then

                    Column = xlSheet_S_Conc_100.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.ColumnLetter
                    Row = xlSheet_S_Conc_100.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.RowNumber

                    If First_Cell_Q = False Then
                        Rng_Q = Rng_Q & "," & Column & Row
                    Else
                        Rng_Q = Column & Row
                        First_Cell_Q = False
                    End If
                End If

                If Val_YNQ_Final(Num_Proc, i) = "N" Then

                    Column = xlSheet_S_Conc_100.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.ColumnLetter
                    Row = xlSheet_S_Conc_100.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.RowNumber

                    If First_Cell_N = False Then
                        Rng_N = Rng_N & "," & Column & Row
                    Else
                        Rng_N = Column & Row
                        First_Cell_N = False
                    End If

                End If
            Next
            Num_Fichier = Num_Fichier + 1
            Num_Proc = Num_Proc + 1
        Next p

        '###########################################################################  SET FORMAT "?" ##########################################
        If First_Cell_Q <> True Then '## Au moins une valeur ?
            xlSheet_S_Conc_100.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_S_Conc_100.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_S_Conc_ppm.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_S_Conc_ppm.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_Choix_S.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_Choix_S.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_S_Conc_ppm_RED.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_S_Conc_ppm_RED.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_S_Conc_100_RED.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_S_Conc_100_RED.Ranges(Rng_Q).Style.Font.Bold = True


        End If '

        '###########################################################################  SET FORMAT "N" ##########################################
        If First_Cell_N <> True Then '## Au moins une valeur ?
            xlSheet_S_Conc_100.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240)
            xlSheet_S_Conc_100.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_S_Conc_ppm.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240)
            xlSheet_S_Conc_ppm.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_Choix_S.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240)
            xlSheet_Choix_S.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_S_Conc_ppm_RED.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240)
            xlSheet_S_Conc_ppm_RED.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_S_Conc_100_RED.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240)
            xlSheet_S_Conc_100_RED.Ranges(Rng_N).Style.Font.Italic = True
        End If


        ''###########################################################################  FORMAT "?" et "N" Feuille "S" ERROR  ##########################################
        Num_Proc = 0
        Num_Fichier = Num_File_Init
        First_Cell_S = True
        First_Cell_Q = True
        First_Cell_N = True
        Rng_Q = ""
        Rng_N = ""
        Rng_S = ""

        For p = 0 To Nb_Process - 1
            ToolStripStatusLabel1.Text = "Set 'S_Conc_Unc' font color, ? = RED, N = Blue, process n°" & CStr(p)
            Application.DoEvents()
            For i = 0 To Nb_Elem_Unique - 1
                'ReDim New_Rng_Q(10)
                If Val_YNQ_Final_1(Num_Proc, i) = "?" Then

                    ' For p = 0 To 1
                    Column = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i * 2 + 3).Address.ColumnLetter
                    Row = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i * 2 + 3).Address.RowNumber

                    If First_Cell_Q = False Then
                        Rng_Q = Rng_Q & "," & Column & Row
                    Else
                        Rng_Q = Column & Row
                        First_Cell_Q = False
                    End If

                    ' Next p
                End If

                If Val_YNQ_Final_1(Num_Proc, i) = "N" Then

                    ' For p = 0 To 1
                    Column = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i * 2 + 3).Address.ColumnLetter
                    Row = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i * 2 + 3).Address.RowNumber

                    If First_Cell_N = False Then
                        Rng_N = Rng_N & "," & Column & Row
                    Else
                        Rng_N = Column & Row
                        First_Cell_N = False
                    End If

                    ' Next p
                End If
            Next i 'Element unique

            Num_Fichier = Num_Fichier + 1
            Num_Proc = Num_Proc + 1
        Next p 'nb_process


        If First_Cell_Q <> True Then '## Au moins une valeur N
            xlSheet_S_Conc_Error_ppm.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_S_Conc_Error_ppm.Ranges(Rng_Q).Style.Font.Bold = True
            xlSheet_S_Conc_Error_100.Ranges(Rng_Q).Style.Font.FontColor = XLColor.Red
            xlSheet_S_Conc_Error_100.Ranges(Rng_Q).Style.Font.Bold = True
        End If

        If First_Cell_N <> True Then '## Au moins une valeur N
            xlSheet_S_Conc_Error_ppm.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240) '(255, 0, 0)
            xlSheet_S_Conc_Error_ppm.Ranges(Rng_N).Style.Font.Italic = True
            xlSheet_S_Conc_Error_100.Ranges(Rng_N).Style.Font.FontColor = XLColor.FromArgb(0, 175, 240) '(255, 0, 0)
            xlSheet_S_Conc_Error_100.Ranges(Rng_N).Style.Font.Italic = True
        End If


        '######################################################################################## BACKGROUND FEUILLES "S"############################
        Dim MyColor

        For t = 0 To Nb_Trc - 1
            Num_Proc = 0
            Num_Fichier = Num_File_Init
            First_Cell_S = True
            Rng_S = ""

            For p = 0 To Nb_Process - 1
                ToolStripStatusLabel1.Text = "Set cell background color process n°" & CStr(p)
                For i = 0 To Nb_Elem_Unique - 1

                    If Val_Choix_S_1(Num_Proc, i) = NomDet_Trc(t) Then
                        Column = xlSheet_S_Conc_100.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.ColumnLetter
                        Row = xlSheet_S_Conc_100.Cell(Num_Fichier + 3 + Offset_Excel, i + 3).Address.RowNumber

                        If First_Cell_S = False Then
                            Rng_S = Rng_S & "," & Column & Row
                        Else
                            Rng_S = Column & Row
                            First_Cell_S = False
                        End If
                    End If
                Next

                Num_Fichier = Num_Fichier + 1
                Num_Proc = Num_Proc + 1
            Next p


            If First_Cell_S <> True Then '## Au moins une valeur CHOIX

                Select Case t
                    Case 0 '"X0"
                        MyColor = XLColor.FromArgb(230, 230, 230)
                    Case 1
                        MyColor = XLColor.FromArgb(200, 200, 200) ' GRIS CLAIRE
                    Case 2
                        MyColor = XLColor.FromArgb(170, 170, 170) ' ORANGE CLAIRE
                    Case 3
                        MyColor = XLColor.FromArgb(145, 145, 145)  ' JAUNE CLAIRE 
                    Case 4
                        MyColor = XLColor.FromArgb(200, 255, 200) ' VERT CLAIRE
                    Case 5
                        MyColor = XLColor.FromArgb(255, 200, 110) 'ORANGE FONCE
                    Case 6
                        MyColor = XLColor.FromArgb(145, 255, 145) ' JAUNE CLAIRE
                    Case 7
                        MyColor = XLColor.FromArgb(180, 180, 225) ' MAUVE FONCE CLAIRE
                    Case 8
                        MyColor = XLColor.FromArgb(255, 255, 200) ' JAUNE CLAIRE
                End Select

                xlSheet_S_Conc_100.ranges(Rng_S).Style.Fill.BackgroundColor = MyColor
                xlSheet_S_Conc_ppm.ranges(Rng_S).Style.Fill.BackgroundColor = MyColor
                xlSheet_S_Conc_ppm_RED.ranges(Rng_S).Style.Fill.BackgroundColor = MyColor
                xlSheet_S_Conc_100_RED.ranges(Rng_S).Style.Fill.BackgroundColor = MyColor
                xlSheet_Choix_S.ranges(Rng_S).Style.Fill.BackgroundColor = MyColor
            End If

        Next t

        '####################################### BACKGROUND S_ERROR
        Rng_S_Err = ""
        For t = 0 To Nb_Trc - 1
            Num_Proc = 0
            Num_Fichier = Num_File_Init
            First_Cell_S = True
            Rng_S_Err = ""

            For p = 0 To Nb_Process - 1
                ToolStripStatusLabel1.Text = "Set 'Conc & Error' cell background color process n°" & CStr(p)

                For i = 0 To Nb_Elem_Unique - 1

                    If Val_Choix_S_1(Num_Proc, i) = NomDet_Trc_1(t) Then

                        ' For p = 0 To 1
                        Column = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i * 2 + 3).Address.ColumnLetter
                        Row = xlSheet_Conc.Cell(Num_Fichier + 3 + Offset_Excel, i * 2 + 3).Address.RowNumber

                        If First_Cell_S = False Then
                            Rng_S_Err = Rng_S_Err & "," & Column & Row
                        Else
                            Rng_S_Err = Column & Row
                            First_Cell_S = False
                        End If
                        'Next p
                    End If

                Next i ' Element_unique

                Num_Fichier = Num_Fichier + 1
                Num_Proc = Num_Proc + 1
                ' ProgressBar1.Value = ProgressBar1.Value + 1
            Next p

            If First_Cell_S <> True Then '## Au moins une valeur CHOIX
                Select Case t
                    Case 0 '"X0"
                        MyColor = XLColor.FromArgb(230, 230, 230)
                    Case 1
                        MyColor = XLColor.FromArgb(200, 200, 200) ' GRIS CLAIRE
                    Case 2
                        MyColor = XLColor.FromArgb(170, 170, 170) ' ORANGE CLAIRE
                    Case 3
                        MyColor = XLColor.FromArgb(145, 145, 145)  ' JAUNE CLAIRE 
                    Case 4
                        MyColor = XLColor.FromArgb(200, 255, 200) ' VERT CLAIRE
                    Case 5
                        MyColor = XLColor.FromArgb(255, 200, 110) 'ORANGE FONCE
                    Case 6
                        MyColor = XLColor.FromArgb(145, 255, 145) ' JAUNE CLAIRE
                    Case 7
                        MyColor = XLColor.FromArgb(180, 180, 225) ' MAUVE FONCE CLAIRE
                    Case 8
                        MyColor = XLColor.FromArgb(255, 255, 200) ' JAUNE CLAIRE
                End Select

                xlSheet_S_Conc_Error_ppm.ranges(Rng_S_Err).Style.Fill.BackgroundColor = MyColor
                xlSheet_S_Conc_Error_100.ranges(Rng_S_Err).Style.Fill.BackgroundColor = MyColor

            End If
        Next t

    End Sub

    Function Retourne_Conc_Mat(Y_N_Q As String, Num_Proc As Integer, Indice As Integer, conc_in_oxide As Boolean) As Double()
        Dim val_return(2) As Double
        Dim i As Integer
        Dim Conc_as_Oxy As Boolean
        Dim j As Integer
        Dim t As Integer

        Conc_as_Oxy = True


        If Ck_AllAsOxy.Checked = True Then
            Conc_as_Oxy = True
        Else
            If Check_Trc_As_Oxy.Checked = True Then
                If conc_in_oxide = True Then
                    Conc_as_Oxy = True
                Else
                    Conc_as_Oxy = False
                End If

                'For t = 0 To Nb_Trc - 1
                '    For j = 0 To Nb_Elements_Trc(t) - 1
                '        If Tab_Info_Trc(t).Z(j) = Z Then 'Z présent en TRACE 
                '            Conc_as_Oxy = False 'donc Conc. MAT et TRC en élementaire
                '            For i = 0 To UBound(Tab_Trc_as_Oxy) ' Z Demandé en OXYDE OU ELEM ?
                '                If Z = Tab_Trc_as_Oxy(i) Then Conc_as_Oxy = True
                '            Next i
                '        End If

                '    Next j
                'Next t
            End If
        End If


        Select Case Y_N_Q

            Case "Y"
                If mnuOxydeOUI.Checked = True And Conc_as_Oxy = True Then
                    val_return(0) = Val_Mat_Oxyde(Num_Proc, Indice)
                    val_return(1) = 0 'Val_Mat_Oxyde(Num_Proc, Indice)
                Else
                    val_return(0) = Val_Mat_Conc(Num_Proc, Indice)
                    val_return(1) = 0 'Val_Mat_Conc(Num_Proc, Indice)
                End If

                val_return(2) = Val_Mat_Best_Yes(Num_Proc, Indice)
            Case "N"
                val_return(0) = 0
                val_return(1) = 0
                val_return(2) = Val_Mat_LOD(Num_Proc, Indice)

            Case "?"
                If mnuOxydeOUI.Checked = True And Conc_as_Oxy = True Then
                    val_return(0) = 0
                    val_return(1) = Val_Mat_Oxyde(Num_Proc, Indice)
                Else
                    val_return(0) = 0
                    val_return(1) = Val_Mat_Conc(Num_Proc, Indice)
                End If
                val_return(2) = Val_Mat_LOD(Num_Proc, Indice)
        End Select

        Return val_return
    End Function


    Function Retourne_Conc_Trc(Y_N_Q As String, Num_Proc As Integer, Indice As Integer, Z As Integer) As Double()
        Dim val_return(2) As Double
        Dim i As Integer
        Dim Conc_as_Oxy As Boolean

        Conc_as_Oxy = False

        If Ck_AllAsOxy.Checked = True Then
            Conc_as_Oxy = True
        Else
            If Check_Trc_As_Oxy.Checked = True Then
                If Array.IndexOf(Tab_Trc_as_Oxy, Z) <> -1 Then
                    Conc_as_Oxy = True
                Else
                    Conc_as_Oxy = False
                End If

                'For i = 0 To UBound(Tab_Trc_as_Oxy)
                '    If Z = Tab_Trc_as_Oxy(i) Then Conc_as_Oxy = True
                'Next i
            End If

        End If


        Select Case Y_N_Q
            Case "Y"
                If mnuOxydeOUI.Checked = True And Conc_as_Oxy = True Then
                    val_return(0) = Val_Trc_Oxyde(Num_Proc, Indice)
                    val_return(1) = 0 'Val_Trc_Oxyde(Num_Proc, Indice)
                Else
                    val_return(0) = Val_Trc_Conc(Num_Proc, Indice)
                    val_return(1) = 0 'Val_Trc_Conc(Num_Proc, Indice)
                End If
            Case "N"
                val_return(2) = Val_Trc_LOD(Num_Proc, Indice)

            Case "?"

                If Conc_as_Oxy = True Then
                    val_return(1) = Val_Trc_Oxyde(Num_Proc, Indice)
                Else
                    val_return(1) = Val_Trc_Conc(Num_Proc, Indice)
                End If
                val_return(2) = Val_Trc_LOD(Num_Proc, Indice)
        End Select

        Return val_return
    End Function
    Public Function merge_z(All_Z_Trc)
        Dim tab_all_z(100) As String
        Dim indx_z As Integer
        Dim i As Integer

        For Each Z In Tab_Info_Mat.Z
            indx_z = Array.IndexOf(tab_all_z, CStr(Z))
            If indx_z = -1 Then
                tab_all_z(i) = Z
                i += 1
            End If
        Next

        For Each Z In All_Z_Trc
            indx_z = Array.IndexOf(tab_all_z, CStr(Z))
            If indx_z = -1 Then
                tab_all_z(i) = Z
                i += 1
            End If
        Next
        ReDim Preserve tab_all_z(i - 1)
        Return tab_all_z
    End Function
    'Sub Calcul_Final_Best_Conc_New_Thread(Num_File As Integer, Num_Proc As Integer, Nb_Trc As Integer) ', nb_data_read As Integer)
    Sub Calcul_Final_Best_Conc_New_Thread(Parametres As parametres_best_conc_Thread) ', nb_data_read As Integer)

        Dim Num_proc As Integer
        Dim Num_File As Integer
        Dim Nb_Trc As Integer

        Nb_Trc = Parametres.nb_trace
        Num_proc = Parametres.num_process
        Num_File = Parametres.indx_file

        Dim i As Integer
        Dim t As Integer
        Dim J As Integer
        Dim K As Integer
        Dim Indice_Mat As Integer
        Dim Indice_Mat_0 As Integer
        Dim Indice_Mat_1 As Integer
        Dim Indice_Trc_0 As Integer
        Dim Indice_Trc_1 As Integer
        Dim Indice_Trc As Integer
        Dim Indice_Trc_10(10, 10) As Integer
        Dim Z_Mat As Integer
        Dim Z_Trc As Integer
        Dim Z_Trc_1(10) As Integer
        Dim Z As Integer
        Dim p As Integer
        Dim El_Mat As String
        Dim El_Trc As String
        Dim El_Trc_1(10) As String
        Dim Best_Stat_0 As Double
        Dim Best_Stat_1 As Double
        Dim Best_Stat_2 As Double
        Dim Best_Stat_3 As Double
        Dim Best_Stat_Trc(10, 10) As Double
        Dim No_Egal As Boolean
        Dim Best_Done As Boolean
        Dim Best_Mat As Boolean
        Dim Best_Trc1 As Boolean
        Dim Best_Trc2 As Boolean
        Dim Nb_Trc_Search As Integer
        Dim Num_Best_Trc As Integer
        Dim Best_LOD_mat_current As Integer
        Dim Best_LOD_mat_previous As Integer
        Dim Best_LOD_Trc_current As Integer
        Dim Best_LOD_Trc_previous As Integer

        Dim Look_4_Trc As Boolean

        Dim NomDet As String
        Dim Somme As Double
        Dim Somme_RED As Double

        Dim Str_Prec As String

        Dim Str_Error As String
        Dim Str_Error_V2 As String
        Dim Dbl_Error As Double
        Dim Error_100 As Double
        Dim Error_PPM As Double
        Dim S_Str As String
        Dim Val_Conc_Digit As Double
        Dim Trc_As_Elem As Boolean
        Dim Nb_Dig As Integer
        Dim Factor_Div As Integer
        Dim Pos_Inferieur As Integer
        Dim Nb_Trc_Search_Trc_Idem(10) As Integer
        Dim Ind_Best_Trc As Integer
        Dim Factor_Round As Double
        Dim Factor_Round_TRC As Double
        Dim Nb_total_elements_trc As Integer
        Dim All_Z_Trc() As Integer
        Dim Num_Trc As Integer
        Dim num_data As Integer
        Dim Nb_Calcul As Integer
        Dim Offset_Indice_Best_Trc As Integer
        Dim Offset_Trc As Integer
        Dim Y_N_Q_Prev As String
        Dim Y_N_Q As String
        Dim All_Y_N_Q() As String
        Dim Conc_Return(2) As Double
        Dim format_return(1) As String
        Dim Signe_before As String
        Dim Pos_etoile As Integer
        Dim Ind_Z_100 As Integer
        Dim Z1 As Integer
        Dim Offset1 As Integer
        Dim Err_Pivot(Nb_Trc - 1) As Double
        Dim Err_Pivot_Mat As Double
        Dim Err_Pivot_Trc(Nb_Trc - 1) As Double
        Dim Total_Error As Single
        Dim num_gamma As Integer
        Dim select_pixe_gamma As String
        Dim gamma_sum As Integer
        Dim million_norm As Integer
        Dim ind_gamma(10) As Integer
        ind_gamma = {-1, -1, -1, -1, -1, -1, -1, -1, -1, -1}
        Dim trc_ok As Boolean
        Dim mat_ok As Boolean
        Dim only_gamma_sum As Integer
        Dim el_only_gamma As Boolean
        Dim num_elem_only_gamma As Integer
        Dim indx_z_in_gamma As Integer
        Dim indx_G As Integer
        Dim indx_mat As Integer
        Dim indx_trc As Integer
        Dim indx_1_mat As Integer
        Dim indx_1_trc As Integer
        Dim Z_gamma As Integer
        Dim Comp_Ok As Boolean
        Dim High_LOD As Boolean
        Dim tab_all_Z(200) As String
        Dim conc_in_oxide As Boolean

        NomDet_Mat = CbDetMat.Text
        Nb_total_elements_trc = 0
        Y_N_Q_Prev = ""
        gamma_sum = 0
        nb_gamma_and_pixe = 0
        ToolStripStatusLabel1.Text = "Calcul Best value " & CStr(Num_proc)
        Application.DoEvents() ':System.Threading.Thread.Sleep(2000)

        For t = 0 To Nb_Trc - 1
            Nb_total_elements_trc = Nb_total_elements_trc + Nb_Elements_Trc(t)
        Next
        'Nb_total_elements_trc = Nb_total_elements_trc - (Nb_Trc - 1)

        ReDim All_Y_N_Q(Nb_total_elements_trc - 1)
        ReDim All_Z_Trc(Nb_total_elements_trc - 1)


        'Calcul Error total du PIVOT matrice
        ''''BEFORE 12/01/2022 
        '''Err_Pivot_Mat = Math.Sqrt((Val_Mat_Stat_Error(Num_Proc, Indice_Pivot_Mat(0)) ^ 2) + (Val_Mat_Fit_Error(Num_Proc, Indice_Pivot_Mat(0)) ^ 2))
        ''''BEFORE 12/01/2022 
        'AFTER 12/01/2022 ### ONLY STAT_ERROR #####


        Err_Pivot_Mat = 0
        If Indice_Pivot_Mat(0) <> -1 Then Err_Pivot_Mat = Val_Mat_Fit_Error(Num_Proc, Indice_Pivot_Mat(0))
        Offset1 = 0
        Offset_Trc = 0

        '#####################################
        For t = 0 To Nb_Trc - 1
            Err_Pivot(t) = 0
            For i = 0 To Nb_Elements_Trc(t) - 1
                'J = 0
                If Tab_Info_Trc(t).Z(i) <> 0 Then
                    All_Z_Trc(Offset_Trc) = Tab_Info_Trc(t).Z(i)
                    All_Y_N_Q(Offset_Trc) = Val_Trc_Y_N_Q(Num_proc, Offset_Trc)
                    Offset_Trc = Offset_Trc + 1
                End If
            Next

            If Use_ext_charge_Trc(Tab_Num_Trc(Num_Trc)) = False Then
                Err_Pivot_Trc(t) = Val_Trc_Fit_Error(Num_proc, Offset1 + Indice_Pivot_trc(t, 0))
            Else
                Err_Pivot_Trc(t) = 0
            End If

            Err_Pivot(t) = Math.Sqrt((Err_Pivot_Mat ^ 2) + (Err_Pivot_Trc(t) ^ 2)) ' ajoute erreur du Pivot Matrice à l'erreur Pivot Trace
            Offset1 = Offset1 + Nb_Elements_Trc(t)
        Next t


        nb_gamma_and_pixe = 0
        tab_all_Z = merge_z(All_Z_Trc) ' Merge Z MAtrix and trace pour parcourir les Z

        For Each Z In tab_all_Z
            'For Z = 11 To 92
            ' ToolStripStatusLabel1.Text = "Calcul Best value " & CStr(Num_proc) & " Z:" & CStr(Z) & " , 1/2"
            Best_Done = False
            Look_4_Trc = False
            Best_Mat = False
            Best_Trc1 = False
            Best_Trc2 = False
            Indice_Mat = -1
            Best_Stat_3 = 0
            Nb_Trc_Search = 0
            Best_Stat_2 = 100000
            Indice_Mat_0 = -1
            Indice_Mat_1 = -1
            Best_Stat_1 = 10000000
            '  Best_LOD_Mat = 0
            mat_ok = False
            trc_ok = False
            indx_1_mat = 0
            indx_1_trc = 0
            indx_mat = -1
            indx_trc = -1

            select_pixe_gamma = "pixe_mode"
            indx_z_in_gamma = Array.IndexOf(info_gamma_z, CStr(Z))

            If indx_z_in_gamma <> -1 Then
                nb_gamma_and_pixe += 1
                If CInt(gamma_conc_init(Num_File, indx_z_in_gamma)) > 0 Then
                    select_pixe_gamma = "gamma_mode" ' Conc de Z est pris dans Gamma_conc
                    num_gamma = K
                    ' Exit For
                End If
            Else
                select_pixe_gamma = "pixe_mode"
            End If

            If Z = 29 Then
                Z = 29
            End If

            Select Case select_pixe_gamma

                Case "pixe_mode"
                    indx_mat = Array.IndexOf(Tab_Info_Mat.Z, Z)

                    Do While indx_mat <> -1
                        i = indx_mat

                        If Array.IndexOf(All_Z_Trc, Z) <> -1 Then
                            If Array.IndexOf(Tab_Trc_as_Oxy, Z) <> -1 Then
                                conc_in_oxide = True
                            Else
                                conc_in_oxide = False
                            End If
                        Else
                            conc_in_oxide = True
                        End If



                        Z_Mat = Z 'Tab_Info_Mat.Z(i)
                        Y_N_Q = Val_Mat_Y_N_Q(Num_proc, indx_mat)

                        If Y_N_Q <> "N" And Info_Mat_Raie(Num_proc, indx_mat) = False Then
                            Best_Done = True
                            Y_N_Q_Prev = Y_N_Q

                            If Indice_Mat_0 = -1 Then ' 1 er valeur 
                                Indice_Mat_0 = indx_mat
                                'BEFORE 12/01/2022
                                'Best_Stat_0 = Math.Sqrt((Val_Mat_Stat_Error(Num_Proc, Indice_Mat_0) ^ 2) + (Val_Mat_Fit_Error(Num_Proc, Indice_Mat_0) ^ 2))
                                Best_Stat_0 = Val_Mat_Fit_Error(Num_proc, Indice_Mat_0)
                                If Best_Stat_0 = 0 Then Best_Stat_0 = 10000000
                            Else                        ' Seconde Z trouvé
                                Indice_Mat_1 = indx_mat
                                'BEFORE 12/01/2022
                                'Best_Stat_1 = Math.Sqrt((Val_Mat_Stat_Error(Num_Proc, Indice_Mat_1) ^ 2) + (Val_Mat_Fit_Error(Num_Proc, Indice_Mat_1) ^ 2))
                                Best_Stat_1 = Val_Mat_Fit_Error(Num_proc, Indice_Mat_1)
                                If Best_Stat_1 = 0 Then Best_Stat_1 = 10000000
                                If Z_Mat = Indice_Pivot_Mat(0) Then '######### Element Pivot on calcul Err_Pivot
                                    Err_Pivot_Mat = Best_Stat_1
                                    Val_Trc_Error_Pivot(Num_proc, Num_Trc) = Best_Stat_1
                                End If
                            End If

                            If Best_Stat_1 < Best_Stat_0 And Info_Mat_Raie(Num_proc, indx_mat) = False Then 'Second Z meilleur
                                Best_Stat_0 = Best_Stat_1
                                Best_Stat_1 = 10000000

                                Conc_Return = Retourne_Conc_Mat(Y_N_Q, Num_proc, Indice_Mat_1, conc_in_oxide) ' IF Oxyde retourne la conc. en oxyde
                                Val_Mat_Best_Yes(Num_proc, Indice_Mat_1) = Conc_Return(0)
                                Val_Mat_Best_Yes_RED(Num_proc, Indice_Mat_1) = Conc_Return(0) 'Retourne aussi valeur 
                                'Val_Mat_Best_Yes_RED(Num_Proc, Indice_Mat_1) = Conc_Return(0)

                                Val_Mat_Best_Yes(Num_proc, Indice_Mat_0) = Nothing
                                Val_Mat_Best_Yes_RED(Num_proc, Indice_Mat_0) = Nothing
                                Indice_Mat_0 = Indice_Mat_1 'Ind1 deviens le meilleur
                            Else 'Cas Simple on ecris la valeur
                                Conc_Return = Retourne_Conc_Mat(Y_N_Q, Num_proc, indx_mat, conc_in_oxide)
                                Val_Mat_Best_Yes(Num_proc, Indice_Mat_0) = Conc_Return(0) ' Si ? ou N alors Conc_Return(0) = 0
                                Val_Mat_Best_Yes_RED(Num_proc, Indice_Mat_0) = Conc_Return(0) ' Valeur conc.
                            End If
                        End If

                        'End If 'Z trouvé en MAT
                        'Next i
                        indx_mat = Array.IndexOf(Tab_Info_Mat.Z, Z, indx_mat + 1) ' Search for other Z in Mat 
                    Loop
pass_mat:   ' Only_Trace

                    If Best_Done = False Then
                        Look_4_Trc = True
                        Best_Stat_0 = 1000000
                        Y_N_Q_Prev = "N"
                    End If

                    Indice_Trc_0 = -1
                    Indice_Trc_1 = -1

                    If Z = 82 And skip_Pb_mtx = True Then
                        Best_Stat_0 = 1000000 'Permet pour le Pb de prendre la valeur en HE quelque soit sont l'erreur Total (Fit+Stat
                        Y_N_Q_Prev = "N" ' Ajout 29/09/2024
                    End If
                    indx_trc = Array.IndexOf(All_Z_Trc, Z) 'Search Z in All_Z_Trc

                    Do While indx_trc <> -1

                        'If indx_trc <> -1 Then
                        i = indx_trc

                        Z_Trc = Z 'All_Z_Trc(i)
                        Nb_Trc_Search_Trc_Idem(T) = 0
                        'If Z_Trc > Z Then 'Exit For
                        Y_N_Q = All_Y_N_Q(indx_trc)
                        Z_Trc = All_Z_Trc(indx_trc)

                        If Info_Trc_Raie(Num_proc, indx_trc) = False Then
                            Indice_Trc_1 = indx_trc
                            Offset1 = 0
                            For t = 0 To Nb_Trc - 1
                                If indx_trc <= Offset1 + Nb_Elements_Trc(t) - 1 And indx_trc >= Offset1 Then
                                    Num_Trc = t
                                    Exit For
                                End If
                                Offset1 = Offset1 + Nb_Elements_Trc(t)
                            Next t

                            If Indice_Pivot_trc(Num_Trc, 0) + Offset1 = indx_trc Then
                                Best_Stat_1 = Err_Pivot(Num_Trc) '+ Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, i) ^ 2 + Val_Trc_Stat_Error(Num_Proc, i) ^ 2)
                                Val_Trc_Pivot_Error(Num_proc, Num_Trc) = Best_Stat_1

                            Else
                                'BEFORE 12/01/2022
                                'Best_Stat_1 = Math.Sqrt(Err_Pivot(Num_Trc) ^ 2 + (Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, i) ^ 2 + Val_Trc_Stat_Error(Num_Proc, i) ^ 2) ^ 2))
                                Best_Stat_1 = Math.Sqrt(Err_Pivot(Num_Trc) ^ 2 + Val_Trc_Fit_Error(Num_proc, indx_trc) ^ 2)
                            End If


                            'BEFORE 12/01/2022
                            'If Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, i) ^ 2 + Val_Trc_Stat_Error(Num_Proc, i) ^ 2) = 0 Then Best_Stat_1 = 10000001
                            If Val_Trc_Fit_Error(Num_proc, indx_trc) = 0 Then Best_Stat_1 = 10000001
                            'Else
                            'Best_Stat_0 = Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, i) ^ 2) + (Val_Trc_Stat_Error(Num_Proc, i) ^ 2)
                            'End If


                            Comp_Ok = True
                            If Y_N_Q = "Y" And Y_N_Q_Prev <> "Y" Then
                                Comp_Ok = True
                            ElseIf Y_N_Q = "Y" And Y_N_Q_Prev = "Y" Then
                                Comp_Ok = True
                            ElseIf Y_N_Q = "?" And Y_N_Q_Prev = "?" Then
                                Comp_Ok = True
                            ElseIf Y_N_Q = "N" And Y_N_Q_Prev = "N" Then
                                Comp_Ok = True
                            ElseIf Y_N_Q = "N" And Y_N_Q_Prev <> "N" Then
                                Comp_Ok = False
                            ElseIf Y_N_Q = "?" And Y_N_Q_Prev = "Y" Then
                                Comp_Ok = False
                            End If

                            If (Best_Stat_1 < Best_Stat_0 And Comp_Ok = True) Or (Y_N_Q = "Y" And Y_N_Q_Prev <> "Y") Then

                                Conc_Return = Retourne_Conc_Trc(Y_N_Q, Num_proc, indx_trc, Z)

                                If Conc_Return(1) < 49999 Then ' ? mais valeur trop haute pour ? 'PbMa en trc par exemple
                                    Best_Stat_0 = Best_Stat_1
                                    Best_Stat_1 = 10000000

                                    Val_Trc_Best_Yes(Num_proc, Indice_Trc_1) = Conc_Return(0)
                                    If Y_N_Q = "?" Then
                                        Val_Trc_Best_Yes_RED(Num_proc, Indice_Trc_1) = Conc_Return(1)
                                    Else
                                        Val_Trc_Best_Yes_RED(Num_proc, Indice_Trc_1) = Conc_Return(0)
                                    End If


                                    If Indice_Mat_0 <> -1 Then
                                        Val_Mat_Best_Yes(Num_proc, Indice_Mat_0) = Nothing
                                        Val_Mat_Best_Yes_RED(Num_proc, Indice_Mat_0) = Nothing
                                        Indice_Mat_0 = -1
                                    ElseIf Indice_Trc_0 <> -1 Then
                                        Val_Trc_Best_Yes(Num_proc, Indice_Trc_0) = Nothing
                                        Val_Trc_Best_Yes_RED(Num_proc, Indice_Trc_0) = Nothing
                                        Indice_Trc_0 = Indice_Trc_1
                                    End If
                                    Indice_Trc_0 = Indice_Trc_1
                                    Y_N_Q_Prev = Y_N_Q
                                Else

                                End If
                            End If
                        End If


                        indx_trc = Array.IndexOf(All_Z_Trc, Z, indx_trc + 1) 'Search Z in All_Z_Trc
                    Loop

                Case "gamma_mode"

                    'For i = 0 To Nb_Elements_Mat - 1
                    indx_mat = Array.IndexOf(Tab_Info_Mat.Z, Z)

                    If indx_mat <> -1 Then
                        'i = indx_mat
                        Z_Mat = Tab_Info_Mat.Z(indx_mat)
                        indx_G = Array.IndexOf(info_gamma_z, CStr(Z))
                        If info_gamma_z(indx_G) = CStr(Z) Then
                            Val_Mat_Best_Yes(Num_proc, indx_mat) = gamma_conc_init(Num_File, indx_G)
                            Val_Mat_Best_Yes_RED(Num_proc, indx_mat) = gamma_conc_init(Num_File, indx_G)
                            gamma_sum = gamma_sum + gamma_conc_init(Num_File, indx_G)
                            ind_gamma(num_gamma) = indx_mat
                        End If
                    End If
            End Select
        Next Z

        Somme = 0
        Somme_RED = 0

        For i = 0 To Nb_Elements_Mat - 1
            On Error Resume Next
            Somme = Somme + Val_Mat_Best_Yes(Num_Proc, i)
            Somme_RED = Somme_RED + Val_Mat_Best_Yes_RED(Num_Proc, i)
        Next i


        For i = 0 To Nb_total_elements_trc - 1
            Somme = Somme + Val_Trc_Best_Yes(Num_Proc, i)
            Somme_RED = Somme_RED + Val_Trc_Best_Yes_RED(Num_Proc, i)

        Next


        el_only_gamma = True
        only_gamma_sum = 0
        num_elem_only_gamma = 0

        For i = 0 To nb_gamma - 1
            el_only_gamma = True
            Z_gamma = info_gamma_z(i)

            indx_mat = Array.IndexOf(Tab_Info_Mat.Z, Z_gamma) 'Recheche Z_gamma dans Z Matrice
            If indx_mat <> -1 Then
                el_only_gamma = False
                '    nb_gamma_and_pixe += 1
            End If

            indx_trc = Array.IndexOf(All_Z_Trc, Z_gamma) 'Recheche Z_gamma dans Z Trace
            If indx_trc <> -1 Then
                el_only_gamma = False
            End If
            If el_only_gamma = True Then
                only_gamma_sum = only_gamma_sum + gamma_conc_init(Num_File, i)
                format_return = Format_Str(gamma_conc_init(Num_File, i))
                Str_Prec = format_return(0)
                Nb_Dig = CInt(format_return(1))

                Val_Conc_S_ppm(Num_proc, Nb_Elem_Unique_sans_external + num_elem_only_gamma) = Strings.Format(Math.Round(gamma_conc_init(Num_File, i), Nb_Dig), 0)
                Val_Conc_S_100(Num_proc, Nb_Elem_Unique_sans_external + num_elem_only_gamma) = Strings.Format(gamma_conc_init(Num_File, i) / 10000, Str_Prec)
                Val_Conc_S_RED_ppm(Num_proc, Nb_Elem_Unique_sans_external + num_elem_only_gamma) = Strings.Format(Math.Round(gamma_conc_init(Num_File, i), Nb_Dig), 0)
                Val_Conc_S_RED100(Num_proc, Nb_Elem_Unique_sans_external + num_elem_only_gamma) = Strings.Format(gamma_conc_init(Num_File, i) / 10000, Str_Prec)
                Val_Choix_S(Num_proc, Nb_Elem_Unique_sans_external + num_elem_only_gamma) = "Gamma"
                Val_YNQ_Final(Num_proc, Nb_Elem_Unique_sans_external + num_elem_only_gamma) = "Y"
                Val_Error_S(Num_proc, Nb_Elem_Unique_sans_external + num_elem_only_gamma) = "n.d."

                num_elem_only_gamma += 1
            Else
                'num_elem_only_gamma += 1
            End If
        Next

        If Z_Elem_Inv <> 8 And Z_Elem_Inv <> 0 Then
            Somme = Somme + Conc_Invisible
            Somme_RED = Somme_RED + Conc_Invisible
        End If


        '#################### enleve de la somme la contriubution Gamma présent dans le pixe aussi Ex: Na2O
        Somme = Somme - gamma_sum
        Somme_RED = Somme_RED - gamma_sum
        million_norm = 1000000 - (gamma_sum + only_gamma_sum)
        num_gamma = 0

        For i = 0 To Nb_Elements_Mat - 1
            indx_G = -1
            indx_G = Array.IndexOf(info_gamma_z, CStr(Tab_Info_Mat.Z(i)))
            If indx_G <> -1 Then
                If gamma_conc(Num_File, indx_G) > 0 Then ' Conc. > 0 on prend la valeur Gamma
                    select_pixe_gamma = "gamma_mode"
                Else
                    select_pixe_gamma = "pixe_mode"
                End If
            Else
                select_pixe_gamma = "pixe_mode" ' Si pas Z en Gamma on normalize la valeur
            End If

            If select_pixe_gamma = "pixe_mode" Then
                If Chk_RoundValue.Checked = True Then
                    Total_Error = Val_Mat_Total_Error(Num_Proc, i)
                    Val_Mat_Best_Yes(Num_Proc, i) = MonArrondi_Conc(Val_Mat_Best_Yes(Num_Proc, i) * (million_norm / Somme), Total_Error)
                Else
                    Val_Mat_Best_Yes(Num_Proc, i) = Val_Mat_Best_Yes(Num_Proc, i) * (million_norm / Somme)
                End If
                Val_Mat_Best_Yes_RED(Num_Proc, i) = Val_Mat_Best_Yes_RED(Num_Proc, i) * (million_norm / Somme_RED)
            End If

        Next i

        For i = 0 To Nb_total_elements_trc - 1
            If Chk_RoundValue.Checked = True Then
                Total_Error = Val_Trc_Total_Error(Num_Proc, i)
                Val_Trc_Best_Yes(Num_Proc, i) = MonArrondi_Conc(Val_Trc_Best_Yes(Num_Proc, i) * (million_norm / Somme), Total_Error)
            Else
                Val_Trc_Best_Yes(Num_Proc, i) = Val_Trc_Best_Yes(Num_Proc, i) * (million_norm / Somme)
            End If
            Val_Trc_Best_Yes_RED(Num_Proc, i) = Val_Trc_Best_Yes_RED(Num_Proc, i) * (million_norm / Somme_RED)
            Conc_Return = Retourne_Conc_Trc("Y", Num_Proc, i, Z)
        Next


        ' For Z = 11 To 92 '############################################################# FINAL & BEST LOD #####################################################################
        For Each Z In tab_all_Z '############################################################# FINAL & BEST LOD #####################################################################
            '    ToolStripStatusLabel1.Text = "Calcul Best value " & CStr(Num_proc) & " Z:" & CStr(Z) & " , 2/2"
            Best_Done = False
            Look_4_Trc = False
            Best_Mat = False
            Best_Trc1 = False
            Best_Trc2 = False
            Indice_Mat = -1
            Best_Stat_3 = 0
            Nb_Trc_Search = 0
            Best_Stat_2 = 100000
            ' Z_Mat = Tab_Info_Mat.Z(i)
            Indice_Mat_0 = -1
            Indice_Mat_1 = -1
            Best_Stat_1 = 10000000
            Best_LOD_mat_current = 0
            Best_LOD_mat_previous = -1
            mat_ok = False
            trc_ok = False


            Ind_Z_100 = Array.IndexOf(Tab_Z_100, Z)
            select_pixe_gamma = "pixe_mode"
            Application.DoEvents()
            indx_G = -1
            indx_G = Array.IndexOf(info_gamma_z, CStr(Z))

            If indx_G <> -1 Then
                If CInt(gamma_conc_init(Num_File, indx_G)) > 0 Then
                    select_pixe_gamma = "gamma_mode"
                Else
                    select_pixe_gamma = "pixe_mode"
                End If
            Else
                select_pixe_gamma = "pixe_mode"
            End If

            Select Case select_pixe_gamma

                Case "pixe_mode"
                    indx_mat = Array.IndexOf(Tab_Info_Mat.Z, Z)

                    If indx_mat = -1 Then 'PAS DE Z EN MATRICE
                        Look_4_Trc = True
                        Best_Stat_0 = 100000000
                        Y_N_Q_Prev = ""
                        Best_LOD_mat_current = -1
                    End If

                    If Z = 82 And skip_Pb_mtx = True Then
                        Z = 82
                        indx_mat = -1
                        Y_N_Q_Prev = "N"
                    End If

                    If Z = 29 Then
                        Z = 29
                    End If

                    Do While indx_mat <> -1
                        'If indx_mat <> -1 Then
                        'i = indx_mat
                        Y_N_Q = Val_Mat_Y_N_Q(Num_Proc, indx_mat)

                        If Array.IndexOf(All_Z_Trc, Z) <> -1 Then
                            If Array.IndexOf(Tab_Trc_as_Oxy, Z) <> -1 Then
                                conc_in_oxide = True
                            Else
                                conc_in_oxide = False
                            End If
                        Else
                            conc_in_oxide = True
                        End If

                        Z_Mat = Tab_Info_Mat.Z(indx_mat)
                        Best_Done = True
                        Y_N_Q_Prev = Y_N_Q

                        If Indice_Mat_0 = -1 Then ' 1 er valeur 
                            Indice_Mat_0 = indx_mat
                            Best_Stat_0 = Val_Mat_Fit_Error(Num_Proc, Indice_Mat_0)
                            If Best_Stat_0 = 0 Then Best_Stat_0 = 10000000
                        Else                        ' Seconde Z trouvé
                            Indice_Mat_1 = indx_mat
                            Best_Stat_1 = Val_Mat_Fit_Error(Num_Proc, Indice_Mat_1)
                            If Best_Stat_1 = 0 Then Best_Stat_1 = 10000000
                        End If

                        If Val_Mat_Best_Yes(Num_proc, indx_mat) <> 0 Then ' ########################################################## VALEUR 100%  Y & ?
                            'Valeur renormalisées
                            format_return = Format_Str(Val_Mat_Best_Yes(Num_proc, indx_mat))
                            Str_Prec = format_return(0)
                            Nb_Dig = CInt(format_return(1))
                            Str_Mat_Conc_100(Num_proc, indx_mat) = Strings.Format(Math.Round((Val_Mat_Best_Yes(Num_proc, indx_mat) / 10000), Nb_Dig), Str_Prec)
                            Val_Mat_Conc_ppm(Num_proc, indx_mat) = Strings.Format(Val_Mat_Best_Yes(Num_proc, indx_mat), 0)
                            Val_Conc_S_RED_ppm(Num_proc, Ind_Z_100) = Strings.Format(Val_Mat_Best_Yes_RED(Num_proc, indx_mat), 0)
                            Val_Conc_S_RED100(Num_proc, Ind_Z_100) = Strings.Format(Math.Round((Val_Mat_Best_Yes_RED(Num_proc, indx_mat) / 10000), Nb_Dig), Str_Prec)
                            Val_Conc_S_100(Num_proc, Ind_Z_100) = Strings.Format(Math.Round((Val_Mat_Best_Yes(Num_proc, indx_mat) / 10000), Nb_Dig), Str_Prec)
                            Val_Conc_S_ppm(Num_proc, Ind_Z_100) = Strings.Format(Val_Mat_Best_Yes(Num_proc, indx_mat), 0)
                            Val_YNQ_Final(Num_proc, Ind_Z_100) = "Y"

                            Val_Choix_S(Num_proc, Ind_Z_100) = NomDet_Mat
                            ''''BEFORE 12/01/2022 
                            'Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Math.Sqrt((Math.Sqrt(Val_Mat_Fit_Error(Num_Proc, i) ^ 2 + Val_Mat_Stat_Error(Num_Proc, i) ^ 2) ^ 2)), 2) 'Strings.Format(Val_Trc_Total_Error(Num_Proc, i), "0.00") 'Val_Trc_Pivot_Error
                            Val_Error_S(Num_proc, Ind_Z_100) = Math.Round(Val_Mat_Fit_Error(Num_proc, indx_mat), 2)

                            If Indice_Mat_0 = -1 Then ' 1 er valeur 
                                Indice_Mat_0 = indx_mat
                                Best_Stat_0 = Val_Mat_Fit_Error(Num_proc, Indice_Mat_0)
                                If Best_Stat_0 = 0 Then Best_Stat_0 = 10000000
                            Else                        ' Seconde Z trouvé
                                Indice_Mat_1 = indx_mat
                                Best_Stat_1 = Val_Mat_Fit_Error(Num_proc, Indice_Mat_1)
                                If Best_Stat_1 = 0 Then Best_Stat_1 = 10000000
                            End If
                            mat_ok = True

                        ElseIf mat_ok = False Then ' VALEUR ?

                            If Y_N_Q = "?" Then ' VALEUR LOD
                                Best_Done = True

                                If Indice_Mat_1 <> -1 Then

                                    If Best_Stat_1 < Best_Stat_0 Then 'Second Z meilleur ############## INDICE MAT 1
                                        Conc_Return = Retourne_Conc_Mat(Y_N_Q, Num_proc, Indice_Mat_1, conc_in_oxide)
                                        Best_Stat_0 = Val_Mat_Fit_Error(Num_Proc, Indice_Mat_1) ' Conc_Return(2) ' Best_Stat_1 'Prend la valeur LOD
                                        Best_Stat_1 = 10000000
                                        'format_return = Format_Str(Conc_Return(2) * 3.3) ' Return LOD BEFORE 2023
                                        format_return = Format_Str(Conc_Return(2) * 3.3) ' Return LOD 2023
                                        Str_Prec = format_return(0)
                                        Nb_Dig = CInt(format_return(1))

                                        '############## 2023 --> 3.3 remplacé par 1 

                                        Str_Mat_Conc_100(Num_Proc, indx_mat) = "< " & Strings.Format((Math.Round(Conc_Return(2) * 3.3 / 10000, Nb_Dig)), Str_Prec) ''3.3 x LOD
                                        Val_Mat_Conc_ppm(Num_Proc, indx_mat) = "< " & Strings.Format(Conc_Return(2) * 3.3, 0) ''3.3 x LOD
                                        Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = Strings.Format(Conc_Return(1) * (million_norm / Somme_RED), 0)
                                        Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = Strings.Format((Math.Round((Conc_Return(1) * (million_norm / Somme_RED)) / 10000, Nb_Dig)), Str_Prec)

                                        Val_Conc_S_100(Num_Proc, Ind_Z_100) = "< " & Strings.Format((Math.Round(Conc_Return(2) * 3.3 / 10000, Nb_Dig)), Str_Prec) '3.3 x LOD
                                        Val_Conc_S_ppm(Num_Proc, Ind_Z_100) = "< " & Strings.Format(Conc_Return(2) * 3.3, 0) '3.3 x LOD
                                        Val_YNQ_Final(Num_Proc, Ind_Z_100) = "?"
                                        Val_Choix_S(Num_Proc, Ind_Z_100) = NomDet_Mat
                                        ''''BEFORE 12/01/2022
                                        'Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Math.Sqrt((Math.Sqrt(Val_Mat_Fit_Error(Num_Proc, i) ^ 2 + Val_Mat_Stat_Error(Num_Proc, i) ^ 2) ^ 2)), 2)  '"n.d."
                                        Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Val_Mat_Fit_Error(Num_Proc, indx_mat), 2)  '"n.d."

                                        Str_Mat_Conc_100(Num_Proc, Indice_Mat_0) = ""
                                        Val_Mat_Conc_ppm(Num_Proc, Indice_Mat_0) = ""
                                        Val_Mat_Conc_RED(Num_Proc, Indice_Mat_0) = ""
                                        Indice_Mat_0 = Indice_Mat_1 'Ind1 deviens le meilleur
                                        ' Val_Choix_S(Num_Proc, Ind_Z_100) = NomDet_Mat
                                        '{############################################################################# INCLURE VALEUR RED ICI
                                    End If
                                Else '                    ######################################### INDICE MAT 0
                                    Conc_Return = Retourne_Conc_Mat(Y_N_Q, Num_proc, Indice_Mat_0, conc_in_oxide)
                                    format_return = Format_Str(Conc_Return(2))
                                    Best_Stat_0 = Val_Mat_Fit_Error(Num_Proc, Indice_Mat_0) 'Conc_Return(2) 'Prend la valeur LOD
                                    Best_Stat_1 = 10000000
                                    Str_Prec = format_return(0)
                                    Nb_Dig = CInt(format_return(1))


                                    ' ##########################################    2023 3.3 remplacé par 1

                                    Str_Mat_Conc_100(Num_Proc, Indice_Mat_0) = "<" & Strings.Format((Math.Round(Conc_Return(2) * 3.3 / 10000, Nb_Dig)), Str_Prec)
                                    Val_Mat_Conc_ppm(Num_Proc, Indice_Mat_0) = "<" & Strings.Format(Conc_Return(2) * 3.3, 0)

                                    If Conc_Return(1) < limite_conc_red_ok Then ' 5 % de val en ROUGE
                                        Val_Conc_S_RED_ppm(Num_proc, Ind_Z_100) = Strings.Format(Conc_Return(1) * (million_norm / Somme_RED), 0)
                                        Val_Conc_S_RED100(Num_proc, Ind_Z_100) = Strings.Format((Math.Round((Conc_Return(1) * (million_norm / Somme_RED)) / 10000, Nb_Dig)), Str_Prec)
                                    Else
                                        Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = "<" & Strings.Format(Conc_Return(2) * 3.3, 0)
                                        Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = "<" & Strings.Format((Math.Round(Conc_Return(2) * 1 / 10000, Nb_Dig)), Str_Prec)
                                    End If

                                    Val_Conc_S_100(Num_Proc, Ind_Z_100) = "<" & Strings.Format((Math.Round(Conc_Return(2) * 3.3 / 10000, Nb_Dig)), Str_Prec)
                                    Val_Conc_S_ppm(Num_Proc, Ind_Z_100) = "<" & Strings.Format(Conc_Return(2) * 3.3, 0)
                                    Val_YNQ_Final(Num_Proc, Ind_Z_100) = "?"
                                    Val_Choix_S(Num_Proc, Ind_Z_100) = NomDet_Mat
                                    ''''BEFORE 12/01/2022
                                    'Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Math.Sqrt((Math.Sqrt(Val_Mat_Fit_Error(Num_Proc, i) ^ 2 + Val_Mat_Stat_Error(Num_Proc, i) ^ 2) ^ 2)), 2) 'Strings.Format(Conc_Return(1), 0) '"n.d."
                                    Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Val_Mat_Fit_Error(Num_Proc, indx_mat), 2)
                                End If

                            End If


                            If Y_N_Q = "N" Then ' VALEUR LOD
                                Best_Done = True

                                If Indice_Mat_1 <> -1 Then 'Deja     
                                    Conc_Return = Retourne_Conc_Mat(Y_N_Q, Num_proc, Indice_Mat_1, conc_in_oxide)
                                    Best_Stat_1 = Conc_Return(2) ' Prend la valeur LOD
                                    Best_LOD_mat_current = Conc_Return(2)
                                Else ' 1ere LOD
                                    Conc_Return = Retourne_Conc_Mat(Y_N_Q, Num_proc, Indice_Mat_0, conc_in_oxide)
                                    Best_Stat_1 = Best_Stat_0
                                    'Best_LOD_mat-previous = Best_LOD_Mat
                                    Best_LOD_mat_current = Conc_Return(2)
                                    Best_LOD_mat_previous = 1000000
                                End If
                                If Best_LOD_mat_current <= Best_LOD_mat_previous Then
                                    If Indice_Mat_1 <> -1 Then Indice_Mat_0 = Indice_Mat_1
                                    Best_Stat_0 = Best_Stat_1
                                    Best_LOD_mat_previous = Best_LOD_mat_current ' Prev LOD prend la valeur 
                                    format_return = Format_Str(Conc_Return(2)) '* 3.3)
                                    Str_Prec = format_return(0)
                                    Nb_Dig = CInt(format_return(1))
                                    Str_Mat_Conc_100(Num_Proc, indx_mat) = "< " & Strings.Format((Math.Round((Conc_Return(2) / 10000), Nb_Dig)), Str_Prec) '* 3.3)
                                    Val_Mat_Conc_ppm(Num_Proc, indx_mat) = "< " & Strings.Format(Conc_Return(2), 0) '* 3.3)
                                    Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = "< " & Strings.Format(Conc_Return(2), 0) '* 3.3)
                                    Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = "< " & Strings.Format((Math.Round((Conc_Return(2) / 10000), Nb_Dig)), Str_Prec) '* 3.3)
                                    Val_Conc_S_100(Num_Proc, Ind_Z_100) = "< " & Strings.Format((Math.Round((Conc_Return(2) / 10000), Nb_Dig)), Str_Prec) '* 3.3)
                                    Val_Conc_S_ppm(Num_Proc, Ind_Z_100) = "< " & Strings.Format(Conc_Return(2), 0) '* 3.3)
                                    Val_YNQ_Final(Num_Proc, Ind_Z_100) = "N"
                                    Val_Choix_S(Num_Proc, Ind_Z_100) = NomDet_Mat
                                    Val_Error_S(Num_Proc, Ind_Z_100) = "n.d." 'Math.Round(Val_Mat_Fit_Error(Num_Proc, i), 2) 'Val_Error_S(Num_Proc, Ind_Z_100) '"< " & Strings.Format(Conc_Return(2), 0) 'Strings.Format(Val_Mat_Total_Error(Num_Proc, i), "0.00") '"n.d." 
                                    '{############################################################################# INCLURE VALEUR RED ICI
                                Else
                                    Best_LOD_mat_current = Best_LOD_mat_previous
                                End If

                            End If

                        End If

                        indx_mat = Array.IndexOf(Tab_Info_Mat.Z, Z, indx_mat + 1)
                    Loop
                    Indice_Trc_0 = -1
                    Indice_Trc_1 = -1
                    Best_LOD_Trc_previous = 1000000

                    '############################## SEARCH IN TRACE
                    indx_trc = Array.IndexOf(All_Z_Trc, Z)
                    Do While indx_trc <> -1
                        If indx_trc <> -1 Then 'el_only_gamma = False Then ' For J = 0 To Nb_total_elements_trc - 1
                            'J = indx_trc
                            Z_Trc = All_Z_Trc(indx_trc)
                            Nb_Trc_Search_Trc_Idem(T) = 0
                            'If Z_Trc > Z and Then Exit For
                            Y_N_Q = All_Y_N_Q(indx_trc)
                            Z_Trc = All_Z_Trc(indx_trc)
                            If Z = 82 Then
                                Z = 82
                            End If
                            'If Z_Trc = Z Then

                            If Val_Trc_Best_Yes(Num_Proc, indx_trc) <> 0 Then
                                format_return = Format_Str(Val_Trc_Best_Yes(Num_Proc, indx_trc))
                                Str_Prec = format_return(0)
                                Nb_Dig = CInt(format_return(1))
                                Y_N_Q_Prev = Y_N_Q
                                ' Y_N_Q_Prev = Y_N_Q
                                Val_Trc_Conc100(Num_Proc, indx_trc) = Strings.Format((Math.Round(Val_Trc_Best_Yes(Num_Proc, indx_trc) / 10000, Nb_Dig)), Str_Prec)
                                Val_Trc_Conc_ppm(Num_Proc, indx_trc) = Strings.Format(Val_Trc_Best_Yes(Num_Proc, indx_trc), 0)
                                Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = Strings.Format(Val_Trc_Best_Yes_RED(Num_Proc, indx_trc), 0)
                                Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = Strings.Format((Math.Round(Val_Trc_Best_Yes_RED(Num_Proc, indx_trc) / 10000, Nb_Dig)), Str_Prec)
                                Val_Conc_S_100(Num_Proc, Ind_Z_100) = Strings.Format((Math.Round(Val_Trc_Best_Yes(Num_Proc, indx_trc) / 10000, Nb_Dig)), Str_Prec)
                                Val_Conc_S_ppm(Num_Proc, Ind_Z_100) = Strings.Format(Val_Trc_Best_Yes(Num_Proc, indx_trc), 0)
                                Val_YNQ_Final(Num_Proc, Ind_Z_100) = "Y"
                                ''''BEFORE 12/01/2022
                                'Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Math.Sqrt(Val_Trc_Pivot_Error(Num_Proc, T) ^ 2 + (Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, j) ^ 2 + Val_Trc_Stat_Error(Num_Proc, j) ^ 2) ^ 2)), 2) 'Strings.Format(Val_Trc_Error_Pivot(Num_Proc, j), "0.00")


                                Offset1 = 0
                                For t = 0 To Nb_Trc - 1
                                    If indx_trc <= Offset1 + Nb_Elements_Trc(t) - 1 And indx_trc >= Offset1 Then
                                        Val_Choix_S(Num_proc, Ind_Z_100) = NomDet_Trc(t)
                                        Val_Error_S(Num_proc, Ind_Z_100) = Math.Round(Math.Sqrt(Val_Trc_Pivot_Error(Num_proc, t) ^ 2 + Val_Trc_Fit_Error(Num_proc, indx_trc) ^ 2), 2)
                                        Exit For
                                    End If
                                    Offset1 = Offset1 + Nb_Elements_Trc(t)
                                Next t

                                If Indice_Mat_0 <> -1 Then
                                    Str_Mat_Conc_100(Num_Proc, Indice_Mat_0) = ""
                                    Val_Mat_Conc_ppm(Num_Proc, Indice_Mat_0) = ""
                                    Indice_Mat_0 = -1
                                ElseIf Indice_Trc_0 <> -1 Then
                                    Val_Trc_Conc100(Num_Proc, Indice_Trc_0) = ""
                                    Val_Trc_Conc_ppm(Num_Proc, Indice_Trc_0) = ""
                                    ' Val_Trc_Conc_RED(Num_Proc, Indice_Trc_0) = ""
                                    Indice_Trc_0 = Indice_Trc_1
                                End If

                                'If Indice_Trc_0 = -1 Then 'And Indice_Mat_0 = -1 Then 'PAs de matrice et 1er Valeur TRC
                                Indice_Trc_1 = indx_trc
                                trc_ok = True
                            ElseIf mat_ok = False And trc_ok = False Then '#### NE recherche LOD que si pas de valeur Y
                                'End If
                                Conc_Return = Retourne_Conc_Trc(Y_N_Q, Num_Proc, indx_trc, Z)
                                High_LOD = False
                                If Conc_Return(1) > 9999 Then High_LOD = True ' N'écris pas la valeur < 3.3 LOD si LOD > 9999 (1%) - ex: PbM en  HE 
                                If Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = "" Then High_LOD = False ' La seul valeur donc on écris qd même < 3.3 LOD même si énorme

                                If Y_N_Q = "?" And Y_N_Q_Prev <> "Y" And High_LOD = False Then 'And Y_N_Q_Prev = "N" Then
                                    Best_Stat_1 = Math.Sqrt(Val_Trc_Pivot_Error(Num_Proc, T) ^ 2 + Val_Trc_Fit_Error(Num_Proc, indx_trc) ^ 2) 'Val_Trc_LOD(Num_Proc, Indice_Trc_1)
                                    If Best_Stat_1 < Best_Stat_0 Or Y_N_Q_Prev = "N" Then
                                        Best_Stat_0 = Best_Stat_1 'LOD
                                        Y_N_Q_Prev = Y_N_Q
                                        Best_Stat_1 = 1000000
                                        'Indice_Trc_1 = i

                                        '################## BEFORE 2023 3.3 remplacé par 1 
                                        format_return = Format_Str(Conc_Return(2) * 3.3)
                                        Str_Prec = format_return(0)
                                        Nb_Dig = CInt(format_return(1))
                                        Val_Trc_Conc100(Num_Proc, Indice_Trc_1) = "<" & Strings.Format((Math.Round(Conc_Return(2) * 3.3 / 10000, Nb_Dig)), Str_Prec)
                                        Val_Trc_Conc_ppm(Num_Proc, Indice_Trc_1) = "<" & Strings.Format(Conc_Return(2) * 3.3, 0)
                                        Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = Strings.Format(Conc_Return(1) * (million_norm / Somme_RED), 0)
                                        Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = Strings.Format((Math.Round((Conc_Return(1) * (million_norm / Somme_RED)) / 10000, Nb_Dig)), Str_Prec)
                                        Val_Conc_S_100(Num_Proc, Ind_Z_100) = "<" & Strings.Format((Math.Round(Conc_Return(2) * 3.3 / 10000, Nb_Dig)), Str_Prec)
                                        Val_Conc_S_ppm(Num_Proc, Ind_Z_100) = "<" & Strings.Format(Conc_Return(2) * 3.3, 0)
                                        Val_YNQ_Final(Num_Proc, Ind_Z_100) = "?"
                                        ''''BEFORE 12/01/2022
                                        'Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Math.Sqrt(Val_Trc_Pivot_Error(Num_Proc, T) ^ 2 + (Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, j) ^ 2 + Val_Trc_Stat_Error(Num_Proc, j) ^ 2) ^ 2)), 2) 'Strings.Format(Val_Trc_Error_Pivot(Num_Proc, j), "0.00")


                                        Offset1 = 0
                                        For t = 0 To Nb_Trc - 1
                                            If indx_trc <= Offset1 + Nb_Elements_Trc(t) - 1 And indx_trc >= Offset1 Then
                                                Val_Choix_S(Num_proc, Ind_Z_100) = NomDet_Trc(t)
                                                Val_Error_S(Num_proc, Ind_Z_100) = Math.Round(Math.Sqrt(Val_Trc_Pivot_Error(Num_proc, t) ^ 2 + Val_Trc_Fit_Error(Num_proc, indx_trc) ^ 2), 2)
                                                Exit For
                                            End If
                                            Offset1 = Offset1 + Nb_Elements_Trc(t)
                                        Next t

                                        If Indice_Mat_0 <> -1 Then
                                            Str_Mat_Conc_100(Num_Proc, Indice_Mat_0) = ""
                                            Val_Mat_Conc_ppm(Num_Proc, Indice_Mat_0) = ""
                                            Indice_Mat_0 = -1
                                        ElseIf Indice_Trc_0 <> -1 Then
                                            Val_Trc_Conc100(Num_Proc, Indice_Trc_0) = ""
                                            Val_Trc_Conc_ppm(Num_Proc, Indice_Trc_0) = ""
                                            ' Val_Trc_Conc_RED(Num_Proc, Indice_Trc_0) = ""
                                            Indice_Trc_0 = Indice_Trc_1
                                        End If
                                    End If

                                ElseIf Y_N_Q_Prev = "Y" And Y_N_Q <> "N" Then
                                    Y_N_Q_Prev = Y_N_Q
                                ElseIf (Y_N_Q = "N" And Y_N_Q_Prev = "N") Or Y_N_Q_Prev = "" Then
                                    If Y_N_Q_Prev = "N" Then
                                        If Best_LOD_mat_current <> -1 Then ' Si LOD en MATRICE on compare LOD-Mat ave LOD-Trc
                                            Best_Stat_0 = Best_LOD_mat_current
                                            Best_LOD_Trc_current = Val_Trc_LOD(Num_Proc, indx_trc) 'Best_LOD_mat_current
                                            Best_LOD_Trc_previous = Best_LOD_mat_current
                                            Best_LOD_mat_current = -1
                                        Else
                                            Best_Stat_1 = Val_Trc_LOD(Num_Proc, indx_trc)
                                            Best_LOD_Trc_current = Val_Trc_LOD(Num_Proc, indx_trc)
                                        End If
                                    ElseIf Y_N_Q_Prev = "?" Or Y_N_Q_Prev = "Y" Then
                                        Best_Stat_1 = 99999999999 'Math.Sqrt(Val_Trc_Pivot_Error(Num_Proc, T) ^ 2 + Val_Trc_Fit_Error(Num_Proc, i) ^ 2)
                                    End If

                                    If Best_LOD_Trc_current <= Best_LOD_Trc_previous Then ' comparaison LOD avec LOD précédent avec Y_N_Q = N
                                        'Best_LOD_Trc_current = 0
                                        Y_N_Q_Prev = "N"
                                        Conc_Return = Retourne_Conc_Trc(Y_N_Q, Num_Proc, indx_trc, Z)
                                        Best_LOD_Trc_previous = Best_LOD_Trc_current
                                        '################## BEFORE 2023 3.3 remplacé par 1 
                                        format_return = Format_Str(Conc_Return(2) * 1)

                                        Str_Prec = format_return(0)
                                        Nb_Dig = CInt(format_return(1))
                                        Val_Trc_Conc100(Num_Proc, Indice_Trc_1) = "<" & Strings.Format((Math.Round((Conc_Return(2) / 10000) * 1, Nb_Dig)), Str_Prec)
                                        Val_Trc_Conc_ppm(Num_Proc, Indice_Trc_1) = "<" & Strings.Format(Conc_Return(2) * 1, 0)
                                        Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = "<" & Strings.Format(Conc_Return(2) * 1, 0)
                                        Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = "<" & Strings.Format((Math.Round((Conc_Return(2) / 10000) * 1, Nb_Dig)), Str_Prec)
                                        Val_Conc_S_100(Num_Proc, Ind_Z_100) = "<" & Strings.Format((Math.Round((Conc_Return(2) / 10000) * 1, Nb_Dig)), Str_Prec)
                                        Val_Conc_S_ppm(Num_Proc, Ind_Z_100) = "<" & Strings.Format(Conc_Return(2) * 1, 0)
                                        Val_YNQ_Final(Num_Proc, Ind_Z_100) = "N"
                                        Val_Error_S(Num_Proc, Ind_Z_100) = "n.d." '"<" & Strings.Format(Conc_Return(2) * 1, 0) '"n.d."
                                        Offset1 = 0
                                        For t = 0 To Nb_Trc - 1
                                            If indx_trc <= Offset1 + Nb_Elements_Trc(t) - 1 And indx_trc >= Offset1 Then
                                                Val_Choix_S(Num_proc, Ind_Z_100) = NomDet_Trc(t)
                                            End If
                                            Offset1 = Offset1 + Nb_Elements_Trc(t)
                                        Next
                                    End If
                                End If

                            End If
                            'End If

                            '      Next J 'TRACE
                        End If
                        indx_trc = Array.IndexOf(All_Z_Trc, Z, indx_trc + 1)
                    Loop


                Case "gamma_mode"
                    indx_mat = Array.IndexOf(Tab_Info_Mat.Z, Z)
                    ' i = indx_mat
                    Y_N_Q = Val_Mat_Y_N_Q(Num_Proc, indx_mat)
                    If indx_mat <> -1 Then 'Tab_Info_Mat.Z(i) = Z Then
                        format_return = Format_Str(Val_Mat_Best_Yes(Num_Proc, indx_mat))
                        Str_Prec = format_return(0)
                        Nb_Dig = CInt(format_return(1))
                        Str_Mat_Conc_100(Num_proc, Ind_Z_100) = Strings.Format(Math.Round((Val_Mat_Best_Yes(Num_proc, indx_mat) / 10000), Nb_Dig), Str_Prec)
                        Val_Mat_Conc_ppm(Num_proc, Ind_Z_100) = Strings.Format(Val_Mat_Best_Yes(Num_proc, indx_mat), 0)
                        Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = Strings.Format(Val_Mat_Best_Yes_RED(Num_Proc, i), 0)
                        Val_Conc_S_RED100(Num_proc, Ind_Z_100) = Strings.Format(Math.Round((Val_Mat_Best_Yes_RED(Num_proc, indx_mat) / 10000), Nb_Dig), Str_Prec)
                        Val_Conc_S_100(Num_proc, Ind_Z_100) = Strings.Format(Math.Round((Val_Mat_Best_Yes(Num_proc, indx_mat) / 10000), Nb_Dig), Str_Prec)
                        Val_Conc_S_ppm(Num_proc, Ind_Z_100) = Strings.Format(Val_Mat_Best_Yes(Num_proc, indx_mat), 0)
                        Val_YNQ_Final(Num_Proc, Ind_Z_100) = "Y"
                        Val_Choix_S(Num_Proc, Ind_Z_100) = "Gamma" 'NomDet_Mat
                        ''''BEFORE 12/01/2022 
                        'Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Math.Sqrt((Math.Sqrt(Val_Mat_Fit_Error(Num_Proc, i) ^ 2 + Val_Mat_Stat_Error(Num_Proc, i) ^ 2) ^ 2)), 2) 'Strings.Format(Val_Trc_Total_Error(Num_Proc, i), "0.00") 'Val_Trc_Pivot_Error
                        Val_Error_S(Num_Proc, Ind_Z_100) = "n.d." 'Math.Round(Val_Mat_Fit_Error(Num_Proc, indx), 2)
                    End If

            End Select
        Next Z

    End Sub


    Sub Calcul_Final_Best_Only_Trace_Conc_New_Thread(Num_File As Integer, Num_Proc As Integer, Nb_Trc As Integer) ', nb_data_read As Integer)
        'Sub Calcul_Final_Best_Conc_Thread(Parametres As Struct_Parametres_Thread) ', nb_data_read As Integer)
        Dim i As Integer
        Dim t As Integer
        Dim J As Integer
        Dim K As Integer
        Dim Indice_Trc_0 As Integer
        Dim Indice_Trc_1 As Integer
        Dim Indice_Trc As Integer
        Dim Indice_Trc_10(10, 10) As Integer
        Dim Z_Trc As Integer
        Dim Z_Trc_1(10) As Integer
        Dim Z As Integer
        Dim p As Integer
        Dim El_Mat As String
        Dim El_Trc As String
        Dim El_Trc_1(10) As String
        Dim Best_Stat_0 As Double
        Dim Best_Stat_1 As Double
        Dim Best_Stat_2 As Double
        Dim Best_Stat_3 As Double
        Dim Best_Stat_Trc(10, 10) As Double
        Dim No_Egal As Boolean
        Dim Best_Done As Boolean
        Dim Best_Mat As Boolean
        Dim Best_Trc1 As Boolean
        Dim Best_Trc2 As Boolean
        Dim Nb_Trc_Search As Integer
        Dim Num_Best_Trc As Integer

        Dim Look_4_Trc As Boolean
        'Dim NomDet_Mat As String
        ' Dim NomDet_Trc As String
        Dim NomDet As String
        Dim Somme As Double
        Dim Somme_RED As Double
        'ReDim Tab_Last_Entete(100)
        Dim Str_Prec As String
        NomDet_Mat = CbDetMat.Text
        Dim Str_Error As String
        Dim Str_Error_V2 As String
        Dim Dbl_Error As Double
        Dim Error_100 As Double
        Dim Error_PPM As Double
        Dim S_Str As String
        Dim Val_Conc_Digit As Double
        Dim Trc_As_Elem As Boolean
        Dim Nb_Dig As Integer
        Dim Factor_Div As Integer
        Dim Pos_Inferieur As Integer
        Dim Nb_Trc_Search_Trc_Idem(10) As Integer
        Dim Ind_Best_Trc As Integer
        Dim Factor_Round As Double
        Dim Factor_Round_TRC As Double
        Dim Nb_total_elements_trc As Integer
        Dim All_Z_Trc() As Integer

        Dim Num_Trc As Integer
        Dim num_data As Integer
        Dim Nb_Calcul As Integer
        Dim Offset_Indice_Best_Trc As Integer
        Dim Offset_Trc As Integer
        Dim Y_N_Q_Prev As String
        Dim Y_N_Q As String
        Dim All_Y_N_Q() As String
        Dim Conc_Return(2) As Double
        Dim format_return(1) As String
        Dim Signe_before As String
        Dim Pos_etoile As Integer
        Dim Ind_Z_100 As Integer
        Dim Z1 As Integer
        Dim Offset1 As Integer
        Dim Err_Pivot(Nb_Trc - 1) As Double
        Dim Err_Pivot_Mat As Double
        Dim Err_Pivot_Trc(Nb_Trc - 1) As Double
        Dim Total_Error As Single

        ' Dim K As Integer

        Nb_total_elements_trc = 0
        ToolStripStatusLabel1.Text = "Calcul Best value " & CStr(Num_Proc)
        Application.DoEvents() ':System.Threading.Thread.Sleep(2000)

        For t = 0 To Nb_Trc - 1
            Nb_total_elements_trc = Nb_total_elements_trc + Nb_Elements_Trc(t)
        Next t
        'Nb_total_elements_trc = Nb_total_elements_trc - (Nb_Trc - 1)
        On Error GoTo 10
        ReDim All_Y_N_Q(Nb_total_elements_trc - 1)
        ReDim All_Z_Trc(Nb_total_elements_trc - 1)
10:
        Offset_Trc = 0
        For t = 0 To Nb_Trc - 1
            For i = 0 To Nb_Elements_Trc(t) - 1
                'J = 0
                If Tab_Info_Trc(t).Z(i) <> 0 Then
                    All_Z_Trc(Offset_Trc) = Tab_Info_Trc(t).Z(i)
                    All_Y_N_Q(Offset_Trc) = Val_Trc_Y_N_Q(Num_Proc, Offset_Trc)
                    Offset_Trc = Offset_Trc + 1
                End If

            Next i
            'UBound(Tab_Info_Trc(T).Z)

        Next t

        Look_4_Trc = True
        Best_Stat_0 = 100000000
        Y_N_Q_Prev = "N"

        For Z = 11 To 100
            Best_Done = False
            Look_4_Trc = False
            Best_Mat = False
            Best_Trc1 = False
            Best_Trc2 = False

            Best_Stat_3 = 0
            Nb_Trc_Search = 0
            Best_Stat_2 = 100000
            Best_Stat_1 = 10000000
            Best_Stat_0 = 10000000

            If Z = 30 Then
                Z = 30
            End If

            If Best_Done = False Then

            End If

            Indice_Trc_0 = -1
            Indice_Trc_1 = -1

            For i = 0 To Nb_total_elements_trc - 1

                Z_Trc = All_Z_Trc(i)
                Nb_Trc_Search_Trc_Idem(T) = 0
                'If Z_Trc > Z Then 'Exit For
                Y_N_Q = All_Y_N_Q(i)
                Z_Trc = All_Z_Trc(i)

                If Z_Trc = Z And Info_Trc_Raie(Num_Proc, i) = False Then
                    Indice_Trc_1 = i

                    For t = 0 To Nb_Trc - 1
                        If i <= Offset1 + Nb_Elements_Trc(t) - 1 And i >= Offset1 Then Num_Trc = t
                    Next t

                    If Indice_Pivot_trc(Num_Trc, 0) = i Then
                        Best_Stat_1 = Err_Pivot(Num_Trc) '+ Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, i) ^ 2 + Val_Trc_Stat_Error(Num_Proc, i) ^ 2)
                    Else
                        ''''BEFORE 12/01/2022
                        ' Best_Stat_1 = Math.Sqrt(Err_Pivot(Num_Trc) ^ 2 + (Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, i) ^ 2 + Val_Trc_Stat_Error(Num_Proc, i) ^ 2) ^ 2))
                        Best_Stat_1 = Math.Sqrt(Err_Pivot(Num_Trc) ^ 2 + Val_Trc_Fit_Error(Num_Proc, i) ^ 2)
                    End If

                    ''''BEFORE 12/01/2022
                    'If Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, i) ^ 2 + Val_Trc_Stat_Error(Num_Proc, i) ^ 2) = 0 Then Best_Stat_1 = 10000001
                    If Val_Trc_Stat_Error(Num_Proc, i) = 0 Then Best_Stat_1 = 10000001

                    If Best_Stat_1 < Best_Stat_0 Then
                        Conc_Return = Retourne_Conc_Trc(Y_N_Q, Num_Proc, i, Z)
                        Best_Stat_0 = Best_Stat_1
                        Best_Stat_1 = 10000000
                        Val_Trc_Best_Yes(Num_Proc, Indice_Trc_1) = Conc_Return(0) ' 0 = Concentration
                        Val_Trc_Best_Yes_RED(Num_Proc, Indice_Trc_1) = Conc_Return(0) ' 0 = Concentration
                        If Indice_Trc_0 <> -1 Then
                            Val_Trc_Best_Yes(Num_Proc, Indice_Trc_0) = Nothing
                            Val_Trc_Best_Yes_RED(Num_Proc, Indice_Trc_0) = Nothing
                            Indice_Trc_0 = Indice_Trc_1
                        End If
                        Indice_Trc_0 = Indice_Trc_1
                    End If
                End If
            Next i

        Next Z

        Somme = 0
        Somme_RED = 0


        For i = 0 To Nb_total_elements_trc - 1
            Somme = Somme + Val_Trc_Best_Yes(Num_Proc, i)
            Somme_RED = Somme_RED + Val_Trc_Best_Yes_RED(Num_Proc, i)

        Next

        Look_4_Trc = True
        Best_Stat_0 = 100000000
        Y_N_Q_Prev = ""

        For Z = 11 To 100 '############################################################# BEST LOD ########################################################################

            Best_Done = False
            Look_4_Trc = False
            Best_Mat = False
            Best_Trc1 = False
            Best_Trc2 = False
            Best_Stat_3 = 0
            Nb_Trc_Search = 0
            Best_Stat_2 = 100000
            ' Z_Mat = Tab_Info_Mat.Z(i)
            Best_Stat_1 = 10000000
            Best_Stat_0 = 10000000
            Y_N_Q_Prev = ""
            For i = 0 To 50
                Z1 = Tab_Z_100(i)
                If Z1 = Z Then
                    Ind_Z_100 = i '############ Indice Z actuel dans Z_100
                    ' Exit For
                End If
            Next

            If Z = 82 Then
                Z = 82
            End If
            Application.DoEvents()


            Indice_Trc_0 = -1
            Indice_Trc_1 = -1

            'For T = 0 To Nb_Trc - 1 '############################## SEARCH IN TRACE
            For i = 0 To Nb_total_elements_trc - 1

                Z_Trc = All_Z_Trc(i)
                Nb_Trc_Search_Trc_Idem(T) = 0
                'If Z_Trc > Z and Then Exit For
                Y_N_Q = All_Y_N_Q(i)
                Z_Trc = All_Z_Trc(i)

                If Z_Trc = Z Then
                    '    Z_Trc_1(T) = Z_Elem_100_Trc(T, J)
                    '   Indice_Trc_1(T, Nb_Trc_Search_Trc_Idem(T)) = Indice_Elem_100_Trc(T, J)
                    Indice_Trc_1 = i
                    Best_Stat_1 = Val_Trc_LOD(Num_Proc, Indice_Trc_1)
                    If Y_N_Q = "Y" Then
                        If Val_Trc_Best_Yes(Num_Proc, i) <> 0 Then
                            format_return = Format_Str(Val_Trc_Best_Yes(Num_Proc, i))
                            Str_Prec = format_return(0)
                            Nb_Dig = CInt(format_return(1))
                            Y_N_Q_Prev = Y_N_Q
                            Val_Trc_Conc100(Num_Proc, i) = Strings.Format((Math.Round(Val_Trc_Best_Yes(Num_Proc, i) / 10000, Nb_Dig)), Str_Prec)
                            Val_Trc_Conc_ppm(Num_Proc, i) = Strings.Format(Val_Trc_Best_Yes(Num_Proc, i), 0)
                            Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = Strings.Format(Val_Trc_Best_Yes_RED(Num_Proc, i), 0)
                            Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = Strings.Format((Math.Round(Val_Trc_Best_Yes_RED(Num_Proc, i) / 10000, Nb_Dig)), Str_Prec)
                            Val_Conc_S_100(Num_Proc, Ind_Z_100) = Strings.Format((Math.Round(Val_Trc_Best_Yes(Num_Proc, i) / 10000, Nb_Dig)), Str_Prec)
                            Val_Conc_S_ppm(Num_Proc, Ind_Z_100) = Strings.Format(Val_Trc_Best_Yes(Num_Proc, i), 0)
                            Val_YNQ_Final(Num_Proc, Ind_Z_100) = "Y"

                            Offset1 = 0
                            For t = 0 To Nb_Trc - 1
                                If i <= Offset1 + Nb_Elements_Trc(t) - 1 And i >= Offset1 Then
                                    Val_Choix_S(Num_Proc, Ind_Z_100) = NomDet_Trc(t)
                                    Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Val_Trc_Fit_Error(Num_Proc, i), 2)
                                End If
                                Offset1 = Offset1 + Nb_Elements_Trc(t)
                            Next t

                            If Indice_Trc_0 <> -1 Then
                                Val_Trc_Conc100(Num_Proc, Indice_Trc_0) = ""
                                Val_Trc_Conc_ppm(Num_Proc, Indice_Trc_0) = ""
                                ' Val_Trc_Conc_RED(Num_Proc, Indice_Trc_0) = ""
                                Indice_Trc_0 = Indice_Trc_1
                            End If
                        End If
                        'If Indice_Trc_0 = -1 Then 'And Indice_Mat_0 = -1 Then 'PAs de matrice et 1er Valeur TRC
                        Indice_Trc_1 = i
                        Best_Stat_1 = Val_Trc_LOD(Num_Proc, Indice_Trc_1)
                        'End If
                    End If

                    If Y_N_Q = "?" And Y_N_Q_Prev <> "Y" Then 'And Y_N_Q_Prev = "N" Then
                        Best_Stat_1 = Val_Trc_LOD(Num_Proc, Indice_Trc_1)
                        If Best_Stat_1 < Best_Stat_0 Or Y_N_Q_Prev = "N" Then
                            Best_Stat_0 = Best_Stat_1 'LOD
                            Y_N_Q_Prev = Y_N_Q
                            Best_Stat_1 = 10000000
                            Indice_Trc_1 = i
                            Conc_Return = Retourne_Conc_Trc(Y_N_Q, Num_Proc, i, Z)
                            format_return = Format_Str(Conc_Return(2))
                            Str_Prec = format_return(0)
                            Nb_Dig = CInt(format_return(1))
                            Val_Trc_Conc100(Num_Proc, Indice_Trc_1) = "<" & Strings.Format((Math.Round(Conc_Return(2) / 10000, Nb_Dig)), Str_Prec)
                            Val_Trc_Conc_ppm(Num_Proc, Indice_Trc_1) = "<" & Strings.Format(Conc_Return(2), 0)

                            If Conc_Return(1) < 9999 Then
                                Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = Strings.Format(Conc_Return(1), 0)
                                Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = Strings.Format((Math.Round(Conc_Return(1) / 10000, Nb_Dig)), Str_Prec)
                            Else
                                Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = "<" & Strings.Format(Conc_Return(2), 0)
                                Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = "<" & Strings.Format((Math.Round(Conc_Return(2) / 10000, Nb_Dig)), Str_Prec)
                            End If
                            Val_Conc_S_100(Num_Proc, Ind_Z_100) = "<" & Strings.Format((Math.Round(Conc_Return(2) / 10000, Nb_Dig)), Str_Prec)
                            Val_Conc_S_ppm(Num_Proc, Ind_Z_100) = "<" & Strings.Format(Conc_Return(2), 0)
                            Val_YNQ_Final(Num_Proc, Ind_Z_100) = "?"

                            Offset1 = 0
                            For t = 0 To Nb_Trc - 1
                                If i <= Offset1 + Nb_Elements_Trc(t) - 1 And i >= Offset1 Then
                                    Val_Choix_S(Num_Proc, Ind_Z_100) = NomDet_Trc(t)
                                    Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Val_Trc_Fit_Error(Num_Proc, i), 2)
                                End If
                                Offset1 = Offset1 + Nb_Elements_Trc(t)
                            Next t

                            If Indice_Trc_0 <> -1 Then
                                Val_Trc_Conc100(Num_Proc, Indice_Trc_0) = ""
                                Val_Trc_Conc_ppm(Num_Proc, Indice_Trc_0) = ""
                                ' Val_Trc_Conc_RED(Num_Proc, Indice_Trc_0) = ""
                                Indice_Trc_0 = Indice_Trc_1
                            End If
                        End If
                    End If

                    If Y_N_Q = "N" Or Y_N_Q_Prev = "N" Or Y_N_Q_Prev = "" Then
                        Indice_Trc_1 = i
                        Best_Stat_1 = Val_Trc_LOD(Num_Proc, Indice_Trc_1)
                        If Best_Stat_1 < Best_Stat_0 Then
                            Conc_Return = Retourne_Conc_Trc(Y_N_Q, Num_Proc, i, Z)
                            format_return = Format_Str(Conc_Return(2)) '* 3.3)
                            Str_Prec = format_return(0)
                            Nb_Dig = CInt(format_return(1))
                            Val_Trc_Conc100(Num_Proc, Indice_Trc_1) = "<" & Strings.Format((Math.Round((Conc_Return(2) / 10000), Nb_Dig)), Str_Prec) 'Math.Round((Conc_Return(2) / 10000) * 3.3
                            Val_Trc_Conc_ppm(Num_Proc, Indice_Trc_1) = "<" & Strings.Format(Conc_Return(2), 0) '* 3.3)
                            Val_Conc_S_RED_ppm(Num_Proc, Ind_Z_100) = "<" & Strings.Format(Conc_Return(2), 0) '* 3.3)
                            Val_Conc_S_RED100(Num_Proc, Ind_Z_100) = "<" & Strings.Format((Math.Round((Conc_Return(2) / 10000), Nb_Dig)), Str_Prec) '* 3.3)
                            Val_Conc_S_100(Num_Proc, Ind_Z_100) = "<" & Strings.Format((Math.Round((Conc_Return(2) / 10000), Nb_Dig)), Str_Prec) '* 3.3)
                            Val_Conc_S_ppm(Num_Proc, Ind_Z_100) = "<" & Strings.Format(Conc_Return(2), 0) '* 3.3)
                            Val_YNQ_Final(Num_Proc, Ind_Z_100) = "N"
                            Best_Stat_0 = Best_Stat_1
                            Y_N_Q_Prev = Y_N_Q
                            Offset1 = 0

                            For t = 0 To Nb_Trc - 1
                                If i <= Offset1 + Nb_Elements_Trc(t) - 1 And i >= Offset1 Then
                                    Val_Choix_S(Num_Proc, Ind_Z_100) = NomDet_Trc(t)
                                    Val_Error_S(Num_Proc, Ind_Z_100) = Math.Round(Val_Trc_Fit_Error(Num_Proc, i), 2)
                                End If
                                Offset1 = Offset1 + Nb_Elements_Trc(t)
                            Next t

                        End If
                    End If


                End If
            Next i
        Next Z
    End Sub




    Private Sub mnuOxydeOUI_Click()
        mnuOxydeOUI.Checked = True
        mnuOxydeNON.Checked = False
    End Sub
    Private Sub mnuOxydeNON_Click()
        mnuOxydeOUI.Checked = False
        mnuOxydeNON.Checked = True
    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextXLS.TextChanged, Pivot_det0.TextChanged
        Pivot1(0) = Pivot_det0.Text
        'Not IsNumeric
    End Sub

    Private Sub Check_det0_CheckedChanged(sender As Object, e As EventArgs) Handles Check_det0.CheckedChanged
        Dim Det_use_Q As Boolean
        Dim i
        Par_det0.Text = ""

        If Check_det0.Checked = True Then
            Select_Par_files = 0
            Ext_Par_Trc = "*" & Check_det0.Text & ".par"
            Maj_Par_Files_Trc(Par_det0, "det0")
            'Maj_Files_Trc("det0", Check_det0.Text)

            ComboBox_Type_F.Items.Add(Check_det0.Text)
            If hdf5_mode = False Then Maj_Files_Trc("det0", Check_det0.Text)

            Det_use_Q = Det_one_use_charge(Check_det0.Text)
            If Det_use_Q = False Then
                Pivot_det0.Enabled = True
                Pivot_det0.Focus()
            Else
                Pivot_det0.Text = "File-Q"
                Pivot_det0.Enabled = False
            End If

            Pivot_det0.Focus()
        Else

            For i = 0 To ComboBox_Type_F.Items.Count - 1
                If ComboBox_Type_F.Items(i) = Check_det0.Text Then
                    ComboBox_Type_F.Items.Remove(Check_det0.Text)
                    Exit For
                End If
            Next

            Select_Par_files = -1
            Par_det0.Text = ""
            Pivot_det0.Text = ""
        End If

    End Sub



    Private Sub Check_det1_CheckedChanged(sender As Object, e As EventArgs) Handles Check_det1.CheckedChanged
        Dim Det_use_Q As Boolean
        Dim i
        If Check_det1.Checked = True Then
            Select_Par_files = 1
            Ext_Par_Trc = "*" & Check_det1.Text & ".par"
            Maj_Par_Files_Trc(Par_det1, "det1")

            If hdf5_mode = False Then Maj_Files_Trc("det1", Check_det1.Text)

            Det_use_Q = Det_one_use_charge(Check_det1.Text)
            If Det_use_Q = False Then
                Pivot_det1.Enabled = True
                Pivot_det1.Focus()
            Else
                Pivot_det1.Text = "File-Q"
                Pivot_det1.Enabled = False
            End If

        Else
            Select_Par_files = -1
            Par_det1.Text = ""
            Pivot_det1.Text = ""
            Pivot_det1.Enabled = True
            For i = 0 To ComboBox_Type_F.Items.Count - 1
                If ComboBox_Type_F.Items(i) = Check_det1.Text Then
                    ComboBox_Type_F.Items.Remove(Check_det1.Text)
                    Exit For
                End If
            Next
        End If
    End Sub



    Private Sub Check_det2_CheckedChanged(sender As Object, e As EventArgs) Handles Check_det2.CheckedChanged
        Dim Det_use_Q As Boolean
        Dim i
        If Check_det2.Checked = True Then
            Select_Par_files = 2
            Ext_Par_Trc = "*" & Check_det2.Text & ".par" '"*HE2*.par"
            ComboBox_Type_F.Items.Add(Check_det2.Text)
            Maj_Par_Files_Trc(Par_det2, "det2")
            If hdf5_mode = False Then Maj_Files_Trc("det2", Check_det2.Text)
            Det_use_Q = Det_one_use_charge(Check_det2.Text)
            If Det_use_Q = False Then
                Pivot_det2.Enabled = True
                Pivot_det2.Focus()
            Else
                Pivot_det2.Text = "File-Q"
                Pivot_det2.Enabled = False
            End If
            Pivot_det2.Focus()
        Else
            Select_Par_files = -1
            Par_det2.Text = ""
            Pivot_det2.Text = ""
            For i = 0 To ComboBox_Type_F.Items.Count - 1
                If ComboBox_Type_F.Items(i) = Check_det2.Text Then
                    ComboBox_Type_F.Items.Remove(Check_det2.Text)
                    Exit For
                End If
            Next
        End If
    End Sub





    Private Sub Check_det3_CheckedChanged(sender As Object, e As EventArgs) Handles Check_det3.CheckedChanged
        Dim Det_use_Q As Boolean
        Dim i As Integer
        If Check_det3.Checked = True Then
            Select_Par_files = 3
            Ext_Par_Trc = "*" & Check_det3.Text & ".par" '"*HE3*.par"
            ComboBox_Type_F.Items.Add(Check_det3.Text)
            Maj_Par_Files_Trc(Par_det3, "det3")
            If hdf5_mode = False Then Maj_Files_Trc("det3", Check_det3.Text)

            Det_use_Q = Det_one_use_charge(Check_det3.Text)
            If Det_use_Q = False Then
                Pivot_det3.Enabled = True
                Pivot_det3.Focus()
            Else
                Pivot_det3.Text = "File-Q"
                Pivot_det3.Enabled = False
            End If
            Pivot_det3.Focus()
        Else
            Select_Par_files = -1
            Par_det3.Text = ""
            Pivot_det3.Text = ""
            For i = 0 To ComboBox_Type_F.Items.Count - 1
                If ComboBox_Type_F.Items(i) = Check_det3.Text Then
                    ComboBox_Type_F.Items.Remove(Check_det3.Text)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub Check_det4_CheckedChanged(sender As Object, e As EventArgs) Handles Check_det4.CheckedChanged
        Dim Det_use_Q As Boolean
        Dim i As Integer

        If Check_det4.Checked = True Then
            Select_Par_files = 4
            Ext_Par_Trc = "*" & Check_det4.Text & ".par" '"*HE4*.par"
            ComboBox_Type_F.Items.Add(Check_det4.Text)
            Maj_Par_Files_Trc(Par_det4, "det4")
            If hdf5_mode = False Then Maj_Files_Trc("det4", Check_det4.Text)

            Det_use_Q = Det_one_use_charge(Check_det4.Text)

            If Det_use_Q = False Then
                Pivot_det4.Enabled = True
                Pivot_det4.Focus()
            Else
                Pivot_det4.Text = "File-Q"
                Pivot_det4.Enabled = False
            End If
            Pivot_det4.Focus()
        Else
            Select_Par_files = Select_Par_files - 1
            Par_det4.Text = ""
            Pivot_det4.Text = ""
            For i = 0 To ComboBox_Type_F.Items.Count - 1
                If ComboBox_Type_F.Items(i) = Check_det4.Text Then
                    ComboBox_Type_F.Items.Remove(Check_det4.Text)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub Check_det5_CheckedChanged(sender As Object, e As EventArgs) Handles Check_det5.CheckedChanged
        Dim Det_use_Q As Boolean
        Dim i As Integer

        If Check_det5.Checked = True Then
            Select_Par_files = 5
            Ext_Par_Trc = "*" & Check_det5.Text & ".par" '"*HE10*.par"
            ComboBox_Type_F.Items.Add(Check_det5.Text)
            Maj_Par_Files_Trc(Par_det5, "det5")
            If hdf5_mode = False Then Maj_Files_Trc("det5", Check_det5.Text)

            Det_use_Q = Det_one_use_charge(Check_det5.Text)
            If Det_use_Q = False Then
                Pivot_det5.Enabled = True
                Pivot_det5.Focus()
            Else
                Pivot_det5.Text = "File-Q"
                Pivot_det5.Enabled = False
            End If
            Pivot_det5.Focus()
        Else
            Select_Par_files = -1
            Par_det5.Text = ""
            Pivot_det5.Text = ""
            For i = 0 To ComboBox_Type_F.Items.Count - 1
                If ComboBox_Type_F.Items(i) = Check_det5.Text Then
                    ComboBox_Type_F.Items.Remove(Check_det5.Text)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub Check_det6_CheckedChanged(sender As Object, e As EventArgs) Handles Check_det6.CheckedChanged
        Dim Det_use_Q As Boolean
        Dim i As Integer

        If Check_det6.Checked = True Then
            Select_Par_files = 6
            Ext_Par_Trc = "*" & Check_det6.Text & ".par" '"*HE11*.par"
            ComboBox_Type_F.Items.Add(Check_det6.Text)
            Maj_Par_Files_Trc(Par_det6, "det6")
            If hdf5_mode = False Then Maj_Files_Trc("det6", Check_det6.Text)

            Det_use_Q = Det_one_use_charge(Check_det6.Text)
            If Det_use_Q = False Then
                Pivot_det6.Enabled = True
                Pivot_det6.Focus()
            Else
                Pivot_det6.Text = "File-Q"
                Pivot_det6.Enabled = False
            End If

            Pivot_det6.Focus()
        Else
            Select_Par_files = -1
            Par_det6.Text = ""
            Pivot_det6.Text = ""
            For i = 0 To ComboBox_Type_F.Items.Count - 1
                If ComboBox_Type_F.Items(i) = Check_det6.Text Then
                    ComboBox_Type_F.Items.Remove(Check_det6.Text)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub Check_det7_CheckedChanged(sender As Object, e As EventArgs) Handles Check_det7.CheckedChanged
        Dim Det_use_Q As Boolean
        Dim i As Integer

        If Check_det7.Checked = True Then
            Select_Par_files = 7
            Ext_Par_Trc = "*" & Check_det7.Text & ".par" '"*HE12*.par"
            ComboBox_Type_F.Items.Add(Check_det7.Text)
            Maj_Par_Files_Trc(Par_det7, "det7")
            If hdf5_mode = False Then Maj_Files_Trc("det7", Check_det7.Text)

            Det_use_Q = Det_one_use_charge(Check_det7.Text)
            If Det_use_Q = False Then
                Pivot_det7.Enabled = True
                Pivot_det7.Focus()
            Else
                Pivot_det7.Text = "File-Q"
                Pivot_det7.Enabled = False
            End If

            Pivot_det7.Focus()
        Else
            Select_Par_files = -1
            Par_det7.Text = ""
            Pivot_det7.Text = ""
            For i = 0 To ComboBox_Type_F.Items.Count - 1
                If ComboBox_Type_F.Items(i) = Check_det7.Text Then
                    ComboBox_Type_F.Items.Remove(Check_det7.Text)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub Check_det8_CheckedChanged(sender As Object, e As EventArgs) Handles Check_det8.CheckedChanged
        Dim Det_use_Q As Boolean
        Dim i As Integer

        If Check_det8.Checked = True Then
            Select_Par_files = 8
            Ext_Par_Trc = "*" & Check_det8.Text & ".par"
            Maj_Par_Files_Trc(Par_det8, "det8")
            ComboBox_Type_F.Items.Add(Check_det8.Text)
            If hdf5_mode = False Then Maj_Files_Trc("det8", Check_det8.Text)

            Det_use_Q = Det_one_use_charge(Check_det8.Text)
            If Det_use_Q = False Then
                Pivot_det8.Enabled = True
                Pivot_det8.Focus()
            Else
                Pivot_det8.Text = "File-Q"
                Pivot_det8.Enabled = False
            End If

            Pivot_det8.Focus()
        Else
            Select_Par_files = -1
            Par_det8.Text = ""
            Pivot_det8.Text = ""
            For i = 0 To ComboBox_Type_F.Items.Count - 1
                If ComboBox_Type_F.Items(i) = Check_det8.Text Then
                    ComboBox_Type_F.Items.Remove(Check_det8.Text)
                    Exit For
                End If
            Next
        End If
    End Sub


    Private Sub LstPar_Mat_MouseDoubleClick1(sender As Object, e As MouseEventArgs) Handles LstPar_Mat.MouseDoubleClick
        Dim Mysep As String = ""

        Par_Mat.Text = LstPar_Mat.SelectedItem ' & ".par"

        Try
            ComboBox_Type_F.Items.RemoveAt(0)
        Catch ex As Exception

        End Try
        ComboBox_Type_F.Items.Insert(0, CbDetMat.Text)

        ComboBox_Type_F.SelectedIndex = 0
        If (PathData IsNot Nothing) Then Mysep = "\"
        PathData = PathData + Mysep + LstPar_Mat.SelectedItem
    End Sub


    Public Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()
        CheckForIllegalCrossThreadCalls = False
        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().

    End Sub

    Private Function List_Create_Swap() As Integer
        Dim Nb_Swap As Integer
        Dim Nb_Swap1 As Integer
        Dim Nb_Swap2 As Integer
        Dim fs, fs0, File_Bat
        Dim i As Integer
        Dim Nb_swap_tmp As Integer
        Dim Rep0 As String

        SwapDrive1 = ""
        SwapDrive2 = ""
        SwapDrive3 = ""
        SwapDrive4 = ""
        '
        ' On Error Resume Next

        fs = CreateObject("Scripting.FileSystemObject")
        Rep0 = CStr(TimeOfDay.Minute) & CStr(TimeOfDay.Second) & CStr(TimeOfDay.Millisecond)
        NomOrdi = System.Net.Dns.GetHostName
        ''If NomOrdi = "Server-aglae" Then
        ''    Drive_Name = DriveInfo.GetDrives()
        ''    MkDir("c:\temp\Traupixe_tmp\" & Mid(Fichier_Matrix(0), 1, 8) & "_" & Rep0)
        ''    My.Computer.FileSystem.CreateDirectory("c:\tmp\Traupixe_tmp\" & Mid(Fichier_Matrix(0), 1, 8) & "_" & Rep0)
        ''    For Each Drive_Name As DriveInfo In DriveInfo.GetDrives()
        ''        ComboBoxDrive.Items.Add(disque)
        ''    Next
        ''End If
        'Drive_Name = DriveInfo.GetDrives()
        'MkDir("c:\temp\Traupixe_tmp\" & Mid(Fichier_Matrix(0), 1, 8) & "_" & Rep0)
        '  My.Computer.FileSystem.CreateDirectory("c:\tmp\Traupixe_tmp\" & Mid(Fichier_Matrix(0), 1, 8) & "_" & Rep0)
        'For Each Drive_Name As DriveInfo In DriveInfo.GetDrives()
        '    ComboBoxDrive.Items.Add(disque)
        'Next


        Nb_Swap = 0
        ToolStripStatusLabel1.Text = "Create Swap folder " & Mid(Fichier_Matrix(0), 1, 8)

        '  On Error Resume Next
        '        For Each Drive_Name In DriveInfo.GetDrives()

        '            Free_Space = 0
        '            Try
        '                SwapDrive_Name = Drive_Name.VolumeLabel
        '            Catch ex As Exception
        '                GoTo NextDrive
        '            End Try
        '            Free_Space = Drive_Name.TotalFreeSpace
        '            TextBox8.Text = TextBox8.Text & SwapDrive_Name & vbCrLf
        '            ToolStripStatusLabel1.Text = "Drive Name  " & SwapDrive_Name

        '            If SwapDrive1 = "" And InStr(1, SwapDrive_Name, "swap1", vbTextCompare) <> 0 And Free_Space > 100000000 Then '(100 Mo)
        '                'SwapDrive1 = Drive_Name.Name & "temp_traupixe\" & Mid(Fichier_Matrix(0), 1, 8) & "_" & Rep0
        '                SwapDrive1 = Drive_Name.Name & "tmp_traupixe\" & Mid(Fichier_Matrix(0), 1, 8) & "_" & Rep0
        '                Nb_Swap = Nb_Swap + 1
        '                My.Computer.FileSystem.CreateDirectory(SwapDrive1)
        '                Tab_Swap(0) = SwapDrive1
        '            End If

        '            If SwapDrive2 = "" And InStr(1, SwapDrive_Name, "swap2", vbTextCompare) <> 0 And Free_Space > 100000000 Then
        '                SwapDrive2 = Drive_Name.Name & "tmp_traupixe\" & Mid(Fichier_Matrix(0), 1, 8) & "_" & Rep0
        '                My.Computer.FileSystem.CreateDirectory(SwapDrive2)
        '                Tab_Swap(1) = SwapDrive2
        '            End If

        '            'If SwapDrive3 = "" And InStr(1, SwapDrive_Name, "swap3", vbTextCompare) <> 0 And Free_Space > 100000000 Then
        '            '    SwapDrive3 = Drive_Name.Name & Mid(Fichier_Matrix(0), 1, 13) & "_" & Rep0
        '            '    Nb_Swap = Nb_Swap + 1
        '            '    MkDir(SwapDrive3)
        '            'End If
        '            ' If Drive_Name.Name = "C:\" Then Exit For
        'NextDrive:
        ' Next

        SwapDrive_Name = "d:\"
        Dim myDocuPath = "c:" 'System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        'Chemin_Data = myDocuPath & "\tmp_Traupixe\"

        If Nb_Swap = 0 Then

            If Not Directory.Exists(Chemin_Processed_Data) Then
                MkDir(Chemin_Processed_Data)
            End If

            If Not Directory.Exists(myDocuPath & "\tmp_Traupixe\") Then '"d:\tmp_Traupixe") Then
                Try
                    'MkDir("d:\tmp_Traupixe\")
                    MkDir(myDocuPath & "\tmp_Traupixe\") 'Chemin_Data)
                Catch ex As Exception
                    MsgBox("Create " & Chemin_Data & " folder and retry", MsgBoxStyle.Information, "Error creating temp folder") ' 'c:\tmp_Traupixe' folder and retry", MsgBoxStyle.Information, "Error creating temp folder")
                End Try
            End If
            'SwapDrive1 = Chemin_Data & "\" & Mid(Fichier_Matrix(0), 1, 8) & "_" & Rep0

            SwapDrive1 = myDocuPath & "\tmp_Traupixe\" & Mid(Fichier_Matrix(0), 1, 8) & "_" & Rep0
            My.Computer.FileSystem.CreateDirectory(SwapDrive1)
            Tab_Swap(0) = SwapDrive1
            Nb_Swap = 1
        End If

        Nb_swap_tmp = Nb_Process / Nb_Swap
        fs0 = CreateObject("Scripting.FileSystemObject")

        Select Case (Nb_Swap)

            Case 1
                ToolStripStatusLabel1.Text = "Copy GUPIX in " & SwapDrive1 & "\" & Rep0
                For i = 0 To Nb_Process - 1
                    My.Computer.FileSystem.CreateDirectory(SwapDrive1 & "\" & Rep0 + i)

                    Try
                        My.Computer.FileSystem.CopyDirectory(Chemin_GupixWin, SwapDrive1 & "\" & Rep0 + i, True)
                    Catch ex As Exception
                        My.Computer.FileSystem.CopyDirectory(Chemin_GupixWin, "c:\" & Rep0 + i, True)
                        'My.Computer.FileSystem.CopyDirectory()
                        SwapDrive1 = "c:"
                    End Try

                    Chemin_GupixWin_Multi(i) = SwapDrive1 & "\" & Rep0 + i
                    File_Bat = File.CreateText(Chemin_GupixWin_Multi(i) & "\pixwin.bat")
                    File_Bat.writeline(Mid(SwapDrive1, 1, 2))
                    File_Bat.writeline("cd " & SwapDrive1 & "\" & Rep0 + i)
                    File_Bat.writeline(".\pixwin.exe")
                    File_Bat.Close()

                Next i

            Case 2

                Nb_Swap1 = Int(Nb_swap_tmp)
                For i = 0 To Nb_Swap1 - 1
                    'fs.CopyFolder(Chemin_GupixWin, SwapDrive1 & "\" & Rep0 + i)
                    Chemin_GupixWin_Multi(i) = SwapDrive1 & "\" & Rep0 + i
                    File_Bat = File.CreateText(Chemin_GupixWin_Multi(i) & "\pixwin.bat")
                    File_Bat.writeline(Mid(SwapDrive1, 1, 2))
                    File_Bat.writeline("cd " & SwapDrive1 & "\" & Rep0 + i)
                    File_Bat.writeline(".\pixwin.exe")
                    File_Bat.Close()

                Next i

                Nb_Swap2 = Nb_Process - Int(Nb_swap_tmp)
                For i = Nb_Swap1 To Nb_Swap1 + Nb_Swap2 - 1
                    My.Computer.FileSystem.CopyDirectory(Chemin_GupixWin, SwapDrive2 & "\" & Rep0 + i)
                    Chemin_GupixWin_Multi(i) = SwapDrive2 & "\" & Rep0 + i
                    File_Bat = File.CreateText(Chemin_GupixWin_Multi(i) & "\pixwin.bat")
                    File_Bat.writeline(Mid(SwapDrive2, 1, 2))
                    File_Bat.writeline("cd " & SwapDrive2 & "\" & Rep0 + i)
                    File_Bat.writeline(".\pixwin.exe")
                    File_Bat.Close()

                Next i


        End Select
        Global_Nb_Swap = Nb_Swap
        Return Nb_Swap

    End Function




    Function Format_Str(valeur As Double) As String()

        Dim Format_String(1) As String

        Select Case valeur

            Case 100000 To 1000000
                Format_String(0) = "0.0"
                Format_String(1) = 1
            Case 1000 To 99999
                Format_String(0) = "0.00"
                Format_String(1) = 2
            Case 100 To 999
                Format_String(0) = "0.000"
                Format_String(1) = 3
            Case 10 To 99
                Format_String(0) = "0.0000"
                Format_String(1) = 4
            Case 1 To 9
                Format_String(0) = "0.00000"
                Format_String(1) = 5
            Case 0
                Format_String(0) = "0.0"
                Format_String(1) = 1
            Case Else
                Format_String(0) = "0.00000"
                Format_String(1) = 5
        End Select

        Return Format_String
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Function MonArrondi(valeur As Integer) As Integer

        If RoundConcentrationToolStripMenuItem.Checked = True Then
            Select Case valeur
                Case 100000 To 3000000
                    valeur = CInt(Math.Round(valeur / 100) * 100)
                Case 10000 To 99999
                    valeur = CInt(Math.Round(valeur / 100) * 100)
                Case 1000 To 9999
                    valeur = CInt(Math.Round(valeur / 10)) * 10
                Case 100 To 999
                    valeur = CInt(Math.Round(valeur / 10)) * 10
                Case 0 To 99
                    valeur = valeur
                Case Else
                    valeur = valeur
            End Select
        Else
            valeur = valeur
        End If

        Return valeur
    End Function

    Function MonArrondi_Conc(valeur As Integer, Total_Error As Single) As Integer
        Dim Arr As Integer
        Arr = 1
        If valeur > 1000 Then

            Select Case valeur
                Case 100000 To 99999999
                    valeur = CInt(Math.Round(valeur / 100) * 100)
                Case 1000 To 99999
                    valeur = CInt(Math.Round(valeur / 10) * 10)
                Case Else
                    valeur = valeur
            End Select
        End If

        Return valeur
    End Function

    Sub Arrondi_Mat_Error(Num_Proc As Integer)
        Dim Total_Error As Single
        Dim i As Integer

        For i = 0 To Nb_Elements_Mat - 1
            Total_Error = Val_Mat_Total_Error(Num_Proc, i)
            Val_Mat_Final_Error(Num_Proc, i) = Math.Round(Total_Error, 2)

        Next i

    End Sub
    Sub Arrondi_Trc_Error(Num_Proc As Integer) '(Num_Trc As Integer, Num_Proc As Integer, Offset_Trc As Integer)
        Dim Total_Error As Single
        Dim i As Integer
        Dim Offset_Trc As Integer

        For t = 0 To Nb_Trc - 1
            If t > 0 Then Offset_Trc = Offset_Trc + Nb_Elements_Trc(t - 1)
            For i = 0 To Nb_Elements_Trc(t) - 1
                Total_Error = Val_Trc_Total_Error(Num_Proc, i)

                If Indice_Pivot_trc(t, 0) = i Then
                    Val_Trc_WithPivot_Error(Num_Proc, i + Offset_Trc) = Math.Round(Val_Trc_Pivot_Error(Num_Proc, t), 2)
                Else
                    Val_Trc_WithPivot_Error(Num_Proc, i + Offset_Trc) = Math.Round(Math.Sqrt(Val_Trc_Pivot_Error(Num_Proc, t) ^ 2 + Val_Trc_Fit_Error(Num_Proc, i + Offset_Trc) ^ 2), 2)
                End If

            Next i
        Next t
    End Sub

    Sub Arrondi_Mat_Elem(Num_Proc As Integer)
        Dim Total_Error As Single
        Dim i As Integer

        For i = 0 To Nb_Elements_Mat - 1
            Total_Error = Val_Mat_Total_Error(Num_Proc, i)
            Val_Mat_Conc(Num_Proc, i) = MonArrondi_Conc(Val_Mat_Conc(Num_Proc, i), Total_Error)
            ' Val_Mat_Final_Error(Num_Proc, i) = Math.Round((Total_Error / 100) * Val_Mat_Conc(Num_Proc, i), 0)
            Val_Mat_Final_Error(Num_Proc, i) = Math.Round(Total_Error, 2)

        Next i

        For i = 0 To Nb_Elem_Unique - 1
            Try
                Total_Error = Val_Error_S(Num_Proc, i)
                Val_Conc_S_RED_ppm(Num_Proc, i) = MonArrondi_Conc(Val_Conc_S_RED_ppm(Num_Proc, i), Total_Error)
                Val_Conc_S_ppm(Num_Proc, i) = MonArrondi_Conc(Val_Conc_S_ppm(Num_Proc, i), Total_Error)

            Catch ex As Exception
                'Do nothing
            End Try

        Next

    End Sub

    Sub Arrondi_Trc_Elem(Num_Proc As Integer) '(Num_Trc As Integer, Num_Proc As Integer, Offset_Trc As Integer)
        Dim Total_Error As Single
        Dim i As Integer
        Dim Offset_Trc As Integer

        For t = 0 To Nb_Trc - 1
            If t > 0 Then Offset_Trc = Offset_Trc + Nb_Elements_Trc(t - 1)
            For i = 0 To Nb_Elements_Trc(t) - 1
                Total_Error = Val_Trc_Total_Error(Num_Proc, i)
                Val_Trc_Conc(Num_Proc, i + Offset_Trc) = MonArrondi_Conc(Val_Trc_Conc(Num_Proc, i + Offset_Trc), Total_Error)
                'Val_Trc_Pivot_Error(Num_Proc, Num_Trc) = 

                ''''BEFORE 12/01/2022
                'Val_Trc_WithPivot_Error(Num_Proc, i + Offset_Trc) = Math.Round(Math.Sqrt(Val_Trc_Pivot_Error(Num_Proc, T) ^ 2 + (Math.Sqrt(Val_Trc_Fit_Error(Num_Proc, i) ^ 2 + Val_Trc_Stat_Error(Num_Proc, i) ^ 2) ^ 2)), 2)
                If Indice_Pivot_trc(t, 0) = i Then
                    Val_Trc_WithPivot_Error(Num_Proc, i + Offset_Trc) = Math.Round(Val_Trc_Pivot_Error(Num_Proc, t), 2)
                Else
                    Val_Trc_WithPivot_Error(Num_Proc, i + Offset_Trc) = Math.Round(Math.Sqrt(Val_Trc_Pivot_Error(Num_Proc, t) ^ 2 + Val_Trc_Fit_Error(Num_Proc, i + Offset_Trc) ^ 2), 2)
                End If

            Next i
        Next t
    End Sub


    Sub Arrondi_Mat_Oxyde(Num_Proc As Integer)
        Dim Total_Error As Single
        Dim i As Integer
        For i = 0 To Nb_Elements_Mat - 1
            Total_Error = Val_Mat_Total_Error(Num_Proc, i)
            Val_Mat_Oxyde(Num_Proc, i) = MonArrondi_Conc(Val_Mat_Oxyde(Num_Proc, i), Total_Error)
            '   Val_Mat_Final_Error(Num_Proc, i) = Math.Round((Total_Error / 100) * Val_Mat_Oxyde(Num_Proc, i), 0)
        Next i
    End Sub

    Sub Arrondi_Trc_Oxyde(Num_Proc As Integer)
        Dim Total_Error As Single
        Dim i As Integer
        Dim Offset_Trc As Integer

        For t = 0 To Nb_Trc - 1
            If t > 0 Then Offset_Trc = Offset_Trc + Nb_Elements_Trc(t - 1)
            For i = 0 To Nb_Elements_Trc(t) - 1
                Total_Error = Val_Trc_Error_Pivot(Num_Proc, i)
                Val_Trc_Oxyde(Num_Proc, i + Offset_Trc) = MonArrondi_Conc(Val_Trc_Oxyde(Num_Proc, i + Offset_Trc), Total_Error)
            Next i
        Next t
    End Sub

    Sub Excel_Create_Sheet_conc_and_error(Num_Proc As Integer)

        Dim i As Integer

        For i = 0 To Nb_Elem_Unique - 1
            Dim pos = InStr(Val_Conc_S_RED_ppm(Num_Proc, i), "<", CompareMethod.Text)

            If pos <> 0 Then
                Val_Conc_And_Error(Num_Proc, i * 2) = Val_Conc_S_RED_ppm(Num_Proc, i)
                Val_Conc_And_Error100(Num_Proc, i * 2) = Val_Conc_S_RED100(Num_Proc, i)
            Else
                Val_Conc_And_Error(Num_Proc, i * 2) = Val_Conc_S_RED_ppm(Num_Proc, i)
                Val_Conc_And_Error100(Num_Proc, i * 2) = Val_Conc_S_RED100(Num_Proc, i)
            End If

            Val_Conc_And_Error(Num_Proc, i * 2 + 1) = Val_Error_S(Num_Proc, i)
            Val_Conc_And_Error100(Num_Proc, i * 2 + 1) = Val_Error_S(Num_Proc, i)

        Next i


    End Sub

    Private Sub Ck_AllAsOxy_CheckedChanged(sender As Object, e As EventArgs) Handles Ck_AllAsOxy.CheckedChanged

        If Ck_AllAsOxy.Checked = True And mnuOxydeOUI.Checked = True Then
            Check_Trc_As_Oxy.Enabled = False
            Text_Lst_Ox_Trc.Text = "ALL TRACE AS OXIDE"
            Check_Trc_As_Oxy.Checked = True
        Else
            Check_Trc_As_Oxy.Enabled = True
            Check_Trc_As_Oxy.Checked = False
            'Text_Lst_Ox_Trc.Text = ""
        End If

    End Sub

    Private Sub RoundConcentrationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RoundConcentrationToolStripMenuItem.Click
        RoundConcentrationToolStripMenuItem.Checked = Not (RoundConcentrationToolStripMenuItem.Checked)
        Chk_RoundValue.Checked = Not (Chk_RoundValue.Checked)
    End Sub



    Private Sub mnuOxydeOUI_Click(sender As Object, e As EventArgs) Handles mnuOxydeOUI.Click
        mnuOxydeOUI.Checked = Not (mnuOxydeOUI.Checked)
        mnuOxydeNON.Checked = Not (mnuOxydeNON.Checked)
        Check_Trc_As_Oxy.Enabled = True
        Ck_AllAsOxy.Enabled = True
        Check_Trc_As_Oxy.Checked = True
    End Sub

    Private Sub mnuOxydeNON_Click(sender As Object, e As EventArgs) Handles mnuOxydeNON.Click
        mnuOxydeOUI.Checked = Not (mnuOxydeOUI.Checked)
        mnuOxydeNON.Checked = Not (mnuOxydeNON.Checked)
        Check_Trc_As_Oxy.Enabled = False
        Check_Trc_As_Oxy.Checked = False
        Ck_AllAsOxy.Enabled = False
        Ck_AllAsOxy.Checked = False
        Text_Lst_Ox_Trc.Text = "No oxide"

    End Sub


    Private Sub Chk_RoundValue_Click(sender As Object, e As EventArgs) Handles Chk_RoundValue.Click
        ' Chk_RoundValue.Checked = Not (Chk_RoundValue.Checked)
        RoundConcentrationToolStripMenuItem.Checked = Not (RoundConcentrationToolStripMenuItem.Checked)
    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim ComputerName As String
        Dim MesFiles(1000) As Object
        Dim Nb_Det_ToRead As Integer
        Dim MesSpectres(10, 200, 2048) As String
        Dim Attrib_Spectrum(10, 10, 10) As String
        Dim MesMaps(1000) As Object
        ComputerName = System.Net.Dns.GetHostName
        Dim HDF_as_Map As Boolean
        '        Dim File = New H5File("c:\data\")
        Dim Myh5

        Myh5 = PureHDF.H5File.OpenRead("C:\Data\2023_Data_Euphro\20230322_globals_NomProjet_IBA.hdf5") 'H5.open("C:\Data\2023_Data_Euphro\20220620_0011_Night_SIBILLA_IBA.hdf5", OpenMode.Binary)
        Dim myH5Group1 = Myh5.Group("/")
        Dim i As Integer = 0

        '        foreach(var link In group.Children)
        '{
        '    var Message = link switch
        '    {
        '        H5Group group >= $"I am a group and my name is '{group.Name}'.",
        '        H5Dataset dataset           => $"I am a dataset, call me '{dataset.Name}'.",
        '        H5CommitedDatatype datatype >= $"I am the data type '{datatype.Name}'.",
        '        H5UnresolvedLink lostLink   => $"I cannot find my link target =( shame on '{lostLink.Name}'."
        '        _                           => throw New Exception("Unknown link type");
        '    }

        '    Console.WriteLine(Message)
        '}
        '  LvFiles.Items.Clear()
        Dim Attrib
        Dim Attrib_ref_obj
        Dim Attrib_Tmp
        Dim Data
        Dim Local_Ref_DataSet_ToRead() As String


        'For Each List_Group As H5Group In myH5Group1.Children 'List les " GROUP" présent à la racine
        '    MesFiles(i) = List_Group
        '    Attrib = List_Group.Attribute("ref object")
        '    Attrib_ref_obj = Attrib.ReadString
        '    'Attrib_ref_obj.value()
        '    LvFiles.Items.Add(MesFiles(i).Name & "_" & Attrib_ref_obj(0))
        '    i += 1
        '    'Data = DataSet.Read(Of Integer)
        '    ' SubItemIndex += 1
        'Next
        Dim MyH5Group
        Try
            MyH5Group = Myh5.Group("/data")
            HDF_as_Map = True
        Catch ex As Exception
            HDF_as_Map = False
        End Try

        i = 0
        Dim DataSet
        Dim j, k
        j = 0
        i = 0
        k = 0
        Ref_DataSet_ToRead(0) = "x0"
        Ref_DataSet_ToRead(1) = "x10"
        For Each Det_ToRead As String In Ref_DataSet_ToRead
            If Det_ToRead <> "" Then
                Nb_Det_ToRead += 1
            End If
        Next
        ReDim Local_Ref_DataSet_ToRead(Nb_Det_ToRead - 1)
        For i = 0 To Nb_Det_ToRead - 1
            Local_Ref_DataSet_ToRead(i) = Ref_DataSet_ToRead(i)
        Next

        If HDF_as_Map = True Then
            For Each List_Dataset As H5Dataset In MyH5Group.Children 'List les " DATASET" présent dans le group "/data"
                MesMaps(i) = List_Dataset
                i += 1
            Next
            Dim TestN = MesMaps(0).Name
            'var group = root.Group("/my/nested/group");
            '// get dataset in groupa
            DataSet = MyH5Group.Dataset("HE1")
            Data = DataSet.Read(Of Integer)

        Else ' Ponctual DATA

            ReDim MesSpectres(LvFiles.SelectedItems.Count - 1, Nb_Det_ToRead - 1, 2048)
            ReDim Attrib_Spectrum(LvFiles.SelectedItems.Count - 1, Nb_Det_ToRead - 1, 7)
            'For Each Grp As Object In MesFiles



            For j = 0 To LvFiles.SelectedItems.Count - 1

                For Each List_Group As H5Group In myH5Group1.Children
                    Attrib = List_Group.Attribute("ref object")
                    Attrib_ref_obj = Attrib.ReadString

                    If LvFiles.SelectedItems(j).Text = List_Group.Name & "_" & Attrib_ref_obj(0) Then
                        k = 0
                        For Each Det_ToRead As String In Local_Ref_DataSet_ToRead
                            DataSet = List_Group.Dataset(Det_ToRead)
                            Data = DataSet.Read(Of Int64)
                            Attrib = DataSet.Attribute("spectrum sum")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 0) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("acquisition time")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 1) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("seconds since midnight")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 2) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("month")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 3) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("year")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 4) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("user comment")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 5) = Attrib_Tmp(0)

                            Attrib_Spectrum(j, k, 6) = Det_ToRead


                            For i = 0 To 2047
                                MesSpectres(j, k, i) = CStr(Data(i))
                            Next
                            k += 1
                        Next
                        Exit For
                    End If

                Next
            Next j


            'Dim DataSet = DataSet_Spectre.Dataset("x0")
            ' Data = DataSet_Spectre.Read(Of Integer)

            'Next


        End If
        Dim Spectre

        'PureHDF.H5Group.Get("/")
        'PureHDF.H5Group
        'PureHDF.H5Object
        'Dim i As Integer
        'For i = 0 To 100
        '    x(i) = i / 10
        '    y(i) = Math.Sin(x(i))
        '    Chart1.Series(0).Points.AddXY(x(i), y(i))

        'Next
        ' PureHDF.H5
    End Sub


    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        Dim Parametres_All_Thread As Struct_Parametres_Thread
        Dim Live As Boolean
        ReDim thread_tab_FitToPNG_TRC(2)

        thread_tab_FitToPNG_TRC(0) = New System.Threading.Thread(AddressOf Th_FitToPNG)

        Parametres_All_Thread.Num_Proc = 0
        Parametres_All_Thread.Num_File = 0
        Parametres_All_Thread.Num_Trc = 1
        Parametres_All_Thread.Num_Data = 0
        Parametres_All_Thread.Nb_Calcul = 1
        Parametres_All_Thread.voie = 2
        Parametres_All_Thread.File_Name = "totot"
        ReDim Chemin_GupixWin_Multi(1)
        Chemin_GupixWin_Multi(0) = "c:\Gupixwin\gupix\"
        ToolStripStatusLabel1.Text = "Start Thread FitToPNG"

        Dim fs As Object
        fs = CreateObject("Scripting.FileSystemObject")
        ToolStripStatusLabel1.Text = "Plot Fit To image PNG " & CStr(0)
        Application.DoEvents() : System.Threading.Thread.Sleep(500)
        Dim Pixtable_ok As Boolean
        Do
            Pixtable_ok = fs.FileExists(Chemin_GupixWin_Multi(0) & "\PIXTABLE.OUT")
            System.Threading.Thread.Sleep(20)
        Loop While Pixtable_ok = False

        thread_tab_FitToPNG_TRC(0).Start(Parametres_All_Thread)
        Chemin_Data = "C:\tmp_Traupixe"

        Do
            Live = thread_tab_FitToPNG_TRC(0).IsAlive
            Application.DoEvents() ' : Sleep(50)
            ToolStripStatusLabel1.Text = "Thread FitToPNG Running"
        Loop While Live = True
        ToolStripStatusLabel1.Text = "Thread FitToPNG Finish"
    End Sub

    Private Sub Button_Extract_Click(sender As Object, e As EventArgs) Handles Button_Extract.Click

        Dim ComputerName As String
        Dim MesFiles(1000) As Object
        Dim Nb_Det_ToRead As Integer
        Dim MesMaps(1000) As Object
        ComputerName = System.Net.Dns.GetHostName
        Dim HDF_as_Map As Boolean
        Dim Myh5

        If hdf5_mode = False Then
            Exit Sub
        End If
        Myh5 = PureHDF.H5File.OpenRead(Chemin_Data & "\" & TxtBox_HDF5_File.Text)
        Dim myH5Group1 = Myh5.Group("/")
        Dim i As Integer = 0

        Dim Attrib
        Dim Attrib_ref_obj
        Dim Attrib_Tmp
        Dim Data
        Dim Local_Ref_DataSet_ToRead(10) As String
        Dim Local_Attrib_Spectrum(20) As String
        Dim MyH5Group
        Dim Som As Long
        Dim MyHeures As String, Second As String, MyDate As String

        Dim MyStrData As String

        Dim MyYear As String
        Dim Second_Temp As Byte
        Dim Minute_Temp As Byte
        Dim Heure_Temp As Long
        Dim tps_cps As String
        Dim commentaire As String

        Try
            MyH5Group = Myh5.Group("/data")
            HDF_as_Map = True
        Catch ex As Exception
            HDF_as_Map = False
        End Try

        i = 0
        Dim DataSet
        Dim j, k
        j = 0
        i = 0
        k = 0
        '''  Ref_DataSet_ToRead(0) = "x0"
        ''''  Ref_DataSet_ToRead(1) = "x10"
        Local_Ref_DataSet_ToRead = {"x0", "x1", "x3", "x4", "x10", "x11", "x12", "x13", "g20", "g70", "r135", "r150"}

        For Each Det_ToRead As String In Local_Ref_DataSet_ToRead
            If Det_ToRead <> "" Then
                Nb_Det_ToRead += 1
            End If
        Next

        Dim Spectres_hdf5()

        'Dim AllSpectres_hdf5_gamma(LvFiles.SelectedItems.Count - 1, Nb_Det_ToRead - 1, 2048)
        'Dim AllSpectres_hdf5_rbs(LvFiles.SelectedItems.Count - 1, Nb_Det_ToRead - 1, 512)
        'Dim Attrib_Spectrum(LvFiles.SelectedItems.Count - 1, Nb_Det_ToRead - 1, 7)
        'For Each Grp As Object In MesFiles

        Try
            MkDir(Chemin_Data & "\GupixFiles\")
        Catch ex As Exception

        End Try

        For j = 0 To LvFiles.SelectedItems.Count - 1

            For Each List_Group As H5Group In myH5Group1.Children
                Attrib = List_Group.Attribute("ref object")
                Attrib_ref_obj = Attrib.ReadString


                If LvFiles.SelectedItems(j).Text = List_Group.Name & "_" & Attrib_ref_obj(0) Then
                    k = 0

                    Application.DoEvents()
                    ToolStripStatusLabel1.Text = "extract dataset:  " & List_Group.Name & "_" & Attrib_ref_obj(0)

                    For Each Det_ToRead As String In Local_Ref_DataSet_ToRead
                        Try
                            DataSet = List_Group.Dataset(Det_ToRead)
                        Catch ex As Exception
                            GoTo Myend
                        End Try

                        Data = DataSet.Read(Of Int32)
                        Attrib = DataSet.Attribute("spectrum sum")
                        Attrib_Tmp = Attrib.ReadString
                        Local_Attrib_Spectrum(0) = Attrib_Tmp(0)

                        Attrib = DataSet.Attribute("acquisition time")
                        Attrib_Tmp = Attrib.ReadString
                        Local_Attrib_Spectrum(1) = Attrib_Tmp(0)

                        Attrib = DataSet.Attribute("seconds since midnight")
                        Attrib_Tmp = Attrib.ReadString
                        Local_Attrib_Spectrum(2) = Attrib_Tmp(0)

                        Attrib = DataSet.Attribute("month")
                        Attrib_Tmp = Attrib.ReadString
                        Local_Attrib_Spectrum(3) = Attrib_Tmp(0)

                        Attrib = DataSet.Attribute("year")
                        Attrib_Tmp = Attrib.ReadString
                        Local_Attrib_Spectrum(4) = Attrib_Tmp(0)

                        Attrib = DataSet.Attribute("user comment")
                        Attrib_Tmp = Attrib.ReadString
                        Local_Attrib_Spectrum(5) = Attrib_Tmp(0)
                        Local_Attrib_Spectrum(6) = Det_ToRead

                        ReDim Spectres_hdf5(DataSet.Space.Dimensions(0))

                        Local_Attrib_Spectrum(7) = CStr(Som)
                        If Som = 0 Then
                            Som = 0
                        End If


                        MyYear = Local_Attrib_Spectrum(4)
                        MyDate = "'" & Local_Attrib_Spectrum(4) & "-" & Local_Attrib_Spectrum(3) & "-01'"

                        Second = Local_Attrib_Spectrum(2)
                        Heure_Temp = Int(CInt(Second) / 3600)
                        Minute_Temp = Int((CInt(Second) Mod 3600) / 60)
                        Second_Temp = CInt(Second) Mod 3600 - Minute_Temp * 60
                        MyHeures = IIf(Heure_Temp < 10, String.Format(Heure_Temp, "0#"), LTrim(Str(Heure_Temp))) & ":" & String.Format(Minute_Temp, "0#") & ":" & String.Format(Second_Temp, "0#")
                        tps_cps = Local_Attrib_Spectrum(1) 'ACQ TIME
                        commentaire = " " & Local_Attrib_Spectrum(5)
                        'Tab_Comment(i) = commentaire
                        MyStrData = ""
                        MyStrData = CStr(DataSet.Space.Dimensions(0)) & " 1" & vbCrLf
                        '2022 07 42292 245 668979  'DrN,1000,1000,25,25,500,40,100000,Proton , 3001 keV , 40mm He , 50 um Al + 20 um Cr , OFF , 50 um Al + 20 um Cr , 50 um Al'

                        'MyStrData = MyStrData & MyDate & " '" & MyHeures & "' " & tps_cps & "    " & Local_Attrib_Spectrum(2) & " " & commentaire & " ' " & vbCrLf FOR PAR FILE
                        MyStrData = MyStrData & MyYear & " 10 " & Local_Attrib_Spectrum(2) & " " & Local_Attrib_Spectrum(1) & " " & Local_Attrib_Spectrum(0) & " " & commentaire & " ' " & vbCrLf ' FOR SPECTRA FILE

                        For i = 0 To DataSet.Space.Dimensions(0) - 1
                            Som = Som + Data(i)
                            MyStrData = MyStrData & CStr(Data(i)) & vbCrLf ' Nb_file, Num_det, canauxSpectres_hdf5(i) = 
                        Next
                        File.WriteAllText(Chemin_Data & "\GupixFiles\" & LvFiles.SelectedItems(j).Text & "." & Det_ToRead, MyStrData & vbCrLf)

Myend:
                        k += 1
                    Next

                    Exit For
                End If

            Next
        Next j

        ToolStripStatusLabel1.Text = "extracting finish : " & Chemin_Data & "\GupixFiles\"
    End Sub

    Private Sub LvFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LvFiles.SelectedIndexChanged
        Progress.Text = "0 / " & CStr(LvFiles.SelectedItems.Count)
    End Sub


    Private Sub Check_Trc_As_Oxy_CheckedChanged(sender As Object, e As EventArgs) Handles Check_Trc_As_Oxy.CheckedChanged

        If Check_Trc_As_Oxy.Checked = False And Ck_AllAsOxy.Checked = False Or mnuOxydeOUI.Checked = False Then
            Text_Lst_Ox_Trc.Text = "No oxide"
        ElseIf Ck_AllAsOxy.Checked = True Then
            Exit Sub
        Else
            Text_Lst_Ox_Trc.Text = "19,20,25,26,82"
        End If
    End Sub





    Private Sub ComboBox_Type_Calc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Type_Calc.SelectedIndexChanged

        If Myinit = True Then

            Select Case CbDetMat.SelectedIndex
                Case 0
                    If ComboBox_Type_Calc.Text = "Ponctual" Then
                        Ext_Mat = "*.x" + CStr(CbDetMat.SelectedIndex)
                    Else
                        Ext_Mat = "*BE0*.edf"
                    End If

                    Ext_Par_Mat = "*BE0*.par"
                Case 1 To 4
                    Ext_Mat = "*.x" + CStr(CbDetMat.SelectedIndex)
                    Ext_Par_Mat = "*HE" + CStr(CbDetMat.SelectedIndex) & "*.par"
                Case 5
                    Ext_Mat = "*.x10"
                    Ext_Par_Mat = "*HE10*.par"
                Case 6
                    Ext_Mat = "*.x11"
                    Ext_Par_Mat = "*HE11*.par"
                Case 7
                    Ext_Mat = "*.x12"
                    Ext_Par_Mat = "*HE12*.par"
                Case 8
                    Ext_Mat = "*.x13"
                    Ext_Par_Mat = "*HE13*.par"
                Case 9
                    Ext_Mat = ""
                    Ext_Par_Mat = ""
                Case Else
                    Ext_Mat = "*.x12"
                    Ext_Par_Mat = "*HE12*.par"

            End Select

            If hdf5_mode = True Then
                List_HDF5_group(Chemin_hdf5)
            Else
                Maj_Files_Mat()
            End If
        End If


    End Sub

    Sub lecture_Spectres_Mat_EDF(Num_Pixel As Integer, Num_File_EDF As Integer)
        Dim i As Long
        Dim Y As Long
        Dim T As Integer
        Dim nb_point As Long
        Dim Spe(2049) As String
        Dim Mydiv As Integer
        Dim Tmp_Spe As String
        Dim Num_File_0 As Integer
        Dim fs
        Dim Mean_Somme As Double
        Dim Tab_Val_Read() As Long
        Dim br As BinaryReader
        fs = CreateObject("Scripting.FileSystemObject")

        'If Dir_EDF <> "c:" Then
        '           fs.CopyFile(Chemin_Data & "\" + Fichier_Matrix(Num_File_EDF), Dir_EDF & "\" & Fichier_Matrix(Num_File_EDF))
        'fil1 = fs.GetFile(Chemin_Data + "\" + Fichier_Matrix(i + Num_Fichier))
        'Open Dir_EDF & "\" & Fichier_Matrix(Num_File_EDF) For Binary Access Read As #1
        'Dim inputFile = IO.File.Open(Dir_EDF & "\" & Fichier_Matrix(Num_File_EDF), FileMode.Open)

        Try
            br = New BinaryReader(New FileStream(Dir_EDF & "\" & Fichier_Matrix(Num_File_EDF), FileMode.Open))
        Catch e As IOException
            Console.WriteLine(e.Message + "\n Cannot open file.")
            Return
        End Try
        'Else
        '            fs.CopyFile(Chemin_Data & "\" + Fichier_Matrix(Num_File_EDF), Dir_Calc1 & "\" & Fichier_Matrix(Num_File_EDF))
        '        Open Dir_Calc1 & "\" & Fichier_Matrix(Num_File_EDF) For Binary Access Read As #1
        'End If

        'Open Chemin + "\" + Fichier_Matrix(Num_File_EDF) For Binary Access Read As #1



        Dim MyHeader = br.ReadBytes(Myseek_Mat)
        'On Error Resume Next
        'If WhereEx = 1 Then MkDir (Chemin + "\spectres\")
        'Seek #1, (Myseek_Mat + 1) '+ (Nb_Canaux * Num_Fichier)

        'i = br.ReadInt32()
        'Console.WriteLine("Integer data: {0}", i)
        'd = br.ReadDouble()


        'Ext_Mat = ".x1"

        nb_point = CLng(CLng(Nb_Canaux) * CLng(Nb_Spectres_EDF))
        Dim bytes() As Byte = br.ReadBytes(nb_point * 4)
        ReDim Tab_Val_Read(nb_point)
        ''''''''''' BitConverter.ToInt32(a, 1)
        'Ext_Mat = ".x1"

        Num_File_0 = 0
        ''''''  Nb_File_Total = 0

        For i = 0 To Nb_Spectres_EDF - 1 'Init somme spectre
            '''''''      Tab_Sum_Spe_Mat(i) = 0
        Next i


        For i = 0 To nb_point - 1
            Tmp_Spe = Tmp_Spe & CStr(Tab_Val_Read(i)) & vbCrLf

            ''''''   Tab_Sum_Spe_Mat(Num_File_0 + Num_Pixel) = _
            '''''   Tab_Sum_Spe_Mat(Num_File_0 + Num_Pixel) + CDbl(Tab_Val_Read(i))
            Mydiv = Y \ (Nb_Canaux - 1)

            Select Case Mydiv

                Case 0 'PAS FINI DE LIRE UN SPECTRE
                    Y = Y + 1
                Case Else


                    If Mean_Somme <> 0 Then
                        '''''         Mean_Somme = (Mean_Somme + Tab_Sum_Spe_Mat(Num_File_0 + Num_Pixel)) / 2
                    Else
                        '''''           Mean_Somme = Tab_Sum_Spe_Mat(Num_File_0 + Num_Pixel)
                    End If

                    Y = 0
                    Num_File_0 = Num_File_0 + 1
                    Tmp_Spe = ""

                    ' ts.Close
            End Select

        Next i



        If Dir_EDF <> "c:" Then
            Kill(Dir_EDF & "\" & Fichier_Matrix(Num_File_EDF))
        Else
            ''''''       Kill(Dir_Calc1 & "\" & Fichier_Matrix(Num_File_EDF))
        End If
    End Sub

    Sub Lect_Dim()
        Dim Fso3
        Dim fil3, File_B3
        Dim str As String
        Dim Str1 As String
        Dim Posval As Integer
        Dim pos_coma As Integer
        Dim pos_egal As Integer


        Fso3 = CreateObject("Scripting.FileSystemObject")
        fil3 = Fso3.GetFile(Chemin_Data + "\" + Fichier_Matrix(0))
        'Set fil3 = Fso3.GetFile("c:\temp\toto.edf")

        ' Text1.Text = "Get OK"
        File_B3 = fil3.OpenAsTextStream(1)
        Posval = 0
        Str1 = ""
        '  Text1.Text = "Open OK"

        Do
            str = File_B3.readline
            Str1 = Str1 + str
            Posval = Posval + 1

            If InStr(1, str, "Dim_1", vbTextCompare) > 0 Then
                Nb_Canaux = Int(Mid(str, 9, 5))
            End If

            If InStr(1, str, "Dim_2", vbTextCompare) > 0 Then
                pos_coma = InStr(1, str, ";", vbTextCompare)
                pos_egal = InStr(1, str, "=", vbTextCompare)
                Nb_Spectres_EDF = Mid(str, 8, pos_coma - pos_egal - 1)
            End If

        Loop While InStr(1, str, "}", vbTextCompare) = 0
        ' Text1.Text = "Dim OK"
        Myseek_Mat = Len(Str1) + Posval
        File_B3.Close()
        Nb_Spectres_EDF = Nb_Spectres_EDF ' - CInt(TextOffset_Pixel)
        'Text1.Text = "Fin lect DIM"
    End Sub

    Private Sub TextF_Step_TextChanged(sender As Object, e As EventArgs) Handles TextF_Step.TextChanged
        TextF_Step.Text = Strings.Replace(TextF_Step.Text, ",", ".")
        TextF_Step.SelectionStart = TextF_Step.MaxLength
    End Sub

    Private Sub TextF_From_TextChanged(sender As Object, e As EventArgs) Handles TextF_From.TextChanged
        TextF_From.Text = Strings.Replace(TextF_From.Text, ",", ".")
        TextF_From.SelectionStart = TextF_From.MaxLength
    End Sub

    Private Sub TextF_To_TextChanged(sender As Object, e As EventArgs) Handles TextF_To.TextChanged
        TextF_To.Text = Strings.Replace(TextF_To.Text, ",", ".")
        TextF_To.SelectionStart = TextF_To.MaxLength
    End Sub

    Sub lecture_Spectres_Trc_EDF(Num_Pixel As Integer, Num_File_EDF As Integer, Num_Trc As Integer)
        Dim i As Long
        Dim Y As Long
        Dim nb_point As Long
        Dim Spe(2049) As String
        Dim Mydiv As Integer
        Dim Tmp_Spe As String
        Dim Num_File_0 As Integer


        Dim Tab_Val_Read() As Long
        Dim File_Trc As String
        Dim fs


        Select Case Tab_Num_Trc(Num_Trc)

            Case 0
                File_Trc = Fichier_Trace0(Num_File_EDF)
            Case 1
                File_Trc = Fichier_Trace1(Num_File_EDF)
            Case 2
                File_Trc = Fichier_Trace2(Num_File_EDF)
            Case 3
                File_Trc = Fichier_Trace3(Num_File_EDF)
            Case 4
                File_Trc = Fichier_Trace4(Num_File_EDF)
            Case 5 ' 1+2
                File_Trc = Fichier_Trace5(Num_File_EDF)
            Case 6 ' All
                File_Trc = Fichier_Trace6(Num_File_EDF)
            Case 7 '2+3
                File_Trc = Fichier_Trace7(Num_File_EDF)
            Case 8
                File_Trc = Fichier_Trace8(Num_File_EDF)

        End Select

        ToolStripStatusLabel1.Text = "Reading TRC spectra in EDF"
        fs = CreateObject("Scripting.FileSystemObject")


        If Len(Dir_EDF) > 0 Then
            fs.CopyFile(Chemin_Data & "\" + File_Trc, Dir_EDF & "\" & File_Trc)
            'Open Dir_EDF & "\" & File_Trc For Binary Access Read As #1
        Else
            '''''''     fs.CopyFile(Chemin_Data & "\" + File_Trc, Dir_Calc1 & "\" & File_Trc)
            'Open Dir_Calc1 & "\" & File_Trc For Binary Access Read As #1
        End If


        'Open Chemin & "\" & file_trc For Binary Access Read As #1 'Fichier_Trace5(Num_File_EDF)

        'On Error Resume Next
        'If WhereEx = 1 Then MkDir (Chemin + "\spectres\")
        '     Seek #1, (Myseek_Trc + 1) '+ (Nb_Canaux * Num_Fichier)
        '

        nb_point = CLng(CLng(Nb_Canaux) * CLng(Nb_Spectres_EDF))
        ReDim Tab_Val_Read(nb_point)

        '   Nb_File_Total = 0
        Num_File_0 = 0

        For i = 0 To nb_point - 1

            Tmp_Spe = Tmp_Spe & CStr(Tab_Val_Read(i)) & vbCrLf
            ' If OnlyCalcTrace = True Then Tmp_Spe = Tmp_Spe & CStr(Tab_Val_Read(i)) & vbCrLf
            '       Else

            '      End If

            'CALCUL DE LA SOMME DU SPECTRE
            'Tab_Sum_Spe_Trc(J * Nb_Spectres_EDF + (Nb_File_Total)) = _
            'Tab_Sum_Spe_Trc(J * Nb_Spectres_EDF + (Nb_File_Total)) + CDbl(Val_Read)

            ''''  Tab_Sum_Spe_Trc(Num_File_0 + Num_Pixel) = _
            ''''    Tab_Sum_Spe_Trc(Num_File_0 + Num_Pixel) + CDbl(Tab_Val_Read(i))


            Mydiv = Y \ (Nb_Canaux - 1)

            Select Case Mydiv

                Case 0 'PAS FINI DE LIRE UN SPECTRE
                    Y = Y + 1
                Case Else

                    'Nb_File_Total = Nb_File_Total + 1
                    ' If mnuSansPivot(5).Checked = False Then

                    'End If

                    ' DoEvents()

                    Spectrum_Trc(Num_Trc, Num_File_0) = "" '+ Num_Pixel
                    Spectrum_Trc(Num_Trc, Num_File_0) = Spectrum_Trc(Num_Trc, Num_File_0) + Tmp_Spe

                    Y = 0
                    Num_File_0 = Num_File_0 + 1
                    Tmp_Spe = ""
                    ' ts.Close
            End Select


        Next i

        'Close #1

        If Len(Dir_EDF) > 0 Then
            Kill(Dir_EDF & "\" & File_Trc)
        Else
            '''''''         Kill(Dir_Calc1 & "\" & File_Trc)
        End If
    End Sub

    Private Sub Nb_Proc_TextChanged(sender As Object, e As EventArgs) Handles Nb_Proc.TextChanged
        nb_process_custom = True
    End Sub

    Private Sub Text_Lst_Ox_Trc_TextChanged(sender As Object, e As EventArgs) Handles Text_Lst_Ox_Trc.TextChanged

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim Toto As String
        Dim Toto1 As String
        Dim Toto2 As String
        Dim elem_valence As String
        Dim oxyde_valence As String
        Dim i As Integer
        Chemin_Data = "c:\Data"
        Dim Mypath = "c:\data\Test.xlsx"
        Load_atomic_masse_csv()

        Toto = "Sb2O5"
        Toto1 = "Mo3O8"
        Toto2 = "Fe"



        'Dim xlBook = New XLWorkbook("c:\data\Test.xlsx")
        'Dim xlBook1 = New XLWorkbook(Mypath)
        'xlBook = Excel_Open(0, "Test.xlsx")
        Dim StartRow = 3
        Dim StartCol = 3

        ReDim Val_Mat_Area_1(20, 20)
        Dim xlBook1 As XLWorkbook
        Dim xlSheet_gamma As IXLWorksheet

        Dim g As Double
        Dim Filter_Step As Single
        Dim o


        Toto = Strings.Format(0, ".")

        Filter_Step = Single.Parse(TextF_Step.Text, USACulture)
        Filter_Step = Single.Parse("0,2", USACulture)
        Dim valeur As Integer

        valeur = CInt(Math.Round(154251 / 100) * 100)
        valeur = CInt(Math.Round(1005 / 10) * 10)

        ' Read_gamma_csv()
        For i = 0 To 20
            Val_Mat_Area_1(i, 0) = i
        Next
        g = 0.1
        Toto = Strings.Format(0.1, "0.0000")
        Toto = Strings.Format(g, "0.0000")

        o = Strings.Replace("0.2000", ",", ".")
        ReDim Tab_Info_Mat.Z(50)
        ReDim Tab_Info_Mat.Inv(50)
        ReDim Val_Inv_Mtx(50, 50)
        'Read_gamma_xls()
        '  ''''  Insert_Matrix_gamma(0, 0)
        'Dim temp(100) As Integer
        'Dim i As Integer
        'For i = 0 To Nb_Elements_Mat - 1
        '    temp(i) = Val_Mat_Conc_1(0, i)
        'Next i

        'Dim list As New List(Of Array)
        'list.Add(temp)

        'With xlSheet_Conc
        '    ' .Range(
        '    '.Cell(StartRow, StartCol).value = Val_Mat_Conc_1 'InsertTable(IEnumerable(Val_Mat_Conc_1))
        '    Rng = .Range(StartRow, StartCol, (UBound(Val_Mat_Area_1, 1) - LBound(Val_Mat_Area_1, 1) + StartRow), (UBound(Val_Mat_Area_1, 2) - LBound(Val_Mat_Area_1, 2) + StartCol))
        '    'Rng = .Ranges(StartRow, StartCol, (UBound(Val_Mat_Area_1, 1) - LBound(Val_Mat_Area_1, 1) + StartRow), (UBound(Val_Mat_Area_1, 2) - LBound(Val_Mat_Area_1, 2) + StartCol))

        'End With
        'Rng.Value = Val_Mat_Conc_1
        ' Rng.HorizontalAlignment = xlCenter


        ' Renvoi un List(Of Array) a partir d'un datatable
        'Dim retArray As New List(Of Array)
        'For i = 0 To 10
        '    retArray.Add(Val_Mat_Area_1(i, 0))
        'Next
        'Return retArray
        Dim Montest As String
        Dim gamma_z, gamma_f, Mycell_z, Mycell_f
        Dim CellTest
        Dim empty As Boolean
        Dim nb_gamma As Integer
        Dim val_gamma(10, 2) As String
        Dim StrR As String
        Dim Start_filename As Integer
        Dim Nb_filename As Integer
        Dim gamma_filename() As String
        Dim gamma_conc(,) As Integer
        Dim tmp_conc As String


        Montest = "openexcel"

        Select Case Montest
            Case "openexcel"
                'Excel_Open(0, "c:\data\text.xlsx")
                Try
                    xlBook1 = New XLWorkbook("c:\data\test.xlsx")
                    xlSheet_gamma = xlBook1.Worksheet("Gamma")
                Catch ex As Exception
                    Exit Sub
                End Try


                '######### READ Z and Form (Na2O or Na)
                Do
                    Mycell_z = xlSheet_gamma.Cell(1, 2 + nb_gamma)
                    empty = Mycell_z.IsEmpty()


                    If empty = False Then
                        gamma_z = Mycell_z.GetString
                        Mycell_f = xlSheet_gamma.Cell(2, 2 + nb_gamma)
                        'val_gamma(0, i) = gamma_z
                        gamma_f = Mycell_f.GetString 'gamma_f
                        val_gamma(nb_gamma, 0) = gamma_z
                        val_gamma(nb_gamma, 1) = gamma_f

                        nb_gamma += 1
                    End If

                Loop While empty = False

                i = 0

                Do '######### READ Filename / concentration

                    CellTest = xlSheet_gamma.Cell(1 + i, 1)
                    empty = CellTest.IsEmpty()

                    If empty = False Then
                        StrR = CellTest.GetString
                        If StrR = "Filename" Then
                            Start_filename = 1 + i
                            Nb_filename = 0
                        Else
                            Nb_filename += 1
                        End If
                    End If
                    i += 1
                Loop While empty = False
                ReDim gamma_filename(Nb_filename - 1)
                ReDim gamma_conc(Nb_filename - 1, nb_gamma)

                For i = 0 To UBound(gamma_filename)
                    For j = 0 To nb_gamma - 1
                        Mycell_f = xlSheet_gamma.Cell(Start_filename + 1 + i, 2 + j)
                        tmp_conc = Mycell_f.GetString
                        gamma_conc(i, j) = CInt(tmp_conc)
                    Next j

                Next i

                    'val_gamma(0, i) = gamma_z





                'Try
                '   

                '    For j = 0 To 2
                '        For i = 0 To 2
                '            multiList(j).Add(Val_Mat_Area_1(j, i))
                '        Next
                '    Next j
                'Catch ex As Exception
                'End Try



            Case "createexcel"

                Excel_Open(0, "c:\data\text.xlsx")
                Dim multiList As New List(Of List(Of Integer))
                multiList.Add(New List(Of Integer))
                multiList.Add(New List(Of Integer))
                Nb_Process = 1

                For j = 0 To Nb_Process - 1
                    For i = 0 To UBound(Val_Mat_Area_1) - 1
                        multiList(j).Add(Val_Mat_Area_1(j, i))
                    Next
                Next j


                xlSheet_Conc.Row(3).value = multiList


                xlSheet_LOD.Cell(StartRow, StartCol).InsertTable(multiList)
                xlSheet_Conc.Cell(StartRow + 10, StartCol).InsertData(multiList)
                xlSheet_Conc.Cell(1, 2).Value = "Laurent Hello "
                xlSheet_Conc.Cell(1, 1).Value = "B1"
                xlSheet_Conc.Cell(1, 2).Value = "C2 "
                Dim Column As String = xlSheet_Conc.Cell(1, 2).Address.ColumnLetter
                Dim Row As String = xlSheet_Conc.Cell(1, 2).Address.RowNumber
                Dim New_Rng_Q As String = "B1,C4"
                xlSheet_Conc.Ranges(New_Rng_Q).Style.Font.FontColor = XLColor.Red
                Excel_Save(1)



        End Select



    End Sub



    Private Sub Box_txtFiltre_TextChanged(sender As Object, e As EventArgs) Handles Box_txtFiltre.TextChanged 'TextChanged
        If Myinit = True Then
            Select_Par_files = 0
            Par_Mat.Text = ""

            If ComboBox_Type_Calc.Text = "" Then ComboBox_Type_Calc.Text = "Ponctual" ''Init
            If Myinit = True And Ext_Mat <> "" Then
                Maj_Par_Files_Mat()

                If hdf5_mode = True Then
                    Dim pos_star = Strings.InStr(Box_txtFiltre.Text, "*", vbTextCompare)
                    If pos_star > 0 Then
                        Box_txtFiltre.Text = Strings.Replace(Box_txtFiltre.Text, "*", "")
                        Exit Sub
                    End If
                    List_HDF5_group(Chemin_hdf5)
                Else
                    Maj_Files_Mat()
                End If

                List_Par_Files_Trc()
            ElseIf Ext_Mat <> "*.x0" Then
                List_Par_Files_Trc()
                LstPar_Mat.Items.Clear()
                List_Par_Files_Trc()
            End If

            '   List_HDF5_group(Chemin_hdf5)

        End If
    End Sub

    Private Sub ListBox_HDF5_DoubleClick(sender As Object, e As EventArgs) Handles ListBox_HDF5.DoubleClick

        TxtBox_HDF5_File.Text = ListBox_HDF5.SelectedItem
        Chemin_hdf5 = Chemin_Data + "\" + TxtBox_HDF5_File.Text
        List_HDF5_group(Chemin_hdf5)
    End Sub


    Private Sub List_HDF5_group(Hdf5file As String)
        Dim MesFiles(1000) As Object
        Dim MesSpectres(50) As Object
        Dim MesMaps(1000) As Object
        Dim Attrib
        Dim Attrib_ref_obj
        Dim Data
        Dim HDF_as_Map As Boolean
        Dim Myh5
        Dim i As Integer = 0

        'Myh5 = PureHDF.H5File.OpenRead("C:\Data\2023_Data_Euphro\20230322_globals_OBJ_PRJ_IBA.hdf5") 
        Try
            Myh5 = PureHDF.H5File.OpenRead(Hdf5file)
        Catch ex As Exception
            hdf5_mode = False
            Exit Sub
        End Try



        Dim myH5Group1 = Myh5.Group("/")


        '        foreach(var link In group.Children)
        '{
        '    var Message = link switch
        '    {
        '        H5Group group >= $"I am a group and my name is '{group.Name}'.",
        '        H5Dataset dataset           => $"I am a dataset, call me '{dataset.Name}'.",
        '        H5CommitedDatatype datatype >= $"I am the data type '{datatype.Name}'.",
        '        H5UnresolvedLink lostLink   => $"I cannot find my link target =( shame on '{lostLink.Name}'."
        '        _                           => throw New Exception("Unknown link type");
        '    }

        '    Console.WriteLine(Message)
        '}
        LvFiles.Items.Clear()

        Try

            For Each List_Group As H5Group In myH5Group1.Children 'List les " GROUP" présent à la racine

                Try
                    Attrib = List_Group.Attribute("ref object")
                    Attrib_ref_obj = Attrib.ReadString
                    'Attrib_ref_obj.value()

                    If InStr(Attrib_ref_obj(0), Box_txtFiltre.Text, CompareMethod.Text) > 0 Or Box_txtFiltre.Text = "*" Then
                        MesFiles(i) = List_Group
                        LvFiles.Items.Add(MesFiles(i).Name & "_" & Attrib_ref_obj(0))
                        i += 1
                    End If

                Catch ex As Exception
                    hdf5_mode = False
                    TxtBox_HDF5_File.Text = "No AGLAE ponctual group corresponding in this hdf5"
                    Button_Extract.Enabled = False
                End Try
            Next

            If i > 0 Then
                hdf5_mode = True '########### Group type Matrice trouvé
                Button_Extract.Enabled = True
            End If

            Dim MyH5Group
            Try
                MyH5Group = Myh5.Group("/data")
                HDF_as_Map = True
                LvFiles.Items.Clear()
                hdf5_mode = False
                TxtBox_HDF5_File.Text = "No AGLAE ponctual group corresponding in this hdf5"
            Catch ex As Exception

                HDF_as_Map = False
            End Try


            'If HDF_as_Map = True Then
            '    For Each List_Dataset As H5Dataset In MyH5Group.Children 'List les " DATASET" présent dans le group "/data"
            '        MesMaps(i) = List_Dataset
            '        i += 1
            '    Next
            '    Dim TestN = MesMaps(0).Name
            '    'var group = root.Group("/my/nested/group");
            '    '// get dataset in group
            '    Dim DataSet = MyH5Group.Dataset("HE1")
            '    Data = DataSet.Read(Of Integer)

            'Else ' Ponctual DATA

            '    For Each Spectre_name As H5Dataset In MesFiles(0).Children
            '        MesSpectres(i) = Spectre_name.Name
            '    Next

            'End If
        Catch ex As Exception
            TxtBox_HDF5_File.Text = "No AGLAE ponctual group corresponding in this hdf5"
            hdf5_mode = False
        End Try
    End Sub

    Private Sub hdf5_Read_Dataset_Attrib()
        Dim ComputerName As String
        Dim MesFiles(1000) As Object
        Dim Nb_Det_ToRead As Integer
        Dim MesMaps(1000) As Object
        ComputerName = System.Net.Dns.GetHostName
        Dim HDF_as_Map As Boolean
        '        Dim File = New H5File("c:\data\")
        Dim Myh5
        'Myh5 = PureHDF.H5File.OpenRead("C:\Data\2023_Data_Euphro\20230322_globals_NomProjet_IBA.hdf5") 'H5.open("C:\Data\2023_Data_Euphro\20220620_0011_Night_SIBILLA_IBA.hdf5", OpenMode.Binary)
        Myh5 = PureHDF.H5File.OpenRead(Chemin_Data & "\" & TxtBox_HDF5_File.Text)
        Dim myH5Group1 = Myh5.Group("/")
        Dim i As Integer = 0

        Dim Attrib
        Dim Attrib_ref_obj
        Dim Attrib_Tmp
        Dim Data
        Dim Local_Ref_DataSet_ToRead() As String
        Dim MyH5Group
        Dim Som As Long

        Try
            MyH5Group = Myh5.Group("/data")
            HDF_as_Map = True
        Catch ex As Exception
            HDF_as_Map = False
        End Try

        i = 0
        Dim DataSet
        Dim j, k
        j = 0
        i = 0
        k = 0
        '''  Ref_DataSet_ToRead(0) = "x0"
      ''''  Ref_DataSet_ToRead(1) = "x10"
        For Each Det_ToRead As String In Ref_DataSet_ToRead
            If Det_ToRead <> "" Then
                Nb_Det_ToRead += 1
            End If
        Next
        ReDim Local_Ref_DataSet_ToRead(Nb_Det_ToRead - 1)
        For i = 0 To Nb_Det_ToRead - 1
            Local_Ref_DataSet_ToRead(i) = Ref_DataSet_ToRead(i)
        Next

        If HDF_as_Map = True Then
            For Each List_Dataset As H5Dataset In MyH5Group.Children 'List les " DATASET" présent dans le group "/data"
                MesMaps(i) = List_Dataset
                i += 1
            Next
            Dim TestN = MesMaps(0).Name
            'var group = root.Group("/my/nested/group");
            '// get dataset in groupa
            DataSet = MyH5Group.Dataset("HE1")
            Data = DataSet.Read(Of Integer)

        Else ' Ponctual DATA

            ReDim AllSpectres_hdf5(LvFiles.SelectedItems.Count - 1, Nb_Det_ToRead - 1, 2048)
            ReDim Attrib_Spectrum(LvFiles.SelectedItems.Count - 1, Nb_Det_ToRead - 1, 7)
            'For Each Grp As Object In MesFiles



            For j = 0 To LvFiles.SelectedItems.Count - 1

                For Each List_Group As H5Group In myH5Group1.Children
                    Attrib = List_Group.Attribute("ref object")
                    Attrib_ref_obj = Attrib.ReadString

                    If LvFiles.SelectedItems(j).Text = List_Group.Name & "_" & Attrib_ref_obj(0) Then
                        k = 0
                        For Each Det_ToRead As String In Local_Ref_DataSet_ToRead
                            DataSet = List_Group.Dataset(Det_ToRead)
                            Data = DataSet.Read(Of Int32)
                            Attrib = DataSet.Attribute("spectrum sum")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 0) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("acquisition time")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 1) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("seconds since midnight")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 2) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("month")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 3) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("year")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 4) = Attrib_Tmp(0)
                            Attrib = DataSet.Attribute("user comment")
                            Attrib_Tmp = Attrib.ReadString
                            Attrib_Spectrum(j, k, 5) = Attrib_Tmp(0)
                            Attrib_Spectrum(j, k, 6) = Det_ToRead

                            For i = 0 To DataSet.Space.Dimensions(0) - 1
                                Som = Som + Data(i)
                                AllSpectres_hdf5(j, k, i) = CStr(Data(i))
                            Next
                            Attrib_Spectrum(j, k, 7) = CStr(Som)
                            If Som = 0 Then
                                Som = 0
                            End If
                            k += 1
                        Next
                        Exit For
                    End If

                Next
            Next j

        End If

    End Sub

    Private Sub SkipPbMatrixToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SkipPbMatrixToolStripMenuItem.Click
        SkipPbMatrixToolStripMenuItem.Checked = Not (SkipPbMatrixToolStripMenuItem.Checked)
        skip_Pb_mtx = Not (skip_Pb_mtx)
    End Sub

    Private Sub GupixLODNWrite0ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GupixLODNWrite0ToolStripMenuItem.Click

    End Sub

    Private Sub ComboBox_Type_F_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Type_F.SelectedIndexChanged

        Select Case ComboBox_Type_F.Text
            Case "X0"
                TextF_Z.Text = "102"
                TextF_From.Text = "0.3"
                TextF_To.Text = "0.8"
                TextF_Step.Text = "0.1"
            Case "X4"
                TextF_Z.Text = "13"
                TextF_From.Text = "95"
                TextF_To.Text = "105"
                TextF_Step.Text = "1"
            Case "X13"
                TextF_Z.Text = "13"
                TextF_From.Text = "48"
                TextF_To.Text = "52"
                TextF_Step.Text = "1"
            Case "X1"
                TextF_Z.Text = "13"
                TextF_From.Text = "195"
                TextF_To.Text = "205"
                TextF_Step.Text = "1"
            Case "X12"
                TextF_Z.Text = "27"
                TextF_From.Text = "25"
                TextF_To.Text = "30"
                TextF_Step.Text = "0.5"
            Case Else
                TextF_Z.Text = "13"
                TextF_From.Text = "48"
                TextF_To.Text = "52"
                TextF_Step.Text = "1"
        End Select
        Adjust_Filter.Enabled = True
    End Sub

    Private Sub chk_external_ok_CheckedChanged(sender As Object, e As EventArgs) Handles chk_external_ok.CheckedChanged

        If chk_external_ok.Checked = True And path_gamma <> Nothing Then
            Text_gamma.Text = ""
            read_gamma_name_csv()
        Else
            gamma_mode = False
            gamma_ok = False
            nb_gamma = 0

            Text_gamma.Text = "No external elements for this processing"
            ReDim info_gamma_z(1)
            ReDim info_gamma_name(1)
            ReDim ext_tech(1)
        End If

    End Sub

    Private Sub LstPar_Mat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LstPar_Mat.SelectedIndexChanged

    End Sub

    Public Sub read_gamma_name_csv()
        Dim gamma_z, gamma_f
        Dim All_line
        Dim Splitline() As String
        Dim Splitline2() As String
        Dim Splitline3() As String
        ReDim info_gamma_z(100)
        ReDim info_gamma_name(100)
        ReDim ext_tech(100)

        gamma_ok = True
        glob_gamma_mode = True
        nb_gamma = 0
        Try
            All_line = IO.File.ReadAllLines(path_gamma)
            GoTo Okread
        Catch ex As Exception
            gamma_ok = False
            MsgBox("error, close 'external-conc.csv' before click OK")
        End Try

        Try
            All_line = IO.File.ReadAllLines(path_gamma)
        Catch ex As Exception
            gamma_ok = False
            MsgBox("error reading External concentration")
            Exit Sub
        End Try
Okread:
        'For i = 0 To UBound(All_line)
        Splitline = Split(All_line(0), ";")
        Splitline2 = Split(All_line(1), ";")
        Splitline3 = Split(All_line(2), ";")


        If Splitline(0) = "Z" Then
            For i = 0 To UBound(Splitline) - 1
                gamma_z = Splitline(nb_gamma + 1)
                gamma_f = Splitline2(nb_gamma + 1)

                If gamma_z = "" Then Exit For
                info_gamma_z(nb_gamma) = gamma_z
                info_gamma_name(nb_gamma) = gamma_f
                ext_tech(nb_gamma) = Splitline3(nb_gamma + 1)
                nb_gamma += 1
            Next i
        Else
            gamma_ok = False
            '    Exit Sub
        End If

        ReDim Preserve info_gamma_z(nb_gamma)
        ReDim Preserve info_gamma_name(nb_gamma - 1)
        ReDim Preserve ext_tech(nb_gamma - 1)
        'Next



        gamma_mode = True

        If gamma_ok = True Then
            Text_gamma.Text = "External conc. for "

            For i = 0 To nb_gamma - 1
                Text_gamma.Text = Text_gamma.Text & info_gamma_name(i) & ", "
            Next i
        Else
            Text_gamma.Text = Text_gamma.Text & " / No external elements"
        End If
    End Sub


    Sub Load_charge_exp_csv()
        Dim str = ""
        Dim SubItemIndex As Integer
        Dim Ind1 As Integer
        Dim SplitText() As String
        Dim SplitText2() As String
        Dim folder As String = CStr(trvFolders.SelectedNode.Tag)
        Dim pos1 As Integer
        Dim find_Q_in_str As Integer


        If Not folder Is Nothing AndAlso Directory.Exists(folder) Then

            Try
                ''''''''''''''''********************************************************************************************** MATRICE DATA FILES
                ' For Each file As String In Directory.GetFiles(folder, "Traupixe-oxide.ini") 'Get Files In Folder
                For Each file As String In Directory.GetFiles(folder, "charge-exp.csv") 'Get Files In Folder
                    'LvFiles.Items.Add(Path.GetFileNameWithoutExtension(file))
                    SubItemIndex += 1
                Next

            Catch ex As Exception 'Something Went Wrong
                MessageBox.Show(ex.Message)
            Finally
                'If SubItemIndex = 0 Then LvFiles.Items.Add("No files ...")
                If SubItemIndex > 0 Then
                    Ind1 = 0
                    'Text_Lst_Ox_Trc.Text = File.ReadAllText(Chemin_Data & "\Traupixe-oxide.ini") 'str
                    str = File.ReadAllText(Chemin_Data & "\charge-exp.csv") 'str

                    SplitText = Split(str, vbCrLf)
                    SplitText2 = Split(SplitText(0), ";")

                    For i = 0 To UBound(SplitText2)
                        pos1 = Strings.InStr(1, SplitText2(i), "Q", vbTextCompare)
                        If pos1 > 0 Then
                            Det_name_with_charge(find_Q_in_str) = Strings.Trim(Strings.Right(SplitText2(i), Len(SplitText2(i)) - pos1))
                            find_Q_in_str += 1
                        End If
                        ReDim Preserve Det_name_with_charge(find_Q_in_str)
                    Next

                    If find_Q_in_str > 0 Then
                        ReDim Charge_Exp(find_Q_in_str - 1, UBound(SplitText) - 2)
                        For i = 0 To UBound(SplitText) - 2
                            SplitText2 = Split(SplitText(i + 1), ";")
                            For j = 0 To find_Q_in_str - 1
                                Charge_Exp(j, i) = Strings.Replace(SplitText2(j + 1), ",", ".")
                            Next j
                        Next i
                    End If
                End If
            End Try
        End If

    End Sub
    Public Function Det_one_use_charge(nom_det As String) As Boolean
        'MAtrix has chareg value in charge.csv file
        Dim det_trouver As Boolean
        det_trouver = False
        For Each nom In Det_name_with_charge
            If nom_det = nom Then
                det_trouver = True
                Exit For
            End If
        Next
        Return det_trouver
    End Function
    Public Sub Det_use_charge()
        Dim num_column_in_csv_file As Integer

        num_column_in_csv_file = 0
        Pivot_det0.Enabled = True
        Pivot_det1.Enabled = True
        Pivot_det2.Enabled = True
        Pivot_det3.Enabled = True
        Pivot_det4.Enabled = True
        Pivot_det5.Enabled = True
        Pivot_det6.Enabled = True
        Pivot_det7.Enabled = True
        Pivot_det8.Enabled = True

        'MAtrix has chareg value in charge.csv file
        For Each nom In Det_name_with_charge
            If CbDetMat.Text = nom Then
                Use_ext_charge_Mat = True
                num_column_charge_csv_MAT = num_column_in_csv_file
                num_column_in_csv_file += 1
            End If


            If Check_det0.Checked = True Then
                If Check_det0.Text = nom Then
                    Use_ext_charge_Trc(0) = True
                    num_column_charge_csv_TRC(0) = num_column_in_csv_file
                    num_column_in_csv_file += 1
                    Pivot_det0.Text = "Q-File"
                    Pivot_det0.Enabled = False
                End If
            End If

            If Check_det1.Checked = True Then
                If Check_det1.Text = nom Then
                    Use_ext_charge_Trc(1) = True
                    num_column_charge_csv_TRC(1) = num_column_in_csv_file
                    num_column_in_csv_file += 1
                    Pivot_det1.Text = "Q-File"
                    Pivot_det1.Enabled = False

                End If
            End If


            If Check_det2.Checked = True Then
                If Check_det2.Text = nom Then
                    Use_ext_charge_Trc(2) = True
                    num_column_charge_csv_TRC(2) = num_column_in_csv_file
                    num_column_in_csv_file += 1
                    Pivot_det2.Text = "Q-File"
                    Pivot_det2.Enabled = False

                End If
            End If

            If Check_det3.Checked = True Then
                If Check_det3.Text = nom Then
                    Use_ext_charge_Trc(3) = True
                    num_column_charge_csv_TRC(3) = num_column_in_csv_file
                    num_column_in_csv_file += 1
                    Pivot_det3.Text = "Q-File"
                    Pivot_det3.Enabled = False

                End If
            End If

            If Check_det4.Checked = True Then
                If Check_det4.Text = nom Then
                    Use_ext_charge_Trc(4) = True
                    num_column_charge_csv_TRC(4) = num_column_in_csv_file
                    num_column_in_csv_file += 1
                    Pivot_det4.Text = "Q-File"
                    Pivot_det4.Enabled = False

                End If
            End If

            If Check_det5.Checked = True Then
                If Check_det5.Text = nom Then
                    Use_ext_charge_Trc(5) = True
                    num_column_charge_csv_TRC(5) = num_column_in_csv_file
                    num_column_in_csv_file += 1
                    Pivot_det5.Text = "Q-File"
                    Pivot_det5.Enabled = False
                End If
            End If

            If Check_det6.Checked = True Then
                If Check_det6.Text = nom Then
                    Use_ext_charge_Trc(6) = True
                    num_column_charge_csv_TRC(6) = num_column_in_csv_file
                    num_column_in_csv_file += 1
                    Pivot_det6.Text = "Q-File"
                    Pivot_det6.Enabled = False
                End If
            End If

            If Check_det7.Checked = True Then
                If Check_det6.Text = nom Then
                    Use_ext_charge_Trc(7) = True
                    num_column_charge_csv_TRC(7) = num_column_in_csv_file
                    num_column_in_csv_file += 1
                    Pivot_det7.Text = "Q-File"
                    Pivot_det7.Enabled = False
                End If
            End If

            If Check_det8.Checked = True Then
                If Check_det8.Text = nom Then
                    Use_ext_charge_Trc(8) = True
                    num_column_charge_csv_TRC(8) = num_column_in_csv_file
                    num_column_in_csv_file += 1
                    Pivot_det8.Text = "Q-File"
                    Pivot_det8.Enabled = False
                End If
            End If

        Next

    End Sub
    Public Sub Load_atomic_masse_csv()
        Dim Splitline() As String
        Dim All_line

        Try
            All_line = IO.File.ReadAllLines(Environment.CurrentDirectory & "\atomic_mass.csv")
        Catch ex As Exception
            Exit Sub
        End Try


        i = 0
        Try

            Do '######### READ Filename / concentration
                Splitline = Split(All_line(i + 1), ";")
                atomic_info_name(i) = Splitline(0)
                atomic_info_Z(i) = Splitline(1)
                atomic_info_mass(i) = Single.Parse(Splitline(2), USACulture)
                i += 1
            Loop While Splitline(0) <> ""

        Catch ex As Exception

        End Try

    End Sub


    Public Sub Load_gamma_csv()
        Dim Nb_filename_G As Integer
        Dim tmp_conc As String
        Dim Splitline() As String
        Dim All_line
        Dim ind_files_selec As Integer
        Dim Ma, Mo, Va, Vo
        Dim TabFiles_Items() As String
        Dim TabFiles_selected_Items() As String
        Dim G As Integer


        Try
            All_line = IO.File.ReadAllLines(path_gamma)
        Catch ex As Exception
            Exit Sub
        End Try


        ReDim gamma_conc(LvFiles.Items.Count, nb_gamma - 1)
        ReDim gamma_conc_init(LvFiles.Items.Count, nb_gamma - 1)
        ReDim gamma_conc_oxide(LvFiles.Items.Count, nb_gamma - 1)
        ReDim sum_gamma_oxide(LvFiles.Items.Count)
        ReDim sum_gamma_conc(LvFiles.Items.Count)

        i = 0
        Try

            Do '######### READ Filename / concentration
                Splitline = Split(All_line(i + 3), ";")
                gamma_filename(i) = Splitline(0)
                If Splitline(0) <> "" Then
                    Nb_filename_G = i
                    i += 1
                End If

            Loop While Splitline(0) <> ""

        Catch ex As Exception

        End Try

        ReDim Preserve gamma_filename(Nb_filename_G)
        ReDim TabFiles_Items(LvFiles.Items.Count - 1)
        ReDim TabFiles_selected_Items(LvFiles.SelectedItems.Count - 1)


        Dim indx_analyse_in_external As Integer
        Dim indx_analyse_selected As Integer
        Dim nb_gamma_and_selected As Integer

        Dim nb_analyse_selected As Integer
        Dim elem_valence As String
        Dim oxyde_valence As String
        Dim idx_z_in_file_mass As Integer


        ind_files_selec = 0

        nb_gamma_and_selected = 0
        nb_analyse_selected = 0
        Load_atomic_masse_csv() ' Load valeur massa atomique 

        For i = 0 To LvFiles.SelectedItems.Count - 1
            TabFiles_selected_Items(i) = LvFiles.SelectedItems(i).Text
        Next

        For i = 0 To LvFiles.Items.Count - 1
            TabFiles_Items(i) = LvFiles.Items(i).Text
        Next

        ' For i = 0 To LvFiles.Items.Count - 1
        For Each name_file_analyse In TabFiles_Items
            indx_analyse_in_external = Array.IndexOf(gamma_filename, name_file_analyse)
            indx_analyse_selected = Array.IndexOf(TabFiles_selected_Items, name_file_analyse)

            If indx_analyse_in_external <> -1 And indx_analyse_selected <> -1 Then
                tab_gamma_external_value_ok(indx_analyse_selected) = True
                Splitline = Split(All_line(indx_analyse_in_external + 3), ";")

                For j = 0 To nb_gamma - 1

                    Try
                        gamma_conc_init(nb_gamma_and_selected, j) = Splitline(j + 1)
                    Catch ex As Exception
                        gamma_conc_init(nb_gamma_and_selected, j) = "0"
                    End Try

                    tmp_conc = gamma_conc_init(nb_gamma_and_selected, j)


                    Mo = 16 'Masse Oxygen
                    Dim withoxide As Boolean
                    withoxide = False
                    elem_valence = 1
                    oxyde_valence = 0

                    For i = 0 To info_gamma_name(j).Length - 1
                        If IsNumeric(info_gamma_name(j)(i)) Then
                            If info_gamma_name(j)(i - 1) <> "O" Then ' Si un O est trouvé alors on cherche valeurs de Valence
                                elem_valence = info_gamma_name(j)(i)
                            Else
                                oxyde_valence = info_gamma_name(j)(i)
                            End If
                        ElseIf info_gamma_name(j)(i) = "O" Then
                            oxyde_valence = "1"
                        End If
                    Next

                    idx_z_in_file_mass = Array.IndexOf(atomic_info_Z, info_gamma_z(j))
                    Ma = atomic_info_mass(idx_z_in_file_mass) ' Recupère la Mass du Z
                    Va = CInt(elem_valence)
                    Vo = CInt(oxyde_valence)
                    If Vo <> 0 Then
                        withoxide = True

                    Else
                        withoxide = False
                    End If
                    'Recupère Rapport Oxide a enlever car élement en Gamma (Voir Q. Lemasson)
                    ' recupère aussi Rapport oxide -> (Vo * Mo / (Va * Ma + Vo * Mo))
                    tab_rapport_oxide_gamma(j) = calc_gamma_conc_oxide(Va, Ma, Vo, Mo, tmp_conc, nb_gamma_and_selected, j, withoxide)
                    gamma_ok = True
                Next j


                nb_gamma_and_selected += 1
                nb_analyse_selected += 1
            Else
                tab_gamma_external_value_ok(nb_gamma_and_selected) = False

                For j = 0 To nb_gamma - 1
                    gamma_conc_init(nb_analyse_selected, j) = 0
                    gamma_conc(nb_analyse_selected, j) = 0
                    gamma_conc_oxide(nb_analyse_selected, j) = 0
                Next j
                nb_analyse_selected += 1
                nb_gamma_and_selected += 1
            End If


        Next

        If gamma_ok = True Then
            Text_gamma.Text = CStr(Nb_filename_G + 1) & " conc. for "
            glob_gamma_mode = True
            For i = 0 To nb_gamma - 1
                Text_gamma.Text = Text_gamma.Text & info_gamma_name(i) & ", "
            Next i
        Else
            Text_gamma.Text = Text_gamma.Text & " / No external conc."
        End If


    End Sub

    Public Function calc_gamma_conc_oxide(Va As Integer, Ma As Double, Vo As Integer, Mo As Integer, tmp_conc As Integer, i As Integer, j As Integer, withoxide As Boolean)
        Dim rapport_y As Single
        Dim rapport_x As Single

        If withoxide = True Then
            rapport_x = (Va * Ma / (Va * Ma + Vo * Mo))
            rapport_y = (Vo * Mo / (Va * Ma + Vo * Mo))
            gamma_conc_init(i, j) = CInt(tmp_conc)
            gamma_conc(i, j) = CInt(CInt(tmp_conc) * rapport_x) 'Extrait conc Na de Na2O
            gamma_conc_oxide(i, j) = CInt(CInt(tmp_conc) * rapport_y) 'Extrait conc O de Na2O
            sum_gamma_conc(i) += gamma_conc(i, j) ' Sum Gamma Elem without O
            sum_gamma_oxide(i) += gamma_conc_oxide(i, j) ' Sum Gamma O
        Else
            gamma_conc_init(i, j) = CInt(tmp_conc)
            gamma_conc(i, j) = gamma_conc_init(i, j)
            sum_gamma_conc(i) += gamma_conc_init(i, j)
            gamma_conc_oxide(i, j) = 0

        End If

        Return rapport_y
    End Function

    Sub ResetAllTab()

        ReDim Fichier_Matrix(-1) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace(-1)
        ReDim Fichier_Trace0(-1)
        ReDim Fichier_Trace1(-1) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace2(-1) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace3(-1) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace4(-1) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace5(-1) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace6(-1) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace7(-1) 'LvFiles.SelectedItems.Count)
        ReDim Fichier_Trace8(-1) 'LvFiles.SelectedItems.Count)
        ReDim Val_Charge_Trc(-1)

        ReDim Tab_Matrix(-1)
        ReDim Tab_Val_Oxyde_Mat(-1)
        ReDim Tab_Val_Oxyde_Trc(-1, -1)
        ReDim Tab_Val_Trc(-1, -1)
        ReDim Tab_Val_Mat(-1)


        ReDim Info_Oxyde_Mat.nom(-1)
        ReDim Info_Oxyde_Mat.Z(-1)

        ReDim Val_Conc_S_ppm(-1, -1)
        ReDim Val_Conc_S_100(-1, -1)
        ReDim Val_Conc_S_RED_ppm(-1, -1)
        ReDim Val_Conc_S_RED100(-1, -1)
        ReDim Val_Choix_S(-1, -1)
        ReDim Val_YNQ_Final(-1, -1)
        ReDim Val_Error_S(-1, -1)
        ReDim Val_Conc_And_Error(-1, -1)
        ReDim Val_Conc_And_Error100(-1, -1)
        ReDim Val_Mat_Area_1(-1, -1)
        ReDim Val_Mat_Conc_1(-1, -1)
        ReDim Val_Mat_LOD_1(-1, -1)
        ReDim Val_Mat_Height_1(-1, -1)

        'Val_Mat_Stat_Error_1(, )
        ReDim Val_Mat_Fit_Error_1(-1, -1)
        ReDim Val_Mat_Total_Error_1(-1, -1)
        ReDim Val_Mat_Final_Error_1(-1, -1)
        ReDim Val_Mat_Y_N_Q_1(-1, -1)
        ReDim Val_Conc_S_100_1(-1, -1)
        ReDim Val_Conc_S_ppm_1(-1, -1)
        ReDim Val_Conc_S_RED_ppm_1(-1, -1)
        ReDim Val_Conc_S_RED100_1(-1, -1)
        ReDim Val_Error_S_1(-1, -1)
        ReDim Val_Conc_And_Error_1(-1, -1)
        ReDim Val_Conc_And_Error100_1(-1, -1)
        ReDim Val_Choix_S_1(-1, -1)
        ReDim Val_Mat_Mtx_1(-1, -1)

        ReDim Val_Trc_Y_N_Q_1(-1, -1)
        ReDim Val_Trc_Conc_1(-1, -1)
        ReDim Val_Trc_LOD_1(-1, -1)
        ReDim Val_Trc_Area_1(-1, -1)
        ReDim Val_Trc_Fit_Error_1(-1, -1)
        ReDim Val_Trc_WithPivot_Error_1(-1, -1)
        ReDim Val_Trc_Height_1(-1, -1)
        ReDim Val_YNQ_Final_1(-1, -1)
    End Sub

End Class




