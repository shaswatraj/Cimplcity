Imports Microsoft.VisualBasic
Imports GefObjectModel
Imports System.Data.OleDb
Imports System.IO
Imports System.Threading
Imports System.Data
Imports System.Diagnostics

Public Class TagExtract
    Inherits System.Web.UI.Page

    Dim strFileName As String
    Dim filname As String
    Dim screenfile As String
    'Dim xlWorkBook As Excel1.Workbook
    'Dim xlWorkSheet As Excel1.Worksheet
    'Dim xlwb As Excel1.Workbook
    'Dim xlws As Excel1.Worksheet
    Dim drow As VariantType
    Dim baradd As String
    Dim modadd As String
    Dim count As Integer
    Dim tsheet As Integer
    Dim i, k, LOC_CNT, sigcnt As Integer
    Dim r As Integer
    Dim J As Integer
    ' Dim hmitag As Excel1.Range
    Dim strloc As VariantType
    Dim searchrow As Array
    Dim modname As String
    Dim stag As Integer
    Dim STAGPREV As Integer
    Dim FPATH As String
    Dim SNAME As String
    Dim GETAG, TAGDESC, INSTKEY, SIGNAME As String
    Dim CUSTTAG As String
    Dim LRNG As String
    Dim HRNG As String
    Dim STAT As String
    Dim IOPR As String
    Dim EU As String
    Dim SIGTYP As String
    Dim FILTER As String
    Dim MINMA As String
    Dim MAXMA As String
    Dim sigto As String
    Dim sigdev As String
    Dim SIGDEF As String
    Dim rcount As Integer
    Dim lopval As Integer
    Dim dummy As Integer
    Dim devtag As String
    Dim configlbl As String
    Dim configval As String
    Dim tags As String
    Dim scount As Integer
    Dim NROW As Integer
    Dim HMITAGPREV As String
    Dim version As Array
    Dim version1 As String
    Dim version2 As String
    Dim CARD As Array
    Dim CARDP As Integer
    Dim CARDT As Object
    Dim TEMP As String
    Dim TEMP1, TEMP2, TEMP3 As Object
    Dim ROWC As Integer
    Dim ROWCPREV As Integer
    Dim MKVIES As Integer
    Dim oSCR As GefScreen
    Dim oSCRs As GefObjectModel.GefObject
    Dim scrobj As GefObjectModel.GefObjects
    Dim scrobj1 As GefObjectModel.GefObject
    Dim objGef As GefObjectModel.GefObject
    Dim objGef1 As GefObjectModel.GefObject
    Dim objGef2 As GefObjectModel.GefObject
    Dim objGef3 As GefObjectModel.GefObject
    Dim objGef4 As GefObjectModel.GefObject
    Dim obvar As GefObjectModel.GefObjectVariable
    Dim THEAPP As GefObjectModel.GefApplication
    Dim oCimFrmContFmt As GefObjectModel.GefFrameContainerFormat
    Dim H As Integer
    Dim W As Integer
    Dim TEST, sigup As Integer
    Dim TESTB As Integer
    Dim TESTP As Integer
    Dim POSX, POSY, posxx As Integer
    Dim shcount As Integer
    Dim file_cnt, cim, cntlist, chkd As Integer
    Dim cimname(300)
    Dim OBJCNT As Integer
    Dim COUNTNO As Integer
    Dim IOLISTFP, SIMUNIT As String
    Dim IOLISTFN As String
    Dim ERRORLOG As String
    Dim GETAGR As Integer
    Dim GETAGC, TAGDESCC, INSTKEYC, CUSTTAGC, SIGCUSTC As Integer
    Dim GETAGL As String
    Dim TAGVAL As String
    Dim logSCREEN As String
    Dim UPDATESCREEN As String
    Dim LOBJN As String
    Dim dt_TAGID As String
    Dim DUAL, TXMSEL As Integer
    Dim SCREENNAME As String
    Dim LOC, varjobtag As String
    Dim VARTAG, ALMTAG As Object
    Dim VARTAGC, PGT, ROWCG, G, RC As Integer
    Dim yourstring As String
    Dim allcaps As Boolean
    Dim TAGVAL1, TAGVAL2, TAGVAL3, TAGVAL4, TAGVAL5, TAGVAL6, TAGVAL7, TAGVAL8 As String
    Public log As String
    Dim GNAME, GACCESS, GUSEDC, GALM, GALMCLS, GDESC As String
    Dim starttime, stoptime, strtexe As Object
    Dim IOSIG, IODEV, moddev As Object
    'Public tabcontrol1 As New TabControl
    'Public tabpage1 As New TabPage
    Public pnumber As Integer = 0
    Public TAG2, TAG3, TAG4, TAG5, TAG6, TAG7, TAG8 As String
    Public TAG2RL, TAG3RL, TAG4RL, TAG5RL, TAG6RL, TAG7RL, TAG8RL As String
    Dim rowindex As Integer
    Public arr(3) As String
    Public itm As ListViewItem
    Dim ds As New DataSet
    Public Shared dt As DataTable
    Dim dr As DataRow
    Dim idCoulumn As DataColumn
    Dim nameCoulumn As DataColumn
    Dim snameCoulumn, genameCoulumn, almtCoulumn, almvCoulumn, unitCoulumn, dtypCoulumn, unameCoulumn As DataColumn
    Dim TAGARR() As Object
    Dim sender1 As Object
    Dim e1 As EventArgs
    Dim tagshow As String
    Public infoc, warc, errc As Integer
    Dim logck As Integer
    Dim Datagridview1, gridview1 As GridView
    Dim chkAlias As CheckBox
    Public Function screen(CheckBoxList1 As CheckBoxList, gridview1l As GridView, checkbox1 As CheckBox, dtagrid As GridView, spath As String) As GridView
        Try
            chkAlias = checkbox1
            dt = New DataTable()
            idCoulumn = New DataColumn("CUSTTAG NAME", Type.GetType("System.String"))
            nameCoulumn = New DataColumn("SIGNAL NAME", Type.GetType("System.String"))
            snameCoulumn = New DataColumn("SCREEN NAME", Type.GetType("System.String"))
            genameCoulumn = New DataColumn("GE TAGNAME", Type.GetType("System.String"))
            almtCoulumn = New DataColumn("ALARM TYPE", Type.GetType("System.String"))
            almvCoulumn = New DataColumn("ALARM VALUE", Type.GetType("System.String"))
            unitCoulumn = New DataColumn("UNIT", Type.GetType("System.String"))
            dtypCoulumn = New DataColumn("DATA TYPE", Type.GetType("System.String"))
            unameCoulumn = New DataColumn("UNITNAME", Type.GetType("System.String"))
            
            dt.Columns.Add(idCoulumn)
            dt.Columns.Add(nameCoulumn)
            dt.Columns.Add(snameCoulumn)
            dt.Columns.Add(genameCoulumn)
            dt.Columns.Add(almtCoulumn)
            dt.Columns.Add(almvCoulumn)
            dt.Columns.Add(unitCoulumn)
            dt.Columns.Add(dtypCoulumn)
            dt.Columns.Add(unameCoulumn)
            FPATH = spath
            cntlist = CheckBoxList1.Items.Count
            chkd = -1
            Call killprc()
            For t = 0 To cntlist - 1
                If CheckBoxList1.Items(t).Selected = True Then
                    chkd = chkd + 1
                    cimname(chkd) = CheckBoxList1.Items(t).Value
                End If
            Next t
            For x = 0 To chkd
                Try
                    Thread.Sleep(50)
                    strtexe = Now.TimeOfDay
                    THEAPP = CreateObject("CimEdit")

                    'Form2.ProgressBar1.BringToFront()
                    SCREENNAME = cimname(x)

                    screenfile = FPATH & "\" & cimname(x)

                    'If found = False Then

                    oSCR = THEAPP.Open(screenfile, , False)

                    oSCR.Activate()
                    OBJCNT = oSCR.Object.Objects.Count - 1
                Catch ex As Exception

                End Try

                For Me.J = OBJCNT To 0 Step -1
                    Try
                        scrobj1 = oSCR.Object.Objects.Item(J)
                        dummy = 0
                        Call HMISIG(scrobj1)
                    Catch ex As Exception

                    End Try
                Next J
                scrobj = Nothing
                'oSCR.Visible = True
                oSCR.Refresh(True)
                oSCR.Close()
                Call killprc()
                THEAPP = Nothing
            Next
            MsgBox("done")


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Datagridview1
    End Function

    Sub killprc()
        For Each prc As Process In Process.GetProcesses
            'ListBox1.Items.Add(p.ProcessName.ToString)
            If prc.ProcessName = "CimEdit" Then
                prc.Kill()
                Exit For
            End If
        Next
    End Sub
    Sub Insert(tag As String, cname As String, sname As String, getag As String, almtyp As String, almval As String, unit As String, datyp As String, unitname As String)


        Try
            dr = dt.NewRow()
            '    dt1 = ViewState("Customers")
            dt.Rows.Add(tag, cname, sname, getag, almtyp, almval, unit, datyp, unitname)

            'ViewState("Customers") = dt1

            'Me.BindGrid()
            ' Insert1(tag, cname, sname)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Function HMISIG(obj As GefObjectModel.GefObject)

        Dim objcnt1 As Integer
        Dim objs As GefObjectModel.GefObjects
        Dim obj1 As GefObjectModel.GefObject
        Dim JOLYCHARPOS As Integer
        Dim varUnlinkdb As GefObjectModel.GefObjectVariable
        Dim varGETAG As GefObjectModel.GefObjectVariable
        Dim varJOBTAG As GefObjectModel.GefObjectVariable
        Dim varJOBTAG1 As GefObjectModel.GefObjectVariable
        Dim tagIDs As GefObjectModel.GefObjectVariable
        Dim tag, GETAGCK As String
        Dim varTAG1 As GefObjectModel.GefObjectVariable
        Dim varTAG2 As GefObjectModel.GefObjectVariable
        Dim varTAG3 As GefObjectModel.GefObjectVariable
        Dim varTAG4 As GefObjectModel.GefObjectVariable
        'Dim varUnlinkdb As GefObjectVariable
        Dim VarSEN_OPT As GefObjectModel.GefObjectVariable
        Dim dbFound As Boolean
        Dim varname As String
        Dim tempvar As String
        Dim tempstring As String
        Dim getagtemp As String
        Dim custagtemp As String
        Dim convrted As Integer
        Dim MPGETAG As String = ""

        On Error Resume Next
        objcnt1 = 0
        GETAGCK = Nothing
        '  Dim adgvc As Integer = AliasOption.dgvAlias.RowCount
        Dim ac As Integer
        If Not obj.Objects Is Nothing And obj.LinkFormat Is Nothing Then
            objcnt1 = obj.Objects.Count - 1
        Else
            objcnt1 = -1
        End If

        tempvar = ""

        If (objcnt1 = -1) Then

            'Call hmireport

            '        If Not obj.LinkFormat Is Nothing Then
            LOBJN = Nothing
            LOBJN = obj.LinkFormat.LinkObjectName.ToString
            If Not LOBJN Is Nothing Then
                If (LOBJN = "SO_INDIALTR" Or LOBJN = "SO_INDISTAT" Or LOBJN = "SO_INDIAMOS") And TXMSEL = 0 And chkAlias.Checked = False Then

                    If LOBJN = "SO_INDISTAT" Then
                        TAGVAL = obj.Variables.Item("INP_BO_SIGNAL").Value
                        Insert(TAGVAL, obj.Variables.Item("INP_BO_SIGNAL").Value, obj.Screen.Name, obj.Variables.Item("dt_GETAG").Value, "", "", "", "DIGSIG", obj.Variables.Item("UNIT_NAME").Value)
                    ElseIf obj.Variables.Item("LOC_AN_DISPOP").Value = 2 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        Insert(TAGVAL, obj.Variables.Item("INP_BO_SIGNAL").Value, obj.Screen.Name, obj.Variables.Item("dt_GETAG").Value, "", "", "", "DIGSIG", obj.Variables.Item("UNIT_NAME").Value)
                    Else
                        TAGVAL = obj.Variables.Item("TAG1").Value

                        Insert("", obj.Variables.Item("INP_BO_SIGNAL").Value, obj.Screen.Name, obj.Variables.Item("dt_GETAG").Value, "", "", "", "DIGSIG", obj.Variables.Item("UNIT_NAME").Value)
                        
                    End If
                    'col = col + 1
                End If
                If LOBJN = "SO_INDIBFAT" And TXMSEL = 0 And chkAlias.Checked = False Then

                    If obj.Variables.Item("LOC_AN_OPTIONIND_THD").Value = 0 Then
                        TAGVAL = obj.Variables.Item("LOC_AN_TAG1_THD").Value
                        Insert(TAGVAL, obj.Variables.Item("OUT_BO_SIGNAL_THD").Value, obj.Screen.Name, obj.Variables.Item("LOC_AN_DTGETAG_THD").Value, "", "", "", "DIGSIG", obj.Variables.Item("UNIT_NAME").Value)

                    End If
                    If obj.Variables.Item("LOC_BO_FLTPRESENT_THD").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("OUT_BO_SIGNALFLT_THD").Value, obj.Screen.Name, obj.Variables.Item("LOC_AN_DTGETAG_THD").Value, "", "", "", "DIGSIG", obj.Variables.Item("UNIT_NAME").Value)


                    End If
                    If obj.Variables.Item("LOC_AN_OPTIONIND_THD").Value = 1 Then

                        If obj.Variables.Item("LOC_BO_ALARMPRESENT_THD").Value = 1 Then
                            TAGVAL = obj.Variables.Item("LOC_AN_TAG1_THD").Value
                            Insert(TAGVAL, obj.Variables.Item("OUT_BO_ALARMSIG_THD").Value, obj.Screen.Name, obj.Variables.Item("LOC_AN_DTGETAG_THD").Value, "", "", "", "DIGSIG", obj.Variables.Item("UNIT_NAME").Value)
                            'col = col + 1
                        End If
                        If obj.Variables.Item("LOC_BO_TRIPPRESENT_THD").Value = 1 Then
                            TAGVAL = obj.Variables.Item("LOC_AN_TAG1_THD").Value
                            Insert(TAGVAL, obj.Variables.Item("OUT_BO_TRIPSIG_THD").Value, obj.Screen.Name, obj.Variables.Item("LOC_AN_DTGETAG_THD").Value, "", "", "", "DIGSIG", obj.Variables.Item("UNIT_NAME").Value)
                            
                        End If
                    End If
                End If
                If LOBJN = "SO_INDIASAX" And (TXMSEL = 0 Or chkAlias.Checked = True) Then
                    TAGVAL = obj.Variables.Item("TAG1").Value
                    Insert(TAGVAL, obj.Variables.Item("SIGNAL").Value, obj.Screen.Name, obj.Variables.Item("dt_GETAG").Value, "", "", "", "ASIGNAL", obj.Variables.Item("UNIT_NAME").Value)

                    If obj.Variables.Item("SF_AVL").Value = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If chkAlias.Checked = True Then
                            'FltAlias(TAGVAL)
                        End If
                        Insert(TAGVAL, obj.Variables.Item("SF_SIG").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)

                        'col = col + 1
                    End If
                End If
                If LOBJN = "SO_INDIENMN" And TXMSEL = 0 And chkAlias.Checked = False Then
                    TAGVAL = ""
                    Insert(TAGVAL, obj.Variables.Item("SIGNAL").Value, obj.Screen.Name, "", "", "", "", "ENUMSIGNAL", obj.Variables.Item("UNIT_NAME").Value)
                End If
                If LOBJN = "SO_INDILAMP" And TXMSEL = 0 And chkAlias.Checked = False Then
                    TAGVAL = ""
                    Insert(TAGVAL, obj.Variables.Item("SIGNAL").Value, obj.Screen.Name, "", "", "", "", "LAMP", obj.Variables.Item("UNIT_NAME").Value)
                End If
                If LOBJN = "SO_PROCENDR" And TXMSEL = 0 And chkAlias.Checked = False Then
                    TAGVAL = ""
                    Insert(TAGVAL, obj.Variables.Item("SIGNAL").Value, obj.Screen.Name, "", "", "", "", "DOOR", obj.Variables.Item("UNIT_NAME").Value)
                End If
                If LOBJN = "SO_PROCBRK2" And TXMSEL = 0 And chkAlias.Checked = False Then
                    TAGVAL = ""
                    Insert(TAGVAL, obj.Variables.Item("SIGNAL").Value, obj.Screen.Name, "", "", "", "", "BREAKER", obj.Variables.Item("UNIT_NAME").Value)
                End If
                If LOBJN = "SO_PROCDAMP" And TXMSEL = 0 And chkAlias.Checked = False Then
                    TAGVAL = ""
                    If obj.Variables.Item("V1").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("INP_BO_OPENFB").Value, obj.Screen.Name, "", "", "", "", "DAMPEROPENFB", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("V2").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("INP_BO_CLOSEFB").Value, obj.Screen.Name, "", "", "", "", "DAMPERCLSFB", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("V3").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("INP_BO_CMDSIG").Value, obj.Screen.Name, "", "", "", "", "DAMPERCMD", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                End If
                If (LOBJN = "SO_VALVDISC" Or LOBJN = "SO_VALVANAG") And TXMSEL = 0 And chkAlias.Checked = False Then
                    If obj.Variables.Item("DEV_CONTROL_CHK").Value = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        MPGETAG = obj.Variables.Item("INP_ST_GETAG").Value
                        Insert(TAGVAL, obj.Variables.Item("CLS_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        Insert(TAGVAL, obj.Variables.Item("OPN_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        If obj.Variables.Item("COMBO").Value <> 2 Then
                            Insert(TAGVAL, obj.Variables.Item("AUTO_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("COMBO").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("MN_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                    Else
                        TAGVAL = "VALVE"
                        MPGETAG = ""
                    End If
                    If obj.Variables.Item("LOC_BO_VAL4").Value = 1 Then
                        If obj.Variables.Item("LOC_BO_FAILOPEN").Value = 0 Then
                            If LOBJN = "SO_VALVANAG" Then
                                Insert(TAGVAL, obj.Variables.Item("INP_AN_SIGNAL").Value, obj.Screen.Name, MPGETAG, "FC", "", "", "VLVCMD", obj.Variables.Item("UNIT_NAME").Value)
                            Else
                                Insert(TAGVAL, obj.Variables.Item("INP_BO_SIGNAL").Value, obj.Screen.Name, MPGETAG, "FC", "", "", "VLVCMD", obj.Variables.Item("UNIT_NAME").Value)
                            End If
                        ElseIf obj.Variables.Item("LOC_BO_FAILOPEN").Value = 1 Then
                            If LOBJN = "SO_VALVANAG" Then
                                Insert(TAGVAL, obj.Variables.Item("INP_AN_SIGNAL").Value, obj.Screen.Name, MPGETAG, "FO", "", "", "VLVCMD", obj.Variables.Item("UNIT_NAME").Value)
                            Else
                                Insert(TAGVAL, obj.Variables.Item("INP_BO_SIGNAL").Value, obj.Screen.Name, MPGETAG, "FO", "", "", "VLVCMD", obj.Variables.Item("UNIT_NAME").Value)
                            End If
                        End If
                        If obj.Variables.Item("LOC_BO_VAL2").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("INP_BO_CLOSEFB").Value, obj.Screen.Name, MPGETAG, "", "", "", "VLVCLSF", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("LOC_BO_VAL1").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("INP_BO_OPENFB").Value, obj.Screen.Name, MPGETAG, "", "", "", "VLVOPNF", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("LOC_BO_VAL3").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("INP_BO_SIGNAL_FLT").Value, obj.Screen.Name, MPGETAG, "", "", "", "VLVFLT", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                    End If
                End If
                '*************************************************************************
                '    FOR MOTORIZED VALVE OBJECT
                '*************************************************************************
                If LOBJN = "SO_VALVDISC" And TXMSEL = 0 And chkAlias.Checked = False Then
                    If obj.Variables.Item("DEV_CONTROL_CHK").Value = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        MPGETAG = obj.Variables.Item("INP_ST_GETAG").Value
                        Insert(TAGVAL, obj.Variables.Item("FULLCLS_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        Insert(TAGVAL, obj.Variables.Item("FULLOPN_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        Insert(TAGVAL, obj.Variables.Item("AUTO_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        If obj.Variables.Item("COMBO").Value <> 0 Then
                            Insert(TAGVAL, obj.Variables.Item("MN_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                            Insert(TAGVAL, obj.Variables.Item("STEPCLS_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                            Insert(TAGVAL, obj.Variables.Item("STEPOPN_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                    Else
                        TAGVAL = "VALVE"
                        MPGETAG = ""
                    End If

                    If obj.Variables.Item("LOC_BO_FAILOPEN").Value = 0 Then
                        If obj.Variables.Item("LOC_BO_VAL4").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("INP_BO_OPENSIGNAL").Value, obj.Screen.Name, MPGETAG, "FC", "", "", "VLVCMD", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("LOC_BO_VAL5").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("INP_BO_CLOSESIGNAL").Value, obj.Screen.Name, MPGETAG, "FC", "", "", "VLVCMD", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                    ElseIf obj.Variables.Item("LOC_BO_FAILOPEN").Value = 1 Then
                        If obj.Variables.Item("LOC_BO_VAL4").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("INP_BO_OPENSIGNAL").Value, obj.Screen.Name, MPGETAG, "FO", "", "", "VLVCMD", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("LOC_BO_VAL5").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("INP_BO_CLOSESIGNAL").Value, obj.Screen.Name, MPGETAG, "FO", "", "", "VLVCMD", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                    End If
                    If obj.Variables.Item("LOC_BO_VAL2").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("INP_BO_CLOSEFB").Value, obj.Screen.Name, MPGETAG, "", "", "", "VLVCLSF", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("LOC_BO_VAL1").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("INP_BO_OPENFB").Value, obj.Screen.Name, MPGETAG, "", "", "", "VLVOPNF", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("LOC_BO_VAL3").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("INP_BO_SIGNAL_FLT").Value, obj.Screen.Name, MPGETAG, "", "", "", "VLVFLT", obj.Variables.Item("UNIT_NAME").Value)
                    End If

                End If



                '**************************************************************************
                If (LOBJN = "SO_BUTTABS1" Or LOBJN = "SO_BUTTTOGL") And TXMSEL = 0 And chkAlias.Checked = False Then

                    TAGVAL = "CMDBUTTON"
                    Insert(TAGVAL, obj.Variables.Item("CMD_SIGNAL").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                    If obj.Variables.Item("FB_AVAIL").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("FB_SIGNAL").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("ECOPT").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("ECPOINT").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("VISI_AVAIL").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("VISI_SIGNAL").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("VISI_AVAIL1").Value = 1 Then
                        Insert(TAGVAL, obj.Variables.Item("VISI_SIGNAL1").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                End If

                If (LOBJN = "SO_PROCMTAC" Or LOBJN = "SO_PROCFANG" Or LOBJN = "SO_PROCMTDC" Or LOBJN = "SO_PROCPUMP" Or LOBJN = "SO_PROCMTFN" Or LOBJN = "SO_PROCPMPC" Or LOBJN = "SO_PROCMTPP" Or LOBJN = "SO_PROCHEAT") And TXMSEL = 0 And chkAlias.Checked = False Then

                    If obj.Variables.Item("POPUP_OPT").Value = 1 Then
                        If obj.Variables.Item("DEV_CONTROL_CHK").Value = 1 Then
                            TAGVAL = obj.Variables.Item("TAG1").Value
                            MPGETAG = obj.Variables.Item("INP_ST_GETAG").Value
                            Insert(TAGVAL, obj.Variables.Item("SP_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                            Insert(TAGVAL, obj.Variables.Item("ST_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                            If obj.Variables.Item("COMBO").Value <> 2 Then
                                Insert(TAGVAL, obj.Variables.Item("AUTO_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                            End If
                            If obj.Variables.Item("COMBO").Value = 1 Then
                                Insert(TAGVAL, obj.Variables.Item("MN_SIG").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                            End If
                        Else
                            TAGVAL = obj.Variables.Item("AUTO_SIG").Value
                            MPGETAG = ""
                        End If
                        Insert(TAGVAL, obj.Variables.Item("SIGNAL_FBK").Value, obj.Screen.Name, MPGETAG, "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        If obj.Variables.Item("CHK_SIGNAL_FLT").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("SIGNAL_FLT").Value, obj.Screen.Name, MPGETAG, "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("CHK_SIGNAL").Value = 1 Then
                            Insert(TAGVAL, obj.Variables.Item("SIGNAL").Value, obj.Screen.Name, MPGETAG, "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                    End If
                End If
                If LOBJN = "SO_INDIANSP" And TXMSEL = 0 And chkAlias.Checked = False Then
                    
                    Insert(obj.Variables.Item("INP_AN_SIGNAL").Value, obj.Variables.Item("INP_AN_SIGNAL").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    'col = col + 1
                    If obj.Variables.Item("LOC_BO_CHKANIM").Value = 1 And obj.Variables.Item("VALUE1").Value = 1 Then
                        
                        Insert(obj.Variables.Item("LOC_HHSIGNAL").Value, obj.Variables.Item("LOC_HHSIGNAL").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)

                    End If
                    If obj.Variables.Item("LOC_BO_CHKANIM").Value = 1 And obj.Variables.Item("VALUE2").Value = 1 Then
                        
                        Insert(obj.Variables.Item("LOC_HSIGNAL").Value, obj.Variables.Item("LOC_HSIGNAL").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)

                    End If
                    If obj.Variables.Item("LOC_BO_CHKANIM").Value = 1 And obj.Variables.Item("VALUE3").Value = 1 Then
                        
                        Insert(obj.Variables.Item("LOC_LLSIGNAL").Value, obj.Variables.Item("LOC_LLSIGNAL").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("LOC_BO_CHKANIM").Value = 1 And obj.Variables.Item("VALUE4").Value = 1 Then
                        
                        Insert(obj.Variables.Item("LOC_LSIGNAL").Value, obj.Variables.Item("LOC_LSIGNAL").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                End If

                '********** FOR BOOLEAN BOX ************************

                If LOBJN = "SO_INDIBOSP" And TXMSEL = 0 And chkAlias.Checked = False Then
                    Insert(obj.Variables.Item("INP_BO_SIGNAL").Value, obj.Variables.Item("INP_BO_SIGNAL").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                End If

                '*********** FOR PID SIGNALS *********************
                If LOBJN = "SO_PROCPIDO" And TXMSEL = 0 And chkAlias.Checked = False Then
                    If obj.Variables.Item("LOC_BO_SETPOINT").Value = 1 Then
                        Insert(obj.Variables.Item("OUT_AN_SETPOINT").Value, obj.Variables.Item("OUT_AN_SETPOINT").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        
                    End If
                    Insert(obj.Variables.Item("OUT_AN_PROCVAL_FBK").Value, obj.Variables.Item("OUT_AN_PROCVAL_FBK").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    
                    Insert(obj.Variables.Item("OUT_AN_PROCVAL_FBK").Value, obj.Variables.Item("OUT_AN_OUTPUT_FBK").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    

                    If obj.Variables.Item("LOC_BO_SETPOINTCTRL_THD").Value = 1 Then
                        Insert(obj.Variables.Item("INP_BO_LOWER_CMD").Value, obj.Variables.Item("INP_BO_LOWER_CMD").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        
                        Insert(obj.Variables.Item("INP_BO_RAISE_CMD").Value, obj.Variables.Item("INP_BO_RAISE_CMD").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        
                    End If
                    If obj.Variables.Item("LOC_BO_MANUALBO_THD").Value = 1 Then
                        Insert(obj.Variables.Item("INP_BO_AUTOMAN_CMD").Value, obj.Variables.Item("INP_BO_AUTOMAN_CMD").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        

                        'col = col + 1
                        Insert(obj.Variables.Item("INP_BO_CLOSEFINE_CMD").Value, obj.Variables.Item("INP_BO_CLOSEFINE_CMD").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)


                        
                        Insert(obj.Variables.Item("INP_BO_CLOSEFULL_CMD").Value, obj.Variables.Item("INP_BO_CLOSEFULL_CMD").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        
                        Insert(obj.Variables.Item("INP_BO_OPENFINE_CMD").Value, obj.Variables.Item("INP_BO_OPENFINE_CMD").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        
                        Insert(obj.Variables.Item("INP_BO_OPENFULL_CMD").Value, obj.Variables.Item("INP_BO_OPENFULL_CMD").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        
                        Insert(obj.Variables.Item("INP_BO_STOPFULL_CMD").Value, obj.Variables.Item("INP_BO_STOPFULL_CMD").Value, obj.Screen.Name, "", "", "", "", "BUTTON", obj.Variables.Item("UNIT_NAME").Value)
                        
                    End If
                    If obj.Variables.Item("LOC_BO_KPKICTRL_THD").Value = 1 Then


                        Insert(obj.Variables.Item("INP_AN_INTTIME_FBK").Value, obj.Variables.Item("INP_AN_INTTIME_FBK").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        Insert(obj.Variables.Item("INP_AN_PROPGAIN_FBK").Value, obj.Variables.Item("INP_AN_PROPGAIN_FBK").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        
                    End If
                End If
                '******** FOR BAR GRAPH SIGNALS*********************
                If LOBJN = "SO_BARSANAL" And TXMSEL = 0 And chkAlias.Checked = False Then
                    If obj.Variables.Item("LOC_BO_BARTYPESELECT").Value = 0 Then
                        
                        Insert(obj.Variables.Item("OUT_AN_SIGNAL").Value, obj.Variables.Item("OUT_AN_SIGNAL").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        If obj.Variables.Item("OUT_BO_FAULT_AVAIL").Value = 1 Then
                            

                            Insert(obj.Variables.Item("OUT_BO_FAULT").Value, obj.Variables.Item("OUT_BO_FAULT").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("LOC_BO_ALARMUSED").Value = 1 Then
                            
                            Insert(obj.Variables.Item("OUT_AN_ALARM_THD").Value, obj.Variables.Item("OUT_AN_ALARM_THD").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("LOC_BO_TRIPUSED").Value = 1 Then
                            
                            Insert(obj.Variables.Item("OUT_AN_ALARM_THD").Value, obj.Variables.Item("OUT_AN_TRIP_THD").Value, obj.Screen.NAME, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("LOC_BO_LOGICALMAVBL").Value = 1 Then
                           
                            Insert(obj.Variables.Item("OUT_BO_LOGICALMSIGNAL").Value, obj.Variables.Item("OUT_BO_LOGICALMSIGNAL").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                        If obj.Variables.Item("LOC_BO_LOGICTRIPAVBL").Value = 1 Then
                            
                            Insert(obj.Variables.Item("OUT_BO_LOGICTRIPSIGNAL").Value, obj.Variables.Item("OUT_BO_LOGICTRIPSIGNAL").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                    End If
                End If

                '  *****FOR TRANSMITTER OBJECT**************************************
                ' BELOW LOGIC IS GUSEDC FOR FETCHING SIGNALS FROM TRANSMITTER OBJECT
                '********************************************************************
                If LOBJN = "SO_INDIASAT" Then
                    If obj.Variables.Item("SEN_OPT").Value > 0 Then
                        If obj.Variables.Item("signal").Value <> obj.Variables.Item("TAGVALUE_1").Value Then
                            TAGVAL = obj.Variables.Item("TAG1").Value
                            If InStr(TAGVAL, ";") Then
                                searchrow = Split(TAGVAL, ";")
                                TAGVAL = searchrow(1)
                                searchrow = Split(obj.Variables.Item("dt_getag").Value, ";")
                                GETAGCK = searchrow(1)
                            End If

                            Insert(TAGVAL, obj.Variables.Item("signal").Value, obj.Screen.Name, GETAGCK, "SELVAL", "", "", "ASIGNAL", obj.Variables.Item("UNIT_NAME").Value)

                            'Insert(TAGVAL, obj.Variables.Item("signal").Value, obj.Screen.Name)

                            

                        Else
                            Insert(obj.Variables.Item("TAG1").Value, obj.Variables.Item("signal").Value, obj.Screen.Name, obj.Variables.Item("dt_getag").Value, "SELVAL", "", "", "ASIGNAL", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                    End If

                    If obj.Variables.Item("SEN_OPT").Value = 0 Then
                        Dim searchrow As Object
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                            searchrow = Split(obj.Variables.Item("dt_getag").Value, ";")
                            GETAGCK = searchrow(1)
                        End If


                        Insert(TAGVAL, obj.Variables.Item("TAGVALUE_1").Value, obj.Screen.Name, GETAGCK, "ASIGNAL1", "", "", "ASIGNAL", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("SEN_OPT").Value = 1 Then

                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)

                        End If
                        If InStr(obj.Variables.Item("dt_getag").Value, ";") Then
                            searchrow = Split(obj.Variables.Item("dt_getag").Value, ";")
                            GETAGCK = searchrow(1)
                        Else
                            GETAGCK = obj.Variables.Item("dt_getag").Value
                        End If

                        Insert(TAGVAL, obj.Variables.Item("TAGVALUE_1").Value, obj.Screen.Name, GETAGCK, "ASIGNAL1", "", "", "ASIGNAL", obj.Variables.Item("UNIT_NAME").Value)
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(2)

                        End If
                        If InStr(obj.Variables.Item("dt_getag").Value, ";") Then
                            searchrow = Split(obj.Variables.Item("dt_getag").Value, ";")
                            GETAGCK = searchrow(2)
                        Else
                            GETAGCK = obj.Variables.Item("dt_getag").Value
                        End If
                        ' Dim dt As DataTable = DirectCast(ViewState("Customers"), DataTable)
                        
                        Insert(TAGVAL, obj.Variables.Item("TAGVALUE_2").Value, obj.Screen.Name, GETAGCK, "ASIGNAL2", "", "", "ASIGNAL", obj.Variables.Item("UNIT_NAME").Value)
                    End If

                    If obj.Variables.Item("SEN_OPT").Value = 2 Then

                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If
                        If InStr(obj.Variables.Item("dt_getag").Value, ";") Then
                            searchrow = Split(obj.Variables.Item("dt_getag").Value, ";")
                            GETAGCK = searchrow(1)
                        Else
                            GETAGCK = obj.Variables.Item("dt_getag").Value
                        End If
                        Insert(TAGVAL, obj.Variables.Item("TAGVALUE_1").Value, obj.Screen.Name, GETAGCK, "ASIGNAL1", "", "", "ASIGNAL", obj.Variables.Item("UNIT_NAME").Value)
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(2)
                        End If
                        If InStr(obj.Variables.Item("dt_getag").Value, ";") Then
                            searchrow = Split(obj.Variables.Item("dt_getag").Value, ";")
                            GETAGCK = searchrow(2)
                        Else
                            GETAGCK = obj.Variables.Item("dt_getag").Value
                        End If
                        Insert(TAGVAL, obj.Variables.Item("TAGVALUE_2").Value, obj.Screen.Name, GETAGCK, "ASIGNAL2", "", "", "ASIGNAL", obj.Variables.Item("UNIT_NAME").Value)
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(3)
                        End If
                        If InStr(obj.Variables.Item("dt_getag").Value, ";") Then
                            searchrow = Split(obj.Variables.Item("dt_getag").Value, ";")
                            GETAGCK = searchrow(3)
                        Else
                            GETAGCK = obj.Variables.Item("dt_getag").Value
                        End If
                        'Dim dt As DataTable = DirectCast(ViewState("Customers"), DataTable)
                        Insert(TAGVAL, obj.Variables.Item("TAGVALUE_3").Value, obj.Screen.Name, GETAGCK, "ASIGNAL3", "", "", "ASIGNAL", obj.Variables.Item("UNIT_NAME").Value)
                    End If

                    If obj.Variables.Item("SF_TAGVALUE_1").Value <> "" And obj.Variables.Item("SF_TAGVALUE_1").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 Then
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("SF_TAGVALUE_1").Value
                        If obj.Variables.Item("SF_TAGVALUE_2").Value <> "" And obj.Variables.Item("SF_TAGVALUE_2").Value = obj.Variables.Item("signal").Value Then
                            'For J = 1 To UNITNO
                            '    If J = 1 Then
                            TAGVAL = obj.Variables.Item("TAG1").Value

                        Else

                            TAGVAL = obj.Variables.Item("TAG1").Value
                        End If
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If

                        If chkAlias.Checked Then
                            ' FltAlias(TAGVAL)
                            convrted = 0
                        End If
                        Insert(TAGVAL, obj.Variables.Item("SF_TAGVALUE_1").Value, obj.Screen.Name, "", "FAULT1", "", "", "", obj.Variables.Item("UNIT_NAME").Value)

                    End If

                    If obj.Variables.Item("SF_TAGVALUE_2").Value <> "" And obj.Variables.Item("SF_TAGVALUE_2").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 Then
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("SF_TAGVALUE_2").Value
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 12) = obj.Variables.Item("dt_getag").Value
                        'For J = 1 To UNITNO
                        'If J = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        'ElseIf J = 2 Then
                        '    TAGVAL = obj.Variables.Item("TAG2").Value
                        'ElseIf J = 3 Then
                        '    TAGVAL = obj.Variables.Item("TAG3").Value
                        'ElseIf J = 4 Then
                        '    TAGVAL = obj.Variables.Item("TAG4").Value
                        'ElseIf J = 5 Then
                        '    TAGVAL = obj.Variables.Item("TAG5").Value
                        'End If
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(2)
                        End If
                        If chkAlias.Checked Then
                            'FltAlias(TAGVAL)
                            convrted = 0
                        End If

                        Insert(TAGVAL, obj.Variables.Item("SF_TAGVALUE_2").Value, obj.Screen.Name, "", "FAULT2", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If

                    If obj.Variables.Item("SF_TAGVALUE_3").Value <> "" And obj.Variables.Item("SF_TAGVALUE_3").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 And obj.Variables.Item("SEN_OPT").Value = 2 Then
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("SF_TAGVALUE_2").Value
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 12) = obj.Variables.Item("dt_getag").Value
                        'For J = 1 To UNITNO
                        'If J = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        'ElseIf J = 2 Then
                        '    TAGVAL = obj.Variables.Item("TAG2").Value
                        'ElseIf J = 3 Then
                        '    TAGVAL = obj.Variables.Item("TAG3").Value
                        'ElseIf J = 4 Then
                        '    TAGVAL = obj.Variables.Item("TAG4").Value
                        'ElseIf J = 5 Then
                        '    TAGVAL = obj.Variables.Item("TAG5").Value
                        'End If
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(3)
                        End If
                        If chkAlias.Checked Then
                            ' FltAlias(TAGVAL)
                            convrted = 0
                        End If

                        'rowindex = DataGridView1.Rows.Count - 1
                        'DataGridView1.Rows(rowindex).Cells(0).Text = TAGVAL
                        'DataGridView1.Rows(rowindex).Cells(1).Text = obj.Variables.Item("signal").Value
                        'DataGridView1.Rows(rowindex).Cells(2).Text = obj.Screen.Name
                        Insert(TAGVAL, obj.Variables.Item("SF_TAGVALUE_3").Value, obj.Screen.Name, "", "FAULT3", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("SF_SIG").Value <> "" And obj.Variables.Item("SF_SIG").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 Then
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("SF_SIG").Value
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 12) = obj.Variables.Item("dt_getag").Value
                        '   If Not obj.Variables.Item("SF_TAGVALUE_2").Value Like "*FLT*" Then
                        'For J = 1 To UNITNO
                        '    If J = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(0)
                        End If
                        '    ElseIf J = 2 Then
                        '        TAGVAL = obj.Variables.Item("TAG_2").Value
                        '    ElseIf J = 3 Then
                        '        TAGVAL = obj.Variables.Item("TAG_3").Value
                        '    ElseIf J = 4 Then
                        '        TAGVAL = obj.Variables.Item("TAG_4").Value
                        '    ElseIf J = 5 Then
                        '        TAGVAL = obj.Variables.Item("TAG_5").Value
                        '    End If

                        '   Else
                        If chkAlias.Checked Then
                            ' FltAlias(TAGVAL)
                            convrted = 0
                        End If

                        Insert(TAGVAL, obj.Variables.Item("SF_SIG").Value, obj.Screen.Name, "", "", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("LOC_AN_FAILALLSIGNAL").Value <> "" And obj.Variables.Item("LOC_AN_FAILALLSIGNAL").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 Then
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("LOC_AN_FAILALLSIGNAL").Value
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 12) = obj.Variables.Item("dt_getag").Value
                        'For J = 1 To UNITNO
                        '    If J = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(0)
                        End If
                        '    ElseIf J = 2 Then
                        '        TAGVAL = obj.Variables.Item("TAG_2").Value
                        '    ElseIf J = 3 Then
                        '        TAGVAL = obj.Variables.Item("TAG_3").Value
                        '    ElseIf J = 4 Then
                        '        TAGVAL = obj.Variables.Item("TAG_4").Value
                        '    ElseIf J = 5 Then
                        '        TAGVAL = obj.Variables.Item("TAG_5").Value
                        '    End If
                        If chkAlias.Checked Then
                            ' FltAlias(TAGVAL)
                            convrted = 0
                        End If
                        convrted = 0

                        Insert(TAGVAL, obj.Variables.Item("LOC_AN_FAILALLSIGNAL").Value, obj.Screen.Name, "", "FAILALL", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("LOC_BO_SPREADALM").Value <> "" And obj.Variables.Item("LOC_BO_SPREADALM").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 Then
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("LOC_BO_SPREADALM").Value
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 12) = obj.Variables.Item("dt_getag").Value
                        'For J = 1 To UNITNO
                        '    If J = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(0)
                        End If
                        If chkAlias.Checked Then
                            'If convrted = 0 Then
                            '    For ac = 0 To adgvc - 1
                            '        If AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value <> "" Then
                            '            If InStr(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value) And convrted = 0 Then
                            '                TAGVAL = Replace(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value, AliasOption.dgvAlias.Rows(ac).Cells("SPDALM").Value)
                            '                convrted = 1
                            '                Exit For
                            '            End If
                            '        End If
                            '    Next
                            'End If
                            'If InStr(TAGVAL, "I") And TAGVAL.Contains("IT") = False And convrted = 0 Then
                            '    TAGVAL = Replace(TAGVAL, "I", AliasOption.txtaspd.Text)
                            '    convrted = 1
                            'End If
                            'If InStr(TAGVAL, "IT") And convrted = 0 Then
                            '    TAGVAL = Replace(TAGVAL, "IT", AliasOption.txtaspd.Text)
                            '    convrted = 1
                            'End If
                            'If InStr(TAGVAL, "TT") And convrted = 0 Then
                            '    TAGVAL = Replace(TAGVAL, "TT", "T" & AliasOption.txtaspd.Text)
                            '    convrted = 1
                            'End If
                            'If InStr(TAGVAL, "E") And convrted = 0 Then
                            '    TAGVAL = Replace(TAGVAL, "E", AliasOption.txtaspd.Text)
                            '    convrted = 1
                            'End If
                            'If InStr(TAGVAL, "T") And TAGVAL.Contains("TT") = False And TAGVAL.Contains("IT") = False And convrted <> 1 And TAGVAL.Contains("T48") = False And TAGVAL.Contains("T2") = False And TAGVAL.Contains("T3") = False And TAGVAL.Contains("TEPT") = False Then
                            '    TAGVAL = Replace(TAGVAL, "T", AliasOption.txtaspd.Text, 2, 1)
                            'End If
                            convrted = 0
                        End If
                        convrted = 0
                        Insert(TAGVAL, obj.Variables.Item("LOC_BO_SPREADALM").Value, obj.Screen.Name, "", "SPREAD_ALM", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        If chkAlias.Checked = False Then
                            Insert(TAGVAL, obj.Variables.Item("LOC_AN_SPREADTHD").Value, obj.Screen.Name, "", "SPREAD_SP", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                        End If
                    End If
                    If obj.Variables.Item("signal_alm1").Value <> "" And obj.Variables.Item("signal_alm1").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 Then
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("signal_alm1").Value
                        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 12) = obj.Variables.Item("dt_getag").Value
                        'For J = 1 To UNITNO
                        '    If J = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If
                        '    ElseIf J = 2 Then
                        '        TAGVAL = obj.Variables.Item("TAG_2").Value
                        '    ElseIf J = 3 Then
                        '        TAGVAL = obj.Variables.Item("TAG_3").Value
                        '    ElseIf J = 4 Then
                        '        TAGVAL = obj.Variables.Item("TAG_4").Value
                        '    ElseIf J = 5 Then
                        '        TAGVAL = obj.Variables.Item("TAG_5").Value
                        '    End If
                        '   TAGVAL = obj.Variables.Item("dt_jobtagdb").Value
                        '    If chkAlias.Checked Then
                        '        If convrted = 0 Then
                        '            For ac = 0 To adgvc - 1
                        '                If AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value <> "" Then
                        '                    If InStr(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value) And convrted = 0 Then
                        '                        TAGVAL = Replace(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value, AliasOption.dgvAlias.Rows(ac).Cells("HHALM").Value)
                        '                        convrted = 1
                        '                        Exit For
                        '                    End If
                        '                End If
                        '            Next
                        '        End If
                        '        If InStr(TAGVAL, "I") And TAGVAL.Contains("IT") = False And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "I", AliasOption.txtHHAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "IT") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "IT", AliasOption.txtHHAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "TT") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "TT", "T" & AliasOption.txtHHAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "E") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "E", AliasOption.txtHHAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "T") And TAGVAL.Contains("TT") = False And TAGVAL.Contains("IT") = False And convrted <> 1 Then
                        '            TAGVAL = Replace(TAGVAL, "T", AliasOption.txtHHAlm.Text, 2, 1)
                        '        End If

                        '    End If
                        '    convrted = 0

                        Insert(TAGVAL, obj.Variables.Item("signal_alm1").Value, obj.Screen.Name, "", "HHALM", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("signal_alm2").Value <> "" And obj.Variables.Item("signal_alm2").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 Then
                        '    'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("signal_alm2").Value
                        '    'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 12) = obj.Variables.Item("dt_getag").Value
                        '    'For J = 1 To UNITNO
                        '    '    If J = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If
                        '    '    ElseIf J = 2 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_2").Value
                        '    '    ElseIf J = 3 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_3").Value
                        '    '    ElseIf J = 4 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_4").Value
                        '    '    ElseIf J = 5 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_5").Value
                        '    '    End If
                        '    '   TAGVAL = obj.Variables.Item("dt_jobtagdb").Value
                        '    If chkAlias.Checked Then
                        '        If convrted = 0 Then
                        '            For ac = 0 To adgvc - 1
                        '                If AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value <> "" Then
                        '                    If InStr(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value) And convrted = 0 Then
                        '                        TAGVAL = Replace(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value, AliasOption.dgvAlias.Rows(ac).Cells("LLALM").Value)
                        '                        convrted = 1
                        '                        Exit For
                        '                    End If
                        '                End If
                        '            Next
                        '        End If
                        '        If InStr(TAGVAL, "I") And TAGVAL.Contains("IT") = False And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "I", AliasOption.txtLLAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "IT") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "IT", AliasOption.txtLLAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "TT") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "TT", "T" & AliasOption.txtLLAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "E") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "E", AliasOption.txtLLAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "T") And TAGVAL.Contains("TT") = False And TAGVAL.Contains("IT") = False And convrted <> 1 Then
                        '            TAGVAL = Replace(TAGVAL, "T", AliasOption.txtLLAlm.Text, 2, 1)
                        '        End If

                        '    End If
                        '    convrted = 0

                        Insert(TAGVAL, obj.Variables.Item("signal_alm2").Value, obj.Screen.Name, "", "LLALM", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("signal_alm3").Value <> "" And obj.Variables.Item("signal_alm3").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 Then
                        '    'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("signal_alm3").Value
                        '    'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 12) = obj.Variables.Item("dt_getag").Value
                        '    'For J = 1 To UNITNO
                        '    '    If J = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If
                        '    '    ElseIf J = 2 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_2").Value
                        '    '    ElseIf J = 3 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_3").Value
                        '    '    ElseIf J = 4 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_4").Value
                        '    '    ElseIf J = 5 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_5").Value
                        '    '    End If
                        '    '   TAGVAL = obj.Variables.Item("dt_jobtagdb").Value
                        '    If chkAlias.Checked Then
                        '        If convrted = 0 Then
                        '            For ac = 0 To adgvc - 1
                        '                If AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value <> "" Then
                        '                    If InStr(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value) And convrted = 0 Then
                        '                        TAGVAL = Replace(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value, AliasOption.dgvAlias.Rows(ac).Cells("HALM").Value)
                        '                        convrted = 1
                        '                        Exit For
                        '                    End If
                        '                End If
                        '            Next
                        '        End If
                        '        If InStr(TAGVAL, "I") And TAGVAL.Contains("IT") = False And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "I", AliasOption.txtHalm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "IT") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "IT", AliasOption.txtHalm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "TT") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "TT", "T" & AliasOption.txtHalm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "E") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "E", AliasOption.txtHalm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "T") And TAGVAL.Contains("TT") = False And TAGVAL.Contains("IT") = False And convrted <> 1 Then
                        '            TAGVAL = Replace(TAGVAL, "T", AliasOption.txtHalm.Text, 2, 1)
                        '        End If
                        '    End If
                        '    convrted = 0

                        Insert(TAGVAL, obj.Variables.Item("signal_alm3").Value, obj.Screen.Name, "", "HALM", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("signal_alm4").Value <> "" And obj.Variables.Item("signal_alm4").Value <> obj.Variables.Item("signal").Value And TXMSEL = 0 Then
                        '    'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("signal_alm4").Value
                        '    'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 12) = obj.Variables.Item("dt_getag").Value
                        '    'For J = 1 To UNITNO
                        '    '    If J = 1 Then
                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If
                        '    '    ElseIf J = 2 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_2").Value
                        '    '    ElseIf J = 3 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_3").Value
                        '    '    ElseIf J = 4 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_4").Value
                        '    '    ElseIf J = 5 Then
                        '    '        TAGVAL = obj.Variables.Item("TAG_5").Value
                        '    '    End If
                        '    '   TAGVAL = obj.Variables.Item("dt_jobtagdb").Value
                        '    '   If InStr(TAGVAL, "I") And Not InStr(TAGVAL, "IT") Then
                        '    If chkAlias.Checked Then
                        '        If convrted = 0 Then
                        '            For ac = 0 To adgvc - 1
                        '                If AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value <> "" Then
                        '                    If InStr(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value) And convrted = 0 Then
                        '                        TAGVAL = Replace(TAGVAL, AliasOption.dgvAlias.Rows(ac).Cells("SIGNAME").Value, AliasOption.dgvAlias.Rows(ac).Cells("LALM").Value)
                        '                        convrted = 1
                        '                        Exit For
                        '                    End If
                        '                End If
                        '            Next
                        '        End If
                        '        If InStr(TAGVAL, "I") And TAGVAL.Contains("IT") = False And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "I", AliasOption.txtLAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "IT") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "IT", AliasOption.txtLAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "TT") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "TT", "T" & AliasOption.txtLAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "E") And convrted = 0 Then
                        '            TAGVAL = Replace(TAGVAL, "E", AliasOption.txtLAlm.Text)
                        '            convrted = 1
                        '        End If
                        '        If InStr(TAGVAL, "T") And TAGVAL.Contains("TT") = False And TAGVAL.Contains("IT") = False And convrted <> 1 Then
                        '            TAGVAL = Replace(TAGVAL, "T", AliasOption.txtLAlm.Text, 2, 1)
                        '        End If

                        '    End If
                        '    convrted = 0

                        Insert(TAGVAL, obj.Variables.Item("signal_alm4").Value, obj.Screen.Name, "", "LALM", "", "", "", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.LinkFormat.LinkObjectName = "SO_INDIASAT" And TXMSEL = 0 And chkAlias.Checked = False Then
                        '    If obj.Variables.Item("ALM1_LMT").Value <> "" And obj.Variables.Item("ALM1_LMT").Value <> obj.Variables.Item("signal").Value Then
                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("ALM1_LMT").Value


                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If
                        '        'TAGVAL = Replace(TAGVAL, """", "")
                        '        '   TAGVAL = obj.Variables.Item("dt_jobtagdb").Value

                        '        '   ThisWorkbook.Worksheets("alias").Cells(logLOC + col, 3) = obj.Variables.Item("dt_jobtagdb").Value

                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 3) = TAGVAL
                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 6) = "HIHIALM"

                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 11) = obj.Screen.Name

                        '        'col = col + 1
                        Insert(TAGVAL, obj.Variables.Item("ALM1_LMT").Value, obj.Screen.Name, "", "HHSP", "", "", "CONSTANT", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("ALM2_LMT").Value <> "" And obj.Variables.Item("ALM2_LMT").Value <> obj.Variables.Item("signal").Value Then
                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("ALM2_LMT").Value


                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If
                        '        'TAGVAL = Replace(TAGVAL, """", "")
                        '        '   TAGVAL = obj.Variables.Item("dt_jobtagdb").Value

                        '        '   ThisWorkbook.Worksheets("alias").Cells(logLOC + col, 3) = obj.Variables.Item("dt_jobtagdb").Value

                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 3) = TAGVAL
                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 6) = "LOLOALM"

                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 11) = obj.Screen.Name

                        '        'col = col + 1
                        Insert(TAGVAL, obj.Variables.Item("ALM2_LMT").Value, obj.Screen.Name, "", "LLSP", "", "", "CONSTANT", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("ALM3_LMT").Value <> "" And obj.Variables.Item("ALM3_LMT").Value <> obj.Variables.Item("signal").Value Then
                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("ALM3_LMT").Value


                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If
                        '        'TAGVAL = Replace(TAGVAL, """", "")
                        '        '   TAGVAL = obj.Variables.Item("dt_jobtagdb").Value

                        '        '   ThisWorkbook.Worksheets("alias").Cells(logLOC + col, 3) = obj.Variables.Item("dt_jobtagdb").Value

                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 3) = TAGVAL
                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 6) = "HIALM"

                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 11) = obj.Screen.Name

                        '        'col = col + 1
                        Insert(TAGVAL, obj.Variables.Item("ALM3_LMT").Value, obj.Screen.Name, "", "HSP", "", "", "CONSTANT", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                    If obj.Variables.Item("ALM4_LMT").Value <> "" And obj.Variables.Item("ALM4_LMT").Value <> obj.Variables.Item("signal").Value Then
                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 2) = obj.Variables.Item("ALM4_LMT").Value


                        TAGVAL = obj.Variables.Item("TAG1").Value
                        If InStr(TAGVAL, ";") Then
                            searchrow = Split(TAGVAL, ";")
                            TAGVAL = searchrow(1)
                        End If
                        '        'TAGVAL = Replace(TAGVAL, """", "")
                        '        '   TAGVAL = obj.Variables.Item("dt_jobtagdb").Value

                        '        '   ThisWorkbook.Worksheets("alias").Cells(logLOC + col, 3) = obj.Variables.Item("dt_jobtagdb").Value

                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 3) = TAGVAL
                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 6) = "LOALM"

                        '        'ThisWorkbook.Worksheets("alias").Cells(logloc + col, 11) = obj.Screen.Name

                        '        'col = col + 1
                        Insert(TAGVAL, obj.Variables.Item("ALM4_LMT").Value, obj.Screen.Name, "", "LSP", "", "", "CONSTANT", obj.Variables.Item("UNIT_NAME").Value)
                    End If
                End If


            End If
        End If

        Dim xxxx As Integer
        For x = 0 To objcnt1
            If Not obj.Objects Is Nothing Then
                obj1 = obj.Objects.Item(x)
                xxxx = 0
                'If Not obj1.LinkFormat Is Nothing Then
                Call HMISIG(obj1)
                'End If
            End If
        Next x

        '    End If
    End Function

    Public Sub Add(Of T)(ByRef arr As T(), item As T)
        Array.Resize(arr, arr.Length + 1)
        arr(arr.Length - 1) = item
    End Sub

End Class
