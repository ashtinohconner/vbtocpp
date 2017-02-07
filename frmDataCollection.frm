Test
VERSION 5.00
Begin VB.Form frmDataCollection 
   Caption         =   "GTACS NTT - Data Collection Screen"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdDataColl 
      Caption         =   "Begin Data Collection"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Frame fraTypeColl 
      Caption         =   "Collection Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5880
      TabIndex        =   7
      Top             =   2460
      Width           =   1695
      Begin VB.OptionButton optS311 
         Caption         =   "S311"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optExtended 
         Caption         =   "9 Bit Extended"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "9 Bit Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame fraModem 
      Caption         =   "Modem Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1920
      TabIndex        =   3
      Top             =   2460
      Width           =   1695
      Begin VB.CheckBox chkModemC 
         Caption         =   "Modem 3"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkModemB 
         Caption         =   "Modem 2"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkModemA 
         Caption         =   "Modem 1"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCreateDB 
      Caption         =   "Create Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmDataCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Option Explicit

Private cnnDB As ADODB.Connection
Private rstTarget As ADODB.Recordset
Private blnRstOpen As Boolean
Private usbModemA As Arades.UsbPipe
Private usbModemB As Arades.UsbPipe
Private usbModemC As Arades.UsbPipe
Private usbS311 As Arades.UsbPipe
Private blnExit As Boolean
Private blnScan As Boolean
Private blnModemA As Boolean
Private blnModemB As Boolean
Private blnModemC As Boolean
Private blnExtColl As Boolean
Private blnS311Coll As Boolean
Private intUBound As Integer
Private intLBound As Integer
Private intArraySize As Integer
Private intDynArrayPntr As Integer
Private intWrkRegAPntr As Integer
Private intWrkRegBPntr As Integer
Private intWrkRegCPntr As Integer
Private intWrkRegS311Pntr As Integer
Private blnMessModeA As Boolean
Private blnMessModeB As Boolean
Private blnMessModeC As Boolean
Private blnMessModeS311 As Boolean
Private blnMessProb As Boolean
Private lngDBCnt As Long
Private intScnCnt As Integer
Private sglXcoor As Single
Private sglYCoor As Single
Private strModemA As String
Private strModemB As String
Private strModemC As String
Private strS311File As String

Public blnDBOpen As Boolean

Private Sub cmdCreateDB_Click()

    'This procedure checks the PIM connection status, loads the
    'PIM files, obtains site and radiosonde data from the user,
    'and creates the data collection database.

    On Error GoTo errorhandler

    frmDataCollection.blnDBOpen = True
    cmdCreateDB.Enabled = False
    cmdBack.Enabled = False

    Dim usbPIM As Arades.UsbPipe
    Dim strFile As String

    Set usbPIM = New Arades.UsbPipe
    If usbPIM.ConnectionStatus = 1 Then
        MsgBox "A PIM is not connected to this computer.", vbOKOnly
        Set usbPIM = Nothing
        cmdCreateDB.Enabled = True
        cmdBack.Enabled = True
        GoTo subexit
    End If

    strFile = "C:\Program Files\GTACSNTT\GTT_FW_V00.hex"
    usbPIM.LoadFirmware strFile
    
    frmSiteInfo.cboCrystal.Text = "D"
    frmSiteInfo.cboFreq.Text = "8"
    frmSiteInfo.Show vbModal
    
    strFile = "C:\Program Files\GTACSNTT\GTT_FPGA_V00.HEX"
    usbPIM.LoadFPGA strFile
    
    Set usbPIM = Nothing
    
    frmRadiosonde.Show vbModal
    
    Call MakeDB
    
    Call FillDB
    
    cmdDataColl.Enabled = True
    cmdDataColl.SetFocus
    cmdBack.Enabled = True
    GoTo subexit
    
subexit:
    frmDataCollection.blnDBOpen = False
    Exit Sub
    
errorhandler:
    MsgBox "Error Number = " & Err.Number & ", Error Description is " & Err.Description, vbOKOnly
    cmdBack.Enabled = True
    cmdBack.SetFocus
    GoTo subexit

End Sub

Private Sub cmdDataColl_Click()

    'This is the main procedure for the data collection program.  The
    'loop for collecting all data from the PIM is located in this procedure.

    On Error GoTo errorhandler

    cmdDataColl.Enabled = False
    cmdBack.Enabled = False

    Call ConfCheck
    If blnExit = True Then
        GoTo subexit
    End If
    
    Call SWCollConf
    
    If blnModemA = True Then
        Open strModemA For Binary As #1
    End If
    If blnModemB = True Then
        Open strModemB For Binary As #2
    End If
    If blnModemC = True Then
        Open strModemC For Binary As #3
    End If
    If blnS311Coll = True Then
        Open strS311File For Binary As #4
    End If
    
    Dim blnDataColl As Boolean
    Dim blnComplete As Boolean
    Dim bteWrkRegA(15) As Byte
    Dim bteWrkRegB(15) As Byte
    Dim bteWrkRegC() As Byte
    Dim bteWrkRegS311(21) As Byte
    Dim varPIMData As Variant
    Dim bteDynArray() As Byte
    Dim intModemSel As Integer
    Dim lngTargCnt(3) As Long
    Dim lngTargCnt30(3, 29) As Long
    Dim lngFaultCnt(2) As Long
    Dim sglMean(3) As Single
    Dim lngMean(3) As Long
    Dim lngTargCntTot(3) As Long
    Dim intScanSelect As Integer
    Dim varBookmark As Variant
    Dim dblDispTime As Double
    Dim dblCurrTime As Double
    Dim i As Integer
    Dim j As Integer
    blnComplete = False
    blnRstOpen = False
    blnScan = False
    blnMessProb = False
    blnMessModeA = False
    blnMessModeB = False
    blnMessModeC = False
    blnMessModeS311 = False
    intWrkRegAPntr = 0
    intWrkRegBPntr = 0
    intWrkRegCPntr = 0
    intWrkRegS311Pntr = 0
    lngDBCnt = 0
    intScnCnt = 0
    intScanSelect = 0
    dblDispTime = timeGetTime / 1000
    For i = 0 To 3
        For j = 0 To 29
            lngTargCnt30(i, j) = 0
        Next j
        lngTargCnt(i) = 0
        sglMean(i) = 0
        lngMean(i) = 0
        lngTargCntTot(i) = 0
        If i < 3 Then
            lngFaultCnt(i) = 0
        End If
    Next i
    If blnExtColl = True Then
            ReDim bteWrkRegC(31) As Byte
        Else
            ReDim bteWrkRegC(15) As Byte
    End If
    
    Call PIMConnect
    
    blnDataColl = True
    
    Call DBOpen
    
    Call PIMClear(varPIMData)
    If blnExit = True Then
        cmdBack.Enabled = True
        cmdBack.SetFocus
        GoTo subexit
    End If
    
    Call PreparePltFrm
    DoEvents
    
    Do Until blnDataColl = False
        
        If blnModemA = True Then
            intModemSel = 1
            Call MdmProc(varPIMData, bteDynArray(), intModemSel, bteWrkRegA(), bteWrkRegB(), bteWrkRegC(), bteWrkRegS311(), lngTargCnt(), lngFaultCnt())
            If blnExit = True Then
                GoTo subexit
            End If
        End If
        If blnModemB = True Then
            intModemSel = 2
            Call MdmProc(varPIMData, bteDynArray(), intModemSel, bteWrkRegA(), bteWrkRegB(), bteWrkRegC(), bteWrkRegS311(), lngTargCnt(), lngFaultCnt())
            If blnExit = True Then
                GoTo subexit
            End If
        End If
        If blnModemC = True Then
            intModemSel = 3
            Call MdmProc(varPIMData, bteDynArray(), intModemSel, bteWrkRegA(), bteWrkRegB(), bteWrkRegC(), bteWrkRegS311(), lngTargCnt(), lngFaultCnt())
            If blnExit = True Then
                GoTo subexit
            End If
        End If
        If blnS311Coll = True Then
            intModemSel = 4
            Call MdmProc(varPIMData, bteDynArray(), intModemSel, bteWrkRegA(), bteWrkRegB(), bteWrkRegC(), bteWrkRegS311(), lngTargCnt(), lngFaultCnt())
            If blnExit = True Then
                GoTo subexit
            End If
        End If
        
        If blnScan = True Then
            Call NorthDisp(lngTargCnt(), lngTargCnt30(), sglMean(), lngMean(), lngTargCntTot(), intScanSelect, varBookmark)
            blnScan = False
        End If
        dblCurrTime = timeGetTime / 1000
        If (dblCurrTime - dblDispTime) >= 0.5 Then
            DoEvents
            If frmPlot.cmdStop.Enabled = False Then
                blnDataColl = False
                blnComplete = True
            End If
            dblDispTime = dblCurrTime
        End If
        
    Loop
    
    Call TargStat(lngTargCntTot(), lngTargCnt(), lngFaultCnt())

    frmPlot.cmdBack.Enabled = True
    frmPlot.cmdBack.SetFocus

    GoTo subexit
        
subexit:
    If (blnExit = True And frmPlot.Visible = True) Then
        frmDataCollection.Show
        frmPlot.Hide
        cmdBack.Enabled = True
        cmdBack.SetFocus
    End If
    If blnRstOpen = True Then
        rstTarget.Close
        blnRstOpen = False
        Set rstTarget = Nothing
    End If
    If frmDataCollection.blnDBOpen = True Then
        cnnDB.Close
        frmDataCollection.blnDBOpen = False
        Set cnnDB = Nothing
    End If
    If (blnModemA = True And blnDataColl = True) Or (blnModemA = True And blnComplete = True) Then
        If usbModemA.Connected = True Then
            usbModemA.Disconnect
        End If
        Set usbModemA = Nothing
    End If
    If (blnModemB = True And blnDataColl = True) Or (blnModemB = True And blnComplete = True) Then
        If usbModemB.Connected = True Then
            usbModemB.Disconnect
        End If
        Set usbModemB = Nothing
    End If
    If (blnModemC = True And blnDataColl = True) Or (blnModemC = True And blnComplete = True) Then
        If usbModemC.Connected = True Then
            usbModemC.Disconnect
        End If
        Set usbModemC = Nothing
    End If
    If (blnS311Coll = True And blnDataColl = True) Or (blnS311Coll = True And blnComplete = True) Then
        If usbS311.Connected = True Then
            usbS311.Disconnect
        End If
        Set usbS311 = Nothing
    End If
    If blnModemA = True Then
        Close #1
    End If
    If blnModemB = True Then
        Close #2
    End If
    If blnModemC = True Then
        Close #3
    End If
    If blnS311Coll = True Then
        Close #4
    End If
    Exit Sub

errorhandler:
    MsgBox "Error Number = " & Err.Number & ", Error Description is " & Err.Description, vbOKOnly
    frmDataCollection.Show
    frmPlot.Hide
    cmdBack.Enabled = True
    cmdBack.SetFocus
    GoTo subexit

End Sub

Private Sub cmdBack_Click()

    'This procedure takes the user back to the database file
    'naming screen.

    cmdCreateDB.Enabled = True
    cmdDataColl.Enabled = False
    cmdBack.Enabled = True
    chkModemA.Value = 1
    chkModemB.Value = 0
    chkModemC.Value = 0
    optNormal.Value = True
    frmDBFileNaming.Show
    frmDataCollection.Hide
    
End Sub

Private Sub CalcDvalue(sglDvalue() As Single, sglDdelta() As Single)

    'This procedure calculates the D values and stores them in the
    'database.

    Dim sglMalt(7) As Single
    Dim blnSiteFlg As Boolean
    Dim sglSurPres As Single
    Dim sglRise As Single
    Dim sglSlope As Single
    Dim sglYint As Single
    Dim sglAlt1013 As Single
    Dim rstDvalue As ADODB.Recordset
    
    sglMalt(2) = frmRadiosonde.txtAlt500.Text * 39.37 / 12
    sglMalt(3) = frmRadiosonde.txtAlt400.Text * 39.37 / 12
    sglMalt(4) = frmRadiosonde.txtAlt300.Text * 39.37 / 12
    sglMalt(5) = frmRadiosonde.txtAlt250.Text * 39.37 / 12
    sglMalt(6) = frmRadiosonde.txtAlt200.Text * 39.37 / 12
    sglMalt(7) = frmRadiosonde.txtAlt150.Text * 39.37 / 12
    If frmSiteInfo.txtAntHgt.Text = 0 Then
        frmSiteInfo.txtAntHgt.Text = 1
        blnSiteFlg = True
    End If
    sglSurPres = frmSiteInfo.txtPress.Text / Exp(-1 * (frmSiteInfo.txtAntHgt.Text * 12) / (39.37 * 1000 * 7))

        sglRise = frmSiteInfo.txtPress.Text - sglSurPres
        
        sglSlope = sglRise / frmSiteInfo.txtAntHgt.Text
    sglYint = frmSiteInfo.txtPress.Text - (sglSlope * frmSiteInfo.txtAntHgt.Text)
    sglAlt1013 = (1013 - sglYint) / sglSlope
    sglMalt(1) = sglAlt1013
    If frmSiteInfo.txtAntHgt.Text = 1 And blnSiteFlg = True Then
        frmSiteInfo.txtAntHgt.Text = 0
    End If
    sglDvalue(1) = sglMalt(1)
    sglDvalue(2) = sglMalt(2) - 18289
    sglDvalue(3) = sglMalt(3) - 23574
    sglDvalue(4) = sglMalt(4) - 30065
    sglDvalue(5) = sglMalt(5) - 33999
    sglDvalue(6) = sglMalt(6) - 38662
    sglDvalue(7) = sglMalt(7) - 44647
    Set rstDvalue = New ADODB.Recordset
    rstDvalue.ActiveConnection = cnnDB
    rstDvalue.Source = "Radiosonde"
    rstDvalue.Open Options:=adCmdTable, LockType:=adLockOptimistic, CursorType:=adOpenKeyset
    rstDvalue.Move (8)
    rstDvalue.Fields("A2_Value").Value = sglDvalue(1)
    rstDvalue.Update
    rstDvalue.MoveNext
    rstDvalue.Fields("A2_Value").Value = sglDvalue(2)
    rstDvalue.Update
    rstDvalue.MoveNext
    rstDvalue.Fields("A2_Value").Value = sglDvalue(3)
    rstDvalue.Update
    rstDvalue.MoveNext
    rstDvalue.Fields("A2_Value").Value = sglDvalue(4)
    rstDvalue.Update
    rstDvalue.MoveNext
    rstDvalue.Fields("A2_Value").Value = sglDvalue(5)
    rstDvalue.Update
    rstDvalue.MoveNext
    rstDvalue.Fields("A2_Value").Value = sglDvalue(6)
    rstDvalue.Update
    rstDvalue.MoveNext
    rstDvalue.Fields("A2_Value").Value = sglDvalue(7)
    rstDvalue.Update
    rstDvalue.Close
    Set rstDvalue = Nothing
    sglDdelta(1) = sglDvalue(2) - sglDvalue(1)
    sglDdelta(2) = sglDvalue(3) - sglDvalue(2)
    sglDdelta(3) = sglDvalue(4) - sglDvalue(3)
    sglDdelta(4) = sglDvalue(5) - sglDvalue(4)
    sglDdelta(5) = sglDvalue(6) - sglDvalue(5)
    sglDdelta(6) = sglDvalue(7) - sglDvalue(6)

End Sub

Private Sub ConfCheck()

    'This procedure checks that the user has selected a
    'valid data collection cofiguration.

    blnExit = False
    
    If optS311.Value = True Then
        chkModemA.Value = 0
        chkModemB.Value = 0
        chkModemC.Value = 0
    End If
    
    If optS311.Value = False Then
        If (chkModemA.Value = 0 And chkModemB.Value = 0 And chkModemC.Value = 0) Then
            MsgBox "You must select at least one modem to collect data from.", vbOKOnly
            cmdDataColl.Enabled = True
            cmdBack.Enabled = True
            blnExit = True
        End If
    End If
    
    If optExtended.Value = True And (chkModemA.Value = 1 Or chkModemB.Value = 1) Then
        MsgBox "Extended data collection is available only on Modem 3.", vbOKOnly
        cmdDataColl.Enabled = True
        cmdBack.Enabled = True
        blnExit = True
    End If

End Sub

Private Sub DBOpen()

    'This procedure opens the database for the data collection,
    'writes the configuration data to the database, and opens the
    'targets table in the database to accept target data.

    Dim rstSite As ADODB.Recordset
    Dim intModem As Integer

    Set cnnDB = New ADODB.Connection
    cnnDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51; " _
                                & "Data Source= " & frmDBFileNaming.strDB
    cnnDB.Open
    frmDataCollection.blnDBOpen = True
    
    Set rstSite = New ADODB.Recordset
    rstSite.ActiveConnection = cnnDB
    rstSite.Source = "Site"
    rstSite.Open Options:=adCmdTable, LockType:=adLockOptimistic, CursorType:=adOpenKeyset
    rstSite.Move (28)
    If frmDataCollection.optNormal = True Then
            rstSite.Fields("A2_Value").Value = 0
        ElseIf frmDataCollection.optExtended = True Then
            rstSite.Fields("A2_Value").Value = 1
        ElseIf frmDataCollection.optS311 = True Then
            rstSite.Fields("A2_Value").Value = 2
        Else
            rstSite.Fields("A2_Value").Value = -1
    End If
    rstSite.Update
    rstSite.MoveNext
    intModem = 0
    If frmDataCollection.chkModemA.Value = 1 Then
        intModem = intModem + 1
    End If
    If frmDataCollection.chkModemB.Value = 1 Then
        intModem = intModem + 2
    End If
    If frmDataCollection.chkModemC.Value = 1 Then
        intModem = intModem + 4
    End If
    rstSite.Fields("A2_Value").Value = intModem
    rstSite.Update
    rstSite.Close
    Set rstSite = Nothing
    
    Set rstTarget = New ADODB.Recordset
    rstTarget.ActiveConnection = cnnDB
    rstTarget.Source = "Targets"
    rstTarget.Open Options:=adCmdTable, LockType:=adLockOptimistic, CursorType:=adOpenKeyset
    blnRstOpen = True

End Sub

Private Sub FillParity(intModemSel As Integer, bteWrkRegA() As Byte, bteWrkRegB() As Byte, bteWrkRegC() As Byte, bteWrkRegS311() As Byte, bteParity() As Byte, blnEndProc As Boolean)

    'This procedure fills the parity register with target data.

    Dim intMaxRegPos As Integer
    Dim i As Integer
    
    If (blnS311Coll = True And intModemSel = 4) Then
            intMaxRegPos = 21
        ElseIf (intModemSel = 3 And blnExtColl = True) Then
            intMaxRegPos = 31
        Else
            intMaxRegPos = 15
    End If

    Select Case intModemSel
        Case 1
            For i = 0 To intMaxRegPos
                bteParity(i) = bteWrkRegA(i)
            Next i
        Case 2
            For i = 0 To intMaxRegPos
                bteParity(i) = bteWrkRegB(i)
            Next i
        Case 3
            For i = 0 To intMaxRegPos
                bteParity(i) = bteWrkRegC(i)
            Next i
        Case 4
            For i = 0 To intMaxRegPos
                bteParity(i) = bteWrkRegS311(i)
            Next i
        Case Else
            blnEndProc = True
    End Select

End Sub

Private Sub FillWrkReg(bteDynArray() As Byte, intModemSel As Integer, bteWrkRegA() As Byte, bteWrkRegB() As Byte, bteWrkRegC() As Byte, bteWrkRegS311() As Byte, blnEndProc As Boolean)

    'This procedure fills the working register with the target data.

    Dim intLPntr As Integer
    Dim intMaxRegPos As Integer
    Dim i As Integer
    
    If (blnS311Coll = True And intModemSel = 4) Then
            intMaxRegPos = 21
        ElseIf (intModemSel = 3 And blnExtColl = True) Then
            intMaxRegPos = 31
        Else
            intMaxRegPos = 15
    End If
    
    Select Case intModemSel
        Case 1
            If intWrkRegAPntr >= 16 Then
                intWrkRegAPntr = 0
            End If
            intLPntr = intWrkRegAPntr
        Case 2
            If intWrkRegBPntr >= 16 Then
                intWrkRegBPntr = 0
            End If
            intLPntr = intWrkRegBPntr
        Case 3
            If intWrkRegCPntr >= intMaxRegPos + 1 Then
                intWrkRegCPntr = 0
            End If
            intLPntr = intWrkRegCPntr
        Case 4
            If intWrkRegS311Pntr >= intMaxRegPos + 1 Then
                intWrkRegS311Pntr = 0
            End If
            intLPntr = intWrkRegS311Pntr
        Case Else
            blnEndProc = True
            GoTo subexit
    End Select
    
    For i = intLPntr To intMaxRegPos
        Select Case intModemSel
            Case 1
                bteWrkRegA(i) = bteDynArray(intDynArrayPntr - 1)
                intWrkRegAPntr = intWrkRegAPntr + 1
                intDynArrayPntr = intDynArrayPntr + 1
            Case 2
                bteWrkRegB(i) = bteDynArray(intDynArrayPntr - 1)
                intWrkRegBPntr = intWrkRegBPntr + 1
                intDynArrayPntr = intDynArrayPntr + 1
            Case 3
                bteWrkRegC(i) = bteDynArray(intDynArrayPntr - 1)
                intWrkRegCPntr = intWrkRegCPntr + 1
                intDynArrayPntr = intDynArrayPntr + 1
            Case 4
                bteWrkRegS311(i) = bteDynArray(intDynArrayPntr - 1)
                intWrkRegS311Pntr = intWrkRegS311Pntr + 1
                intDynArrayPntr = intDynArrayPntr + 1
            Case Else
                blnEndProc = True
                GoTo subexit
        End Select
        If intDynArrayPntr > intArraySize Then
            If i < intMaxRegPos Then
                    blnEndProc = True
                    GoTo subexit
                Else
                    'out of data but message register is full
            End If
        End If
    Next i

    GoTo subexit
    
subexit:
    Exit Sub

End Sub

Private Sub KFactor(sglKFact As Single)

    'This procedure retrieves the k factor from the database.

    Dim rstKfact As ADODB.Recordset
    Set rstKfact = New ADODB.Recordset
    rstKfact.ActiveConnection = cnnDB
    rstKfact.Source = "Site"
    rstKfact.Open Options:=adCmdTable, LockType:=adLockOptimistic, CursorType:=adOpenKeyset
    rstKfact.Move (5)
    sglKFact = rstKfact.Fields("A2_Value").Value
    rstKfact.Close
    Set rstKfact = Nothing

End Sub

Private Sub MdmProc(varPIMData As Variant, bteDynArray() As Byte, intModemSel As Integer, bteWrkRegA() As Byte, bteWrkRegB() As Byte, bteWrkRegC() As Byte, bteWrkRegS311() As Byte, lngTargCnt() As Long, lngFaultCnt() As Long)

    'This procedure processes the data received from the radar
    'modems.

    Dim blnMessWrdBits() As Boolean
    Dim bteParity() As Byte
    Dim blnWord0 As Boolean
    Dim intParityWrds As Integer
    Dim intParityCnt As Integer
    Dim blnEndProc As Boolean
    Dim blnMessProc As Boolean
    Dim i As Integer
    blnEndProc = False
    If (blnS311Coll = True And intModemSel = 4) Then
            ReDim blnMessWrdBits(21, 7) As Boolean
            ReDim bteParity(21) As Byte
        ElseIf (blnExtColl = True And intModemSel = 3) Then
                ReDim blnMessWrdBits(31, 7) As Boolean
                ReDim bteParity(31) As Byte
        Else
            ReDim blnMessWrdBits(15, 7) As Boolean
            ReDim bteParity(15) As Byte
    End If

    Call PIMGetData(varPIMData, bteDynArray(), intModemSel, blnEndProc)
    If blnExit = True Then
        GoTo subexit
    End If
    If blnEndProc = True Then
        GoTo subexit
    End If
    intDynArrayPntr = 1
    
    Do Until intDynArrayPntr = intArraySize + 1
        
        Call FillWrkReg(bteDynArray(), intModemSel, bteWrkRegA(), bteWrkRegB(), bteWrkRegC(), bteWrkRegS311(), blnEndProc)
        If blnEndProc = True Then
            GoTo subexit
        End If
        
        Call Word0_HdrChk(intModemSel, bteWrkRegA(), bteWrkRegB(), bteWrkRegC(), bteWrkRegS311(), blnWord0, blnEndProc)
        If blnEndProc = True Then
            GoTo subexit
        End If
        
        If blnWord0 = True Then
            Call FillParity(intModemSel, bteWrkRegA(), bteWrkRegB(), bteWrkRegC(), bteWrkRegS311(), bteParity(), blnEndProc)
        End If
        If blnEndProc = True Then
            GoTo subexit
        End If
        
        intParityCnt = 0
        If blnWord0 = True Then
            Call Parity(bteParity(), blnMessWrdBits(), intParityCnt)
        End If
        
        blnMessProc = False
        Call MessPres(blnWord0, blnMessProc, intModemSel, intParityCnt, bteParity(), blnMessWrdBits(), lngTargCnt(), blnEndProc, lngFaultCnt())
        If blnEndProc = True Then
            GoTo subexit
        End If
        
        Call WrkReg(intModemSel, blnMessProc, bteWrkRegA(), bteWrkRegB(), bteWrkRegC(), bteWrkRegS311(), blnEndProc)
        If blnEndProc = True Then
            GoTo subexit
        End If
        
    Loop

    GoTo subexit

subexit:
    Exit Sub

End Sub

Private Sub MessPres(blnWord0 As Boolean, blnMessProc As Boolean, intModemSel As Integer, intParityCnt As Integer, bteParity() As Byte, blnMessWrdBits() As Boolean, lngTargCnt() As Long, blnEndProc As Boolean, lngFaultCnt() As Long)

    'This procedure checks to determine if a message is in the
    'registers containing the modem data from the radar.

    If blnWord0 = True Then
                Select Case intModemSel
                    Case 1
                        If (blnMessModeA = True And intParityCnt >= 7) Then
                                Call MessProc(bteParity(), intParityCnt, blnMessWrdBits(), lngTargCnt(), lngFaultCnt())
                                blnMessProc = True
                                If intParityCnt = 8 Then
                                        blnMessModeA = True
                                    Else
                                        lngTargCnt(3) = lngTargCnt(3) + 1
                                        blnMessModeA = False
                                End If
                            ElseIf (blnMessModeA = False And intParityCnt >= 8) Then
                                Call MessProc(bteParity(), intParityCnt, blnMessWrdBits(), lngTargCnt(), lngFaultCnt())
                                blnMessProc = True
                                blnMessModeA = True
                        End If
                        If (blnMessModeA = True And intParityCnt <= 6) Then
                            blnMessModeA = False
                            lngTargCnt(3) = lngTargCnt(3) + 1
                        End If
                    Case 2
                        If (blnMessModeB = True And intParityCnt >= 7) Then
                                Call MessProc(bteParity(), intParityCnt, blnMessWrdBits(), lngTargCnt(), lngFaultCnt())
                                blnMessProc = True
                                If intParityCnt = 8 Then
                                        blnMessModeB = True
                                    Else
                                        lngTargCnt(3) = lngTargCnt(3) + 1
                                        blnMessModeB = False
                                End If
                            ElseIf (blnMessModeB = False And intParityCnt = 8) Then
                                Call MessProc(bteParity(), intParityCnt, blnMessWrdBits(), lngTargCnt(), lngFaultCnt())
                                blnMessProc = True
                                blnMessModeB = True
                        End If
                        If (blnMessModeB = True And intParityCnt <= 6) Then
                            lngTargCnt(3) = lngTargCnt(3) + 1
                            blnMessModeB = False
                        End If
                    Case 3
                        If (blnMessModeC = True And intParityCnt >= 7) Then
                                Call MessProc(bteParity(), intParityCnt, blnMessWrdBits(), lngTargCnt(), lngFaultCnt())
                                blnMessProc = True
                                If intParityCnt = 8 Then
                                        blnMessModeC = True
                                    Else
                                        lngTargCnt(3) = lngTargCnt(3) + 1
                                        blnMessModeC = False
                                End If
                            ElseIf (blnMessModeC = False And intParityCnt = 8) Then
                                Call MessProc(bteParity(), intParityCnt, blnMessWrdBits(), lngTargCnt(), lngFaultCnt())
                                blnMessProc = True
                                blnMessModeC = True
                        End If
                        If (blnMessModeC = True And intParityCnt <= 6) Then
                            lngTargCnt(3) = lngTargCnt(3) + 1
                            blnMessModeC = False
                        End If
                    Case 4
                        If (blnMessModeS311 = True And intParityCnt >= 15) Then
                                Call MessProc(bteParity(), intParityCnt, blnMessWrdBits(), lngTargCnt(), lngFaultCnt())
                                blnMessProc = True
                                If intParityCnt = 16 Then
                                        blnMessModeS311 = True
                                    Else
                                        lngTargCnt(3) = lngTargCnt(3) + 1
                                        blnMessModeS311 = False
                                End If
                            ElseIf (blnMessModeS311 = False And intParityCnt >= 16) Then
                                Call MessProc(bteParity(), intParityCnt, blnMessWrdBits(), lngTargCnt(), lngFaultCnt())
                                blnMessProc = True
                                blnMessModeS311 = True
                        End If
                        If (blnMessModeS311 = True And intParityCnt <= 14) Then
                            lngTargCnt(3) = lngTargCnt(3) + 1
                            blnMessModeS311 = False
                        End If
                    Case Else
                        blnEndProc = True
                End Select
            Else
                Select Case intModemSel
                    Case 1
                        blnMessModeA = False
                    Case 2
                        blnMessModeB = False
                    Case 3
                        blnMessModeC = False
                    Case 4
                        blnMessModeS311 = False
                    Case Else
                        blnEndProc = True
                End Select
        End If

End Sub

Private Sub MessProc(bteParity() As Byte, intParityCnt As Integer, blnMessWrdBits() As Boolean, lngTargCnt() As Long, lngFaultCnt() As Long)

    'This procedure processes the target data, writes records
    'to the database, and plots the targets on the PPI screen.

    Dim sglRange As Single
    Dim sglAzmth As Single
    Dim intRunLgth As Integer
    Dim lngRdrHgt As Long
    Dim intModeTot(3) As Integer
    Dim intModeCCbit As Integer
    Dim intType As Integer
    Dim strTargType As String
    Dim dblMessTime As Double
    Static dblScanTime As Double
    Dim blnHgtBit(7) As Boolean
    Dim blnHgtGrey(7) As Boolean
    Dim lngAltitude As Long
    Static sglDvalue(7) As Single
    Static sglDdelta(6) As Single
    Dim lngRawBcnAlt As Long
    Dim sglBcnFracCor As Single
    Dim sglBcnCorFact As Single
    Dim lngMCCorHgt As Long
    Dim sglCurvHgt As Single
    Static sglKFact As Single
    Dim sglRdrHgtCor As Single
    Dim sglRdrElAng As Single
    Dim sglBcnHgtCor As Single
    Dim sglBcnElAng As Single
    Dim sglTimStore As Single
    Dim intM4Cnt As Integer
    Dim intEquipFlt As Integer
    Dim intCollFlt As Integer
    Dim intPrimBmAmpl As Integer
    Dim intScnBmAmpl As Integer
    Dim intPrimBmCnt As Integer
    Dim intScnBmCnt As Integer
    Dim intDynTilt As Integer
    Dim intSatDet As Integer
    Dim intHgtMTICnt As Integer
    Dim intLastFreq As Integer
    Dim intFreqMode As Integer
    Dim intMstFreqBmPr As Integer
    Dim intMstFreqBmPrCnt As Integer
    Dim intScnFreqBmPr As Integer
    Dim intScnFreqBmPrCnt As Integer
    Dim intMstFreqBmPrPrimBm As Integer
    Dim intMstFreqBmPrPrimCnt As Integer
    Dim intMstFreqBmPrScnBm As Integer
    Dim intMstFreqBmPrScnCnt As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    If lngDBCnt = 0 Then
        Call CalcDvalue(sglDvalue(), sglDdelta())
        Call KFactor(sglKFact)
    End If
    
    intPrimBmAmpl = -1
    intScnBmAmpl = -1
    intPrimBmCnt = -1
    intScnBmCnt = -1
    intDynTilt = -1
    intSatDet = -1
    intHgtMTICnt = -1
    intLastFreq = -1
    intFreqMode = -1
    intMstFreqBmPr = -1
    intMstFreqBmPrCnt = -1
    intScnFreqBmPr = -1
    intScnFreqBmPrCnt = -1
    intMstFreqBmPrPrimBm = -1
    intMstFreqBmPrPrimCnt = -1
    intMstFreqBmPrScnBm = -1
    intMstFreqBmPrScnCnt = -1
    
    intType = 0
    If blnS311Coll = False Then
            intType = CInt((bteParity(1) And CByte(248)))
            Select Case intType
                Case 184
                    strTargType = "sqc"
                    dblMessTime = timeGetTime / 1000
                    If intScnCnt > 0 Then
                        If (dblMessTime - dblScanTime) < 5 Then
                            GoTo subexit
                        End If
                    End If
                    dblScanTime = dblMessTime
                    intScnCnt = intScnCnt + 1
                    blnScan = True
                Case 96
                    strTargType = "srchstrb"
                Case 32
                    strTargType = "search"
                    lngTargCnt(0) = lngTargCnt(0) + 1
                Case 16
                    strTargType = "beacon"
                    lngTargCnt(1) = lngTargCnt(1) + 1
                Case 48
                    strTargType = "correlated"
                    lngTargCnt(2) = lngTargCnt(2) + 1
                Case 56
                    strTargType = "corrtest"
                    lngTargCnt(2) = lngTargCnt(2) + 1
                Case 40
                    strTargType = "srchtest"
                    lngTargCnt(0) = lngTargCnt(0) + 1
                Case 24
                    strTargType = "bcntest"
                    lngTargCnt(1) = lngTargCnt(1) + 1
                Case Else
                    GoTo subexit
            End Select
        Else
            intType = CInt((bteParity(6) And CByte(64)))
            If intType > 0 Then
                intType = 500
            End If
            intType = intType + CInt((bteParity(5) And CByte(15)))
            Select Case intType
                Case 507
                    strTargType = "sqc"
                    dblMessTime = timeGetTime / 1000
                    If intScnCnt > 0 Then
                        If (dblMessTime - dblScanTime) < 5 Then
                            GoTo subexit
                        End If
                    End If
                    dblScanTime = dblMessTime
                    intScnCnt = intScnCnt + 1
                    blnScan = True
                Case 12
                    strTargType = "srchstrb"
                Case 4
                    strTargType = "search"
                    lngTargCnt(0) = lngTargCnt(0) + 1
                Case 2
                    strTargType = "beacon"
                    lngTargCnt(1) = lngTargCnt(1) + 1
                Case 6
                    strTargType = "correlated"
                    lngTargCnt(2) = lngTargCnt(2) + 1
                Case 7
                    strTargType = "corrtest"
                    lngTargCnt(2) = lngTargCnt(2) + 1
                Case 5
                    strTargType = "srchtest"
                    lngTargCnt(0) = lngTargCnt(0) + 1
                Case 3
                    strTargType = "bcntest"
                    lngTargCnt(1) = lngTargCnt(1) + 1
                Case Else
                    GoTo subexit
            End Select
    End If
    
    sglRange = 0
    If blnS311Coll = False Then
            sglRange = (CInt((bteParity(1) And CByte(1))) * 2048) + _
                        (CInt(bteParity(2)) * 8) + _
                            (CInt((bteParity(3) And CByte(224))) / 32)
            sglRange = sglRange * 0.0625
        Else
            sglRange = CInt((bteParity(8) And CByte(63)) * 256) + _
                CInt(bteParity(9) And CByte(255))
            sglRange = sglRange * 0.015625
    End If
    
    sglAzmth = 0
    If blnS311Coll = False Then
            sglAzmth = (CInt((bteParity(3) And CByte(15))) * 256) + _
                            CInt(bteParity(4))
        Else
            sglAzmth = CInt((bteParity(6) And CByte(63)) * 64) + _
                    CInt((bteParity(7) And CByte(252)) / 4)
    End If
    sglAzmth = sglAzmth * 0.0878906
    
    intRunLgth = 0
    If blnS311Coll = False Then
            intRunLgth = (CInt((bteParity(5) And CByte(254))) / 2)
        Else
            intRunLgth = CInt(bteParity(11) And CByte(127))
    End If
    
    lngRdrHgt = 0
    If blnS311Coll = False Then
            lngRdrHgt = CInt(bteParity(6))
        Else
            If blnMessWrdBits(11, 7) = True Then
                    lngRdrHgt = 1
                Else
                    lngRdrHgt = 0
            End If
            lngRdrHgt = lngRdrHgt + (CInt(bteParity(10) And CByte(127)) * 2)
    End If
    lngRdrHgt = lngRdrHgt * 500
    
    For i = 0 To 3
        intModeTot(i) = 0
    Next i
    
    If blnS311Coll = False Then
            intModeTot(0) = ((CInt(bteParity(8) And CByte(56))) / 8) * 1000
            intModeTot(1) = ((CInt(bteParity(9) And CByte(56))) / 8) * 1000
            intModeTot(2) = ((CInt(bteParity(11) And CByte(56))) / 8) * 1000
            intModeTot(3) = ((CInt(bteParity(13) And CByte(56))) / 8) * 1000
            intModeTot(0) = (((CInt(bteParity(8) And CByte(6))) / 2) * 100) + intModeTot(0)
            intModeTot(1) = ((CInt(bteParity(9) And CByte(7))) * 100) + intModeTot(1)
            intModeTot(2) = ((CInt(bteParity(11) And CByte(7))) * 100) + intModeTot(2)
            intModeTot(3) = ((CInt(bteParity(13) And CByte(7))) * 100) + intModeTot(3)
            intModeTot(1) = (((CInt(bteParity(10) And CByte(224))) / 32) * 10) + intModeTot(1)
            intModeTot(2) = (((CInt(bteParity(12) And CByte(224))) / 32) * 10) + intModeTot(2)
            intModeTot(3) = (((CInt(bteParity(14) And CByte(224))) / 32) * 10) + intModeTot(3)
            intModeCCbit = (((CInt(bteParity(14) And CByte(224))) / 32) * 10)
            intModeTot(1) = ((CInt(bteParity(10) And CByte(28))) / 4) + intModeTot(1)
            intModeTot(2) = ((CInt(bteParity(12) And CByte(28))) / 4) + intModeTot(2)
            intModeTot(3) = ((CInt(bteParity(14) And CByte(24))) / 4) + intModeTot(3)
        Else
            intModeTot(0) = ((CInt(bteParity(12) And CByte(56))) / 8) * 1000
            intModeTot(1) = ((CInt(bteParity(14) And CByte(56))) / 8) * 1000
            intModeTot(2) = ((CInt(bteParity(16) And CByte(56))) / 8) * 1000
            intModeTot(3) = ((CInt(bteParity(18) And CByte(56))) / 8) * 1000
            intModeTot(0) = (((CInt(bteParity(12) And CByte(6))) / 2) * 100) + intModeTot(0)
            intModeTot(1) = ((CInt(bteParity(14) And CByte(7))) * 100) + intModeTot(1)
            intModeTot(2) = ((CInt(bteParity(16) And CByte(7))) * 100) + intModeTot(2)
            intModeTot(3) = ((CInt(bteParity(18) And CByte(7))) * 100) + intModeTot(3)
            intModeTot(1) = (((CInt(bteParity(15) And CByte(224))) / 32) * 10) + intModeTot(1)
            intModeTot(2) = (((CInt(bteParity(17) And CByte(224))) / 32) * 10) + intModeTot(2)
            intModeTot(3) = (((CInt(bteParity(19) And CByte(224))) / 32) * 10) + intModeTot(3)
            intModeCCbit = (((CInt(bteParity(19) And CByte(224))) / 32) * 10)
            intModeTot(1) = ((CInt(bteParity(15) And CByte(28))) / 4) + intModeTot(1)
            intModeTot(2) = ((CInt(bteParity(17) And CByte(28))) / 4) + intModeTot(2)
            intModeTot(3) = ((CInt(bteParity(19) And CByte(24))) / 4) + intModeTot(3)
    End If
    
    If blnS311Coll = False Then
            blnHgtGrey(7) = blnMessWrdBits(14, 3)
            blnHgtGrey(6) = blnMessWrdBits(14, 4)
            blnHgtGrey(5) = blnMessWrdBits(13, 3)
            blnHgtGrey(4) = blnMessWrdBits(13, 4)
            blnHgtGrey(3) = blnMessWrdBits(13, 5)
            blnHgtGrey(2) = blnMessWrdBits(13, 0)
            blnHgtGrey(1) = blnMessWrdBits(13, 1)
            blnHgtGrey(0) = blnMessWrdBits(13, 2)
        Else
            blnHgtGrey(7) = blnMessWrdBits(19, 3)
            blnHgtGrey(6) = blnMessWrdBits(19, 4)
            blnHgtGrey(5) = blnMessWrdBits(18, 3)
            blnHgtGrey(4) = blnMessWrdBits(18, 4)
            blnHgtGrey(3) = blnMessWrdBits(18, 5)
            blnHgtGrey(2) = blnMessWrdBits(18, 0)
            blnHgtGrey(1) = blnMessWrdBits(18, 1)
            blnHgtGrey(0) = blnMessWrdBits(18, 2)
    End If
    
    blnHgtBit(7) = blnHgtGrey(7)
    For i = 6 To 0 Step -1
        blnHgtBit(i) = blnHgtBit(i + 1) Xor blnHgtGrey(i)
    Next i
    lngAltitude = 0
    For i = 0 To 7
        lngAltitude = (Abs(blnHgtBit(i)) * (2 ^ i)) + lngAltitude
    Next i
    lngAltitude = lngAltitude * 5
    lngAltitude = lngAltitude - 10
    Select Case intModeCCbit
        Case 10
            If blnHgtBit(0) = False Then
                    lngAltitude = lngAltitude + 2
                Else
                    lngAltitude = lngAltitude - 2
            End If
        Case 20
        Case 30
            If blnHgtBit(0) = False Then
                    lngAltitude = lngAltitude + 1
                Else
                    lngAltitude = lngAltitude - 1
            End If
        Case 40
            If blnHgtBit(0) = False Then
                    lngAltitude = lngAltitude - 2
                Else
                    lngAltitude = lngAltitude + 2
            End If
        Case 60
            If blnHgtBit(0) = False Then
                    lngAltitude = lngAltitude - 1
                Else
                    lngAltitude = lngAltitude + 1
            End If
        Case Else
            'invalid
    End Select
    
    lngRawBcnAlt = lngAltitude * 100
    If lngRawBcnAlt < 0 Then
        lngRawBcnAlt = 0
        ElseIf lngRawBcnAlt <= 18289 Then
            sglBcnFracCor = lngRawBcnAlt / 18289
            sglBcnCorFact = sglDvalue(1) + (sglDdelta(1) * sglBcnFracCor)
        ElseIf lngRawBcnAlt <= 23574 Then
            sglBcnFracCor = (lngRawBcnAlt - 18289) / (23574 - 18289)
            sglBcnCorFact = sglDvalue(2) + (sglDdelta(2) * sglBcnFracCor)
        ElseIf lngRawBcnAlt <= 30065 Then
            sglBcnFracCor = (lngRawBcnAlt - 23574) / (30065 - 23574)
            sglBcnCorFact = sglDvalue(3) + (sglDdelta(3) * sglBcnFracCor)
        ElseIf lngRawBcnAlt <= 33999 Then
            sglBcnFracCor = (lngRawBcnAlt - 30065) / (33999 - 30065)
            sglBcnCorFact = sglDvalue(4) + (sglDdelta(4) * sglBcnFracCor)
        ElseIf lngRawBcnAlt <= 38662 Then
            sglBcnFracCor = (lngRawBcnAlt - 33999) / (38662 - 33999)
            sglBcnCorFact = sglDvalue(5) + (sglDdelta(5) * sglBcnFracCor)
        Else
            sglBcnFracCor = (lngRawBcnAlt - 38662) / (44647 - 38662)
            sglBcnCorFact = sglDvalue(6) + (sglDdelta(6) * sglBcnFracCor)
    End If
    lngMCCorHgt = CLng((lngRawBcnAlt + sglBcnCorFact) / 100)
    
    If sglRange = 0 Then
        sglRange = 0.01
    End If
    
    sglCurvHgt = (0.88314 / sglKFact) * sglRange ^ 2
    sglRdrHgtCor = (lngRdrHgt - sglCurvHgt - frmSiteInfo.txtAntHgt.Text) / (6076 * sglRange)
    If sglRdrHgtCor = 0 Then
        sglRdrElAng = 0
        ElseIf -sglRdrHgtCor * sglRdrHgtCor + 1 < 0 Then
            sglRdrElAng = -99
            Else
            sglRdrElAng = Atn(sglRdrHgtCor / Sqr(-sglRdrHgtCor * sglRdrHgtCor + 1)) * (180 / 3.14)
    End If
    
    sglBcnHgtCor = ((lngMCCorHgt * 100) - sglCurvHgt - frmSiteInfo.txtAntHgt.Text) / (6076 * sglRange)
    If sglBcnHgtCor = 0 Then
        sglBcnElAng = 0
        ElseIf -sglBcnHgtCor * sglBcnHgtCor + 1 < 0 Then
            sglBcnElAng = -99
            Else
            sglBcnElAng = Atn(sglBcnHgtCor / Sqr(-sglBcnHgtCor * sglBcnHgtCor + 1)) * (180 / 3.14)
    End If
    
    If sglRange = 0.01 Then
        sglRange = 0
    End If
    
    dblMessTime = timeGetTime / 1000
    
    sglXcoor = sglRange * Sin(sglAzmth * 3.14156 / 180)
    sglYCoor = sglRange * Cos(sglAzmth * 3.14156 / 180)
    
    If blnS311Coll = False Then
            intPrimBmAmpl = (CInt(bteParity(7) And CByte(63))) * 32
            If UBound(blnMessWrdBits(), 1) > 15 Then
                    sglTimStore = -1
                Else
                    sglTimStore = (CSng(bteParity(7) And CByte(63))) * 0.1
            End If
        Else
           sglTimStore = (CSng(bteParity(4) And CByte(127))) * 0.1
    End If
    
    If blnS311Coll = False Then
            intM4Cnt = (CInt(bteParity(14) And CByte(3)))
        Else
            intM4Cnt = (CInt(bteParity(13) And CByte(12))) / 4
    End If
    
    If blnS311Coll = False Then
        intPrimBmAmpl = ((CInt(bteParity(8) And CByte(62))) / 2) + intPrimBmAmpl
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intScnBmAmpl = CInt(bteParity(26))
                intScnBmAmpl = ((CInt(bteParity(27) And CByte(224))) * 8) + intScnBmAmpl
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intPrimBmCnt = CInt(bteParity(25) And CByte(31))
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intScnBmCnt = CInt(bteParity(27) And CByte(31))
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intDynTilt = ((CInt(bteParity(15) And CByte(248))) / 8)
                intDynTilt = ((CInt(bteParity(16) And CByte(7))) * 32) + intDynTilt
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intSatDet = CInt(bteParity(15) And CByte(7))
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intHgtMTICnt = ((CInt(bteParity(18) And CByte(192))) / 64)
                intHgtMTICnt = ((CInt(bteParity(19) And CByte(31))) * 4) + intHgtMTICnt
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intLastFreq = ((CInt(bteParity(19) And CByte(224))) / 32)
                intLastFreq = ((CInt(bteParity(20) And CByte(1))) * 8) + intLastFreq
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intFreqMode = ((CInt(bteParity(20) And CByte(14))) / 2)
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intMstFreqBmPr = ((CInt(bteParity(16) And CByte(56))) / 8)
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intMstFreqBmPrCnt = ((CInt(bteParity(16) And CByte(192))) / 64)
                intMstFreqBmPrCnt = ((CInt(bteParity(17) And CByte(15))) * 4) + intMstFreqBmPrCnt
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intScnFreqBmPr = ((CInt(bteParity(17) And CByte(112))) / 16)
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intScnFreqBmPrCnt = ((CInt(bteParity(17) And CByte(128))) / 128)
                intScnFreqBmPrCnt = ((CInt(bteParity(18) And CByte(31))) * 2) + intScnFreqBmPrCnt
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intMstFreqBmPrPrimBm = ((CInt(bteParity(22) And CByte(224))) / 32)
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intMstFreqBmPrPrimCnt = ((CInt(bteParity(21) And CByte(128))) / 128)
                intMstFreqBmPrPrimCnt = ((CInt(bteParity(22) And CByte(31))) * 2) + intMstFreqBmPrPrimCnt
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intMstFreqBmPrScnBm = ((CInt(bteParity(24) And CByte(224))) / 32)
        End If
        
        If UBound(blnMessWrdBits(), 1) > 15 Then
                intMstFreqBmPrScnCnt = ((CInt(bteParity(23) And CByte(128))) / 128)
                intMstFreqBmPrScnCnt = ((CInt(bteParity(24) And CByte(31))) * 2) + intMstFreqBmPrScnCnt
        End If
    End If
    
    intEquipFlt = 0
    If blnS311Coll = False Then
            If blnMessWrdBits(1, 1) = True Then
                intEquipFlt = 1 + intEquipFlt
                lngFaultCnt(0) = lngFaultCnt(0) + 1
            End If
            If blnMessWrdBits(1, 2) = True Then
                intEquipFlt = 2 + intEquipFlt
                lngFaultCnt(1) = lngFaultCnt(1) + 1
            End If
            If blnMessWrdBits(3, 4) = True Then
                intEquipFlt = 4 + intEquipFlt
                lngFaultCnt(2) = lngFaultCnt(2) + 1
            End If
            If UBound(blnMessWrdBits(), 1) > 15 Then
                If blnMessWrdBits(18, 5) = True Then
                    intEquipFlt = 8 + intEquipFlt
                    lngFaultCnt(2) = lngFaultCnt(2) + 1
                End If
            End If
        Else
            If blnMessWrdBits(5, 6) = True Then
                intEquipFlt = 1 + intEquipFlt
                lngFaultCnt(0) = lngFaultCnt(0) + 1
            End If
            If blnMessWrdBits(5, 4) = True Then
                intEquipFlt = 2 + intEquipFlt
                lngFaultCnt(1) = lngFaultCnt(1) + 1
            End If
            If blnMessWrdBits(5, 7) = True Then
                intEquipFlt = 4 + intEquipFlt
                lngFaultCnt(2) = lngFaultCnt(2) + 1
            End If
            If blnMessWrdBits(5, 5) = True Then
                intEquipFlt = 16 + intEquipFlt
            End If
    End If
    
    intCollFlt = 0
    If blnS311Coll = False Then
            If intParityCnt < 8 Then
                intCollFlt = 1 + intCollFlt
            End If
        Else
            If intParityCnt < 16 Then
                intCollFlt = 1 + intCollFlt
            End If
    End If
    
    lngDBCnt = lngDBCnt + 1
    With rstTarget
        .AddNew
        .Fields(0) = lngDBCnt
        Select Case intType
            Case 184, 507
                .Fields(1) = "sqc"
                .Fields(28) = 9999
                .Fields(30) = -99
                .Fields(31) = -99
                .Fields(32) = -99
            Case 96, 12
                .Fields(1) = "srchstrb"
                .Fields(28) = 9999
                .Fields(30) = -99
                .Fields(31) = sglRdrElAng
                .Fields(32) = -99
            Case 32, 4
                .Fields(1) = "search"
                .Fields(28) = 9999
                .Fields(30) = -99
                .Fields(31) = sglRdrElAng
                .Fields(32) = -99
            Case 16, 2
                .Fields(1) = "beacon"
                .Fields(28) = 9999
                .Fields(30) = sglBcnElAng
                .Fields(31) = -99
                .Fields(32) = -99
            Case 48, 6
                .Fields(1) = "correlated"
                .Fields(28) = (lngRdrHgt / 100) - lngMCCorHgt
                .Fields(30) = sglBcnElAng
                .Fields(31) = sglRdrElAng
                .Fields(32) = sglRdrElAng - sglBcnElAng
            Case 40, 5
                .Fields(1) = "srchtest"
                .Fields(28) = 9999
                .Fields(30) = -99
                .Fields(31) = sglRdrElAng
                .Fields(32) = -99
            Case 56, 7
                .Fields(1) = "corrtest"
                .Fields(28) = (lngRdrHgt / 100) - lngMCCorHgt
                .Fields(30) = sglBcnElAng
                .Fields(31) = sglRdrElAng
                .Fields(32) = sglRdrElAng - sglBcnElAng
            Case 24, 3
                .Fields(1) = "bcntest"
                .Fields(28) = 9999
                .Fields(30) = sglBcnElAng
                .Fields(31) = -99
                .Fields(32) = -99
        End Select
        .Fields(2) = sglRange
        .Fields(3) = sglAzmth
        .Fields(4) = intRunLgth
        .Fields(5) = lngRdrHgt / 100
        If blnS311Coll = False Then
                If blnMessWrdBits(5, 0) = True Then
                    .Fields(6) = True
                End If
            Else
                If blnMessWrdBits(8, 6) = True Then
                    .Fields(6) = True
                End If
        End If
        .Fields(7) = intModeTot(3)
        If blnS311Coll = False Then
                If blnMessWrdBits(13, 6) = True Then
                    .Fields(8) = True
                End If
            Else
                If blnMessWrdBits(18, 6) = True Then
                    .Fields(8) = True
                End If
        End If
        .Fields(9) = lngAltitude
        .Fields(10) = lngMCCorHgt
        .Fields(11) = intModeTot(2)
        If blnS311Coll = False Then
                If blnMessWrdBits(12, 1) = True Then
                    .Fields(12) = True
                End If
            Else
                If blnMessWrdBits(17, 1) = True Then
                    .Fields(12) = True
                End If
        End If
        If blnS311Coll = False Then
                If blnMessWrdBits(11, 6) = True Then
                    .Fields(13) = True
                End If
            Else
                If blnMessWrdBits(16, 6) = True Then
                    .Fields(13) = True
                End If
        End If
        .Fields(14) = intModeTot(0)
        If blnS311Coll = False Then
                If blnMessWrdBits(8, 0) = True Then
                    .Fields(15) = True
                End If
            Else
                If blnMessWrdBits(13, 1) = True Then
                    .Fields(15) = True
                End If
        End If
        If blnS311Coll = False Then
                If blnMessWrdBits(8, 6) = True Then
                    .Fields(16) = True
                End If
            Else
                If blnMessWrdBits(12, 6) = True Then
                    .Fields(16) = True
                End If
        End If
        .Fields(17) = intModeTot(1)
        If blnS311Coll = False Then
                If blnMessWrdBits(10, 1) = True Then
                    .Fields(18) = True
                End If
            Else
                If blnMessWrdBits(15, 1) = True Then
                    .Fields(18) = True
                End If
        End If
        If blnS311Coll = False Then
                If blnMessWrdBits(9, 6) = True Then
                    .Fields(19) = True
                End If
            Else
                If blnMessWrdBits(14, 6) = True Then
                    .Fields(19) = True
                End If
        End If
        If blnS311Coll = False Then
                If blnMessWrdBits(10, 0) = True Then
                    .Fields(20) = True
                End If
            Else
                If blnMessWrdBits(19, 0) = True Then
                    .Fields(20) = True
                End If
        End If
        If blnS311Coll = False Then
                If blnMessWrdBits(7, 7) = True Then
                    .Fields(21) = True
                End If
            Else
                If intModeTot(2) = 7500 Then
                    .Fields(21) = True
                End If
        End If
        If blnS311Coll = False Then
                If blnMessWrdBits(7, 6) = True Then
                    .Fields(22) = True
                End If
            Else
                If intModeTot(2) = 7600 Then
                    .Fields(22) = True
                End If
        End If
        If intModeTot(2) = 7700 Then
            .Fields(23) = True
        End If
        If blnS311Coll = False Then
                If blnMessWrdBits(1, 3) = True Then
                    .Fields(27) = True
                End If
            Else
                If blnMessWrdBits(5, 0) = True Then
                    .Fields(27) = True
                End If
        End If
        .Fields(29) = -99
        .Fields(33) = sglXcoor
        .Fields(34) = sglYCoor
        .Fields(35) = intScnCnt
        .Fields(36) = 0
        .Fields(37) = dblMessTime
        .Fields(38) = intEquipFlt
        .Fields(39) = sglTimStore
        .Fields(40) = intM4Cnt
        If blnS311Coll = False Then
                If blnMessWrdBits(14, 2) = True Then
                    .Fields(41) = True
                End If
            Else
                If blnMessWrdBits(13, 4) = True Then
                    .Fields(41) = True
                End If
        End If
        .Fields(42) = intPrimBmAmpl
        .Fields(43) = intScnBmAmpl
        .Fields(44) = intPrimBmCnt
        .Fields(45) = intScnBmCnt
        .Fields(46) = intDynTilt
        .Fields(47) = intSatDet
        .Fields(48) = intHgtMTICnt
        .Fields(49) = intLastFreq
        .Fields(50) = intFreqMode
        .Fields(51) = intMstFreqBmPr
        .Fields(52) = intMstFreqBmPrCnt
        .Fields(53) = intScnFreqBmPr
        .Fields(54) = intScnFreqBmPrCnt
        .Fields(55) = intMstFreqBmPrPrimBm
        .Fields(56) = intMstFreqBmPrPrimCnt
        .Fields(57) = intMstFreqBmPrScnBm
        .Fields(58) = intMstFreqBmPrScnCnt
        If blnS311Coll = False Then
                If blnMessWrdBits(12, 0) = True Then
                    .Fields(59) = True
                End If
            Else
                If blnMessWrdBits(19, 1) = True Then
                    .Fields(59) = True
                End If
        End If
        .Fields(60) = intCollFlt
        .Update
    End With
    
    If (intType = 184 Or intType = 507) Then
            frmPlot.FillColor = RGB(255, 255, 0)
            frmPlot.Circle (sglXcoor, sglYCoor), 0.5, RGB(255, 255, 0)
        ElseIf (intType = 32 Or intType = 40 Or intType = 4 Or intType = 5) Then
            frmPlot.FillColor = RGB(0, 0, 255)
            frmPlot.Circle (sglXcoor, sglYCoor), 0.5, RGB(0, 0, 255)
        ElseIf (intType = 16 Or intType = 24 Or intType = 2 Or intType = 3) Then
            frmPlot.FillColor = RGB(255, 0, 0)
            frmPlot.Circle (sglXcoor, sglYCoor), 0.5, RGB(255, 0, 0)
        ElseIf (intType = 48 Or intType = 56 Or intType = 6 Or intType = 7) Then
            frmPlot.FillColor = RGB(0, 255, 0)
            frmPlot.Circle (sglXcoor, sglYCoor), 0.5, RGB(0, 255, 0)
        ElseIf (intType = 96 Or intType = 12) Then
            frmPlot.FillColor = RGB(255, 128, 128)
            frmPlot.Circle (sglXcoor, sglYCoor), 0.5, RGB(0, 255, 0)
        Else
            'don't plot
    End If
    
    GoTo subexit
    
subexit:
    Exit Sub

End Sub

Private Sub NorthDisp(lngTargCnt() As Long, lngTargCnt30() As Long, sglMean() As Single, lngMean() As Long, lngTargCntTot() As Long, intScanSelect As Integer, varBookmark As Variant)

    'This procedure updates scan and 30 scan average counts on the
    'PPI screen, deletes target plots from 30 scans ago, and updates
    'the total correlated target count on the PPI screen.

    Dim i As Integer

    For i = 0 To 3
        lngMean(i) = lngMean(i) - lngTargCnt30(i, intScanSelect) + lngTargCnt(i)
        If intScnCnt >= 30 Then
                sglMean(i) = lngMean(i) / 30
            Else
                sglMean(i) = lngMean(i) / intScnCnt
        End If
        lngTargCnt30(i, intScanSelect) = lngTargCnt(i)
        lngTargCntTot(i) = lngTargCntTot(i) + lngTargCnt(i)
    Next i
    frmPlot.txtBcn1.Text = Str(lngTargCnt(1))
    frmPlot.txtCorr1.Text = Str(lngTargCnt(2))
    frmPlot.txtSrch1.Text = Str(lngTargCnt(0))
    frmPlot.txtErr1.Text = Str(lngTargCnt(3))
    frmPlot.txtBcn30.Text = Str(CLng(sglMean(1)))
    frmPlot.txtCorr30.Text = Str(CLng(sglMean(2)))
    frmPlot.txtSrch30.Text = Str(CLng(sglMean(0)))
    frmPlot.txtErr30.Text = Str(CLng(sglMean(3)))
    frmPlot.txtScnCnt.Text = Str(intScnCnt)
    frmPlot.txtTotCorrCnt.Text = Str(lngTargCntTot(2))
    If intScnCnt >= 30 Then
        If IsEmpty(varBookmark) = True Then
                rstTarget.MoveFirst
                Do Until rstTarget.Fields("D9_Scn_Num").Value = 1
                    sglXcoor = rstTarget.Fields("D7_XPltCoor").Value
                    sglYCoor = rstTarget.Fields("D8_YPltCoor").Value
                    frmPlot.FillColor = RGB(0, 0, 0)
                    frmPlot.Circle (sglXcoor, sglYCoor), 0.5, RGB(0, 0, 0)
                    rstTarget.MoveNext
                Loop
                varBookmark = rstTarget.Bookmark
            Else
                rstTarget.Bookmark = varBookmark
                Do Until rstTarget.Fields("D9_Scn_Num").Value = intScnCnt - 29
                    sglXcoor = rstTarget.Fields("D7_XPltCoor").Value
                    sglYCoor = rstTarget.Fields("D8_YPltCoor").Value
                    frmPlot.FillColor = RGB(0, 0, 0)
                    frmPlot.Circle (sglXcoor, sglYCoor), 0.5, RGB(0, 0, 0)
                    rstTarget.MoveNext
                Loop
                varBookmark = rstTarget.Bookmark
        End If
    End If
    For i = 1 To 5
        frmPlot.Circle (0, 0), 50 * i, &H8000000F
    Next i
    For i = 0 To 2
        If (lngTargCnt(i) < (0.5 * sglMean(i))) Or (lngTargCnt(i) > (1.5 * sglMean(i))) Then
                blnMessProb = True
                i = 4
            Else
                blnMessProb = False
        End If
    Next i
    If blnMessProb = True Then
            frmPlot.shpMess.FillColor = RGB(255, 255, 0)
        Else
            frmPlot.shpMess.FillColor = RGB(0, 255, 0)
    End If
    DoEvents
    If intScanSelect < 29 Then
            intScanSelect = intScanSelect + 1
        Else
            intScanSelect = 0
    End If
    For i = 0 To 3
        lngTargCnt(i) = 0
    Next i

End Sub

Private Sub Parity(bteParity() As Byte, blnMessWrdBits() As Boolean, intParityCnt As Integer)

    'This procedure determines individual target message bits and
    'develops the parity counts for the message.

    Dim intWrdCnt As Integer
    Dim intBitParity(1, 7) As Integer
    Dim bteResult As Byte
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 1
        For j = 0 To 7
            intBitParity(i, j) = 0
        Next j
    Next i
    
    intWrdCnt = UBound(blnMessWrdBits(), 1)
    
    If blnS311Coll = False Then
            For i = 0 To intWrdCnt
                For j = 0 To 7
                    bteResult = bteParity(i) And CByte(2 ^ j)
                    If bteResult <> 0 Then
                            blnMessWrdBits(i, j) = True
                            intBitParity(0, j) = intBitParity(0, j) + 1
                        Else
                            blnMessWrdBits(i, j) = False
                    End If
                Next j
            Next i
            For i = 0 To 7
                Select Case intBitParity(0, i)
                    Case 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31
                        intParityCnt = intParityCnt + 1
                End Select
            Next i
        Else
            For i = 4 To 20 Step 2
                For j = 0 To 7
                    bteResult = bteParity(i) And CByte(2 ^ j)
                    If bteResult <> 0 Then
                            blnMessWrdBits(i, j) = True
                            intBitParity(0, j) = intBitParity(0, j) + 1
                        Else
                            blnMessWrdBits(i, j) = False
                    End If
                Next j
            Next i
            For i = 5 To 21 Step 2
                For j = 0 To 7
                    bteResult = bteParity(i) And CByte(2 ^ j)
                    If bteResult <> 0 Then
                            blnMessWrdBits(i, j) = True
                            intBitParity(1, j) = intBitParity(1, j) + 1
                        Else
                            blnMessWrdBits(i, j) = False
                    End If
                Next j
            Next i
            For i = 0 To 1
                For j = 0 To 7
                    Select Case intBitParity(i, j)
                        Case 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29
                            intParityCnt = intParityCnt + 1
                    End Select
                Next j
            Next i
    End If

End Sub

Private Sub PIMClear(varPIMData As Variant)

    'This procedure clears out any data in the PIM prior to
    'beginning a data collection.

    blnExit = False
    
    If blnModemA = True Then
        If usbModemA.PipeStatus = 0 Then
                If usbModemA.DataReady = True Then
                        usbModemA.GetData varPIMData
                    Else
                        'no old data waiting in the modem so do nothing
                End If
            ElseIf usbModemA.PipeStatus = 1 Then
                MsgBox "You have lost the connection to modem A.", vbOKOnly
                Set usbModemA = Nothing
                blnExit = True
                GoTo subexit
            Else
                MsgBox "A usb port error has occurred.", vbOKOnly
                blnExit = True
                GoTo subexit
        End If
    End If
    If blnModemB = True Then
        If usbModemB.PipeStatus = 0 Then
                If usbModemB.DataReady = True Then
                        usbModemB.GetData varPIMData
                    Else
                        'no old data waiting in the modem so do nothing
                End If
            ElseIf usbModemB.PipeStatus = 1 Then
                MsgBox "You have lost the connection to modem B.", vbOKOnly
                Set usbModemB = Nothing
                blnExit = True
                GoTo subexit
            Else
                MsgBox "A usb port error has occurred.", vbOKOnly
                blnExit = True
                GoTo subexit
        End If
    End If
    If blnModemC = True Then
        If usbModemC.PipeStatus = 0 Then
                If usbModemC.DataReady = True Then
                        usbModemC.GetData varPIMData
                    Else
                        'no old data waiting in the modem so do nothing
                End If
            ElseIf usbModemC.PipeStatus = 1 Then
                MsgBox "You have lost the connection to modem C.", vbOKOnly
                Set usbModemC = Nothing
                blnExit = True
                GoTo subexit
            Else
                MsgBox "A usb port error has occurred.", vbOKOnly
                blnExit = True
                GoTo subexit
        End If
    End If
    If blnS311Coll = True Then
        If usbS311.PipeStatus = 0 Then
                If usbS311.DataReady = True Then
                        usbS311.GetData varPIMData
                    Else
                        'no old data waiting in the modem so do nothing
                End If
            ElseIf usbS311.PipeStatus = 1 Then
                MsgBox "You have lost the connection to the S311 channel.", vbOKOnly
                Set usbS311 = Nothing
                blnExit = True
                GoTo subexit
            Else
                MsgBox "A usb port error has occurred.", vbOKOnly
                blnExit = True
                GoTo subexit
        End If
    End If

    GoTo subexit

subexit:
    Exit Sub

End Sub

Private Sub PIMConnect()

    'This procedure sets the PIM channel connections based on the
    'user configuration inputs.

    If blnModemA = True Then
        Set usbModemA = New Arades.UsbPipe
        usbModemA.Connect CByte(131)
        usbModemA.ModemMode = 0
    End If
    If blnModemB = True Then
        Set usbModemB = New Arades.UsbPipe
        usbModemB.Connect CByte(132)
        usbModemB.ModemMode = 0
    End If
    If blnModemC = True Then
        Set usbModemC = New Arades.UsbPipe
        usbModemC.Connect CByte(133)
        If blnExtColl = True Then
                usbModemC.ModemMode = 1
            Else
                usbModemC.ModemMode = 0
        End If
    End If
    If blnS311Coll = True Then
        Set usbS311 = New Arades.UsbPipe
        usbS311.Connect CByte(134)
        usbS311.ModemMode = 0
    End If

End Sub

Private Sub PIMGetData(varPIMData As Variant, bteDynArray() As Byte, intModemSel As Integer, blnEndProc As Boolean)

    'This procedure retrieves data from the PIM and stores it in registers
    'for use in later message decoding.

    Dim i As Integer
    Dim j As Integer

    intUBound = 0
    intLBound = 0
    intArraySize = 0
    intDynArrayPntr = 0

    If intModemSel = 1 Then
            If usbModemA.PipeStatus = 0 Then
                    If usbModemA.DataReady = True Then
                            usbModemA.GetData varPIMData
                            intUBound = UBound(varPIMData)
                            intLBound = LBound(varPIMData)
                            intArraySize = (intUBound - intLBound) + 1
                            ReDim bteDynArray(intArraySize)
                            j = 0
                            For i = intLBound To intUBound
                                bteDynArray(j) = varPIMData(i)
                                Put #1, , bteDynArray(j)
                                j = j + 1
                            Next i
                        Else
                            blnEndProc = True
                            GoTo subexit
                    End If
                ElseIf usbModemA.PipeStatus = 1 Then
                    MsgBox "You have lost the connection to modem A.", vbOKOnly
                    Set usbModemA = Nothing
                    blnExit = True
                    GoTo subexit
                Else
                    MsgBox "A usb port error has occurred.", vbOKOnly
                    blnExit = True
                    GoTo subexit
            End If
        ElseIf intModemSel = 2 Then
            If usbModemB.PipeStatus = 0 Then
                    If usbModemB.DataReady = True Then
                            usbModemB.GetData varPIMData
                            intUBound = UBound(varPIMData)
                            intLBound = LBound(varPIMData)
                            intArraySize = (intUBound - intLBound) + 1
                            ReDim bteDynArray(intArraySize)
                            j = 0
                            For i = intLBound To intUBound
                                bteDynArray(j) = varPIMData(i)
                                Put #2, , bteDynArray(j)
                                j = j + 1
                            Next i
                        Else
                            blnEndProc = True
                            GoTo subexit
                    End If
                ElseIf usbModemB.PipeStatus = 1 Then
                    MsgBox "You have lost the connection to modem B.", vbOKOnly
                    Set usbModemB = Nothing
                    blnExit = True
                    GoTo subexit
                Else
                    MsgBox "A usb port error has occurred.", vbOKOnly
                    blnExit = True
                    GoTo subexit
            End If
        ElseIf intModemSel = 3 Then
            If usbModemC.PipeStatus = 0 Then
                    If usbModemC.DataReady = True Then
                            usbModemC.GetData varPIMData
                            intUBound = UBound(varPIMData)
                            intLBound = LBound(varPIMData)
                            intArraySize = (intUBound - intLBound) + 1
                            ReDim bteDynArray(intArraySize)
                            j = 0
                            For i = intLBound To intUBound
                                bteDynArray(j) = varPIMData(i)
                                Put #3, , bteDynArray(j)
                                j = j + 1
                            Next i
                        Else
                            blnEndProc = True
                            GoTo subexit
                    End If
                ElseIf usbModemC.PipeStatus = 1 Then
                    MsgBox "You have lost the connection to modem C.", vbOKOnly
                    Set usbModemC = Nothing
                    blnExit = True
                    GoTo subexit
                Else
                    MsgBox "A usb port error has occurred.", vbOKOnly
                    blnExit = True
                    GoTo subexit
            End If
        ElseIf intModemSel = 4 Then
            If usbS311.PipeStatus = 0 Then
                    If usbS311.DataReady = True Then
                            usbS311.GetData varPIMData
                            intUBound = UBound(varPIMData)
                            intLBound = LBound(varPIMData)
                            intArraySize = (intUBound - intLBound) + 1
                            ReDim bteDynArray(intArraySize)
                            j = 0
                            For i = intLBound To intUBound
                                bteDynArray(j) = varPIMData(i)
                                Put #4, , bteDynArray(j)
                                j = j + 1
                            Next i
                        Else
                            blnEndProc = True
                            GoTo subexit
                    End If
                ElseIf usbS311.PipeStatus = 1 Then
                    MsgBox "You have lost the connection to the S311 channel.", vbOKOnly
                    Set usbS311 = Nothing
                    blnExit = True
                    GoTo subexit
                Else
                    MsgBox "A usb port error has occurred.", vbOKOnly
                    blnExit = True
                    GoTo subexit
            End If
        Else
            blnEndProc = True
            GoTo subexit
    End If

    GoTo subexit

subexit:
    Exit Sub

End Sub

Private Sub SWCollConf()

    'This procedure checks the user data collection configuration
    'and sets the appropriate flag values for the rest of the program.

    If chkModemA.Value = 1 Then
            blnModemA = True
        Else
            blnModemA = False
    End If
    If chkModemB.Value = 1 Then
            blnModemB = True
        Else
            blnModemB = False
    End If
    If chkModemC.Value = 1 Then
            blnModemC = True
        Else
            blnModemC = False
    End If
    If optExtended.Value = True Then
            blnExtColl = True
        Else
            blnExtColl = False
    End If
    If optS311.Value = True Then
            blnS311Coll = True
        Else
            blnS311Coll = False
    End If
    
End Sub

Private Sub TargStat(lngTargCntTot() As Long, lngTargCnt() As Long, lngFaultCnt() As Long)

    'This procedure loads the targets count statistics in the database
    'at the conclusion of the data collection.
    
    Dim rstStat As ADODB.Recordset
    Set rstStat = New ADODB.Recordset
    rstStat.ActiveConnection = cnnDB
    rstStat.Source = "TargStat"
    rstStat.Open Options:=adCmdTable, LockType:=adLockOptimistic, CursorType:=adOpenKeyset
    rstStat.Move (2)
    rstStat.Fields("A2_Value").Value = lngTargCntTot(2) + lngTargCnt(2)
    rstStat.Update
    rstStat.MoveNext
    rstStat.Fields("A2_Value").Value = lngTargCntTot(0) + lngTargCnt(0)
    rstStat.Update
    rstStat.MoveNext
    rstStat.Fields("A2_Value").Value = lngTargCntTot(1) + lngTargCnt(1)
    rstStat.Update
    rstStat.MoveNext
    rstStat.Fields("A2_Value").Value = intScnCnt
    rstStat.Update
    rstStat.Move (3)
    rstStat.Fields("A2_Value").Value = lngTargCntTot(3) + lngTargCnt(3)
    rstStat.Update
    rstStat.Move (4)
    rstStat.Fields("A2_Value").Value = lngDBCnt
    rstStat.Update
    rstStat.Close
    
    Set rstStat = New ADODB.Recordset
    rstStat.ActiveConnection = cnnDB
    rstStat.Source = "Site"
    rstStat.Open Options:=adCmdTable, LockType:=adLockOptimistic, CursorType:=adOpenKeyset
    rstStat.Move (46)
    rstStat.Fields("A2_Value").Value = lngFaultCnt(0)
    rstStat.Update
    rstStat.MoveNext
    rstStat.Fields("A2_Value").Value = lngFaultCnt(1)
    rstStat.Update
    rstStat.MoveNext
    rstStat.Fields("A2_Value").Value = lngFaultCnt(2)
    rstStat.Update
    rstStat.Close
    Set rstStat = Nothing

End Sub

Private Sub Word0_HdrChk(intModemSel As Integer, bteWrkRegA() As Byte, bteWrkRegB() As Byte, bteWrkRegC() As Byte, bteWrkRegS311() As Byte, blnWord0 As Boolean, blnEndProc As Boolean)

    'This procedure checks the first word in the message register
    'to determine if a valid message is in the register.

    Dim intWord0Cnt As Integer
    Dim intCntWord(3) As Integer
    Dim i As Integer
    
    Select Case intModemSel
        Case 1
            intWord0Cnt = CInt(bteWrkRegA(0))
            If intWord0Cnt = 0 Then
                    blnWord0 = True
                Else
                    blnWord0 = False
            End If
        Case 2
            intWord0Cnt = CInt(bteWrkRegB(0))
            If intWord0Cnt = 0 Then
                    blnWord0 = True
                Else
                    blnWord0 = False
            End If
        Case 3
            intWord0Cnt = CInt(bteWrkRegC(0))
            If intWord0Cnt = 0 Then
                    blnWord0 = True
                Else
                    blnWord0 = False
            End If
        Case 4
            blnWord0 = True
            For i = 0 To 3
                intCntWord(i) = CInt(bteWrkRegS311(i))
                Select Case i
                    Case 0
                        If intCntWord(0) <> 128 Then
                            blnWord0 = False
                        End If
                    Case 1
                        If intCntWord(1) <> 240 Then
                            blnWord0 = False
                        End If
                    Case 2
                        If intCntWord(2) <> 255 Then
                            blnWord0 = False
                        End If
                    Case 3
                        If intCntWord(3) <> 143 Then
                            blnWord0 = False
                        End If
                End Select
            Next i
        Case Else
            blnEndProc = True
    End Select
    
End Sub

Private Sub WrkReg(intModemSel As Integer, blnMessProc As Boolean, bteWrkRegA() As Byte, bteWrkRegB() As Byte, bteWrkRegC() As Byte, bteWrkRegS311() As Byte, blnEndProc As Boolean)

    'This procedure determines if a message was processed in the
    'registers and sets the register pointer if required.

    Dim i As Integer

    Select Case intModemSel
        Case 1
            If blnMessProc = False Then
                    For i = 0 To 14
                        bteWrkRegA(i) = bteWrkRegA(i + 1)
                    Next i
                    intWrkRegAPntr = 15
                Else
                    intWrkRegAPntr = 0
            End If
         Case 2
            If blnMessProc = False Then
                    For i = 0 To 14
                        bteWrkRegB(i) = bteWrkRegB(i + 1)
                    Next i
                    intWrkRegBPntr = 15
                Else
                    intWrkRegBPntr = 0
            End If
         Case 3
            If blnMessProc = False Then
                    For i = 0 To (UBound(bteWrkRegC()) - 1)
                        bteWrkRegC(i) = bteWrkRegC(i + 1)
                    Next i
                    intWrkRegCPntr = UBound(bteWrkRegC())
                Else
                    intWrkRegCPntr = 0
            End If
        Case 4
            If blnMessProc = False Then
                    For i = 0 To 20
                        bteWrkRegS311(i) = bteWrkRegS311(i + 1)
                    Next i
                    intWrkRegS311Pntr = 21
                Else
                    intWrkRegS311Pntr = 0
            End If
        Case Else
            blnEndProc = True
    End Select

End Sub

Private Sub PreparePltFrm()

    'This procedure initializes all text boxes,
    'clears the graphics, and sets the scale on the plot screen.

    Dim i As Integer

    frmPlot.Cls
    
    frmPlot.txtBcn1.Text = ""
    frmPlot.txtBcn30.Text = ""
    frmPlot.txtCorr1.Text = ""
    frmPlot.txtCorr30.Text = ""
    frmPlot.txtErr1.Text = ""
    frmPlot.txtErr30.Text = ""
    frmPlot.txtScnCnt.Text = ""
    frmPlot.txtSrch1.Text = ""
    frmPlot.txtSrch30.Text = ""
    frmPlot.txtTotCorrCnt = ""
    
    frmPlot.Scale (-325, 325)-(325, -325)
    
    frmPlot.shpMess.FillColor = RGB(0, 0, 255)
    
    frmDataCollection.Hide
    frmPlot.Show
    frmPlot.cmdStop.SetFocus
    
    For i = 1 To 5
        frmPlot.Circle (0, 0), 50 * i, &H8000000F
    Next i
    DoEvents
    
End Sub

Private Sub FillDB()

    'This procedure fills in the initial and default values for the
    'site, filter, radiosonde, and target statistics tables
    'in the database.

    Dim cnnDB As ADODB.Connection
    Dim rstFill(3) As ADODB.Recordset
    Dim intModem As Integer
    Dim i As Integer
    Dim ktemp As Single
    Dim term1 As Single
    Dim term2 As Single
    Dim term3 As Single
    Dim term4 As Single
    Dim satpres As Single
    Dim parpres As Single
    Dim indref As Single
    Dim kfact As Single
    
    ktemp = frmSiteInfo.txtTemp.Text + 273.15
    If ktemp = 0 Then
        ktemp = 0.1
    End If
    term1 = 5.02808 * (Log(ktemp) / Log(10#))
    term2 = 1.3816 * 10 ^ (11.334 - (0.0303998 * ktemp) - 7)
    term3 = 8.1328 * 10 ^ (3.49149 - 3 - (1302.8844 / ktemp))
    term4 = 2949.076 / ktemp
    satpres = 10 ^ (23.832241 - term1 - term2 + term3 - term4)
    parpres = (frmSiteInfo.txtRelHum.Text / 100) * satpres
    indref = ((77.66 * frmSiteInfo.txtPress.Text) / ktemp) + ((parpres * 3.73 * 10 ^ 5) / ktemp ^ 2)
    kfact = ((1.1116 * 10 ^ -5) * indref ^ 2) - ((4.68 * 10 ^ -3) * indref) + 1.686
    If kfact = 0 Then
        kfact = 0.01
    End If
    
    Set cnnDB = New ADODB.Connection
    cnnDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51; " _
                                & "Data Source= " & frmDBFileNaming.strDB
    cnnDB.Open
    For i = 0 To 3
        Set rstFill(i) = New ADODB.Recordset
    Next i
    For i = 0 To 3
        rstFill(i).ActiveConnection = cnnDB
        If i = 0 Then
                rstFill(i).Source = "Site"
            ElseIf i = 1 Then
                rstFill(i).Source = "Radiosonde"
            ElseIf i = 2 Then
                rstFill(i).Source = "Filter"
            ElseIf i = 3 Then
                rstFill(i).Source = "TargStat"
        End If
        rstFill(i).Open Options:=adCmdTable, LockType:=adLockOptimistic, CursorType:=adOpenKeyset
    Next i
    
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Temperature"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtTemp.Text)
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Pressure"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtPress.Text)
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "RelHumidity"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtRelHum.Text)
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Elevation"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntHgt.Text)
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "TiltAngle"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text)
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "KFactor"
    rstFill(0).Fields("A2_Value").Value = kfact
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "1XO1"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text) + 1.31
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "1XO2"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text) + 3.08
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "1XO3"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text) + 5.19
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "1XO4"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text) + 7.88
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "1XO5"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text) + 12.49
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "2XO1"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text) + 2.1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "2XO2"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text) + 3.7
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "2XO3"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text) + 5.9
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "2XO4"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.txtAntTlt.Text) + 8.6
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm1/2Perc"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm2/3Perc"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm3/4Perc"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm4/5Perc"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm5/6Perc"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm1/2Cnt"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm2/3Cnt"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm3/4Cnt"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm4/5Cnt"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm5/6Cnt"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "CrysSet"
        Select Case frmSiteInfo.cboCrystal.Text
            Case "A"
                rstFill(0).Fields("A2_Value").Value = 1
            Case "B"
                rstFill(0).Fields("A2_Value").Value = 2
            Case "C"
                rstFill(0).Fields("A2_Value").Value = 3
            Case "D"
                rstFill(0).Fields("A2_Value").Value = 4
        End Select
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Freq"
    rstFill(0).Fields("A2_Value").Value = CSng(frmSiteInfo.cboFreq.Text)
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "AntSerNum"
    rstFill(0).Fields("A2_Value").Value = CSng(frmDBFileNaming.txtAntSerNum.Text)
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "CollType"
        If frmDataCollection.optNormal = True Then
                rstFill(0).Fields("A2_Value").Value = 0
            ElseIf frmDataCollection.optExtended = True Then
                rstFill(0).Fields("A2_Value").Value = 1
            ElseIf frmDataCollection.optS311 = True Then
                rstFill(0).Fields("A2_Value").Value = 2
            Else
                rstFill(0).Fields("A2_Value").Value = -1
        End If
    rstFill(0).Update
    intModem = 0
    If frmDataCollection.chkModemA.Value = 1 Then
        intModem = intModem + 1
    End If
    If frmDataCollection.chkModemB.Value = 1 Then
        intModem = intModem + 2
    End If
    If frmDataCollection.chkModemC.Value = 1 Then
        intModem = intModem + 4
    End If
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Modems"
    rstFill(0).Fields("A2_Value").Value = intModem
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm1/2SD"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm2/3SD"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm3/4SD"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm4/5SD"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm5/6SD"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm1/2SD+"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm2/3SD+"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm3/4SD+"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm4/5SD+"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Bm5/6SD+"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "HgtBias"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "HgtRMS"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "HgtSD"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Hgt5250P"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "Hgt5250M"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "HgtTot"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "RdrFlt"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "ProcFlt"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = "OFFlt"
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    rstFill(0).AddNew
    rstFill(0).Fields("A1_Parameter").Value = frmDBFileNaming.txtUntDes.Text
    rstFill(0).Fields("A2_Value").Value = -1
    rstFill(0).Update
    
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "SurfPress"
    rstFill(1).Fields("A2_Value").Value = CSng(frmRadiosonde.txtSurfPress.Text)
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "SurfAlt"
    rstFill(1).Fields("A2_Value").Value = CSng(frmRadiosonde.txtAltSurf.Text)
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "P500Alt"
    rstFill(1).Fields("A2_Value").Value = CSng(frmRadiosonde.txtAlt500.Text)
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "P400Alt"
    rstFill(1).Fields("A2_Value").Value = CSng(frmRadiosonde.txtAlt400.Text)
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "P300Alt"
    rstFill(1).Fields("A2_Value").Value = CSng(frmRadiosonde.txtAlt300.Text)
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "P250Alt"
    rstFill(1).Fields("A2_Value").Value = CSng(frmRadiosonde.txtAlt250.Text)
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "P200Alt"
    rstFill(1).Fields("A2_Value").Value = CSng(frmRadiosonde.txtAlt200.Text)
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "P150Alt"
    rstFill(1).Fields("A2_Value").Value = CSng(frmRadiosonde.txtAlt150.Text)
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "D1Value"
    rstFill(1).Fields("A2_Value").Value = 3
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "D2Value"
    rstFill(1).Fields("A2_Value").Value = 4
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "D3Value"
    rstFill(1).Fields("A2_Value").Value = 5
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "D4Value"
    rstFill(1).Fields("A2_Value").Value = 6
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "D5Value"
    rstFill(1).Fields("A2_Value").Value = 7
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "D6Value"
    rstFill(1).Fields("A2_Value").Value = 8
    rstFill(1).Update
    rstFill(1).AddNew
    rstFill(1).Fields("A1_Parameter").Value = "D7Value"
    rstFill(1).Fields("A2_Value").Value = 9
    rstFill(1).Update
    
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtAz1Strt"
    rstFill(2).Fields("A2_Value").Value = -1
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtAz1Stp"
    rstFill(2).Fields("A2_Value").Value = -1
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtAz2Strt"
    rstFill(2).Fields("A2_Value").Value = -1
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtAz2Stp"
    rstFill(2).Fields("A2_Value").Value = -1
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtCumRngMin"
    rstFill(2).Fields("A2_Value").Value = 10
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtCumRngMax"
    rstFill(2).Fields("A2_Value").Value = 180
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtBcnHgtMin"
    rstFill(2).Fields("A2_Value").Value = 6
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtBcnHgtMax"
    rstFill(2).Fields("A2_Value").Value = 50
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtRnlgthMin"
    rstFill(2).Fields("A2_Value").Value = 0
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtRnlgthMax"
    rstFill(2).Fields("A2_Value").Value = 48
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtScrnAng"
    rstFill(2).Fields("A2_Value").Value = 0
    rstFill(2).Update
    rstFill(2).AddNew
    rstFill(2).Fields("A1_Parameter").Value = "HgtMaxHgtDif"
    rstFill(2).Fields("A2_Value").Value = 6
    rstFill(2).Update
        
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "ByteCount"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "MessCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "CorrCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "SrchCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "BcnCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "NMCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "SMCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "IdleCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "ParErrCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "HdrErrCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "SyncErrCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "MSBErrCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    rstFill(3).AddNew
    rstFill(3).Fields("A1_Parameter").Value = "DBRecCnt"
    rstFill(3).Fields("A2_Value").Value = -1
    rstFill(3).Update
    
    For i = 0 To 3
        rstFill(i).Close
        Set rstFill(i) = Nothing
    Next i
    cnnDB.Close
    Set cnnDB = Nothing

End Sub

Private Sub MakeDB()

    'This procedure creates the database for the data collection.

    Dim fsoDirectory As FileSystemObject
    Dim strAntSerNum As String
    Dim strPath As String
    Dim tblDB(4) As ADOX.Table
    Dim catDB As ADOX.Catalog
    Dim intLength As Integer
    Dim strModem As String
    Dim i As Integer
    
    Set fsoDirectory = New FileSystemObject
    strPath = "C:\GTACSNTT"
    If fsoDirectory.FolderExists(strPath) = False Then
        fsoDirectory.CreateFolder (strPath)
    End If
    strAntSerNum = frmDBFileNaming.txtAntSerNum.Text
    strPath = "C:\GTACSNTT\SN10" & strAntSerNum
    If fsoDirectory.FolderExists(strPath) = False Then
        fsoDirectory.CreateFolder (strPath)
    End If
    
    Set catDB = New ADOX.Catalog
    catDB.Create "Provider=Microsoft.Jet.OLEDB.3.51; " _
                                & "Data Source= " & frmDBFileNaming.strDB
    For i = 0 To 4
        Set tblDB(i) = New ADOX.Table
    Next i
    With tblDB(0)
        .Name = "Targets"
        Set .ParentCatalog = catDB
        With .Columns
            .Append "A1_Targ_ID", adInteger
            .Append "A2_Targ_Type", adVarChar, 12
            .Append "A3_Range", adSingle
            .Append "A4_Azimuth", adSingle
            .Append "A5_Runlength", adSmallInt
            .Append "A6_Rdr_Hgt", adSmallInt
            .Append "A7_V_Rdr_Hgt", adBoolean
            .Append "A8_MC_Code", adSmallInt
            .Append "A9_V_MC", adBoolean
            .Append "B1_MC_Bin_Hgt", adSmallInt
            .Append "B2_MC_Cor_Hgt", adSmallInt
            .Append "B3_M3_Code", adSmallInt
            .Append "B4_M3_X", adBoolean
            .Append "B5_V_M3", adBoolean
            .Append "B6_M1_Code", adSmallInt
            .Append "B7_M1_X", adBoolean
            .Append "B8_V_M1", adBoolean
            .Append "B9_M2_Code", adSmallInt
            .Append "C1_M2_X", adBoolean
            .Append "C2_V_M2", adBoolean
            .Append "C3_SPI", adBoolean
            .Append "C4_7500", adBoolean
            .Append "C5_7600", adBoolean
            .Append "C6_7700", adBoolean
            .Append "C7_Reinforce", adBoolean
            .Append "C8_AF", adBoolean
            .Append "C9_FAA", adBoolean
            .Append "D1_Test", adBoolean
            .Append "D2_D_Hgt", adSmallInt
            .Append "D3_Scr_Ang", adSingle
            .Append "D4_B_El_Ang", adSingle
            .Append "D5_R_El_Ang", adSingle
            .Append "D6_D_El_Ang", adSingle
            .Append "D7_XPltCoor", adSingle
            .Append "D8_YPltCoor", adSingle
            .Append "D9_Scn_Num", adSmallInt
            .Append "E1_Sth_Cnt", adSmallInt
            .Append "E2_Tim_Stmp", adDouble
            .Append "E3_Eqmt_Flt", adSmallInt
            .Append "E4_Tim_Store", adSingle
            .Append "E5_M4_Cnt", adSmallInt
            .Append "E6_V_M4", adBoolean
            .Append "E7_Prime_Bm_Amp", adSmallInt
            .Append "E8_Scndry_Bm_Amp", adSmallInt
            .Append "E9_Prime_Bm_Cnt", adSmallInt
            .Append "F1_Scndry_Bm_Cnt", adSmallInt
            .Append "F2_Dyn_Tlt", adSmallInt
            .Append "F3_Sat_Det", adSmallInt
            .Append "F4_Hgt_MTI_Cnt", adSmallInt
            .Append "F5_Last_Freq", adSmallInt
            .Append "F6_Freq_Mode", adSmallInt
            .Append "F7_MFBP_Num", adSmallInt
            .Append "F8_MFBP_Cnt", adSmallInt
            .Append "F9_2MFBP_Num", adSmallInt
            .Append "G1_2MFBP_Cnt", adSmallInt
            .Append "G2_MFBP_1Num", adSmallInt
            .Append "G3_MFBP_1Cnt", adSmallInt
            .Append "G4_MFBP_2Num", adSmallInt
            .Append "G5_MFBP_2Cnt", adSmallInt
            .Append "G6_Emer_Bit", adBoolean
            .Append "G7_Coll_Flt", adSmallInt
        End With
    End With
    For i = 1 To 4
        With tblDB(i)
            If i = 1 Then
                    .Name = "Site"
                ElseIf i = 2 Then
                    .Name = "Radiosonde"
                ElseIf i = 3 Then
                    .Name = "Filter"
                ElseIf i = 4 Then
                    .Name = "TargStat"
            End If
            Set .ParentCatalog = catDB
            With .Columns
                .Append "A1_Parameter", adVarChar, 12
                .Append "A2_Value", adSingle
            End With
        End With
    Next i
    For i = 0 To 4
        catDB.Tables.Append tblDB(i)
    Next i
    For i = 0 To 4
        Set tblDB(i) = Nothing
    Next i
    Set catDB = Nothing

    intLength = Len(frmDBFileNaming.strDB)
    strModem = Left(frmDBFileNaming.strDB, intLength - 4)
    strModemA = strModem & "mdmA.dat"
    strModemB = strModem & "mdmB.dat"
    strModemC = strModem & "mdmC.dat"
    strS311File = strModem & "S311.dat"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'This procedure quits the application if the user selects the X
    'in the control box while not performing database operations.

    If (frmDataCollection.blnDBOpen = True And UnloadMode = vbFormControlMenu) Then
            MsgBox "Database operations are in progress, please wait until complete to exit the application.", vbOKOnly
            Cancel = 1
        Else
            Call frmMain.AppQuit
    End If

End Sub
