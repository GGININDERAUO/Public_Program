Imports AUO.SubSystemControl
Imports AUO.InspectionFlow
Imports AUO.Cell.AreaFrameIOIF

Imports AUO.Cell.ProberIOIF
Imports AUO.Cell.PatGenIOIF
Imports AUO.Cell.DeMuraIOIF
Imports AUO.Cell.AlarmRule
Imports AUO.Cell.LogRecorder
Imports AUO.Cell.PatGenRS232_DynaColor
Imports AUO.MQEquipStatusReporter
Imports System.Collections.Specialized

Imports System.Windows.Forms
Imports System.IO

Imports AUO.CIM.Agents
Imports System.Xml
Imports System.Text
Imports System.Timers
Imports System.Net
Imports System.Security.Cryptography
Imports System.Data
Imports AUTH
Imports AUTH.LogIn

' <remarks>
' <para>Version History.</para>
' <para>====================================================================================================================</para>
' <para> v2020.06.01 | 2020-06-28 | ??? | 新增DarwinBL -PLC Class </para>
' <para>====================================================================================================================</para>
' </remarks>

Public Class ClsKernel

#Region "Variable"

    Private m_BootRecipe As ClsBootConfig       ' 參數物件
    Private m_TimeoutRecipe As ClsTimeoutRecipe       ' Timeout參數物件
    Private m_NetworkInfo As ClsNetWorkConfig         ' 網路設定
    Private m_AlarmRule As CAlarmRule ' Alarm Rule

    Private m_SystemLog As ClsLogRecorder             ' System Log
    Private m_SpeedLog As ClsLogRecorder
    Private m_ErrorLog As ClsLogRecorder
    'Private m_InspectLog As CInspectionLogRecorder  ' Inspect Log
    Private m_InspectLog As ClsLogRecorder
    Private m_InspectLogForPG2 As ClsLogRecorder
    Private m_AuthLog As ClsLogRecorder               ' Auth Log
    Private m_DefectMonitorLog As ClsLogRecorder

    '------------------------------------------------------------------------------
    'Private m_Auth As CAuth                          ' User Level Control
    Private WithEvents m_NewAuth As ClsAUTH
    Private m_AuthUserInfo As ClsAuthResult

    '------------------------------------------------------------------------------
    Public m_InspectionFlow As ClsInspectionFlow     '檢測流程
    Private m_CurrentModel As ClsModel                '跑貨產品
    Private m_CurrentPattern As ClsPattern
    Private m_FurtherPattern As ClsPattern

    Private m_LastPatGenSignal_1 As eSignal
    Private m_IdxFlow As Integer

    '------------------------------------------------------------------------------
    Private m_Ui As ClsUiIf

    Private WithEvents m_AreaFrame As ClsAreaFrameIOIF

    Private WithEvents m_ProcessReset As ProceRest

    Private m_AreaFrameCommType As eAreaFrameCommType

    Private m_bootRecipeFilename As String

    Private WithEvents m_JudgerManager As ClsJudgerIOIF
    Private WithEvents m_AreaGrabberManager As ClsAreaGrabberIOIF  ' Area Scan CCD Manager
    Private WithEvents m_AreaGrabberManager2 As ClsAreaGrabberIOIF  ' Area Scan CCD Manager2

    '2017 AI 
    Private WithEvents m_AIManager As ClsAIIOIF
    Private m_AIWorker As clsAIWorker
    Private m_MuraMultiFlow As ClsMuraOtherFlow


    'DlgErrBox
    Private WithEvents m_DlgErrBox As DlgErrBox
    'DlgMessBox
    Private WithEvents m_DlgMessBox As DlgMessInfoBox

    '-----WaitPanelContactProcess--------------------------------------------
    Public m_PanelIdFromUI As String            '刷入UI的Panel Id
    Private m_ProberPanelId As String           ' Panel Id
    Private m_ProberPanelId2 As String           'Panel Id for two inspect panel mode
    Private m_ProberPCBId As String            ' PCB Id
    Private m_ProberIsContact As Boolean        ' 是否已經Contact
    Private m_ProberCommExMsg As String         ' 通訊失敗錯誤訊息
    '------------------------------------------------------------------------------

    '----- S06 wait panel rotation to AOI position-----------------------------
    Private m_PLCCommExMsg As String         ' 通訊失敗錯誤訊息
    '------------------------------------------------------------------------------

    Private m_StartTime As Integer                  ' 計算檢測一片的時間
    Private m_StopInspection As Boolean = True            ' 是否按下停止檢測

    '------------------------------------------------------------------------------
    Private m_forwardPanelId As String = ""             '前一片Panel id
    Private m_IsLastPattern As String

    '------------------------------------------------------------------------------
    Private m_PreparationThread As System.Threading.Thread      ' 啟動準備
    Private m_InspectionThread As System.Threading.Thread       ' 跑貨程序
    Private m_AutoAdjustExpTimeThread As System.Threading.Thread       '自動調整曝光時間程序
    '------------------------------------------------------------------------------
    Private JudgerPatternInfoThread As System.Threading.Thread = Nothing       ' Grabber Pattern Info Thread
    Private GrabberPatternInfoThread As System.Threading.Thread = Nothing  ' Judger Pattern Info Thread
    Private ImgProcThread As System.Threading.Thread = Nothing    ' Prober Set Filter Thread
    Private ResetGrabberThread As System.Threading.Thread = Nothing
    Private ResetJudgerThread As System.Threading.Thread = Nothing
    Private GrabberSetExpThread As System.Threading.Thread = Nothing
    Private GrabberGrabThread As System.Threading.Thread = Nothing

    Private m_ResetGrabberExMsg As String
    Private m_ResetJudgerExMsg As String
    Private m_JudgerPatternInfoExMsg As String
    Private m_GrabberPatternInfoExMsg As String
    Private m_ImgProcExMsg As String
    Private m_ImgProcTimeoutExMsg As String
    Private m_WaitContactExMsg As String
    Private m_GrabberSetExpExMsg As String
    Private m_GrabberGrabExMsg As String
    Private m_JudgerResultProcessExMsg As String
    Private m_UnloadPanelExMsgExMsg As String

    Private Alarmdata1 As String = ""
    Private Alarmdata2 As String = ""
    Private Alarmdata3 As String = ""

    Private intSpeed_ResetJudger, intSpeed_ResetGrabber As Integer
    Private intSpeed_PatGenSetPattern, intSpeed_PatGenSetPattern_t1, intSpeed_PatGenSetPattern_t2 As Integer
    Private intSpeed_JudgerPatternInfo As Integer
    Private intSpeed_GrabberPatternInfo As Integer
    Private intSpeed_AreaFrameMove As Integer
    Private intSpeed_ImgProcProcess As Integer
    Private intSpeed_GrabberGrab, intSpeed_GrabberGrab_t1, intSpeed_GrabberGrab_t2 As Integer
    Private intSpeed_GrabberSetExp As Integer
    '------------------------------------------------------------------------------
    Private m_InspectStatus As String
    Private m_InspectOtherReason As String
    Private m_GrabberOutputFilename As String
    Private m_JudgerInputFilename As String
    Private m_AreaGrabberImgProcErrMsg As String
    Private m_Current As Integer
    Private m_RecipeNo As Integer
    Private m_RecipeName As String
    Private m_FuncStationResult As Boolean

    Private m_Other_CCD As String
    Private m_Other_Pattern As String
    Private m_CCDOneOther_CCD As String             'for S06
    Private m_CCDOneOther_Pattern As String         'for S06
    Private m_CCDTwoOther_CCD As String             'for S06
    Private m_CCDTwoOther_Pattern As String         'for S06

    'for S17CT1 one CCD  check two Panel
    Private m_InspectStatus_1 As String             'one Panel or more
    Private m_Other_Pattern_1 As String             'one Panel or more
    Private m_InspectOtherReason_1 As String        'one Panel or more
    Private m_InspectStatus_2 As String             'more Panel
    Private m_Other_Pattern_2 As String             'more Panel
    Private m_InspectOtherReason_2 As String        'more Panel

    '------------------------------------------------------------------------------
    Public Event WorkPoint(ByVal msg As String)
    Public Event StatusMsg(ByVal msg As String)
    Public Event StatusUserMsg(ByVal msg As String)
    Public Event InspectResult(ByVal DateTime As String, ByVal PanelId As String, ByVal MqRank As String, ByVal CstRank As String, ByVal AgsRank As String, ByVal MainDefectCode As String, ByVal ReasonCode As String, ByVal DefectCode As String, ByVal CCD As String, ByVal Pattern As String, ByVal Coordinate As String, ByVal StageNo As String, ByVal TactTime As String)

    Public Event GetChipIDRatio(ByVal ReadFailCount As String)

    Public Event StopInspect()
    '*************************
    Public Event GetPanelID(ByVal id As String)  '將PanelID顯示到UI上
    Public Event GetPanelIDFromUI()              '獲取UI上PanelID
    Public Event GetCurrentUser(ByVal id As String)

    Public Event ChangeRole()
    Public Event InspectionFolwStep(ByVal i As Integer)
    Public Event TxtPanelIDFocus()
    Public Event MainFormEnable(ByVal v As Boolean)


    Public Event UpdateInspectionFlowInfo()
    Public Event ChangeInspectionFlowBackColor(ByVal PatternIndex As String)



    Public Event MainViewShowDefectWindow(ByVal AOIResult As ClsAOIResult, ByVal PatternName As String)
    Public Event SlideViewShowDefect(ByVal AOIResult As ClsAOIResult, ByVal CCD As Integer, ByVal PatternName As String)
    Public Event ShowResult(ByVal JudgeRankResult As Boolean, ByVal MainDefect As String, ByVal LuminaceResult As ClsLuminaceResult)
    Public Event InspectionUIChange(ByVal StopSignal As Boolean)
    Public Event UpdateInformation()

    Public Event ShowAreaFramePCInfo() 'add by xiangping 2015-05-22

    Public Event MCMQConnect()
    Public Event MCMQDisConnect()

    Public Event JudgerConnect()
    Public Event JudgerDisConnect()


    Public Event AreaGrabber1Connect()
    Public Event AreaGrabber2Connect()
    Public Event AreaGrabber3Connect()
    Public Event AreaGrabber4Connect()
    Public Event AreaGrabber1DisConnect()
    Public Event AreaGrabber2DisConnect()
    Public Event AreaGrabber3DisConnect()
    Public Event AreaGrabber4DisConnect()
    Public Event AreaGrabberDisConnect()

    Public Event AIConnect()
    Public Event AIDisconnect()

    Public Event RS232Connect()
    Public Event RS232DisConnect()

    Public Event AreaFrameConnect()
    Public Event AreaFrameDisConnect()


    Public Event MQConnect()
    Public Event MQDisConnect()
    Public Event ErrBoxShow(ByVal strTitle As String, ByVal strCode As String, ByVal strMsg As String, ByVal strDetail As String)


    Public Event StatusChange(ByVal Status As String, ByVal UploadStatus As Boolean)
    Public Event RefreshYieldAfterInspect()
    Public Event NormalFormSize()


    '------------------------------------------------------------------------------
    Private m_RecipeClient As RecipeClient.RecipeClient
    Private m_AlarmRuleRecipe As CAlarmRuleRecipe

    'Add for New Judger String
    Private m_GrabberOutput(3) As String
    Dim m_BypassAbnormalMsg As String
    'Add for New IP & Grabber 
    Private m_AllGrabberOutput As String   ' for Grabber & IP structure
    Private m_Other_CCD_AllDefect As String 'freetomove 集合所有的defect並顯示 for M11 for mulit CCD
    Private m_AllGrabberOutputForRuleBase As String

    'S06 LL CT2
    Private m_CurrentPatternIdx As Integer
    Public m_PatternReady As Boolean
    Public m_PanelCurrentCount As String

    'S16 
    Private m_StickCurrentCount As Integer
    Private m_LastStickCurrentCount As Integer
    Private m_FutherStickCurrentCount As Integer
    Public m_SidelightBacklightReady As Boolean
    'Private m_SetSidelightBacklightExMsg As String
    Private m_FirstActivePosition As Integer
    Public m_MovingFlag As Boolean
    Public m_FilterChanging As Boolean
    Public m_FilterReady As Boolean
    Public m_SetExpReady As Boolean
    Public m_SignalReady As Boolean

    Private m_WaitCanGrabCount As Integer
    Private m_WaitCanSwapCount As Integer
    Public m_YieldNeedUpdate As Boolean
    Public m_ChipInfoMode As Integer
    Public m_GrabberFinished As Boolean
    Public m_MoveFinished As Boolean
    Public PreFilterPos As String
    Public CurrentCCDPosition As Integer
    Public WaitContactTime As Integer

    Private JudgeRank As CJudgeRank = Nothing
    Public Line As String
    Public User_ID As String
    Private m_intInspectionCount As Integer = 1
    '20160121 add by Leon
    Private iInspect_Panel_Cnt As Integer    ' 20160113 add by Leon for Alarm rule re-count

    Private ImageProcessResultArray(4, 4) As String 'first "4" : Inspect Status , InspectOtherReason ,Other_Pattern,Other_CCD  ; Sec "4" CCD1~CCD2 
    Public Event GetPanelIDformUIText()


    Public PanelIDfromUIText As String
    Public PanelIDfromAreaFrame As String
    Public CurrentMQTransationMode As Integer
    '******20161206 M11 JI0506
    Public PCBNoFromUIText As String
    Public Event GetPCBNoFormUIText()
    '***********************

    'For S13
    Public m_bConnectIPFinished As Boolean = False
    Public m_bLoadRecipeFinished As Boolean = False
    Public m_bLoadJudgerRecipeFinished As Boolean = False
    Public m_RS232_DevBoard As ClsRS232Control
    Private IsStartInspect_ForRS232 As Boolean = False

    Public m_MainViewPictureBox As PictureBox = New PictureBox()


    Private Const IdxType As Integer = 0
    Private Const IdxArea As Integer = 3
    Private Const IdxData As Integer = 4
    Private Const IdxGate As Integer = 5
    Private Const IdxPType As Integer = 6
    Private Const IdxHistory As Integer = 11
    Private Const IdxCCDNo As Integer = 12
    Private Const IdxImageFilePath As Integer = 14
    Private Const IdxFileName As Integer = 15
    
    Private m_ViewerIMP As clsViewerIMP
#End Region

#Region "Property"
    Public ReadOnly Property AuthUserInfo As ClsAuthResult
        Get
            Return Me.m_AuthUserInfo
        End Get
    End Property
    Public ReadOnly Property BootRecipe() As ClsBootConfig
        Get
            Return Me.m_BootRecipe
        End Get
    End Property

    Public ReadOnly Property TimeOutRecipe() As ClsTimeoutRecipe
        Get
            Return Me.m_TimeoutRecipe
        End Get
    End Property

    Public ReadOnly Property NetworkInfo() As ClsNetWorkConfig
        Get
            Return Me.m_NetworkInfo
        End Get
    End Property

    Public ReadOnly Property AreaFrame() As ClsAreaFrameIOIF
        Get
            Return Me.m_AreaFrame
        End Get
    End Property

    Public ReadOnly Property AlarmRule() As CAlarmRule
        Get
            Return Me.m_AlarmRule
        End Get
    End Property
    'Public ReadOnly Property FrameRecipe() As CAlarmRule
    '    Get
    '        Return Me.m_FrameRecipe
    '    End Get
    'End Property


    Public ReadOnly Property AlarmRuleRecipe() As CAlarmRuleRecipe
        Get
            Return Me.m_AlarmRuleRecipe
        End Get
    End Property


    Public ReadOnly Property RecipeClient() As RecipeClient.RecipeClient
        Get
            Return Me.m_RecipeClient
        End Get
    End Property

    Public ReadOnly Property SystemLog() As ClsLogRecorder
        Get
            Return Me.m_SystemLog
        End Get
    End Property

    Public ReadOnly Property SpeedLog() As ClsLogRecorder
        Get
            Return Me.m_SpeedLog
        End Get
    End Property

    Public ReadOnly Property ErrorLog() As ClsLogRecorder
        Get
            Return Me.m_ErrorLog
        End Get
    End Property


    Public ReadOnly Property InspectLog() As ClsLogRecorder
        Get
            Return Me.m_InspectLog
        End Get
    End Property
    Public ReadOnly Property InspectLogForPG2() As ClsLogRecorder
        Get
            Return Me.m_InspectLogForPG2
        End Get
    End Property

    Public ReadOnly Property DefectParmsMonitorLog() As ClsLogRecorder
        Get
            Return Me.m_DefectMonitorLog
        End Get
    End Property


    Public ReadOnly Property InspectionFlow() As ClsInspectionFlow
        Get
            Return Me.m_InspectionFlow
        End Get
    End Property

    Public ReadOnly Property CurrentModel() As ClsModel
        Get
            Return Me.m_CurrentModel
        End Get
    End Property

    Public Property StopInspection() As Boolean
        Set(ByVal value As Boolean)
            Me.m_StopInspection = value
        End Set
        Get
            Return Me.m_StopInspection
        End Get
    End Property

    Public Property Ui() As ClsUiIf
        Set(ByVal value As ClsUiIf)
            Me.m_Ui = value
        End Set
        Get
            Return Me.m_Ui
        End Get
    End Property

#End Region

#Region "---Event---"
    Public Event OnLogIn()
    Public Event OnLogOut()

#End Region
    '---function switch------------------------------------------------------------
    Private m_blnAutoChangePatGenModel As Boolean = True

    Private m_blnAutoChkPanelIdLength As Boolean = False
    Private m_blnAutoChkRecipeNo_Prober_InspectFlow As Boolean = True
    Private m_blnAutoChkRecipeName_Prober_InspectFlow As Boolean = False
    Private m_Plc As clsPlcUtil
    Private Const SYSTEM_LOG_PREFIX As String = "System"
    Private Const INSPECT_LOG_PREFIX As String = "Inspect"
    Private Const AUTH_LOG_PREFIX As String = "Auth"
    Private Const DEFECT_MONITOR_LOG_PREFIX As String = "DefectMonitor"
    Private Const SPEED_LOG_PREFIX As String = "Speed"
    Private Const COMM_LOG_PREFIX As String = "Comm"
    Private Const ERROR_LOG_PREFIX As String = "Error"
    Friend Const AI_WORKER_PATH As String = "C:\AOI_System\Recipe\Controller\AIWorkerTable.xml"
    Friend Const AI_PERCENTAGE_TH_PATH As String = "C:\AOI_System\Recipe\Controller\AIPercentageTHTable.xml"
    Friend Const AUTH_FILE_PATH As String = "C:\AOI_System\Recipe\Controller\AUTHBootConfig.xml"

    Public Sub InitController(ByVal bootRecipeFilename As String, Optional ByVal mPlc As clsPlcUtil = Nothing, Optional ByVal Line As String = "", Optional ByVal User_ID As String = "")
        Dim mainLogPath As String
        Dim SystemLogPath As String
        Dim InspectLogPath As String
        Dim AuthLogPath As String
        Dim strtmp As String

        Dim SubSystemResult As CResponseResult = Nothing

        Try
            If Me.m_ViewerIMP Is Nothing Then Me.m_ViewerIMP = New clsViewerIMP()
            Try
                Me.m_BootRecipe = AOIController.ClsBootConfig.ReadXML(bootRecipeFilename)
                Me.m_bootRecipeFilename = bootRecipeFilename
            Catch ex As Exception
                Throw New Exception("BootConfig File: " & bootRecipeFilename & " load fail!")
            End Try
            If Not mPlc Is Nothing Then
                Me.m_Plc = mPlc
            End If
            strtmp = Me.m_BootRecipe.CONTROLLER_LOG_PATH.Trim
            If strtmp.Substring(strtmp.Length - 1, 1) <> "\" Then Me.m_BootRecipe.CONTROLLER_LOG_PATH = strtmp & "\"
            strtmp = Me.m_BootRecipe.JUDGE_RESULT_PATH.Trim
            If strtmp.Substring(strtmp.Length - 1, 1) <> "\" Then Me.m_BootRecipe.JUDGE_RESULT_PATH = strtmp & "\"
            strtmp = Me.m_BootRecipe.LASER_REPAIR_PATH.Trim
            If strtmp.Substring(strtmp.Length - 1, 1) <> "\" Then Me.m_BootRecipe.LASER_REPAIR_PATH = strtmp & "\"
            strtmp = Me.m_BootRecipe.GRABBER_OUTPUT_PATH.Trim
            If strtmp.Substring(strtmp.Length - 1, 1) <> "\" Then Me.m_BootRecipe.GRABBER_OUTPUT_PATH = strtmp & "\"
            strtmp = Me.m_BootRecipe.GRABBER_OUTPUT_BACKUP_PATH.Trim
            If strtmp.Substring(strtmp.Length - 1, 1) <> "\" Then Me.m_BootRecipe.GRABBER_OUTPUT_BACKUP_PATH = strtmp & "\"
            strtmp = Me.m_BootRecipe.FALSE_DEFECT_TABLE_BACKUP_PATH.Trim
            If strtmp.Substring(strtmp.Length - 1, 1) <> "\" Then Me.m_BootRecipe.FALSE_DEFECT_TABLE_BACKUP_PATH = strtmp & "\"

            If Me.m_BootRecipe.JUDGER_MQ_REPORT_BACKUP_PATH = String.Empty Then Me.m_BootRecipe.JUDGER_MQ_REPORT_BACKUP_PATH = "D:\"
            strtmp = Me.m_BootRecipe.JUDGER_MQ_REPORT_BACKUP_PATH.Trim
            If strtmp.Substring(strtmp.Length - 1, 1) <> "\" Then Me.m_BootRecipe.JUDGER_MQ_REPORT_BACKUP_PATH = strtmp & "\"

            mainLogPath = Me.m_BootRecipe.CONTROLLER_LOG_PATH

            '------------ 檢查系統主要Log路徑是否存在 ------
            Try
                If Not System.IO.Directory.Exists(mainLogPath) Then System.IO.Directory.CreateDirectory(mainLogPath)
            Catch ex As Exception
                Throw New Exception("建立System Log檔案的路徑失敗! Build log directory fail! [" & mainLogPath & "]")
            End Try
            '---------End 檢查系統主要Log路徑是否存在 ------

            '---------------- 開啟系統LOG記錄 --------------
            SystemLogPath = mainLogPath & SYSTEM_LOG_PREFIX & "\"
            Try
                If Not System.IO.Directory.Exists(SystemLogPath) Then System.IO.Directory.CreateDirectory(SystemLogPath)
            Catch ex As Exception
                Throw New Exception("建立System Log檔案的路徑失敗! Build System-Log directory fail! [" & SystemLogPath & "]")
            End Try

            If Me.m_SystemLog Is Nothing Then Me.m_SystemLog = New ClsLogRecorder
            Try
                Me.m_SystemLog.Open(SystemLogPath, SYSTEM_LOG_PREFIX, Now)
            Catch ex As Exception
                Me.m_SystemLog = Nothing
                Throw New Exception("開啟System Log檔案失敗! Open System-Log file fail!")
            End Try
            Me.m_SystemLog.WriteLog(Now, ", [Info]Start init Controller.")
            '--------------End 開啟系統LOG記錄 --------------

            '--------------- -Run Speed Log------------------
            If Me.m_BootRecipe.USE_SPEED_LOG Then
                If Me.m_SpeedLog Is Nothing Then Me.m_SpeedLog = New ClsLogRecorder
                Try
                    Me.m_SpeedLog.Open(SystemLogPath, SPEED_LOG_PREFIX, Now)
                Catch ex As Exception
                    Me.m_SpeedLog = Nothing
                    Throw New Exception("開啟Speed Log檔案失敗! Open Speed-Log file fail!")
                End Try
            End If
            '--------------End Run Speed Log------------------

            '----------------- Run Error Log------------------
            If Me.m_BootRecipe.USE_ERROR_LOG Then
                If Me.m_ErrorLog Is Nothing Then Me.m_ErrorLog = New ClsLogRecorder
                Try
                    Me.m_ErrorLog.Open(SystemLogPath, ERROR_LOG_PREFIX, Now)
                Catch ex As Exception
                    Me.m_ErrorLog = Nothing
                    Throw New Exception("開啟Error Log檔案失敗! Open Error-Log file fail!")
                End Try
            End If
            '---------------End Run Error Log-----------------

            '-------------- 開啟檢測LOG記錄 ------------------
            InspectLogPath = mainLogPath & INSPECT_LOG_PREFIX & "\"
            Try
                If Not System.IO.Directory.Exists(InspectLogPath) Then System.IO.Directory.CreateDirectory(InspectLogPath)
            Catch ex As Exception
                Throw New Exception("建立Inspect Log檔案的路徑失敗! Build Inspect-Log directory fail! [" & InspectLogPath & "]")
            End Try
            If Me.m_InspectLog Is Nothing Then Me.m_InspectLog = New ClsLogRecorder

            Try
                Me.m_InspectLog.Open(InspectLogPath, INSPECT_LOG_PREFIX, Now)
            Catch ex As Exception
                Me.m_InspectLog = Nothing
                Throw New Exception("開啟Inspect Log檔案失敗! Open Inspect-Log file fail!")
            End Try
            '-------------End 開啟檢測LOG記錄 ---------------

            '------------- 開啟權限控管LOG記錄 --------------
            'If Me.m_BootRecipe.USE_AUTH Then
            AuthLogPath = mainLogPath & AUTH_LOG_PREFIX & "\"
            Try
                If Not System.IO.Directory.Exists(AuthLogPath) Then System.IO.Directory.CreateDirectory(AuthLogPath) ' 檢查Auth Log的檔案路徑是否存在
            Catch ex As Exception
                Throw New Exception("建立Auth Log檔案的路徑失敗! Build Auth-Log directory fail! [" & AuthLogPath & "]")
            End Try

            If Me.m_AuthLog Is Nothing Then Me.m_AuthLog = New ClsLogRecorder
            Try
                Me.m_AuthLog.Open(AuthLogPath, AUTH_LOG_PREFIX, Now)
            Catch ex As Exception
                Me.m_AuthLog = Nothing
                Throw New Exception("開啟Auth Log檔案失敗! Open Auth-Log file fail!")
            End Try

            If Me.m_NewAuth Is Nothing Then
                Me.m_NewAuth = New ClsAUTH
                RemoveHandler Me.m_NewAuth.OnLogIn, AddressOf Me.AUTH_OnLogIn
                RemoveHandler Me.m_NewAuth.OnLogOut, AddressOf Me.AUTH_OnLogOut
                RemoveHandler Me.m_NewAuth.OnShowDialog, AddressOf Me.AUTH_OnShowDialog

                AddHandler Me.m_NewAuth.OnLogIn, AddressOf Me.AUTH_OnLogIn
                AddHandler Me.m_NewAuth.OnLogOut, AddressOf Me.AUTH_OnLogOut
                AddHandler Me.m_NewAuth.OnShowDialog, AddressOf Me.AUTH_OnShowDialog

                Me.m_NewAuth.Initial(AUTH_FILE_PATH)
            End If
            'End If

            '------------End 開啟權限控管LOG記錄 -------------

            '--------------- 載入Network Setting -------------
            If System.IO.File.Exists(Me.m_BootRecipe.NETWORK_CONFIG_FILE) Then
                Try
                    Me.m_NetworkInfo = ClsNetWorkConfig.ReadXML(Me.m_BootRecipe.NETWORK_CONFIG_FILE)
                Catch ex As Exception
                    Me.m_SystemLog.WriteLog(Now, ", [Error]" & ex.Message)
                    Throw New Exception("載入NetWork-Config檔案錯誤! Load NetWork-Config file: " & Me.m_BootRecipe.NETWORK_CONFIG_FILE & " fail!")
                End Try
            Else
                Throw New Exception("NetWork-Config檔案不存在! NetWork-Config file: " & Me.m_BootRecipe.NETWORK_CONFIG_FILE & " doesn't exist!")
            End If
            '------------End  載入Network Setting ---------------

            '--- --------------設定檢查廠區差異 -----------------
            If Not Me.CheckSite() Then Throw New Exception("Boot-Config中的Generatrion或Phase錯誤! The Generatrion or Phase in Boot-Config is invalid!")
            '-------------End 設定檢查廠區差異 ------------------

            '---------------- 載入 Timeout Recipe ----------------
            If System.IO.File.Exists(Me.m_BootRecipe.TIMEOUT_RECIPE_FILE) Then
                Try
                    Me.m_TimeoutRecipe = ClsTimeoutRecipe.ReadXML(Me.m_BootRecipe.TIMEOUT_RECIPE_FILE)
                Catch ex As Exception
                    Me.m_SystemLog.WriteLog(Now, ", [Error]" & ex.Message)
                    Throw New Exception("載入Timeout-Recipe檔案錯誤! Load Timeout-Recipe file: " & Me.m_BootRecipe.TIMEOUT_RECIPE_FILE & " fail!")
                End Try
            Else
                Throw New Exception("Timeout-Recipe檔案不存在! Timeout-Recipe file: " & Me.m_BootRecipe.TIMEOUT_RECIPE_FILE & " doesn't exist!")
            End If
            '-------------End 載入 Timeout Recipe ---------------

            '-----------------Load Inspection Flow --------------
            If Me.m_InspectionFlow Is Nothing Then Me.m_InspectionFlow = New ClsInspectionFlow
            Me.LoadInspectionFlow()
            '------------End  Load Inspection Flow ---------------

            '--- ----------------Alarm Rule ----------------------
            If Me.m_AlarmRule Is Nothing Then Me.m_AlarmRule = New CAlarmRule()
            Me.m_AlarmRule.Init(Me.m_BootRecipe.ALARM_RULE_FILE)
            '---------------End Alarm Rule ------------------------

            '-----------------------UI IF -------------------------
            If Me.Ui Is Nothing Then Me.Ui = New ClsUiIf

            '------------------ 初始化周邊設備 --------------------
            If Not Me.m_BootRecipe.USE_TEST Then
                RaiseEvent WorkPoint("Check EtherNet IP Setting ... ...")
                Me.CheckIP()
            End If

            '----AI_INIIALIZE------
            If Me.m_CurrentModel.UseAISystem Then
                Me.InitAI()
                If Me.m_AIWorker Is Nothing Then Me.m_AIWorker = New clsAIWorker(AI_WORKER_PATH)
                Me.m_AIWorker.ThresTable = New ClsAIThresholdTable(AI_PERCENTAGE_TH_PATH)
                Me.m_AIWorker.Init()
                'Mura Other Type--LC_BUBBLE...etc
                Me.m_MuraMultiFlow = New ClsMuraOtherFlow()
                Me.m_MuraMultiFlow.AITimeout = Me.m_TimeoutRecipe.AI
                Me.m_MuraMultiFlow.ActionScanInterval = 10
                Me.m_MuraMultiFlow.AI_IMF = Me.m_AIManager
                Me.m_MuraMultiFlow.AI_OUTPUT_PATH = Me.m_BootRecipe.AI_OUTPUT_FOLDER

            End If
            '-------------------End UI IF -------------------------

            '--------- Initial  AreaFrame   --------
            If Me.BootRecipe.USE_AREAFRAME = True Then
                RaiseEvent WorkPoint("Initial AreaFrame Communication ... ...")
                Me.InitAreaFrame()
            Else
                RaiseEvent WorkPoint("UseAreaFrame is False ... ...")
            End If
            '----------------------------------------

            If Me.m_BootRecipe.GENERATION = "S13" AndAlso Me.m_BootRecipe.USE_RS232_DEV_BOARD Then
                Me.InitRS232FreqConvert()
            End If

            '----End  Initial  Pattern Generator  & 確認Pattern Generator版本 ----

            '------------------------ Initial  Error Box --------------------------
            If Me.m_DlgErrBox Is Nothing Then Me.m_DlgErrBox = New DlgErrBox
            '-------------------- End Initial  Error Box --------------------------

            '------------------------ Initial  Mess Box ---------------------------
            If Me.m_DlgMessBox Is Nothing Then Me.m_DlgMessBox = New DlgMessInfoBox
            '-------------------- End Initial  Messr Box --------------------------
            If Me.m_AuthUserInfo Is Nothing Then Me.m_AuthUserInfo = New ClsAuthResult()
            'If Not Me.m_BootRecipe.USE_AUTH Then
            Me.m_AuthUserInfo.Role = "3"
            Me.m_AuthUserInfo.UserID = Me.m_BootRecipe.EQ_ID
            'Else
            '    Me.m_Role = "-1"
            '    Me.m_UserId = "0000000"
            'End If

            ''--- Init SubUnit ---
            Me.CreateSubUnit()

        Catch ex As Exception
            If Me.m_SystemLog IsNot Nothing Then Me.m_SystemLog.WriteLog(Now, ", [Error]" & ex.Message)
            If Me.m_ErrorLog IsNot Nothing Then Me.m_ErrorLog.WriteLog(Now, ", [Error]" & ex.Message)
            Throw New Exception("[InitController]" & ex.ToString)
        Finally
            RaiseEvent WorkPoint("Hide")
        End Try
    End Sub

    Public Sub OpenInspectLog()
        Dim InspectLogPath As String

        InspectLogPath = Me.m_BootRecipe.CONTROLLER_LOG_PATH & INSPECT_LOG_PREFIX & "\"
        Try
            If Not System.IO.Directory.Exists(InspectLogPath) Then System.IO.Directory.CreateDirectory(InspectLogPath)
        Catch ex As Exception
            Throw New Exception("建立Inspect Log檔案的路徑失敗! Build Inspect-Log directory fail! [" & InspectLogPath & "]")
        End Try

        If Me.m_InspectLog Is Nothing Then Me.m_InspectLog = New ClsLogRecorder
        Try
            Me.m_InspectLog.Open(InspectLogPath, INSPECT_LOG_PREFIX, Now)
        Catch ex As Exception
            Me.m_InspectLog = Nothing
            Throw New Exception("開啟Inspect Log檔案失敗! Open Inspect-Log file fail!")
        End Try
        If Me.BootRecipe.GENERATION = "S02_PT" Then
            If Me.m_InspectLogForPG2 Is Nothing Then Me.m_InspectLogForPG2 = New ClsLogRecorder
            Try
                Me.m_InspectLogForPG2.Open(InspectLogPath, INSPECT_LOG_PREFIX & "_2", Now)
            Catch ex As Exception
                Me.m_InspectLogForPG2 = Nothing
                Throw New Exception("開啟Inspect Log檔案失敗! Open Inspect-Log_2 file fail!")
            End Try
        End If

    End Sub

    Public Sub CloseController()
        Try
            Me.DisconnectSubUnit()
            Me.DisconnectAreaFrame()
            Me.DisconnectAI()
            Me.DisconnectFreqConvert()

            If Me.m_AuthLog IsNot Nothing Then Me.m_AuthLog.Close()
            If Me.m_InspectLog IsNot Nothing Then Me.m_InspectLog.Close()
            If Me.m_InspectLogForPG2 IsNot Nothing Then Me.m_InspectLogForPG2.Close()
            If Me.m_SpeedLog IsNot Nothing Then Me.m_SpeedLog.Close()

            If Me.m_ErrorLog IsNot Nothing Then Me.m_ErrorLog.Close()
            If Me.m_SystemLog IsNot Nothing Then Me.m_SystemLog.WriteLog(Now, ", [Info]Close Controller.")
            If Me.m_SystemLog IsNot Nothing Then Me.m_SystemLog.Close()
            If Me.m_MuraMultiFlow IsNot Nothing Then Me.m_MuraMultiFlow.CloseLog()

        Catch ex As Exception
            Throw New Exception("[CloseController] " & ex.Message)
        End Try
    End Sub

    Private Function CheckSite() As Boolean

        Select Case Me.m_BootRecipe.GENERATION.ToUpper
            Case "S13"

            Case Else
        End Select

        '--- Comm Type ---
        Select Case Me.m_NetworkInfo.AREAFRAME_COMM_TYPE.ToUpper
            Case UCase("EtherNet_MXcomponent_DarwinBL")
                Me.m_AreaFrameCommType = eAreaFrameCommType.EtherNet_MXcomponent_DarwinBL
            Case Else
                Me.m_AreaFrameCommType = eAreaFrameCommType.None
        End Select

        Return True
    End Function

#Region "About Comm Sub"

    Public Sub CheckIP()

        If Not My.Computer.Network.IsAvailable Then Throw New Exception("[CheckIP] 請確認網路線是否鬆脫! Please check EtherNet cable!")

        Select Case Me.m_BootRecipe.GENERATION
            Case "S13"
                If Not My.Computer.Network.Ping(Me.m_NetworkInfo.JUDGER_IP) Then Throw New Exception("[CheckIP] ping Judger IP: " & Me.m_NetworkInfo.JUDGER_IP & " fail!")
                If Not My.Computer.Network.Ping(Me.m_NetworkInfo.GRABBER_1_IP) Then Throw New Exception("[CheckIP] ping Grabber 1: " & Me.m_NetworkInfo.GRABBER_1_IP & " fail!")

        End Select
    End Sub

    Public Sub CreateSubUnit()
        Try
            Me.InitJudger()
            Me.InitGrabber1()
            'If Me.m_CurrentModel.UseAISystem Then Me.InitAI()

        Catch ex As Exception
            Throw New Exception("[CreateSubUnit] " & ex.Message)
        End Try

        If Me.m_SystemLog IsNot Nothing Then Me.m_SystemLog.WriteLog(Now, ", [Info]Connect SubUnit.")
    End Sub

    Public Sub InitJudger()

        If Me.m_JudgerManager Is Nothing Then Me.m_JudgerManager = New ClsJudgerIOIF
        Try
            If Not Me.m_JudgerManager.IsConnect Then
                If Me.m_BootRecipe.USE_SUB_SYSTEM_CONTROL_LOG Then
                    Me.m_JudgerManager.Connect(Me.m_NetworkInfo.JUDGER_IP, CInt(Me.m_NetworkInfo.JUDGER_PORT), "D:\AOI_Data\Log\SubSystemControl\")
                Else
                    Me.m_JudgerManager.Connect(Me.m_NetworkInfo.JUDGER_IP, CInt(Me.m_NetworkInfo.JUDGER_PORT))
                End If
                RaiseEvent JudgerConnect()
            End If

        Catch ex As Exception
            Throw New Exception("[InitJudger] " & ex.Message)
        End Try
    End Sub

    Public Sub InitGrabber1()

        If Me.m_AreaGrabberManager Is Nothing Then Me.m_AreaGrabberManager = New ClsAreaGrabberIOIF
        If Me.m_AreaGrabberManager2 Is Nothing Then Me.m_AreaGrabberManager2 = New ClsAreaGrabberIOIF

        Try
            If Not Me.m_AreaGrabberManager.G1IsConnect Then
                If Me.m_BootRecipe.USE_SUB_SYSTEM_CONTROL_LOG Then
                    Me.m_AreaGrabberManager.G1Connect(Me.m_NetworkInfo.GRABBER_1_IP, CInt(Me.m_NetworkInfo.GRABBER_1_1_PORT), "D:\AOI_Data\Log\SubSystemControl\")
                Else
                    Me.m_AreaGrabberManager.G1Connect(Me.m_NetworkInfo.GRABBER_1_IP, CInt(Me.m_NetworkInfo.GRABBER_1_1_PORT))
                End If
            End If

            If Not Me.m_AreaGrabberManager2.G1IsConnect Then
                If Me.m_BootRecipe.USE_SUB_SYSTEM_CONTROL_LOG Then
                    Me.m_AreaGrabberManager2.G1Connect(Me.m_NetworkInfo.GRABBER_1_IP, CInt(Me.m_NetworkInfo.GRABBER_1_2_PORT), "D:\AOI_Data\Log\SubSystemControl\")
                Else
                    Me.m_AreaGrabberManager2.G1Connect(Me.m_NetworkInfo.GRABBER_1_IP, CInt(Me.m_NetworkInfo.GRABBER_1_2_PORT))
                End If

            End If
            RaiseEvent AreaGrabber1Connect()
        Catch ex As Exception
            Throw New Exception("[InitGrabber1] " & ex.Message)
        End Try
    End Sub

    Public Sub InitRS232FreqConvert()
        'Dim RS232_Sidelight1Status, RS232_Sidelight2Status As Boolean
        Dim PatGenRetryLimit As Integer = 0

        '--- 開啟連接Back light and Side light ---
        If Me.m_RS232_DevBoard Is Nothing Then Me.m_RS232_DevBoard = New ClsRS232Control(Me.NetworkInfo.FREQ_CONVERT_PORT, Me.NetworkInfo.FREQ_CONVERT_BAUD_RATE)
        RemoveHandler Me.m_RS232_DevBoard.OnReceiveData, AddressOf m_RS232_DevBoard_OnReceiveData
        AddHandler Me.m_RS232_DevBoard.OnReceiveData, AddressOf m_RS232_DevBoard_OnReceiveData

        Try
            Me.m_RS232_DevBoard.Open()
        Catch ex As Exception
            Me.m_RS232_DevBoard = Nothing  'wait
            Throw New Exception("[InitRS232FreqConvert] Freq-Convert RS232連線錯誤! " & ex.Message)
        End Try

    End Sub

    Public Sub DisconnectFreqConvert()
        Dim PatGenRetryLimit As Integer = 0

        '--- 關閉連接Back light and Side light ---
        Try
            If Me.m_RS232_DevBoard IsNot Nothing Then
                Me.m_RS232_DevBoard.Close()
                RaiseEvent StatusMsg("[DisconnectFreqConvert] Freq-Convert RS232 Disconnect.......")
            End If
        Catch ex As Exception
            Me.m_RS232_DevBoard = Nothing 'wait
            Throw New Exception("[DisconnectFreqConvert] Freq-Convert RS232 斷線錯誤! " & ex.Message)
        End Try
    End Sub

    Public Sub SetFreqConvertSignal(ByVal OnOff As Boolean)
        Try
            Dim strOnOff As String = IIf(OnOff, "1", "0")
            Dim tmpComm As String = String.Format("COV,{0}", strOnOff)
            Me.m_RS232_DevBoard.Enter_Command(tmpComm)
        Catch ex As Exception
            Throw New Exception("[SetFreqConvertSignal] Freq-Convert RS232 傳輸錯誤! " & ex.Message)
        End Try
    End Sub

    Public Sub DisconnectJudger()
        Try
            If Me.m_JudgerManager IsNot Nothing Then Me.m_JudgerManager.Disconnect()
            Me.m_JudgerManager = Nothing
            RaiseEvent JudgerDisConnect()

        Catch ex As Exception
            Throw New Exception("[DisconnectSubUnit] " & ex.Message)
        End Try
    End Sub

    Public Sub DisconnectGrabber()
        Try
            If Me.m_AreaGrabberManager IsNot Nothing Then Me.m_AreaGrabberManager.Disconnect()
            Me.m_AreaGrabberManager = Nothing
            If Me.m_AreaGrabberManager2 IsNot Nothing Then Me.m_AreaGrabberManager2.Disconnect()
            Me.m_AreaGrabberManager2 = Nothing
            RaiseEvent AreaGrabberDisConnect()

        Catch ex As Exception
            Throw New Exception("[DisconnectSubUnit] " & ex.Message)
        End Try
    End Sub

    Public Sub DisconnectAI()
        Try
            If Me.m_AIManager IsNot Nothing Then Me.m_AIManager.Disconnect()
            RaiseEvent AIDisconnect()
        Catch ex As Exception
            Throw New Exception("[DisconnectAI] " & ex.ToString)
        End Try
    End Sub

    Public Sub DisconnectSubUnit()
        Try
            If Me.m_SystemLog IsNot Nothing Then Me.m_SystemLog.WriteLog(Now, ", [Info]Disconnect SubUnit.")

            If Me.m_JudgerManager IsNot Nothing Then Me.DisconnectJudger()
            If Me.m_AreaGrabberManager IsNot Nothing Then Me.DisconnectGrabber()
            If Me.m_AIManager IsNot Nothing Then Me.DisconnectAI()

        Catch ex As Exception
            Throw New Exception("[DisconnectSubUnit] " & ex.Message)
        End Try
    End Sub

    Public Sub InitAreaFrame()

        ' --- 開啟連接AreaFrame ---
        If Me.m_AreaFrame Is Nothing Then
            Me.m_AreaFrame = New ClsAreaFrameIOIF(Me.m_AreaFrameCommType)
        End If

        Try
            If Not Me.m_AreaFrame.IsConnect Then

                Select Case Me.m_AreaFrameCommType

                    Case eAreaFrameCommType.EtherNet_MXcomponent_DarwinBL
                        Me.m_AreaFrame.Connect(Me.m_NetworkInfo.AREAFRAME_COM_PORT_1)
                        RaiseEvent AreaFrameConnect()

                End Select
                Me.m_AreaFrame.RetryLimit = 2
            End If

        Catch ex As Exception
            Me.m_AreaFrame = Nothing
            Throw New Exception("[InitAreaFrame] AreaFrame連線錯誤! Connect AreaFrame fail! " & ex.Message)
        End Try

    End Sub

    Public Sub DisconnectAreaFrame()

        If Me.m_SystemLog IsNot Nothing Then Me.m_SystemLog.WriteLog(Now, ", [Info]Disconnect AreaFrame Comunication.")

        If Me.m_AreaFrame IsNot Nothing Then
            Me.m_AreaFrame.Disconnect()
            Me.m_AreaFrame = Nothing
            RaiseEvent AreaFrameDisConnect()
        End If
    End Sub


#End Region

#Region "Comm Event"

    Private Sub m_AreaFrame_RemoteDisconnect() Handles m_AreaFrame.RemoteDisconnect
        Me.DisconnectAreaFrame()
        RaiseEvent AreaFrameDisConnect()
    End Sub

    Private Sub m_JudgerManager_RemoteDisconnect() Handles m_JudgerManager.RemoteDisconnect
        Me.DisconnectJudger()
        RaiseEvent JudgerDisConnect()
    End Sub

    Private Sub m_AreaGrabberManager_RemoteDisconnect() Handles m_AreaGrabberManager.RemoteDisconnect
        Me.DisconnectGrabber()
        RaiseEvent AreaGrabberDisConnect()
    End Sub
#End Region

    Public Sub StartInspect()
        Dim tmpDir As DirectoryInfo
        Dim tmpFileList As FileInfo()
        Dim i As Integer

        Me.m_SystemLog.WriteLog(Now, ", [Info]User(" & Me.m_AuthUserInfo.UserID & ") Start Inspect=" & [Enum].GetName(GetType(eRunMode), Me.m_Ui.RunMode))
        If Me.m_AuthLog IsNot Nothing Then Me.m_AuthLog.WriteLog(Now, ", [Info]User(" & Me.m_AuthUserInfo.UserID & ") Start Inspect=" & [Enum].GetName(GetType(eRunMode), Me.m_Ui.RunMode))

        If Not Me.m_BootRecipe.USE_TEST Then
            tmpDir = New DirectoryInfo(Me.m_BootRecipe.GRABBER_OUTPUT_PATH)
            tmpFileList = tmpDir.GetFiles()

            For i = 0 To tmpFileList.Length - 1
                File.Delete(tmpFileList(i).FullName)
            Next
        End If

        Me.m_PreparationThread = New System.Threading.Thread(AddressOf Me.PreparationProcess)
        Me.m_PreparationThread.Name = "Preparation"
        Me.m_PreparationThread.Start()
    End Sub

    Private Sub PreparationProcess()
        Dim AreaFrameResult As New ClsAreaFrameResult
        Dim PointCount As ClsPointCount
        Dim FalseCount As ClsFalseCount
        Dim PointCountA, PointCountB As New ClsPointCount
        Dim FalseCountA, FalseCountB As New ClsFalseCount
        Dim SubSystemResult As CResponseResult = Nothing
        Dim Point(4) As Integer
        Dim FunPoint(2) As Integer
        Dim MuraPoint(2) As Integer

        Try
            Me.m_StopInspection = False

            RaiseEvent StatusMsg("檢測啟動中...!!")
            RaiseEvent StatusChange("RUN", False)

            If Not Me.m_InspectionFlow.hasCurrentModel Then Throw New Exception("沒有可使用的產品檢測流! There is no InspectionFlow Model for Inspection!")

            If Not Me.CheckImgProcRecipe Then Throw New Exception("檢測流程的影像處理Recipe未指定檢Func or Mura, 請重新確認! InspectionFlow's ImgProc Recipe don't match for use, please check it!")

            '-----------------Load Inspection Flow --------------
            If Me.m_BootRecipe.GENERATION <> "S13" Then
                If Me.m_InspectionFlow Is Nothing Then Me.m_InspectionFlow = New ClsInspectionFlow
                Me.LoadInspectionFlow()
            End If

            '--- InspectionFlow Recipe---
            Me.m_RecipeNo = Me.m_InspectionFlow.CurrentModel.MODEL_NO
            Me.m_RecipeName = Me.m_InspectionFlow.CurrentModel.MODEL_NAME
            '------------End  Load Inspection Flow -------------

            '--- Init SubUnit ---
            'Me.CreateSubUnit()
            '--------------------    

            If m_CurrentModel.UseAISystem Then
                '-------- AI Load Model -------
                If Me.m_AIWorker.SetCurrentModel(Me.m_CurrentModel.MODEL_NAME) = False Then
                    Dim ERR As String = String.Empty
                    If Me.m_AIWorker.CreateNewModel(Me.m_CurrentModel.MODEL_NAME, ERR) = False Then
                        Throw New Exception(ERR)
                    End If

                    Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[AI_WORKER][CREATE_NEW_MODEL]未建立MODEL, 自動產生. MODEL_NAME : " & Me.m_CurrentModel.MODEL_NAME)
                End If

                Try
                    Me.AIStopWork(Me.m_CurrentModel.UseAISystem)
                    Me.AILoadModel(Me.m_CurrentModel.UseAISystem)
                Catch ex As Exception
                    Throw New Exception("[AI] AOI <-> AI 連線異常. 請確認AI通訊設定並重啟. [ERR_MSG] : " + ex.Message)
                End Try
                '------------------------

                'Mura Multi Flow Run & Set
                If Me.m_CurrentModel.EDGE_AI_TURN_ON AndAlso Me.m_CurrentModel.UseAISystem Then
                    Me.m_MuraMultiFlow.AIWorker = Me.m_AIWorker.CurrentModel
                    Me.m_MuraMultiFlow.Run()
                End If
            End If
            '------END AI LOAD MODEL------

            '----------------------------Frame & PG & Prober Reset,Connect------------------------------------
            If Me.m_Ui.RunMode <> eRunMode.LoadImage Then

                '---------------------------------------Initial AOI Frame-------------------------------------------------
                If Me.BootRecipe.USE_AREAFRAME = True Then
                    '--- InspectionFlow Recipe---
                    Me.m_RecipeNo = Me.m_InspectionFlow.CurrentModel.MODEL_NO
                    Me.m_RecipeName = Me.m_InspectionFlow.CurrentModel.MODEL_NAME
                    If Me.m_AreaFrame Is Nothing Then
                        RaiseEvent StatusMsg("AreaFram正在重新連線!")
                        Me.InitAreaFrame()
                        ' -------------- Reset AOI-PC alarm (MJC Frame)--------
                        Me.ResetAOIpcAlarm()
                        ' ------------End Reset AOI-PC alarm (MJC Frame)------
                    Else
                        If Not Me.m_AreaFrame.IsConnect Then
                            RaiseEvent StatusMsg("AreaFram正在重新連線!")
                            Me.InitAreaFrame()
                            ' -------------- Reset AOI-PC alarm (MJC Frame)--------
                            Me.ResetAOIpcAlarm()
                            ' ------------End Reset AOI-PC alarm (MJC Frame)------
                        End If
                    End If
                End If
            End If
            '--------------------------End Initial AOI Frame------------------------------------

            If Me.m_CurrentModel.UseAreaGrabber Then

                If m_bConnectIPFinished = False Then

                    RaiseEvent StatusMsg("Let AreaGrabber to Connect_IP....")
                    'Let AreaGrabber to Connect_IP
                    Me.m_SystemLog.WriteLog(Now, ", [Cmd=CONNECT_IP")
                    Me.m_AreaGrabberManager.PrepareAllRequest("CONNECT_IP", Me.m_Ui.STR_CCD_COUNT, , , , , , , , 5000)
                    SubSystemResult = Me.m_AreaGrabberManager.SendRequest(5000)
                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then
                        Throw New Exception("Grabber執行CONNECT_IP" & CStr(Me.m_AreaGrabberManager.RetryLimit) & "次失敗!" & vbCrLf & "Communicate with Grabber fail! <CONNECT_IP>")
                    Else
                        m_bConnectIPFinished = True
                    End If
                End If
                'Grabber切換為Auto(Mode)
                RaiseEvent StatusMsg("Grabber設定自動模式中(Info Grabber Setting to Auto Mode)...")
                Me.m_AreaGrabberManager.PrepareAllRequest("SET_MODE", "AUTO", , , , , , , , 10000)
                Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=SET_MODE]Grabber Set to Auto Mode")
                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(10000)
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber Set to Auto Mode" & CStr(Me.m_AreaGrabberManager.RetryLimit) & "次失敗! Communicate with Grabber fail! <SET_MODE>")

                If Me.m_bLoadRecipeFinished = False Then
                    Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=LOAD_RECIPE]Grabber載入Recipe")
                    Me.m_AreaGrabberManager.PrepareAllRequest("LOAD_RECIPE", Me.m_CurrentModel.MODEL_NAME.ToUpper, , , , , , , , Me.m_TimeoutRecipe.Grabber.LOAD_RECIPE)
                    SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.LOAD_RECIPE)
                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber連續載入Recipe" & CStr(Me.m_AreaGrabberManager.RetryLimit) & "次失敗!" & vbCrLf & "Communicate with Grabber fail! <LOAD_RECIPE>")
                    Me.m_bLoadRecipeFinished = True
                End If
            End If

            ' Judger load Recipe

            If Me.m_bLoadJudgerRecipeFinished = False Then

                If Me.m_Ui.blnJudge Then
                    RaiseEvent StatusMsg("Judger設定Mode中(Info Judger Set Mode)...")
                    Me.m_JudgerManager.PrepareAllRequest("SET_MODE", 1, 1)
                    Me.m_SystemLog.WriteLog(Now, ", [Cmd=SET_MODE]Judger設定Mode : " & "One CCD inspect one panel ") 'One CCD inspect one panel

                    SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.SET_MODE)
                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Judger設定Mode失敗! Communicate with Judger fail! <SET_MODE>")

                    If m_CurrentModel.UseAISystem Then
                        RaiseEvent StatusMsg("Judger設定AI_SYATEM")
                        Me.m_JudgerManager.PrepareAllRequest("AI_SYSTEM", "True")
                        Me.m_SystemLog.WriteLog(Now, ", [Cmd=AI_SYSTEM]Judger設定AI_SYSTEM:True")
                        Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.SET_EQ_ID)
                    End If

                    RaiseEvent StatusMsg("Judger設定設備編號中(Info Judger Setting EQ_ID...)")
                    Me.m_JudgerManager.PrepareAllRequest("SET_EQ_ID", Me.m_BootRecipe.EQ_ID.ToUpper, Me.m_BootRecipe.GENERATION, Me.m_BootRecipe.PHASE)
                    Me.m_SystemLog.WriteLog(Now, ", [Cmd=SET_EQ_ID]Judger設定設備編號")
                    SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.SET_EQ_ID)
                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Judger設定設備編號失敗 [SET_EQ_ID]!" & vbCrLf & "Communicate with Judger failed [SET_EQ_ID]")

                    RaiseEvent StatusMsg("Judger載入Recipe中(Info Judger Loading Recipe...)")
                    Me.m_JudgerManager.PrepareAllRequest("LOAD_RECIPE", Me.m_CurrentModel.JUDGER_MODEL.ToUpper)
                    Me.m_SystemLog.WriteLog(Now, ", [Cmd=LOAD_RECIPE]Judger載入Recipe")
                    SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.LOAD_RECIPE)
                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Judger載入Recipe失敗 [LOAD_RECIPE]!" & vbCrLf & "Communicat with Judger failed [LOAD_RECIPE]")

                    ' Judger False Defect DP,BP Count
                    RaiseEvent StatusMsg("查詢Judger False Defect總數中(Querying Judger False Defects...)")
                    PointCount = Me.QueryPointCount()
                    Me.m_AlarmRule.setDPBP(PointCount.DP, PointCount.BP)
                    If Me.m_AlarmRule.IsBPDPFullAlarm Then Throw New Exception("Judger False Defect超過設定值(請清潔擴散版或更換玻璃, 再清除Judger False Defect的資料)!" & vbCrLf & vbCrLf & "Judger False Defect counts are over setting value! (Please clear the diffuser or change glass, then delete Judger False Defect data.)")
                    ' Grabber False Defect DP,BP Count
                    RaiseEvent StatusMsg("查詢Grabber False Defect總數中(Querying Grabber False Defects...)")
                    FalseCount = Me.QueryGrabberPointCount()
                    Me.m_AlarmRule.setGrabberFalseDefect(FalseCount.Func, FalseCount.Mura)
                    If Me.m_AlarmRule.IsGrabberFasleDefectFullAlarm Then Throw New Exception("Grabber False Defect超過設定值(請檢查Grabber False Defect的資料)!" & vbCrLf & vbCrLf & "Grabber False Defect counts are over setting value! (Please check Grabber False Defect data.)")
                End If

                Me.m_bLoadJudgerRecipeFinished = True
            End If


            ''Change Auth
            Me.ChangeSubUnitAuth("-1", Me.m_AuthUserInfo.UserID)

            'for DMtest function
            Me.m_LastPatGenSignal_1 = eSignal.TurnOff

            ''---end caroline

            ' 啟動檢測流
            Me.m_InspectionThread = New System.Threading.Thread(AddressOf Me.InspectionProcess)
            Me.m_InspectionThread.Name = "S13_InspectionProcess"
            Me.m_InspectionThread.Start()

            '----------------------------------------------------------------------------------------------------------

        Catch ex As System.Threading.ThreadAbortException
            '--- 緊急停止 ---
            Me.m_SystemLog.WriteLog(Now, ", [Error]PreparationProcess ThreadAbortException")
            Me.m_ErrorLog.WriteLog(Now, ", [Error]PreparationProcess ThreadAbortException")
            MessageBox.Show("[Kernel] [PreparationProcess] " & ex.Message, "Controller Err", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.PreparationProcessException()

        Catch ex As Exception
            Me.m_SystemLog.WriteLog(Now, ", [Error]PreparationProcess Exception => " & ex.ToString)
            Me.m_ErrorLog.WriteLog(Now, ", [Error]PreparationProcess Exception => " & ex.ToString)
            RaiseEvent StatusMsg("發生錯誤，錯誤訊息(System Error, the error message)：" & ex.Message)
            MessageBox.Show("[Kernel] [PreparationProcess] " & ex.Message, "Controller Err", MessageBoxButtons.OK, MessageBoxIcon.Error)
            RaiseEvent InspectionUIChange(True)
            Me.PreparationProcessException()
        End Try
    End Sub

    Private Sub PreparationProcessException()

        'Me.ChangeSubUnitAuth(Me.m_Role, Me.m_UserId)
        'Me.DisconnectSubUnit()
        RaiseEvent ChangeRole()
        RaiseEvent StopInspect()
    End Sub

    Private Sub InspectionProcess()

        Dim WaitPanelContactThread As System.Threading.Thread = Nothing
        Dim WaiMuraAIAnalysisThread As System.Threading.Thread = Nothing
        Dim SubSystemResult As CResponseResult
        Dim i, j As Integer
        Dim InspectDatetime As Date
        Dim strFalseDefectEnable As String
        Dim JudgeRank As CJudgeRank = Nothing
        Dim ImageProcIsFuncDefectIdx As Integer   '影像處理產生Function Defect的Pattern Index

        Dim PointCount As ClsPointCount
        Dim FalseCount As ClsFalseCount

        Dim IsAlarm As CAlarmRule.eAlarmType
        Dim AlarmMsg As String = ""

        Dim strTemp, strBackupPath As String
        Dim t1, t2 As Integer

        Dim RestartCount As Integer = 0
        Dim StopWatch As New Stopwatch

        ' 2008.4.14 Add for Dynamic Exp
        Dim DynamicExpIncCount As Integer = 0
        Dim HasDynamicExp As Boolean = False
        '20210408 LuminancePredictData
        Dim outputLuminanceData As String = ""
        Dim SlideViewErrMsg As String = ""
        Dim LumErrMsg As String = ""
        Dim IS_RUN_MURA_MULTI_AI As Boolean

        Try

            While Not Me.m_StopInspection
                Try

                    ImageProcIsFuncDefectIdx = -1
                    AlarmMsg = ""
                    outputLuminanceData = ""
                    SlideViewErrMsg = ""
                    LumErrMsg = ""
                    Me.m_AreaGrabberImgProcErrMsg = ""
                    Me.m_IsLastPattern = ""
                    Me.m_GrabberOutputFilename = ""
                    Me.m_InspectStatus = "OK"
                    Me.m_InspectOtherReason = ""
                    Me.m_ProberIsContact = False

                    If Me.m_CurrentModel.UseAISystem Then
                        IS_RUN_MURA_MULTI_AI = False
                        Me.m_MuraMultiFlow.InitDataTable()
                        'Me.MuraMultiFlow.TurnOff()
                        'Threading.Thread.Sleep(100)
                        Me.m_MuraMultiFlow.TurnOn()
                    End If

                    'Add by Leon 20150114
                    '清空ImageProcessResultArray
                    For i = 0 To 3
                        For j = 0 To 3
                            Me.ImageProcessResultArray(i, j) = ""
                        Next j
                    Next i
                    Me.ImageProcessResultArray(0, 0) = "OK"
                    Me.ImageProcessResultArray(1, 0) = "OK"

                    '清空GrabberOutput資料
                    For i = 0 To Me.m_GrabberOutput.Length - 1
                        Me.m_GrabberOutput(i) = ""
                    Next i

                    Try
                        If Me.m_Ui.USE_FREQ_CONVERT Then
                            Me.SetFreqConvertSignal(True)
                            Me.m_SystemLog.WriteLog(Now, ", [SetFreqConvertSignal][True] FREQ_CONVERT_DELAY_TIME : " & Me.m_Ui.FREQ_CONVERT_DELAY_TIME & " ms")
                            System.Threading.Thread.Sleep(Me.m_Ui.FREQ_CONVERT_DELAY_TIME)
                        End If

                        If Me.m_Ui.RunMode = eRunMode.Auto Then
                            If Me.m_BootRecipe.USE_AREAFRAME Then
                                RaiseEvent StatusMsg("[等待PLC啟動檢測訊號...]")
                                WaitPanelContactThread = New System.Threading.Thread(AddressOf Me.WaitPanelRotationToAOIPositionProcess)
                                WaitPanelContactThread.Name = "WaitPanelContact"
                                WaitPanelContactThread.Start()
                                WaitPanelContactThread.Join()
                            ElseIf Me.m_BootRecipe.USE_RS232_DEV_BOARD Then
                                RaiseEvent StatusMsg("[等待RS232啟動檢測訊號...]")
                                WaitPanelContactThread = New System.Threading.Thread(AddressOf Me.WaitInspectionSignal)
                                WaitPanelContactThread.Name = "WaitPanelContact"
                                WaitPanelContactThread.Start()
                                WaitPanelContactThread.Join()
                            End If

                            If Me.m_StopInspection <> True Then
                                Me.m_StartTime = Environment.TickCount          ' 從Contact後開始計算檢測時間
                                '從PLC中獲取 Panel ID
                                RaiseEvent StatusMsg("獲取Panel ID (Get Panel ID)...")
                                t1 = Environment.TickCount
                                RaiseEvent GetPanelIDformUIText()
                                If Me.PanelIDfromUIText <> "" Then
                                    Me.PanelIDfromUIText = Me.PanelIDfromUIText.Trim.ToUpper
                                Else
                                    Me.PanelIDfromUIText = DateTime.Now.ToString("yyyyMMdd-HHmmss")
                                End If
                                t2 = Environment.TickCount
                                If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Read Panel ID => " & t2 - t1)
                                Me.m_ProberPanelId = Me.PanelIDfromUIText

                                Me.m_ProberIsContact = True
                                Me.m_ProberCommExMsg = ""
                                'Me.m_StartTime = Environment.TickCount          ' 從Contact後開始計算檢測時間
                                Me.m_SystemLog.WriteLog(Now, ", [Get Panel ID] Start Inapsection , CHIP : " & Me.m_ProberPanelId)
                                RaiseEvent GetPanelID(Me.m_ProberPanelId)
                                If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", [PanelID]=" & Me.m_ProberPanelId & "," & [Enum].GetName(GetType(eRunMode), Me.m_Ui.RunMode) & " Mode")
                            End If

                        ElseIf Me.m_Ui.RunMode = eRunMode.Sim Then
                            '====Simulation Inspect mode=====
                            Me.m_StartTime = Environment.TickCount          ' 從Contact後開始計算檢測時間
                            Me.m_ProberPanelId = "Sim" & Format(Now, "yyMMddHHmmss")

                            Me.m_ProberIsContact = True
                            Me.m_ProberCommExMsg = ""
                            RaiseEvent GetPanelID(Me.m_ProberPanelId)
                        Else
                            '====AdjustEXPtime mode======
                            Me.m_StartTime = Environment.TickCount          ' 從Contact後開始計算檢測時間
                            Me.m_ProberPanelId = "AdjustEXPtime" & Format(Now, "yyMMddHHmmss")
                            Me.m_ProberIsContact = True
                            Me.m_ProberCommExMsg = ""
                        End If
                    Catch ex As Exception
                        Me.m_ErrorLog.WriteLog(Now, ", Step1: Get Panel ID => " & ex.Message)
                        Throw New Exception("[Step1: Get Panel ID] " & ex.Message)
                    End Try

                    RestartCount = 0
                    If Me.m_StopInspection Then Exit While ' "停止檢測"被按下

                    '上一片若有動態調整曝光時間則重新計算,反之則累加
                    If Not HasDynamicExp Then
                        DynamicExpIncCount += 1
                    Else
                        DynamicExpIncCount = 1
                        HasDynamicExp = False
                    End If

                    ' 同一片檢測第二次不加FalseDefect
                    strFalseDefectEnable = IIf(Me.m_forwardPanelId <> Me.m_ProberPanelId, "YES", "NO")
                    ' Auto Run,Enable=YES
                    If Me.m_Ui.RunMode <> eRunMode.Auto Then
                        strFalseDefectEnable = "NO"
                    End If
                    InspectDatetime = Now
                Catch ex As Exception
                    Throw New Exception("[Step1: WaitPanel] " & ex.Message)
                End Try
RESTART:
                '檢測次數加一
                RestartCount += 1
                Me.m_Other_CCD = ""
                Me.m_Other_Pattern = ""
                Me.m_CurrentPatternIdx = 0
                Try
                    If Me.m_InspectStatus = "OK" Then
                        Me.m_SystemLog.WriteLog(Now, ", [Info]--- Start Panel Inspection ---")

                        '--------------- Pattern inspect loop ------------------------------------------
                        For i = 0 To Me.m_CurrentModel.PatternCount - 1

                            Me.m_Current = i

                            'If ImageProcIsFuncDefectIdx <> -1 Then Exit For

                            If ImageProcIsFuncDefectIdx = -1 Then

                                If Me.ImageProcessResultArray(0, 0) = "OK" AndAlso Me.ImageProcessResultArray(1, 0) = "OK" Then '1.check is Area type 2.check Prober is contact
                                    '----- Area scan type -------------------------------------                                            
                                    Me.m_CurrentPattern = Me.m_CurrentModel.UsedPatterns(i)

                                    If Me.m_CurrentPattern.TYPE.ToString = "LINE" Then
                                        Throw New Exception("本機台不提供Line Scan 檢測功能")
                                    End If

                                    Me.m_IsLastPattern = IIf(Me.m_CurrentModel.IsLastAreaPattern(i), "YES", "NO")

                                    '-------Auto and Sim mode-----------------------------------
                                    If Me.m_Ui.RunMode <> eRunMode.LoadImage Then
                                        If i = 0 Then
                                            '第一波非同步 第一個Pattern
                                            Try
                                                RaiseEvent ChangeInspectionFlowBackColor(1)
                                                Me.FlowSeq1stThread()
                                            Catch ex As Exception
                                                Me.m_SystemLog.WriteLog(Now, ", [Error]FlowSeq1stThread => " & ex.Message)
                                                Me.m_ErrorLog.WriteLog(Now, ", [Error]FlowSeq1stThread => " & ex.Message)
                                                If Me.m_BootRecipe.USE_AREAFRAME Then
                                                    Me.m_AreaFrame.AlarmControl(eAlarmControl.SetAlarm)
                                                End If
                                                MessageBox.Show("[Kernel] [FlowSeq1stThread] " & ex.Message, "Controller Err", MessageBoxButtons.OK, MessageBoxIcon.Error)

                                                If Me.m_BootRecipe.USE_AREAFRAME Then
                                                    Me.m_AreaFrame.AlarmControl(eAlarmControl.Reset)
                                                End If

                                                'Me.PreparationProcessException()
                                                Exit While
                                            End Try

                                            '------ 擷取影像 ---
                                            If Me.m_Ui.blnGrabImage Then
                                                t1 = Environment.TickCount
                                                Me.m_CurrentPatternIdx = Me.m_CurrentPatternIdx + 1
                                                RaiseEvent StatusMsg(" 擷取第" & CStr(i + 1) & "個Pattern中...")

                                                Me.m_AreaGrabberManager.PrepareAllRequest("GRAB", , , , , , , , , Me.m_TimeoutRecipe.Grabber.GRAB)
                                                Me.m_SystemLog.WriteLog(Now, ", [Cmd=GRAB]Grabber擷取第" & CStr(i + 1) & "個Pattern " & Me.m_CurrentModel.Pattern(Me.m_CurrentPatternIdx - 1).PATTERN_NAME)
                                                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.GRAB)
                                                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Grabber擷取影像失敗! Communicate with Grabber fail! <GRAB>")
                                                t2 = Environment.TickCount
                                                If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", GRAB Image,Seq=" & CStr(i + 1) & " => " & t2 - t1)

                                            End If
                                            '----end of 擷取影像--
                                        End If
                                    Else '-------Load mode-----------------------------------------
                                        Me.FlowSeq1stThread()
                                        '--- 載入影像 ---
                                        t1 = Environment.TickCount
                                        RaiseEvent StatusMsg("載入第" & CStr(i + 1) & "個Pattern中(Loading the " & CStr(i + 1) & " Pattern)...")

                                        Me.m_AreaGrabberManager.PrepareAllRequest("LOAD_IMAGE", Me.m_ProberPanelId, Me.m_CurrentPattern.IMG_PROC_RECIPE, Date.Now.ToString("yyyyMMdd-HHmmss"), , , , , , Me.m_TimeoutRecipe.Grabber.LOAD_IMAGE)
                                        Me.m_SystemLog.WriteLog(Now, ", [Cmd=LOAD_IMAGE]Grabber載入第" & CStr(i + 1) & "個Pattern")
                                        SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.LOAD_IMAGE)
                                        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Grabber載入影像失敗! Communicate with Grabber fail! <LOAD_IMAGE>")
                                        t2 = Environment.TickCount
                                        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Load Image,Seq=" & CStr(i + 1) & " => " & t2 - t1)

                                    End If

                                    If Me.m_Ui.RunMode <> eRunMode.LoadImage AndAlso Me.m_Ui.RunMode <> eRunMode.AutoAdjustExp Then
                                        '20150610 加入Func+Mura(同一個Pattern)流程簡化,縮短TT
                                        'Func(Pattern) Recipe 才需SetWeap,接下來的Mura Pattern(Recipe) bypass
                                        If Me.m_CurrentPattern.FUNC <> ClsPattern.eFunctionType.FUNC_MURA Then
                                            '----- GRABBER SET_SWAP -----
                                            t1 = Environment.TickCount
                                            RaiseEvent StatusMsg("Grabber第" & CStr(i + 1) & "次 SetSwap中(Grabber Swapping)...")
                                            Me.m_AreaGrabberManager.PrepareAllRequest("SET_SWAP", , , , , , , , , Me.m_TimeoutRecipe.Grabber.SET_SWAP)
                                            Me.m_SystemLog.WriteLog(Now, ", [SendCommand=SET_SWAP")
                                            SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.SET_SWAP)
                                            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber切換Grab Image失敗！Communicate with Grabber failed [Swap Grabber Image]" & SubSystemResult.ErrMessage)
                                            t2 = Environment.TickCount
                                            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Grabber Swap Image,Seq=" & CStr(i + 1) & " => " & t2 - t1)
                                            '----------------------------------
                                        End If
                                    End If

                                    'first image save
                                    If i = 0 Then
                                        If Me.m_Ui.blnSaveImage Then
                                            t1 = Environment.TickCount
                                            If Me.m_CurrentPattern.SAVE Then
                                                RaiseEvent StatusMsg("儲存第" & CStr(i + 1) & "個Pattern中(Saving the " & CStr(i + 1) & " Pattern)...")
                                                Dim CutImageAIType As String = ""
                                                If Me.m_CurrentModel.EDGE_AI_TURN_ON AndAlso Me.m_CurrentModel.UseAISystem Then
                                                    CutImageAIType = "CLT"
                                                End If
                                                Me.m_AreaGrabberManager.PrepareAllRequest("SAVE", , , CutImageAIType, , , , , , Me.m_TimeoutRecipe.Grabber.SAVE)
                                                Me.m_SystemLog.WriteLog(Now, ", [Cmd=SAVE]Grabber儲存第" & CStr(i + 1) & "個Pattern " & Me.m_CurrentModel.Pattern(Me.m_CurrentPatternIdx - 1).PATTERN_NAME)
                                                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.SAVE)
                                                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Grabber儲存影像失敗! Communicate with Grabber fail! <SAVE>")
                                                If Me.m_CurrentModel.EDGE_AI_TURN_ON AndAlso Me.m_CurrentModel.UseAISystem Then
                                                    Dim tmpList As List(Of ClsMuraOtherFlow.clsImgInput) = Me.m_MuraMultiFlow.DecorderImagePathToList(SubSystemResult.Responses(0).Param2)
                                                    Me.m_MuraMultiFlow.AddImagePathList(tmpList)
                                                    IS_RUN_MURA_MULTI_AI = True
                                                End If
                                            End If
                                            t2 = Environment.TickCount
                                            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", SAVE Image,Seq=" & CStr(i + 1) & " => " & t2 - t1)
                                        End If
                                    End If

                                    '--- 儲存影像 ---
                                    If i > 0 AndAlso Not i > Me.m_CurrentModel.PatternCount - 1 Then
                                        If Me.m_Ui.blnSaveImage Then
                                            t1 = Environment.TickCount
                                            If Me.m_CurrentPattern.SAVE Then
                                                RaiseEvent StatusMsg("儲存第" & CStr(i + 1) & "個Pattern中(Saving the " & CStr(i + 1) & " Pattern)...")
                                                If Me.m_CurrentPattern.FUNC <> ClsPattern.eFunctionType.FUNC_MURA Then
                                                    Dim CutImageAIType As String = ""
                                                    If Me.m_CurrentModel.EDGE_AI_TURN_ON AndAlso Me.m_CurrentModel.UseAISystem Then
                                                        CutImageAIType = "CLT"
                                                    End If
                                                    Me.m_AreaGrabberManager.PrepareAllRequest("SAVE", , , CutImageAIType, , , , , , Me.m_TimeoutRecipe.Grabber.SAVE)
                                                Else
                                                    Me.m_AreaGrabberManager.PrepareAllRequest("SAVE", , Me.m_CurrentPattern.IMG_PROC_RECIPE, , , , , , , Me.m_TimeoutRecipe.Grabber.SAVE)
                                                End If
                                                Me.m_SystemLog.WriteLog(Now, ", [Cmd=SAVE]Grabber儲存第" & CStr(i + 1) & "個Pattern " & Me.m_CurrentModel.Pattern(Me.m_CurrentPatternIdx - 1).PATTERN_NAME)
                                                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.SAVE)
                                                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Grabber儲存影像失敗! Communicate with Grabber fail! <SAVE>")

                                                If Me.m_CurrentModel.EDGE_AI_TURN_ON AndAlso Me.m_CurrentModel.UseAISystem Then
                                                    Dim tmpList As List(Of ClsMuraOtherFlow.clsImgInput) = Me.m_MuraMultiFlow.DecorderImagePathToList(SubSystemResult.Responses(0).Param2)
                                                    Me.m_MuraMultiFlow.AddImagePathList(tmpList)
                                                    IS_RUN_MURA_MULTI_AI = True
                                                End If

                                            End If
                                            t2 = Environment.TickCount
                                            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", SAVE Image,Seq=" & CStr(i + 1) & " => " & t2 - t1)
                                        End If
                                    End If
                                    '----------------
                                    '-------Judger Pattern Info
                                    RaiseEvent StatusMsg("Judger 設定Pattern Infomation(Judger set patter infomation)...")
                                    Me.m_JudgerManager.PrepareAllRequest("PATTERN_INFO", Me.m_CurrentPattern.IMG_PROC_RECIPE, , , , , , , , Me.m_TimeoutRecipe.Judger.PATTERN_INFO)
                                    Me.m_SystemLog.WriteLog(Now, ", [Cmd=PATTERN_INFO]Judger 設定Pattern Infomation(Judger set patter infomation)")
                                    SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.CLEAR)
                                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Judger 設定PATTERN_INFO 失敗! Communicate with Judger fail! <PATTERN_INFO>")


                                    '-------
                                    ' ----- 第二波非同步 Grabber Async 影像處理同步進行其他運動動作 ------
                                    Me.m_FurtherPattern = Me.m_CurrentPattern
                                    If Not i >= Me.m_CurrentModel.PatternCount - 1 Then
                                        Me.m_CurrentPattern = Me.m_CurrentModel.UsedPatterns(i + 1)
                                    End If
                                    Try
                                        Me.FlowThread(i)
                                    Catch ex As Exception
                                        Me.m_SystemLog.WriteLog(Now, ", [Error]FlowThread => " & ex.Message)
                                        Me.m_ErrorLog.WriteLog(Now, ", [Error]FlowThread => " & ex.Message)
                                        RaiseEvent StatusMsg("發生錯誤，錯誤訊息(System Error, the error message)：" & ex.Message)

                                        If Me.m_BootRecipe.USE_AREAFRAME Then
                                            Me.m_AreaFrame.AlarmControl(eAlarmControl.SetAlarm)
                                        End If
                                        MessageBox.Show("[Kernel] [FlowThread] " & ex.Message, "Controller= Err", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        'Me.PreparationProcessException()

                                        If Me.m_BootRecipe.USE_AREAFRAME Then
                                            Me.m_AreaFrame.AlarmControl(eAlarmControl.Reset)
                                        End If
                                        Exit While
                                    End Try

                                    '--- 影像處理發生錯誤 ---
                                    If Me.m_AreaGrabberImgProcErrMsg <> "" Then Throw New Exception(Me.m_AreaGrabberImgProcErrMsg)

                                    '---- Luminance Measurement Process ----
                                    Try
                                        m_AreaGrabberImgProcErrMsg = ""
                                        Me.LuminanceMeasurementProcess(outputLuminanceData)
                                        Me.m_SystemLog.WriteLog(Date.Now, "[LuminanceMeasurementProcess] Predictor Value : " & outputLuminanceData)
                                        '20210408 - Judger Add Data
                                        If Me.m_AreaGrabberImgProcErrMsg <> "" Then Throw New Exception(Me.m_AreaGrabberImgProcErrMsg)
                                        If outputLuminanceData <> String.Empty Then
                                            outputLuminanceData = outputLuminanceData.Replace("#", ";").TrimEnd(";")
                                            Me.m_JudgerManager.PrepareAllRequest("ADD_LUMINANCE_STREAM", outputLuminanceData, , , , , , , , )
                                            Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_GRABBER_FILTER_STREAM]Judger載入判片資料 ")
                                            SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_LUMINANCE_DATA)
                                            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("IMG_PROC_RECIPE : " & Me.m_FurtherPattern.IMG_PROC_RECIPE & " Judger載入亮度資料失敗! Communicate with Judger fail! <ADD_LUMINANCE_DATA>")
                                        End If
                                    Catch ex As Exception
                                        Throw New Exception("[InspectionProcess][LuminanceMeasurementProcess] ERROR : " & ex.Message)
                                    End Try


                                    If Me.m_InspectStatus = "OK" Then

                                        t1 = Environment.TickCount

                                        If Me.m_Ui.blnJudge Then
                                            RaiseEvent StatusMsg("載入判片資料中(Info Judger loading data)...")

                                            Me.m_AllGrabberOutput = Me.m_GrabberOutput(0)

                                            If Me.m_FurtherPattern.IMG_PROC_RECIPE.Substring(0, 1).ToUpper = "F" Then
                                                ' for Function
                                                If m_AllGrabberOutput <> "" Then
                                                    strBackupPath = IIf(Me.m_Ui.blnBackupFalseDefectTable, Me.m_BootRecipe.FALSE_DEFECT_TABLE_BACKUP_PATH, "")
                                                    Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_FILTER_STREAM", Me.m_AllGrabberOutput, Me.m_FurtherPattern.JUDGER_PARTICLE_RULE.ToString, Me.m_FurtherPattern.JUDGER_PARTICLE_FILTER.ToString, , strBackupPath, , , , )
                                                    Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_GRABBER_FILTER_STREAM]Judger載入判片資料 ")
                                                    SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_LINE_CCD_DATA)
                                                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Judger載入Func判片資料失敗! Communicate with Judger fail! <ADD_GRABBER_FILTER_STREAM>")
                                                End If
                                            Else
                                                ' for Mura
                                                strBackupPath = IIf(Me.m_Ui.blnBackupFalseDefectTable, Me.m_BootRecipe.FALSE_DEFECT_TABLE_BACKUP_PATH, "")
                                                'Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_MURA_STREAM", Me.m_AllGrabberOutput, , , Grabber_Mura_Rim, strBackupPath, , , , "AREA;D:\")--------------------5
                                                Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_MURA_STREAM", Me.m_AllGrabberOutput, Me.m_FurtherPattern.JUDGER_PARTICLE_RULE.ToString, Me.m_FurtherPattern.JUDGER_PARTICLE_FILTER.ToString, , strBackupPath, , , , )
                                                Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_CCD_MURA_STREAM]Judger載入判片資料 ")
                                                SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_CCD_MURA_DATA)
                                                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Judger載入Mura判片資料失敗! Communicate with Judger fail! <ADD_GRABBER_MURA_STREAM>")
                                            End If

                                            If Me.m_Ui.blnPreJudge Then
                                                If Me.m_FurtherPattern.IMG_PROC_RECIPE.Substring(0, 1).ToUpper = "F" Then
                                                    '--- Line Defect is S Grade ---
                                                    RaiseEvent StatusMsg("執行提前判片中(check Line defect)...")
                                                    Me.m_JudgerManager.PrepareAllRequest("PRE_JUDGE_LINE_DEFECT", strFalseDefectEnable)
                                                    Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=PRE_JUDGE_LINE_DEFECT]Judger執行提前判片")
                                                    SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.PRE_JUDGE_LINE_DEFECT)
                                                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Judger提前判片失敗! Communicate with Judger fail! <PRE_JUDGE_LINE_DEFECT>")
                                                    '多片檢不提前退片()
                                                    'If (Me.m_Ui.blnUseCCD1 = True And Me.m_Ui.blnUseCCD2 = False) Or (Me.m_Ui.blnUseCCD1 = False And Me.m_Ui.blnUseCCD2 = True) Then
                                                    '    If SubSystemResult.Responses(0).Param1 = "YES" Then
                                                    '        ImageProcIsFuncDefectIdx = i  ' 產生提前退片
                                                    '    End If
                                                    'End If
                                                Else
                                                    ''--- Filter false Mura ---
                                                    'Me.m_JudgerManager.PrepareAllRequest("FILTER_FALSE_MURA")
                                                    'Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=FILTER_FALSE_MURA]Judger濾除false Mura")
                                                    'SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.FILTER_FALSE_MURA)
                                                    'If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Judger濾除False Mura失敗! Communicate with Judger fail! <FILTER_FALSE_MURA>")

                                                    '--- Mura Defect is S Grade ---
                                                    RaiseEvent StatusMsg("執行提前判片中(check Mura defect)...")

                                                    Me.m_JudgerManager.PrepareAllRequest("PRE_JUDGE_MURA_DEFECT", strFalseDefectEnable)
                                                    Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=PRE_JUDGE_MURA_DEFECT]Judger執行提前判片")
                                                    SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.PRE_JUDGE_MURA_DEFECT)
                                                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Judger提前判片失敗! Communicate with Judger fail! <PRE_JUDGE_MURA_DEFECT>")
                                                    '多片檢不提前退片
                                                    'If (Me.m_Ui.blnUseCCD1 = True And Me.m_Ui.blnUseCCD2 = False) Or (Me.m_Ui.blnUseCCD1 = False And Me.m_Ui.blnUseCCD2 = True) Then
                                                    '    If SubSystemResult.Responses(0).Param1 = "YES" Then
                                                    '        ImageProcIsFuncDefectIdx = i  ' 產生提前退片
                                                    '    End If
                                                    'End If
                                                End If
                                            End If 'end of PreJudge提前退片

                                            '--- Point Count is S Grade ---
                                            If Me.m_Ui.RunMode <> eRunMode.AutoAdjustExp Then
                                                If ImageProcIsFuncDefectIdx = -1 Then

                                                    If Me.m_CurrentModel.IsLastAreaPattern(i) Then ' 最後一個Function Pattern

                                                        If Not Me.m_Ui.blnPreJudge Then
                                                        End If
                                                        '--- Line Defect is S Grade ---
                                                        RaiseEvent StatusMsg("執行提前判片中(check Line defect)...")
                                                        Me.m_JudgerManager.PrepareAllRequest("PRE_JUDGE_LINE_DEFECT", strFalseDefectEnable)
                                                        Me.m_SystemLog.WriteLog(Now, ", [Cmd=PRE_JUDGE_LINE_DEFECT]Judger執行提前判片")
                                                        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.PRE_JUDGE_LINE_DEFECT)
                                                        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Judger提前判片失敗! Communicate with Judger fail! <PRE_JUDGE_LINE_DEFECT>" & vbCrLf & SubSystemResult.ErrMessage)

                                                        RaiseEvent StatusMsg("執行提前判片中(check Point Count defect)...")
                                                        Me.m_JudgerManager.PrepareAllRequest("PRE_JUDGE_POINT_COUNT", strFalseDefectEnable)
                                                        Me.m_SystemLog.WriteLog(Now, ", [Cmd=PRE_JUDGE_POINT_COUNT]Judger執行提前判片")
                                                        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.PRE_JUDGE_POINT_COUNT)
                                                        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Judger提前判片失敗! Communicate with Judger fail! <PRE_JUDGE_POINT_COUNT>")

                                                        RaiseEvent StatusMsg("執行Mura判片中(check Mura defect)...")
                                                        Me.m_JudgerManager.PrepareAllRequest("PRE_JUDGE_MURA_DEFECT", strFalseDefectEnable)
                                                        Me.m_SystemLog.WriteLog(Now, ", [Cmd=PRE_JUDGE_MURA_DEFECT]Judger執行Mura判片")
                                                        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.PRE_JUDGE_MURA_DEFECT)
                                                        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Judger Mura判片失敗! Communicate with Judger fail! <PRE_JUDGE_MURA_DEFECT>")

                                                    End If
                                                End If 'end of 無提前退片
                                            End If 'blnJudge

                                            t2 = Environment.TickCount
                                            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Judge Process,Seq=" & CStr(i + 1) & " => " & t2 - t1)

                                        Else
                                            If Me.m_Ui.RunMode <> eRunMode.AutoAdjustExp Then
                                                ' 產生提前退片
                                                ImageProcIsFuncDefectIdx = i  ' 記錄影像處理產生Function Defect的Pattern Index
                                            End If
                                        End If 'end of 影像處理無Abnormal
                                    End If 'end of Inspection=OK
                                    'End If 'UseMultiCCDInspectOnePanelMode or not
                                End If 'end of ImageProcIsFuncDefectIdx = -1, 無提前退片
                            End If 'end of  Me.m_Ui.RunMode <> eRunMode.AutoAdjustExp
                        Next i 'end of pattern inspect loop
                    End If

                Catch ex As Exception
                    'If RestartCount > 1 Then
                    Throw New Exception("[Step5: ImageInspect] " & ex.ToString)
                End Try



                '-------------Mura AI Analysis----------------
                Try
                    If Me.m_CurrentModel.UseAISystem Then
                        If IS_RUN_MURA_MULTI_AI Then '20180731
                            strBackupPath = IIf(Me.m_Ui.blnBackupFalseDefectTable, Me.m_BootRecipe.FALSE_DEFECT_TABLE_BACKUP_PATH, "")
                            WaiMuraAIAnalysisThread = New System.Threading.Thread(AddressOf Me.WaitMuraAIAnalysisComplete)
                            WaiMuraAIAnalysisThread.Name = "WaiMuraAIAnalysis"
                            WaiMuraAIAnalysisThread.IsBackground = True
                            WaiMuraAIAnalysisThread.Start()
                            WaiMuraAIAnalysisThread.Join()
                            Me.MuraFlowControlSendData(strBackupPath)

                            Me.m_JudgerManager.PrepareAllRequest("PRE_JUDGE_MURA_DEFECT", strFalseDefectEnable)
                            Me.m_SystemLog.WriteLog(Now, ", [Cmd=PRE_JUDGE_MURA_DEFECT]Judger執行Mura判片")
                            SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.PRE_JUDGE_MURA_DEFECT)
                            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(i + 1) & " Judger Mura判片失敗! Communicate with Judger fail! <PRE_JUDGE_MURA_DEFECT>")
                        End If

                        'MuraFlowControl 必須停止不然會影響進片後開啟Task
                        Me.m_MuraMultiFlow.TurnOff()
                    End If
                Catch ex As Exception
                    Throw New Exception("[Step 5.5: Mura AI Analysis] " & ex.Message)
                End Try
                '---------------Mura AI Analysis-----------------------------------------

                Try
                    '--- 判片處理 ---
                    If Me.m_Ui.blnJudge Then
                        RaiseEvent StatusMsg("執行判片處理中(Judging)...")

                        '*******************************************************************************************************************************************
                        ' ImageProcessResultArray(4, 4) As String 'first "4" : Inspect Status , InspectOtherReason ,Other_Pattern,Other_CCD  ; Sec "4" CCD1~CCD2  record by pattaen
                        '__                                                                                                                            __
                        '| Penal 1 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 1 (IP1) Current Pattern ImageProcess Result
                        '| Penal 2 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 2 (IP2) Current Pattern ImageProcess Result
                        '| Penal 2 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 3 (IP3) Current Pattern ImageProcess Result
                        '| Penal 3 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 4 (IP4) Current Pattern ImageProcess Result
                        '__                                                                                                                          __
                        '*******************************************************************************************************************************************

                        'Get Final RanK....
                        t1 = Environment.TickCount
                        Me.m_InspectStatus = Me.ImageProcessResultArray(0, 0)
                        Me.m_InspectOtherReason = Me.ImageProcessResultArray(0, 1)
                        'End If
                        JudgeRank = Me.GetFinalRank(Me.m_InspectStatus, Me.m_InspectOtherReason)
                        t2 = Environment.TickCount
                        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", GetFinalRank => " & t2 - t1)

                        Dim blnUploadReport As Boolean = True
                        If m_CurrentModel.UseAISystem AndAlso Not Me.m_InspectStatus.ToUpper.Contains("OTHER_") Then
                            '啟動AI，先不上報，只SAVE 檢測結果檔案(XML)，
                            blnUploadReport = False
                        Else
                            If Me.m_Ui.RunMode <> eRunMode.Auto Then
                                blnUploadReport = False
                            End If
                        End If
                        Me.MQReport(Me.m_InspectStatus, Me.m_InspectOtherReason, "", Me.m_BootRecipe.JUDGER_MQ_REPORT_BACKUP_PATH, blnUploadReport)
                        'LogJudgeResult....
                        t1 = Environment.TickCount
                        Me.LogJudgeResult(Me.m_InspectStatus, Me.m_InspectOtherReason, Me.m_BootRecipe.JUDGE_RESULT_PATH, JudgeRank.CstRank) 'Parameter 6 : 目前的檢測工位 (決定ViewJudgerReult 更新哪一Form秀出檢測結果)
                        t2 = Environment.TickCount
                        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", LogJudgeResult => " & t2 - t1)

                        Me.m_intInspectionCount += 1

                        If JudgeRank.CCD = "" Then
                            JudgeRank.CCD = "1"
                        End If

                        If JudgeRank.Pattern = "" Then
                            JudgeRank.Pattern = Me.m_FurtherPattern.IMG_PROC_RECIPE
                        End If
                    End If
                Catch ex As Exception
                    Throw New Exception("[Step7: Finally Judge] " & ex.Message)
                End Try


                Try
                    '20170908 AI
                    If Me.m_CurrentModel.UseAISystem AndAlso Not Me.m_InspectStatus.ToUpper.Contains("OTHER_") Then
                        Me.AIProcessFlowControl(False, JudgeRank, "", "")

                        'After AI Process - Get Final Rank again.
                        'Get Final RanK....
                        t1 = Environment.TickCount
                        JudgeRank = Me.GetFinalRank(Me.m_InspectStatus, Me.m_InspectOtherReason)
                        t2 = Environment.TickCount
                        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", [After AI Process]GetFinalRank => " & t2 - t1)

                        Dim blnUploadReport As Boolean = True
                        If Me.m_Ui.RunMode <> eRunMode.Auto Then
                            blnUploadReport = False
                        End If
                        Me.MQReport(Me.m_InspectStatus, Me.m_InspectOtherReason, "", Me.m_BootRecipe.JUDGER_MQ_REPORT_BACKUP_PATH, blnUploadReport)
                        'LogJudgeResult....
                        t1 = Environment.TickCount
                        Me.LogJudgeResult(Me.m_InspectStatus, Me.m_InspectOtherReason, Me.m_BootRecipe.JUDGE_RESULT_PATH, JudgeRank.CstRank) 'Parameter 6 : 目前的檢測工位 (決定ViewJudgerReult 更新哪一Form秀出檢測結果)
                        t2 = Environment.TickCount
                        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", [After AI Process]LogJudgeResult => " & t2 - t1)

                        If JudgeRank.CCD = "" Then
                            JudgeRank.CCD = "1"
                        End If

                        If JudgeRank.Pattern = "" Then
                            JudgeRank.Pattern = Me.m_FurtherPattern.IMG_PROC_RECIPE
                        End If
                    End If

                Catch ex As Exception
                    Throw New Exception("[Step7: AI Processing][ERR] : " & ex.ToString)
                End Try

                '可透過Thread加速處理--功能正常再修改
                '----AOI 檢測結果顯示處理----
                Try
                    Dim judgeResultPath As String = Me.m_BootRecipe.JUDGE_RESULT_PATH
                    Dim EqID As String = Me.m_BootRecipe.EQ_ID.ToUpper
                    Dim testDate As String = InspectDatetime.ToString("yyyyMMdd")
                    Dim fullJudgeResultPath As String = judgeResultPath & "\" & EqID & "\" & testDate & "\"
                    Dim AOIResultFilePath As String = String.Format("{0}{1}.txt", fullJudgeResultPath, Me.m_ProberPanelId)
                    Dim ErrMsg As String = ""
                    If Me.m_ViewerIMP.GetAOIResult(AOIResultFilePath, ErrMsg) = False Then
                        Throw New Exception(ErrMsg)
                    End If

                    'If JudgeRank.MainDefectCode <> "NO_DEFECT" Then
                    RaiseEvent MainViewShowDefectWindow(Me.m_ViewerIMP.AOICenterData, Me.m_CurrentPattern.IMG_PROC_RECIPE)
                    RaiseEvent SlideViewShowDefect(Me.m_ViewerIMP.AOICenterData, 2, Me.m_CurrentPattern.IMG_PROC_RECIPE)
                    RaiseEvent SlideViewShowDefect(Me.m_ViewerIMP.AOICenterData, 3, Me.m_CurrentPattern.IMG_PROC_RECIPE)
                Catch ex As Exception
                    SlideViewErrMsg = "[Step - AOIResult Processing][ERR] : " & ex.Message
                    'Throw New Exception("[Step - AOIResult Processing][ERR] : " & ex.ToString)
                End Try
                '----------------------------
                'End If

                '----------------------------
                '----LuminanceResult Processing 亮度量測結果處理----
                Try
                    Dim judgeResultPath As String = Me.m_BootRecipe.JUDGE_RESULT_PATH
                    Dim EqID As String = Me.m_BootRecipe.EQ_ID.ToUpper
                    Dim testDate As String = InspectDatetime.ToString("yyyyMMdd")
                    Dim fullJudgeResultPath As String = judgeResultPath & "\" & EqID & "\" & testDate & "\"

                    Dim LumResultFilePath As String = String.Format("{0}{1}_Measure.txt", fullJudgeResultPath, Me.m_ProberPanelId)
                    Dim ErrMsg As String = ""
                    If Me.m_ViewerIMP.GetLuminanceResult(LumResultFilePath, ErrMsg) = False Then
                        Throw New Exception(ErrMsg)
                    End If
                Catch ex As Exception
                    LumErrMsg = "[Step - LuminanceResult Processing][ERR] : " & ex.Message
                    'Throw New Exception("[Step - LuminanceResult Processing][ERR] : " & ex.ToString)
                End Try
                '------------------------
                '----ShowResult----
                RaiseEvent ShowResult(JudgeRank.FinalOKNG, JudgeRank.MainDefectCode, Me.m_ViewerIMP.LuminanceMeasData)
                If SlideViewErrMsg <> "" Then
                    Throw New Exception("[Step - AOIResult Processing][ERR] : " & SlideViewErrMsg)
                End If
                If LumErrMsg <> "" Then
                    Throw New Exception("[Step - LuminanceResult Processing][ERR] : " & LumErrMsg)
                End If
                '------------------



                '--- 記錄檢測結果 ---
                If Me.m_Ui.RunMode = eRunMode.Auto Or Me.m_Ui.RunMode = eRunMode.Sim Then
                    RaiseEvent StatusMsg("記錄檢測結果中(Recording inspect result)...")
                    'strTemp = IIf(Me.m_InspectStatus <> "OK" And Me.m_InspectOtherReason <> "", JudgeRank.MainDefectCode & "(" & Me.m_InspectOtherReason & ")", JudgeRank.MainDefectCode)
                    strTemp = JudgeRank.MainDefectCode
                    If strTemp Is "NO_DEFECT" And Me.m_ViewerIMP.LuminanceMeasData.Final_OKNG <> True Then
                        strTemp = "LUMINANCE"
                    End If

                    Me.InspectResultLog(Format(Now, "yyyy/MM/dd HH:mm:ss "), Me.m_ProberPanelId, JudgeRank.MqRank, JudgeRank.CstRank, JudgeRank.AGSRank, strTemp, JudgeRank.ReasonCode,
                        JudgeRank.MainDefectCode, JudgeRank.CCD, JudgeRank.Pattern, JudgeRank.CoordData, JudgeRank.CoordGate, "",
                        Format(((Environment.TickCount - Me.m_StartTime) / 1000), "0.00"))

                    RaiseEvent RefreshYieldAfterInspect()
                End If

                'Unload
                Me.m_forwardPanelId = Me.m_ProberPanelId
                'Auto-Flow
                If Me.m_BootRecipe.USE_AREAFRAME = False Then Me.m_StopInspection = True

                'S13 - For Freq-Convert
                If Me.m_Ui.USE_FREQ_CONVERT Then
                    Me.SetFreqConvertSignal(False)
                    Me.m_SystemLog.WriteLog(Now, ", [SetFreqConvertSignal][False] FREQ_CONVERT_DELAY_TIME : " & Me.m_Ui.FREQ_CONVERT_DELAY_TIME & " ms")
                    System.Threading.Thread.Sleep(Me.m_Ui.FREQ_CONVERT_DELAY_TIME)
                End If

                Try
                    ' --- Alarm Rule ---
                    If Me.m_Ui.RunMode = eRunMode.Auto Then

                        Me.iInspect_Panel_Cnt = Me.iInspect_Panel_Cnt + 1  '最高累加檢測機片數 ->32767 (controller 重開才重cnt)

                        ' Process Error Count
                        If Me.m_InspectStatus = "TFT_BREAK" Then             '連續發生
                            Me.m_AlarmRule.IncTFTBreakCount()
                        Else
                            Me.m_AlarmRule.ResetTFTBreakCount()
                        End If


                        If Me.m_InspectStatus = "OTHER_GLASS_DEFECT" Then    '連續發生
                            Me.m_AlarmRule.IncAreaOtherGlassCount()
                        Else
                            Me.m_AlarmRule.ResetAreaOtherGlassCount()
                        End If

                        If Me.m_InspectStatus = "OTHER_ALIGN_DEFECT" Then      '連續發生
                            Me.m_AlarmRule.IncAreaOtherAlignCount()
                        Else
                            Me.m_AlarmRule.ResetAreaOtherAlignCount()
                        End If


                        If Me.m_InspectStatus = "OTHER_LINE_DEFECT" Then         '連續發生
                            Me.m_AlarmRule.IncAreaOtherLineCount()
                        Else
                            Me.m_AlarmRule.ResetAreaOtherLineCount()
                        End If

                        If JudgeRank.CstRank.ToUpper.Equals("W") Then             '連續發生
                            Me.m_AlarmRule.IncCstWCount()
                        Else
                            Me.m_AlarmRule.ResetCstWCount()
                        End If

                        ' Judger False Defect DP,BP Count
                        PointCount = Me.QueryPointCount()
                        Me.m_AlarmRule.setDPBP(PointCount.DP, PointCount.BP)

                        ' Grabber False Defect Func,Mura Count
                        FalseCount = Me.QueryGrabberPointCount()
                        Me.m_AlarmRule.setGrabberFalseDefect(FalseCount.Func, FalseCount.Mura)

                        'Rank Ratio Monitor
                        Me.m_AlarmRule.EnQueueRank(JudgeRank.CstRank)

                        ' V-Line/V-Open Ratio Monitor
                        If JudgeRank.MainDefectCode = "V_LINE" Or JudgeRank.MainDefectCode = "V_OPEN" Then
                            JudgeRank.MainDefectCode = "V_LINE" '名稱統一置換
                            Me.m_AlarmRule.EnQueueVLineDefect(JudgeRank.MainDefectCode)
                        End If


                        ' H-Line/H-Open Ratio Monitor
                        If JudgeRank.MainDefectCode = "H_LINE" Or JudgeRank.MainDefectCode = "H_OPEN" Then
                            JudgeRank.MainDefectCode = "H_LINE" '名稱統一置換
                            Me.m_AlarmRule.EnQueueHLineDefect(JudgeRank.MainDefectCode)
                        End If

                        ' V-Band Ratio Monitor
                        If JudgeRank.MainDefectCode = "V_BAND_MURA" Then
                            Me.m_AlarmRule.EnQueueVBandDefect(JudgeRank.MainDefectCode)
                        End If

                        ' H-Band Ratio Monitor
                        If JudgeRank.MainDefectCode = "H_BAND_MURA" Then
                            Me.m_AlarmRule.EnQueueHBandDefect(JudgeRank.MainDefectCode)
                        End If

                        ' White Mura Ratio Monitor
                        If JudgeRank.MainDefectCode = "AROUND_GAP_MURA_WHITE" Or JudgeRank.MainDefectCode = "WHITE_MURA" Or JudgeRank.MainDefectCode = "WHITE_SPOT" Then
                            JudgeRank.MainDefectCode = "WHITE MURA" ''名稱統一置換
                            Me.m_AlarmRule.EnQueueWhiteMuraDefect(JudgeRank.MainDefectCode)
                        End If

                        ' Black Mura Ratio Monitor
                        If JudgeRank.MainDefectCode = "AROUND_GAP_MURA_BLACK" Or JudgeRank.MainDefectCode = "BLACK_MURA" Or JudgeRank.MainDefectCode = "BLACK_SPOT" Then
                            JudgeRank.MainDefectCode = "BLACK MURA" ''名稱統一置換
                            Me.m_AlarmRule.EnQueueBlackMuraDefect(JudgeRank.MainDefectCode)
                        End If

                        ' DP Kind Ratio Monitor
                        If JudgeRank.MainDefectCode = "DP_NEAR" Or JudgeRank.MainDefectCode = "DP_CLUSTER" Or JudgeRank.MainDefectCode = "DP_ADJ" Or JudgeRank.MainDefectCode = "DP_PAIR" Or JudgeRank.MainDefectCode = "DP" Then
                            JudgeRank.MainDefectCode = "DP Kind" ''名稱統一置換
                            Me.m_AlarmRule.EnQueueDPKindDefect(JudgeRank.MainDefectCode)
                        End If

                        ' BP Kind Ratio Monitor
                        If JudgeRank.MainDefectCode = "BP_NEAR" Or JudgeRank.MainDefectCode = "BP_CLUSTER" Or JudgeRank.MainDefectCode = "BP_ADJ" Or JudgeRank.MainDefectCode = "BP_PAIR" Or JudgeRank.MainDefectCode = "BP" Or JudgeRank.MainDefectCode = "CELL_PARTICLE" Or JudgeRank.MainDefectCode = "SMALL_BP" Or JudgeRank.MainDefectCode = "WEAK_BP" Or JudgeRank.MainDefectCode = "GROUP_SMALL_BP" Then
                            JudgeRank.MainDefectCode = "BP Kind" ''名稱統一置換
                            Me.m_AlarmRule.EnQueueBPKindDefect(JudgeRank.MainDefectCode)
                        End If

                        ' OGD Ratio Monitor
                        If JudgeRank.MainDefectCode = "OTHER_GLASS_DEFECT" Then
                            Me.m_AlarmRule.EnQueueOADDefect(JudgeRank.DefectCode)
                        End If

                        ' OAD Ratio Monitor
                        If JudgeRank.MainDefectCode = "OTHER_ALIGN_DEFECT" Then
                            Me.m_AlarmRule.EnQueueOGDDefect(JudgeRank.DefectCode)
                        End If

                        ' Check if matching alarm criterion
                        IsAlarm = Me.m_AlarmRule.matchAlarmCriteria

                        ' 判斷是否需要清除各Quene
                        Me.m_AlarmRule.IfResetQuene(Me.iInspect_Panel_Cnt)

                        Select Case IsAlarm
                            Case CAlarmRule.eAlarmType.TFT_Break
                                AlarmMsg = "[Prober異常]發生連續對位失敗(請確認點燈機的對位CCD)!" & vbCrLf & vbCrLf & "Prober continually failed in image alignment, please check the CCD alignment."
                            Case CAlarmRule.eAlarmType.Area_Other_Glass
                                AlarmMsg = "[點線檢測異常]影像處理發現影像連續異常(請確認影像圖檔)!" & vbCrLf & vbCrLf & "Grabber is coutinually processing unusual images. please check the images."
                            Case CAlarmRule.eAlarmType.Area_Other_Line
                                AlarmMsg = "[點線檢測異常]影像處理發生連續OTHER_LINE_DEFECT(請確認影像處理參數)!" & vbCrLf & vbCrLf & "Grabber continually OTHER_LINE in Function image processing, please check the image process recipe setting."
                            Case CAlarmRule.eAlarmType.Area_Other_Align
                                AlarmMsg = "[Mura檢測異常]影像處理發生連續對位失敗(請確認影像的四個定位點)!" & vbCrLf & vbCrLf & "Grabber continually failed in Mura image alignment, please check the four corner points of the image."
                            Case CAlarmRule.eAlarmType.R_Ratio
                                AlarmMsg = "[檢測異常]B比例過高(請清潔擴散版)!" & vbCrLf & vbCrLf & "The B ratio is too high, please clear the diffuser."
                            Case CAlarmRule.eAlarmType.W_Ratio
                                AlarmMsg = "[檢測異常]W比例過高(請確認探針與影像圖檔)!" & vbCrLf & vbCrLf & "The W ratio is too high, please check the probe and image."
                            Case CAlarmRule.eAlarmType.X_Ratio
                                AlarmMsg = "[檢測異常]X比例過高(請清潔擴散版)!" & vbCrLf & vbCrLf & "The X ratio is too high, please clear the diffuser."
                            Case CAlarmRule.eAlarmType.Y_Ratio
                                AlarmMsg = "[檢測異常]Y比例過高(請確認探針與影像圖檔)!" & vbCrLf & vbCrLf & "The Y ratio is too high, please check the probe and image."
                            Case CAlarmRule.eAlarmType.VLine_Ratio
                                AlarmMsg = "[檢測異常]V-Line or V-Open比例過高(請確認)!"
                            Case CAlarmRule.eAlarmType.HLine_Ratio
                                AlarmMsg = "[檢測異常]H-Line or H-Open比例過高(請確認)"
                            Case CAlarmRule.eAlarmType.X_Ratio
                                AlarmMsg = "[檢測異常]X比例過高(請確認)!"
                            Case CAlarmRule.eAlarmType.Y_Ratio
                                AlarmMsg = "[檢測異常]Y比例過高(請確認)!"
                            Case CAlarmRule.eAlarmType.ContiCstW
                                AlarmMsg = "[檢測異常]W發生連續(請確認)!"
                            Case CAlarmRule.eAlarmType.BP_Kind_Ratio
                                AlarmMsg = "[檢測異常]BP類Defect比例過高(請確認)"
                            Case CAlarmRule.eAlarmType.DP_Kind_Ratio
                                AlarmMsg = "[檢測異常]DP類Defect比例過高(請確認)!"
                            Case CAlarmRule.eAlarmType.VBand_Mura_Ratio
                                AlarmMsg = "[檢測異常]V Band Mura類Defect比例過高(請確認)!"
                            Case CAlarmRule.eAlarmType.HBand_Mura_Ratio
                                AlarmMsg = "[檢測異常]H Band Mura類Defect發生連續(請確認)!"
                            Case CAlarmRule.eAlarmType.White_Mura_Ratio
                                AlarmMsg = "[檢測異常]White Mura類Defect比例過高(請確認)!"
                            Case CAlarmRule.eAlarmType.Black_Mura_Ratio
                                AlarmMsg = "[檢測異常]BlackMura類Defect發生連續(請確認)!"
                            Case CAlarmRule.eAlarmType.OAD_Ratio
                                AlarmMsg = "[檢測異常]OTHER_ALIGN_DEFECT比例過高!" & vbCrLf & vbCrLf & "The OTHER_ALIGN_DEFECT ratio is too high!!"
                            Case CAlarmRule.eAlarmType.OGD_Ratio
                                AlarmMsg = "[檢測異常]OTHER_GLASS_DEFECT比例過高!" & vbCrLf & vbCrLf & "The OTHER_GLASS_DEFECT ratio is too high!!"
                        End Select
                        Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss.fff ") & "[Info]" & "10")
                        If Me.m_AlarmRule.IsBPDPFullAlarm Then AlarmMsg = "[Judger異常]False Defect超過設定值(請清潔擴散版或更換玻璃後, 再清除False Defect的資料)!" & vbCrLf & vbCrLf & "False Defect counts are over setting value, please clear the diffuser or change glass,then delete False Defect data."
                        If Me.m_AlarmRule.IsGrabberFasleDefectFullAlarm Then AlarmMsg = "[Grabber異常]False Defect超過設定值(請確認False Defect的資料)!" & vbCrLf & vbCrLf & "False Defect counts are over setting value, please check the False Defect data."
                        If Me.m_Ui.RunMode = eRunMode.Auto AndAlso (IsAlarm <> CAlarmRule.eAlarmType.OK Or Me.m_AlarmRule.IsBPDPFullAlarm Or Me.m_AlarmRule.IsGrabberFasleDefectFullAlarm) Then Throw New Exception(AlarmMsg)
                        If AlarmMsg <> "" Then Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss.fff ") & "[Info]" & AlarmMsg)
                    End If
                Catch ex As Exception
                    Me.SetAOIpcAlarm()
                    Me.m_DlgErrBox.strMsg = "[AlarmRule]  " & ex.Message
                    Me.m_DlgErrBox.strTitle = "Alarm"
                    Me.m_DlgErrBox.ShowDialog()
                    Throw New Exception("[Step8: AlarmRule] " & ex.Message)
                End Try

                'OK
                'TFT_EXSERT
                'TFT_BREAK
                'MASK_ID_REJECT
                'ImgProcProcess's Param2 and Param3
                Try
                    '--- 刪除影像 ---
                    If Me.m_Ui.blnSaveImage Then
                        t1 = Environment.TickCount

                        If JudgeRank IsNot Nothing Then
                            If (JudgeRank.CstRank = "Z" And Not Me.m_Ui.blnSaveZ) OrElse
                                                        (JudgeRank.CstRank = "N" And Not Me.m_Ui.blnSaveN) OrElse
                                                        (JudgeRank.CstRank = "S" And Not Me.m_Ui.blnSaveS) OrElse
                                                        (JudgeRank.CstRank = "H" And Not Me.m_Ui.blnSaveH) OrElse
                                                        (JudgeRank.CstRank = "G" And Not Me.m_Ui.blnSaveG) OrElse
                                                        (JudgeRank.CstRank = "B" And Not Me.m_Ui.blnSaveB) OrElse
                                                        (JudgeRank.CstRank = "W" And Not Me.m_Ui.blnSaveW) OrElse
                                                        (JudgeRank.CstRank = "R" And Not Me.m_Ui.blnSaveR) OrElse
                                                        (JudgeRank.CstRank = "J" And Not Me.m_Ui.blnSaveJ) OrElse
                                                        (JudgeRank.CstRank = "X" And Not Me.m_Ui.blnSaveX) OrElse
                                                        (Me.m_InspectStatus <> "OK" And Not Me.m_Ui.blnSaveAbnormal) Then

                                RaiseEvent StatusMsg("刪除檢測影像中(Delete image)...")

                                Me.m_AreaGrabberManager.PrepareAllRequest("DELETE", , , , , , , , , Me.m_TimeoutRecipe.Grabber.DELETE)
                                Me.m_SystemLog.WriteLog(Now, ", [Cmd=DELETE]Grabber刪除檢測影像")
                                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.DELETE)

                                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("刪除檢測影像失敗! Delete image fail! <DELETE>")
                            End If
                        End If

                        t2 = Environment.TickCount
                        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Delete Image => " & t2 - t1)
                    End If

                Catch ex As Exception
                    Throw New Exception("[Step10: DeleteImage] " & ex.Message)
                End Try

                If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Inspect time => " & CStr(Environment.TickCount - Me.m_StartTime))
                If Me.m_Ui.RunMode = eRunMode.Sim OrElse Me.m_Ui.RunMode = eRunMode.LoadImage OrElse AlarmMsg <> "" Then Me.m_StopInspection = True ' 模擬測試, 只跑一次
                RaiseEvent ChangeInspectionFlowBackColor(0)
                RaiseEvent InspectionUIChange(Me.m_StopInspection)

            End While 'Loop on Me.m_StopInspection

            Me.ChangeSubUnitAuth(Me.m_AuthUserInfo.Role, Me.m_AuthUserInfo.UserID)
            RaiseEvent StatusChange("DOWN", False)
            Me.m_SystemLog.WriteLog(Now, ", [Info]User(" & Me.m_AuthUserInfo.UserID & ") Stop Inspect!")

        Catch ex As System.Threading.ThreadAbortException
            '--- 緊急停止或設備發生錯誤 ---
            Me.m_SystemLog.WriteLog(Now, ", [Error]InspectionProcess ThreadAbortException,Panel ID= " & Me.m_ProberPanelId)
            Me.m_ErrorLog.WriteLog(Now, ", [Error]InspectionProcess ThreadAbortException,Panel ID= " & Me.m_ProberPanelId)
            Me.SetAOIpcAlarm()
            RaiseEvent ChangeInspectionFlowBackColor(0)
        Catch ex As Exception

            RaiseEvent StatusMsg("產生Alarm中...")
            Me.SetAOIpcAlarm()
            RaiseEvent StatusMsg("")

            RaiseEvent MainFormEnable(False)    ' 使用強制回應的畫面

            If Me.m_Ui.RunMode = eRunMode.Auto AndAlso (IsAlarm <> CAlarmRule.eAlarmType.OK Or Me.m_AlarmRule.IsBPDPFullAlarm Or Me.m_AlarmRule.IsGrabberFasleDefectFullAlarm) Then
                If AlarmMsg <> "" Then Me.m_SystemLog.WriteLog(Now, ", [Info]" & AlarmMsg)
                MessageBox.Show("[Kernal] [InspectionProcess-警示規則] " & AlarmMsg, "Controller Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                Me.m_SystemLog.WriteLog(Now, ", [Error]InspectionProcess Exception(Panel ID= " & Me.m_ProberPanelId & ") => " & ex.Message & ex.StackTrace)
                Me.m_ErrorLog.WriteLog(Now, ", [Error]InspectionProcess Exception(Panel ID= " & Me.m_ProberPanelId & ") => " & ex.Message & ex.StackTrace)
                MessageBox.Show("[Kernal] [InspectionProcess-檢測流程] " & ex.Message, "Controller Err", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            RaiseEvent MainFormEnable(True)

            RaiseEvent StatusMsg("取消Alarm中...")
            Me.ResetAOIpcAlarm()
            RaiseEvent StatusMsg("")
            RaiseEvent ChangeInspectionFlowBackColor(0)
            RaiseEvent InspectionUIChange(True)
            '' --- 解Alarm ---
            'If Me.m_BootRecipe.USE_AUTH And (Me.m_Ui.RunMode = eRunMode.Auto) Then
            '    If Me.m_Role <> "-2" Then

            '        RaiseEvent MainFormEnable(False)    ' 使用強制回應的畫面

            '        Me.m_Auth.OpenAuthDialog()
            '        While Not Me.m_Auth.m_AuthResult.Result

            '            MessageBox.Show("[Kernal] [InspectionProcess-登入錯誤] " & Me.m_Auth.m_AuthResult.ErrMessage, "Controller", MessageBoxButtons.OK, MessageBoxIcon.Error)

            '            If Me.m_Auth.m_AuthResult.Role = "-2" Then Exit While
            '            Me.m_Auth.OpenAuthDialog()
            '        End While

            '        RaiseEvent MainFormEnable(True)

            '        Me.m_Role = Me.m_Auth.m_AuthResult.Role
            '        Me.m_UserId = Me.m_Auth.m_AuthResult.UserId
            '        Me.m_AuthLog.WriteLog(Now, ", [Info]User(" & Me.m_UserId & ") Reset Alarm!")
            '    End If
            'End If
            'Me.ChangeSubUnitAuth(Me.m_AuthUserInfo.Role, Me.m_AuthUserInfo.UserID)

            RaiseEvent InspectionUIChange(True)
            Me.m_SystemLog.WriteLog(Now, ", [Info]User(" & Me.m_AuthUserInfo.UserID & ") Reset Alarm!")

        Finally

            'Grabber切換為auto Mode
            RaiseEvent StatusMsg("Grabber設定手動模式中(Info Grabber Setting to Manual Mode)...")

            If m_AreaGrabberManager IsNot Nothing Then
                Me.m_AreaGrabberManager.PrepareAllRequest("SET_MODE", "MANUAL", , , , , , , , Me.m_TimeoutRecipe.Grabber.SET_MODE)
                Me.m_SystemLog.WriteLog(Now, ", [Cmd=SET_MODE]Grabber Set to Manual Mode")
                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.SET_MODE)
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber Set to Manual Mode" & CStr(Me.m_AreaGrabberManager.RetryLimit) & "次失敗! Communicate with Grabber fail! <SET_MODE>")
            End If

            'AI STOP
            If Me.m_CurrentModel.UseAISystem Then
                Me.AIStopWork(Me.m_CurrentModel.UseAISystem)
            End If
            Me.m_SystemLog.WriteLog(Now, ", [Info]--- Stop Inspection Process ---")
            RaiseEvent StatusMsg("停止運作中(Stopping Inspection)...")
            RaiseEvent ChangeRole()
            RaiseEvent ChangeInspectionFlowBackColor(0)
            'Me.DisconnectSubUnit()

            If Me.GrabberPatternInfoThread IsNot Nothing Then Me.GrabberPatternInfoThread.Abort()
            If Me.JudgerPatternInfoThread IsNot Nothing Then Me.JudgerPatternInfoThread.Abort()
            If Me.ImgProcThread IsNot Nothing Then Me.ImgProcThread.Abort()

            RaiseEvent StopInspect()
            RaiseEvent StatusMsg("請按""啟動檢測""開始執行! Please push ""Start Inspection"" button to start!")
        End Try
    End Sub

    Public Sub InspectResultLog(ByVal DateTime As String, ByVal PanelId As String, ByVal MqRank As String, ByVal CstRank As String, ByVal AgsRank As String, ByVal MainDefectCode As String, ByVal ReasonCode As String, ByVal DefectCode As String, ByVal CCD As String, ByVal Pattern As String, ByVal CoordData As String, ByVal CoordGate As String, ByVal StageNo As String, ByVal TactTime As String, Optional ByVal PG_ID As String = "-1")
        Dim Coordinate As String
        Coordinate = "(" & CoordData & "," & CoordGate & ")"
        RaiseEvent InspectResult(DateTime, PanelId, MqRank, CstRank, AgsRank, MainDefectCode, ReasonCode, DefectCode, CCD, Pattern, Coordinate, StageNo, TactTime)
        Me.m_InspectLog.WriteLog(Now, vbTab & PanelId & vbTab & MqRank & vbTab & CstRank & vbTab & AgsRank & vbTab & MainDefectCode & vbTab & CCD & vbTab & Pattern & vbTab & Coordinate & vbTab & StageNo & vbTab & TactTime)

    End Sub

    Public Sub GetReadChipIDFailRatio(ByVal ReadFailCount As String)
        RaiseEvent GetChipIDRatio(ReadFailCount)
    End Sub

    Private Sub m_dlgErrBox_BuzzerStop() Handles m_DlgErrBox.BuzzerStop
        Me.ResetAOIpcAlarm()
    End Sub

    Public Sub GetFileFilterDefect(ByRef DefectStream As String, ByRef DefectStream_RB As String, ByVal ShowTrueDefect As Boolean) 'ByVal path As String, ByRef str As String)
        Dim OpenFileRetry As Integer
        Dim FileName As String = "D:\AOI_Data\Temp\DefectLog1.txt"
        'Dim DefectFileCopyPath As String = "D:\AOI_Data\GrabImage\AreaGrabber1\Stick Temp\DefectBackup\"
        Dim strs1(), str2, strs2(), strs3(), strDefect(1500) As String
        Dim OffsetX, OffsetY, DefectCountSource As Integer
        Dim DefectCount As Integer
        Dim i, NewDefectCount As Integer
        Dim str As String
        Dim sr As StreamReader = Nothing
        Dim WeakPointflag As Boolean
        Dim ppa As Double
        'Dim p As ClsMuraPosition
        Try
            'strs1 傳字串過來 , strs2 解暫時檔案 , 
            'If Not System.IO.Directory.Exists("D:\AOI_Data\GrabImage\AreaGrabber1\Stick Temp\") Then System.IO.Directory.CreateDirectory("D:\AOI_Data\GrabImage\AreaGrabber1\Stick Temp\")
            'If Not System.IO.Directory.Exists(DefectFileCopyPath) Then System.IO.Directory.CreateDirectory(DefectFileCopyPath)
            'If Not System.IO.Directory.Exists(DefectFileCopyPath & "WeakBP\" & Format(Now, "yyyyMMdd") & "\") Then System.IO.Directory.CreateDirectory(DefectFileCopyPath & "WeakBP\" & Format(Now, "yyyyMMdd") & "\")
            'If Not System.IO.Directory.Exists(DefectFileCopyPath & "BP\" & Format(Now, "yyyyMMdd") & "\") Then System.IO.Directory.CreateDirectory(DefectFileCopyPath & "BP\" & Format(Now, "yyyyMMdd") & "\")
            'If Not System.IO.Directory.Exists(DefectFileCopyPath & "DP\" & Format(Now, "yyyyMMdd") & "\") Then System.IO.Directory.CreateDirectory(DefectFileCopyPath & "DP\" & Format(Now, "yyyyMMdd") & "\")
            'Strs1():Graber Calculate  
            strs1 = DefectStream.Split(";")
            If strs1.Length = 4 Then
                DefectStream_RB = DefectStream
                Exit Sub
            End If
            i = 0
            OpenFileRetry = 0
            NewDefectCount = 0
            DefectCount = 0
            DefectStream_RB = ""
            Try
OpenFile:       sr = File.OpenText(FileName)
            Catch ex As Exception
                OpenFileRetry = OpenFileRetry + 1
                System.Threading.Thread.Sleep(300)
                If OpenFileRetry < 3 Then
                    Me.m_SystemLog.WriteLog(Now, ",[GetFileFilterDefect] Open File Retry Count " & OpenFileRetry)
                    GoTo OpenFile
                Else
                    Throw New Exception(ex.Message)
                End If
            End Try

            strs2 = sr.ReadLine().Split(":")  'OffsetX
            OffsetX = strs2(1)
            strs2 = sr.ReadLine().Split(":")  'OffsetY
            OffsetY = strs2(1)
            strs2 = sr.ReadLine().Split(":")  'Count
            DefectCountSource = strs2(1)
            sr.ReadLine() 'Title

            Do Until (sr.EndOfStream)
                strDefect(DefectCount) = sr.ReadLine()
                DefectCount = DefectCount + 1
            Loop
            sr.Close()
            If DefectCount > 0 Then
                DefectStream = ""
                DefectStream_RB = ""
            End If

            For i = 0 To DefectCount - 1
                str2 = strDefect(i)
                strs2 = str2.Split(vbTab)
                If strs2.Length < 14 Then Exit For 'last data
                ppa = Microsoft.VisualBasic.Val(strs2(16)) * Microsoft.VisualBasic.Val(strs2(16)) / Microsoft.VisualBasic.Val(strs2(3))
                If strs1.Length > 4 Then
                    strs3 = strs1(i + 4).Split(",")
                    DefectStream_RB = DefectStream_RB & i + 1 & "," & strs3(1) & "," & strs3(2) & "," & strs3(3) & "," & strs3(4) & "," & strs3(5) & "," & strs3(6) & "," & strs3(7) & "," & strs3(8) & "," & strs3(9) & "," & strs3(10) & "," & strs2(9) & "," & strs2(10) & "," & strs2(11) & "," & strs2(12) & "," & strs2(13) & "," & strs2(14) & "," & ppa & "," & strs2(16) & ";"
                Else
                    str = "D:\AOI_Data\GrabImage\AreaGrabber1\" & Format(Now, "yyyyMMdd") & "\" & Format(Now, "hh") & "\Defect\Func\," & Me.m_ProberPanelId & "_C1_PF_L0_TBP_X" & strs2(1) & "_Y" & strs2(2) & "_D" & strs2(4) & "_G" & strs2(5) & ".bmp"
                    DefectStream_RB = DefectStream_RB & i + 1 & "," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & "L0" & "," & strs2(8) & "," & str & "," & strs2(9) & "," & strs2(10) & "," & strs2(11) & "," & strs2(12) & "," & strs2(13) & "," & strs2(14) & "," & ppa & "," & strs2(16) & ";"
                End If
                'Filter Rule
                If Me.IsTrueDefectTest2(strs2, WeakPointflag) <> 0 Then
                    NewDefectCount += 1
                    If strs1.Length > 4 Then
                        DefectStream = DefectStream & NewDefectCount - 1 & "," & strs3(1) & "," & strs3(2) & "," & strs3(3) & "," & strs3(4) & "," & strs3(5) & "," & strs3(6) & "," & strs3(7) & "," & strs3(8) & "," & strs3(9) & "," & strs3(10) & ";"
                    Else
                        DefectStream = DefectStream & NewDefectCount - 1 & "," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & strs2(9) & "," & strs2(10) & ";"
                    End If
                Else
                    'Me.UI_txtLog(str & " is particle")
                End If
                'i += 1
            Next
            'If DefectCount > 0 Then
            'DefectStream = strs1(0) & ";" & strs1(1) & ";" & NewDefectCount & ";" & strs1(3) & ";" & DefectStream
            DefectStream_RB = OffsetX & ";" & OffsetY & ";" & DefectCount & ";10000;" & DefectStream_RB
            'End If
            'DefectStream_RB = DefectStream
            'sr.Close()
            'Exit Sub
            'Else
            'End If
            'If DefectStream_RB = "" Then

            '    'DefectStream = "0;0;" & NewDefectCount & ";10000;" & DefectStream
            '    DefectStream_RB = "0;0;" & i & ";10000;" & DefectStream_RB
            'End If
            'DefectStream_RB = DefectStream
            If ShowTrueDefect Then MsgBox(DefectStream)
            'sr.Close()

        Catch ex As Exception
            If sr IsNot Nothing Then sr.Close()
            Throw New Exception("[GetFileFilterDefect]" & ex.Message)
        End Try
    End Sub

    Public Sub RetRuleBaseDefectStream(ByRef DefectStream As String, ByRef DefectStream_RB As String) 'ByVal path As String, ByRef str As String)
        Dim DefectFileCopyPath As String = "D:\AOI_Data\GrabImage\IP1\Stick Temp\DefectBackup\"
        Dim strs1(), strs2() As String
        Dim DefectStreamTemp As String = ""
        Dim i As Integer

        If Not System.IO.Directory.Exists("D:\AOI_Data\GrabImage\AreaGrabber1\Stick Temp\") Then System.IO.Directory.CreateDirectory("D:\AOI_Data\GrabImage\AreaGrabber1\Stick Temp\")
        If Not System.IO.Directory.Exists(DefectFileCopyPath) Then System.IO.Directory.CreateDirectory(DefectFileCopyPath)
        If Not System.IO.Directory.Exists(DefectFileCopyPath & "WeakBP\" & Format(Now, "yyyyMMdd") & "\") Then System.IO.Directory.CreateDirectory(DefectFileCopyPath & "WeakBP\" & Format(Now, "yyyyMMdd") & "\")
        If Not System.IO.Directory.Exists(DefectFileCopyPath & "GSBP\" & Format(Now, "yyyyMMdd") & "\") Then System.IO.Directory.CreateDirectory(DefectFileCopyPath & "GSBP\" & Format(Now, "yyyyMMdd") & "\")
        If Not System.IO.Directory.Exists(DefectFileCopyPath & "BP\" & Format(Now, "yyyyMMdd") & "\") Then System.IO.Directory.CreateDirectory(DefectFileCopyPath & "BP\" & Format(Now, "yyyyMMdd") & "\")
        If Not System.IO.Directory.Exists(DefectFileCopyPath & "DP\" & Format(Now, "yyyyMMdd") & "\") Then System.IO.Directory.CreateDirectory(DefectFileCopyPath & "DP\" & Format(Now, "yyyyMMdd") & "\")
        If Not System.IO.Directory.Exists(DefectFileCopyPath & "BPDP\" & Format(Now, "yyyyMMdd") & "\") Then System.IO.Directory.CreateDirectory(DefectFileCopyPath & "BPBP\" & Format(Now, "yyyyMMdd") & "\")

        strs1 = DefectStream_RB.Split(";")
        If strs1.Length > 4 Then
            DefectStream = strs1(0) & ";" & strs1(1) & ";" & strs1(2) & ";" & strs1(3) & ";"
        End If
        For i = 0 To strs1.Length - 5
            strs2 = strs1(i + 4).Split(",")
            If strs2.Length > 10 Then

                If strs2(6) = 6 Then
                    Me.m_DefectMonitorLog.WriteLog(Now, ",CellParticle," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & strs2(11) & "," & strs2(12) & "," & strs2(13) & "," & strs2(14) & "," & strs2(15) & "," & strs2(16) & "," & strs2(17) & "," & strs2(18))
                    If File.Exists(strs2(9) & strs2(10)) Then
                        File.Move(strs2(9) & strs2(10), DefectFileCopyPath & "CellParticle\" & Format(Now, "yyyyMMdd") & "\" & strs2(10))
                        DefectStream = DefectStream & strs2(0) & "," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & DefectFileCopyPath & "CellParticle\" & Format(Now, "yyyyMMdd") & "\" & "," & strs2(10) & ";"
                    End If
                ElseIf strs2(6) = 3 Then
                    Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & strs2(11) & "," & strs2(12) & "," & strs2(13) & "," & strs2(14) & "," & strs2(15) & "," & strs2(16) & "," & strs2(17) & "," & strs2(18))
                    If File.Exists(strs2(9) & strs2(10)) Then
                        File.Move(strs2(9) & strs2(10), DefectFileCopyPath & "BP\" & Format(Now, "yyyyMMdd") & "\" & strs2(10))
                        DefectStream = DefectStream & strs2(0) & "," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & DefectFileCopyPath & "BP\" & Format(Now, "yyyyMMdd") & "\" & "," & strs2(10) & ";"
                    End If
                ElseIf strs2(6) = 5 Then
                    Me.m_DefectMonitorLog.WriteLog(Now, ",GSBP," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & strs2(11) & "," & strs2(12) & "," & strs2(13) & "," & strs2(14) & "," & strs2(15) & "," & strs2(16) & "," & strs2(17) & "," & strs2(18))
                    If File.Exists(strs2(9) & strs2(10)) Then
                        File.Move(strs2(9) & strs2(10), DefectFileCopyPath & "GSBP\" & Format(Now, "yyyyMMdd") & "\" & strs2(10))
                        DefectStream = DefectStream & strs2(0) & "," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & DefectFileCopyPath & "GSBP\" & Format(Now, "yyyyMMdd") & "\" & "," & strs2(10) & ";"
                    End If
                ElseIf strs2(6) = 1 Then
                    Me.m_DefectMonitorLog.WriteLog(Now, ",DP," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & strs2(11) & "," & strs2(12) & "," & strs2(13) & "," & strs2(14) & "," & strs2(15) & "," & strs2(16) & "," & strs2(17) & "," & strs2(18))
                    If File.Exists(strs2(9) & strs2(10)) Then
                        File.Move(strs2(9) & strs2(10), DefectFileCopyPath & "DP\" & Format(Now, "yyyyMMdd") & "\" & strs2(10))
                        DefectStream = DefectStream & strs2(0) & "," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & DefectFileCopyPath & "DP\" & Format(Now, "yyyyMMdd") & "\" & "," & strs2(10) & ";"
                    End If
                ElseIf strs2(6) = 7 Then
                    Me.m_DefectMonitorLog.WriteLog(Now, ",BPDP," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & strs2(11) & "," & strs2(12) & "," & strs2(13) & "," & strs2(14) & "," & strs2(15) & "," & strs2(16) & "," & strs2(17) & "," & strs2(18))
                    If File.Exists(strs2(9) & strs2(10)) Then
                        File.Move(strs2(9) & strs2(10), DefectFileCopyPath & "BPDP\" & Format(Now, "yyyyMMdd") & "\" & strs2(10))
                        DefectStream = DefectStream & strs2(0) & "," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & DefectFileCopyPath & "BPDP\" & Format(Now, "yyyyMMdd") & "\" & "," & strs2(10) & ";"
                    End If
                ElseIf strs2(6) = 8 Then
                    Me.m_DefectMonitorLog.WriteLog(Now, ",BPDP," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & strs2(11) & "," & strs2(12) & "," & strs2(13) & "," & strs2(14) & "," & strs2(15) & "," & strs2(16) & "," & strs2(17) & "," & strs2(18))
                    If File.Exists(strs2(9) & strs2(10)) Then
                        File.Move(strs2(9) & strs2(10), DefectFileCopyPath & "BPDP\" & Format(Now, "yyyyMMdd") & "\" & strs2(10))
                        DefectStream = DefectStream & strs2(0) & "," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & DefectFileCopyPath & "BPDP\" & Format(Now, "yyyyMMdd") & "\" & "," & strs2(10) & ";"
                    End If
                Else
                    Me.m_DefectMonitorLog.WriteLog(Now, ",Line," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & strs2(11) & "," & strs2(12) & "," & strs2(13) & "," & strs2(14) & "," & strs2(15) & "," & strs2(16) & "," & strs2(17) & "," & strs2(18))
                    DefectStream = DefectStream & strs2(0) & "," & strs2(1) & "," & strs2(2) & "," & strs2(3) & "," & strs2(4) & "," & strs2(5) & "," & strs2(6) & "," & strs2(7) & "," & strs2(8) & "," & strs2(9) & "," & strs2(10) & ";"
                End If

            End If
        Next

    End Sub

    Private Function IsTrueDefectTest2(ByVal strs() As String, ByRef WeakPoint As Boolean) As Integer
        Dim sr As StreamReader = Nothing
        Dim ppa As Double
        Try

            'Return 1
            ppa = Microsoft.VisualBasic.Val(strs(16)) * Microsoft.VisualBasic.Val(strs(16)) / Microsoft.VisualBasic.Val(strs(3))
            If strs(6) = 3 Then
                WeakPoint = False
                If Microsoft.VisualBasic.Val(strs(10)) >= 3800 And 16 > ppa And ppa > 14 Then '實亮點
                    If 115 >= Microsoft.VisualBasic.Val(strs(3)) And Microsoft.VisualBasic.Val(strs(3)) > 60 Then
                        '小顆
                        If Microsoft.VisualBasic.Val(strs(14)) = 0 Then
                            'Elogration
                            If 1.1 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.8 Then
                                'Fullness
                                If 0.98 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.75 Then
                                    'compactness
                                    If 1.3 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.13 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Math.Abs(Microsoft.VisualBasic.Val(strs(14))) = 45 Then
                            'Elogration
                            If 1.15 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.95 Then
                                'Fullness
                                If 0.85 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.7 Then
                                    'compactness
                                    If 1.35 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Math.Abs(Microsoft.VisualBasic.Val(strs(14))) = 67.5 Then
                            'Elogration
                            If 1.15 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.95 Then
                                'Fullness
                                If 0.85 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.7 Then
                                    'compactness
                                    If 1.35 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.2 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Microsoft.VisualBasic.Val(strs(14)) = 90 Then
                            'Elogration
                            If 1.45 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 1 Then
                                'Fullness
                                If 0.95 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.7 Then
                                    'compactness
                                    If 2.1 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        Else
                            Return 0
                        End If

                    ElseIf 200 >= Microsoft.VisualBasic.Val(strs(3)) And Microsoft.VisualBasic.Val(strs(3)) > 115 Then
                        '2連點
                        If Microsoft.VisualBasic.Val(strs(14)) = 0 Then
                            'Elogration
                            If 1.1 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.45 Then
                                'Fullness
                                If 0.98 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.7 Then
                                    'compactness
                                    If 1.65 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Math.Abs(Microsoft.VisualBasic.Val(strs(14))) = 45 Then
                            'Elogration
                            If 1.15 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.9 Then
                                'Fullness
                                If 0.8 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.4 Then
                                    'compactness
                                    If 2.6 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Math.Abs(Microsoft.VisualBasic.Val(strs(14))) = 67.5 Then
                            'Elogration
                            If 1.5 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 1 Then
                                'Fullness
                                If 0.85 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.65 Then
                                    'compactness
                                    If 1.55 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then

                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Microsoft.VisualBasic.Val(strs(14)) = 90 Then
                            'Elogration
                            If 2.4 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 1 Then
                                'Fullness
                                If 0.95 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.52 Then
                                    'compactness
                                    If 2.2 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        Else
                            Return 0
                        End If

                    ElseIf 300 >= Microsoft.VisualBasic.Val(strs(3)) And Microsoft.VisualBasic.Val(strs(3)) > 200 Then
                        '3連點
                        If Microsoft.VisualBasic.Val(strs(14)) = 0 Then
                            'Return 0
                            'Elogration
                            If 1.1 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.8 Then
                                'Fullness
                                If 0.9 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.6 Then
                                    'compactness
                                    If 3 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Math.Abs(Microsoft.VisualBasic.Val(strs(14))) = 45 Then
                            'Return 0
                            'Elogration
                            If 1.15 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.9 Then
                                'Fullness
                                If 0.88 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.68 Then
                                    'compactness
                                    If 1.5 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Math.Abs(Microsoft.VisualBasic.Val(strs(14))) = 67.5 Then
                            'Elogration
                            If 1.5 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.9 Then
                                'Fullness
                                If 0.88 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.67 Then
                                    'compactness
                                    If 1.4 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Microsoft.VisualBasic.Val(strs(14)) = 90 Then
                            'Elogration
                            If 1.8 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 1 Then
                                'Fullness
                                If 0.95 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.7 Then
                                    'compactness
                                    If 1.5 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        Else
                            Return 0
                        End If

                    ElseIf Microsoft.VisualBasic.Val(strs(3)) > 300 Then
                        If Microsoft.VisualBasic.Val(strs(14)) = 0 Then
                            'Return 0
                            'Elogration
                            If 1.1 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.8 Then
                                'Fullness
                                If 0.9 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.6 Then
                                    'compactness
                                    If 3 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Math.Abs(Microsoft.VisualBasic.Val(strs(14))) = 45 Then
                            'Return 0
                            'Elogration
                            If 1.15 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.9 Then
                                'Fullness
                                If 0.88 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.68 Then
                                    'compactness
                                    If 1.5 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Math.Abs(Microsoft.VisualBasic.Val(strs(14))) = 67.5 Then
                            'Elogration
                            If 1.5 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.9 Then
                                'Fullness
                                If 0.88 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.67 Then
                                    'compactness
                                    If 1.4 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        ElseIf Microsoft.VisualBasic.Val(strs(14)) = 90 Then
                            'Elogration
                            If 1.8 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 1 Then
                                'Fullness
                                If 0.95 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) > 0.7 Then
                                    'compactness
                                    If 1.5 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.1 Then
                                        Me.m_DefectMonitorLog.WriteLog(Now, ",BP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                        Return 1
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        Else
                            Return 0
                        End If

                    End If
                Else '淡點
                    '9/13 Area Max 110 > 130 > 210
                    '9/13 Add FeretAngle
                    Return 0
                    If 210 >= Microsoft.VisualBasic.Val(strs(3)) And Microsoft.VisualBasic.Val(strs(3)) > 35 And (Microsoft.VisualBasic.Val(strs(14)) = 0 Or Microsoft.VisualBasic.Val(strs(14)) = 67.5 Or Microsoft.VisualBasic.Val(strs(14)) = 90) Then
                        '小顆
                        'Elogration  9/13 max 1.26 > 2.14 ,  
                        If 2.14 >= Microsoft.VisualBasic.Val(strs(11)) And Microsoft.VisualBasic.Val(strs(11)) > 0.99 Then
                            'Fullness 9/13 max .93 > .94 , Min .75 >.74 
                            If 0.94 >= Microsoft.VisualBasic.Val(strs(12)) And Microsoft.VisualBasic.Val(strs(12)) >= 0.74 Then
                                'compactness 9/13 max 1.34 > 1.37 > 1.49
                                If 1.49 >= Microsoft.VisualBasic.Val(strs(13)) And Microsoft.VisualBasic.Val(strs(13)) > 1.16 Then
                                    If Microsoft.VisualBasic.Val(strs(10)) > 1800 Then
                                        '9/17 max/mean fix 1.95  
                                        If Microsoft.VisualBasic.Val(strs(10)) / Microsoft.VisualBasic.Val(strs(8)) > 1.51 Then
                                            WeakPoint = True
                                            Me.m_DefectMonitorLog.WriteLog(Now, ",WeakBP," & strs(1) & "," & strs(2) & "," & strs(3) & "," & strs(4) & "," & strs(5) & "," & strs(6) & "," & strs(7) & "," & strs(8) & "," & strs(9) & "," & strs(10) & "," & strs(11) & "," & strs(12) & "," & strs(13) & "," & strs(14) & "," & strs(15) & "," & strs(16))
                                            Return 2
                                        Else : Return 0
                                        End If
                                    Else : Return 0
                                    End If
                                Else : Return 0
                                End If
                            Else : Return 0
                            End If
                        Else : Return 0
                        End If
                    End If

                End If
            Else
                Return 3
            End If
            Return 0
            'Return 0 False Defect ,  1 True Defect , 2 Review ,  3 Line or Func
        Catch ex As Exception
            Return 0
        End Try

    End Function

    Private Sub FlowSeq1stPartThread()
        Try
            ' 非同步Reset Judger
            Me.m_SystemLog.WriteLog(Now, ", [SendCommand Juddger  = CLEAR  & SET_DATE & PANEL_ID")
            ResetJudgerThread = New System.Threading.Thread(AddressOf Me.ResetJudgerProcess)
            ResetJudgerThread.Name = "ResetJudger"
            ResetJudgerThread.SetApartmentState(Threading.ApartmentState.STA)
            ResetJudgerThread.Start()

            ' 非同步Reset Grabber 
            Me.m_SystemLog.WriteLog(Now, ", [SendCommand Grabber  = CLEAR  & SET_DATE & PANEL_ID")
            ResetGrabberThread = New System.Threading.Thread(AddressOf Me.ResetGrabberProcess)
            ResetGrabberThread.Name = "ResetGrabber"
            ResetGrabberThread.SetApartmentState(Threading.ApartmentState.STA)
            ResetGrabberThread.Start()

            '同步Reset Judger
            If Not ResetJudgerThread.Join(Me.m_TimeoutRecipe.Judger.CLEAR + 6000) Then
                Me.ResetJudgerThread.Abort()
                Throw New Exception("Judger執行Reset命令逾時")
            Else
                If Me.m_ResetJudgerExMsg <> "" Then Throw New Exception("執行ResetJudger發生錯誤! " & Me.m_ResetJudgerExMsg)
            End If
            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Async ResetJudger,TT => " & Me.intSpeed_ResetJudger)

            '同步Reset Grabber
            If Not ResetGrabberThread.Join(Me.m_TimeoutRecipe.Grabber.CLEAR + 1000) Then
                Me.ResetGrabberThread.Abort()
                Throw New Exception("Grabber執行Reset命令逾時")
            Else
                If Me.m_ResetGrabberExMsg <> "" Then Throw New Exception("執行ResetGrabber發生錯誤! " & Me.m_ResetGrabberExMsg)
            End If
            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Async ResetGrabber,TT => " & Me.intSpeed_ResetGrabber)

        Catch ex As Exception
            Throw New Exception("FlowSeq1stPartThread " & ex.Message)
        End Try
    End Sub

#Region "非同步Function"

    Private Sub FlowSeq1stThread()
        Dim strTemp As String
        Dim i As Integer = 0

        Try
            '第一波非同步 第一個Pattern

            If Me.m_Ui.RunMode <> eRunMode.LoadImage Then


                If Me.m_BootRecipe.GENERATION <> "M11_SKD_INLINE" Or Me.m_BootRecipe.GENERATION <> "S02_PT" Then '在前面的流程做掉,此處不執行
                    ' 非同步Reset Judger
                    Me.m_SystemLog.WriteLog(Now, ", [SendCommand Juddger  = CLEAR  & SET_DATE & PANEL_ID")
                    ResetJudgerThread = New System.Threading.Thread(AddressOf Me.ResetJudgerProcess)
                    ResetJudgerThread.Name = "ResetJudger"
                    ResetJudgerThread.SetApartmentState(Threading.ApartmentState.STA)
                    ResetJudgerThread.Start()

                    ' 非同步Reset Grabber 
                    Me.m_SystemLog.WriteLog(Now, ", [SendCommand Grabber  = CLEAR  & SET_DATE & PANEL_ID")
                    ResetGrabberThread = New System.Threading.Thread(AddressOf Me.ResetGrabberProcess)
                    ResetGrabberThread.Name = "ResetGrabber"
                    ResetGrabberThread.SetApartmentState(Threading.ApartmentState.STA)
                    ResetGrabberThread.Start()
                End If

                Me.m_SystemLog.WriteLog(Now, ", [SendCommand=SET_PATTERN : " & Me.m_CurrentPattern.PATTERN_NAME)

            End If

            If Me.m_BootRecipe.GENERATION <> "M11_SKD_INLINE" Or Me.m_BootRecipe.GENERATION <> "S02_PT" Then
                '同步Reset Judger
                If Not ResetJudgerThread.Join(Me.m_TimeoutRecipe.Judger.CLEAR + 6000) Then
                    Me.ResetJudgerThread.Abort()
                    Throw New Exception("Judger執行Reset命令逾時")
                Else
                    If Me.m_ResetJudgerExMsg <> "" Then Throw New Exception("執行ResetJudger發生錯誤! " & Me.m_ResetJudgerExMsg)
                End If
                If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Async ResetJudger,TT => " & Me.intSpeed_ResetJudger)

                '同步Reset Grabber
                If Not ResetGrabberThread.Join(Me.m_TimeoutRecipe.Grabber.CLEAR + 1000) Then
                    Me.ResetGrabberThread.Abort()
                    Throw New Exception("Grabber執行Reset命令逾時")
                Else
                    If Me.m_ResetGrabberExMsg <> "" Then Throw New Exception("執行ResetGrabber發生錯誤! " & Me.m_ResetGrabberExMsg)
                End If
                If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Async ResetGrabber,TT => " & Me.intSpeed_ResetGrabber)
            End If

            ' Modify for IP need to Know Inline or offline Mode
            ' 非同步GrabberSetEXP
            If (Me.m_Ui.blnGrabImage Or Me.m_Ui.RunMode = eRunMode.LoadImage) Then
                If Me.m_CurrentPattern.LOAD_DEMURA_IMAGE = False Then
                    Me.m_SystemLog.WriteLog(Now, ", [SendCommand=SET_EXP")
                    GrabberSetExpThread = New System.Threading.Thread(AddressOf Me.GrabberSetExpProcess)
                    GrabberSetExpThread.Name = "GrabberSetExp"
                    GrabberSetExpThread.SetApartmentState(Threading.ApartmentState.STA)
                    GrabberSetExpThread.Start()
                End If
            End If

            If Me.m_Ui.RunMode <> eRunMode.LoadImage Then

                ' Modify for IP need to Know Inline or offline Mode
                '同步Grabber Exp
                If Me.m_CurrentPattern.LOAD_DEMURA_IMAGE = False Then
                    If Not GrabberSetExpThread.Join(Me.m_TimeoutRecipe.Grabber.PATTERN_INFO + 1000) Then
                        Me.GrabberSetExpThread.Abort()
                        Throw New Exception("Seq=" & Me.m_CurrentModel.Pattern(m_CurrentPatternIdx).PATTERN_NAME & " Grabber執行命令逾時(" & CStr(Me.m_TimeoutRecipe.Grabber.PATTERN_INFO) & "ms)! " & Me.m_GrabberSetExpExMsg)
                    Else
                        If Me.m_GrabberSetExpExMsg <> "" Then Throw New Exception("Seq=" & Me.m_CurrentModel.Pattern(m_CurrentPatternIdx).PATTERN_NAME & " Grabber執行命令發生錯誤! " & Me.m_GrabberSetExpExMsg)
                    End If
                    strTemp = IIf((Me.m_BootRecipe.GENERATION = "S03_CT1" Or Me.m_BootRecipe.GENERATION = "S16" Or Me.m_BootRecipe.GENERATION = "L5B" Or Me.m_BootRecipe.GENERATION = "L6B" Or Me.m_BootRecipe.GENERATION = "L5C" Or Me.m_BootRecipe.GENERATION = "L6A" Or Me.m_BootRecipe.GENERATION = "L7B" Or Me.m_BootRecipe.GENERATION = "L7A"), "Async Grabber SetExp,Seq=", "Mura檢測設定Pattern資訊,Seq=")
                    If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", " & strTemp & Me.m_CurrentModel.Pattern(m_CurrentPatternIdx).PATTERN_NAME & " => " & Me.intSpeed_GrabberSetExp)
                End If
            End If

            ' Modify for IP need to Know Inline or offline Mode
            '同步Grabber Exp
            If Me.m_CurrentPattern.LOAD_DEMURA_IMAGE = False Then
                If Not GrabberSetExpThread.Join(Me.m_TimeoutRecipe.Grabber.PATTERN_INFO + 1000) Then
                    Me.GrabberSetExpThread.Abort()
                    Throw New Exception("Seq=" & Me.m_CurrentPattern.PATTERN_NAME & " Grabber執行命令逾時(" & CStr(Me.m_TimeoutRecipe.Grabber.PATTERN_INFO) & "ms)! " & Me.m_GrabberSetExpExMsg)
                Else
                    If Me.m_GrabberSetExpExMsg <> "" Then Throw New Exception("Seq=" & Me.m_CurrentPattern.PATTERN_NAME & " Grabber執行命令發生錯誤! " & Me.m_GrabberSetExpExMsg)
                End If
                strTemp = IIf((Me.m_BootRecipe.GENERATION = "S03_CT1" Or Me.m_BootRecipe.GENERATION = "S16" Or Me.m_BootRecipe.GENERATION = "L5B" Or Me.m_BootRecipe.GENERATION = "L6B" Or Me.m_BootRecipe.GENERATION = "L5C" Or Me.m_BootRecipe.GENERATION = "L6A" Or Me.m_BootRecipe.GENERATION = "L7B" Or Me.m_BootRecipe.GENERATION = "L7A"), "Async Grabber SetExp,Seq=", "Mura檢測設定Pattern資訊,Seq=")
                If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, "Async GrabberSetExposure time ,Seq= " & strTemp & Me.m_CurrentPattern.PATTERN_NAME & " => " & Me.intSpeed_GrabberSetExp)
            End If
        Catch ex As Exception
            Throw New Exception("FlowSeq1stThread " & ex.Message)
        End Try
    End Sub

    Private Sub FlowThread(ByVal i As Integer)
        Dim strTemp As String

        Try
            ' 非同步影像處理 
            If Me.m_Ui.blnProcessImage Then
                'Defect裡面1~N隻CCD(1~N片Panel)的資訊用 @ 隔開
                Me.m_SystemLog.WriteLog(Now, ", [SendCommand=PATTERN_INFO,ALIGN,CALCULATE to Singal-Grabber ")
                ImgProcThread = New System.Threading.Thread(AddressOf Me.ImgProcProcess)
                ImgProcThread.Name = "ImgProc"
                ImgProcThread.SetApartmentState(Threading.ApartmentState.STA)
                ImgProcThread.Start()
            End If


            If Me.m_Ui.RunMode <> eRunMode.LoadImage Then
                If Not i >= Me.m_CurrentModel.PatternCount - 1 Then
                    RaiseEvent StatusMsg("切換Pattern中(Changing Pattern)...............")
                    Me.m_SystemLog.WriteLog(Now, "切換Pattern中(Changing Pattern)............... ")
                    RaiseEvent ChangeInspectionFlowBackColor(i + 1 + 1)

                    If m_CurrentPattern.SIDE_LIGHT.ToUpper = "TRUE" Then
                        m_Plc.WriteDeviceBlock("D7101", 3, New Integer() {1, 0, 1})
                    End If

                    If Me.m_CurrentPattern.LOAD_DEMURA_IMAGE = False Then
                        ' 非同步GrabberSetEXP
                        If Me.m_Ui.blnGrabImage Then
                            '20150610 加入Func+Mura(同一個Pattern)流程簡化,縮短TT
                            'Func(Pattern) Recipe 才需設定exposure time,接下來的Mura Pattern(Recipe) bypass
                            If Me.m_CurrentPattern.FUNC <> ClsPattern.eFunctionType.FUNC_MURA Then
                                Me.m_SystemLog.WriteLog(Now, ", [SendCommand=SET_EXP  " & "Current Recipe:" & Me.m_CurrentPattern.IMG_PROC_RECIPE)
                                GrabberSetExpThread = New System.Threading.Thread(AddressOf Me.GrabberSetExpProcess)
                                GrabberSetExpThread.Name = "GrabberSetExp"
                                GrabberSetExpThread.SetApartmentState(Threading.ApartmentState.STA)
                                GrabberSetExpThread.Start()
                            End If
                        End If

                        '同步 GrabberSetExp
                        '20150610 加入Func+Mura(同一個Pattern)流程簡化,縮短TT
                        'Func(Pattern) Recipe 才需設定exposure time,接下來的Mura Pattern(Recipe) bypass
                        If Me.m_CurrentPattern.FUNC <> ClsPattern.eFunctionType.FUNC_MURA Then
                            If Not GrabberSetExpThread.Join(Me.m_TimeoutRecipe.Grabber.PATTERN_INFO + 1000) Then
                                Me.GrabberSetExpThread.Abort()
                                Throw New Exception("Seq=" & Me.m_CurrentModel.Pattern(m_CurrentPatternIdx).PATTERN_NAME & " Grabber執行命令逾時(" & CStr(Me.m_TimeoutRecipe.Grabber.PATTERN_INFO) & "ms)! " & Me.m_GrabberSetExpExMsg)
                            Else
                                If Me.m_GrabberSetExpExMsg <> "" Then Throw New Exception("Seq=" & Me.m_CurrentModel.Pattern(m_CurrentPatternIdx).PATTERN_NAME & " Grabber執行命令發生錯誤! " & Me.m_GrabberSetExpExMsg)
                            End If
                            strTemp = IIf((Me.m_BootRecipe.GENERATION = "S03_CT1" Or Me.m_BootRecipe.GENERATION = "S16" Or Me.m_BootRecipe.GENERATION = "L5B" Or Me.m_BootRecipe.GENERATION = "L6B" Or Me.m_BootRecipe.GENERATION = "L5C" Or Me.m_BootRecipe.GENERATION = "L6A" Or Me.m_BootRecipe.GENERATION = "L7B"), "Async Grabber SetExp,Seq=", "Mura檢測設定Pattern資訊,Seq=")
                            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", " & strTemp & Me.m_CurrentModel.Pattern(m_CurrentPatternIdx).PATTERN_NAME & " => " & Me.intSpeed_GrabberSetExp)
                        End If

                    End If 'If Me.m_CurrentPattern.LOAD_DEMURA_IMAGE = False


                    ' 非同步Grabber擷取影像
                    If Me.m_Ui.blnGrabImage Then
                        If Me.m_CurrentPattern.LOAD_DEMURA_IMAGE = False Then
                            '' 非同步SetPattern 
                            '20150610 加入Func+Mura(同一個Pattern)流程簡化,縮短TT
                            'Func(Pattern) Recipe 才需設定取像,接下來的Mura Pattern(Recipe) bypass
                            If Me.m_CurrentPattern.FUNC <> ClsPattern.eFunctionType.FUNC_MURA Then

                                Me.m_SystemLog.WriteLog(Now, ", [SendCommand=GRAB")
                                GrabberGrabThread = New System.Threading.Thread(AddressOf Me.GrabberGrabProcess)
                                GrabberGrabThread.Name = "GrabberGrab"
                                GrabberGrabThread.SetApartmentState(Threading.ApartmentState.STA)
                                GrabberGrabThread.Start()

                                '同步Grabber擷取影像
                                If Not GrabberGrabThread.Join(Me.m_TimeoutRecipe.Grabber.GRAB + 5000) Then
                                    Me.GrabberGrabThread.Abort()
                                    Throw New Exception("Seq=" & CStr(i + 1) & "+1 Grabber執行命令逾時(" & CStr(Me.m_TimeoutRecipe.Grabber.GRAB) & "ms)! " & Me.m_GrabberGrabExMsg)
                                Else
                                    If Me.m_GrabberGrabExMsg <> "" Then Throw New Exception("Seq=" & CStr(i + 1) & "+1 Grabber執行命令發生錯誤! " & Me.m_GrabberGrabExMsg)
                                End If
                                If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Async Grab,Seq=" & CStr(i + 1) & "+1 => " & Me.intSpeed_GrabberGrab)
                            Else
                                Me.m_CurrentPatternIdx = Me.m_CurrentPatternIdx + 1
                            End If
                        End If
                    End If

                End If
            End If

            ' 同步 ImgProc 
            If Me.m_Ui.blnProcessImage Then
                If Not ImgProcThread.Join(Me.m_TimeoutRecipe.Grabber.ALIGN + Me.m_TimeoutRecipe.Grabber.CALCULATE + 1000) Then
                    Me.ImgProcThread.Abort()
                    Throw New Exception("Seq=" & CStr(i + 1) & " Grabber執行命令逾時(" & CStr(Me.m_TimeoutRecipe.Grabber.ALIGN + Me.m_TimeoutRecipe.Grabber.CALCULATE) & "ms)! " & Me.m_ImgProcExMsg)
                Else
                    If Me.m_ImgProcExMsg <> "" Then Throw New Exception("Seq=" & CStr(i + 1) & " Grabber執行命令發生錯誤! " & Me.m_ImgProcExMsg)
                End If
                If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Async ImgProcProcess,Seq=" & CStr(i + 1) & "=> " & Me.intSpeed_ImgProcProcess)
            End If

        Catch ex As Exception
            Throw New Exception("FlowThread " & ex.ToString)
        End Try
    End Sub

    Public Sub ShowGrabber()
        Dim t1 As Integer
        Dim t2 As Integer
        Dim SubSystemResult As CResponseResult
        Try
            Me.m_GrabberSetExpExMsg = ""

            t1 = Environment.TickCount

            Me.m_AreaGrabberManager.PrepareAllRequest("SHOW_GRABBER", , , , , , , , , 5000)
            SubSystemResult = Me.m_AreaGrabberManager.SendRequest(5000)
            Me.m_SystemLog.WriteLog(Now, ", [Cmd=SHOW_GRABBER]Grabber视窗显示")

            t2 = Environment.TickCount
            Me.intSpeed_GrabberSetExp = t2 - t1
            '------------------------------
        Catch ex As Exception
            Me.m_GrabberSetExpExMsg = "[GrabberShowGrabber] " & ex.Message
        End Try
    End Sub
    Public Sub HideGrabber()
        Dim t1 As Integer
        Dim t2 As Integer
        Dim SubSystemResult As CResponseResult
        Try
            Me.m_GrabberSetExpExMsg = ""

            t1 = Environment.TickCount

            Me.m_AreaGrabberManager.PrepareAllRequest("HIDE_GRABBER", , , , , , , , , 5000)
            SubSystemResult = Me.m_AreaGrabberManager.SendRequest(5000)
            Me.m_SystemLog.WriteLog(Now, ", [Cmd=HIDE_GRABBER]Grabber视窗隐藏")

            t2 = Environment.TickCount
            Me.intSpeed_GrabberSetExp = t2 - t1
            '------------------------------
        Catch ex As Exception
            Me.m_GrabberSetExpExMsg = "[GrabberShowGrabber] " & ex.Message
        End Try
    End Sub

    Public Sub ShowJudger()
        Dim t1 As Integer
        Dim t2 As Integer
        Dim SubSystemResult As CResponseResult
        Try
            Me.m_GrabberSetExpExMsg = ""

            t1 = Environment.TickCount

            RaiseEvent StatusMsg("Judger视窗显示")

            Me.m_JudgerManager.PrepareAllRequest("SHOW_Judger")
            SubSystemResult = Me.m_JudgerManager.SendRequest(5000)

            t2 = Environment.TickCount
            Me.intSpeed_GrabberSetExp = t2 - t1
            '------------------------------
        Catch ex As Exception
            Me.m_GrabberSetExpExMsg = "[GrabberShowGrabber] " & ex.Message
        End Try
    End Sub
    Public Sub HideJudger()
        Dim t1 As Integer
        Dim t2 As Integer
        Dim SubSystemResult As CResponseResult
        Try
            Me.m_GrabberSetExpExMsg = ""

            t1 = Environment.TickCount

            RaiseEvent StatusMsg("Judger视窗隐藏")

            Me.m_JudgerManager.PrepareAllRequest("HIDE_Judger", , , , , , , , , 5000)
            SubSystemResult = Me.m_JudgerManager.SendRequest(5000)

            t2 = Environment.TickCount
            Me.intSpeed_GrabberSetExp = t2 - t1
            '------------------------------
        Catch ex As Exception
            Me.m_GrabberSetExpExMsg = "[GrabberShowGrabber] " & ex.Message
        End Try
    End Sub

    Private Function GetPatternNo(ByVal patternary() As String, ByVal patternname As String) As Integer
        Dim i As Integer
        For i = 0 To patternary.Length - 1
            If patternary(i) = patternname Then
                Return i + 1
            End If
        Next
        Return 0
    End Function

    Private Sub GrabberSetExpProcess()
        Dim t1 As Integer
        Dim t2 As Integer
        Dim SubSystemResult As CResponseResult

        Try
            Me.m_GrabberSetExpExMsg = ""
            '--- 設定Pattern資訊 ---
            RaiseEvent StatusMsg("Grabber設定Pattern曝光資訊中(Info Grabber setting Pattern Exposure time)...")

            t1 = Environment.TickCount

            Me.m_AreaGrabberManager.PrepareAllRequest("SET_EXP", Me.m_CurrentPattern.IMG_PROC_RECIPE, , , , , , , , Me.m_TimeoutRecipe.Grabber.PATTERN_INFO)
            SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.PATTERN_INFO)
            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber設定Pattern曝光資訊失敗 [SET_EXP]！Communicate with Grabber failed [SET_EXP]" & SubSystemResult.ErrMessage)
            t2 = Environment.TickCount
            Me.intSpeed_GrabberSetExp = t2 - t1
            '------------------------------
        Catch ex As Exception
            Me.m_GrabberSetExpExMsg = "[GrabberSetExpProcess] " & ex.Message
        End Try
    End Sub

    Private Function UseGrabberGetChipID(ByVal PanelID As String) As CResponseResult

        Dim t1 As Integer
        Dim t2 As Integer
        Dim SubSystemResult As CResponseResult
        Dim inspectDatetime As Date = Now
        'Grabber_Clear
        t1 = Environment.TickCount
        Me.m_AreaGrabberManager.PrepareAllRequest("CLEAR", , , , , , , , , Me.m_TimeoutRecipe.Grabber.CLEAR)
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=CLEAR]Grabber檢測重置")
        SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.CLEAR)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber檢測連續重置" & CStr(Me.m_AreaGrabberManager.RetryLimit) & "次失敗! Communicate with Grabber fail! <CLEAR>")
        t2 = Environment.TickCount
        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Grabber_Clear" & "=> " & t2 - t1)
        'Grabber_SET_DATE
        t1 = Environment.TickCount
        Me.m_AreaGrabberManager.PrepareAllRequest("SET_DATE", Format(inspectDatetime, "yyyy"), Format(inspectDatetime, "MM"), Format(inspectDatetime, "dd"), Format(inspectDatetime, "HH"), Format(inspectDatetime, "mm"), Format(inspectDatetime, "ss"), , , Me.m_TimeoutRecipe.Grabber.SET_DATE)
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=SET_DATE]Grabber檢測設定檢測日期")
        SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.SET_DATE)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber檢測設定檢測日期失敗! Communicate with Grabber fail! <SET_DATE>")
        t2 = Environment.TickCount
        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Grabber_SET_DATE" & "=> " & t2 - t1)
        '非同步GrabberSetEXP()
        If (Me.m_Ui.blnGrabImage Or Me.m_Ui.RunMode = eRunMode.LoadImage) Then
            'If Me.m_CurrentPattern.LOAD_DEMURA_IMAGE = False Then
            '    Me.m_SystemLog.WriteLog(Now, ", [SendCommand=SET_EXP")
            '    GrabberSetExpThread = New System.Threading.Thread(AddressOf Me.GrabberSetExpProcess)
            '    GrabberSetExpThread.Name = "GrabberSetExp"
            '    GrabberSetExpThread.SetApartmentState(Threading.ApartmentState.STA)
            '    GrabberSetExpThread.Start()
            '    GrabberSetExpThread.Join()
            'End If
            t1 = Environment.TickCount
            GrabberSetExpProcess()
            t2 = Environment.TickCount
            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", " & Me.m_CurrentPattern.PATTERN_NAME & ": GrabberSetEXP" & "=> " & t2 - t1)
        End If
        'SetPattern Delay(确保背光已开启)
        Threading.Thread.Sleep(CInt(Me.m_CurrentPattern.DELAY_TIME))
        '拍照(獲取Chip_id)
        t1 = Environment.TickCount
        Me.m_AreaGrabberManager.PrepareAllRequest("GRAB", , , , , , , , , Me.m_TimeoutRecipe.Grabber.GRAB)
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=GRAB]Grabber擷取圖像中（獲取Chip_id）")
        SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.GRAB)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then
            Me.m_GrabberGrabExMsg = ("Grabber擷取 影像失敗!")
        End If
        t2 = Environment.TickCount
        Me.intSpeed_GrabberGrab = t2 - t1
        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Grabber 擷取第1個patten(GraGetChipID) ..." & "=> " & t2 - t1)
        'Save Image
        If Me.m_Ui.blnSaveImage Then
            t1 = Environment.TickCount
            If Me.m_CurrentPattern.SAVE Then
                RaiseEvent StatusMsg("SAVE," & CStr(1) & "Pattern-(Saving the " & CStr(1) & " Pattern)...")
                Me.m_AreaGrabberManager.PrepareAllRequest("SAVE", , , , , , , , , Me.m_TimeoutRecipe.Grabber.SAVE)
                Me.m_SystemLog.WriteLog(Now, ", [Cmd=SAVE]Grabber save," & CStr(1) & "Pattern " & Me.m_CurrentPattern.PATTERN_NAME)
                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.SAVE)
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Seq=" & CStr(1) & " Grabber2Xq1W! Communicate with Grabber fail! <SAVE>")
            End If
            t2 = Environment.TickCount
            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", SAVE Image,Seq=" & Me.m_CurrentPattern.PATTERN_NAME & " => " & t2 - t1)
        End If
        'get Chip_id
        t1 = Environment.TickCount
        'GET_PANEL_ID指令加PanelID,為了存二維碼時，文件有Name值
        Me.m_AreaGrabberManager.PrepareAllRequest("GET_PANEL_ID", PanelID, , , , , , , , Me.m_TimeoutRecipe.Grabber.PANEL_ID)
        SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.PANEL_ID)
        t2 = Environment.TickCount
        Me.intSpeed_GrabberGrab = t2 - t1
        If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Grabber_get_Chip_id" & "=> " & t2 - t1)

        Return SubSystemResult

    End Function

    Private Sub GrabberGrabProcess()
        Dim t1 As Integer
        Dim t2 As Integer
        Dim SubSystemResult As CResponseResult
        Dim AreaFrameResult As New ClsAreaFrameResult
        Dim RetryCount As Integer
        Dim StopWatch As New Stopwatch

        Try
            Me.m_GrabberGrabExMsg = ""
            RetryCount = 0

            '--------------------- Grab Image (Dynamic Exposure) ---------------------------------------------
            t1 = Environment.TickCount

            intSpeed_GrabberGrab_t1 = t1
            m_CurrentPatternIdx = m_CurrentPatternIdx + 1

            ' Dynamic Exp (include Grab)
            RaiseEvent StatusMsg(" Grabber 擷取第" & CStr(m_CurrentPatternIdx) & "個Pattern中...Pattern : " & Me.m_CurrentPattern.PATTERN_NAME)
            Me.m_AreaGrabberManager.PrepareAllRequest("GRAB", , , , , , , , , Me.m_TimeoutRecipe.Grabber.GRAB)
            Me.m_SystemLog.WriteLog(Now, ", [Cmd=GRAB]Grabber擷取第" & CStr(m_CurrentPatternIdx) & "個Pattern " & Me.m_CurrentPattern.PATTERN_NAME)
            SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.GRAB)
            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then
                Me.m_GrabberGrabExMsg = ("Grabber擷取" & Me.m_CurrentPattern.PATTERN_NAME & " Pattern 影像失敗! Communicate with Grabber fail! <GRAB>")
            End If
            t2 = Environment.TickCount
            Me.intSpeed_GrabberGrab = t2 - t1
            If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Grabber 擷取第" & CStr(m_CurrentPatternIdx) & "個Pattern中..." & "=> " & t2 - t1)

        Catch ex As Exception
            Me.m_GrabberGrabExMsg = "[GrabberGrabProcess] " & ex.Message
        End Try
    End Sub

    Private Sub GrabberPatternInfoProcess()
        Dim SubSystemResult As CResponseResult
        Dim t1, t2 As Integer

        Try
            Me.m_GrabberPatternInfoExMsg = ""

            Me.intSpeed_GrabberPatternInfo = 0
            t1 = Environment.TickCount

            If Me.m_BootRecipe.GENERATION = "L5B" Or Me.m_BootRecipe.GENERATION = "L6B" Or Me.m_BootRecipe.GENERATION = "L5C" Or Me.m_BootRecipe.GENERATION = "L6A" Or Me.m_BootRecipe.GENERATION = "L7B" Or Me.m_BootRecipe.GENERATION = "L7A" Then
                RaiseEvent StatusMsg("Grabber設定Pattern資訊中(Info Grabber setting Pattern)...")
            Else
                RaiseEvent StatusMsg("Mura檢測設定Pattern資訊中...")
            End If

            Me.m_AreaGrabberManager2.PrepareAllRequest("PATTERN_INFO", Me.m_CurrentPattern.PATTERN_NAME, Me.m_CurrentPattern.IMG_PROC_RECIPE, Me.m_IsLastPattern, , , , , , Me.m_TimeoutRecipe.Grabber.PATTERN_INFO)
            SubSystemResult = Me.m_AreaGrabberManager2.SendRequest(Me.m_TimeoutRecipe.Grabber.PATTERN_INFO)

            If Me.m_BootRecipe.GENERATION = "L5B" Or Me.m_BootRecipe.GENERATION = "L6B" Or Me.m_BootRecipe.GENERATION = "L5C" Or Me.m_BootRecipe.GENERATION = "L6A" Or Me.m_BootRecipe.GENERATION = "L7B" Or Me.m_BootRecipe.GENERATION = "L7A" Then
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber設定Pattern資訊失敗! Communicate with Grabber fail! <PATTERN_INFO>")
            Else
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Mura檢測設定Pattern資訊失敗 <PATTERN_INFO>")
            End If

            t2 = Environment.TickCount
            Me.intSpeed_GrabberPatternInfo = t2 - t1

        Catch ex As Exception
            Me.m_GrabberPatternInfoExMsg = "[GrabberPatternInfoProcess] " & ex.Message
        End Try
    End Sub

    '********************************New       Controller <====> Grabber <=====> IP        Structure *********************************
    'Multi CCD inspect same Panel Mode          :  GrabberOutput = IP2-1 output # IP2-2 output   ; CCD parameter

    'Each CCD inspect different Panel Mode      :  GrabberOutput = IP2-1 output @ IP2-2 output ......

    '   Area Frame Structure            8吋以下產品 (Each CCD inspect different Panel Mode)                                                      8吋以上產品 (Multi CCD inspect same Panel Mode)
    '     __________________________________________________________________________________________                            __________________________________________________________________________________________         
    '    | _X asix__________________________________________________________________________________ |                         | _X asix__________________________________________________________________________________ |
    '           |  CCD 1 |            |  CCD 2 |                     shift                                                                                                             |  CCD 1 |            |  CCD 2 |               shift  
    '           |________ |             |________ |               <--------->                                                                            or                      |________ |            |________ |           <--------->
    '                  
    '                 
    '        _____________        _____________                                         _____________       _____________                                                          _____________                                                         _____________
    '       | ___工位1___ |     | ___工位2___ |                                       | ___工位3___ |     | ___工位4___ |                                                      | ___工位1___ |                                                      | ___工位2___ |

    ' ImageProcessResultArray(4, 4) As String 'first "4" : Inspect Status , InspectOtherReason ,Other_Pattern,Other_CCD  ; Sec "4" CCD1~CCD2  record by pattaen
    '__                                                                                                                            __
    '| Penal 1 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 1 (IP1) Current Pattern ImageProcess Result
    '| Penal 2 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 2 (IP2) Current Pattern ImageProcess Result
    '| Penal 2 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 3 (IP3) Current Pattern ImageProcess Result
    '| Penal 3 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 4 (IP4) Current Pattern ImageProcess Result
    '__ 

    '*******************************************************************************************************************************************

    Private Sub ImgProcProcess()
        Dim j As Integer
        Dim SubSystemResult As CResponseResult
        Dim strTemp, strBackupPath As String
        Dim t1, t2 As Integer
        Dim result As String = String.Empty

        Try
            Me.m_ImgProcExMsg = ""
            Me.intSpeed_GrabberPatternInfo = 0
            t1 = Environment.TickCount

            If Me.m_FurtherPattern.IMG_PROC_RECIPE.Substring(0, 1).ToUpper = "F" Then
                RaiseEvent StatusMsg("Grabber設定Pattern資訊中(Info Grabber setting Pattern)...")
            Else
                RaiseEvent StatusMsg("Mura檢測設定Pattern資訊中...")
            End If

            Me.m_AreaGrabberManager2.PrepareAllRequest("PATTERN_INFO", Me.m_FurtherPattern.IMG_PROC_RECIPE, Me.m_IsLastPattern, "Gray", , , , , , Me.m_TimeoutRecipe.Grabber.PATTERN_INFO)
            SubSystemResult = Me.m_AreaGrabberManager2.SendRequest(Me.m_TimeoutRecipe.Grabber.PATTERN_INFO)
            Me.m_SystemLog.WriteLog(Now, ", [Cmd=PATTERN_INFO]," & Me.m_FurtherPattern.PATTERN_NAME & "," & Me.m_FurtherPattern.IMG_PROC_RECIPE & "," & Me.m_IsLastPattern & "," & "," & Me.m_TimeoutRecipe.Grabber.PATTERN_INFO)

            If Me.m_FurtherPattern.IMG_PROC_RECIPE.Substring(0, 1).ToUpper = "F" Then
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber設定Pattern資訊失敗! Communicate with Grabber fail! <PATTERN_INFO>")
            Else
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Mura檢測設定Pattern資訊失敗 <PATTERN_INFO>")
            End If

            't2 = Environment.TickCount
            'Me.intSpeed_GrabberPatternInfo = t2 - t1
            'End If
        Catch ex As Exception
            Me.m_ImgProcExMsg = "[ImgProcProcess <PATTERN_INFO>] " & ex.ToString
        End Try

        '********************************New       Controller <====> Grabber <=====> IP        Structure *********************************
        'Multi CCD inspect same Panel Mode           :  GrabberOutput = IP2-1 output # IP2-2 output   ; CCD parameter

        'Each CCD inspect different Panel Mode      :  GrabberOutput = IP2-1 output @ IP2-2 output

        '   Area Frame Structure            8吋以下產品 (Each CCD inspect different Panel Mode)                                                      8吋以上產品 (Multi CCD inspect same Panel Mode)
        '     __________________________________________________________________________________________                            __________________________________________________________________________________________         
        '    | _X asix__________________________________________________________________________________ |                         | _X asix__________________________________________________________________________________ |
        '           |  CCD 1 |            |  CCD 2 |                     shift                                                                                                             |  CCD 1 |            |  CCD 2 |               shift  
        '           |________ |             |________ |               <--------->                                                                            or                      |________ |            |________ |           <--------->
        '                  
        '                 
        '        _____________        _____________                                         _____________       _____________                                                          _____________                                                         _____________
        '       | ___工位1___ |     | ___工位2___ |                                       | ___工位3___ |     | ___工位4___ |                                                      | ___工位1___ |                                                      | ___工位2___ |

        ' ImageProcessResultArray(4, 4) As String 'first "4" : Inspect Status , InspectOtherReason ,Other_Pattern,Other_CCD  ; Sec "4" CCD1~CCD2  record by pattaen
        '__                                                                                                                            __
        '| Penal 1 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 1 (IP1) Current Pattern ImageProcess Result
        '| Penal 2 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 2 (IP2) Current Pattern ImageProcess Result
        '| Penal 2 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 3 (IP3) Current Pattern ImageProcess Result
        '| Penal 3 InspectStatus  ,InspectOtherReason,Other_Pattern,Other_CCD|          'Panel 4 (IP4) Current Pattern ImageProcess Result
        '__                                                                                                                          __
        '*******************************************************************************************************************************************

        Try
            Me.intSpeed_ImgProcProcess = 0
            t1 = Environment.TickCount

            RaiseEvent StatusMsg("影像處理第" & (Me.m_Current + 1).ToString & "個Pattern中(Image processing the " & (Me.m_Current + 1).ToString & " Pattern)...")

            '------------------------------------------- ALIGN ---------------------------------------------------------------------
            If Me.m_FurtherPattern.ALIGNMENT Then
                Me.m_AreaGrabberManager2.PrepareAllRequest("ALIGN", , , , , , , , , Me.m_TimeoutRecipe.Grabber.ALIGN)
                SubSystemResult = Me.m_AreaGrabberManager2.SendRequest(Me.m_TimeoutRecipe.Grabber.ALIGN)

                Dim i As Integer
                Dim StrSpilit() As String
                Dim Strs() As String

                Select Case SubSystemResult.CommResult
                    Case AUO.SubSystemControl.eCommResult.OK
                        For j = 0 To SubSystemResult.Responses.Length - 1  '都只連一隻Grabber,原則上 SubSystemResult.Responses.Length都為1             
                            '*********Each CCD inspect different Panel Mode**********
                            ' Parameter 1 : IP1 OK/OTHER_IMGPROC ; IP2   OK/OTHER_IMGPROC ; IP3   OK/OTHER_IMGPROC ; IP4   OK/OTHER_IMGPROC
                            StrSpilit = SubSystemResult.Responses(j).Param1.ToUpper.Split("@")
                            For i = 0 To StrSpilit.Length - 1    'Panel 1~Panel4 
                                '--- Align Abnormal ---
                                If StrSpilit(i) = "" Then Continue For

                                If StrSpilit(i) <> "OK" Then
                                    Select Case StrSpilit(i)
                                        Case "OTHER_IMGPROC"
                                            ''--- 記錄Other CCD與Pattern ---
                                            Me.m_Other_CCD_AllDefect = SubSystemResult.Responses(j).Param2   ' Parameter 2 : IP1 NA/OTHER_ALIGN_DEFECT ; IP2 NA/OTHER_ALIGN_DEFECT ; IP3 NA/OTHER_ALIGN_DEFECT ; IP4 NA/OTHER_ALIGN_DEFECT
                                            'Defect裡面1~N隻CCD(Panel)的資訊用 # 隔開
                                            Strs = Me.m_Other_CCD_AllDefect.Split(",")

                                            'InspectStatus
                                            Me.m_InspectStatus = "OTHER_ALIGN_DEFECT"
                                            Me.ImageProcessResultArray(i, 0) = "OTHER_ALIGN_DEFECT"

                                            'InspectOtherReason 
                                            Me.ImageProcessResultArray(i, 1) = Strs(3)

                                            'Other_Pattern
                                            Me.ImageProcessResultArray(i, 2) = Me.m_FurtherPattern.PATTERN_NAME


                                            'Other_CCD                    
                                            Me.ImageProcessResultArray(i, 3) = i + 1
                                            '------------------------------
                                            Me.m_SystemLog.WriteLog(Now, ", [Warning]AreaGrabber ALIGN Error," & Me.ImageProcessResultArray(i, 0) & "," & Me.ImageProcessResultArray(i, 1))
                                            'Throw New Exception("Align Error, " & Me.ImageProcessResultArray(i, 0) & "," & Me.ImageProcessResultArray(i, 1)
                                        Case Else
                                            Throw New Exception("Align Error, 影像處理回傳參數1(Param1)有誤! 內容為(" & SubSystemResult.Responses(j).Param1 & ")")
                                    End Select
                                    Exit Select
                                    'Else 'StrSpilit(i) = "OK"
                                    '    Me.ImageProcessResultArray(i, 0) = "OK"
                                    '    'Me.m_InspectStatus = "OK"  ' 只要有一片是OK即可,因多片檢無法個別提前退片
                                End If
                            Next i
                        Next j

                    Case AUO.SubSystemControl.eCommResult.ERR
                        Me.m_AreaGrabberImgProcErrMsg = "Area Grabber 影像處理發生錯誤!<ALIGN>," & SubSystemResult.ErrMessage

                    Case AUO.SubSystemControl.eCommResult.TIMEOUT
                        Me.m_AreaGrabberImgProcErrMsg = "Area Grabber 影像處理超過時間沒有回應! Communicate with Grabber is timeout! <ALIGN>"
                End Select
            End If

        Catch ex As Exception
            Me.m_ImgProcExMsg = "[ImgProcProcess <ALIGN>] " & ex.ToString
        End Try

        '********************************New       Controller <====> Grabber <=====> IP        Structure *********************************
        'Multi CCD inspect same Panel Mode           :  GrabberOutput = IP2-1 output # IP2-2 output  


        'Each CCD inspect different Panel Mode      :  GrabberOutput = IP2-1 output @ IP2-2 output

        '*******************************************************************************************************************************************

        Try
            '----------------------------------------------- CALCULATE ---------------------------------------------------
            If Me.m_InspectStatus = "OK" Then

                Dim i As Integer
                Dim StrSpilit() As String
                Dim Strs() As String
                Dim StrReason() As String

                '清空GrabberOutput資料
                For j = 0 To Me.m_GrabberOutput.Length - 1
                    Me.m_GrabberOutput(j) = ""
                Next j

                If Me.m_FurtherPattern.IMG_PROC_RECIPE.Substring(0, 1).ToUpper = "F" Then
                    strTemp = "P" & Me.m_FurtherPattern.IMG_PROC_RECIPE & "_FFunc_C#.txt"
                Else
                    strTemp = "P" & Me.m_FurtherPattern.IMG_PROC_RECIPE & "_FMura_C#.txt"
                End If

                Me.m_GrabberOutputFilename = Me.m_BootRecipe.GRABBER_OUTPUT_PATH & strTemp

                strBackupPath = IIf(Me.m_Ui.blnBackupPatternDefect, Me.m_BootRecipe.GRABBER_OUTPUT_BACKUP_PATH, "")
                Me.m_AreaGrabberManager2.PrepareAllRequest("CALCULATE", Me.m_GrabberOutputFilename, strBackupPath, Me.m_FurtherPattern.GRABBER_PARTICLE_RULE.ToString, , "STREAM", , , , Me.m_TimeoutRecipe.Grabber.CALCULATE)
                SubSystemResult = Me.m_AreaGrabberManager2.SendRequest(Me.m_TimeoutRecipe.Grabber.CALCULATE)
                Me.m_SystemLog.WriteLog(Now, ", [Cmd=CALCULATE]  ," & "," & Me.m_GrabberOutputFilename & "," & strBackupPath & "," & Me.m_FurtherPattern.GRABBER_PARTICLE_RULE.ToString & "," & "" & "," & "STREAM" & "," & Me.m_TimeoutRecipe.Grabber.CALCULATE)


                Select Case SubSystemResult.CommResult
                    Case AUO.SubSystemControl.eCommResult.OK
                        For j = 0 To SubSystemResult.Responses.Length - 1           '都只連一隻Grabber,原則上 SubSystemResult.Responses.Length都為1             

                            '*********Each CCD inspect different Panel Mode**********
                            ' Parameter 1 : IP1 OK/OTHER_IMGPROC # IP2   OK/OTHER_IMGPROC # IP3   OK/OTHER_IMGPROC # IP4   OK/OTHER_IMGPROC
                            StrSpilit = SubSystemResult.Responses(j).Param1.ToUpper.Split("@")
                            For i = 0 To StrSpilit.Length - 1    'Panel 1~Panel4 
                                '--- Align Abnormal ---
                                If StrSpilit(i) = "" Then Continue For
                                If StrSpilit(i) <> "OK" Then
                                    Select Case StrSpilit(i)                                              '
                                        Case "OTHER_IMGPROC"
                                            ''--- 記錄Other CCD與Pattern ---
                                            Me.m_Other_CCD_AllDefect = SubSystemResult.Responses(j).Param2   ' Parameter 2 : IP1 NA/OTHER_ALIGN_DEFECT # IP2 NA/OTHER_ALIGN_DEFECT # IP3 NA/OTHER_ALIGN_DEFECT # IP4 NA/OTHER_ALIGN_DEFECT
                                            'Defect裡面1~N隻CCD的資訊用 # 隔開
                                            Strs = Me.m_Other_CCD_AllDefect.Split("@")

                                            'InspectStatus & InspectOtherReason
                                            If Not Strs(i).ToString.Equals("") Then
                                                StrReason = Strs(i).Split(",")

                                                'InspectStatus
                                                Me.ImageProcessResultArray(i, 0) = StrReason(2)
                                                'InspectOtherReason
                                                Me.ImageProcessResultArray(i, 1) = StrReason(3)
                                            End If

                                            'Other_Pattern
                                            Me.ImageProcessResultArray(i, 2) = Me.m_FurtherPattern.PATTERN_NAME

                                            'Other_CCD                    
                                            Me.ImageProcessResultArray(i, 3) = i + 1
                                            '------------------------------
                                            Me.m_SystemLog.WriteLog(Now, ", [Warning]AreaGrabber CALCULATE Error," & Me.ImageProcessResultArray(i, 0) & "," & Me.ImageProcessResultArray(i, 1))
                                        Case "OTHER_MACHINE"
                                            ''--- 記錄Other CCD與Pattern ---
                                            Me.m_Other_CCD_AllDefect = SubSystemResult.Responses(j).Param2   ' Parameter 2 : IP1 NA/OTHER_ALIGN_DEFECT # IP2 NA/OTHER_ALIGN_DEFECT # IP3 NA/OTHER_ALIGN_DEFECT # IP4 NA/OTHER_ALIGN_DEFECT
                                            'Defect裡面1~N隻CCD的資訊用 # 隔開
                                            Strs = Me.m_Other_CCD_AllDefect.Split("@")

                                            'InspectStatus & InspectOtherReason
                                            If Not Strs(i).ToString.Equals("") Then
                                                StrReason = Strs(i).Split(",")

                                                'InspectStatus
                                                Me.ImageProcessResultArray(i, 0) = StrReason(2)
                                                'InspectOtherReason
                                                Me.ImageProcessResultArray(i, 1) = StrReason(3)
                                            End If

                                            'Other_Pattern
                                            Me.ImageProcessResultArray(i, 2) = Me.m_FurtherPattern.PATTERN_NAME

                                            'Other_CCD                    
                                            Me.ImageProcessResultArray(i, 3) = i + 1
                                            '------------------------------
                                            Me.m_SystemLog.WriteLog(Now, ", [Warning]AreaGrabber CALCULATE Error," & Me.ImageProcessResultArray(i, 0) & "," & Me.ImageProcessResultArray(i, 1))
                                        Case Else
                                            Throw New Exception("CALCULATE Error, 影像處理回傳參數1(Param1)有誤! 內容為(" & SubSystemResult.Responses(j).Param1 & "," & SubSystemResult.Responses(j).Param2 & ")")
                                    End Select
                                    'Exit Select
                                Else 'StrSpilit(i) = "OK"
                                    '收集GrabberOutput
                                    Me.m_GrabberOutput(j) = SubSystemResult.Responses(j).Param2
                                    If Me.ImageProcessResultArray(i, 0) <> "OTHER_ALIGN_DEFECT" AndAlso Me.ImageProcessResultArray(i, 0) <> "OTHER_GLASS_DEFECT" Then
                                        Me.ImageProcessResultArray(i, 0) = "OK"
                                    End If
                                End If
                            Next i
                        Next j
                    Case AUO.SubSystemControl.eCommResult.ERR
                        Me.m_AreaGrabberImgProcErrMsg = "Area Grabber 影像處理發生錯誤!<CALCULATE>, " & SubSystemResult.ErrMessage

                    Case AUO.SubSystemControl.eCommResult.TIMEOUT
                        Me.m_AreaGrabberImgProcErrMsg = "Area Grabber 影像處理超過時間沒有回應! Communicate with Grabber is timeout! <CALCULATE>"
                End Select
            End If

            t2 = Environment.TickCount
            Me.intSpeed_ImgProcProcess = t2 - t1

        Catch ex As Exception
            Me.m_ImgProcExMsg = "[ImgProcProcess <CALCULATE>] " & ex.Message
        End Try
    End Sub

#Region "--- AI Process ---"

    Public Sub InitAI()
        If m_AIManager Is Nothing Then Me.m_AIManager = New ClsAIIOIF("AI_Handshake")
        Try
            If Not Me.m_AIManager.IsConnect Then
                If Me.m_BootRecipe.USE_SUB_SYSTEM_CONTROL_LOG Then
                    Me.m_AIManager.Connect(Me.m_NetworkInfo.AI_IP, CInt(Me.m_NetworkInfo.AI_PORT), "D:\AOI_Data\Log\SubSystemControl\AI\falseDefect")
                Else
                    Me.m_AIManager.Connect(Me.m_NetworkInfo.AI_IP, CInt(Me.m_NetworkInfo.AI_PORT))
                End If
            End If
            RaiseEvent AIConnect()
        Catch ex As Exception
            Throw New Exception("[InitAI] " & ex.ToString)
        End Try
    End Sub

    Private Sub WaitMuraAIAnalysisComplete()
        Try
            While (Me.m_MuraMultiFlow.Status = ClsMuraOtherFlow.eStatus.BUSY)
                If Me.m_MuraMultiFlow.CLImgList.Count <= 0 Then
                    Exit While
                End If
            End While
            Me.m_MuraMultiFlow.TurnOff()
        Catch ex As Exception
            Me.m_SystemLog.WriteLog(Now, ", [Error]Wait Mura AI Analysis Complete Exception => " & ex.Message)
        End Try
    End Sub

    Public Sub MuraFlowControlSendData(ByVal strBackupPath As String)
        Dim SubSystemResult As CResponseResult
        If Me.m_MuraMultiFlow.Status = ClsMuraOtherFlow.eStatus.FAIL Then
            Throw New Exception("[Mura-Multi Flow] ERR_MSG : " & Me.m_MuraMultiFlow.ERR_MSG)
        End If
        For k As Integer = 0 To Me.m_MuraMultiFlow.CLAIDataList.Count - 1
            '----Change Mura Defect Name----------
            Dim tmpOutput As String = String.Empty
            Dim strData As ClsMuraOtherFlow.clsMuraAIResult = Me.m_MuraMultiFlow.CLAIDataList(k)

            If Me.m_MuraMultiFlow.IsHaveMuraDefectCount(strData.Output_Data) Then
                If Me.m_InspectStatus.Contains("OTHER_") Then
                    Me.m_InspectStatus = "OK"
                End If
            End If

            Me.m_MuraMultiFlow.AnalysisMuraGrabberOutput(strData.Output_Data)
            tmpOutput = Me.m_MuraMultiFlow.GetPatternDataLine()
            tmpOutput = strData.Output_Data
            '----------------------------------------------

            '----------Send Mura Data------------
            Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_CCD_MURA_STREAM]Judger載入判片資料")
            Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_MURA_STREAM", tmpOutput, "NONE", "EnableFilter", , strBackupPath, , , , )
            SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_CCD_MURA_DATA)
            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception(" Judger載入Mura判片資料失敗! Communicate with Judger fail! <ADD_GRABBER_MURA_STREAM>")
            '------------------------------
        Next

    End Sub
    Private Function AIProcessFlowControl(ByVal MASK_ID_AOI As Boolean, ByVal JudgeRank As CJudgeRank, ByRef SecondMQRankForAI As String, ByRef SecondAGSRankForAI As String) As CJudgeRank
        If Me.m_AIWorker.CurrentModel Is Nothing Then Throw New Exception("[InspectionProcess][AI_Processing] AI_Worker Table is not object !!")

        Try
            '20170908 AI
            Dim t1, t2 As Integer
            Dim SubSystemResult As CResponseResult
            Dim strBackupPath As String
            Dim tmpJudgeRank As CJudgeRank

            '20180227 Judger AI Restore Data
            strBackupPath = IIf(Me.m_Ui.blnBackupFalseDefectTable, Me.m_BootRecipe.FALSE_DEFECT_TABLE_BACKUP_PATH, "")
            'AI_DEFECT_RESTORE
            Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_FILTER_STREAM", "1;1_1;342;554;1;3000;", "AI_DEFECT_RESTORE", "EnableFilter", , strBackupPath, , , , 6000)
            Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_GRABBER_FILTER_STREAM]Judger載入判片資料 ")
            SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_AREA_CCD_DATA)
            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("[AI Processing] Judger載入Func判片資料失敗! Communicate with Judger fail! <ADD_GRABBER_FILTER_STREAM>" & vbCrLf & SubSystemResult.ErrMessage)
            'AI_PARTICLE_REFILTER
            Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_FILTER_STREAM", "1;1_1;342;554;1;3000;", "AI_PARTICLE_REFILTER", "EnableFilter", , strBackupPath, , , , 6000)
            Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_GRABBER_FILTER_STREAM]Judger載入判片資料 ")
            SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_AREA_CCD_DATA)
            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("[AI Processing] Judger載入Func判片資料失敗! Communicate with Judger fail! <ADD_GRABBER_FILTER_STREAM>" & vbCrLf & SubSystemResult.ErrMessage)
            'AI_DEFECT_RESTORE
            Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_MURA_STREAM", "1;1_1;342;554;1;3000;", "AI_DEFECT_RESTORE", "EnableFilter", "10#10#10", strBackupPath, , , , 6000)
            Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_GRABBER_FILTER_STREAM]Judger載入判片資料 ")
            SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_AREA_CCD_DATA)
            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("[AI Processing] Judger載入Func判片資料失敗! Communicate with Judger fail! <ADD_GRABBER_FILTER_STREAM>" & vbCrLf & SubSystemResult.ErrMessage)
            'AI_PARTICLE_REFILTER
            Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_MURA_STREAM", "1;1_1;342;554;1;3000;", "AI_PARTICLE_REFILTER", "EnableFilter", "10#10#10", strBackupPath, , , , 6000)
            Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_GRABBER_FILTER_STREAM]Judger載入判片資料 ")
            SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_AREA_CCD_DATA)
            If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("[AI Processing] Judger載入Func判片資料失敗! Communicate with Judger fail! <ADD_GRABBER_FILTER_STREAM>" & vbCrLf & SubSystemResult.ErrMessage)

            'AI Data Analysis & Judger DATA Stream
            Select Case Me.m_BootRecipe.AI_Ver
                Case eAIVer.Original_Ver
                    Throw New Exception("[AI_PROCESS][AI_VERSION] : " & Me.m_BootRecipe.AI_Ver.ToString() & ", 已不提供支持")
                Case eAIVer.Public_Ver
                    Dim AIFilePath As String = String.Format("{0}{1}{2}{3}.xml", Me.m_BootRecipe.JUDGER_MQ_REPORT_BACKUP_PATH, Format(Now, "yyyyMMdd") & "\", Me.m_ProberPanelId, Me.m_BootRecipe.REPORT_TYPE)
                    Dim CLTWorkerID As Integer = 0
                    Dim ODTWorkerID As Integer = 0
                    Dim AETWorkerID As Integer = 0
                    Dim tmpWorker As String = String.Empty
                    For Each worker As clsAIWorker.ClsAIWorkInfo In Me.m_AIWorker.CurrentModel.CurrentUseAIModelList
                        If worker Is Nothing Then Throw New Exception("[InspectionProcess][AI_Processing]Get_worker Fail")
                        If worker.IsEnable Then

                            Select Case worker.AIKernel
                                Case eAIKernel.NONE, eAIKernel.CLT
                                    tmpWorker = String.Format("{0};{1}", CLTWorkerID, worker.AIClassType.ToString())
                                Case eAIKernel.AET
                                    tmpWorker = String.Format("{0};{1}", AETWorkerID, worker.AIClassType.ToString())
                                Case eAIKernel.ODT
                                    tmpWorker = String.Format("{0};{1}", ODTWorkerID, worker.AIClassType.ToString())
                            End Select

                            Select Case worker.AIClassType
                                Case eAIClassType.FUNC, eAIClassType.MURA
                                    Me.AIProcessing(worker.AIClassType, tmpWorker, AIFilePath, worker.AnalysisDefect, worker.AnalysisItem, worker.AnalysisPattern)
                                    Select Case worker.AIKernel
                                        Case eAIKernel.NONE, eAIKernel.CLT
                                            CLTWorkerID += 1
                                        Case eAIKernel.AET
                                            AETWorkerID += 1
                                        Case eAIKernel.ODT
                                            ODTWorkerID += 1
                                    End Select
                            End Select
                        End If
                    Next
            End Select

            Return tmpJudgeRank
        Catch ex As Exception
            Throw New Exception("[Step7: AI Processing][ERR] : " & ex.ToString)
        End Try
    End Function

    'AI Processing--For Public version
    Private Function AIProcessing(ByVal AIClass As eAIClassType, ByVal WorkID As String, ByVal xmlFilePath As String, ByVal AnalysisDefect As Boolean, ByVal AnalysisItem As String, ByVal AnalysisPattern As String) As Boolean
        Dim SubSystemResult As CResponseResult
        Dim strBackupPath As String = String.Empty
        Dim bERR As Boolean = False
        Dim bRecipeLoadERR As Boolean = False

        If AnalysisItem <> String.Empty Then
            AnalysisItem = AnalysisItem.TrimEnd(";")
        End If
        If AnalysisPattern <> String.Empty Then
            AnalysisPattern = AnalysisPattern.TrimEnd(";")
        End If

        '[AI] cmd-AI_ANALYSIS, 回傳Response結果
        Dim AIResult As AUO.SubSystemControlV2.CResponseResult = Me.m_AIManager.PredictFolder(WorkID, xmlFilePath, AnalysisDefect, AnalysisItem, AnalysisPattern, Me.m_BootRecipe.AI_OUTPUT_FOLDER, Me.m_TimeoutRecipe.AI.PREDICT_FOLDER)
        If AIResult.CommResult <> AUO.SubSystemControlV2.TCPEnum.eCommResult.OK Then Throw New Exception("[AIProcessing][PredictFolder]WORKER_ID : " & WorkID & ", AI Communication Fail !!")
        If AIResult.Responses.Result <> "OK" Then
            If AIResult.Responses.Param1 = eAIERR.NO_RECIPE_LOAD Then
                bRecipeLoadERR = True
            End If
            bERR = True
            Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[AIProcessing][PredictFolder]WORKER_ID : " & WorkID & ", AI分析回覆ERR, [ERR_MSG] : " & AIResult.Responses.Param2)
        Else
            Me.m_AllGrabberOutput = AIResult.Responses.Param1
            ' for Function---AI
            strBackupPath = IIf(Me.m_Ui.blnBackupFalseDefectTable, Me.m_BootRecipe.FALSE_DEFECT_TABLE_BACKUP_PATH, "")
        End If
        'AI_ADD
        Select Case AIClass
            Case eAIClassType.FUNC
                If bERR = False Then
                    Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_FILTER_STREAM", Me.m_AllGrabberOutput, "AI_ADD", "EnableFilter", , strBackupPath, , , , 6000)
                    Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_GRABBER_FILTER_STREAM]Judger載入判片資料 ")
                    SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_AREA_CCD_DATA)
                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("[AI Processing] Judger載入Func判片資料失敗! Communicate with Judger fail! <ADD_GRABBER_FILTER_STREAM>" & vbCrLf & SubSystemResult.ErrMessage)
                End If
            Case eAIClassType.MURA
                If bERR = False Then
                    Me.m_JudgerManager.PrepareAllRequest("ADD_GRABBER_MURA_STREAM", Me.m_AllGrabberOutput, "AI_ADD", "EnableFilter", "10#10#10", strBackupPath, , , , 6000)
                    Me.m_SystemLog.WriteLog(Now, Format(Now, "yyyy/MM/dd HH:mm:ss ") & "[Cmd=ADD_GRABBER_MURA_STREAM]Judger載入判片資料 ")
                    SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.ADD_AREA_CCD_DATA)
                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("[AI Processing] Judger載入Func判片資料失敗! Communicate with Judger fail! <ADD_GRABBER_FILTER_STREAM>" & vbCrLf & SubSystemResult.ErrMessage)
                End If
        End Select

        ''Return Result
        Return IIf(bERR And bRecipeLoadERR, False, True)
    End Function

    Private Sub AILoadModel(ByVal Enable As Boolean)
        If Enable = False Then Return

        Dim CLTWorkerID As Integer = 0
        Dim AETWorkerID As Integer = 0
        Dim ODTWorkerID As Integer = 0
        Dim FUNC_TH As String
        Dim MURA_TH As String

        '----AI MODEL LOAD----
        If Me.m_AIManager.IsConnect Then
            If Me.m_AIWorker.CurrentModel Is Nothing Or Me.m_AIWorker.ModelCount <= 0 Then Throw New Exception("[Initializable_controller][AI_WORKER_TABLE]WORKER TABLE has not data!!")

            Dim aiKernel As String = String.Empty
            For Each worker As clsAIWorker.ClsAIWorkInfo In Me.m_AIWorker.CurrentModel.CurrentUseAIModelList
                If worker.IsEnable Then
                    Dim tmpWorkerID As Integer = 0
                    Select Case worker.AIClassType
                        Case eAIClassType.MURA, eAIClassType.FUNC, eAIClassType.MURA_MULTI
                            'load TH By Worker Name 
                            Me.m_AIWorker.ThresTable.SetCurrentModel(m_AIWorker.CurrentModel.ModelName, worker.WorkName, worker.AnalysisDefect,
                                                                     worker.AnalysisItem, worker.AIKernel)
                            '--------------------------------------------------------------------------------

                            Select Case worker.AIKernel
                                Case eAIKernel.NONE, eAIKernel.CLT
                                    tmpWorkerID = CLTWorkerID
                                Case eAIKernel.AET
                                    tmpWorkerID = AETWorkerID
                                Case eAIKernel.ODT
                                    tmpWorkerID = ODTWorkerID
                            End Select

                            If worker.AIKernel = eAIKernel.NONE Or worker.AIKernel = eAIKernel.CLT Then
                                aiKernel = String.Empty
                            Else
                                aiKernel = worker.AIKernel.ToString()
                            End If
                            FUNC_TH = Me.m_AIWorker.ThresTable.GetFuncTH().TrimEnd(";")
                            MURA_TH = Me.m_AIWorker.ThresTable.GetMuraTH().TrimEnd(";")
                            Dim AIResult As AUO.SubSystemControlV2.CResponseResult = Me.m_AIManager.LoadRecipe(tmpWorkerID, worker.RecipePath, FUNC_TH, MURA_TH, _
                                                                                                               aiKernel, worker.AILoadType.ToString(), _
                                                                                                               worker.AIOutputType.ToString(), _
                                                                                                               Me.m_TimeoutRecipe.AI.LOAD_RECIPE)
                            If AIResult.CommResult <> AUO.SubSystemControlV2.TCPEnum.eCommResult.OK Then Throw New Exception("[AI_LOAD_RECIPE] Worker_Name : " & worker.WorkName & ", LOAD RECIPE  COMMUNICATION FAIL!!")
                            If AIResult.Responses.Result <> "OK" Then Throw New Exception("[AI_LOAD_RECIPE][FUNC/MURA]LOAD RECIP FAIL!!")

                            Select Case worker.AIKernel
                                Case eAIKernel.NONE, eAIKernel.CLT
                                    CLTWorkerID += 1
                                Case eAIKernel.AET
                                    AETWorkerID += 1
                                Case eAIKernel.ODT
                                    ODTWorkerID += 1
                            End Select
                    End Select
                End If
            Next

            Me.m_SystemLog.WriteLog(Now, ", [檢測啟動][PreparationProcess] C-> CELL_TEST_AI CMD_LoadRecipe, Complete")
        End If

    End Sub

    Public Sub AIStopWork(ByVal Enable As Boolean)
        If Enable = False Then Return

        If Me.m_AIManager IsNot Nothing AndAlso Me.m_AIManager.IsConnect Then
            If Me.m_AIWorker.CurrentModel Is Nothing OrElse Me.m_AIWorker.ModelCount = 0 Then Throw New Exception("[AI_STOP_WORK][AI_WORKER_TABLE]WORKER TABLE has not data!!")
            Dim CLTWorkerID As Integer = 0
            Dim AETWorkerID As Integer = 0
            Dim ODTWorkerID As Integer = 0

            For Each worker As clsAIWorker.ClsAIWorkInfo In Me.m_AIWorker.CurrentModel.CurrentUseAIModelList
                If worker.IsEnable Then
                    Dim tmpWorkerID As String = 0
                    Select Case worker.AIKernel
                        Case eAIKernel.NONE, eAIKernel.CLT
                            tmpWorkerID = CLTWorkerID
                        Case eAIKernel.AET
                            tmpWorkerID = AETWorkerID
                        Case eAIKernel.ODT
                            tmpWorkerID = ODTWorkerID
                    End Select

                    Select Case worker.AIClassType
                        Case eAIClassType.MURA, eAIClassType.FUNC, eAIClassType.MURA_MULTI
                            Dim AIResult As AUO.SubSystemControlV2.CResponseResult = Me.m_AIManager.AIStop(tmpWorkerID, Me.m_TimeoutRecipe.AI.AI_STOP)
                            If AIResult.CommResult <> AUO.SubSystemControlV2.TCPEnum.eCommResult.OK Then Throw New Exception("[AI_STOP_WORK] Worker_Name : " & worker.WorkName & ", STOP_WORK COMMUNICATION FAIL!!")
                            If AIResult.Responses.Result <> "OK" Then Throw New Exception("[AI_STOP_WORK][FUNC/MURA]AI_STOP FAIL!!")

                            Select Case worker.AIKernel
                                Case eAIKernel.NONE, eAIKernel.CLT
                                    CLTWorkerID += 1
                                Case eAIKernel.AET
                                    AETWorkerID += 1
                                Case eAIKernel.ODT
                                    ODTWorkerID += 1
                            End Select
                    End Select
                End If
            Next

            Me.m_SystemLog.WriteLog(Now, ", [檢測啟動][AI_STOP_WORK] C->CELL_TEST_AI STOP, Complete")
        End If
    End Sub

#End Region

    'LuminanceMeasurementProcess
    Private Sub LuminanceMeasurementProcess(ByRef outputLuminanceData As String)
        Dim SubSystemResult As CResponseResult
        If Me.m_FurtherPattern.USE_LUMINANCE_MEASUREMENT Then
            Me.m_AreaGrabberManager2.PrepareAllRequest("CALCULATE_LUMINANCE", , , , , , , , , Me.m_TimeoutRecipe.Grabber.CALCULATE_LUMINANCE)
            SubSystemResult = Me.m_AreaGrabberManager2.SendRequest(Me.m_TimeoutRecipe.Grabber.CALCULATE_LUMINANCE)

            Select Case SubSystemResult.CommResult
                Case AUO.SubSystemControl.eCommResult.OK
                    For k As Integer = 0 To SubSystemResult.Responses.Length - 1
                        outputLuminanceData = SubSystemResult.Responses(k).Param1
                    Next k
                Case AUO.SubSystemControl.eCommResult.ERR
                    Me.m_AreaGrabberImgProcErrMsg = "Area Grabber 計算亮度值發生錯誤!<LuminanceMeasurementProcess>," & SubSystemResult.ErrMessage

                Case AUO.SubSystemControl.eCommResult.TIMEOUT
                    Me.m_AreaGrabberImgProcErrMsg = "Area Grabber 計算亮度值超過時間沒有回應! Communicate with Grabber is timeout! <LuminanceMeasurementProcess>"
            End Select
        End If

    End Sub
#End Region


    Public Sub ShowDlgBootConfig()
        Dim dlgBootConfig As dlgBootConfig = New dlgBootConfig(Me.m_bootRecipeFilename)

        If dlgBootConfig.ShowDialog() = DialogResult.OK Then
            Me.m_BootRecipe = AOIController.ClsBootConfig.ReadXML(m_bootRecipeFilename)
            Me.m_bootRecipeFilename = m_bootRecipeFilename
            Me.CloseController()
            Me.InitController(m_bootRecipeFilename)

            RaiseEvent UpdateInformation()
        End If
    End Sub

    Public Sub ShowDlgNetWorkConfig()
        Dim dlgNetWorkConfig As dlgNetWorkConfig = New dlgNetWorkConfig(Me.m_BootRecipe.NETWORK_CONFIG_FILE)
        dlgNetWorkConfig.ShowDialog()
    End Sub


    Public Sub ShowDlgMQConfig()
        Dim dlgMQConfig As dlgMQConfig = New dlgMQConfig(Me.m_BootRecipe.MQCONFIG_PATH)
        dlgMQConfig.ShowDialog()
    End Sub

    'Show System Log
    Public Sub ShowDlgSystemLog()
    End Sub

    'Show Inspect Log
    Public Sub ShowDlgInspectLog()
    End Sub

    'Show Error Log
    Public Sub ShowDlgErrorLog()
    End Sub

    'Show Speed Log
    Public Sub ShowDlgSpeedLog()
    End Sub

#Region "AreaFrame & PatGen Function"

    Private Sub WaitPanelRotationToAOIPositionProcess()
        Dim QueryStatus As ClsAreaFrameResult
        Me.m_PLCCommExMsg = ""
        Try
            While Not Me.m_StopInspection
                If Me.m_BootRecipe.GENERATION = "S13" AndAlso Me.m_BootRecipe.USE_AREAFRAME Then
                    '--- 讀取PLC狀態 ---       
                    QueryStatus = Me.m_AreaFrame.ReadStatus(Me.m_TimeoutRecipe.AreaFrame.ReadStatus)
                    If QueryStatus.CommResult = eCommResult.OK Then
                        If QueryStatus.Status_DarwinBL.IsStartInspect Then
                            Me.m_SystemLog.WriteLog(Now, ", [Info]Panel In AOI Position now, PLC Status=" & QueryStatus.Status_DarwinBL.RawData)
                            Me.m_StartTime = Environment.TickCount          ' 從Contact後開始計算檢測時間
                            Exit Sub
                        End If
                    Else
                        If Me.m_AreaFrame.CommCount >= Me.m_AreaFrame.RetryLimit Then Exit Sub
                    End If
                    System.Threading.Thread.Sleep(100)
                End If
            End While
        Catch ex As Exception
            Me.m_SystemLog.WriteLog(Now, ", [Error]Wait Panel rotation to AOI Position Exception => " & ex.Message)
            Me.m_PLCCommExMsg = ex.Message
        End Try
    End Sub

    'Public Function AreaFrameIsAuto() As Boolean
    '    Dim AreaFrameResult As ClsAreaFrameResult
    '    Dim blnRtn As Boolean = False

    '    AreaFrameResult = Me.m_AreaFrame.ReadStatus(Me.m_TimeoutRecipe.AreaFrame.ReadStatus)
    '    If AreaFrameResult.CommResult = eCommResult.OK AndAlso AreaFrameResult.Status.IsValid Then
    '        If AreaFrameResult.Status.IsAutoMode Then blnRtn = True
    '    Else
    '        Throw New Exception("[AreaFrameIsAuto] AreaFrame通訊發生錯誤! Communicate with AreaFrame fail! <Get Status>")
    '    End If

    '    Return blnRtn
    'End Function

    Private Sub SetAOIpcAlarm()
        Try
            Dim t1, t2 As Integer
            Dim AreaFrameResult As New ClsAreaFrameResult
            If Me.m_AreaFrame IsNot Nothing Then
                If Me.m_AreaFrame.IsConnect Then
                    AreaFrameResult = Me.m_AreaFrame.AlarmControl(eAlarmControl.SetAlarm)
                    t1 = Environment.TickCount
                    If AreaFrameResult.CommResult = eCommResult.OK Then
                        Me.m_SystemLog.WriteLog(Now, ", [Cmd= AlarmControl] SetAlarm OK!!")
                    ElseIf AreaFrameResult.CommResult = eCommResult.ERR Then
                        Me.m_ErrorLog.WriteLog(Now, ",  [Cmd= AlarmControl] SetAlarm Error!!")
                    Else
                        Me.m_ErrorLog.WriteLog(Now, ", [Cmd= AlarmControl] SetAlarm Time out!!")
                    End If
                    t2 = Environment.TickCount
                    If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", [Cmd= AlarmControl] SetAlarm => " & t2 - t1)
                End If

            End If
        Catch ex As Exception
            MessageBox.Show("[Kernel] [SetAOIpcAlarm] " & ex.Message, "Controller Err", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SetAOIpcWarning()
        Try
            Dim t1, t2 As Integer
            Dim AreaFrameResult As New ClsAreaFrameResult
            If Me.m_AreaFrame IsNot Nothing Then
                If Me.m_AreaFrame.IsConnect Then
                    AreaFrameResult = Me.m_AreaFrame.AlarmControl(eAlarmControl.SetWarning)
                    t1 = Environment.TickCount
                    If AreaFrameResult.CommResult = eCommResult.OK Then
                        Me.m_SystemLog.WriteLog(Now, ", [Cmd= AlarmControl] SetWarning OK!!")
                    ElseIf AreaFrameResult.CommResult = eCommResult.ERR Then
                        Me.m_ErrorLog.WriteLog(Now, ",  [Cmd= AlarmControl] SetWarning Error!!")
                    Else
                        Me.m_ErrorLog.WriteLog(Now, ", [Cmd= AlarmControl] SetWarning Time out!!")
                    End If
                    t2 = Environment.TickCount
                    If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", [Cmd= AlarmControl] SetWarning => " & t2 - t1)
                End If

            End If

        Catch ex As Exception
            MessageBox.Show("[Kernel] [SetAOIpcWarning] " & ex.Message, "Controller Err", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ResetAOIpcAlarm()

        Try
            Dim t1, t2 As Integer
            Dim AreaFrameResult As New ClsAreaFrameResult
            If Me.m_AreaFrame IsNot Nothing Then
                If Me.m_AreaFrame.IsConnect Then
                    AreaFrameResult = Me.m_AreaFrame.AlarmControl(eAlarmControl.Reset)
                    t1 = Environment.TickCount
                    If AreaFrameResult.CommResult = eCommResult.OK Then
                        Me.m_SystemLog.WriteLog(Now, ", [Cmd= AlarmControl] Reset OK!!")
                    ElseIf AreaFrameResult.CommResult = eCommResult.ERR Then
                        Me.m_ErrorLog.WriteLog(Now, ",  [Cmd= AlarmControl] Reset Error!!")
                    Else
                        Me.m_ErrorLog.WriteLog(Now, ", [Cmd= AlarmControl] Reset Time out!!")
                    End If
                    t2 = Environment.TickCount
                    If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", [Cmd= AlarmControl] Reset => " & t2 - t1)
                End If

            End If

        Catch ex As Exception
            MessageBox.Show("[Kernel] [ResetAOIpcAlarm] " & ex.Message, "Controller Err", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "Judger Grabber Function"

    Private Sub WaitResetJudgerGrabberProcess()



        Me.m_SystemLog.WriteLog(Now, ", [SendCommand Juddger  = CLEAR  & SET_DATE & PANEL_ID")
        ResetJudgerThread = New System.Threading.Thread(AddressOf Me.ResetJudgerProcess)
        ResetJudgerThread.Name = "ResetJudger"
        ResetJudgerThread.SetApartmentState(Threading.ApartmentState.STA)
        ResetJudgerThread.Start()

        Me.m_SystemLog.WriteLog(Now, ", [SendCommand Grabber  = CLEAR  & SET_DATE & PANEL_ID")
        ResetGrabberThread = New System.Threading.Thread(AddressOf Me.ResetGrabberProcess)
        ResetGrabberThread.Name = "ResetGrabber"
        ResetGrabberThread.SetApartmentState(Threading.ApartmentState.STA)
        ResetGrabberThread.Start()

        If Not ResetJudgerThread.Join(Me.m_TimeoutRecipe.Judger.CLEAR + 3000) Then
            Me.ResetJudgerThread.Abort()
            Throw New Exception("Seq=" & Me.m_CurrentModel.Pattern(m_CurrentPatternIdx).PATTERN_NAME & " Judger執行Reset命令逾時(" & CStr(Me.m_TimeoutRecipe.Judger.CLEAR) & "ms)! ")
        End If

        If Not ResetGrabberThread.Join(Me.m_TimeoutRecipe.Grabber.CLEAR + 1000) Then
            Me.ResetGrabberThread.Abort()
            Throw New Exception("Seq=" & Me.m_CurrentModel.Pattern(m_CurrentPatternIdx).PATTERN_NAME & " Grabber執行Reset命令逾時(" & CStr(Me.m_TimeoutRecipe.Grabber.CLEAR) & "ms)! " & Me.m_GrabberSetExpExMsg)
        End If

    End Sub

    Private Sub ResetJudgerProcess()
        Dim t1, t2 As Integer
        Dim SubSystemResult As CResponseResult
        Dim InspectDatetime As Date

        Me.intSpeed_ResetJudger = 0
        Me.m_ResetJudgerExMsg = ""

        Try
            t1 = Environment.TickCount
            InspectDatetime = Now
            ' Judger判片重置 
            If Me.m_Ui.blnJudge Then
                RaiseEvent StatusMsg("Judger重置中(Info Judger Resetting)...")
                Me.m_JudgerManager.PrepareAllRequest("CLEAR")
                Me.m_SystemLog.WriteLog(Now, ", [Cmd=CLEAR]Judger重置")
                SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.CLEAR)
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Judger重置失敗! Communicate with Judger fail! <CLEAR>")

                RaiseEvent StatusMsg("Judger設定檢測日期中(Info Judger Setting Date)...")
                Me.m_JudgerManager.PrepareAllRequest("SET_DATE", Format(InspectDatetime, "yyyy"), Format(InspectDatetime, "MM"), Format(InspectDatetime, "dd"), Format(InspectDatetime, "HH"), Format(InspectDatetime, "mm"), Format(InspectDatetime, "ss"))
                Me.m_SystemLog.WriteLog(Now, ", [Cmd=SET_DATE]Judger設定檢測日期")
                SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.SET_DATE)
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Judger設定檢測日期失敗! Communicate with Judger fail! <SET_DATE>")


                RaiseEvent StatusMsg("Judger設定PanelID中(Info Judger Setting PanelID)...")
                Me.m_JudgerManager.PrepareAllRequest("PANEL_ID", Me.m_ProberPanelId, [Enum].GetName(GetType(eRunMode), Me.m_Ui.RunMode), , , , "0", , , ) 'Parameter 6 : 目前的檢測工位 (決定ViewJudgerReult 更新哪一Form秀出"檢測中")

                Me.m_SystemLog.WriteLog(Now, ", [Cmd=PANEL_ID]Judger設定PanelID")
                SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.PANEL_ID)
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then
                    Throw New Exception("Judger設定PanelID失敗! Communicate with Judger fail! <PANEL_ID>")
                Else
                    If SubSystemResult.Responses(0).Param1.ToUpper = "YES" Then
                        Me.m_InspectStatus = "MASK_ID_REJECT"
                        Me.m_SystemLog.WriteLog(Now, ", [Info]Unload panel of MASK_ID_REJECT.")
                    End If
                End If

                t2 = Environment.TickCount
                intSpeed_ResetJudger = t2 - t1
                'If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Asyscn Judger Reset => " & t2 - t1)
            End If
        Catch ex As Exception
            Me.m_ResetJudgerExMsg = "[ResetJudgerProcess] " & ex.Message

            Me.m_ErrorLog.WriteLog(Now, ", [Error]Step1: Judger Reset => " & ex.Message)
            'Throw New Exception("[Step1: Judger Reset] " & ex.Message)
        End Try

    End Sub

    Private Sub ResetGrabberProcess()
        Dim SubSystemResult As CResponseResult
        Dim t1, t2 As Integer
        Dim InspectDatetime As Date

        Me.intSpeed_ResetGrabber = 0
        Me.m_ResetGrabberExMsg = ""
        Try

            t1 = Environment.TickCount
            InspectDatetime = Now
            ' Area Grabber重置
            If Me.m_CurrentModel.UseAreaGrabber Then

                RaiseEvent StatusMsg("Grabber檢測重置中(Info Grabber Resetting)...")

                Me.m_AreaGrabberManager.PrepareAllRequest("CLEAR", , , , , , , , , Me.m_TimeoutRecipe.Grabber.CLEAR)
                Me.m_SystemLog.WriteLog(Now, ", [Cmd=CLEAR]Grabber檢測重置")

                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.CLEAR)
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber檢測連續重置" & CStr(Me.m_AreaGrabberManager.RetryLimit) & "次失敗! Communicate with Grabber fail! <CLEAR>")

                RaiseEvent StatusMsg("Grabber檢測設定檢測日期中(Info Grabber Setting Date)...")

                Me.m_AreaGrabberManager.PrepareAllRequest("SET_DATE", Format(InspectDatetime, "yyyy"), Format(InspectDatetime, "MM"), Format(InspectDatetime, "dd"), Format(InspectDatetime, "HH"), Format(InspectDatetime, "mm"), Format(InspectDatetime, "ss"), , , Me.m_TimeoutRecipe.Grabber.SET_DATE)
                Me.m_SystemLog.WriteLog(Now, ", [Cmd=SET_DATE]Grabber檢測設定檢測日期")
                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.SET_DATE)
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber檢測設定檢測日期失敗! Communicate with Grabber fail! <SET_DATE>")

                RaiseEvent StatusMsg("Grabber檢測設定PanelID中(Info Grabber Setting PanelID)...")

                Me.m_AreaGrabberManager.PrepareAllRequest("PANEL_ID", Me.m_ProberPanelId, [Enum].GetName(GetType(eRunMode), Me.m_Ui.RunMode), , , , , , , Me.m_TimeoutRecipe.Grabber.PANEL_ID)

                Me.m_SystemLog.WriteLog(Now, ", [Cmd=PANEL_ID]Grabber檢測設定PanelID")
                SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.PANEL_ID)
                If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Grabber檢測設定PanelID失敗! Communicate with Grabber fail! <PANEL_ID>")
            End If
            t2 = Environment.TickCount
            'If Me.m_BootRecipe.USE_SPEED_LOG Then Me.m_SpeedLog.WriteLog(Now, ", Grabber Reset => " & t2 - t1)
            intSpeed_ResetGrabber = t2 - t1

        Catch ex As Exception
            Me.m_ErrorLog.WriteLog(Now, ", [Error]Step2: Grabber Reset => " & ex.Message)
            Me.m_ResetGrabberExMsg = "[ResetGrabberProcess] " & ex.Message

            'Throw New Exception("[Step2: Grabber Reset] " & ex.Message)
        End Try
    End Sub

    Private Function GetFinalRank(ByVal InspectStatus As String, ByVal InspectOtherReason As String) As CJudgeRank
        Dim SubSystemResult As CResponseResult
        Dim JudgeRank As New CJudgeRank

        Me.m_JudgerManager.PrepareAllRequest("GET_FINAL_RANK", InspectStatus, InspectOtherReason, , , , , , )
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=GET_FINAL_RANK] " & InspectStatus & "," & InspectOtherReason)

        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.GET_FINAL_RANK)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("判片失敗! Communicate with Judger fail! <GET_FINAL_RANK> (" & InspectStatus & "," & InspectOtherReason & ")")

        JudgeRank.MqRank = SubSystemResult.Responses(0).Param1

        JudgeRank.CstRank = SubSystemResult.Responses(0).Param2

        JudgeRank.AGSRank = SubSystemResult.Responses(0).Param3

        JudgeRank.MainDefectCode = SubSystemResult.Responses(0).Param4

        If JudgeRank.MainDefectCode = "NO_DEFECT" Then
            JudgeRank.FinalOKNG = True
        Else
            JudgeRank.FinalOKNG = False
        End If


        JudgeRank.ReasonCode = SubSystemResult.Responses(0).Param5
        JudgeRank.DefectCode = SubSystemResult.Responses(0).Param6
        'Param7 = strCCD & "$" & strPattern & "$" & strCoordData & "$" & strCoordGate
        Dim Strs() As String
        Strs = SubSystemResult.Responses(0).Param7.Split("$")
        JudgeRank.CCD = Strs(0)
        JudgeRank.Pattern = Strs(1)
        JudgeRank.CoordData = Strs(2)
        JudgeRank.CoordGate = Strs(3)

        Return JudgeRank
    End Function

    Public Sub LogJudgeResult(ByVal InspectStatus As String, ByVal InspectOtherReason As String, ByVal SaveJudgePath As String, ByVal CstRank As String)
        Dim SubSystemResult As CResponseResult
        Me.m_JudgerManager.PrepareAllRequest("LOG_JUDGE_RESULT", InspectStatus, InspectOtherReason, SaveJudgePath, CstRank, , "0", , , )


        Me.m_SystemLog.WriteLog(Now, ", [Cmd=LOG_JUDGE_RESULT]Judger儲存判片結果")

        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.LOG_JUDGE_RESULT)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("儲存判片結果失敗! Communicate with Judger fail! <LOG_JUDGE_RESULT>")
    End Sub

    Public Function MQReport(ByVal InspectStatus As String, ByVal InspectOtherReason As String, ByVal CstRank As String, Optional ByVal MQReportBackupPath As String = "", Optional ByVal blnUseAI As Boolean = True) As String
        Dim SubSystemResult As CResponseResult


        Me.m_JudgerManager.PrepareAllRequest("MQ_REPORT", InspectStatus, InspectOtherReason, CstRank, MQReportBackupPath, , blnUseAI, , Me.ImageProcessResultArray(0, 2))
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=MQ_REPORT]Judger MQ上報")

        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.MQ_REPORT)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then
            Return "MQ上報失敗! Communicate with Judger fail! <MQ_REPORT>  => " & SubSystemResult.ErrMessage
        Else
            Return ""
        End If

    End Function

    Public Sub MQSetDate(ByVal InspectDatetime As Date)
        Dim SubSystemResult As CResponseResult

        Me.m_JudgerManager.PrepareAllRequest("SET_DATE", Format(InspectDatetime, "yyyy"), Format(InspectDatetime, "MM"), Format(InspectDatetime, "dd"), Format(InspectDatetime, "HH"), Format(InspectDatetime, "mm"), Format(InspectDatetime, "ss"))
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=MQ_REPORT]Judger Set Date")

        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.SET_DATE)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("MQ Set Date失敗! Communicate with Judger fail! <SET_DATE>")
    End Sub

    Public Sub MQSetPanel_ID(ByVal Panel_ID As String)
        Dim SubSystemResult As CResponseResult

        Me.m_JudgerManager.PrepareAllRequest("PANEL_ID", Panel_ID, [Enum].GetName(GetType(eRunMode), Me.m_Ui.RunMode))
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=PANEL_ID]Judger Panel ID info.")

        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.PANEL_ID)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("Panel ID通知失敗! Communicate with Judger fail! <PANEL_ID>")
    End Sub

    Private Sub LaserReport(ByVal LaserFilePath As String, ByVal CstRank As String)
        Dim SubSystemResult As CResponseResult

        Me.m_JudgerManager.PrepareAllRequest("LASER_REPORT", LaserFilePath, CstRank)
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=LASER_REPORT]Judger LASER上報")

        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.LASER_REPORT)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("LASER上報失敗! Communicate with Judger fail! <LASER_REPORT>")
    End Sub

    Private Function QueryPointCount() As ClsPointCount
        Dim SubSystemResult As CResponseResult
        Dim PointCount As New ClsPointCount

        Me.m_JudgerManager.PrepareAllRequest("GET_FALSE_DEFECT_COUNT")
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=GET_FALSE_DEFECT_COUNT]Judger查詢FalseDefect數")

        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.GET_FALSE_DEFECT_COUNT)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("查詢Judger亮暗點總數失敗! Communicate with Judger fail! <GET_FALSE_DEFECT_COUNT>")

        PointCount.DP = System.Convert.ToInt32(SubSystemResult.Responses(0).Param1)
        PointCount.BP = System.Convert.ToInt32(SubSystemResult.Responses(0).Param2)

        Return PointCount
    End Function

    Private Sub QueryS06JudgerPointCount(ByRef point() As Integer)
        Dim SubSystemResult As CResponseResult
        'Dim PointCount As New ClsPointCount

        Me.m_JudgerManager.PrepareAllRequest("GET_FALSE_DEFECT_COUNT")
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=GET_FALSE_DEFECT_COUNT]Judger查詢FalseDefect數")

        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.GET_FALSE_DEFECT_COUNT)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("查詢Judger亮暗點總數失敗! Communicate with Judger fail! <GET_FALSE_DEFECT_COUNT>")

        Dim Strs() As String
        Strs = SubSystemResult.Responses(0).Param1.Split("@")
        '
        point(0) = System.Convert.ToInt32(Strs(0))
        point(1) = System.Convert.ToInt32(Strs(1))

        Strs = SubSystemResult.Responses(0).Param2.Split("@")
        point(2) = System.Convert.ToInt32(Strs(0))
        point(3) = System.Convert.ToInt32(Strs(1))

    End Sub

    Private Function QueryGrabberPointCount() As ClsFalseCount
        Dim SubSystemResult As CResponseResult
        Dim FalseCount As New ClsFalseCount
        Dim i As Integer

        Me.m_AreaGrabberManager.PrepareAllRequest("GET_FALSE_DEFECT_COUNT", , , , , , , , , Me.m_TimeoutRecipe.Grabber.GET_FALSE_DEFECT_COUNT)
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=GET_FALSE_DEFECT_COUNT]Grabber查詢FalseDefect數")

        SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.GET_FALSE_DEFECT_COUNT)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("查詢Grabber False點數失敗! Communicate with Grabber fail! <GET_FALSE_DEFECT_COUNT>")
        For i = 0 To SubSystemResult.Responses.Length - 1
            FalseCount.Func += System.Convert.ToInt32(SubSystemResult.Responses(i).Param1)
            FalseCount.Mura += System.Convert.ToInt32(SubSystemResult.Responses(i).Param2)
        Next i

        Return FalseCount
    End Function

    Private Sub QueryS06GrabberPointCount(ByRef FunFalseCount() As Integer, ByRef MuraFalseCount() As Integer)
        Dim SubSystemResult As CResponseResult
        Dim i As Integer

        Me.m_AreaGrabberManager.PrepareAllRequest("GET_FALSE_DEFECT_COUNT", , , , , , , , , Me.m_TimeoutRecipe.Grabber.GET_FALSE_DEFECT_COUNT)
        Me.m_SystemLog.WriteLog(Now, ", [Cmd=GET_FALSE_DEFECT_COUNT]Grabber查詢FalseDefect數")

        SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.GET_FALSE_DEFECT_COUNT)
        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then Throw New Exception("查詢Grabber False點數失敗! Communicate with Grabber fail! <GET_FALSE_DEFECT_COUNT>")

        For i = 0 To SubSystemResult.Responses.Length - 1
            FunFalseCount(i) = System.Convert.ToInt32(SubSystemResult.Responses(i).Param1)
            MuraFalseCount(i) = System.Convert.ToInt32(SubSystemResult.Responses(i).Param2)
        Next i
    End Sub

    Private Sub ClearDefectWindow()
        For Each item As Control In Me.m_MainViewPictureBox.Controls
            If TypeOf (item) Is FrmShowDefect Then
                Me.m_MainViewPictureBox.Controls.Remove(item)
            End If
        Next
    End Sub

#End Region

#Region "About Flow"

    ' 檢測ImgProc Recipe是否合法
    Private Function CheckImgProcRecipe() As Boolean
        Dim i As Integer

        For i = 0 To Me.m_CurrentModel.UsedPatterns.Count - 1
            If Not (Me.m_CurrentModel.UsedPatterns(i).IMG_PROC_RECIPE.Substring(0, 1).ToUpper = "F" Or Me.m_CurrentModel.UsedPatterns(i).IMG_PROC_RECIPE.Substring(0, 1).ToUpper = "M") Then
                Return False
            End If
        Next i
        Return True
    End Function

    Public Sub EditInspectionFlow()
        Dim dlgFlow As DlgInspectionFlow
        dlgFlow = New DlgInspectionFlow(Me.m_BootRecipe.INSPECTION_FLOW_FILE, Me.m_BootRecipe.GENERATION, Me.BootRecipe.PHASE)
        dlgFlow.ShowDialog()

        Me.LoadInspectionFlow()
 
        Me.m_bLoadRecipeFinished = False
    End Sub

    Public Sub ShowAIWorkEditor()
        If Me.m_AIWorker Is Nothing Then Throw New Exception("[ShowAIWorkEditor][ERR]人工智能功能尚未开启 !!")
        If Me.m_AIWorker.LoadPath = String.Empty Then Throw New Exception("[ShowAIWorkEditor][AI_LOAD_PATH]It isn't to set load path!!")
        Me.m_AIWorker.ShowDialog()
        Me.m_AIWorker.Init()
    End Sub

    Public Sub LoadInspectionFlow()

        If System.IO.File.Exists(Me.m_BootRecipe.INSPECTION_FLOW_FILE) Then
            Try
                Me.m_InspectionFlow.ReadXML(Me.m_BootRecipe.INSPECTION_FLOW_FILE)
            Catch ex As Exception
                Me.m_SystemLog.WriteLog(Now, ", [Error]" & ex.Message)
                Throw New Exception("載入Inspection-Flow檔案錯誤! Load InspectionFlow file: " & Me.m_BootRecipe.INSPECTION_FLOW_FILE & " fail!")
            End Try

            If Me.m_InspectionFlow.hasCurrentModel Then
                Me.m_CurrentModel = Me.m_InspectionFlow.getCurrentModel
                RaiseEvent UpdateInspectionFlowInfo()
            Else
                Throw New Exception("InspectionFlow未設定Run貨的產品!" & vbCrLf & vbCrLf & "There is no InspectionFlow Model for Inspection!")
            End If

        Else
            Throw New Exception("Inspection-Flow檔案不存在! InspectionFlow file: " & Me.m_BootRecipe.INSPECTION_FLOW_FILE & " doesn't exist!")
        End If
    End Sub

#End Region

#Region "About Auth"
#Region "ChangeSubUnitAuth"

    Public Sub ChangeSubUnitAuth(ByVal userRole As String, ByVal userId As String)
        Dim SubSystemResult As CResponseResult

        Try
            If userRole = "-2" Then
                userRole = "-1"
                userId = "0000000"
            End If

            If Me.m_AreaGrabberManager IsNot Nothing Then
                If Me.m_AreaGrabberManager.G1IsConnect Or Me.m_AreaGrabberManager.G2IsConnect Then

                    Me.m_AreaGrabberManager.PrepareAllRequest("AUTH_INFO", userRole, userId, , , , , , , Me.m_TimeoutRecipe.Grabber.AUTH_INFO)
                    SubSystemResult = Me.m_AreaGrabberManager.SendRequest(Me.m_TimeoutRecipe.Grabber.AUTH_INFO)

                    If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then
                        RaiseEvent StatusMsg("Area-Grabber權限變更失敗, 請重新啟動Area-Grabber!")

                        Me.m_SystemLog.WriteLog(Now, ", [Error]Area-Grabber權限變更失敗!")
                        Throw New Exception("Area-Grabber權限變更失敗!")
                    End If
                End If
            End If

            If Me.m_Ui.blnJudge Then
                If Me.m_JudgerManager IsNot Nothing Then
                    If Me.m_JudgerManager.IsConnect Then

                        Me.m_JudgerManager.PrepareAllRequest("AUTH_INFO", userRole, userId)
                        SubSystemResult = Me.m_JudgerManager.SendRequest(Me.m_TimeoutRecipe.Judger.AUTH_INFO)

                        If SubSystemResult.CommResult <> AUO.SubSystemControl.eCommResult.OK Then
                            RaiseEvent StatusMsg("Judger權限變更失敗, 請重新啟動Judger!")

                            Me.m_SystemLog.WriteLog(Now, ", [Error]Judger權限變更失敗!")
                            Throw New Exception("Judger權限變更失敗!")
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            Throw New Exception("[ChangeSubUnitAuth] " & ex.Message)
        End Try
    End Sub

#End Region

    Public Function Login(ByRef errMsg As String) As Boolean
        Dim blnRtn As Boolean = False

        'If Not Me.m_BootRecipe.USE_AUTH Then Exit Function

        Try
            Me.m_NewAuth.ShowDialog()

            If Me.m_AuthUserInfo.Result Then
                blnRtn = True
                'Me.m_Role = Me.m_NewAuth.AuthResult.Role
                'Me.m_UserId = Me.m_NewAuth.AuthResult.UserID

                errMsg = ""

                'Me.CreateSubUnit()
                'Me.ChangeSubUnitAuth(Me.m_Role, Me.m_UserId)
                'Me.m_AuthLog.WriteLog(Now, ", [Login] User(" & Me.m_UserId & ")登入成功.")

                '' 主Server掛掉
                'If Me.m_NewAuth.AuthResult.ServerIsDown Then
                '    Me.m_AuthLog.WriteLog(Now, ", [Login] 本次登入成功, 但主伺服器連結失敗: URL為" & Me.m_Auth.m_AuthResult.ErrMessage & ", 請檢查該伺服器!")
                '    Throw New Exception("本次登入成功, 但主伺服器連結失敗: URL為" & Me.m_Auth.m_AuthResult.ErrMessage & ", 請檢查該伺服器!")
                'End If

            Else
                blnRtn = False
                errMsg = Me.m_AuthUserInfo.ErrMessage
                'Me.m_UserId = Me.m_NewAuth.AuthResult.UserID
                Me.m_AuthLog.WriteLog(Now, ", [Login] User(" & Me.m_AuthUserInfo.UserID & ")登入失敗! " & errMsg)
            End If

        Catch ex As Exception
            Me.m_AuthLog.WriteLog(Now, ", [Error] Login Exception" & ex.Message)
            Me.m_SystemLog.WriteLog(Now, ", [Error] Login Exception" & ex.Message)
            Throw New Exception("[Login] Login fail! " & ex.Message)

        Finally
            'Me.DisconnectSubUnit()
        End Try
        Return blnRtn
    End Function

    Public Sub Logout()

        'If Not Me.m_BootRecipe.USE_AUTH Then Exit Sub

        Try
            Me.m_NewAuth.LogOutAUTH()
            'Me.CreateSubUnit()
            'Me.ChangeSubUnitAuth("-1", "0000000")
            Me.m_AuthLog.WriteLog(Now, ", [Logout]User(" & Me.m_AuthUserInfo.UserID & ")登出成功.")

        Catch ex As Exception
            Throw New Exception("[Logout] Logout fail! " & ex.Message)

        Finally
            'Me.DisconnectSubUnit()
        End Try
    End Sub

    Public Sub ChangePassword()

        'If Not Me.m_BootRecipe.USE_AUTH Then Exit Sub

        Try
            Me.m_NewAuth.openChangePWDialog()
            Me.m_AuthLog.WriteLog(Now, ", [ChangePassword]User(" & Me.m_AuthUserInfo.UserID & ")變更密碼.")

        Catch ex As Exception
            Me.m_AuthLog.WriteLog(Now, ", [Error] ChangePassword Exception," & ex.Message)
            Me.m_SystemLog.WriteLog(Now, ", [Error] ChangePassword Exception," & ex.Message)
            Throw New Exception("[ChangePassword] " & ex.Message)
        End Try
    End Sub

#End Region

#Region "---LogIn Component---"
    Private Sub AUTH_OnLogIn(sender As ClsAuthResult)
        '依據使用者需要自行編寫
        Me.m_AuthUserInfo.Result = sender.Result
        Me.m_AuthUserInfo.Role = sender.Role
        Me.m_AuthUserInfo.RoleName = sender.RoleName
        Me.m_AuthUserInfo.UserID = sender.UserID
        If Me.m_AuthUserInfo.Result = False Then
            MessageBox.Show(sender.ErrMessage)
        End If
        RaiseEvent OnLogIn()
    End Sub
    Private Sub AUTH_OnLogOut(sender As ClsAuthResult)
        '依據使用者需要自行編寫

        RaiseEvent OnLogOut()
    End Sub
    Private Sub AUTH_OnShowDialog()
        '依據使用者需要自行編寫
    End Sub
    Public Sub AuthBootConfigSetting()
        Me.m_NewAuth.changeBootData()
    End Sub
#End Region

    Private Sub m_RS232_DevBoard_OnReceiveData(OutputData As Object)
        Try
            Dim tmpData As String = Convert.ToString(OutputData)
            Dim tmpStr() As String = tmpData.TrimEnd().Split(",")
            Dim cmd As String = tmpStr(0)
            Select Case cmd
                Case "STA"
                    Me.IsStartInspect_ForRS232 = IIf(tmpStr(1) = "1", True, False)
            End Select
            Me.m_SystemLog.WriteLog(Now, ", [DevBoard_OnReceiveData] Data : " & tmpData)
        Catch ex As Exception
            Me.m_ErrorLog.WriteLog(Now, ",  [DevBoard_OnReceiveData] Error!! MSG : " & ex.Message)
        End Try

    End Sub

    Private Sub WaitInspectionSignal()
        Try
            While Not Me.m_StopInspection
                If Me.IsStartInspect_ForRS232 Then
                    Me.m_SystemLog.WriteLog(Now, ", [Info]Panel In AOI Position now!!")
                    Me.m_StartTime = Environment.TickCount          ' 從Contact後開始計算檢測時間
                    Me.IsStartInspect_ForRS232 = False
                    Exit Sub
                End If
                Threading.Thread.Sleep(100)
            End While
        Catch ex As Exception
            Me.m_SystemLog.WriteLog(Now, ", [Error]Wait Panel rotation to AOI Position Exception => " & ex.Message)
            Me.m_PLCCommExMsg = ex.Message
        End Try
    End Sub

    Private Class CJudgeRank
        Public CstRank As String
        Public MqRank As String
        Public AGSRank As String
        Public MainDefectCode As String
        Public ReasonCode As String
        Public DefectCode As String
        Public CCD As String
        Public Pattern As String
        Public CoordData As String
        Public CoordGate As String
        Public DefectType As String 'S03 SKD,S16 CT1
        Public FinalOKNG As Boolean


        Public Sub New()
            CstRank = ""
            MqRank = ""
            AGSRank = ""
            MainDefectCode = ""
            ReasonCode = ""
            DefectCode = ""
            CCD = ""
            Pattern = ""
            CoordData = ""
            CoordGate = ""
            DefectType = ""
        End Sub
    End Class

    Private Class ClsPointCount
        Public BP As Integer
        Public DP As Integer
    End Class

    Private Class ClsFalseCount
        Public Func As Integer
        Public Mura As Integer
    End Class

    Private Function m_Picture_Datie() As String
        Throw New NotImplementedException
    End Function

    Private Function ex() As Object
        Throw New NotImplementedException
    End Function

End Class

Public Class ProceRest

    Dim PanelId As String
    Dim InspectDate As Date

    Sub New(ByVal sPanelId As String, ByVal sInspectDate As Date)
        PanelId = sPanelId
        InspectDate = sInspectDate

    End Sub
End Class

'Public Enum eROLE As Integer
'    DBdown = -2
'    UnAuth = -1
'    OP = 0
'    PM = 1
'    ENG = 2
'    SUPER = 3
'End Enum

Public Enum eRunMode As Byte
    Sim = 0
    Auto = 1
    LoadImage = 2
    AutoAdjustExp = 3
    LoadImageInspectAfterDemura = 4
End Enum
