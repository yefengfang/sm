VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_LockStock 
   Caption         =   "锁库明晰"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   13770
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton bt_Close 
      Caption         =   "退出"
      Height          =   615
      Left            =   12000
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton bt_UnLock 
      Caption         =   "解锁"
      Height          =   615
      Left            =   12000
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton bt_LockStock 
      Caption         =   "锁库"
      Height          =   615
      Left            =   12000
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dg_skxx 
      Height          =   3615
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   6376
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dg_djxx 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5953
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_LockStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim billinfo As New Adodb.Recordset
Private Sub bt_Close_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim obill() As String
    Dim connStr As String
    '获取选择的单据内码和分录号
    Call ReadFile(obill, connStr)
    
    '连接数据库
    
    Dim cn As New Adodb.Connection
    cn.ConnectionString = connStr
    cn.CursorLocation = adUseClient
    cn.Open
    
    Dim CMD As New Adodb.Command
    Dim sql As String
    
    sql = "" & _
" SELECT" & _
"     O.FBillNo AS FBillNo," & _
"     I1.FName AS FCustName," & _
"     ICI.FNumber AS FNumber," & _
"     ICI.FName AS FName," & _
"     U.FName AS FUnit," & _
"     OE.FQty AS FQty," & _
"     CASE WHEN ISNULL(L.FQty,0)>=OE.FQty THEN 'Y' ELSE '' END AS FLockState " & _
" FROM SEOrder AS O " & _
" INNER JOIN SEOrderEntry AS OE ON O.FInterID=OE.FInterID " & _
" INNER JOIN t_Item AS I1 ON I1.FItemID=O.FCustID " & _
" INNER JOIN t_ICItem AS ICI ON ICI.FItemID=OE.FItemID " & _
" INNER JOIN t_MeasureUnit AS U ON U.FItemID=ICI.FUnitID " & _
" LEFT JOIN t_LockStock AS L ON OE.FInterID=L.FInterID AND OE.FEntryID=L.FEntryID AND OE.FItemID=L.FItemID "

    Dim where As String
    where = makeWhere_orderbill(obill)
'    Dim rs As New ADODB.Recordset
'    CMD.CommandText = sql + where
'    CMD.ActiveConnection = cn
'
'    Set billinfo = CMD.Execute
    
    Dim billinfo As New Adodb.Recordset
    billinfo.Open sql + where, cn
    
    
    'Set dg_djxx.DataSource = Nothing
    Set dg_djxx.DataSource = billinfo


   
    
    
    
    
    
    
    
    
    
    
End Sub
Public Function makeWhere_orderbill(arr() As String) As String
    Dim where As String
    where = " WHERE 1=0 "
    For I = 0 To UBound(arr)
        where = where + "OR (O.FInterID=" + arr(I, 0) + " AND OE.FEntryID=" + arr(I, 1) + ") "
    Next
End Function
Public Function ReadFile(obill() As String, connStr As String) As String
    If Dir(App.Path & "\temp.txt") = "" Then
    '不存在
        MsgBox ("临时文件未找到，不能进行锁库！")
        Unload Me
    Else
    '存在
        Open (App.Path & "\temp.txt") For Input As #1
        Dim text As String
        Do While Not EOF(1)
            Input #1, b
            text = text & b
        Loop
        Close #1
    End If
    
    Dim s1() As String
    Dim s2() As String
    s1() = Split(text, "I")
    ReDim Bill(UBound(s1) - 1, 1) As String
    
    For I = 1 To UBound(s1)
        s2() = Split(s1(I), "E")
        Bill(I - 1, 0) = s2(0)
        Bill(I - 1, 1) = s2(1)
    Next
    obill = Bill
    
    If Dir(App.Path & "\conn.txt") = "" Then
    '不存在
        MsgBox ("临时文件未找到，不能进行锁库！")
        Unload Me
    Else
    '存在
        connStr = ""
        Open (App.Path & "\conn.txt") For Input As #1
        Do While Not EOF(1)
            Input #1, b
            connStr = connStr & Split(b, "|")(0)
        Loop
        Close #1
    End If
    Kill (App.Path & "\temp.txt")
    Kill (App.Path & "\conn.txt")
End Function


            
        'ByVal sKey As String, oList As Object, ByRef bCancel As Boolean
        '通过Set vectBill = oList.GetSelected 可以获取当前选中序时薄数据
        
        '返回记录集方式
        'Set rs = obj.Execute("select * from t_icitem")
        
        '执行存储过程方式
        ' obj.Execute3 ("exec KY_PlanQty")
        
        
        
        
        
        
            'Dim vectBill As KFO.Vector
    'Dim lmul As Long
    ' Dim rs As ADODB.Recordset
    'Dim InBatch As Form
    'Set InBatch = New InBatch
    
    'Set OBJ = CreateObject("K3Connection.AppConnection")
    
