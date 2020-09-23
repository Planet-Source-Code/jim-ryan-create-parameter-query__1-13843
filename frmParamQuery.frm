VERSION 5.00
Begin VB.Form frmParamQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get Selected Colors"
   ClientHeight    =   2796
   ClientLeft      =   1092
   ClientTop       =   336
   ClientWidth     =   5400
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2796
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   4200
      TabIndex        =   4
      Top             =   1680
      Width           =   972
   End
   Begin VB.ListBox lstNew 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2160
      ItemData        =   "frmParamQuery.frx":0000
      Left            =   240
      List            =   "frmParamQuery.frx":0002
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   3732
   End
   Begin VB.CommandButton cmdGetColors 
      Caption         =   "Get Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   4200
      TabIndex        =   2
      Top             =   552
      Width           =   972
   End
   Begin VB.ListBox lstColors 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2208
      ItemData        =   "frmParamQuery.frx":0004
      Left            =   240
      List            =   "frmParamQuery.frx":0006
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   3732
   End
   Begin VB.Label Label1 
      Caption         =   "Color Selector"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1932
   End
End
Attribute VB_Name = "frmParamQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       ParamQuery
' FORM:         frmParamQuery
' DATABASE:     Text.mdb
' TABLE:        Color
' AUTHOR:       Jim Ryan
' EMAIL:        jprism@prism2000.net
' CREATED:      24-Dec-2000
'
' DESCRIPTION:
'    This program demonstrates how one can
'    allow a user to:
'    1) Make one or more choices from a complete
'       recordset (at runtime)
'    2) Build a Parameter Query based on the selected
'       items
'    3) Run the Parameter Query
'
' MODIFICATION HISTORY:
' 1.0       24-Dec-2000
'           Jim Ryan
'           Initial Version
'*******************************************************************************

Option Explicit
Private Rst As adodb.Recordset
Attribute Rst.VB_VarHelpID = -1

Private Sub cmdGetColors_Click()
   Dim Cmd As adodb.Command
   Dim Prm As adodb.Parameter
   Dim Cnn As adodb.Connection
   Dim Sql As String
   Dim Colors() As Integer
   Dim i As Integer
   Dim c As Integer

   On Error GoTo ErrorTrap
   
   ' check to see if user has selected any
   ' colors from the lstcolors list
   If lstColors.SelCount = 0 Then
      MsgBox "You must first select one or more colors, then click the Get Colors button"
      Exit Sub
   End If
   
   Label1.Caption = "Selected Colors"
   
   ' load the colors array with
   ' the itemdata i.e. colorno
   For i = 0 To lstColors.ListCount - 1
      If lstColors.Selected(i) Then
         ReDim Preserve Colors(c)
         Colors(c) = lstColors.ItemData(i)
         c = c + 1
      End If
   Next
   
   ' Open a connection to the
   ' Color.mdb database
   Set Cnn = New Connection
   Cnn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source= " & App.Path & "\Color.mdb;"

   ' Setup the new command and
   ' query
   Set Cmd = New Command
   Cmd.CommandType = adCmdText
   ' build SQL string for all selected
   ' colors
   Sql = "SELECT  [color].[colorno], [color].[colorname] " & _
         "FROM [color] WHERE"
   For i = 0 To UBound(Colors)
      Sql = Sql & " [color].[colorno] = ? OR"
   Next
   Sql = Left(Sql, Len(Sql) - 3) & " ORDER BY [color].[colorname]"
   Cmd.CommandText = Sql
   Cmd.Name = "adoCommand"
    
   ' set each colors parameter to its
   ' colorno
   For i = 0 To UBound(Colors)
      Set Prm = New Parameter
      Prm.Name = "colorno"
      Prm.Type = adInteger
      Prm.Size = 4
      Prm.Value = Colors(i)
      Cmd.Parameters.Append Prm
      Set Prm = Nothing
   Next

   ' set active connection and
   ' execute recordset
   Set Cmd.ActiveConnection = Cnn
   Set Rst = Cmd.Execute
   
   ' move through the recordset
   ' and add the selected colors
   ' to the lstnew list
   lstNew.Clear
   While Not Rst.EOF
      lstNew.AddItem Rst![colorname]
      Rst.MoveNext
   Wend
   lstColors.Visible = False
   lstNew.Visible = True
   
ExitOk:
   On Error Resume Next
   Set Cmd = Nothing
   Set Prm = Nothing
   Rst.Close
   Set Rst = Nothing
   Cnn.Close
   Set Cnn = Nothing
   On Error GoTo 0
   Exit Sub
   
ErrorTrap:
   MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
   Resume ExitOk
End Sub

Private Sub cmdReset_Click()
   Dim i As Integer
   
   ' reset the selected items in
   ' the lstcolors list
   For i = 0 To lstColors.ListCount - 1
      lstColors.Selected(i) = False
   Next
   ' clear the lstnew list and make
   ' invisible
   lstNew.Clear
   lstNew.Visible = False
   ' show the lstcolors list
   lstColors.Visible = True
End Sub

Private Sub Form_Load()
   Dim Src As String
   Dim Sql As String
   Dim MDBname As String

   On Error GoTo ErrorTrap
   
   ' check for the existence of the
   ' Color.mdb database
   MDBname = App.Path & "\Color.mdb"
   If Dir(MDBname, vbNormal) = "" Then
      MsgBox "The MDB database Color.mdb could not be found!"
      Exit Sub
   End If
   
   ' Open a recordset on the color table
   ' Src = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source= " & MDBname
   Src = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & MDBname
   Sql = "SELECT [color].[colorno], [color].[colorname] " & _
         "FROM [color] ORDER BY [color].[colorname]"
   Set Rst = New adodb.Recordset
   Rst.CursorLocation = adUseClient
   Rst.Open Sql, Src, adOpenForwardOnly, adLockReadOnly
   lstColors.Clear
   
   ' fill the lstColors list with
   ' the colorno and colorname columns
   While Not Rst.EOF
      lstColors.AddItem Rst![colorname]
      lstColors.ItemData(lstColors.NewIndex) = Rst![colorno]
      Rst.MoveNext
   Wend
   
   ' notify the user if the lstcolors list
   ' contains no records
   If lstColors.ListCount < 0 Then
      MsgBox "The table Color contains NO records..."
   End If

ExitOk:
   On Error Resume Next
   Rst.Close
   Set Rst = Nothing
   On Error GoTo 0
   Exit Sub
   
ErrorTrap:
   MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
   Resume ExitOk
End Sub
