VERSION 5.00
Begin {17016CEE-E118-11D0-94B8-00A0C91110ED} wbcPortal 
   ClientHeight    =   7860
   ClientLeft      =   750
   ClientTop       =   1425
   ClientWidth     =   7725
   _ExtentX        =   13626
   _ExtentY        =   13864
   MajorVersion    =   0
   MinorVersion    =   8
   StateManagementType=   1
   ASPFileName     =   ""
   DIID_WebClass   =   "{12CBA1F6-9056-11D1-8544-00A024A55AB0}"
   DIID_WebClassEvents=   "{12CBA1F5-9056-11D1-8544-00A024A55AB0}"
   TypeInfoCookie  =   30
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   3
      BeginProperty WebItem1 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "tplNewUser"
         DISPID          =   1280
         Template        =   "NewUser1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{7B1277D0-9529-11D3-BB26-00105A1BDA1E}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "F:\SourceCode\Articles\Inside ASP\Web Class\test\NewUser.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   1
            BeginProperty Event0 {193556D3-4486-11D1-9C70-00C04FB987DF} 
               Name            =   "Form1"
               DISPID          =   1280
               Type            =   1
               OriginalHREF    =   ""
               TagType         =   6619241
               BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
                  AttribCount     =   1
                  BeginProperty Attrib0 {FA6A55FC-458A-11D1-9C71-00C04FB987DF} 
                     TagType         =   1
                     Attribute       =   "ACTION"
                     State           =   3
                     TagName         =   "Form1"
                     OriginalURL     =   "default.ASP?WCI=tplNewUser&WCE=Form1&WCU"
                     Parent          =   ""
                     Template        =   "tplNewUser"
                     BoundEvent      =   "Form1"
                     BoundItem       =   ""
                     Suffix          =   ""
                     UsesAnonymousName=   -1
                     TagNumber       =   1
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem2 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "tplPortal"
         DISPID          =   1281
         Template        =   "Portal1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{7B1277BB-9529-11D3-BB26-00105A1BDA1E}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "F:\SourceCode\Articles\Inside ASP\Web Class\test\Portal.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem3 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "tplWelcome"
         DISPID          =   1282
         Template        =   "Welcome1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{F64C0728-94BE-11D3-BB22-00105A1BDA1E}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "F:\SourceCode\Articles\Inside ASP\Web Class\test\Welcome.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   1
            BeginProperty Event0 {193556D3-4486-11D1-9C70-00C04FB987DF} 
               Name            =   "Form1"
               DISPID          =   1280
               Type            =   1
               OriginalHREF    =   ""
               TagType         =   6619241
               BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
                  AttribCount     =   1
                  BeginProperty Attrib0 {FA6A55FC-458A-11D1-9C71-00C04FB987DF} 
                     TagType         =   1
                     Attribute       =   "ACTION"
                     State           =   3
                     TagName         =   "Form1"
                     OriginalURL     =   "default.ASP?WCI=tplWelcome&WCE=Form1&WCU"
                     Parent          =   ""
                     Template        =   "tplWelcome"
                     BoundEvent      =   "Form1"
                     BoundItem       =   ""
                     Suffix          =   ""
                     UsesAnonymousName=   -1
                     TagNumber       =   1
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
   EndProperty
   NameInURL       =   "Portal"
End
Attribute VB_Name = "wbcPortal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private mconConnection As ADODB.Connection
Private mrsUser As ADODB.Recordset

Private mstrUser As String
Private mstrFavoriteURL As String
Private mdteDate As Date

Private Sub tplNewUser_Form1()
    '*****************************
    'save user info to database
    '*****************************
    
    'init
    Set mconConnection = New ADODB.Connection
    Set mrsUser = New ADODB.Recordset
    
    mconConnection.Open "DBQ=" & App.Path & "/portal.mdb" & ";DRIVER={Microsoft Access Driver (*.mdb)}"
    mrsUser.Open "SELECT * from USER where" & vbCrLf & _
        "name=''", mconConnection, adOpenForwardOnly, adLockPessimistic
        
    mrsUser.AddNew
    mrsUser("name") = Request("txtName")
    mrsUser("password") = Request("txtPassword")
    mrsUser("FavoriteURL") = Request("txtFavoriteURL")
    mrsUser.Update
    
    mrsUser.Close
    mconConnection.Close
    
    tplPortal.WriteTemplate
    
End Sub
Private Sub tplNewUser_ProcessTag(ByVal TagName As String, TagContents As String, SendTags As Boolean)
    Select Case (TagName)
        Case "WC@txtName"
            TagContents = Request("txtName")
            
        Case "WC@txtPassword"
            TagContents = Request("txtPassword")
            
    End Select
End Sub

Private Sub tplPortal_ProcessTag(ByVal TagName As String, TagContents As String, SendTags As Boolean)

    Select Case (TagName)
        Case "WC@Init"
            '*****************************
            'read user info from database
            '*****************************
            
            'init
            Set mconConnection = New ADODB.Connection
            Set mrsUser = New ADODB.Recordset
            
            mconConnection.Open "DBQ=" & App.Path & "/portal.mdb" & ";DRIVER={Microsoft Access Driver (*.mdb)}"
            mrsUser.Open "SELECT * from USER where" & vbCrLf & _
                "name='" & Request("txtName") & "'" & vbCrLf & _
                "AND password='" & Request("txtPassword") & "'", mconConnection
                
            
            mstrUser = mrsUser("name")
            mstrFavoriteURL = mrsUser("FavoriteURL")
            mdteDate = mrsUser("SignUpDate")
            
            mrsUser.Close
            mconConnection.Close
            TagContents = ""
            
        Case "WC@txtName"
            TagContents = mstrUser
        Case "WC@dteDate"
            TagContents = mdteDate
        Case "WC@txtFavoriteURL"
            TagContents = "<a href=" & Chr(34) & mstrFavoriteURL _
                & Chr(34) & ">" & mstrFavoriteURL & _
                "</a>"
    End Select

End Sub

Private Sub tplWelcome_Form1()
    
    '*******************************************
    'validate user id and password from database
    '*******************************************
    
    'init
    Set mconConnection = New ADODB.Connection
    Set mrsUser = New ADODB.Recordset
    
    mconConnection.Open "DBQ=" & App.Path & "/portal.mdb" & ";DRIVER={Microsoft Access Driver (*.mdb)}"
    mrsUser.Open "SELECT * from USER where" & vbCrLf & _
        "name='" & Request("txtName") & "'" & vbCrLf & _
        "AND password='" & Request("txtPassword") & "'", mconConnection
        
    If mrsUser.EOF Then
        'user not registered
        'show new user screen
        
        mrsUser.Close
        mconConnection.Close
        
        tplNewUser.WriteTemplate
    Else
        'user registered
        'show portal screen
        
        mrsUser.Close
        mconConnection.Close
        
        tplPortal.WriteTemplate
    End If
    
    
End Sub

Private Sub WebClass_Start()
    
    'show default class
    tplWelcome.WriteTemplate


End Sub
