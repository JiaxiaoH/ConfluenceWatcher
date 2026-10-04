Attribute VB_Name = "ConfluenceWatcher"
Option Explicit

' Sheets are referenced by position so that Japanese/Chinese sheet names
' cannot break the module:
'   1 = Notes, 2 = Config, 3 = CurrentState, 4 = History, 5 = Changes, 6 = RunLog
Private Const CFG_SHEET_INDEX As Long = 2
Private Const STATE_SHEET_INDEX As Long = 3
Private Const HIST_SHEET_INDEX As Long = 4
Private Const CHANGE_SHEET_INDEX As Long = 5
Private Const LOG_SHEET_INDEX As Long = 6

' Cached session cookie obtained from /rest/auth/1/session
Private mSessionCookie As String

' A single shared WinHttp instance. The Jira tool that works in this
' environment reuses one object, which keeps the cookie jar alive between
' the login call and every later request. Creating a new object per request
' discards the session that the gateway handed out.
Private mHttp As Object
Private mLoginHeaders As String

Private Function GetHttp() As Object
    If mHttp Is Nothing Then
        Set mHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If
    Set GetHttp = mHttp
End Function

' Drops the shared instance so the next call starts a brand new session.
Public Sub ResetConfluenceSession()
    Set mHttp = Nothing
    mSessionCookie = ""
    mLoginHeaders = ""
    MsgBox "Session reset. The next run will log in again.", vbInformation, "Confluence Watcher"
End Sub

'==================================================================
' Main entry point
'==================================================================
Public Sub CheckConfluenceTree()
    Dim stage As String
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Dim savedErrSource As String
    On Error GoTo EH

    stage = "Initialize"
    Application.ScreenUpdating = False
    Application.StatusBar = "Scanning the Confluence page tree..."

    Dim baseUrl As String, rootId As String, authMode As String
    Dim userName As String, secret As String, pageSize As Long, timeoutSec As Long

    stage = "Read Config sheet"
    baseUrl = Trim$(CStr(Cfg(4, 2)))
    rootId = Trim$(CStr(Cfg(5, 2)))
    authMode = Trim$(CStr(Cfg(6, 2)))
    userName = Trim$(CStr(Cfg(7, 2)))
    secret = GetSecret(authMode, CStr(Cfg(8, 2)))
    timeoutSec = SafeLong(Cfg(9, 2))
    pageSize = SafeLong(Cfg(10, 2))

    If baseUrl = "" Or rootId = "" Then
        Err.Raise 5, "Config", "Base URL or root page ID is empty."
    End If
    If secret = "" Then
        Err.Raise 5, "Config", "Credential is empty. Set the environment variable or fill cell B8 on the Config sheet."
    End If

    stage = "Authenticate"
    secret = PrepareAuth(authMode, baseUrl, userName, secret, timeoutSec)

    Dim oldState As Object, newState As Object
    stage = "Load current state"
    Set oldState = LoadCurrentState()
    Set newState = CreateObject("Scripting.Dictionary")

    stage = "Call Confluence REST API"
    ScanTree baseUrl, rootId, authMode, userName, secret, timeoutSec, pageSize, newState

    Dim firstRun As Boolean
    firstRun = (oldState.Count = 0)

    Dim changes As Long
    stage = "Compare page versions"
    changes = WriteChanges(oldState, newState, firstRun And Not CBool(Cfg(11, 2)))

    stage = "Write CurrentState sheet"
    RewriteCurrentState newState

    stage = "Write History sheet"
    If CBool(Cfg(13, 2)) Then
        AppendHistory newState
    End If

    stage = "Write RunLog sheet"
    AddRunLog "INFO", "Scan completed successfully", newState.Count, changes

    ThisWorkbook.Save
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Scan completed." & vbCrLf & vbCrLf & _
           "Pages: " & CStr(newState.Count) & vbCrLf & _
           "Changes: " & CStr(changes), vbInformation, "Confluence Watcher"
    Exit Sub

EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    savedErrSource = Err.Source
    On Error Resume Next
    AddRunLog "ERROR", "Stage: " & stage & " | Error: " & CStr(savedErrNumber) & _
              " | Source: " & savedErrSource & " | Description: " & savedErrDescription, 0, 0
    Application.StatusBar = False
    Application.ScreenUpdating = True
    On Error GoTo 0
    If Len(savedErrDescription) = 0 Then
        savedErrDescription = "No error description was returned by VBA."
    End If
    MsgBox "Scan failed." & vbCrLf & vbCrLf & _
           "Stage: " & stage & vbCrLf & _
           "Error number: " & CStr(savedErrNumber) & vbCrLf & _
           "Source: " & savedErrSource & vbCrLf & _
           "Description: " & savedErrDescription, vbCritical, "Confluence Watcher"
End Sub

'==================================================================
' Session login - the approach used by the working Jira tool.
' Posts credentials to /rest/auth/1/session and keeps the JSESSIONID.
'==================================================================
Private Function SessionLogin(ByVal baseUrl As String, ByVal userName As String, _
                              ByVal password As String, ByVal timeoutSec As Long) As String
    Dim h As Object
    Dim payload As String
    Dim rawHeaders As String
    Dim cookie As String
    Dim bodyText As String

    If Len(userName) = 0 Then
        Err.Raise 5, "SessionLogin", "User name is empty. Fill cell B7 on the Config sheet."
    End If
    If Len(password) = 0 Then
        Err.Raise 5, "SessionLogin", "Password is empty. Fill cell B8 on the Config sheet or set CONFLUENCE_PASSWORD."
    End If
    If timeoutSec < 5 Then
        timeoutSec = 30
    End If

    baseUrl = TrimSlash(baseUrl)
    payload = "{""username"":""" & JsonEscape(userName) & """,""password"":""" & JsonEscape(password) & """}"

    Set h = GetHttp()
    h.SetTimeouts timeoutSec * 1000, timeoutSec * 1000, timeoutSec * 1000, timeoutSec * 1000
    h.Open "POST", baseUrl & "/rest/auth/1/session", False
    h.SetRequestHeader "Content-Type", "application/json"
    h.SetRequestHeader "Accept", "application/json"
    h.SetRequestHeader "Origin", baseUrl

    On Error Resume Next
    h.SetProxy 1
    On Error GoTo 0

    h.Send payload

    bodyText = Utf8BytesToText(h.ResponseBody)

    If h.Status <> 200 Then
        Err.Raise vbObjectError + 102, "SessionLogin", _
                  "Login failed. HTTP " & CStr(h.Status) & " " & h.StatusText & _
                  " | Response: " & Left$(bodyText, 300)
    End If

    rawHeaders = ""
    On Error Resume Next
    rawHeaders = h.GetAllResponseHeaders
    On Error GoTo 0

    mLoginHeaders = rawHeaders
    cookie = ExtractSessionCookie(rawHeaders)

    If Len(cookie) = 0 Then
        cookie = CookieFromLoginBody(bodyText)
    End If

    If Len(cookie) = 0 Then
        Err.Raise vbObjectError + 103, "SessionLogin", _
                  "Login returned HTTP 200 but no session cookie was found. Response: " & Left$(bodyText, 300)
    End If

    mSessionCookie = cookie
    SessionLogin = cookie
End Function

' Pulls every Set-Cookie name=value pair out of the raw response headers.
Private Function ExtractSessionCookie(ByVal rawHeaders As String) As String
    Dim parts() As String
    Dim i As Long
    Dim ln As String
    Dim value As String
    Dim semi As Long
    Dim result As String

    If Len(rawHeaders) = 0 Then
        Exit Function
    End If

    rawHeaders = Replace(rawHeaders, vbCrLf, vbLf)
    rawHeaders = Replace(rawHeaders, vbCr, vbLf)
    parts = Split(rawHeaders, vbLf)

    For i = LBound(parts) To UBound(parts)
        ln = Trim$(parts(i))
        If LCase$(Left$(ln, 11)) = "set-cookie:" Then
            value = Trim$(Mid$(ln, 12))
            semi = InStr(value, ";")
            If semi > 0 Then
                value = Left$(value, semi - 1)
            End If
            value = Trim$(value)
            If Len(value) > 0 And InStr(value, "=") > 0 Then
                If Len(result) > 0 Then
                    result = result & "; "
                End If
                result = result & value
            End If
        End If
    Next i

    ExtractSessionCookie = result
End Function

' Fallback: build the cookie from the JSON body when headers are unavailable.
Private Function CookieFromLoginBody(ByVal bodyText As String) As String
    Dim sessionObj As String
    Dim cookieName As String
    Dim cookieValue As String

    sessionObj = JsonObjectForKey(bodyText, "session")
    If Len(sessionObj) = 0 Then
        Exit Function
    End If

    cookieName = JsonString(sessionObj, "name")
    cookieValue = JsonString(sessionObj, "value")

    If Len(cookieName) > 0 And Len(cookieValue) > 0 Then
        CookieFromLoginBody = cookieName & "=" & cookieValue
    End If
End Function

Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, Chr$(34), "\" & Chr$(34))
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

' Converts session mode into a usable cookie before any REST call is made.
Private Function PrepareAuth(ByRef authMode As String, ByVal baseUrl As String, _
                             ByVal userName As String, ByVal secret As String, _
                             ByVal timeoutSec As Long) As String
    If LCase$(authMode) = "session" Then
        PrepareAuth = SessionLogin(baseUrl, userName, secret, timeoutSec)
    Else
        PrepareAuth = secret
    End If
End Function

'==================================================================
' Diagnostic helpers
'==================================================================
Public Sub TestSessionLogin()
    Dim baseUrl As String, userName As String, password As String
    Dim timeoutSec As Long, cookie As String, body As String
    On Error GoTo EH
    baseUrl = Trim$(CStr(Cfg(4, 2)))
    userName = Trim$(CStr(Cfg(7, 2)))
    password = GetSecret("session", CStr(Cfg(8, 2)))
    timeoutSec = SafeLong(Cfg(9, 2))

    Set mHttp = Nothing
    cookie = SessionLogin(baseUrl, userName, password, timeoutSec)
    body = HttpGet(TrimSlash(baseUrl) & "/rest/api/user/current", "session", "", cookie, timeoutSec)

    MsgBox "Session login succeeded." & vbCrLf & vbCrLf & _
           "User: " & userName & vbCrLf & _
           "Cookie captured: " & Left$(cookie, 120) & vbCrLf & vbCrLf & _
           "Verify response: " & Left$(body, 300), vbInformation, "Confluence Watcher"
    Exit Sub
EH:
    MsgBox "Session login failed." & vbCrLf & vbCrLf & _
           "Error number: " & CStr(Err.Number) & vbCrLf & _
           "Source: " & Err.Source & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Confluence Watcher"
End Sub

' Shows the raw Set-Cookie headers returned by the login call.
' Run this if TestSessionLogin still fails.
Public Sub ShowLoginHeaders()
    Dim baseUrl As String, userName As String, password As String
    Dim timeoutSec As Long
    On Error GoTo EH
    baseUrl = Trim$(CStr(Cfg(4, 2)))
    userName = Trim$(CStr(Cfg(7, 2)))
    password = GetSecret("session", CStr(Cfg(8, 2)))
    timeoutSec = SafeLong(Cfg(9, 2))

    Set mHttp = Nothing
    SessionLogin baseUrl, userName, password, timeoutSec

    MsgBox "Login response headers:" & vbCrLf & vbCrLf & _
           Left$(mLoginHeaders, 1200), vbInformation, "Confluence Watcher"
    Exit Sub
EH:
    MsgBox "Login failed." & vbCrLf & vbCrLf & _
           "Error number: " & CStr(Err.Number) & vbCrLf & _
           "Source: " & Err.Source & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & vbCrLf & _
           "Headers captured so far:" & vbCrLf & Left$(mLoginHeaders, 800), _
           vbCritical, "Confluence Watcher"
End Sub

Public Sub TestConfluenceAuth()
    Dim baseUrl As String, authMode As String, userName As String, secret As String
    Dim timeoutSec As Long, body As String
    On Error GoTo EH
    baseUrl = Trim$(CStr(Cfg(4, 2)))
    authMode = Trim$(CStr(Cfg(6, 2)))
    userName = Trim$(CStr(Cfg(7, 2)))
    secret = GetSecret(authMode, CStr(Cfg(8, 2)))
    timeoutSec = SafeLong(Cfg(9, 2))

    If secret = "" Then
        MsgBox "No credential was found." & vbCrLf & _
               "Auth mode: " & authMode, vbExclamation, "Confluence Watcher"
        Exit Sub
    End If

    secret = PrepareAuth(authMode, baseUrl, userName, secret, timeoutSec)
    body = HttpGet(TrimSlash(baseUrl) & "/rest/api/user/current", authMode, userName, secret, timeoutSec)
    MsgBox "Authentication succeeded." & vbCrLf & vbCrLf & _
           "Auth mode: " & authMode & vbCrLf & _
           "Credential length: " & CStr(Len(secret)) & vbCrLf & vbCrLf & _
           "Response: " & Left$(body, 400), vbInformation, "Confluence Watcher"
    Exit Sub
EH:
    MsgBox "Authentication failed." & vbCrLf & vbCrLf & _
           "Error number: " & CStr(Err.Number) & vbCrLf & _
           "Source: " & Err.Source & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Confluence Watcher"
End Sub

' Reads only the root page. Use this to isolate problems without recursion.
Public Sub TestSinglePage()
    Dim baseUrl As String, rootId As String, authMode As String
    Dim userName As String, secret As String, timeoutSec As Long
    Dim p As Object
    On Error GoTo EH
    baseUrl = Trim$(CStr(Cfg(4, 2)))
    rootId = Trim$(CStr(Cfg(5, 2)))
    authMode = Trim$(CStr(Cfg(6, 2)))
    userName = Trim$(CStr(Cfg(7, 2)))
    secret = GetSecret(authMode, CStr(Cfg(8, 2)))
    timeoutSec = SafeLong(Cfg(9, 2))
    secret = PrepareAuth(authMode, baseUrl, userName, secret, timeoutSec)

    Set p = GetPage(baseUrl, rootId, "", authMode, userName, secret, timeoutSec)
    MsgBox "Root page read successfully." & vbCrLf & vbCrLf & _
           "ID: " & p("id") & vbCrLf & _
           "Title: " & p("title") & vbCrLf & _
           "Version: " & CStr(p("version")) & vbCrLf & _
           "Updated: " & p("updatedAt") & vbCrLf & _
           "By: " & p("updatedBy") & vbCrLf & _
           "URL: " & p("url"), vbInformation, "Confluence Watcher"
    Exit Sub
EH:
    MsgBox "Single page read failed." & vbCrLf & vbCrLf & _
           "Error number: " & CStr(Err.Number) & vbCrLf & _
           "Source: " & Err.Source & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Confluence Watcher"
End Sub

' Reads only the direct children of the root page.
Public Sub TestChildPages()
    Dim baseUrl As String, rootId As String, authMode As String
    Dim userName As String, secret As String, timeoutSec As Long, pageSize As Long
    Dim kids As Collection, msg As String, i As Long
    On Error GoTo EH
    baseUrl = Trim$(CStr(Cfg(4, 2)))
    rootId = Trim$(CStr(Cfg(5, 2)))
    authMode = Trim$(CStr(Cfg(6, 2)))
    userName = Trim$(CStr(Cfg(7, 2)))
    secret = GetSecret(authMode, CStr(Cfg(8, 2)))
    timeoutSec = SafeLong(Cfg(9, 2))
    pageSize = SafeLong(Cfg(10, 2))
    secret = PrepareAuth(authMode, baseUrl, userName, secret, timeoutSec)

    Set kids = GetChildren(baseUrl, rootId, authMode, userName, secret, timeoutSec, pageSize)
    msg = "Direct children found: " & CStr(kids.Count) & vbCrLf & vbCrLf
    For i = 1 To kids.Count
        If i > 15 Then
            msg = msg & "..."
            Exit For
        End If
        msg = msg & kids(i)("id") & "  v" & CStr(kids(i)("version")) & "  " & kids(i)("title") & vbCrLf
    Next i
    MsgBox msg, vbInformation, "Confluence Watcher"
    Exit Sub
EH:
    MsgBox "Child page read failed." & vbCrLf & vbCrLf & _
           "Error number: " & CStr(Err.Number) & vbCrLf & _
           "Source: " & Err.Source & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Confluence Watcher"
End Sub

'==================================================================
' Tree traversal
'==================================================================
Private Sub ScanTree(ByVal baseUrl As String, ByVal rootId As String, ByVal authMode As String, _
                     ByVal userName As String, ByVal secret As String, ByVal timeoutSec As Long, _
                     ByVal pageSize As Long, ByRef pages As Object)
    Dim stack As Collection
    Dim visited As Object
    Dim root As Object
    Dim parentId As String
    Dim children As Collection
    Dim i As Long
    Dim child As Object
    Dim childId As String

    Set stack = New Collection
    Set visited = CreateObject("Scripting.Dictionary")

    Set root = GetPage(baseUrl, rootId, "", authMode, userName, secret, timeoutSec)
    If Not pages.Exists(rootId) Then
        pages.Add rootId, root
    End If
    stack.Add rootId

    Do While stack.Count > 0
        parentId = CStr(stack.Item(stack.Count))
        stack.Remove stack.Count

        If Not visited.Exists(parentId) Then
            visited.Add parentId, True
            Set children = GetChildren(baseUrl, parentId, authMode, userName, secret, timeoutSec, pageSize)
            For i = 1 To children.Count
                Set child = children.Item(i)
                childId = CStr(child("id"))
                If Len(childId) > 0 Then
                    If Not pages.Exists(childId) Then
                        pages.Add childId, child
                        stack.Add childId
                    End If
                End If
            Next i
        End If

        Application.StatusBar = "Scanning... " & CStr(pages.Count) & " pages found"
        DoEvents
    Loop
End Sub

Private Function GetPage(ByVal baseUrl As String, ByVal pageId As String, ByVal parentId As String, _
                         ByVal authMode As String, ByVal userName As String, ByVal secret As String, _
                         ByVal timeoutSec As Long) As Object
    Dim body As String
    On Error GoTo EH
    body = HttpGet(TrimSlash(baseUrl) & "/rest/api/content/" & pageId & "?status=current&expand=version", _
                   authMode, userName, secret, timeoutSec)
    Set GetPage = PageFromJson(body, baseUrl, parentId)
    Exit Function
EH:
    Err.Raise Err.Number, "GetPage(pageId=" & pageId & ")", Err.Description
End Function

Private Function GetChildren(ByVal baseUrl As String, ByVal parentId As String, ByVal authMode As String, _
                             ByVal userName As String, ByVal secret As String, ByVal timeoutSec As Long, _
                             ByVal pageSize As Long) As Collection
    Dim all As Collection
    Dim startAt As Long
    Dim body As String
    Dim arr As String
    Dim objs As Collection
    Dim i As Long

    On Error GoTo EH
    Set all = New Collection
    startAt = 0
    If pageSize < 1 Then
        pageSize = 25
    End If

    Do
        body = HttpGet(TrimSlash(baseUrl) & "/rest/api/content/" & parentId & "/child/page?start=" & _
                       CStr(startAt) & "&limit=" & CStr(pageSize) & "&expand=version", _
                       authMode, userName, secret, timeoutSec)
        arr = JsonArrayForKey(body, "results")
        Set objs = SplitTopLevelObjects(arr)

        For i = 1 To objs.Count
            all.Add PageFromJson(CStr(objs.Item(i)), baseUrl, parentId)
        Next i

        If objs.Count < pageSize Then
            Exit Do
        End If
        startAt = startAt + objs.Count
    Loop

    Set GetChildren = all
    Exit Function
EH:
    Err.Raise Err.Number, "GetChildren(parentId=" & parentId & ", start=" & CStr(startAt) & ")", Err.Description
End Function

Private Function PageFromJson(ByVal js As String, ByVal baseUrl As String, ByVal parentId As String) As Object
    Dim d As Object
    Dim ver As String
    Dim byObj As String
    Dim links As String
    Dim webui As String
    Dim updater As String

    On Error GoTo EH
    Set d = CreateObject("Scripting.Dictionary")
    ver = JsonObjectForKey(js, "version")
    byObj = JsonObjectForKey(ver, "by")
    links = JsonObjectForKey(js, "_links")
    webui = JsonString(links, "webui")

    d("id") = JsonString(js, "id")
    d("title") = JsonString(js, "title")
    d("version") = JsonNumber(ver, "number")
    d("updatedAt") = JsonString(ver, "when")

    updater = JsonString(byObj, "displayName")
    If Len(updater) = 0 Then
        updater = JsonString(byObj, "username")
    End If
    d("updatedBy") = updater
    d("parentId") = parentId

    If Len(webui) = 0 Then
        d("url") = TrimSlash(baseUrl) & "/pages/viewpage.action?pageId=" & d("id")
    ElseIf Left$(webui, 1) = "/" Then
        d("url") = TrimSlash(baseUrl) & webui
    Else
        d("url") = webui
    End If

    Set PageFromJson = d
    Exit Function
EH:
    Err.Raise Err.Number, "PageFromJson", Err.Description & " | JSON head: " & Left$(js, 200)
End Function

'==================================================================
' HTTP
'==================================================================
Private Function HttpGet(ByVal url As String, ByVal authMode As String, ByVal userName As String, _
                         ByVal secret As String, ByVal timeoutSec As Long) As String
    Dim h As Object
    Dim responsePreview As String

    If timeoutSec < 5 Then
        timeoutSec = 30
    End If

    Set h = GetHttp()
    h.SetTimeouts timeoutSec * 1000, timeoutSec * 1000, timeoutSec * 1000, timeoutSec * 1000
    h.Open "GET", url, False
    h.SetRequestHeader "Accept", "application/json"

    Select Case LCase$(authMode)
        Case "basic"
            h.SetRequestHeader "Authorization", "Basic " & Base64Encode(userName & ":" & secret)
        Case "cookie"
            h.SetRequestHeader "Cookie", secret
        Case "session"
            ' The shared instance replays the session cookies automatically.
            ' Setting the header manually would override that jar.
        Case Else
            h.SetRequestHeader "Authorization", "Bearer " & secret
    End Select

    On Error Resume Next
    h.SetProxy 1
    On Error GoTo 0

    h.Send

    If h.Status < 200 Or h.Status >= 300 Then
        responsePreview = Left$(Utf8BytesToText(h.ResponseBody), 400)
        Err.Raise vbObjectError + 101, "HttpGet", _
                  "HTTP " & CStr(h.Status) & " " & h.StatusText & " | " & url & " | Response: " & responsePreview
    End If

    HttpGet = Utf8BytesToText(h.ResponseBody)
End Function

Private Function Utf8BytesToText(ByVal bodyBytes As Variant) As String
    Dim stream As Object

    If Not IsArray(bodyBytes) Then
        Utf8BytesToText = CStr(bodyBytes)
        Exit Function
    End If

    On Error GoTo Fallback
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bodyBytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    Utf8BytesToText = stream.ReadText
    stream.Close
    Exit Function

Fallback:
    On Error Resume Next
    Utf8BytesToText = StrConv(bodyBytes, vbUnicode)
End Function

Private Function Base64Encode(ByVal plain As String) As String
    Dim bytes() As Byte
    Dim dom As Object
    Dim node As Object
    bytes = StrConv(plain, vbFromUnicode)
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    Base64Encode = Replace(Replace(node.Text, vbCr, ""), vbLf, "")
End Function

'==================================================================
' Worksheet state
'==================================================================
Private Function LoadCurrentState() As Object
    Dim d As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim pageId As String
    Dim p As Object

    Set d = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(STATE_SHEET_INDEX)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 4 To lastRow
        pageId = CellText(ws.Cells(r, 1))
        If Len(pageId) > 0 Then
            If Not d.Exists(pageId) Then
                Set p = CreateObject("Scripting.Dictionary")
                p("id") = pageId
                p("title") = CellText(ws.Cells(r, 2))
                p("version") = SafeLong(ws.Cells(r, 3).Value)
                p("updatedAt") = CellText(ws.Cells(r, 4))
                p("updatedBy") = CellText(ws.Cells(r, 5))
                p("parentId") = CellText(ws.Cells(r, 6))
                p("url") = CellText(ws.Cells(r, 7))
                d.Add pageId, p
            End If
        End If
    Next r

    Set LoadCurrentState = d
End Function

Private Function WriteChanges(ByVal oldState As Object, ByVal newState As Object, _
                              ByVal suppressAdds As Boolean) As Long
    Dim count As Long
    Dim id As Variant
    Dim p As Object
    Dim oldP As Object

    For Each id In newState.Keys
        Set p = newState(id)
        If Not oldState.Exists(id) Then
            If Not suppressAdds Then
                AppendChange "Added", p, "", p("version"), ""
                count = count + 1
            End If
        Else
            Set oldP = oldState(id)
            If SafeLong(oldP("version")) <> SafeLong(p("version")) Then
                AppendChange "Updated", p, oldP("version"), p("version"), "Version changed"
                count = count + 1
            ElseIf CStr(oldP("title")) <> CStr(p("title")) Then
                AppendChange "Renamed", p, oldP("version"), p("version"), "Previous title: " & oldP("title")
                count = count + 1
            ElseIf CStr(oldP("parentId")) <> CStr(p("parentId")) Then
                AppendChange "Moved", p, oldP("version"), p("version"), "Previous parent: " & oldP("parentId")
                count = count + 1
            End If
        End If
    Next id

    For Each id In oldState.Keys
        If Not newState.Exists(id) Then
            Set oldP = oldState(id)
            AppendChange "Removed or inaccessible", oldP, oldP("version"), "", _
                         "The page is no longer present in the visible page tree."
            count = count + 1
        End If
    Next id

    WriteChanges = count
End Function

Private Sub RewriteCurrentState(ByVal pages As Object)
    Dim ws As Worksheet
    Dim r As Long
    Dim id As Variant
    Dim p As Object
    Dim scanTime As Date

    Set ws = ThisWorkbook.Worksheets(STATE_SHEET_INDEX)
    ws.Rows("4:" & ws.Rows.Count).ClearContents
    ws.Columns(1).NumberFormat = "@"
    ws.Columns(6).NumberFormat = "@"
    r = 4
    scanTime = Now

    For Each id In pages.Keys
        Set p = pages(id)
        ws.Cells(r, 1).Resize(1, 8).Value = Array(p("id"), p("title"), p("version"), p("updatedAt"), _
                                                  p("updatedBy"), p("parentId"), p("url"), scanTime)
        r = r + 1
    Next id

    FormatData ws, 8, r - 1
End Sub

Private Sub AppendHistory(ByVal pages As Object)
    Dim ws As Worksheet
    Dim r As Long
    Dim id As Variant
    Dim p As Object
    Dim scanTime As Date

    Set ws = ThisWorkbook.Worksheets(HIST_SHEET_INDEX)
    ws.Columns(2).NumberFormat = "@"
    ws.Columns(7).NumberFormat = "@"
    r = NextRow(ws)
    scanTime = Now

    For Each id In pages.Keys
        Set p = pages(id)
        ws.Cells(r, 1).Resize(1, 8).Value = Array(scanTime, p("id"), p("title"), p("version"), _
                                                  p("updatedAt"), p("updatedBy"), p("parentId"), p("url"))
        r = r + 1
    Next id

    FormatData ws, 8, r - 1
End Sub

Private Sub AppendChange(ByVal kind As String, ByVal p As Object, ByVal oldVer As Variant, _
                         ByVal newVer As Variant, ByVal note As String)
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(CHANGE_SHEET_INDEX)
    ws.Columns(3).NumberFormat = "@"
    ws.Columns(9).NumberFormat = "@"
    r = NextRow(ws)
    ws.Cells(r, 1).Resize(1, 11).Value = Array(Now, kind, p("id"), p("title"), oldVer, newVer, _
                                               p("updatedAt"), p("updatedBy"), p("parentId"), p("url"), note)
    FormatData ws, 11, r
End Sub

Private Sub AddRunLog(ByVal level As String, ByVal message As String, _
                      ByVal pageCount As Long, ByVal changeCount As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_INDEX)
    r = NextRow(ws)
    ws.Cells(r, 1).Resize(1, 5).Value = Array(Now, level, message, pageCount, changeCount)
End Sub

Private Sub FormatData(ByVal ws As Worksheet, ByVal cols As Long, ByVal lastRow As Long)
    If lastRow < 4 Then
        Exit Sub
    End If
    ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, cols)).VerticalAlignment = xlTop
    ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, cols)).WrapText = False
    ws.Columns(1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Private Function NextRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 4 Then
        r = 4
    End If
    NextRow = r
End Function

Private Function Cfg(ByVal r As Long, ByVal c As Long) As Variant
    Cfg = ThisWorkbook.Worksheets(CFG_SHEET_INDEX).Cells(r, c).Value
End Function

Private Function GetSecret(ByVal authMode As String, ByVal fallback As String) As String
    Dim v As String
    Select Case LCase$(authMode)
        Case "basic", "session"
            v = Environ$("CONFLUENCE_PASSWORD")
        Case "cookie"
            v = Environ$("CONFLUENCE_COOKIE")
        Case Else
            v = Environ$("CONFLUENCE_TOKEN")
    End Select
    If Len(v) = 0 Then
        v = Trim$(fallback)
    End If
    GetSecret = v
End Function

'==================================================================
' Safe conversions
' Page IDs such as 4480574806 exceed the Long limit of 2147483647,
' so CLng raises overflow (error 6). These helpers never overflow and
' never lose digits to scientific notation.
'==================================================================
Private Function SafeLong(ByVal v As Variant) As Long
    Dim d As Double
    On Error GoTo Bail

    If IsEmpty(v) Or IsNull(v) Then
        SafeLong = 0
        Exit Function
    End If
    If IsError(v) Then
        SafeLong = 0
        Exit Function
    End If

    d = Val(Trim$(CStr(v)))

    If d > 2147483647# Then
        SafeLong = 2147483647
    ElseIf d < -2147483648# Then
        SafeLong = -2147483648#
    Else
        SafeLong = CLng(Int(d))
    End If
    Exit Function
Bail:
    SafeLong = 0
End Function

' Reads a cell as text without turning long IDs into 4.48057E+09.
Private Function CellText(ByVal c As Range) As String
    Dim v As Variant
    On Error GoTo Bail

    v = c.Value
    If IsEmpty(v) Or IsNull(v) Then
        CellText = ""
        Exit Function
    End If
    If IsError(v) Then
        CellText = ""
        Exit Function
    End If

    If VarType(v) = vbDouble Or VarType(v) = vbSingle Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        If v = Int(v) And Abs(v) < 1E+15 Then
            CellText = Format$(v, "0")
        Else
            CellText = Trim$(CStr(v))
        End If
    Else
        CellText = Trim$(CStr(v))
    End If
    Exit Function
Bail:
    CellText = ""
End Function

Private Function TrimSlash(ByVal s As String) As String
    Do While Right$(s, 1) = "/"
        s = Left$(s, Len(s) - 1)
    Loop
    TrimSlash = s
End Function

'==================================================================
' Minimal JSON reader
'==================================================================
Private Function JsonString(ByVal js As String, ByVal key As String) As String
    Dim p As Long
    Dim out As String
    Dim ch As String
    Dim esc As Boolean
    Dim hexCode As String

    p = KeyValueStart(js, key)
    If p = 0 Then
        Exit Function
    End If
    If Mid$(js, p, 1) <> Chr$(34) Then
        Exit Function
    End If
    p = p + 1

    Do While p <= Len(js)
        ch = Mid$(js, p, 1)
        If esc Then
            Select Case ch
                Case Chr$(34)
                    out = out & Chr$(34)
                Case "\"
                    out = out & "\"
                Case "/"
                    out = out & "/"
                Case "b"
                    out = out & Chr$(8)
                Case "f"
                    out = out & Chr$(12)
                Case "n"
                    out = out & vbLf
                Case "r"
                    out = out & vbCr
                Case "t"
                    out = out & vbTab
                Case "u"
                    hexCode = Mid$(js, p + 1, 4)
                    If Len(hexCode) = 4 Then
                        out = out & UnicodeChar(HexToLong(hexCode))
                        p = p + 4
                    End If
                Case Else
                    out = out & ch
            End Select
            esc = False
        ElseIf ch = "\" Then
            esc = True
        ElseIf ch = Chr$(34) Then
            Exit Do
        Else
            out = out & ch
        End If
        p = p + 1
    Loop

    JsonString = out
End Function

Private Function HexToLong(ByVal hexCode As String) As Long
    Dim i As Long
    Dim ch As String
    Dim digit As Long
    Dim total As Long

    For i = 1 To Len(hexCode)
        ch = LCase$(Mid$(hexCode, i, 1))
        digit = InStr("0123456789abcdef", ch) - 1
        If digit < 0 Then
            HexToLong = 0
            Exit Function
        End If
        total = total * 16 + digit
    Next i

    HexToLong = total
End Function

' ChrW rejects arguments above 32767, which covers most Japanese characters.
Private Function UnicodeChar(ByVal codePoint As Long) As String
    If codePoint < 0 Then
        UnicodeChar = ""
    ElseIf codePoint > 32767 Then
        UnicodeChar = ChrW$(codePoint - 65536)
    Else
        UnicodeChar = ChrW$(codePoint)
    End If
End Function

Private Function JsonNumber(ByVal js As String, ByVal key As String) As Long
    Dim p As Long
    Dim s As String
    Dim ch As String

    p = KeyValueStart(js, key)
    If p = 0 Then
        Exit Function
    End If

    Do While p <= Len(js)
        ch = Mid$(js, p, 1)
        If InStr("0123456789-", ch) = 0 Then
            Exit Do
        End If
        s = s & ch
        p = p + 1
    Loop

    If Len(s) = 0 Then
        JsonNumber = 0
    Else
        JsonNumber = SafeLong(s)
    End If
End Function

Private Function KeyValueStart(ByVal js As String, ByVal key As String) As Long
    Dim p As Long

    If Len(js) = 0 Then
        Exit Function
    End If

    p = InStr(1, js, Chr$(34) & key & Chr$(34), vbBinaryCompare)
    If p = 0 Then
        Exit Function
    End If

    p = InStr(p + Len(key) + 2, js, ":")
    If p = 0 Then
        Exit Function
    End If
    p = p + 1

    Do While p <= Len(js)
        If InStr(" " & vbCr & vbLf & vbTab, Mid$(js, p, 1)) = 0 Then
            Exit Do
        End If
        p = p + 1
    Loop

    KeyValueStart = p
End Function

Private Function JsonObjectForKey(ByVal js As String, ByVal key As String) As String
    Dim p As Long
    p = KeyValueStart(js, key)
    If p = 0 Then
        Exit Function
    End If
    If Mid$(js, p, 1) <> "{" Then
        Exit Function
    End If
    JsonObjectForKey = BalancedPart(js, p, "{", "}")
End Function

Private Function JsonArrayForKey(ByVal js As String, ByVal key As String) As String
    Dim p As Long
    p = KeyValueStart(js, key)
    If p = 0 Then
        Exit Function
    End If
    If Mid$(js, p, 1) <> "[" Then
        Exit Function
    End If
    JsonArrayForKey = BalancedPart(js, p, "[", "]")
End Function

Private Function BalancedPart(ByVal js As String, ByVal startPos As Long, _
                              ByVal openCh As String, ByVal closeCh As String) As String
    Dim depth As Long
    Dim i As Long
    Dim inString As Boolean
    Dim esc As Boolean
    Dim ch As String

    For i = startPos To Len(js)
        ch = Mid$(js, i, 1)
        If inString Then
            If esc Then
                esc = False
            ElseIf ch = "\" Then
                esc = True
            ElseIf ch = Chr$(34) Then
                inString = False
            End If
        Else
            If ch = Chr$(34) Then
                inString = True
            ElseIf ch = openCh Then
                depth = depth + 1
            ElseIf ch = closeCh Then
                depth = depth - 1
                If depth = 0 Then
                    BalancedPart = Mid$(js, startPos, i - startPos + 1)
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function SplitTopLevelObjects(ByVal arr As String) As Collection
    Dim col As Collection
    Dim i As Long
    Dim part As String

    Set col = New Collection
    If Len(arr) < 2 Then
        Set SplitTopLevelObjects = col
        Exit Function
    End If

    i = 2
    Do While i < Len(arr)
        If Mid$(arr, i, 1) = "{" Then
            part = BalancedPart(arr, i, "{", "}")
            If Len(part) = 0 Then
                Exit Do
            End If
            col.Add part
            i = i + Len(part)
        Else
            i = i + 1
        End If
    Loop

    Set SplitTopLevelObjects = col
End Function
