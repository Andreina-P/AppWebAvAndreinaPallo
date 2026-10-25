<%
'=====================================================================
' PoliMusic · ASP Clásico (VBScript)
' inc_db.asp - Conexión a SQL Server, funciones auxiliares y layout
'---------------------------------------------------------------------
' Este archivo se incluye en todas las páginas .asp del sitio.
' No crear tablas adicionales: solo se usan las tablas del script
' BDD_PoliMusic.sql.
'=====================================================================
Option Explicit

Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = True

'--- Constantes ADO ---
Const adCmdText    = 1
Const adParamInput = 1
Const adInteger    = 3
Const adVarChar    = 200
Const adDate       = 7
Const adBoolean    = 11

'--- Cadena de conexión: SQL Server local, base BDD_PoliMusic ---
' Si cambia el servidor o las credenciales, modificar aquí.
Const CONN_STRING = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=BDD_PoliMusic;User Id=usr_polimusic;Password=Politecnica1;"

'---------------------------------------------------------------------
' Acceso a la base de datos (consultas parametrizadas anti SQL-Injection)
'---------------------------------------------------------------------
Function OpenConn()
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CONN_STRING
    Set OpenConn = conn
End Function

' Agrega parámetros posicionales (?) al comando.
' types/values son arreglos del mismo tamaño. Tipos: "s" texto, "i" entero,
' "d" fecha, "b" booleano.
Sub AddParams(cmd, types, values)
    Dim i, t, typ, sz
    If Not IsArray(values) Then Exit Sub
    For i = 0 To UBound(values)
        t = LCase(CStr(types(i)))
        Select Case t
            Case "i": typ = adInteger : sz = 4
            Case "d": typ = adDate : sz = 8
            Case "b": typ = adBoolean : sz = 2
            Case Else: typ = adVarChar : sz = 255
        End Select
        cmd.Parameters.Append cmd.CreateParameter("?", typ, adParamInput, sz, values(i))
    Next
End Sub

Function MakeCmd(conn, sql, types, values)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText
    AddParams cmd, types, values
    Set MakeCmd = cmd
End Function

' Devuelve un Recordset (la página llamadora debe cerrarlo).
Function QueryRS(conn, sql, types, values)
    Set QueryRS = MakeCmd(conn, sql, types, values).Execute
End Function

' Devuelve el primer campo de la primera fila (Null si no hay filas).
Function Scalar(conn, sql, types, values)
    Dim rs
    Set rs = QueryRS(conn, sql, types, values)
    If Not rs.EOF Then
        Scalar = rs.Fields(0).Value
    Else
        Scalar = Null
    End If
    rs.Close
End Function

' INSERT / UPDATE / DELETE -> devuelve la cantidad de filas afectadas.
Function NonQueryP(conn, sql, types, values)
    Dim ra
    MakeCmd(conn, sql, types, values).Execute ra
    NonQueryP = ra
End Function

' INSERT ...; SELECT SCOPE_IDENTITY() -> devuelve el ID generado.
Function InsertId(conn, sql, types, values)
    Dim rs
    Set rs = MakeCmd(conn, sql, types, values).Execute
    If Not rs.EOF Then
        InsertId = CLng(rs.Fields(0).Value)
    Else
        InsertId = 0
    End If
    rs.Close
End Function

'---------------------------------------------------------------------
' Sesión y roles (0 = usuario normal, 1 = administrador)
'---------------------------------------------------------------------
Function IsLoggedIn()
    IsLoggedIn = (Not IsEmpty(Session("ID_USER")))
End Function

Function IsAdmin()
    IsAdmin = (IsLoggedIn() And Session("USER_TYPE") = 1)
End Function

Sub RequireLogin()
    If Not IsLoggedIn() Then
        Response.Redirect "login.asp"
    End If
End Sub

Sub RequireAdmin()
    If Not IsLoggedIn() Then
        Response.Redirect "login.asp"
    ElseIf Not IsAdmin() Then
        Response.Redirect "songs.asp?msg=acceso_denegado"
    End If
End Sub

'---------------------------------------------------------------------
' Utilidades de presentación
'---------------------------------------------------------------------
Function Esc(s)
    If IsNull(s) Then
        Esc = ""
    Else
        Esc = Server.HTMLEncode(CStr(s))
    End If
End Function

' IIf no existe en VBScript (es de VB/VBA): se define aqui.
Function IIf(cond, a, b)
    If cond Then
        IIf = a
    Else
        IIf = b
    End If
End Function

Function Nz(v, dflt)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = dflt
    Else
        Nz = v
    End If
End Function

' Escapa una cadena para incrustarla de forma segura en JavaScript.
Function JsStr(s)
    If IsNull(s) Then JsStr = "" : Exit Function
    s = CStr(s)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, "'", "\'")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, "</", "<\/")
    JsStr = s
End Function

' Revierte el escape de entidades que produce FOR XML PATH.
Function XmlUnescape(s)
    If IsNull(s) Then XmlUnescape = "" : Exit Function
    s = CStr(s)
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&apos;", "'")
    s = Replace(s, "&amp;", "&")
    XmlUnescape = s
End Function

'---------------------------------------------------------------------
' Resolución de recursos contra la carpeta SongFilesAndImages_resources
'---------------------------------------------------------------------
Function ResBase()
    ResBase = Server.MapPath("SongFilesAndImages_resources")
End Function

Function FileNameOnly(p)
    p = Replace(CStr(p), "\", "/")
    If InStrRev(p, "/") > 0 Then p = Mid(p, InStrRev(p, "/") + 1)
    FileNameOnly = p
End Function

Function FolderOf(p)
    Dim s
    s = Replace(CStr(p), "\", "/")
    If InStrRev(s, "/") > 0 Then
        FolderOf = LCase(Left(s, InStrRev(s, "/") - 1))
    Else
        FolderOf = ""
    End If
End Function

' Convierte una ruta de la BD (ej. ~/songFiles/x.mp3) en la URL real del
' archivo dentro de SongFilesAndImages_resources. Devuelve "" si no existe.
Function FindResource(dbPath)
    Dim fs, base, fileName, pref, fld, url
    FindResource = ""
    If IsNull(dbPath) Then Exit Function
    dbPath = Trim(CStr(dbPath))
    If dbPath = "" Then Exit Function
    base = ResBase()
    fileName = FileNameOnly(dbPath)
    pref = FolderOf(dbPath)
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(base) Then Exit Function
    ' 1) Coincidencia en la subcarpeta indicada por la ruta (sin distinguir mayúsculas)
    For Each fld In fs.GetFolder(base).SubFolders
        If LCase(fld.Name) = pref Then
            If fs.FileExists(fld.Path & "\" & fileName) Then
                url = "SongFilesAndImages_resources/" & fld.Name & "/" & fs.GetFile(fld.Path & "\" & fileName).Name
                FindResource = url
                Exit Function
            End If
        End If
    Next
    ' 2) Cualquier subcarpeta que contenga un archivo con ese nombre
    For Each fld In fs.GetFolder(base).SubFolders
        If fs.FileExists(fld.Path & "\" & fileName) Then
            url = "SongFilesAndImages_resources/" & fld.Name & "/" & fs.GetFile(fld.Path & "\" & fileName).Name
            FindResource = url
            Exit Function
        End If
    Next
End Function

' Imagen por defecto cuando la ruta de la BD no existe en los recursos.
Function FallbackUrl(kind)
    Select Case LCase(CStr(kind))
        Case "user": FallbackUrl = "SongFilesAndImages_resources/UserImages/default.jpg"
        Case "artist": FallbackUrl = "SongFilesAndImages_resources/UserImages/default.jpg"
        Case Else: FallbackUrl = "SongFilesAndImages_resources/PlaylistImages/default.jpg"
    End Select
End Function

' URL de imagen con reserva (kind: user | artist | album).
Function ImgUrl(dbPath, kind)
    Dim u
    u = FindResource(dbPath)
    If u = "" Then u = FallbackUrl(kind)
    ImgUrl = u
End Function

' URL del mp3; "" si el archivo no está en los recursos.
Function SongUrl(dbPath)
    SongUrl = FindResource(dbPath)
End Function

'---------------------------------------------------------------------
' Layout compartido
'---------------------------------------------------------------------
Function NavClass(target, active)
    If target = active Then
        NavClass = "active"
    Else
        NavClass = ""
    End If
End Function

Sub Topbar(active)
    Dim u, photo
    u = Nz(Session("USERNAME"), "")
    photo = ImgUrl(Session("USER_PHOTO"), "user")
%>
<header class="topbar">
    <a class="brand" href="songs.asp">
        <img class="brand-mark" src="assets/img/logo.svg" alt="">
        <span class="brand-name">Poli<strong>Music</strong></span>
    </a>
    <% If IsAdmin() Then %>
    <nav class="admin-nav" aria-label="Administración del catálogo">
        <a class="<%=NavClass("admin_songs.asp", active)%>" href="admin_songs.asp">Canciones</a>
        <a class="<%=NavClass("admin_artists.asp", active)%>" href="admin_artists.asp">Artistas</a>
        <a class="<%=NavClass("admin_albums.asp", active)%>" href="admin_albums.asp">Álbumes</a>
    </nav>
    <% End If %>
    <div class="topbar-right">
        <span class="user-chip">
            <img src="<%=Esc(photo)%>" alt="">
            <span class="user-meta">
                <strong><%=Esc(u)%></strong>
                <em><% If IsAdmin() Then %>Administrador<% Else %>Usuario<% End If %></em>
            </span>
        </span>
        <a class="btn btn-ghost btn-sm" href="logout.asp">Salir</a>
    </div>
</header>
<%
End Sub

Sub PageTop(title, active)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><%=Esc(title)%> · PoliMusic</title>
<link rel="icon" type="image/svg+xml" href="assets/img/logo.svg">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="assets/css/styles.css">
</head>
<body data-page="<%=Esc(active)%>">
<%
    Topbar active
End Sub

Sub PageBottom()
%>
<script src="assets/js/player.js"></script>
</body>
</html>
<%
End Sub

' Mensaje flash proveniente de un redirect (?msg=...)
Sub FlashIfAny()
    Dim m, txt
    m = Request.QueryString("msg")
    Select Case m
        Case "guardado": txt = "Los cambios se guardaron correctamente."
        Case "borrado": txt = "El registro se eliminó correctamente."
        Case "acceso_denegado": txt = "Acceso denegado: esta sección es solo para administradores."
        Case "sesion": txt = "Sesión finalizada. Vuelve a iniciar sesión para continuar."
        Case Else: txt = ""
    End Select
    If txt <> "" Then
%>
<div class="flash" role="status"><%=txt%></div>
<%
    End If
End Sub
%>
