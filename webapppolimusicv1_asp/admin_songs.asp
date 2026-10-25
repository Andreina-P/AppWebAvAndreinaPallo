<%@Language="VBScript" CodePage="65001"%>
<!--#include file="inc_db.asp" -->
<%
'=====================================================================
' admin_songs.asp · CRUD de canciones (solo rol administrador)
' Tablas usadas: TBL_SONG, TBL_PLAYER (artistas de cada canción)
'=====================================================================
RequireAdmin

Dim conn, rs, sql
Dim editId, formName, formPath, formPlays
Dim errMsg, msgForm, action, id
Dim fName, fPath, fPlays, songId
Dim parts, j, assigned

editId = 0
formName = ""
formPath = ""
formPlays = "0"
errMsg = ""
msgForm = "Nueva canción"
songId = 0

If Request.QueryString("edit") <> "" And IsNumeric(Request.QueryString("edit")) Then
    editId = CLng(Request.QueryString("edit"))
End If

Set conn = OpenConn()

'----------------- POST: guardar o eliminar -----------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    action = Request.Form("action")
    id = Nz(Request.Form("id"), 0)
    If Not IsNumeric(id) Then id = 0

    If action = "delete" Then
        If CLng(id) > 0 Then
            NonQueryP conn, "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array("i"), Array(CLng(id))
            NonQueryP conn, "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array("i"), Array(CLng(id))
            NonQueryP conn, "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?", Array("i"), Array(CLng(id))
            NonQueryP conn, "DELETE FROM TBL_SONG WHERE ID_SONG = ?", Array("i"), Array(CLng(id))
        End If
        conn.Close
        Set conn = Nothing
        Response.Redirect "admin_songs.asp?msg=borrado"
    Else
        editId = CLng(id)
        fName = Trim(Request.Form("song_name"))
        fPath = Trim(Request.Form("song_path"))
        fPlays = Request.Form("plays")
        If Not IsNumeric(fPlays) Then fPlays = 0
        formName = fName
        formPath = fPath
        formPlays = CStr(fPlays)
        If fName = "" Then errMsg = errMsg & "El nombre de la canción es obligatorio.<br>"
        If fPath = "" Then errMsg = errMsg & "La ruta del archivo MP3 es obligatoria.<br>"

        If errMsg = "" Then
            If editId > 0 Then
                NonQueryP conn, "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?", Array("s", "s", "i", "i"), Array(fName, fPath, CLng(fPlays), editId)
                songId = editId
            Else
                songId = InsertId(conn, "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, ?); SELECT CAST(SCOPE_IDENTITY() AS int)", Array("s", "s", "i"), Array(fName, fPath, CLng(fPlays)))
            End If
            ' Sincroniza los artistas que interpretan la canción (TBL_PLAYER)
            NonQueryP conn, "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array("i"), Array(songId)
            If Request.Form("artist_ids") <> "" Then
                parts = Split(Request.Form("artist_ids"), ",")
                For j = 0 To UBound(parts)
                    If IsNumeric(Trim(parts(j))) And CLng(Trim(parts(j))) > 0 Then
                        NonQueryP conn, "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)", Array("i", "i"), Array(songId, CLng(Trim(parts(j))))
                    End If
                Next
            End If
            conn.Close
            Set conn = Nothing
            Response.Redirect "admin_songs.asp?msg=guardado"
        End If
        msgForm = "Corrige los datos del formulario"
    End If
End If

'----------------- Cargar canción a editar -----------------
If editId > 0 And errMsg = "" Then
    Set rs = QueryRS(conn, "SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG WHERE ID_SONG = ?", Array("i"), Array(editId))
    If Not rs.EOF Then
        formName = rs("SONG_NAME")
        formPath = rs("SONG_PATH")
        formPlays = CStr(Nz(rs("PLAYS"), 0))
        msgForm = "Editar canción #" & editId
    Else
        editId = 0
    End If
    rs.Close
End If

' Artistas actualmente asignados
Set assigned = Server.CreateObject("Scripting.Dictionary")
If errMsg <> "" And Request.Form("artist_ids") <> "" Then
    parts = Split(Request.Form("artist_ids"), ",")
    For j = 0 To UBound(parts)
        If IsNumeric(Trim(parts(j))) And Not assigned.Exists(CLng(Trim(parts(j)))) Then
            assigned.Add CLng(Trim(parts(j))), True
        End If
    Next
ElseIf editId > 0 Then
    Set rs = QueryRS(conn, "SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = ?", Array("i"), Array(editId))
    Do While Not rs.EOF
        If Not assigned.Exists(CLng(rs("ID_ARTIST").Value)) Then
            assigned.Add CLng(rs("ID_ARTIST").Value), True
        End If
        rs.MoveNext
    Loop
    rs.Close
End If
%>
<% PageTop "Administrar canciones", "admin_songs.asp" %>

<main class="container">
    <% FlashIfAny %>

    <div class="admin-grid">
        <section class="card panel">
            <h2><%=msgForm%></h2>
            <p class="muted">Los campos marcados son obligatorios. La ruta apunta al mp3 dentro de <b>SongFilesAndImages_resources/SongFiles</b>.</p>
            <% If errMsg <> "" Then %><div class="form-error" role="alert"><%=errMsg%></div><% End If %>
            <form method="post" action="admin_songs.asp" class="form-grid">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="id" value="<%=editId%>">
                <div class="field">
                    <label for="song_name">Nombre de la canción *</label>
                    <input id="song_name" name="song_name" type="text" value="<%=Esc(formName)%>" maxlength="50" required>
                </div>
                <div class="field">
                    <label for="song_path">Ruta del MP3 *</label>
                    <input id="song_path" name="song_path" type="text" value="<%=Esc(formPath)%>" maxlength="255" placeholder="~/songFiles/bensound-adventure.mp3" required>
                </div>
                <div class="field">
                    <label for="plays">Reproducciones</label>
                    <input id="plays" name="plays" type="number" min="0" value="<%=Esc(formPlays)%>">
                </div>
                <div class="field">
                    <label>Artistas que la interpretan</label>
                    <div class="check-pills">
<%
Set rs = QueryRS(conn, "SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME", Array(), Array())
Do While Not rs.EOF
    Dim artLabel, artChecked
    artLabel = Trim(Nz(rs("ARTIST_NAME"), "") & " " & Nz(rs("ARTIST_LASTNAME"), ""))
    artChecked = assigned.Exists(CLng(rs("ID_ARTIST").Value))
%>
                        <label class="check-pill">
                            <input type="checkbox" name="artist_ids" value="<%=rs("ID_ARTIST")%>"<%=IIf(artChecked, " checked", "")%>>
                            <span><%=Esc(artLabel)%></span>
                        </label>
<%
    rs.MoveNext
Loop
rs.Close
%>
                    </div>
                </div>
                <div class="row-actions">
                    <a class="btn btn-ghost" href="admin_songs.asp">Limpiar</a>
                    <button type="submit" class="btn btn-primary">Guardar canción</button>
                </div>
            </form>
        </section>

        <section class="card panel">
            <h2>Canciones registradas</h2>
            <p class="muted">Edita o elimina canciones. Al eliminar se quitan también sus relaciones con artistas, álbumes y listas.</p>
            <div class="table-wrap">
                <table class="table admin-table">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Canción</th>
                            <th>Ruta</th>
                            <th class="num">Plays</th>
                            <th>Artistas</th>
                            <th>Audio</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody>
<%
sql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, s.PLAYS, " & _
      "STUFF((SELECT ', ' + a.ARTIST_NAME FROM TBL_PLAYER pl INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = pl.ID_ARTIST WHERE pl.ID_SONG = s.ID_SONG FOR XML PATH('')), 1, 2, '') AS ARTISTAS " & _
      "FROM TBL_SONG s ORDER BY s.ID_SONG"
Set rs = QueryRS(conn, sql, Array(), Array())
If rs.EOF Then
%>
                        <tr><td colspan="7" class="empty">No hay canciones registradas.</td></tr>
<%
End If
Do While Not rs.EOF
    Dim listArtists, hasAudio
    listArtists = XmlUnescape(rs("ARTISTAS"))
    If listArtists = "" Then listArtists = "—"
    hasAudio = (SongUrl(rs("SONG_PATH")) <> "")
%>
                        <tr>
                            <td><%=rs("ID_SONG")%></td>
                            <td><strong><%=Esc(rs("SONG_NAME"))%></strong></td>
                            <td class="wrap-cell" style="color:var(--muted);font-size:.8rem"><%=Esc(rs("SONG_PATH"))%></td>
                            <td class="num"><%=Nz(rs("PLAYS"), 0)%></td>
                            <td><%=Esc(listArtists)%></td>
                            <td>
                                <% If hasAudio Then %><span class="file-ok">disponible</span><% Else %><span class="file-missing">sin archivo</span><% End If %>
                            </td>
                            <td>
                                <div class="row-actions">
                                    <a class="btn btn-ghost btn-sm" href="admin_songs.asp?edit=<%=rs("ID_SONG")%>">Editar</a>
                                    <form method="post" action="admin_songs.asp" class="inline-form" onsubmit="return confirm('¿Eliminar la canción «<%=Esc(rs("SONG_NAME"))%>»? Esta acción no se puede deshacer.');">
                                        <input type="hidden" name="action" value="delete">
                                        <input type="hidden" name="id" value="<%=rs("ID_SONG")%>">
                                        <button type="submit" class="btn btn-danger btn-sm">Eliminar</button>
                                    </form>
                                </div>
                            </td>
                        </tr>
<%
    rs.MoveNext
Loop
rs.Close
%>
                    </tbody>
                </table>
            </div>
        </section>
    </div>
</main>

<% PageBottom %>
<%
conn.Close
Set conn = Nothing
%>
