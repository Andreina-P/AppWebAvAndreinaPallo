<%@Language="VBScript" CodePage="65001"%>
<!--#include file="inc_db.asp" -->
<%
'=====================================================================
' admin_albums.asp · CRUD de álbumes (solo rol administrador)
' Tablas usadas: TBL_ALBUM, TBL_TRACKLIST (canciones del álbum)
'=====================================================================
RequireAdmin

Dim conn, rs, sql
Dim editId, formTitle, formYear, formImage
Dim errMsg, msgForm, action, id
Dim fTitle, fYear, fImage, newId
Dim parts, j, selected

editId = 0
formTitle = ""
formYear = ""
formImage = ""
errMsg = ""
msgForm = "Nuevo álbum"
newId = 0

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
            NonQueryP conn, "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array("i"), Array(CLng(id))
            NonQueryP conn, "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array("i"), Array(CLng(id))
        End If
        conn.Close
        Set conn = Nothing
        Response.Redirect "admin_albums.asp?msg=borrado"
    Else
        editId = CLng(id)
        fTitle = Trim(Request.Form("album_title"))
        fYear = Trim(Request.Form("release_year"))
        fImage = Trim(Request.Form("album_image"))
        formTitle = fTitle
        formYear = fYear
        formImage = fImage
        If fTitle = "" Then errMsg = errMsg & "El título del álbum es obligatorio.<br>"
        If fYear <> "" And (Not IsNumeric(fYear) Or CLng(fYear) < 1500 Or CLng(fYear) > 2100) Then
            errMsg = errMsg & "El año de lanzamiento debe ser un número entre 1500 y 2100.<br>"
        End If

        If errMsg = "" Then
            If editId > 0 Then
                If fYear = "" Then
                    NonQueryP conn, "UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = NULL, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?", Array("s", "s", "i"), Array(fTitle, fImage, editId)
                Else
                    NonQueryP conn, "UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?", Array("s", "i", "s", "i"), Array(fTitle, CLng(fYear), fImage, editId)
                End If
                newId = editId
            Else
                If fYear = "" Then
                    newId = InsertId(conn, "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (?, NULL, ?); SELECT CAST(SCOPE_IDENTITY() AS int)", Array("s", "s"), Array(fTitle, fImage))
                Else
                    newId = InsertId(conn, "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (?, ?, ?); SELECT CAST(SCOPE_IDENTITY() AS int)", Array("s", "i", "s"), Array(fTitle, CLng(fYear), fImage))
                End If
            End If
            ' Sincroniza las canciones del álbum (TBL_TRACKLIST)
            NonQueryP conn, "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array("i"), Array(newId)
            If Request.Form("song_ids") <> "" Then
                parts = Split(Request.Form("song_ids"), ",")
                For j = 0 To UBound(parts)
                    If IsNumeric(Trim(parts(j))) And CLng(Trim(parts(j))) > 0 Then
                        NonQueryP conn, "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", Array("i", "i"), Array(newId, CLng(Trim(parts(j))))
                    End If
                Next
            End If
            conn.Close
            Set conn = Nothing
            Response.Redirect "admin_albums.asp?msg=guardado"
        End If
        msgForm = "Corrige los datos del formulario"
    End If
End If

'----------------- Cargar álbum a editar -----------------
If editId > 0 And errMsg = "" Then
    Set rs = QueryRS(conn, "SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array("i"), Array(editId))
    If Not rs.EOF Then
        formTitle = Nz(rs("TITLE_ALBUM"), "")
        formYear = CStr(Nz(rs("RELEASE_YEAR"), ""))
        If formYear = "0" Then formYear = ""
        formImage = Nz(rs("ALBUM_IMAGE_PATH"), "")
        msgForm = "Editar álbum #" & editId
    Else
        editId = 0
    End If
    rs.Close
End If

' Canciones actualmente en el álbum
Set selected = Server.CreateObject("Scripting.Dictionary")
If errMsg <> "" And Request.Form("song_ids") <> "" Then
    parts = Split(Request.Form("song_ids"), ",")
    For j = 0 To UBound(parts)
        If IsNumeric(Trim(parts(j))) And Not selected.Exists(CLng(Trim(parts(j)))) Then
            selected.Add CLng(Trim(parts(j))), True
        End If
    Next
ElseIf editId > 0 Then
    Set rs = QueryRS(conn, "SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array("i"), Array(editId))
    Do While Not rs.EOF
        If Not selected.Exists(CLng(rs("ID_SONG").Value)) Then
            selected.Add CLng(rs("ID_SONG").Value), True
        End If
        rs.MoveNext
    Loop
    rs.Close
End If
%>
<% PageTop "Administrar álbumes", "admin_albums.asp" %>

<main class="container">
    <% FlashIfAny %>

    <div class="admin-grid">
        <section class="card panel">
            <h2><%=msgForm%></h2>
            <p class="muted">El título es obligatorio. Marca las canciones que forman parte del álbum (TBL_TRACKLIST).</p>
            <% If errMsg <> "" Then %><div class="form-error" role="alert"><%=errMsg%></div><% End If %>
            <form method="post" action="admin_albums.asp" class="form-grid">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="id" value="<%=editId%>">
                <div class="field">
                    <label for="album_title">Título del álbum *</label>
                    <input id="album_title" name="album_title" type="text" value="<%=Esc(formTitle)%>" maxlength="50" required>
                </div>
                <div class="field">
                    <label for="release_year">Año de lanzamiento</label>
                    <input id="release_year" name="release_year" type="number" min="1500" max="2100" value="<%=Esc(formYear)%>">
                </div>
                <div class="field">
                    <label for="album_image">Ruta de la imagen</label>
                    <input id="album_image" name="album_image" type="text" value="<%=Esc(formImage)%>" maxlength="255" placeholder="~/albumImages/album1.png">
                </div>
                <div class="field">
                    <label>Canciones del álbum</label>
                    <div class="check-pills">
<%
Set rs = QueryRS(conn, "SELECT ID_SONG, SONG_NAME FROM TBL_SONG ORDER BY ID_SONG", Array(), Array())
Do While Not rs.EOF
    Dim songChecked
    songChecked = selected.Exists(CLng(rs("ID_SONG").Value))
%>
                        <label class="check-pill">
                            <input type="checkbox" name="song_ids" value="<%=rs("ID_SONG")%>"<%=IIf(songChecked, " checked", "")%>>
                            <span><%=Esc(rs("SONG_NAME"))%></span>
                        </label>
<%
    rs.MoveNext
Loop
rs.Close
%>
                    </div>
                </div>
                <div class="row-actions">
                    <a class="btn btn-ghost" href="admin_albums.asp">Limpiar</a>
                    <button type="submit" class="btn btn-primary">Guardar álbum</button>
                </div>
            </form>
        </section>

        <section class="card panel">
            <h2>Álbumes registrados</h2>
            <p class="muted">Al eliminar un álbum se quitan también sus canciones asociadas (TBL_TRACKLIST).</p>
            <div class="table-wrap">
                <table class="table admin-table">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Portada</th>
                            <th>Título</th>
                            <th class="num">Año</th>
                            <th>Imagen</th>
                            <th class="num">Canciones</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody>
<%
sql = "SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
      "(SELECT COUNT(*) FROM TBL_TRACKLIST tl WHERE tl.ID_ALBUM = al.ID_ALBUM) AS CANCIONES " & _
      "FROM TBL_ALBUM al ORDER BY al.ID_ALBUM"
Set rs = QueryRS(conn, sql, Array(), Array())
If rs.EOF Then
%>
                        <tr><td colspan="7" class="empty">No hay álbumes registrados.</td></tr>
<%
End If
Do While Not rs.EOF
    Dim albPhoto
    albPhoto = ImgUrl(rs("ALBUM_IMAGE_PATH"), "album")
%>
                        <tr>
                            <td><%=rs("ID_ALBUM")%></td>
                            <td><img class="thumb" src="<%=Esc(albPhoto)%>" alt=""></td>
                            <td><strong><%=Esc(Nz(rs("TITLE_ALBUM"), ""))%></strong></td>
                            <td class="num"><%=Nz(rs("RELEASE_YEAR"), "—")%></td>
                            <td class="wrap-cell" style="color:var(--muted);font-size:.8rem"><%=Esc(Nz(rs("ALBUM_IMAGE_PATH"), ""))%></td>
                            <td class="num"><%=rs("CANCIONES")%></td>
                            <td>
                                <div class="row-actions">
                                    <a class="btn btn-ghost btn-sm" href="admin_albums.asp?edit=<%=rs("ID_ALBUM")%>">Editar</a>
                                    <form method="post" action="admin_albums.asp" class="inline-form" onsubmit="return confirm('¿Eliminar el álbum «<%=Esc(Nz(rs("TITLE_ALBUM"), ""))%>»?');">
                                        <input type="hidden" name="action" value="delete">
                                        <input type="hidden" name="id" value="<%=rs("ID_ALBUM")%>">
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
