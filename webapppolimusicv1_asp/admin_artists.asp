<%@Language="VBScript" CodePage="65001"%>
<!--#include file="inc_db.asp" -->
<%
'=====================================================================
' admin_artists.asp · CRUD de artistas (solo rol administrador)
' Tablas usadas: TBL_ARTIST, TBL_PLAYER (relación con canciones)
'=====================================================================
RequireAdmin

Dim conn, rs, sql
Dim editId, formName, formLastname, formImage
Dim errMsg, msgForm, action, id
Dim fName, fLastname, fImage, newId

editId = 0
formName = ""
formLastname = ""
formImage = ""
errMsg = ""
msgForm = "Nuevo artista"
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
            NonQueryP conn, "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array("i"), Array(CLng(id))
            NonQueryP conn, "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array("i"), Array(CLng(id))
        End If
        conn.Close
        Set conn = Nothing
        Response.Redirect "admin_artists.asp?msg=borrado"
    Else
        editId = CLng(id)
        fName = Trim(Request.Form("artist_name"))
        fLastname = Trim(Request.Form("artist_lastname"))
        fImage = Trim(Request.Form("artist_image"))
        formName = fName
        formLastname = fLastname
        formImage = fImage
        If fName = "" Then errMsg = errMsg & "El nombre del artista es obligatorio.<br>"

        If errMsg = "" Then
            If editId > 0 Then
                NonQueryP conn, "UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?", Array("s", "s", "s", "i"), Array(fName, fLastname, fImage, editId)
            Else
                newId = InsertId(conn, "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?); SELECT CAST(SCOPE_IDENTITY() AS int)", Array("s", "s", "s"), Array(fName, fLastname, fImage))
            End If
            conn.Close
            Set conn = Nothing
            Response.Redirect "admin_artists.asp?msg=guardado"
        End If
        msgForm = "Corrige los datos del formulario"
    End If
End If

'----------------- Cargar artista a editar -----------------
If editId > 0 And errMsg = "" Then
    Set rs = QueryRS(conn, "SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array("i"), Array(editId))
    If Not rs.EOF Then
        formName = Nz(rs("ARTIST_NAME"), "")
        formLastname = Nz(rs("ARTIST_LASTNAME"), "")
        formImage = Nz(rs("ARTIST_IMAGE"), "")
        msgForm = "Editar artista #" & editId
    Else
        editId = 0
    End If
    rs.Close
End If
%>
<% PageTop "Administrar artistas", "admin_artists.asp" %>

<main class="container">
    <% FlashIfAny %>

    <div class="admin-grid">
        <section class="card panel">
            <h2><%=msgForm%></h2>
            <p class="muted">El nombre es obligatorio. La imagen debe existir en <b>SongFilesAndImages_resources</b> (si no, se usa una imagen por defecto).</p>
            <% If errMsg <> "" Then %><div class="form-error" role="alert"><%=errMsg%></div><% End If %>
            <form method="post" action="admin_artists.asp" class="form-grid">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="id" value="<%=editId%>">
                <div class="field">
                    <label for="artist_name">Nombre *</label>
                    <input id="artist_name" name="artist_name" type="text" value="<%=Esc(formName)%>" maxlength="100" required>
                </div>
                <div class="field">
                    <label for="artist_lastname">Apellido</label>
                    <input id="artist_lastname" name="artist_lastname" type="text" value="<%=Esc(formLastname)%>" maxlength="100">
                </div>
                <div class="field">
                    <label for="artist_image">Ruta de la imagen</label>
                    <input id="artist_image" name="artist_image" type="text" value="<%=Esc(formImage)%>" maxlength="255" placeholder="~/ArtistImages/artist1.jpg">
                </div>
                <div class="row-actions">
                    <a class="btn btn-ghost" href="admin_artists.asp">Limpiar</a>
                    <button type="submit" class="btn btn-primary">Guardar artista</button>
                </div>
            </form>
        </section>

        <section class="card panel">
            <h2>Artistas registrados</h2>
            <p class="muted">Al eliminar un artista se quitan también sus canciones asociadas (TBL_PLAYER).</p>
            <div class="table-wrap">
                <table class="table admin-table">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Foto</th>
                            <th>Nombre</th>
                            <th>Apellido</th>
                            <th>Imagen</th>
                            <th>Canciones</th>
                            <th></th>
                        </tr>
                    </thead>
                    <tbody>
<%
sql = "SELECT ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " & _
      "(SELECT COUNT(*) FROM TBL_PLAYER pl WHERE pl.ID_ARTIST = ar.ID_ARTIST) AS CANCIONES " & _
      "FROM TBL_ARTIST ar ORDER BY ar.ID_ARTIST"
Set rs = QueryRS(conn, sql, Array(), Array())
If rs.EOF Then
%>
                        <tr><td colspan="7" class="empty">No hay artistas registrados.</td></tr>
<%
End If
Do While Not rs.EOF
    Dim fullName, artPhoto
    fullName = Trim(Nz(rs("ARTIST_NAME"), "") & " " & Nz(rs("ARTIST_LASTNAME"), ""))
    artPhoto = ImgUrl(rs("ARTIST_IMAGE"), "artist")
%>
                        <tr>
                            <td><%=rs("ID_ARTIST")%></td>
                            <td><img class="thumb" src="<%=Esc(artPhoto)%>" alt=""></td>
                            <td><strong><%=Esc(Nz(rs("ARTIST_NAME"), ""))%></strong></td>
                            <td><%=Esc(Nz(rs("ARTIST_LASTNAME"), ""))%></td>
                            <td class="wrap-cell" style="color:var(--muted);font-size:.8rem"><%=Esc(Nz(rs("ARTIST_IMAGE"), ""))%></td>
                            <td class="num"><%=rs("CANCIONES")%></td>
                            <td>
                                <div class="row-actions">
                                    <a class="btn btn-ghost btn-sm" href="admin_artists.asp?edit=<%=rs("ID_ARTIST")%>">Editar</a>
                                    <form method="post" action="admin_artists.asp" class="inline-form" onsubmit="return confirm('¿Eliminar al artista «<%=Esc(fullName)%>»? También se quitará de sus canciones.');">
                                        <input type="hidden" name="action" value="delete">
                                        <input type="hidden" name="id" value="<%=rs("ID_ARTIST")%>">
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
