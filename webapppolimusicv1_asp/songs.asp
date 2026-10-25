<%@Language="VBScript" CodePage="65001"%>
<!--#include file="inc_db.asp" -->
<%
'=====================================================================
' songs.asp · Pantalla principal: mensaje de bienvenida, listado de
' canciones (Song List) en tabla HTML con estilos y reproductor.
'=====================================================================
RequireLogin

Dim conn, rs, sql
Dim nSongs, nArtists, nAlbums
Dim json, idx, cover, songSrc, artists, albums, avail

Set conn = OpenConn()

nSongs = Scalar(conn, "SELECT COUNT(*) FROM TBL_SONG", Array(), Array())
nArtists = Scalar(conn, "SELECT COUNT(*) FROM TBL_ARTIST", Array(), Array())
nAlbums = Scalar(conn, "SELECT COUNT(*) FROM TBL_ALBUM", Array(), Array())

' Canciones con sus artistas (TBL_PLAYER) y álbumes (TBL_TRACKLIST)
sql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, s.PLAYS, " & _
      "(SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = tl.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS COVER, " & _
      "STUFF((SELECT ', ' + a.ARTIST_NAME + ISNULL(' ' + NULLIF(a.ARTIST_LASTNAME, ''), '') FROM TBL_PLAYER pl INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = pl.ID_ARTIST WHERE pl.ID_SONG = s.ID_SONG FOR XML PATH('')), 1, 2, '') AS ARTISTAS, " & _
      "STUFF((SELECT ', ' + al.TITLE_ALBUM FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = tl.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG FOR XML PATH('')), 1, 2, '') AS ALBUMES " & _
      "FROM TBL_SONG s ORDER BY s.ID_SONG"

Set rs = QueryRS(conn, sql, Array(), Array())

json = ""
idx = 0
%>
<% PageTop "Song List", "songs" %>

<main class="container">
    <% FlashIfAny %>

    <section class="welcome card">
        <img class="avatar" src="<%=Esc(ImgUrl(Session("USER_PHOTO"), "user"))%>" alt="Foto de perfil">
        <div class="welcome-text">
            <h2 class="welcome-title">Bienvenido, <strong><%=Esc(Session("USERNAME"))%></strong></h2>
            <p class="welcome-sub">Has iniciado sesión en PoliMusic con el rol de <span class="role-pill"><%=IIf(IsAdmin(), "Administrador", "Usuario normal")%></span>.</p>
        </div>
        <% If IsAdmin() Then %>
        <a class="btn btn-primary" href="admin_songs.asp">Administrar catálogo</a>
        <% End If %>
    </section>

    <header class="list-head">
        <p class="eyebrow">Biblioteca PoliMusic</p>
        <h1 class="page-title">Song List</h1>
        <p class="page-sub">A continuación se presenta el listado de canciones registradas en PoliMusic. Pulsa el botón de reproducción para escuchar cada canción.</p>
        <div class="stat-row">
            <span class="chip"><b><%=nSongs%></b> canciones</span>
            <span class="chip"><b><%=nArtists%></b> artistas</span>
            <span class="chip"><b><%=nAlbums%></b> álbumes</span>
        </div>
        <div class="list-tools">
            <div class="search-box">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
                <input id="song-search" type="search" placeholder="Buscar por canción, artista o álbum…" aria-label="Buscar canciones">
            </div>
            <span class="result-count" id="result-count"></span>
        </div>
    </header>

    <div class="table-wrap card">
        <table class="table song-table">
            <caption class="sr-only">Listado de canciones de PoliMusic</caption>
            <thead>
                <tr>
                    <th class="col-idx">#</th>
                    <th class="col-cover">Portada</th>
                    <th>Canción</th>
                    <th>Artista(s)</th>
                    <th>Álbum(es)</th>
                    <th class="num">Reproducciones</th>
                    <th class="col-action">Acciones</th>
                </tr>
            </thead>
            <tbody>
<%
If rs.EOF Then
%>
                <tr><td colspan="7" class="empty">Todavía no hay canciones registradas en la base de datos.</td></tr>
<%
End If

Do While Not rs.EOF
    cover = ImgUrl(rs("COVER"), "album")
    songSrc = SongUrl(rs("SONG_PATH"))
    avail = (songSrc <> "")
    artists = XmlUnescape(rs("ARTISTAS"))
    albums = XmlUnescape(rs("ALBUMES"))
    If artists = "" Then artists = "—"
    If albums = "" Then albums = "—"

    If json <> "" Then json = json & ","
    json = json & "{""id"":" & rs("ID_SONG") & _
                   ",""title"":""" & JsStr(rs("SONG_NAME")) & _
                   ",""artists"":""" & JsStr(artists) & _
                   ",""cover"":""" & JsStr(cover) & _
                   ",""src"":""" & JsStr(songSrc) & _
                   ",""available"":" & LCase(CStr(avail)) & _
                   ",""plays"":" & Nz(rs("PLAYS"), 0) & "}"
%>
                <tr class="song-row" data-index="<%=idx%>">
                    <td class="col-idx"><span class="idx-num"><%=idx + 1%></span></td>
                    <td class="col-cover"><img class="cover" src="<%=Esc(cover)%>" alt="Portada de <%=Esc(rs("SONG_NAME"))%>" loading="lazy"></td>
                    <td class="td-title">
                        <span class="song-title"><%=Esc(rs("SONG_NAME"))%></span>
                        <span class="song-file"><%=Esc(rs("SONG_PATH"))%></span>
                    </td>
                    <td class="td-artists"><%=Esc(artists)%></td>
                    <td class="td-albums"><%=Esc(albums)%></td>
                    <td class="num td-plays" data-plays="<%=Nz(rs("PLAYS"), 0)%>"><%=Nz(rs("PLAYS"), 0)%></td>
                    <td class="col-action">
                        <% If avail Then %>
                        <button type="button" class="btn-play" data-index="<%=idx%>" aria-label="Reproducir <%=Esc(rs("SONG_NAME"))%>">
                            <svg class="i-play" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5.5v13l11-6.5z"></path></svg>
                            <span class="eq" aria-hidden="true"><i></i><i></i><i></i></span>
                        </button>
                        <% Else %>
                        <span class="no-audio" title="El archivo MP3 no se encuentra en SongFilesAndImages_resources/SongFiles">Sin audio</span>
                        <% End If %>
                    </td>
                </tr>
<%
    idx = idx + 1
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
            </tbody>
        </table>
        <p class="empty-row" id="no-results" hidden>No se encontraron canciones que coincidan con la búsqueda.</p>
    </div>
    <noscript><p class="empty-row">Habilita JavaScript para usar el reproductor y la búsqueda.</p></noscript>
</main>

<% '---------- Reproductor fijo ----------%>
<div class="player" id="player">
    <div class="player-track">
        <img class="player-cover" id="pm-cover" src="<%=FallbackUrl("album")%>" alt="">
        <div class="player-meta">
            <strong id="pm-title">Selecciona una canción</strong>
            <span id="pm-artists">PoliMusic</span>
        </div>
    </div>
    <div class="player-center">
        <div class="ctrl-row">
            <button type="button" class="ctrl" id="pm-prev" aria-label="Canción anterior">
                <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 5h2v14H6zM20 5v14L9.5 12z"></path></svg>
            </button>
            <button type="button" class="ctrl ctrl-main" id="pm-toggle" aria-label="Reproducir">
                <svg class="i-play" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5.5v13l11-6.5z"></path></svg>
                <svg class="i-pause" viewBox="0 0 24 24" aria-hidden="true"><path d="M7 5h3.5v14H7zM13.5 5H17v14h-3.5z"></path></svg>
            </button>
            <button type="button" class="ctrl" id="pm-next" aria-label="Siguiente canción">
                <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M16 5h2v14h-2zM4 5v14l10.5-7z"></path></svg>
            </button>
        </div>
        <div class="seek-row">
            <span class="time" id="pm-cur">0:00</span>
            <input type="range" id="pm-seek" min="0" max="1000" value="0" aria-label="Progreso de reproducción">
            <span class="time" id="pm-dur">0:00</span>
        </div>
    </div>
    <div class="player-extra">
        <span class="queue-count" id="pm-count">0 en cola</span>
        <button type="button" class="ctrl" id="pm-mute" aria-label="Silenciar">
            <svg class="i-vol" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 9v6h4l5 4V5L8 9z"></path><path d="M16.5 8.5a5 5 0 0 1 0 7" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"></path><path d="M19 6a8.5 8.5 0 0 1 0 12" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"></path></svg>
            <svg class="i-muted" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 9v6h4l5 4V5L8 9z"></path><path d="M16 9.5l5 5M21 9.5l-5 5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"></path></svg>
        </button>
        <input type="range" id="pm-vol" min="0" max="100" value="80" aria-label="Volumen">
    </div>
</div>

<div class="toast" id="pm-toast" role="status" aria-live="polite"></div>

<script>
window.PM_QUEUE = [<%=json%>];
</script>

<% PageBottom %>
