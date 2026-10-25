<%@Language="VBScript" CodePage="65001"%>
<!--#include file="inc_db.asp" -->
<%
'=====================================================================
' api_play.asp · Suma una reproducción (PLAYS) a una canción.
' Llamado por AJAX desde assets/js/player.js. Responde "OK" o "ERR|..."
'=====================================================================
If Not IsLoggedIn() Then
    Response.Status = "403 Forbidden"
    Response.Write "ERR|sin sesion"
    Response.End
End If

Dim sid, conn
sid = Request.Form("id")

If Not IsNumeric(sid) Or CLng(sid) <= 0 Then
    Response.Write "ERR|id invalido"
    Response.End
End If

Set conn = OpenConn()
NonQueryP conn, "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?", Array("i"), Array(CLng(sid))
conn.Close
Set conn = Nothing

Response.Write "OK"
