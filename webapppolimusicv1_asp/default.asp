<%@Language="VBScript" CodePage="65001"%>
<!--#include file="inc_db.asp" -->
<%
'=====================================================================
' default.asp · Punto de entrada: envía al listado o al login
'=====================================================================
If IsLoggedIn() Then
    Response.Redirect "songs.asp"
Else
    Response.Redirect "login.asp"
End If
%>
