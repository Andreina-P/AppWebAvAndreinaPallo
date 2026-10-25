<%@Language="VBScript" CodePage="65001"%>
<!--#include file="inc_db.asp" -->
<%
'=====================================================================
' logout.asp · Cierra la sesión y regresa al login
'=====================================================================
Session.Abandon()
Response.Redirect "login.asp?msg=sesion"
%>
