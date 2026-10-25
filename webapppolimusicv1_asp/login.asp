<%@Language="VBScript" CodePage="65001"%>
<!--#include file="inc_db.asp" -->
<%
'=====================================================================
' login.asp · Pantalla de inicio de sesión de PoliMusic
' Valida las credenciales contra TBL_USER y guarda las variables de
' sesión que usarán las demás páginas.
'=====================================================================
If IsLoggedIn() Then Response.Redirect "songs.asp"

Dim errMsg, infoMsg, fUser
errMsg = ""
infoMsg = ""
fUser = Trim(Request.Form("username"))

If Request.QueryString("msg") = "sesion" Then infoMsg = "Sesión finalizada correctamente."

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim fPass, conn, rs
    fPass = Request.Form("password")
    If fUser = "" Or fPass = "" Then
        errMsg = "Debes ingresar usuario y contraseña."
    Else
        Set conn = OpenConn()
        Set rs = QueryRS(conn, "SELECT ID_USER, USERNAME, USER_TYPE, USER_PHOTO FROM TBL_USER WHERE USERNAME = ? AND PASSWORD = ?", Array("s", "s"), Array(fUser, fPass))
        If rs.EOF Then
            errMsg = "Usuario o contraseña incorrectos. Verifica tus datos."
            rs.Close
        Else
            ' Variables de sesión (las leen songs.asp y las páginas admin)
            Session("ID_USER") = rs("ID_USER")
            Session("USERNAME") = rs("USERNAME")
            Session("USER_TYPE") = CInt(rs("USER_TYPE"))
            Session("USER_PHOTO") = Nz(rs("USER_PHOTO"), "")
            Session.Timeout = 60
            rs.Close
            conn.Close
            Set conn = Nothing
            Response.Redirect "songs.asp"
        End If
        conn.Close
        Set conn = Nothing
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Iniciar sesión · PoliMusic</title>
<link rel="icon" type="image/svg+xml" href="assets/img/logo.svg">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="assets/css/styles.css">
</head>
<body class="login-body">
<main class="login-wrap">
    <section class="brand-panel">
        <a class="brand" href="default.asp">
            <img class="brand-mark" src="assets/img/logo.svg" alt="">
            <span class="brand-name">Poli<strong>Music</strong></span>
        </a>
        <div class="vinyl" aria-hidden="true"></div>
        <div>
            <h1 class="brand-title">Tu música,<br><strong>sin interrupciones</strong></h1>
            <p class="brand-tag">PoliMusic es el reproductor web del politécnico: escucha el catálogo completo de canciones, artistas y álbumes desde cualquier navegador.</p>
        </div>
        <ul class="brand-points">
            <li>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 18V5l12-2v13"/><circle cx="6" cy="18" r="3"/><circle cx="18" cy="16" r="3"/></svg>
                Catálogo completo de canciones
            </li>
            <li>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l7 3v5c0 4.5-3 7.5-7 9-4-1.5-7-4.5-7-9V6z"/></svg>
                Acceso con roles: usuario y administrador
            </li>
            <li>
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M10 8.5l6 3.5-6 3.5z" fill="currentColor"/></svg>
                Reproducción integrada en la lista
            </li>
        </ul>
        <div class="eq-big" aria-hidden="true"><i></i><i></i><i></i><i></i><i></i></div>
    </section>
    <section class="form-panel">
        <div class="login-card">
            <h2>Iniciar sesión</h2>
            <p class="muted">Ingresa con tu cuenta de PoliMusic para continuar.</p>
            <% If errMsg <> "" Then %><div class="alert alert-error" role="alert"><%=errMsg%></div><% End If %>
            <% If infoMsg <> "" Then %><div class="alert alert-ok" role="status"><%=infoMsg%></div><% End If %>
            <form method="post" action="login.asp">
                <div class="field">
                    <label for="username">Usuario</label>
                    <input id="username" name="username" type="text" value="<%=Esc(fUser)%>" placeholder="userNormal" autocomplete="username" autofocus required>
                </div>
                <div class="field">
                    <label for="password">Contraseña</label>
                    <div class="pass-wrap">
                        <input id="password" name="password" type="password" placeholder="••••••••" autocomplete="current-password" required>
                        <button type="button" class="pass-toggle" id="pass-toggle" aria-label="Mostrar contraseña">Mostrar</button>
                    </div>
                </div>
                <button type="submit" class="btn btn-primary btn-block">Entrar a PoliMusic</button>
            </form>
            <div class="demo-creds">
                <h3>Cuentas de prueba</h3>
                <p>Usuario normal: <code>userNormal</code> · <code>12345678</code></p>
                <p>Administrador: <code>userAdmin</code> · <code>12345678</code></p>
            </div>
        </div>
    </section>
</main>
<script>
document.getElementById("pass-toggle").addEventListener("click", function () {
    var input = document.getElementById("password");
    var show = input.type === "password";
    input.type = show ? "text" : "password";
    this.textContent = show ? "Ocultar" : "Mostrar";
    this.setAttribute("aria-label", show ? "Ocultar contraseña" : "Mostrar contraseña");
});
</script>
</body>
</html>
