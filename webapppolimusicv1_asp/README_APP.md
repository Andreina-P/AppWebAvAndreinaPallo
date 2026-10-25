# PoliMusic · Reproductor web (ASP Clásico + SQL Server)

Sitio web estilo Spotify restringido al contenido de la base de datos
**BDD_PoliMusic** (script `BDD_PoliMusic.sql`). Desarrollado en **ASP clásico
con VBScript** (páginas `.asp`), HTML + CSS + JavaScript propios, sin
frameworks ni tablas adicionales a las del script.

## Requisitos

1. **SQL Server** en `localhost` con la base creada:
   - Ejecutar `BDD_PoliMusic.sql` en SQL Server Management Studio.
   - Crea el login `usr_polimusic` / `Politecnica1` (dueño de la base).
2. **IIS con la característica ASP clásico habilitada**:
   - Panel de control → Programas → Activar/desactivar características de Windows →
     **Internet Information Services → Características de desarrollo de
     aplicaciones → ASP**.
   - Opcional por consola (como administrador):
     `dism /online /enable-feature /featurename:IIS-ASP /all`

## Publicar el sitio en IIS

1. Copiar esta carpeta (`webapppolimusicv1_asp`) al servidor, p. ej.
   `C:/inetpub/wwwroot/PoliMusic` (o usar esta ruta directamente).
2. En el Administrador de IIS: *Sitios → Agregar sitio web* (o *Agregar
   aplicación* dentro de Default Web Site) con:
   - Ruta física: la carpeta del proyecto.
   - Grupo de aplicaciones: sin código administrado (.NET CLR = "Sin código
     administrado"), modo integrado.
3. Verificar que *Documento predeterminado* incluya `default.asp`
   (o entrar directo a `login.asp`).
4. Abrir `http://localhost/PoliMusic/` (redirige a `login.asp`).

> La cadena de conexión está en `inc_db.asp`
> (`Provider=SQLOLEDB;Data Source=localhost;...`). Si el servidor o las
> credenciales cambian, modificar esa constante.

## Cuentas de prueba

| Usuario | Contraseña | Rol |
|---|---|---|
| `userNormal` | `12345678` | Usuario normal (solo escucha) |
| `userAdmin` | `12345678` | Administrador (CRUD completo) |

Notas:
- El usuario `username1` del script guarda su contraseña como hash SHA-1;
  el inicio de sesión compara texto plano, por lo que no podrá ingresar
  (las dos cuentas anteriores sí funcionan).

## Funcionalidades

- **Login** (`login.asp`): valida contra `TBL_USER` y guarda las variables de
  sesión `ID_USER`, `USERNAME`, `USER_TYPE`, `USER_PHOTO`.
- **Song List** (`songs.asp`): mensaje de bienvenida con el nombre del usuario
  leído de la sesión, tabla HTML con estilos de todas las canciones
  (`TBL_SONG` + artistas de `TBL_PLAYER` + álbumes de `TBL_TRACKLIST`),
  buscador en vivo y reproductor de audio (play/pausa, anterior/siguiente,
  barra de progreso, volumen; cada reproducción suma en `PLAYS` vía
  `api_play.asp`).
- **Roles**: `USER_TYPE = 0` usuario normal, `1` administrador. Solo el
  administrador ve el menú de administración.
- **CRUD administrativo** (solo admin): canciones (`admin_songs.asp`,
  incluye asignación de artistas), artistas (`admin_artists.asp`) y álbumes
  (`admin_albums.asp`, incluye tracklist). Usa únicamente las tablas del
  script; al eliminar se limpian primero las filas hijas para no violar las
  llaves foráneas.
- **Cierre de sesión** (`logout.asp`): `Session.Abandon()` y regreso al login.

## Estructura

```
webapppolimusicv1_asp/
├── BDD_PoliMusic.sql            Script de la base de datos
├── SongFilesAndImages_resources/  Recursos (mp3 e imágenes)
├── inc_db.asp                   Conexión, helpers, sesión/roles y layout
├── login.asp  logout.asp  default.asp  api_play.asp
├── songs.asp                    Listado de canciones + reproductor
├── admin_songs.asp  admin_artists.asp  admin_albums.asp
└── assets/
    ├── css/styles.css           Diseño propio (paleta carbón + ámbar)
    ├── js/player.js             Reproductor, búsqueda y conteo de plays
    └── img/logo.svg             Logotipo
```

## Recursos y comportamiento esperado

- El código resuelve las rutas de la BD (`~/songFiles/...`, `~/userImages/...`,
  etc.) contra la carpeta `SongFilesAndImages_resources` (sin distinguir
  mayúsculas) y busca el archivo en cualquier subcarpeta.
- Si un archivo **no existe**, se usa una imagen por defecto; si falta el mp3,
  la canción muestra la etiqueta **"Sin audio"** y el botón se deshabilita.
  Al agregar los archivos con los nombres del script, aparecerán
  automáticamente (también puede crear subcarpetas nuevas como
  `AlbumImages/` o `ArtistImages/`).
- Solo `bensound-adventure.mp3` está incluido en los recursos de este
  proyecto; los demás mp3 del script pueden copiarse a
  `SongFilesAndImages_resources/SongFiles/` con los nombres indicados en
  `TBL_SONG.SONG_PATH`.
