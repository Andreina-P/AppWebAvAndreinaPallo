<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>PoliMusic - Songs</title>
    <style type="text/css">
        body { margin: 0; font-family: Segoe UI, Arial, sans-serif; background: #f3ecfb; color: #3b2a55; }
        .topbar { background: #b79ae0; color: #ffffff; padding: 14px 30px; overflow: hidden; }
        .brand { float: left; font-size: 22px; font-weight: bold; }
        .user { float: right; font-size: 15px; }
        .icon { display: inline-block; width: 30px; height: 30px; line-height: 30px; text-align: center;
                background: #ffffff; color: #8e6fc4; border-radius: 50%; margin-right: 8px; }
        .user a { color: #ffffff; margin-left: 15px; text-decoration: underline; }
        .content { width: 800px; margin: 30px auto; background: #ffffff; padding: 25px; border-radius: 12px; }
        h2 { color: #8e6fc4; margin-top: 0; }
        .grid { width: 100%; border-collapse: collapse; }
        .grid th { background: #d9c8f2; color: #3b2a55; text-align: left; padding: 10px; }
        .grid td { padding: 10px; border-bottom: 1px solid #eadff7; }
        .grid tr:hover td { background: #f8f3fd; }
        audio { width: 300px; height: 32px; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div class="topbar">
            <span class="brand">PoliMusic</span>
            <span class="user">
                <span class="icon">&#128100;</span>
                <asp:Label ID="lblUserName" runat="server"></asp:Label>
                <asp:HyperLink ID="lnkLogout" runat="server" NavigateUrl="~/Logout.aspx">Cerrar sesión</asp:HyperLink>
            </span>
        </div>
        <div class="content">
            <h2>Songs</h2>
            <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False" CssClass="grid" GridLines="None">
                <Columns>
                    <asp:BoundField DataField="ID_SONG" HeaderText="#" />
                    <asp:BoundField DataField="SONG_NAME" HeaderText="Song" />
                    <asp:BoundField DataField="PLAYS" HeaderText="Plays" />
                    <asp:TemplateField HeaderText="Play">
                        <ItemTemplate>
                            <audio controls="controls" preload="none" src='<%# ResolveUrl(Eval("SONG_PATH").ToString()) %>'>
                                Your browser does not support audio.
                            </audio>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
            <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
        </div>
    </form>
</body>
</html>
