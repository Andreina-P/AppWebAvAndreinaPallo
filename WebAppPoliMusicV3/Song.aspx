<%@ Page Title="Canciones" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV3.Song" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <style>
        .pm-wrap { background:#B8A9E8; color:#2A1F4D; padding:28px; border-radius:18px; margin-top:20px; }
        .pm-header { display:flex; align-items:center; gap:16px; margin-bottom:24px; }
        .pm-header img { width:72px; height:72px; border-radius:50%; object-fit:cover; border:4px solid #FFD43B; background:#fff; }
        .pm-hello { font-size:.8rem; text-transform:uppercase; letter-spacing:2px; opacity:.8; display:block; }
        .pm-username { font-size:1.6rem; font-weight:800; }
        .pm-title { font-weight:800; margin:0 0 14px 0; }
        .pm-title span { background:#FFD43B; padding:2px 12px; border-radius:20px; font-size:1rem; margin-left:8px; }
        .pm-grid { width:100%; border-collapse:separate; border-spacing:0 8px; }
        .pm-grid th { text-transform:uppercase; font-size:.75rem; letter-spacing:1px; padding:6px 12px; text-align:left; }
        .pm-grid td { background:#CFC3F2; padding:10px 12px; vertical-align:middle; }
        .pm-grid tr:hover td { background:#E3DAFB; }
        .pm-grid td:first-child { border-radius:12px 0 0 12px; }
        .pm-grid td:last-child { border-radius:0 12px 12px 0; }
        .pm-num { background:#FFD43B !important; font-weight:800; text-align:center; width:48px; }
        .pm-grid audio { width:100%; min-width:220px; height:36px; }
        .pm-cover { width:52px; height:52px; border-radius:8px; object-fit:cover; border:2px solid #FFD43B; background:#fff; }
        .pm-song { display:flex; align-items:center; gap:14px; }
        .pm-song-name { font-weight:700; display:block; }
        .pm-song-artist { font-size:.85rem; opacity:.8; display:block; }
    </style>

    <div class="pm-wrap">
        <div class="pm-header">
            <asp:Image ID="imgUser" runat="server" />
            <div>
                <span class="pm-hello">Bienvenido</span>
                <asp:Label ID="lblUser" runat="server" CssClass="pm-username"></asp:Label>
            </div>
        </div>

        <h2 class="pm-title">Canciones<span>PoliMusic</span></h2>

        <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False" CssClass="pm-grid" GridLines="None">
            <Columns>
                <asp:TemplateField HeaderText="#">
                    <ItemTemplate><%# Container.DataItemIndex + 1 %></ItemTemplate>
                    <ItemStyle CssClass="pm-num" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Título">
                    <ItemTemplate>
                        <div class="pm-song">
                            <img class="pm-cover"
                                 src='<%# ResolveUrl(Eval("COVER").ToString()) %>'
                                 alt='<%# Eval("SONG_NAME") %>'
                                 onerror="this.onerror=null;this.src='<%# ResolveUrl("~/PlaylistImages/default.jpg") %>'" />
                            <div>
                                <span class="pm-song-name"><%# Eval("SONG_NAME") %></span>
                                <span class="pm-song-artist"><%# Eval("ARTISTS") %></span>
                            </div>
                        </div>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="ALBUM" HeaderText="Álbum" />
                <asp:BoundField DataField="PLAYS" HeaderText="Reproducciones" />
                <asp:TemplateField HeaderText="Reproducir">
                    <ItemTemplate>
                        <audio controls preload="none" src='<%# ResolveUrl(Eval("SONG_PATH").ToString()) %>'></audio>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>

        <asp:Label ID="lblMessage" runat="server" Font-Bold="true"></asp:Label>
    </div>
</asp:Content>