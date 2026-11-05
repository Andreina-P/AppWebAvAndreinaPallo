using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV3
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadUser();
                LoadSongs();
            }
        }

        private void LoadUser()
        {
            EntityLayer.PoliMusicV3.User user = Session[EntityLayer.PoliMusicV3.Util.Constants.USER] as EntityLayer.PoliMusicV3.User;

            // Si la sesión se perdió (recompilación, reinicio), se recupera desde la BD
            // usando el nombre de usuario de la cookie de login.
            if (user == null && User.Identity.IsAuthenticated)
            {
                string name = User.Identity.Name;
                DataTable dtUsers = new BLL.PoliMusicV3.User(strConnString).Read();
                DataRow[] rows = dtUsers.Select("USERNAME = '" + name.Replace("'", "''") + "'");
                if (rows.Length > 0)
                {
                    user = new EntityLayer.PoliMusicV3.User();
                    user.ID = Convert.ToInt32(rows[0]["ID_USER"]);
                    user.UserName = name;
                    user.Photo = rows[0]["USER_PHOTO"].ToString();
                    user.Email = rows[0]["EMAIL"].ToString();
                    user.UserType = Convert.ToInt32(rows[0]["USER_TYPE"]);
                    Session[EntityLayer.PoliMusicV3.Util.Constants.USER] = user;
                }
            }

            if (user != null)
            {
                lblUser.Text = user.UserName;
                imgUser.ImageUrl = string.IsNullOrEmpty(user.Photo) ? "~/UserImages/default.jpg" : user.Photo;
                imgUser.AlternateText = user.UserName;
            }
            else
            {
                lblUser.Text = User.Identity.Name;
                imgUser.ImageUrl = "~/UserImages/default.jpg";
            }
        }

        private void LoadSongs()
        {
            DataTable dtSong = new BLL.PoliMusicV3.Song(strConnString).Read();
            if (dtSong != null && dtSong.Rows.Count > 0)
            {
                gvSongs.DataSource = dtSong;
                gvSongs.DataBind();
                lblMessage.Text = string.Empty;
            }
            else
            {
                lblMessage.Text = "No songs found.";
            }
        }
    }
}
