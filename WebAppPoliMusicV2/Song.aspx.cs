using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;

using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                Entities.User user = Session[Util.Constants.USER] as Entities.User;
                if (user != null)
                {
                    lblUserName.Text = user.UserName;
                }
                else
                {
                    lblUserName.Text = User.Identity.Name;
                }
                LoadSongs();
            }
        }

        private void LoadSongs()
        {
            try
            {
                string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;
                DataTable dtSong = new BLL.Song(strConnString).Read();
                gvSongs.DataSource = dtSong;
                gvSongs.DataBind();
            }
            catch (Exception ex)
            {
                lblMessage.Text = "Could not load songs: " + ex.Message;
            }
        }
    }
}
