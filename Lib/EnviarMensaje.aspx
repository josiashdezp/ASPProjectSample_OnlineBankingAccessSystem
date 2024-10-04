<%@ Page Language="C#" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web" %>
<%@ Import Namespace="System.Web.Security" %>
<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="System.Web.UI.WebControls" %>
<%@ Import Namespace="System.Web.UI.WebControls.WebParts" %>
<%@ Import Namespace="System.Web.UI.HtmlControls" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">
 protected void Page_Load(object sender, EventArgs e)
    {

     string where = Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath) + "Var\\";
     
     string To = Request["To"].ToString();
     string Cc = Request["Cc"].ToString();
     
     string MSG_Id = "MSG" + Request["MSG"].ToString() + ".html";

     string Path = where + MSG_Id;

     System.IO.FileInfo filesys = new System.IO.FileInfo(Path);
     System.IO.StreamReader Contenido;

     Contenido = filesys.OpenText();

     System.Net.Mail.MailMessage Mensaje = new System.Net.Mail.MailMessage();

     Mensaje.From = new System.Net.Mail.MailAddress("banca_virtual@dpcf.bandec.cu");
     Mensaje.To.Add(To);

     if (Cc != "")  Mensaje.CC.Add(Cc);

     Mensaje.Subject = "Bandec Online. Estado de Cuentas: " + DateTime.Now.ToString();
     Mensaje.Body = Contenido.ReadToEnd();
     Mensaje.IsBodyHtml = true;

     Mensaje.Priority = System.Net.Mail.MailPriority.High;

     System.Net.Mail.SmtpClient Servidor = new System.Net.Mail.SmtpClient();
     Servidor.Host = "mail.dpcf.bandec.cu";
     Servidor.Credentials = new System.Net.NetworkCredential("banca_virtual","bandeccfg");

     try
     {
         Servidor.Send(Mensaje);
         lbl_Mensaje.Text = "Mensaje enviado satisfactoriamente";
         lbl_Error.Text = DateTime.Now.ToString();
         

     }
     catch (Exception ex)
     {
         lbl_Mensaje.Text = "ERROR: " + ex.Message;
         lbl_Mensaje.Style.Add(HtmlTextWriterStyle.Color, "Tomato");
     }

     Mensaje.Dispose();
     Contenido.Close();  
    }
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="text-align: center">
        <br />
        <br />
        <br />
        &nbsp;
        <div style="width: 454px; height: 210px; background-color: #ffffcc">
           <br />
            <img src="../images/email.gif" alt="Mensaje Enviado" />
            <p>
              <asp:Label ID="lbl_Mensaje" runat="server" Font-Bold="True" Font-Size="Large" Text="Label"></asp:Label>
                <br />
                <br />
                <asp:Label ID="lbl_Error" runat="server" Font-Bold="True" Font-Size="Medium" ForeColor="Tomato"
                Text="Label"></asp:Label>
              <br />
                <br />
              Volver a &nbsp;
              <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../Servicios/Estado_Cuenta.asp">Estado de Cuentas ...</asp:HyperLink>
              <br />
                <br />
                </p>
        </div>
    </div>
    </form>
</body>
</html>
