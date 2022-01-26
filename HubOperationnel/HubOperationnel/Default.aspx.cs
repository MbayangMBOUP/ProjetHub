using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Web.UI;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;
using System.Xml;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Diagnostics;

public partial class _Default : Page
{
    protected void Page_Load()
    {

        if (!IsPostBack) //Permet de ne pas recharger la page après une requête "client"
        {
            SqlConnection sqlConnection = Connexion();

            string professeurs = "SELECT * from Professeurs";
            SqlDataAdapter adpt = new SqlDataAdapter(professeurs, sqlConnection);
            DataTable dt = new DataTable();
            adpt.Fill(dt);

            if (dt != null && dt.Rows.Count > 0)
            {
                Button1.Visible = false;
                LoadDropDowns();
            }
            else
            {
                Button1.Visible = true;
            }
                

        }

    }
    protected void LoadDropDowns()
    {
        
        string professeurs = "SELECT * from Professeurs";
        DropDownData(professeurs, "idProf", "nomProf", DropDownProfesseurs);

        string modules = "SELECT * from Modules m JOIN enseignerModules em ON m.idModule = em.idModule where em.idProfesseur = '" +
           DropDownProfesseurs.SelectedValue + "'";
        DropDownData(modules, "idModule", "nomModule", DropDownListModules);

        string promo = "SELECT * from Promotions";
        DropDownData(promo, "idPromotions", "nomPromotions", DropDownListPromotions);

        string groupe = "SELECT * from Groupes";
        DropDownData(groupe, "idGroupes", "nomGroupes", DropDownListGroupes);


    }

        protected void LoadDropDownModule(object sender, EventArgs e)
    {
      
        string modules = "SELECT * from Modules m JOIN enseignerModules em ON m.idModule = em.idModule where em.idProfesseur = '" +
           DropDownProfesseurs.SelectedValue + "'";

        DropDownData(modules, "idModule", "nomModule", DropDownListModules);

    }

    //pour charger les dropdown
    protected void DropDownData(String req, String id, String nom, DropDownList dropd)
    {
        SqlConnection sqlConnection = Connexion();
        
        SqlDataAdapter adpt = new SqlDataAdapter(req, sqlConnection);
        DataTable dt = new DataTable();
        adpt.Fill(dt);
        dropd.DataSource = dt;
        dropd.DataBind();
        dropd.DataTextField = nom;
        dropd.DataValueField = id;
        dropd.DataBind();
    }


    // Recuperer les données d'un fichier JSON
    protected String ListDataJson(string nomFichier){
         
        StreamReader r = new StreamReader("C:\\USERS\\YANGM_9VNEY6U\\SOURCE\\REPOS\\ProjetHub\\BD\\" + nomFichier);//enseignerModules4.json
        string jsonString = r.ReadToEnd();
        return jsonString;
    }
    

// Connection à la base de données
    protected SqlConnection Connexion() {

        string PARAMS_INTEROP =
             "Data Source = (LocalDB)\\MSSQLLocalDB;" +
               "AttachDbFilename = C:\\USERS\\YANGM_9VNEY6U\\SOURCE\\REPOS\\PROJETHUB\\HUBOPERATIONNEL\\HUBOPERATIONNEL\\APP_DATA\\DATAHUB.MDF;" +
                "Integrated Security = True";

        SqlConnection connection = new SqlConnection(PARAMS_INTEROP);
        return connection;
    }


    //charger la table enseignerModules
    protected void DataEnseignerModules()
    {
        SqlConnection connection = Connexion();
        connection.Open();

        String jsonString = ListDataJson("enseignerModules.json");

        List<DataEnseignModul> m = JsonConvert.DeserializeObject<List<DataEnseignModul>>(jsonString);

        foreach (var ligne in m)
        {
            string query = "INSERT INTO enseignerModules (idModule, idProfesseur) VALUES(@idModule, @idProfesseur)";
            SqlCommand command = new SqlCommand(query, connection);
        
            command.Parameters.AddWithValue("@idModule", ligne.IdModule);
            command.Parameters.AddWithValue("@idProfesseur", ligne.IdProfesseur);
            command.ExecuteNonQuery();
        }  

    }


    //charger la table Professeurs
    protected void DataProfesseur()
    {
        SqlConnection connection = Connexion();
        connection.Open();

        String jsonString = ListDataJson("professeurs.json");
        List<DataProf> m = JsonConvert.DeserializeObject<List<DataProf>>(jsonString);

        foreach (var ligne in m)
        {
            string query = "INSERT INTO Professeurs (idProf, nomProf, prenomProf) VALUES(@idProf, @nomProf, @prenomProf)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@idProf", ligne.IdProf);
            command.Parameters.AddWithValue("@nomProf", ligne.NomProf);
            command.Parameters.AddWithValue("@prenomProf", ligne.PrenomProf);
            command.ExecuteNonQuery();
        }

    }


    //charger la table Modules
    protected void DataModules()
    {
        SqlConnection connection = Connexion();
        connection.Open();

        String jsonString = ListDataJson("modules.json");

        List<DataModul> m = JsonConvert.DeserializeObject<List<DataModul>>(jsonString);

        foreach (var ligne in m)
        {
            string query = "INSERT INTO Modules (idModule, nomModule) VALUES(@idModule, @nomModule)";
            SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.AddWithValue("@idModule", ligne.IdModule);
            command.Parameters.AddWithValue("@nomModule", ligne.NomModule);
            command.ExecuteNonQuery();
        }

    }

    protected OleDbConnection ExcelConnexion()
    {
        String Fournisseur = "Provider=Microsoft.Jet.OLEDB.4.0";
        String Adresse_Donnees = "Data Source=C:\\USERS\\YANGM_9VNEY6U\\SOURCE\\REPOS\\ProjetHub\\BD\\Promotions_Groupes.xls";
        String Outils_Concernes = " Extended Properties=Excel 8.0";
        String Specification_Connexion = Fournisseur + ";" + Adresse_Donnees + ";" + Outils_Concernes;
        OleDbConnection Obj_Interop = new OleDbConnection(Specification_Connexion);
        return Obj_Interop;
    }


    protected void LoadData(object sender, EventArgs e)
    {

        SqlConnection sqlConnection = Connexion();
        sqlConnection.Open();

        OleDbConnection excelConnection = ExcelConnexion();
        excelConnection.Open();


        List<string> sheets = new List<string>();
        DataTable dtExcelSchema;

        //Obtenir le Schema du fichier Excel
        dtExcelSchema = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        foreach (DataRow ds in dtExcelSchema.Rows)
        {
            string s = ds["TABLE_NAME"].ToString();
            sheets.Add(s);
        }
        DataTable dataGroupes = GetData(sheets[0], excelConnection);

        foreach (DataRow ds in dataGroupes.Rows)
        {
            string query = "INSERT INTO Groupes (idGroupes,nomGroupes) VALUES(@idGroupes,@nomGroupes)";
            SqlCommand command = new SqlCommand(query, sqlConnection);

            command.Parameters.AddWithValue("@idGroupes", ds[0]);
            command.Parameters.AddWithValue("@nomGroupes", ds[1]);
            command.ExecuteNonQuery();
        }

        DataTable dataPromotions = GetData(sheets[1], excelConnection);

        foreach (DataRow ds in dataPromotions.Rows)
        {
            string query = "INSERT INTO Promotions (idPromotions,nomPromotions) VALUES(@idPromotions,@nomPromotions)";
            SqlCommand command = new SqlCommand(query, sqlConnection);

            command.Parameters.AddWithValue("@idPromotions", ds[0]);
            command.Parameters.AddWithValue("@nomPromotions", ds[1]);
            command.ExecuteNonQuery();
        }

        DataEnseignerModules();
        DataProfesseur();
        DataModules();

        LoadDropDowns();

        Button1.Visible = false;

    }

    protected DataTable GetData(string sheetName, OleDbConnection nameConnection)
    {
        OleDbCommand Cmnd_Selection = new OleDbCommand("SELECT * FROM [" + sheetName + "]", nameConnection);

        // Créér un adaptateur pour récupérer les valeurs des cellules Excel
        OleDbDataAdapter Adaptateur = new OleDbDataAdapter
        {

        //transfert des données depuis le fichier Execl vers l'adaptateur
            SelectCommand = Cmnd_Selection
        };

        DataSet Ens_Donnees = new DataSet();

        //remplir le Data Set avec le contenu de l'adaptateur
        Adaptateur.Fill(Ens_Donnees, sheetName);

        //Création d'une datatable
        DataTable data = Ens_Donnees.Tables[sheetName];

        return data;
    }


    //génération du fichier XML
    protected void GeneratePDF(object sender, EventArgs e)
    {
        DataTable ds = LoadEtudiant();
        int i = 0;
        int yPoint = 0;

        string nomEtudiant = null;
        string prenomEtudiant = null;

        PdfDocument pdf = new PdfDocument();
        pdf.Info.Title = "Liste étudiants";
        PdfPage pdfPage = pdf.AddPage();
        XGraphics graph = XGraphics.FromPdfPage(pdfPage);
        XFont font = new XFont("Verdana", 20, XFontStyle.Regular);

        yPoint = yPoint + 100;

       
        for (i = 0; i < ds.Rows.Count; i++)
            {
                nomEtudiant = ds.Rows[i].ItemArray[0].ToString();
                prenomEtudiant = ds.Rows[i].ItemArray[1].ToString();

                graph.DrawString(nomEtudiant, font, XBrushes.Black, new XRect(100, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft);

                graph.DrawString(prenomEtudiant, font, XBrushes.Black, new XRect(380, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft);

                yPoint = yPoint + 40;
            }


        string pdfFilename = "List_PDF.pdf";
        pdf.Save("C:/USERS/YANGM_9VNEY6U/SOURCE/REPOS/ProjetHub/BD/" + pdfFilename);
        //Process.Start(pdfFilename);
    }



    // récupérer les données de la liste
    protected DataTable LoadEtudiant()
    {

        // Récupérer les données pour avoir les etudiants concernés
        String idProf = DropDownProfesseurs.SelectedValue;
        String idModule = DropDownListModules.SelectedValue;
        String idPromo = DropDownListPromotions.SelectedValue;
        String idGroupe = DropDownListGroupes.SelectedValue;


        SqlConnection sqlConnection = Connexion();
        String req = "SELECT e.nomEtudiant, e.prenomEtudiant from etudiants e JOIN appartenirgroupe apg ON e.idEtudiant" +
            "= apg.idEtudiant JOIN appartenirpromotion app ON e.idEtudiant = app.idEtudiant where apg.idGroupe = " + idGroupe +
            " and app.idPromotion = " + idPromo;
        SqlDataAdapter adpt = new SqlDataAdapter(req, sqlConnection);
        DataTable dt = new DataTable();
        adpt.Fill(dt);

        return dt;
    }




        //génération du fichier XML
    protected void GenerateXML(object sender, EventArgs e)
    {

        DataTable dt = LoadEtudiant();

        // Création de fichier XML qui sera initié avec la variable writer de type   XmlWriter
        using (XmlWriter writer = XmlWriter.Create("C:\\USERS\\YANGM_9VNEY6U\\SOURCE\\REPOS\\ProjetHub\\BD\\liste_XML.xml"))
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("PROFESSEUR", DropDownProfesseurs.SelectedItem.Text.ToString());
            writer.WriteStartElement("MODULE", DropDownListModules.SelectedItem.Text.ToString());
            writer.WriteStartElement("PROMOTION", DropDownListPromotions.SelectedItem.Text.ToString());
            writer.WriteStartElement("GROUPE", DropDownListGroupes.SelectedItem.Text.ToString());

            writer.WriteStartElement("ETUDIANTS");
            foreach (DataRow dr in dt.Rows)
            {
                // Affichage des valeurs des champs pour s'assurer 
                // que la capture des données via le DataSet s'est bien passé
                /*  Response.Write
                    (dr[0].ToString() + "   " + dr[1].ToString() + "<br>" +nomProf +nomPromo); */

                writer.WriteStartElement("ETUDIANT");
                writer.WriteElementString("NOM", dr[0].ToString());
                writer.WriteElementString("PRENOM", dr[1].ToString());
                writer.WriteEndElement();

            }

            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

    }

}


[JsonObject(MemberSerialization = MemberSerialization.OptIn)]
public class DataEnseignModul
{
    [JsonProperty(PropertyName = "idModule")]
    public string IdModule { get; set;}

    [JsonProperty(PropertyName = "idProfesseur")]
    public string IdProfesseur { get; set;}
}

[JsonObject(MemberSerialization = MemberSerialization.OptIn)]
public class DataProf
{
    [JsonProperty(PropertyName = "idProf")]
    public string IdProf { get; set; }

    [JsonProperty(PropertyName = "nomProf")]
    public string NomProf { get; set; }

    [JsonProperty(PropertyName = "prenomProf")]
    public string PrenomProf { get; set; }

}

[JsonObject(MemberSerialization = MemberSerialization.OptIn)]
public class DataModul
{
    [JsonProperty(PropertyName = "idModule")]
    public string IdModule { get; set; }

    [JsonProperty(PropertyName = "nomModule")]
    public string NomModule { get; set; }
}