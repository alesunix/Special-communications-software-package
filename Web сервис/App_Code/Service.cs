using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Services;

[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// Чтобы разрешить вызывать веб-службу из скрипта с помощью ASP.NET AJAX, раскомментируйте следующую строку.
// [System.Web.Script.Services.ScriptService]

public class Service : System.Web.Services.WebService
{
    SqlConnection con = new SqlConnection("Data Source=192.168.0.3,1433;Network Library=DBMSSOCN;Initial Catalog=ccsbase;User ID=Lan;Password=Samsung0;");
    SqlDataAdapter da;
    DataSet ds;

    public Service()
    {
        //Раскомментируйте следующую строку в случае использования сконструированных компонентов 
        //InitializeComponent(); 
    }

    [WebMethod]
    public DataSet Search(string N_zakaza)
    {
        //string q = "SELECT N_zakaza, oblast, punkt, familia, summ, status, prichina, client FROM [Table_1] WHERE N_zakaza LIKE '%" + N_zakaza + "%'";
        //da = new SqlDataAdapter(q, con);
        //ds = new DataSet();
        //da.Fill(ds);
        //return ds;

        con.Open();//открыть соединение
        SqlCommand cmd = new SqlCommand("SELECT N_zakaza, oblast, punkt, familia, summ, status, prichina, client FROM [Table_1]" +
            "WHERE N_zakaza = @N_zakaza", con);
        cmd.Parameters.AddWithValue("@N_zakaza", N_zakaza);
        cmd.ExecuteNonQuery();
        da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
        ds = new DataSet();
        da.Fill(ds);
        return ds;

        //con = new SqlConnection("Data Source=192.168.0.3,1433;Network Library=DBMSSOCN;Initial Catalog=ccsbase;User ID=Lan;Password=Samsung0;");
        //string q = "SELECT * FROM [Table_1] WHERE N_zakaza LIKE '%" + N_zakaza + "%' AND client = 'ИП Атантаева Н.Т.'";
        //da = new SqlDataAdapter(q, con);
        //ds = new DataSet();
        //da.Fill(ds);
        //return ds;
    }
    [WebMethod]
    public DataSet Update(string Prichina, string Status)//test
    {
        con.Open();//открыть соединение
        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET prichina = @prichina, status= @status WHERE N_zakaza = @N_zakaza", con);
        cmd.Parameters.AddWithValue("@N_zakaza", "KG109981");
        cmd.Parameters.AddWithValue("@prichina", Prichina);
        cmd.Parameters.AddWithValue("@status", Status);
        cmd.ExecuteNonQuery();
        da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
        ds = new DataSet();
        da.Fill(ds);
        return ds;
    }

    [WebMethod]
    public DataSet BETWEEN(DateTime StartDate, DateTime EndDate)
    {
        con.Open();//открыть соединение
        SqlCommand cmd = new SqlCommand("SELECT N_zakaza, oblast, punkt, familia, summ, status, prichina, client FROM [Table_1]" +
            "WHERE (data_zapisi BETWEEN @StartDate AND @EndDate AND client = @client)", con);
        cmd.Parameters.AddWithValue("@StartDate", StartDate);
        cmd.Parameters.AddWithValue("@EndDate", EndDate);
        cmd.Parameters.AddWithValue("@client", "'ИП Атантаева Н.Т.'");
        cmd.ExecuteNonQuery();
        da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
        ds = new DataSet();
        da.Fill(ds);
        return ds;
    }
    [WebMethod]
    public DateTime[] GetDateTimes()
    {
        return new DateTime[] {
        DateTime.Now,
        DateTime.Now.ToLocalTime (),
        DateTime.Now.ToUniversalTime ()};
    }


}