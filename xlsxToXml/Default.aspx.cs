using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

namespace xlsxToXml
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        protected void btncnvrt_Click(object sender, EventArgs e)
        {
            string id, id1 = "";
            string connStr = "";
            int i = 0;
            int oItem = 0;

            if (file1.HasFile)
            {
                string filename = Path.GetFileName(file1.PostedFile.FileName);
                string fileExtension = Path.GetExtension(file1.PostedFile.FileName);
                string filelocation = @"C:\Users\Nurlan Eyvazov\Desktop\" + filename;
                /*if (fileExtension == ".xls" || fileExtension == ".XLS")
                {
                    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filelocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";

                }
                else if (fileExtension == ".xlsx" || fileExtension == ".XLSX")
                {
                    connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @"D:/coding/c#/work/excel_to_xml/numune.xlsx" + @";Extended Properties=""Excel 12.0 Xml;HDR=YES""";
                }*/

                connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+ filelocation + @";Extended Properties=""Excel 12.0 Xml;HDR=YES""";

                OleDbConnection conn = new OleDbConnection(connStr);
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                conn.Open();
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = dtSheet.Rows[0]["TABLE_NAME"].ToString();
                cmd.CommandText = "select * from [" + sheetName + "]";
                da.SelectCommand = cmd;
                da.Fill(dt);
                conn.Close();
                grdExcel.DataSource = dt;
                grdExcel.DataBind();

                DataSet ds = new DataSet();
                ds.Tables.Add(dt);
                oItem = dt.Rows.Count;
                oItem -= 1;


                XmlDocument doc = new XmlDocument();
                XmlDeclaration declaire = doc.CreateXmlDeclaration("1.0", "utf-8", null); //+
                // -----------------------create root-----------------------------  
                //XmlElement rootnode = doc.CreateElement("root"); 
                XmlElement creditInfo = doc.CreateElement("CreditInfo");
                doc.InsertBefore(declaire, doc.DocumentElement);
                doc.AppendChild(creditInfo);

                while (i < oItem)
                {
                    id = dt.Rows[i].ItemArray[0].ToString();

                    if (id != id1)
                    {
                        id1 = id;


                        DateTime d = Convert.ToDateTime(dt.Rows[i].ItemArray[4]);

                        string str = Convert.ToDateTime(d.ToString("dd-MMMM-yyyy")).ToString("MMMM-dd-yyyy");

                        string finaldate = str.Replace(d.Year.ToString(), DateTime.Now.ToString("yyyy"));

                        XmlElement dobEle = doc.CreateElement("DOB");

                        dobEle.SetAttribute("date", finaldate);



                        XmlElement header = doc.CreateElement("Header");    //header yuxari qaldirmaq lazimdi

                        XmlElement bankId = doc.CreateElement("BankID");    //header's children
                        XmlElement bankName = doc.CreateElement("BankName");
                        XmlElement reportingDate = doc.CreateElement("ReportingDate");
                        creditInfo.AppendChild(header);

                        do
                        {
                            //<?xml version="1.0" encoding="UTF-8"?>
                            //credit info
                            //header
                            //credits
                            //credit
                            //group -> borrower
                            //others



                            #region main 
                            

                            XmlElement credits = doc.CreateElement("Credits");

                            XmlElement idd = doc.CreateElement("id");             //tag borrower
                            XmlElement name = doc.CreateElement("name");
                            XmlElement countryCode = doc.CreateElement("CountryCode");
                            XmlElement bankruptcyStatus = doc.CreateElement("BankruptcyStatus");
                            XmlElement dateOfBirth = doc.CreateElement("DateOfBirth");
                            XmlElement placeOfBirth = doc.CreateElement("PlaceOfBirth");
                            XmlElement pinCode = doc.CreateElement("PinCode");


                            XmlElement accountNo = doc.CreateElement("AccountNo");     //
                            XmlElement currencyOfCredit = doc.CreateElement("CurrencyOfCredit");
                            XmlElement creditType = doc.CreateElement("CreditType");
                            XmlElement initialAmountOfCredit = doc.CreateElement("InitialAmountOfCredit");
                            XmlElement creditLineAmount = doc.CreateElement("creditLineAmount");
                            XmlElement disoutAmountOfCredit = doc.CreateElement("DisoutAmountOfCredit");
                            XmlElement annualInterestRate = doc.CreateElement("AnnualInterestRate");
                            XmlElement purposeOfCredit = doc.CreateElement("PurposeOfCredit");
                            XmlElement creditPeriodInMonths = doc.CreateElement("CreditPeriodInMonths");
                            XmlElement dateOfGrant = doc.CreateElement("DateOfGrant");
                            XmlElement dueTimeFirstContract = doc.CreateElement("DueTimeFirstContract");
                            XmlElement DueTimeLastContract = doc.CreateElement("DueTimeLastContract");
                            XmlElement LastPaymentDate = doc.CreateElement("LastPaymentDate");
                            XmlElement MonthlyPaymentAmount = doc.CreateElement("MonthlyPaymentAmount");
                            XmlElement daysMainSumIsOverdue = doc.CreateElement("DaysMainSumIsOverdue");
                            XmlElement daysInterestIsOverdue = doc.CreateElement("DaysInterestIsOverdue");
                            XmlElement oiaForRepperiod = doc.CreateElement("OiaForRepperiod");
                            XmlElement numberOfProlongs = doc.CreateElement("NumberOfProlongs");
                            XmlElement creditClassCode = doc.CreateElement("CreditClassCode");
                            XmlElement creditStatusCode = doc.CreateElement("CreditStatusCode");
                            XmlElement collateralTypeCode = doc.CreateElement("CollateralTypeCode");

                            #endregion



                            XmlElement emp = doc.CreateElement("EmpDetails");
                            
                            XmlElement name1 = doc.CreateElement("Name");
                            emp.AppendChild(name);
                            XmlElement desig = doc.CreateElement("Designation");
                            emp.AppendChild(desig);
                            XmlElement dept = doc.CreateElement("Dept");
                            emp.AppendChild(dept);
                            XmlElement loc = doc.CreateElement("Location");
                            emp.AppendChild(loc);

                            name1.InnerText = dt.Rows[i].ItemArray[1].ToString();
                            desig.InnerText = dt.Rows[i].ItemArray[2].ToString();
                            dept.InnerText = desig.InnerText;
                            loc.InnerText = dt.Rows[i].ItemArray[3].ToString();
                            

                            dobEle.AppendChild(emp);
                            i++;
                        } while (id1 == dt.Rows[i].ItemArray[0].ToString() && i < oItem);


                        doc.DocumentElement.AppendChild(dobEle);

                    }
                }

                doc.Save(@"C:\Users\Nurlan Eyvazov\Desktop\Output.xml");
                Response.Write("Created");
            }
        }
    }
}