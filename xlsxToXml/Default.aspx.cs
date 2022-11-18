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
                XmlElement creditInfo = doc.CreateElement("CreditInfo");
                doc.InsertBefore(declaire, doc.DocumentElement);
                doc.AppendChild(creditInfo);

                while (i < oItem)
                {
                    id = dt.Rows[i].ItemArray[0].ToString();

                    if (id != id1)
                    {
                        id1 = id;

                        XmlElement header = doc.CreateElement("Header");    //header yuxari qaldirmaq lazimdi

                        XmlElement bankId = doc.CreateElement("BankID");    //header's children
                        XmlElement bankName = doc.CreateElement("BankName");
                        XmlElement reportingDate = doc.CreateElement("ReportingDate");

                        bankId.InnerText = dt.Rows[i].ItemArray[0].ToString();
                        bankName.InnerText = dt.Rows[i].ItemArray[1].ToString();
                        reportingDate.InnerText = dt.Rows[i].ItemArray[2].ToString();

                        header.AppendChild(bankId);
                        header.AppendChild(bankName);
                        header.AppendChild(reportingDate);
                        creditInfo.AppendChild(header);

                        doc.DocumentElement.AppendChild(header);

                        XmlElement credits = doc.CreateElement("Credits");
                        do
                        {
                            #region main 
                            
                            XmlElement credit = doc.CreateElement("Credit");
                            XmlElement borrower = doc.CreateElement("Borrower");

                            XmlElement idd = doc.CreateElement("id");             //tag borrower
                            XmlElement name = doc.CreateElement("name");
                            XmlElement countryCode = doc.CreateElement("CountryCode");
                            XmlElement bankruptcyStatus = doc.CreateElement("BankruptcyStatus");
                            XmlElement dateOfBirth = doc.CreateElement("DateOfBirth");
                            XmlElement placeOfBirth = doc.CreateElement("PlaceOfBirth");
                            XmlElement pinCode = doc.CreateElement("PinCode");

                            idd.InnerText = dt.Rows[i].ItemArray[3].ToString();
                            name.InnerText = dt.Rows[i].ItemArray[4].ToString();
                            countryCode.InnerText = dt.Rows[i].ItemArray[5].ToString();
                            bankruptcyStatus.InnerText = dt.Rows[i].ItemArray[6].ToString();
                            dateOfBirth.InnerText = dt.Rows[i].ItemArray[7].ToString();
                            placeOfBirth.InnerText = dt.Rows[i].ItemArray[8].ToString();
                            pinCode.InnerText = dt.Rows[i].ItemArray[9].ToString();

                            borrower.AppendChild(idd);
                            borrower.AppendChild(name);
                            borrower.AppendChild(countryCode);
                            borrower.AppendChild(bankruptcyStatus);
                            borrower.AppendChild(dateOfBirth);
                            borrower.AppendChild(placeOfBirth);
                            borrower.AppendChild(pinCode);


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
                            XmlElement dueTimeLastContract = doc.CreateElement("DueTimeLastContract");
                            XmlElement lastPaymentDate = doc.CreateElement("LastPaymentDate");
                            XmlElement monthlyPaymentAmount = doc.CreateElement("MonthlyPaymentAmount");
                            XmlElement daysMainSumIsOverdue = doc.CreateElement("DaysMainSumIsOverdue");
                            XmlElement daysInterestIsOverdue = doc.CreateElement("DaysInterestIsOverdue");
                            XmlElement oiaForRepperiod = doc.CreateElement("OiaForRepperiod");
                            XmlElement numberOfProlongs = doc.CreateElement("NumberOfProlongs");
                            XmlElement creditClassCode = doc.CreateElement("CreditClassCode");
                            XmlElement creditStatusCode = doc.CreateElement("CreditStatusCode");

                            accountNo.InnerText = dt.Rows[i].ItemArray[10].ToString();
                            currencyOfCredit.InnerText = dt.Rows[i].ItemArray[11].ToString();
                            creditType.InnerText = dt.Rows[i].ItemArray[12].ToString();
                            initialAmountOfCredit.InnerText = dt.Rows[i].ItemArray[13].ToString();
                            creditLineAmount.InnerText = dt.Rows[i].ItemArray[14].ToString();
                            disoutAmountOfCredit.InnerText = dt.Rows[i].ItemArray[15].ToString();
                            annualInterestRate.InnerText = dt.Rows[i].ItemArray[16].ToString();
                            purposeOfCredit.InnerText = dt.Rows[i].ItemArray[17].ToString();
                            creditPeriodInMonths.InnerText = dt.Rows[i].ItemArray[18].ToString();
                            dateOfGrant.InnerText = dt.Rows[i].ItemArray[19].ToString();
                            dueTimeFirstContract.InnerText = dt.Rows[i].ItemArray[20].ToString();
                            dueTimeLastContract.InnerText = dt.Rows[i].ItemArray[21].ToString();

                            lastPaymentDate.InnerText = dt.Rows[i].ItemArray[22].ToString();
                            monthlyPaymentAmount.InnerText = dt.Rows[i].ItemArray[23].ToString();
                            daysMainSumIsOverdue.InnerText = dt.Rows[i].ItemArray[24].ToString();
                            daysInterestIsOverdue.InnerText = dt.Rows[i].ItemArray[25].ToString();
                            oiaForRepperiod.InnerText = dt.Rows[i].ItemArray[26].ToString();
                            numberOfProlongs.InnerText = dt.Rows[i].ItemArray[27].ToString();
                            creditClassCode.InnerText = dt.Rows[i].ItemArray[28].ToString();
                            creditStatusCode.InnerText = dt.Rows[i].ItemArray[29].ToString();



                            XmlElement collateral = doc.CreateElement("Collateral"); //tag collateral

                            XmlElement collateralTypeCode = doc.CreateElement("CollateralTypeCode");
                            collateralTypeCode.InnerText = dt.Rows[i].ItemArray[30].ToString();

                            collateral.AppendChild(collateralTypeCode);


                            credit.AppendChild(borrower);

                            credit.AppendChild(accountNo);
                            credit.AppendChild(currencyOfCredit);
                            credit.AppendChild(creditType);
                            credit.AppendChild(initialAmountOfCredit);
                            credit.AppendChild(creditLineAmount);
                            credit.AppendChild(disoutAmountOfCredit);
                            credit.AppendChild(annualInterestRate);
                            credit.AppendChild(purposeOfCredit);
                            credit.AppendChild(creditPeriodInMonths);
                            credit.AppendChild(dateOfGrant);
                            credit.AppendChild(dueTimeFirstContract);
                            credit.AppendChild(dueTimeLastContract);
                            credit.AppendChild(lastPaymentDate);
                            credit.AppendChild(monthlyPaymentAmount);
                            credit.AppendChild(daysMainSumIsOverdue);
                            credit.AppendChild(daysInterestIsOverdue);
                            credit.AppendChild(oiaForRepperiod);
                            credit.AppendChild(numberOfProlongs);
                            credit.AppendChild(creditClassCode);
                            credit.AppendChild(creditStatusCode);

                            credit.AppendChild(collateral);

                            #endregion

                            credits.AppendChild(credit);
                            creditInfo.AppendChild(credits);
                            i++;
                        } while (id1 == dt.Rows[i].ItemArray[0].ToString() && i < oItem);
                        doc.DocumentElement.AppendChild(credits);
                    }
                }

                doc.Save(@"C:\Users\Nurlan Eyvazov\Desktop\Output.xml");
                Response.Write("Created");
            }
        }
    }
}