using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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

        protected void btncnvrt_Click(object sender, EventArgs  e)
        {
            string id, id1 = "";
            string connStr = "";
            int i = 0;
            int oItem = 0;

            if (file1.HasFile)
            {
                string filename = Path.GetFileName(file1.PostedFile.FileName);
                string fileExtension = Path.GetExtension(file1.PostedFile.FileName);
                //string fullPath = Path.GetFullPath(file1.PostedFile.FileName);
                //string fullPath2 = file1.PostedFile.FileName;
                string filelocation = @"D:\coding\c#\work\excel_to_xml\" + filename;
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

                        if (dt.Rows[i].ItemArray[0].ToString() != "")
                        {
                            bankId.InnerText = dt.Rows[i].ItemArray[0].ToString();
                        }
                        if (dt.Rows[i].ItemArray[1].ToString() != "")
                        {
                            bankName.InnerText = dt.Rows[i].ItemArray[1].ToString();
                        }
                        if (dt.Rows[i].ItemArray[2].ToString() != "")
                        {
                            reportingDate.InnerText = dt.Rows[i].ItemArray[2].ToString();
                        }

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

                            if (dt.Rows[i].ItemArray[3].ToString() != "")
                            {
                                idd.InnerText = dt.Rows[i].ItemArray[3].ToString();
                            }
                            if (dt.Rows[i].ItemArray[4].ToString() != "")
                            {
                                name.InnerText = dt.Rows[i].ItemArray[4].ToString();
                            }
                            if (dt.Rows[i].ItemArray[5].ToString() != "")
                            {
                                countryCode.InnerText = dt.Rows[i].ItemArray[5].ToString();
                            }
                            if (dt.Rows[i].ItemArray[6].ToString() != "")
                            {
                                bankruptcyStatus.InnerText = dt.Rows[i].ItemArray[6].ToString();
                            }
                            if (dt.Rows[i].ItemArray[7].ToString() != "")
                            {
                                dateOfBirth.InnerText = dt.Rows[i].ItemArray[7].ToString();
                            }
                            if (dt.Rows[i].ItemArray[8].ToString() != "")
                            {
                                placeOfBirth.InnerText = dt.Rows[i].ItemArray[8].ToString();
                            }
                            if (dt.Rows[i].ItemArray[9].ToString() != "")
                            {
                                pinCode.InnerText = dt.Rows[i].ItemArray[9].ToString();
                            }

                            borrower.AppendChild(idd);
                            borrower.AppendChild(name);
                            borrower.AppendChild(countryCode);
                            borrower.AppendChild(bankruptcyStatus);
                            borrower.AppendChild(dateOfBirth);
                            borrower.AppendChild(placeOfBirth);
                            borrower.AppendChild(pinCode);

                            

                            XmlElement guarantees = doc.CreateElement("Guarantees");
                            XmlElement guaranteeG = doc.CreateElement("Guarantee");

                            XmlElement iddG = doc.CreateElement("id");             //tag guarantee
                            XmlElement nameG = doc.CreateElement("name");
                            XmlElement countryCodeG = doc.CreateElement("CountryCode");
                            XmlElement dateOfBirthG = doc.CreateElement("DateOfBirth");
                            XmlElement placeOfBirthG = doc.CreateElement("PlaceOfBirth");
                            XmlElement pinCodeG = doc.CreateElement("PinCode");


                            if (dt.Rows[i].ItemArray[10].ToString() != "")
                            {
                                iddG.InnerText = dt.Rows[i].ItemArray[10].ToString().Trim();
                            }
                            if (dt.Rows[i].ItemArray[11].ToString() != "")
                            {
                                nameG.InnerText = dt.Rows[i].ItemArray[11].ToString().Trim();
                            }
                            if (dt.Rows[i].ItemArray[12].ToString() != "")
                            {
                                countryCodeG.InnerText = dt.Rows[i].ItemArray[12].ToString().Trim();
                            }
                            if (dt.Rows[i].ItemArray[13].ToString() != "")
                            {
                                dateOfBirthG.InnerText = dt.Rows[i].ItemArray[13].ToString().Trim();
                            }

                            if (dt.Rows[i].ItemArray[14].ToString() != "")
                            {
                                placeOfBirthG.InnerText = dt.Rows[i].ItemArray[14].ToString().Trim();
                            }
                            if (dt.Rows[i].ItemArray[15].ToString() != "")
                            {
                                pinCodeG.InnerText = dt.Rows[i].ItemArray[15].ToString().Trim();
                            }

                            guaranteeG.AppendChild(iddG);
                            guaranteeG.AppendChild(nameG);
                            guaranteeG.AppendChild(countryCodeG);
                            guaranteeG.AppendChild(dateOfBirthG);
                            guaranteeG.AppendChild(placeOfBirthG);
                            guaranteeG.AppendChild(pinCodeG);


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

                            if (dt.Rows[i].ItemArray[16].ToString() != "")
                            {
                                accountNo.InnerText = dt.Rows[i].ItemArray[16].ToString();
                            }
                            if (dt.Rows[i].ItemArray[17].ToString() != "")
                            {
                                currencyOfCredit.InnerText = dt.Rows[i].ItemArray[17].ToString();
                            }
                            if (dt.Rows[i].ItemArray[18].ToString() != "")
                            {
                                creditType.InnerText = dt.Rows[i].ItemArray[18].ToString(); //
                            }
                            if (dt.Rows[i].ItemArray[19].ToString() != "")
                            {
                                initialAmountOfCredit.InnerText = dt.Rows[i].ItemArray[19].ToString();
                            }
                            if (dt.Rows[i].ItemArray[20].ToString() != "")
                            {
                                creditLineAmount.InnerText = dt.Rows[i].ItemArray[20].ToString();
                            }
                            if (dt.Rows[i].ItemArray[21].ToString() != "")
                            {
                                disoutAmountOfCredit.InnerText = dt.Rows[i].ItemArray[21].ToString();
                            }
                            if (dt.Rows[i].ItemArray[22].ToString() != "")
                            {
                                annualInterestRate.InnerText = dt.Rows[i].ItemArray[22].ToString();
                            }

                            if (dt.Rows[i].ItemArray[23].ToString() != "")
                            {
                                purposeOfCredit.InnerText = dt.Rows[i].ItemArray[23].ToString(); //
                            }
                            if (dt.Rows[i].ItemArray[24].ToString() != "")
                            {
                                creditPeriodInMonths.InnerText = dt.Rows[i].ItemArray[24].ToString();
                            }
                            if (dt.Rows[i].ItemArray[25].ToString() != "")
                            {
                                dateOfGrant.InnerText = dt.Rows[i].ItemArray[25].ToString();
                            }
                            if (dt.Rows[i].ItemArray[26].ToString() != "")
                            {
                                dueTimeFirstContract.InnerText = dt.Rows[i].ItemArray[26].ToString();
                            }
                            if (dt.Rows[i].ItemArray[27].ToString() != "")
                            {
                                dueTimeLastContract.InnerText = dt.Rows[i].ItemArray[27].ToString();
                            }

                            if (dt.Rows[i].ItemArray[28].ToString() != "")
                            {
                                lastPaymentDate.InnerText = dt.Rows[i].ItemArray[28].ToString();
                            }
                            if (dt.Rows[i].ItemArray[29].ToString() != "")
                            {
                                monthlyPaymentAmount.InnerText = dt.Rows[i].ItemArray[29].ToString();
                            }
                            if (dt.Rows[i].ItemArray[30].ToString() != "")
                            {
                                daysMainSumIsOverdue.InnerText = dt.Rows[i].ItemArray[30].ToString();
                            }
                            if (dt.Rows[i].ItemArray[31].ToString() != "")
                            {
                                daysInterestIsOverdue.InnerText = dt.Rows[i].ItemArray[31].ToString();
                            }
                            if (dt.Rows[i].ItemArray[32].ToString() != "")
                            {
                                oiaForRepperiod.InnerText = dt.Rows[i].ItemArray[32].ToString();
                            }
                            if (dt.Rows[i].ItemArray[33].ToString() != "")
                            {
                                numberOfProlongs.InnerText = dt.Rows[i].ItemArray[33].ToString();
                            }
                            if (dt.Rows[i].ItemArray[34].ToString() != "")
                            {
                                creditClassCode.InnerText = dt.Rows[i].ItemArray[34].ToString(); //
                            }
                            if (dt.Rows[i].ItemArray[35].ToString() != "")
                            {
                                creditStatusCode.InnerText = dt.Rows[i].ItemArray[35].ToString(); //
                            }



                            XmlElement collateral = doc.CreateElement("Collateral"); //tag collateral

                            XmlElement collateralTypeCode = doc.CreateElement("CollateralTypeCode");
                            XmlElement anyInfoToDisting = doc.CreateElement("AnyInfoToDisting");
                            XmlElement marketValue = doc.CreateElement("MarketValue");
                            XmlElement registryNo = doc.CreateElement("RegistryNo");
                            XmlElement registryDate = doc.CreateElement("RegistryDate");
                            XmlElement registryAgency = doc.CreateElement("RegistryAgency");

                            if (dt.Rows[i].ItemArray[36].ToString() != "")
                            {
                                collateralTypeCode.InnerText = dt.Rows[i].ItemArray[36].ToString(); //
                            }
                            if (dt.Rows[i].ItemArray[37].ToString() != "")
                            {
                                anyInfoToDisting.InnerText = dt.Rows[i].ItemArray[37].ToString(); //
                            }
                            if (dt.Rows[i].ItemArray[38].ToString() != "")
                            {
                                marketValue.InnerText = dt.Rows[i].ItemArray[38].ToString(); //
                            }
                            if (dt.Rows[i].ItemArray[39].ToString() != "")
                            {
                                registryNo.InnerText = dt.Rows[i].ItemArray[39].ToString(); //
                            }
                            /*if (registryNo == null)
                            {

                            }*/



                            if (dt.Rows[i].ItemArray[40].ToString() != "")
                            {
                                registryDate.InnerText = dt.Rows[i].ItemArray[40].ToString();
                            }
                            if (dt.Rows[i].ItemArray[41].ToString() != "")
                            {
                                registryAgency.InnerText = dt.Rows[i].ItemArray[41].ToString(); //
                            }

                            collateral.AppendChild(collateralTypeCode);
                            collateral.AppendChild(anyInfoToDisting);
                            collateral.AppendChild(marketValue);
                            collateral.AppendChild(registryNo);
                            collateral.AppendChild(registryDate);
                            collateral.AppendChild(registryAgency);

                            //guarantees.AppendChild(guaranteeG); //guarantee
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

                            borrower.AppendChild(guarantees);
                            guarantees.AppendChild(guaranteeG);
                            credits.AppendChild(credit);
                            creditInfo.AppendChild(credits);
                            i++;
                        } while (id1 == dt.Rows[i].ItemArray[0].ToString() && i < oItem);
                        doc.DocumentElement.AppendChild(credits);
                    }
                }

                doc.Save($"C:\\Users\\nurlan.eyvazov\\Desktop\\{filename}.xml");
                Response.Write("Created");
            }
        }

        protected void btncnvrt_Click_without_Guarantees(object sender, EventArgs e)
        {
            string id, id1 = "";
            string connStr = "";
            int i = 0;
            int oItem = 0;

            if (file1.HasFile)
            {
                #region connection 
                string filename = Path.GetFileName(file1.PostedFile.FileName);
                string fileExtension = Path.GetExtension(file1.PostedFile.FileName);
                //string fullPath = Path.GetFullPath(file1.PostedFile.FileName);
                //string fullPath2 = file1.PostedFile.FileName;
                //string filelocation = @"D:\coding\c#\work\1WORK\excel_to_xml\" + filename;

                string filelocation = Server.MapPath("~/Uploaded_Files/") + filename;
                /*if (fileExtension == ".xls" || fileExtension == ".XLS")
                {
                    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filelocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";

                }
                else if (fileExtension == ".xlsx" || fileExtension == ".XLSX")
                {
                    connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @"D:/coding/c#/work/excel_to_xml/numune.xlsx" + @";Extended Properties=""Excel 12.0 Xml;HDR=YES""";
                }*/

                connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filelocation + @";Extended Properties=""Excel 12.0 Xml;HDR=YES""";

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

                #endregion

                DataSet ds = new DataSet();
                ds.Tables.Add(dt);
                oItem = dt.Rows.Count;
                oItem -= 1;

                #region xml
                XmlDocument doc = new XmlDocument();
                XmlDeclaration declaire = doc.CreateXmlDeclaration("1.0", "utf-8", null); //+
                // -----------------------create root-----------------------------  
                XmlElement creditInfo = doc.CreateElement("CreditInfo");
                doc.InsertBefore(declaire, doc.DocumentElement);
                doc.AppendChild(creditInfo);
                #endregion

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

                        if (dt.Rows[i].ItemArray[0].ToString() != "")
                        {
                            bankId.InnerText = dt.Rows[i].ItemArray[0].ToString();
                        }
                        else
                        {
                            bankId.InnerText = "EMPTY_VALUE";
                        }

                        if (dt.Rows[i].ItemArray[1].ToString() != "")
                        {
                            bankName.InnerText = dt.Rows[i].ItemArray[1].ToString();
                        }
                        else
                        {
                            bankName.InnerText = "EMPTY_VALUE";
                        }

                        if (dt.Rows[i].ItemArray[2].ToString() != "")
                        {
                            reportingDate.InnerText = dt.Rows[i].ItemArray[2].ToString();
                        }
                        else
                        {
                            reportingDate.InnerText = "EMPTY_VALUE";
                        }

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

                            if (dt.Rows[i].ItemArray[3].ToString() != "")
                            {
                                idd.InnerText = dt.Rows[i].ItemArray[3].ToString();
                            }
                            else
                            {
                                idd.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[4].ToString() != "")
                            {
                                name.InnerText = dt.Rows[i].ItemArray[4].ToString();
                            }
                            else
                            {
                                name.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[5].ToString() != "")
                            {
                                countryCode.InnerText = dt.Rows[i].ItemArray[5].ToString();
                            }
                            else
                            {
                                countryCode.InnerText= "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[6].ToString() != "")
                            {
                                bankruptcyStatus.InnerText = dt.Rows[i].ItemArray[6].ToString();
                            }
                            else
                            {
                                bankruptcyStatus.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[7].ToString() != "")
                            {
                                dateOfBirth.InnerText = dt.Rows[i].ItemArray[7].ToString();
                            }
                            else
                            {
                                dateOfBirth.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[8].ToString() != "")
                            {
                                placeOfBirth.InnerText = dt.Rows[i].ItemArray[8].ToString();
                            }
                            else
                            {
                                placeOfBirth.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[9].ToString() != "")
                            {
                                pinCode.InnerText = dt.Rows[i].ItemArray[9].ToString();
                            }
                            else
                            {
                                pinCode.InnerText = "EMPTY_VALUE";
                            }

                            borrower.AppendChild(idd);
                            borrower.AppendChild(name);
                            borrower.AppendChild(countryCode);
                            borrower.AppendChild(bankruptcyStatus);
                            borrower.AppendChild(dateOfBirth);
                            borrower.AppendChild(placeOfBirth);
                            borrower.AppendChild(pinCode);



                            /*XmlElement guarantees = doc.CreateElement("Guarantees");
                            XmlElement guaranteeG = doc.CreateElement("Guarantee");

                            XmlElement iddG = doc.CreateElement("id");             //tag guarantee
                            XmlElement nameG = doc.CreateElement("name");
                            XmlElement countryCodeG = doc.CreateElement("CountryCode");
                            XmlElement dateOfBirthG = doc.CreateElement("DateOfBirth");
                            XmlElement placeOfBirthG = doc.CreateElement("PlaceOfBirth");
                            XmlElement pinCodeG = doc.CreateElement("PinCode");


                            if (dt.Rows[i].ItemArray[10].ToString() != "")
                            {
                                iddG.InnerText = dt.Rows[i].ItemArray[10].ToString().Trim();
                            }
                            if (dt.Rows[i].ItemArray[11].ToString() != "")
                            {
                                nameG.InnerText = dt.Rows[i].ItemArray[11].ToString().Trim();
                            }
                            if (dt.Rows[i].ItemArray[12].ToString() != "")
                            {
                                countryCodeG.InnerText = dt.Rows[i].ItemArray[12].ToString().Trim();
                            }
                            if (dt.Rows[i].ItemArray[13].ToString() != "")
                            {
                                dateOfBirthG.InnerText = dt.Rows[i].ItemArray[13].ToString().Trim();
                            }

                            if (dt.Rows[i].ItemArray[14].ToString() != "")
                            {
                                placeOfBirthG.InnerText = dt.Rows[i].ItemArray[14].ToString().Trim();
                            }
                            if (dt.Rows[i].ItemArray[15].ToString() != "")
                            {
                                pinCodeG.InnerText = dt.Rows[i].ItemArray[15].ToString().Trim();
                            }

                            guaranteeG.AppendChild(iddG);
                            guaranteeG.AppendChild(nameG);
                            guaranteeG.AppendChild(countryCodeG);
                            guaranteeG.AppendChild(dateOfBirthG);
                            guaranteeG.AppendChild(placeOfBirthG);
                            guaranteeG.AppendChild(pinCodeG);*/


                            XmlElement accountNo = doc.CreateElement("AccountNo");     // 
                            XmlElement currencyOfCredit = doc.CreateElement("CurrencyOfCredit");
                            XmlElement creditType = doc.CreateElement("CreditType");
                            XmlElement initialAmountOfCredit = doc.CreateElement("InitialAmountOfCredit");
                            XmlElement creditLineAmount = doc.CreateElement("CreditLineAmount");
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
                            XmlElement numberOfProlongs = doc.CreateElement("NumberOfProlongs"); //0
                            XmlElement creditClassCode = doc.CreateElement("CreditClassCode"); //0
                            XmlElement creditStatusCode = doc.CreateElement("CreditStatusCode"); //0

                            if (dt.Rows[i].ItemArray[10].ToString() != "")
                            {
                                accountNo.InnerText = dt.Rows[i].ItemArray[10].ToString();
                            }
                            else
                            {
                                accountNo.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[11].ToString() != "")
                            {
                                currencyOfCredit.InnerText = dt.Rows[i].ItemArray[11].ToString();
                            }
                            else
                            {
                                currencyOfCredit.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[12].ToString() != "")
                            {
                                creditType.InnerText = dt.Rows[i].ItemArray[12].ToString(); //
                            }
                            else
                            {
                                creditType.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[13].ToString() != "")
                            {
                                initialAmountOfCredit.InnerText = dt.Rows[i].ItemArray[13].ToString();
                            }
                            else
                            {
                                initialAmountOfCredit.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[14].ToString() != "")
                            {
                                creditLineAmount.InnerText = dt.Rows[i].ItemArray[14].ToString();
                            }
                            else
                            {
                                creditLineAmount.InnerText = "EMPTY_VALUE";

                            }

                            if (dt.Rows[i].ItemArray[15].ToString() != "")
                            {
                                disoutAmountOfCredit.InnerText = dt.Rows[i].ItemArray[15].ToString();
                            }
                            else
                            {
                                disoutAmountOfCredit.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[16].ToString() != "")
                            {
                                annualInterestRate.InnerText = dt.Rows[i].ItemArray[16].ToString();
                            }
                            else
                            {
                                annualInterestRate.InnerText = "0";
                            }

                            if (dt.Rows[i].ItemArray[17].ToString() != "")
                            {
                                purposeOfCredit.InnerText = dt.Rows[i].ItemArray[17].ToString(); //
                            }
                            else
                            {
                                purposeOfCredit.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[18].ToString() != "")
                            {
                                creditPeriodInMonths.InnerText = dt.Rows[i].ItemArray[18].ToString();
                            }
                            else
                            {
                                creditPeriodInMonths.InnerText = "EMPTY_VALUE";

                            }

                            if (dt.Rows[i].ItemArray[19].ToString() != "")
                            {
                                dateOfGrant.InnerText = dt.Rows[i].ItemArray[19].ToString();
                            }
                            else
                            {
                                dateOfGrant.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[20].ToString() != "")
                            {
                                dueTimeFirstContract.InnerText = dt.Rows[i].ItemArray[20].ToString();
                            }
                            else
                            {
                                dueTimeFirstContract.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[21].ToString() != "")
                            {
                                dueTimeLastContract.InnerText = dt.Rows[i].ItemArray[21].ToString();
                            }
                            else
                            {
                                dueTimeLastContract.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[22].ToString() != "")
                            {
                                lastPaymentDate.InnerText = dt.Rows[i].ItemArray[22].ToString();
                            }
                            else
                            {
                                lastPaymentDate.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[23].ToString() != "")
                            {
                                monthlyPaymentAmount.InnerText = dt.Rows[i].ItemArray[23].ToString();
                            }
                            else
                            {
                                monthlyPaymentAmount.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[24].ToString() != "")
                            {
                                daysMainSumIsOverdue.InnerText = dt.Rows[i].ItemArray[24].ToString();
                            }
                            else
                            {
                                daysMainSumIsOverdue.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[25].ToString() != "")
                            {
                                daysInterestIsOverdue.InnerText = dt.Rows[i].ItemArray[25].ToString();
                            }
                            else
                            {
                                daysInterestIsOverdue.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[26].ToString() != "")
                            {
                                oiaForRepperiod.InnerText = dt.Rows[i].ItemArray[26].ToString();
                            }
                            else
                            {
                                oiaForRepperiod.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[27].ToString() != "")
                            {
                                numberOfProlongs.InnerText = dt.Rows[i].ItemArray[27].ToString();
                            }
                            else
                            {
                                numberOfProlongs.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[28].ToString() != "")
                            {
                                creditClassCode.InnerText = dt.Rows[i].ItemArray[28].ToString(); //
                            }
                            else
                            {
                                creditClassCode.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[29].ToString() != "")
                            {
                                creditStatusCode.InnerText = dt.Rows[i].ItemArray[29].ToString(); //
                            }
                            else
                            {
                                creditStatusCode.InnerText = "EMPTY_VALUE";
                            }


                            XmlElement collateral = doc.CreateElement("Collateral"); //tag collateral

                            XmlElement collateralTypeCode = doc.CreateElement("CollateralTypeCode"); //006
                            XmlElement anyInfoToDisting = doc.CreateElement("AnyInfoToDisting"); //003
                            XmlElement marketValue = doc.CreateElement("MarketValue"); //007
                            XmlElement registryNo = doc.CreateElement("RegistryNo");
                            XmlElement registryDate = doc.CreateElement("RegistryDate");
                            XmlElement registryAgency = doc.CreateElement("RegistryAgency");

                            if (dt.Rows[i].ItemArray[30].ToString() != "")
                            {
                                collateralTypeCode.InnerText = dt.Rows[i].ItemArray[30].ToString(); //
                            }
                            else
                            {
                                collateralTypeCode.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[31].ToString() != "")
                            {
                                anyInfoToDisting.InnerText = dt.Rows[i].ItemArray[31].ToString(); //
                            }
                            else
                            {
                                anyInfoToDisting.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[32].ToString() != "")
                            {
                                marketValue.InnerText = dt.Rows[i].ItemArray[32].ToString(); //
                            }
                            else
                            {
                                marketValue.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[33].ToString() != "")
                            {
                                registryNo.InnerText = dt.Rows[i].ItemArray[33].ToString(); //
                            }
                            else
                            {
                                registryNo.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[34].ToString() != "")
                            {
                                registryDate.InnerText = dt.Rows[i].ItemArray[34].ToString();
                            }
                            else
                            {
                                registryDate.InnerText = "EMPTY_VALUE";
                            }

                            if (dt.Rows[i].ItemArray[35].ToString() != "")
                            {
                                registryAgency.InnerText = dt.Rows[i].ItemArray[35].ToString(); //
                            }
                            else
                            {
                                registryAgency.InnerText = "EMPTY_VALUE";
                            }


                            collateral.AppendChild(collateralTypeCode);
                            collateral.AppendChild(anyInfoToDisting);
                            collateral.AppendChild(marketValue);
                            collateral.AppendChild(registryNo);
                            collateral.AppendChild(registryDate);
                            collateral.AppendChild(registryAgency);

                            //guarantees.AppendChild(guaranteeG); //guarantee
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

                            //borrower.AppendChild(guarantees);
                            //guarantees.AppendChild(guaranteeG);
                            credits.AppendChild(credit);
                            creditInfo.AppendChild(credits);
                            i++;
                        } while (id1 == dt.Rows[i].ItemArray[0].ToString() && i < oItem);
                        doc.DocumentElement.AppendChild(credits);
                    }
                }

                doc.Save($"C:\\Users\\nurlan.eyvazov\\Desktop\\{filename}.xml");
                Response.Write("Created");
            }
        }

        /*void UploadFile()
        {
            try
            {
                if (fuFileUpload.HasFile)
                {
                    //Check File Extension
                    string fileExtension = System.IO.Path.GetExtension(fuFileUpload.FileName);

                    if (fileExtension.ToLower() != ".xls" && fileExtension.ToLower() != ".xlsx")
                    {
                        lblMessage.Text = "Files with .xls and .xlsx only allowed.";
                        lblMessage.ForeColor = System.Drawing.Color.Red;
                    }
                    else
                    {
                        string folderPath = Server.MapPath("~/UploadedFiles/");

                        //Check whether Directory is available or not.
                        if (!Directory.Exists(folderPath))
                        {
                            //if directory is not available, then create a folder.
                            Directory.CreateDirectory(folderPath);
                        }

                        //save the file to folder
                        fuFileUpload.SaveAs(folderPath + Path.GetFileName(fuFileUpload.FileName));

                        //dispaly message after uploaded
                        lblMessage.Text = "File " + Path.GetFileName(fuFileUpload.FileName) + " uploaded successfully...!";
                    }
                }
                else
                {
                    lblMessage.Text = "Please select a file to upload.";
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;
                lblMessage.ForeColor = System.Drawing.Color.Red;
            }
        }*/
    }
}