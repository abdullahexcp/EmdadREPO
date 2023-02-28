using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using JR.AX;
using System.Data.SqlClient;
using RestSharp;
using System.Net;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Bibliography;


public class masterDetails
{
    public string journalID { get; set; }
    public string journalName { get; set; }
    public string documentNum { get; set; }
    public DateTime TransDate
    {
        set { TransDate = value; }
        get
        {
            return TransDate.ToString("MM/dd/yyyy");
        }
    }
    public string Description { set; get; }
    public int ledgerJournalACType { set; get; }
    public string account { set; get; }
    public string AccountDim { set; get; }
    public string amountDebit { set; get; }
    public string AmountCredit { set; get; }
    public string account { set; get; }
    public int OffsetaccTyp { set; get; }
    public int offSetAccountCode { set; get; }
    public string OffsetDim { set; get; }
    public string PostingProfile { set; get; }
    public string bearaerObj { set; get; }
}


public partial class AX_UploadDetailsJournal : GlobalPage
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void btnUpload_Click(object sender, EventArgs e)
    {
        AXObject bearaerObj = new AXObject();

        IRestResponse errorMessage = null;


        string[] allFormats ={"yyyy/MM/dd","yyyy/M/d",
                "dd/MM/yyyy","d/M/yyyy",
                "dd/M/yyyy","d/MM/yyyy","yyyy-MM-dd",
                "yyyy-M-d","dd-MM-yyyy","d-M-yyyy",
                "dd-M-yyyy","d-MM-yyyy","yyyy MM dd",
                "yyyy M d","dd MM yyyy","d M yyyy",
                "dd M yyyy","d MM yyyy"};

        DataTable dtsource = ConvertToExcel.ExcelImport.ImportAnyExcel(FileUpload1.FileName, FileUpload1.PostedFile, true).Tables[0];
        DataTable dtError = dtsource.Clone();
        dtError.Columns.Add("Error");
        //DataTable defaultJournalValues_dt = ConvertToExcel.ExcelImport.ImportExcelXLS(Server.MapPath("~/Templates/DefaultJournalDetails.xlsx"), "DefaultJournalDetails.xlsx", true).Tables[0];
        Dictionary<String, String> defaultJournalValues_dict = new Dictionary<string, string>() {
            { "CostCenter", txtBox_defCostCenter.Text  },
            { "Department", txtBox_defDept.Text  },
            { "Activity", txtBox_defActivity.Text  },
        };

        if (dtsource.Rows.Count > 1)
        {
            // AXCRMNetBusConnector connector = new AXCRMNetBusConnector(Session["Password"].ToString(), Session["Login"].ToString());
            string usernamesession = "";
            try
            {
                usernamesession = Session["Login"].ToString();
            }
            catch (Exception)
            {


            }


            // AXCRMNetBusConnector connector = new AXCRMNetBusConnector();
            //Create LedgerHeader

            string ledgerName = txtJournal.Text;
            string ledgerDesc = txtDescription.Text;
            string journalID = string.Empty;
            try
            {
                journalID = txtJournalNum.Text;

                var tokenResult = AXFinanceOperations.GetToken();

                if (tokenResult.IsSuccessful && tokenResult.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    bearaerObj = JsonConvert.DeserializeObject<AXObject>(tokenResult.Content);

                }

                List<DataTable> DtjournalList = dtsource.AsEnumerable()
                .GroupBy(row => new
                {
                    MapAs = row.Field<string>("MapAS"),
                    GroupBy = row.Field<string>("GroupBy")
                })
                .Select(g => g.CopyToDataTable())
                .ToList();

                foreach (var dt in DtjournalList)
                {


                    #region CreateJournalHeader


                    if (txtJournalNum.Text == "")
                    {
                        //journalID = connector.CreateLedgerHeader(ledgerName, ledgerDesc+"  Createdby "+ usernamesession);


                        var returnedData = AXFinanceOperations.CreateLedgerHeader(ledgerName, ledgerDesc + "  Createdby " + usernamesession,
                            bearaerObj);

                        if (returnedData.IsSuccessful && returnedData.StatusCode == HttpStatusCode.OK)
                        {
                            journalID = returnedData.Content;

                            journalID = Regex.Replace(journalID, @"[^0-9a-zA-Z-]+", "");
                        }
                    }
                    #endregion
                    #region Create Journal Lines


                    if (!string.IsNullOrEmpty(journalID) || !string.IsNullOrEmpty(InvoiceID))
                    {
                        int StartIndex = 0;
                        if (dt.Rows[0][0].ToString() == "رقم المستند")
                            StartIndex = 1;

                        for (int i = StartIndex; i < dt.Rows.Count; i++)    //Create Journal Line 
                        {

                            var details = GetJournalInvoiceDetails(journalID, null);
                            errorMessage =
                                            AXFinanceOperations.CreateLineJournalVat(details.journalID, details.documentNum, details.TransDate, details.Description,
                                            details.ledgerJournalACType, details.account, details.AccountDim, details.amountDebit, details.AmountCredit,
                                            details.account, details.OffsetaccType, details.offSetAccountCode, "", details.PostingProfile, details.vatGroup, details.VatTaxItem, details.bearaerObj);
                            if (!errorMessage.IsSuccessful)
                            {
                                GlobalCode.addRowToTable(dtError, dt.Rows[i]);
                                dtError.Rows[dtError.Rows.Count - 1]["Error"] = errorMessage.Content;
                            }
                        }
                    }
                    #endregion


                }


                DtjournalList = null;


            }
            catch (Exception exc)
            {
                throw new ArgumentException(exc.ToString());
            }
        }
        GridView1.DataSource = dtError;
        GridView1.DataBind();
        Response.Write("Done");
    }


    public masterDetails GetJournalInvoiceDetails(int i, DataTable dt)
    {
       

                if (string.IsNullOrEmpty(dt.Rows[i]["AccountType"].ToString())) continue;


                string PostingProfile = "";
                if (dt.Columns.Contains("PostingProfile"))
                {
                    PostingProfile = dt.Rows[i]["PostingProfile"].ToString();
                }
                if (dt.Rows[i]["Date"].ToString() == "")
                {
                    continue;
                }


                string dimensions = ""; // "BusinessUnit,BU-01,Activity,Recruiting,CostCenter,BR-01,Department,DP-02,Project,projectid,Worker,workerid";
                int DescIndex = 0;

                for (int colIndex = DescIndex; colIndex < dt.Columns.Count; colIndex++)
                {





                    if (dt.Columns[colIndex].ColumnName == "BusinessUnit")
                    {
                        DescIndex = colIndex;
                    }
                    if (DescIndex == 0)
                        continue;
                    if (dt.Columns[colIndex].ColumnName == "PostingProfile")
                        break;
                    if (dt.Rows[i][colIndex].ToString() != "")
                    {
                        if (dt.Rows[i][colIndex].ToString().ToLower() != "crm")
                        {
                            var colValue = string.Empty;
                            colValue = !String.IsNullOrEmpty(dt.Rows[i][colIndex].ToString()) ? dt.Rows[i][colIndex].ToString() :
                                        defaultJournalValues_dict.TryGetValue(dt.Columns[colIndex].ColumnName.Trim(), out colValue) ? colValue :
                                        String.Empty;

                            dimensions += "," + dt.Columns[colIndex].ColumnName + "," + colValue;
                        }
                        else
                        {
                            dimensions += "," + dt.Columns[colIndex].ColumnName + ",@" + dt.Columns[colIndex].ColumnName;
                        }
                    }
                    else
                    {

                    }
                }
                if (dimensions.Length > 1)
                {
                    dimensions = dimensions.Substring(1);
                }
                AccountType ledgerJournalACType = 0;
                decimal amountDebit = 0;
                decimal AmountCredit = 0;
                string AccountDim = "";
                string dimensionDesc = "trans01";
                string Description = "";
                string account = "";
                string documentNum = "";
                AccountType OffsetAccountType = 0;
                string offSetAccountCode = "";
                string OffsetDim = "";
                DateTime TransDate = DateTime.Now;

                try
                {
                    documentNum = dt.Rows[i][0].ToString();
                    Description = dt.Rows[i]["Description"].ToString();
                    if (dt.Rows[i]["Date"].ToString() != "")
                    {
                        if (dt.Rows[i]["Date"] is DateTime)
                        {
                            TransDate = (DateTime)dt.Rows[i]["Date"];
                        }
                        else if (dt.Rows[i]["Date"].ToString().Length > 10)
                        {
                            string strdate = dt.Rows[i]["Date"].ToString().Substring(0, 10);
                            try
                            {

                                TransDate = DateTime.ParseExact(strdate, "yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));

                            }
                            catch (Exception exc)
                            {
                                TransDate = DateTime.ParseExact(strdate, allFormats, new System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None);
                            }
                        }
                        else
                        {
                            int days = 0;
                            string strdate = dt.Rows[i]["Date"].ToString();
                            if (int.TryParse(strdate, out days))
                            {

                                if (days > 59) days -= 1; //Excel/Lotus 2/29/1900 bug   
                                TransDate = new DateTime(1899, 12, 31).AddDays(days);
                            }
                            else
                            {



                                TransDate = DateTime.ParseExact(strdate, allFormats, new System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None);
                            }

                        }
                    }

                    if (dt.Rows[i]["Debit"].ToString() != "")
                    {
                        amountDebit = decimal.Parse(dt.Rows[i]["Debit"].ToString());
                        if (amountDebit > 0)
                        {
                            amountDebit = Math.Round(amountDebit, 2);

                        }
                    }
                    if (dt.Rows[i]["Credit"].ToString() != "")
                    {
                        AmountCredit = decimal.Parse(dt.Rows[i]["Credit"].ToString());
                        if (AmountCredit > 0)
                        {
                            AmountCredit = Math.Round(AmountCredit, 2);

                        }
                    }
                    AccountDim = "";
                    OffsetDim = "";
                    offSetAccountCode = "";
                    OffsetAccountType = 0;



                    if (dt.Rows[i]["AccountType"].ToString() != "" && dt.Rows[i]["AccountCode"].ToString() != "")
                    {
                        ledgerJournalACType = (AccountType)Enum.Parse(typeof(AccountType), dt.Rows[i]["AccountType"].ToString());
                        account = dt.Rows[i]["AccountCode"].ToString();//.ToString().Replace("P", "C");
                                                                       //  if (ledgerJournalACType == AccountType.Ledger)
                        {
                            //AccountDim = dimensions;

                            //if (dt.Rows[i]["Worker"].ToString() != "")
                            //{
                            //    string idSearch = dt.Rows[i]["Worker"].ToString();
                            //    AccountDim = dimensions = GetCostCenters(idSearch, AccountDim);
                            //}
                        }
                        AccountType OffsetaccType = AccountType.Bank;
                        if (dt.Rows[i]["OffsetAccontType"].ToString() != "" && dt.Rows[i]["OffsetAccontCode"].ToString() != "")
                        {

                            offSetAccountCode = dt.Rows[i]["OffsetAccontCode"].ToString();
                            OffsetaccType = (AccountType)Enum.Parse(typeof(AccountType), dt.Rows[i]["OffsetAccontType"].ToString());

                            ////if (accType == AccountType.Ledger)
                            //{
                            //    OffsetDim = dimensions;
                            //    if (dt.Rows[i]["Worker"].ToString() != "")
                            //    {
                            //        string idSearch = dt.Rows[i]["Worker"].ToString();
                            //        OffsetDim = GetCostCenters(idSearch, OffsetDim);
                            //    }
                            //}

                        }




                        //dimensions = BusinessUnitDim;

                        /* // ============= new implementation (MKH) =======================================
                         var workerId = dt.Rows[i]["Worker"].ToString();
                         var JournalDimensionMapper = new JournalDimensionMapper(workerId, defaultJournalValues_dict);
                         AccountDim = dimensions = JournalDimensionMapper.MapAllCostCenterDimensions(dimensions);
                         // ==============================================================================
                         if(!chkIsFixedAsset.Checked)
                         {
                             string transCreditID = connector.CreateLineJournal(journalID, documentNum, TransDate.ToString("MM/dd/yyyy"), Description, (int)ledgerJournalACType, account, AccountDim, amountDebit, AmountCredit, account, (int)OffsetaccType, offSetAccountCode, dimensions, PostingProfile);
                         }
                         else
                         {
                             string TransactionType = dt.Rows[i]["TransactionType"].ToString();
                             string transCreditID = connector.CreateLineJournalFA(journalID, documentNum, TransDate.ToString("MM/dd/yyyy"), Description, (int)ledgerJournalACType, account, AccountDim, amountDebit, AmountCredit, account, (int)OffsetaccType, offSetAccountCode, dimensions, PostingProfile, TransactionType);
                         }
                         */






                        // ============= new implementation (MKH) =======================================
                        if (dt.Columns.Contains("Worker"))
                        {
                            var workerId = dt.Rows[i]["Worker"].ToString();
                            if (!string.IsNullOrEmpty(workerId.ToString()))
                            {
                                var JournalDimensionMapper = new JournalDimensionMapper(workerId, defaultJournalValues_dict);
                                AccountDim = dimensions = JournalDimensionMapper.MapAllCostCenterDimensions(dimensions);

                            }
                            else
                            {


                            }


                        }


                        if (dt.Columns.Contains("Customer"))
                        {
                            var customerId = dt.Rows[i]["Customer"].ToString();
                            if (!string.IsNullOrEmpty(customerId.ToString()))
                            {
                                var JournalDimensionMapper = new JournalDimensionMapper(customerId, defaultJournalValues_dict);
                                AccountDim = dimensions = JournalDimensionMapper.MapAllCostCenterDimensions(dimensions);

                            }
                            else
                            {


                            }


                        }







                        if (!chkIsFixedAsset.Checked)
                        {
                            if (dt.Columns.Contains("VatGroup"))
                            {
                                string vatGroup = dt.Rows[i]["VatGroup"].ToString().Replace(" ", "");

                                string VatTaxItem = "VAT";
                                if (string.IsNullOrEmpty(AccountDim))
                                    AccountDim = dimensions;
                                //string transCreditID = connector.CreateLineJournalVat(journalID, documentNum, TransDate.ToString("MM/dd/yyyy"), Description, (int)ledgerJournalACType, account, AccountDim, amountDebit, AmountCredit, account, (int)OffsetaccType, offSetAccountCode, OffsetDim, PostingProfile, vatGroup, VatTaxItem);

                                return new masterDetails(journalID, documentNum, TransDate.ToString("MM/dd/yyyy"), Description, (int)ledgerJournalACType, account, AccountDim, amountDebit, AmountCredit, account, (int)OffsetaccType, offSetAccountCode, OffsetDim, PostingProfile, vatGroup, VatTaxItem, bearaerObj);

                            }
                            else
                            {

                                if (string.IsNullOrEmpty(AccountDim))
                                    AccountDim = dimensions;
                                // string transCreditID = connector.CreateLineJournal(journalID, documentNum, TransDate.ToString("MM/dd/yyyy"), Description, (int)ledgerJournalACType, account, AccountDim, amountDebit, AmountCredit, account, (int)OffsetaccType, offSetAccountCode, OffsetDim, PostingProfile);

                                return new masterDetails(journalID, documentNum, TransDate.ToString("MM/dd/yyyy"), Description, (int)ledgerJournalACType, account, AccountDim, amountDebit, AmountCredit, account, (int)OffsetaccType, offSetAccountCode, OffsetDim, PostingProfile, bearaerObj);


                            }




                        }
                        else
                        {
                            string vatGroup = dt.Rows[i]["VatGroup"].ToString().Replace(" ", "");
                            string VatTaxItem = "VAT";
                            if (string.IsNullOrEmpty(AccountDim))
                                AccountDim = dimensions;
                            string TransactionType = dt.Rows[i]["TransactionType"].ToString();
                            // string transCreditID = connector.CreateLineJournalFA(journalID, documentNum, TransDate.ToString("MM/dd/yyyy"), Description, (int)ledgerJournalACType, account, AccountDim, amountDebit, AmountCredit, account, (int)OffsetaccType, offSetAccountCode, "", PostingProfile, TransactionType,vatGroup,VatTaxItem );

                            return new masterDetails(journalID, documentNum, TransDate.ToString("MM/dd/yyyy"), Description, (int)ledgerJournalACType, account, AccountDim, amountDebit, AmountCredit, account, (int)OffsetaccType, offSetAccountCode, "", PostingProfile, vatGroup, VatTaxItem, bearaerObj);


                        }


                    }

                }
                catch (Exception exc)
                {

                    GlobalCode.addRowToTable(dtError, dt.Rows[i]);
                    dtError.Rows[dtError.Rows.Count - 1]["Error"] = exc.ToString();
                    lblError.Text = exc.ToString();
                }

            
    }

    public string CreateFreeInvoiceHeader(string Customer, string DocNumber, DateTime Duedate, AXObject bearaerObj)
    {

        FreeTextInvoiceHeaders freetextinvoice = new FreeTextInvoiceHeaders()
        {
            BranchPool = String.Empty,
            CustAccount = Customer,
            CustInvoiceId = DocNumber,
            DueDate = Duedate.ToString("yyyy-MM-dd"),
            TransDate = Duedate.ToString("yyyy-MM-dd"),
            PaySchedule = String.Empty
        };

        string InvoiceHeaderId = "";

        //Create
        var returnedData = AXFinanceOperations.CreateFreeTextInvoiceHeaders(freetextinvoice, bearaerObj);
        if (returnedData.IsSuccessful && returnedData.StatusCode == HttpStatusCode.OK)
        {
            InvoiceHeaderId = returnedData.Content;

            InvoiceHeaderId = Regex.Replace(InvoiceHeaderId, @"[^0-9a-zA-Z-]+", "");
        }
        return InvoiceHeaderId;

    }

    public string CreateFreeInvoiceLines(string journalID, string AccountCode, string LedgerAccount, string Description
        , DataRow dr, decimal amountDebit, decimal AmountCredit, AXObject bearaerObj)
    {
        FreeTextInvoiceLines InvoiceLines = new FreeTextInvoiceLines()
        {
            CustInvoiceRecId = Convert.ToInt64(journalID),
            LedgerAccount = AccountCode,
            InvoiceQuantity = 1m,
            Customer = dr["Customer"].ToString(),
            BusinessUnit = dr["BusinessUnit"].ToString(),
            Branch = dr["Branch"].ToString(),
            Contract = dr["Contract"].ToString(),
            Department = dr["Department"].ToString(),
            FixedAsset = string.Empty,
            Worker = dr["Worker"].ToString(),
            Vendor = dr["Vendor"].ToString(),
            Description = Description,
            InvoiceUnitPrice = Convert.ToDecimal(amountDebit > 0m ? amountDebit.ToString() : AmountCredit.ToString()),
            InvoiceId = dr["InvoiceId"].ToString(),

        };
        string VatGroup = dr["VatGroup"].ToString().Replace(" ", "");
        string salesTaxItem = "VAT";
        InvoiceLines.TaxItemGroup = salesTaxItem;
        InvoiceLines.VatGroup = VatGroup;

        var errorMessage = AXFinanceOperations.CreateFreeTextInvoiceLines(InvoiceLines, bearaerObj);
        return errorMessage.Content.ToString();
    }


    public string GetCostCenters(string empNumber, string costCenter, Dictionary<string, string> defaultValues_dict = null)
    {
        Func<string, string, string> GetCostCenterVal;

        if (defaultValues_dict != null)
            GetCostCenterVal = (key, databaseValue) =>
            {
                var dictVal = string.Empty;
                return !String.IsNullOrEmpty(databaseValue) ? databaseValue :
                        (defaultValues_dict.TryGetValue(key, out dictVal) ? dictVal : String.Empty);
            };
        else
            GetCostCenterVal = (key, databaseValue) => !String.IsNullOrEmpty(databaseValue) ? databaseValue : String.Empty;

        //        string sql = @"SELECT     new_EmployeeBase.new_EmpIdNumber, new_EmployeeBase.new_employeeType AS BusinessUnit, ISNULL(TerritoryBase.new_citycode, 'RG1') AS Region, ISNULL(new_CityBase.new_aiwanumber, 'CT101') AS City,case  new_EmployeeBase.new_employeeType  when 1 then 'BR10101' when 2 then 'BUBranch' when 3 then 'IVBranch' else 'BUBranch' END As Branch, ISNULL(new_departmentBase.new_code, 'D07') 
        //                  AS Department, new_projectBase.new_Code AS Project, new_CountryBase.new_axcode AS Nationality, iif(new_EmployeeBase.new_sex=1,'G01','G02') Gender
        //FROM        new_projectBase INNER JOIN
        //                  new_EmployeeBase ON new_projectBase.new_projectId = new_EmployeeBase.new_projectId LEFT OUTER JOIN
        //                  TerritoryBase INNER JOIN
        //                  new_CityBase ON TerritoryBase.TerritoryId = new_CityBase.new_teritory ON new_EmployeeBase.new_dridinglicenceplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawcity = new_CityBase.new_CityId AND 
        //                  new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId LEFT OUTER JOIN
        //                  new_departmentBase ON new_EmployeeBase.new_departmentId = new_departmentBase.new_departmentId INNER JOIN
        //                  new_CountryBase ON new_EmployeeBase.new_nationalityId = new_CountryBase.new_CountryId ";

        string sql = @"SELECT     new_EmployeeBase.new_EmpIdNumber,
           'BU0'+ cast(cast ( iif( new_EmployeeBase.new_employeeType='100000000','1',new_EmployeeBase.new_employeeType) as int)  as nvarchar) AS BusinessUnit,
            TerritoryBase.new_citycode AS Region, 
            IIF(new_CityBase.new_citycode IS Not NULL,'BR'+ new_CityBase.new_citycode,null) AS CostCenter,
            ISNULL(new_CityBase.new_citycode, 'CT101') AS City,
            case  new_EmployeeBase.new_employeeType  when 1 then 'BR10101' when 2 then 'BUBranch' when 3 then 'IVBranch' else null END As Branch, 
            new_departmentBase.new_code   AS Department, 
            new_projectBase.new_Code AS Contract,
            new_CountryBase.new_axcode AS Nationality , 
            iif(new_EmployeeBase.new_sex=1,'G01','G02') Gender,
            new_profession.new_Code Profession,
            new_customergroup.new_code ContractType,
            new_customergroup.new_invoicecode Activity

FROM        new_projectBase INNER JOIN
            new_EmployeeBase ON new_projectBase.new_projectId = new_EmployeeBase.new_projectId LEFT OUTER JOIN
            TerritoryBase INNER JOIN
            new_CityBase ON TerritoryBase.TerritoryId = new_CityBase.new_teritory ON new_EmployeeBase.new_dridinglicenceplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawcity = new_CityBase.new_CityId AND 
            new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId LEFT OUTER JOIN
            new_departmentBase ON new_EmployeeBase.new_departmentId = new_departmentBase.new_departmentId INNER JOIN
            new_CountryBase ON new_EmployeeBase.new_nationalityId = new_CountryBase.new_CountryId  inner join new_profession on new_profession.new_professionid=new_EmployeeBase.new_professionId
			left outer join new_customergroup  on new_customergroup.new_customergroupId=new_projectBase.new_sectortypeid

UNION ALL

SELECT     new_EmployeeBase.new_EmpIdNumber,
           'BU0'+ cast(cast (  iif( new_EmployeeBase.new_employeeType='100000000','1',new_EmployeeBase.new_employeeType) as int)  as nvarchar) AS BusinessUnit,
            TerritoryBase.new_citycode AS Region, 
            IIF(new_CityBase.new_citycode IS Not NULL,'BR'+ new_CityBase.new_citycode,null) AS CostCenter,
            ISNULL(new_CityBase.new_citycode, 'CT101') AS City,
            case  new_EmployeeBase.new_employeeType  when 1 then 'BR10101' when 2 then 'BUBranch' when 3 then 'IVBranch' else null END As Branch,
            new_departmentBase.new_code   AS Department, 
            new_IndvContract.new_contractNumber AS Contract, 
            new_CountryBase.new_axcode AS Nationality , 
            iif(new_EmployeeBase.new_sex=1,'G01','G02') Gender,
            new_profession.new_Code Profession,
            new_customergroup.new_code ContractType,new_customergroup.new_invoicecode Activity

FROM        new_IndvContract INNER JOIN
            new_EmployeeBase ON new_IndvContract.new_indvContractid = new_EmployeeBase.new_indivcontract LEFT OUTER JOIN
            TerritoryBase INNER JOIN
            new_CityBase ON TerritoryBase.TerritoryId = new_CityBase.new_teritory ON new_EmployeeBase.new_dridinglicenceplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawcity = new_CityBase.new_CityId AND 
            new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId LEFT OUTER JOIN
            new_departmentBase ON new_EmployeeBase.new_departmentId = new_departmentBase.new_departmentId INNER JOIN
            new_CountryBase ON new_EmployeeBase.new_nationalityId = new_CountryBase.new_CountryId  inner join new_profession on new_profession.new_professionid=new_EmployeeBase.new_professionId
			left outer join new_customergroup  on new_customergroup.new_customergroupId=new_IndvContract.new_contracttypeId 
                ";
        sql += "      where new_EmployeeBase.new_idnumber ='" + empNumber + "' or new_empidnumber='" + empNumber + "' or new_borderno='" + empNumber + "' or new_passportnumber='" + empNumber + "'";

        DataTable dt = CRMAccessDB.SelectQ(sql).Tables[0];
        if (dt.Rows.Count > 0)
        {
            costCenter = costCenter.Replace(empNumber, dt.Rows[0]["new_EmpIdNumber"].ToString());
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                costCenter = costCenter.Replace("@" + dt.Columns[i].ColumnName,
                   GetCostCenterVal(dt.Columns[i].ColumnName, dt.Rows[0][i].ToString().Trim()));

                //costCenter = costCenter.Replace("@" + dt.Columns[i].ColumnName,
                //   !String.IsNullOrEmpty(dt.Rows[0][i].ToString().Trim()) ? dt.Rows[0][i].ToString().Trim() : defaultValue.Trim());

                //costCenter = costCenter.Replace("@" + dt.Columns[i].ColumnName,
                //!String.IsNullOrEmpty(dt.Rows[0][i].ToString()) ? dt.Rows[0][i].ToString() :
                //(defaultValues.Columns.Contains(dt.Columns[i].ColumnName)? defaultValues.Rows[0][dt.Columns[i].ColumnName].ToString(): String.Empty));
            }
        }

        return costCenter;
    }
    protected void txtCompay_TextChanged(object sender, EventArgs e)
    {

    }
    //
    // Summary:
    //     Convert Datatable with journal structure to AX
    //
    // Parameters:
    //   datatable:
    //
    // Returns:
    //     the datatable that has error with the exception
    public DataTable UploadDtJournalToAx(DataTable dt)
    {


        string[] allFormats ={"yyyy/MM/dd","yyyy/M/d",
                "dd/MM/yyyy","d/M/yyyy",
                "dd/M/yyyy","d/MM/yyyy","yyyy-MM-dd",
                "yyyy-M-d","dd-MM-yyyy","d-M-yyyy",
                "dd-M-yyyy","d-MM-yyyy","yyyy MM dd",
                "yyyy M d","dd MM yyyy","d M yyyy",
                "dd M yyyy","d MM yyyy"};

        //DataTable dt = ConvertToExcel.ExcelImport.ImportAnyExcel(FileUpload1.FileName, FileUpload1.PostedFile, true).Tables[0];
        DataTable dtError = dt.Clone();
        dtError.Columns.Add("Error");
        if (dt.Rows.Count > 1)
        {
            AXCRMNetBusConnector connector = new AXCRMNetBusConnector();
            //Create LedgerHeader
            string ledgerName = txtJournal.Text;
            string ledgerDesc = "Upload Document " + dt.Rows[2][1].ToString();
            string journalID = string.Empty;
            try
            {
                journalID = txtJournalNum.Text;
                if (txtJournalNum.Text == "")
                {
                    journalID = connector.CreateLedgerHeader(ledgerName, ledgerDesc);
                }
                if (!string.IsNullOrEmpty(journalID))
                {
                    int StartIndex = 0;
                    if (dt.Rows[0][0].ToString() == "رقم المستند")
                        StartIndex = 1;
                    for (int i = StartIndex; i < dt.Rows.Count; i++)
                    {
                        string PostingProfile = "";
                        if (dt.Columns.Contains("PostingProfile"))
                        {
                            PostingProfile = dt.Rows[i]["PostingProfile"].ToString();
                        }
                        if (dt.Rows[i]["Date"].ToString() == "")
                        {
                            continue;
                        }
                        string dimensions = ""; // "BusinessUnit,BU-01,Activity,Recruiting,CostCenter,BR-01,Department,DP-02,Project,projectid,Worker,workerid";
                        int DescIndex = 0;

                        for (int colIndex = DescIndex; colIndex < dt.Columns.Count; colIndex++)
                        {
                            if (dt.Columns[colIndex].ColumnName == "BusinessUnit")
                            {
                                DescIndex = colIndex;
                            }
                            if (DescIndex == 0)
                                continue;
                            if (dt.Columns[colIndex].ColumnName == "PostingProfile")
                                break;
                            if (dt.Rows[i][colIndex].ToString() != "")
                            {
                                if (dt.Rows[i][colIndex].ToString().ToLower() != "crm")
                                {
                                    dimensions += "," + dt.Columns[colIndex].ColumnName + "," + dt.Rows[i][colIndex].ToString();
                                }
                                else
                                {
                                    dimensions += "," + dt.Columns[colIndex].ColumnName + ",@" + dt.Columns[colIndex].ColumnName;
                                }
                            }
                            else
                            {

                            }
                        }
                        if (dimensions.Length > 1)
                        {
                            dimensions = dimensions.Substring(1);
                        }
                        AccountType ledgerJournalACType = 0;
                        decimal amountDebit = 0;
                        decimal AmountCredit = 0;
                        string AccountDim = "";
                        string dimensionDesc = "trans01";
                        string Description = "";
                        string account = "";
                        string documentNum = "";
                        AccountType OffsetAccountType = 0;
                        string offSetAccountCode = "";
                        string OffsetDim = "";
                        DateTime TransDate = DateTime.Now;

                        try
                        {
                            documentNum = dt.Rows[i][0].ToString();
                            Description = dt.Rows[i]["Description"].ToString();
                            if (dt.Rows[i]["Date"].ToString() != "")
                            {
                                if (dt.Rows[i]["Date"] is DateTime)
                                {
                                    TransDate = (DateTime)dt.Rows[i]["Date"];
                                }
                                else if (dt.Rows[i]["Date"].ToString().Length > 10)
                                {
                                    string strdate = dt.Rows[i]["Date"].ToString().Substring(0, 10);
                                    try
                                    {

                                        TransDate = DateTime.ParseExact(strdate, "yyyy-MM-dd", new System.Globalization.CultureInfo("en-US"));

                                    }
                                    catch (Exception exc)
                                    {
                                        TransDate = DateTime.ParseExact(strdate, allFormats, new System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None);
                                    }
                                }
                                else
                                {
                                    int days = 0;
                                    string strdate = dt.Rows[i]["Date"].ToString();
                                    if (int.TryParse(strdate, out days))
                                    {

                                        if (days > 59) days -= 1; //Excel/Lotus 2/29/1900 bug   
                                        TransDate = new DateTime(1899, 12, 31).AddDays(days);
                                    }
                                    else
                                    {



                                        TransDate = DateTime.ParseExact(strdate, allFormats, new System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None);
                                    }

                                }
                            }

                            if (dt.Rows[i]["Debit"].ToString() != "")
                            {
                                amountDebit = decimal.Parse(dt.Rows[i]["Debit"].ToString());
                                if (amountDebit > 0)
                                {
                                    amountDebit = Math.Round(amountDebit, 2);

                                }
                            }
                            if (dt.Rows[i]["Credit"].ToString() != "")
                            {
                                AmountCredit = decimal.Parse(dt.Rows[i]["Credit"].ToString());
                                if (AmountCredit > 0)
                                {
                                    AmountCredit = Math.Round(AmountCredit, 2);

                                }
                            }
                            AccountDim = "";
                            OffsetDim = "";
                            offSetAccountCode = "";
                            OffsetAccountType = 0;
                            if (dt.Rows[i]["AccountType"].ToString() != "" && dt.Rows[i]["AccountCode"].ToString() != "")
                            {
                                ledgerJournalACType = (AccountType)Enum.Parse(typeof(AccountType), dt.Rows[i]["AccountType"].ToString());
                                account = dt.Rows[i]["AccountCode"].ToString();//.ToString().Replace("P", "C");
                                //  if (ledgerJournalACType == AccountType.Ledger)
                                {
                                    AccountDim = dimensions;
                                    //if (dt.Rows[i]["ProjectCode"].ToString() != "")
                                    //{
                                    //    AccountDim = AccountDim.Replace("projectid", dt.Rows[i]["ProjectCode"].ToString());
                                    //}


                                    if (dt.Rows[i]["Worker"].ToString() != "")
                                    {
                                        string idSearch = dt.Rows[i]["Worker"].ToString();
                                        AccountDim = GetCostCenters(idSearch, AccountDim);
                                    }
                                }
                                AccountType OffsetaccType = AccountType.Bank;
                                if (dt.Rows[i]["OffsetAccontType"].ToString() != "" && dt.Rows[i]["OffsetAccontCode"].ToString() != "")
                                {

                                    offSetAccountCode = dt.Rows[i]["OffsetAccontCode"].ToString();
                                    OffsetaccType = (AccountType)Enum.Parse(typeof(AccountType), dt.Rows[i]["OffsetAccontType"].ToString());

                                    //if (accType == AccountType.Ledger)
                                    {
                                        OffsetDim = dimensions;
                                        if (dt.Rows[i]["Worker"].ToString() != "")
                                        {
                                            string idSearch = dt.Rows[i]["Worker"].ToString();
                                            OffsetDim = GetCostCenters(idSearch, OffsetDim);
                                        }
                                    }

                                }


                                string transCreditID = connector.CreateLineJournal(journalID, documentNum, TransDate.ToString("MM/dd/yyyy"), Description, (int)ledgerJournalACType, account, AccountDim, amountDebit, AmountCredit, account, (int)OffsetaccType, offSetAccountCode, OffsetDim, PostingProfile);
                            }


                        }
                        catch (Exception exc)
                        {

                            GlobalCode.addRowToTable(dtError, dt.Rows[i]);
                            dtError.Rows[dtError.Rows.Count - 1]["Error"] = exc.ToString();
                            lblError.Text = exc.ToString();
                        }

                    }
                }
            }
            catch (Exception exc)
            {
                throw new ArgumentException(exc.ToString());
            }
        }
        return dtError;
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        string MapFileName = "UploadExcellsheetTemplatewithCOA.xlsx";
        Response.Clear();
        Response.AddHeader("Content-Disposition", "attachment; filename=\"" + MapFileName.Replace("xml", "xls") + "\"");
        Response.TransmitFile(MapFileName);
        Response.End();
    }
    protected void Button2_Click(object sender, EventArgs e)
    {

        string MapFileName = "UploadExcellsheetTemplatewithCOA (1).xml";
        Response.Clear();
        Response.AddHeader("Content-Disposition", "attachment; filename=\"" + MapFileName.Replace("xml", "xls") + "\"");
        Response.TransmitFile(MapFileName);
        Response.End();
    }
    protected void Button3_Click(object sender, EventArgs e)
    {
        Response.Redirect("http://tamk.excprotection.com:8000/AX/CreateFixedAssets.aspx");
    }
}

public class JournalDimensionMapper
{
    protected static string SqlQuery { get; private set; }
    public string WorkerId { get; set; }
    public DataTable dimensionsDataTable { get; private set; }
    public Dictionary<String, String> DefaultDimensionsValuesDict { get; set; }

    static JournalDimensionMapper()
    {
        SqlQuery = @"SELECT     new_EmployeeBase.new_EmpIdNumber,
                                'BU0'+ cast(cast ( iif( new_EmployeeBase.new_employeeType='100000000','1',new_EmployeeBase.new_employeeType) as int)  as nvarchar) AS BusinessUnit,
                                 TerritoryBase.new_citycode AS Region, 
                                 IIF(new_CityBase.new_citycode IS Not NULL,'BR'+ new_CityBase.new_citycode,null) AS CostCenter,
                                 ISNULL(new_CityBase.new_citycode, 'CT101') AS City,
                                 case  new_EmployeeBase.new_employeeType  when 1 then 'BR10101' when 2 then 'BUBranch' when 3 then 'IVBranch' else null END As Branch, 
                                 new_departmentBase.new_code   AS Department, 
                                 new_projectBase.new_Code AS Contract,
                                 new_CountryBase.new_axcode AS Nationality , 
                                 iif(new_EmployeeBase.new_sex=1,'G01','G02') Gender,
                                 new_profession.new_Code Profession,
                                 new_customergroup.new_code ContractType,
                                 new_customergroup.new_invoicecode Activity
                     
                     FROM        new_projectBase INNER JOIN
                                 new_EmployeeBase ON new_projectBase.new_projectId = new_EmployeeBase.new_projectId LEFT OUTER JOIN
                                 TerritoryBase INNER JOIN
                                 new_CityBase ON TerritoryBase.TerritoryId = new_CityBase.new_teritory ON new_EmployeeBase.new_dridinglicenceplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawcity = new_CityBase.new_CityId AND 
                                 new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId LEFT OUTER JOIN
                                 new_departmentBase ON new_EmployeeBase.new_departmentId = new_departmentBase.new_departmentId INNER JOIN
                                 new_CountryBase ON new_EmployeeBase.new_nationalityId = new_CountryBase.new_CountryId  inner join new_profession on new_profession.new_professionid=new_EmployeeBase.new_professionId
                     			left outer join new_customergroup  on new_customergroup.new_customergroupId=new_projectBase.new_sectortypeid
                              WHERE new_EmployeeBase.new_idnumber = @empId or new_empidnumber = @empId or new_borderno = @empId or new_passportnumber = @empId

                     UNION ALL
                     
                     SELECT     new_EmployeeBase.new_EmpIdNumber,
                                'BU0'+ cast(cast (  iif( new_EmployeeBase.new_employeeType='100000000','1',new_EmployeeBase.new_employeeType) as int)  as nvarchar) AS BusinessUnit,
                                 TerritoryBase.new_citycode AS Region, 
                                 IIF(new_CityBase.new_citycode IS Not NULL,'BR'+ new_CityBase.new_citycode,null) AS CostCenter,
                                 ISNULL(new_CityBase.new_citycode, 'CT101') AS City,
                                 case  new_EmployeeBase.new_employeeType  when 1 then 'BR10101' when 2 then 'BUBranch' when 3 then 'IVBranch' else null END As Branch,
                                 new_departmentBase.new_code   AS Department, 
                                 new_IndvContract.new_contractNumber AS Contract, 
                                 new_CountryBase.new_axcode AS Nationality , 
                                 iif(new_EmployeeBase.new_sex=1,'G01','G02') Gender,
                                 new_profession.new_Code Profession,
                                 new_customergroup.new_code ContractType,new_customergroup.new_invoicecode Activity
                     
                     FROM        new_IndvContract INNER JOIN
                                 new_EmployeeBase ON new_IndvContract.new_indvContractid = new_EmployeeBase.new_indivcontract LEFT OUTER JOIN
                                 TerritoryBase INNER JOIN
                                 new_CityBase ON TerritoryBase.TerritoryId = new_CityBase.new_teritory ON new_EmployeeBase.new_dridinglicenceplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawcity = new_CityBase.new_CityId AND 
                                 new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId AND new_EmployeeBase.new_lawworkplace = new_CityBase.new_CityId LEFT OUTER JOIN
                                 new_departmentBase ON new_EmployeeBase.new_departmentId = new_departmentBase.new_departmentId INNER JOIN
                                 new_CountryBase ON new_EmployeeBase.new_nationalityId = new_CountryBase.new_CountryId  inner join new_profession on new_profession.new_professionid=new_EmployeeBase.new_professionId
                     			left outer join new_customergroup  on new_customergroup.new_customergroupId=new_IndvContract.new_contracttypeId 
                                     
                    WHERE new_EmployeeBase.new_idnumber = @empId or new_empidnumber = @empId or new_borderno = @empId or new_passportnumber = @empId";


    }

    public JournalDimensionMapper(string workerId) : this(workerId, null)
    {
    }

    public JournalDimensionMapper(string workerId, Dictionary<String, String> defaultDimensionsValuesDict)
    {
        this.WorkerId = workerId;
        if (defaultDimensionsValuesDict == null)
            defaultDimensionsValuesDict = new Dictionary<string, string>();
        this.DefaultDimensionsValuesDict = defaultDimensionsValuesDict;

        SelectDimensionsDataTable();

    }

    public DataTable SelectDimensionsDataTable()
    {
        var cmd = new SqlCommand(SqlQuery);
        cmd.Parameters.Add(new SqlParameter("@empId", this.WorkerId));
        dimensionsDataTable = CRMAccessDB.SelectQ(cmd).Tables[0];
        return dimensionsDataTable;
    }
    /// <summary>
    /// map only the passed cost center
    /// </summary>
    /// <param name="joinedDimensions"> string in the following Format "[ColumnName],[ColumnName],[ColumnName],[ColumnName]"</param>
    /// <returns>
    /// string in the following Format "[CostCenterKey],[Value],[CostCenterKey],[ColumnValue]"
    /// </returns>
    public string MapCostCenterDimension(string joinedDimensions, string costCenterName)
    {
        if (dimensionsDataTable != null && dimensionsDataTable.Rows.Count > 0)
        {
            var mappedDimensions = joinedDimensions.Replace(this.WorkerId, dimensionsDataTable.Rows[0]["new_EmpIdNumber"].ToString());
            mappedDimensions = mappedDimensions.Replace(String.Format("@{0}", costCenterName), GetCostCenterValue(costCenterName.Trim()));
            return mappedDimensions;
        }
        return joinedDimensions;
    }

    /// <summary>
    /// map all the  cost centers
    /// </summary>
    /// <param name="joinedDimensions"> string in the following Format "[ColumnName],[ColumnName],[ColumnName],[ColumnName]"</param>
    /// <returns>
    /// string in the following Format "[CostCenterKey],[Value],[CostCenterKey],[ColumnValue]"
    /// </returns>
    public string MapAllCostCenterDimensions(string joinedDimensions)
    {
        if (dimensionsDataTable != null && dimensionsDataTable.Rows.Count > 0)
        {
            var mappedDimensions = joinedDimensions.Replace(this.WorkerId, dimensionsDataTable.Rows[0]["new_EmpIdNumber"].ToString());
            for (int i = 0; i < dimensionsDataTable.Columns.Count; i++)
            {
                mappedDimensions = mappedDimensions.Replace(String.Format("@{0}", dimensionsDataTable.Columns[i].ColumnName), GetCostCenterValue(dimensionsDataTable.Columns[i].ColumnName.Trim()));
            }
            return mappedDimensions;
        }
        return joinedDimensions;
    }

    protected string GetCostCenterValue(string costCenterName)
    {
        string value = string.Empty;
        value = (dimensionsDataTable.Columns.Contains(costCenterName) && dimensionsDataTable.Rows[0][costCenterName] != null) ? dimensionsDataTable.Rows[0][costCenterName].ToString()
                :
                DefaultDimensionsValuesDict.TryGetValue(costCenterName, out value) ? value : String.Empty;
        return value;
    }
}