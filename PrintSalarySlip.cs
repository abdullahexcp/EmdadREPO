using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Configuration;
using ConvertToExcel;

public partial class Templates_ProjectInvoice_PrintSalary : GlobalPage
{
    string salarySlipQuery =
        @"SELECT    
ROW_NUMBER()OVER(ORDER BY new_EmployeeBase.new_IDNumber DESC) AS rowxx, 
 DateName( month , DateAdd( month , month(new_invoiceBase.new_todate) , -1 ) ) slipMonth,
year(new_invoiceBase.new_todate) slipYear,
new_EmployeeBase.new_name AS اسمالموظف
, 0 AS بدلالنقل
, 0 AS بدلطعام
, 0 AS بدلالسكن
, (ISNULL(new_invoicdetailBase.new_absenceDeductions,0)+ISNULL(new_invoicdetailBase.new_otherdeduction,0)) AS الغياب
, new_invoiceBase.new_todate AS تاريخالفاتورة,
 DateName( month , DateAdd( month , month(new_invoiceBase.new_todate) , -1 ) ) slipMonth,
year(new_invoiceBase.new_todate) slipYear

, new_invoicdetailBase.new_totalsalary AS الراتبالاساسي
, new_invoicdetailBase.new_customerloan AS سلفةعميل
, new_invoicdetailBase.new_loanamount AS سلفةموارد
, new_projectBase.new_nameenglish AS ProjectName
, new_projectBase.new_name AS اسمالمشروع
, new_invoicdetailBase.new_absenceDeductions + new_invoicdetailBase.new_loanamount + new_invoicdetailBase.new_customerloan + new_invoicdetailBase.new_otherdeduction AS اجماليالخصم
, new_invoicdetailBase.new_overtime AS قيمةالوقتالاضافي
, new_invoicdetailBase.new_nettoemployee AS صافيالراتب
, new_invoicdetailBase.new_othercommission AS زياداتاخرى
, new_invoicdetailBase.new_overtime + new_invoicdetailBase.new_othercommission AS اجماليالاضافي
,convert(nvarchar(11),new_EmployeeBase.new_DateOfJoining,103) joindtex
,new_EmployeeBase.new_cognizant_emp_code cogempco
, new_EmployeeBase.new_EmpIdNumber AS employeenumber
, new_EmployeeBase.new_IDNumber AS Iqamanumber
,CONVERT(int, new_invoicdetailBase.new_workingdays)
, new_invoiceBase.new_invoicedate as InvoiceDate,

new_projectbase.new_contractcode ContractNumber,
new_invoicdetailBase.new_invoicenumber InvoiceNumber
,replace(new_EmployeeBase.new_EmpIdNumber ,'ML-','') Empnumber
, new_EmployeeBase.new_idnumber IDNumber 
/*,left(new_EmployeeBase.new_name,20) اسمالموظف*/
,convert(varchar, dateadd(hh,3, new_invoicdetailBase.new_startworkdate), 103) StartDate
,new_Country.new_nameenglish Country
,left(new_professionBase.new_ProfessionEnglish,20) Profession
,left(new_professionBase.new_name,20) ProfX
,cast(isnull(new_invoicdetailBase.new_workingdays,0) as int) Days_
,cast(isnull(new_invoicdetailBase.new_gossifees,0) as decimal(10,2)) Gosi

,cast(isnull(new_EmployeeBase.new_basicsalary,0) as decimal(10,2)) BasicSalary
,cast(isnull(new_EmployeeBase.new_HousingAllowance,0) as decimal(10,2)) HousXRH
,cast(isnull(new_EmployeeBase.new_overtimeallowance,0) as decimal(10,2)) TRA

/*,cast(isnull(new_EmployeeBase.new_totalSalary,0) as decimal(10,2)) Salary*/

,
cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2)) AS ovrTimeAllow
,cast(isnull(new_invoicdetailBase.new_nettoemployee,0) as decimal(10,2)) duesalaries
,cast(isnull(new_invoicdetailBase.new_abscentdays,0) as decimal(10,2)) abscentdays
, cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2)) AS otherallowance
, cast(isnull(new_invoicdetailBase.new_duesalary,0) as decimal(10,2)) AS dusalry
,isnull(new_invoicdetailBase.new_workingdays,0) workdys
,isnull(new_invoicdetailBase.new_workingdays,0) workDays


,cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2)) abcded /* الغياب*/
,cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))   emploan /* سلفةموارد*/
,cast(isnull(new_invoicdetailBase.new_customerloan,0) as decimal(10,2))   loanamount /* سلفةعميل */
,cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2)) OTP /* زياداتاخرى*/
,cast(isnull(new_invoicdetailBase.new_bonus,0) as decimal(10,2)) bonusX
,cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2)) overtime /* قيمةالوقتالاضافي*/
,cast(isnull(new_invoicdetailBase.new_totalsalary,0)as decimal(10,2)) xsal /*بركات*/
,cast(isnull(new_invoicdetailBase.new_hardworkallowance,0) as decimal(10,2)) projectallownce
,cast(isnull(new_invoicdetailBase.new_missionallownce,0) as decimal(10,2)) missionallownce
,cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_internaldeduction,0) as decimal(10,2)) otherdeduction
,cast(isnull(new_invoicdetailBase.new_HRADeduction,0) as decimal(10,2)) HRRADeduction

/*,(cast(isnull(new_EmployeeBase.new_basicsalary,0) as decimal(10,2))+cast(isnull(new_EmployeeBase.new_HousingAllowance,0) as decimal(10,2))+cast(isnull(new_EmployeeBase.new_TrasnAllowance,0) as decimal(10,2)))-cast(isnull(new_invoicdetailBase.new_abscentdays,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_hardworkallowance,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_missionallownce,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2))-cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2))-cast(isnull(new_invoicdetailBase.new_HRADeduction,0) as decimal(10,2)) -cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))-cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2)) nettoemployee*/
,cast((
(cast(isnull(new_EmployeeBase.new_basicsalary,0) as decimal(10,2))
+cast(isnull(new_EmployeeBase.new_HousingAllowance,0) as decimal(10,2))
+cast(isnull(new_EmployeeBase.new_TrasnAllowance,0) as decimal(10,2)))
/30*cast(isnull(new_invoicdetailBase.new_workingdays,0) as int))-
-cast(isnull(new_invoicdetailBase.new_abscentdays,0) as decimal(10,2))
+cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2))
+cast(isnull(new_invoicdetailBase.new_hardworkallowance,0) as decimal(10,2))
+cast(isnull(new_invoicdetailBase.new_missionallownce,0) as decimal(10,2))
+cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2))
-cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2))
-cast(isnull(new_invoicdetailBase.new_HRADeduction,0) as decimal(10,2)) 
-cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))
-cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2))  as decimal(10,2)) nettoemployee
/*,cast(isnull(new_invoicdetailBase.new_nettoemployee,0) as decimal(10,2)) nettoemployee*/

,cast(isnull(new_invoicdetailBase.new_duesalary,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2)) as totCredit,
cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_customerloan,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_employeegossi,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2)) as totDebit,
(cast(isnull(new_invoicdetailBase.new_duesalary,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2)))-
(cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_customerloan,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_employeegossi,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_internaldeduction,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2))) netSalary,
cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2)) AS othIncrease,
cast(isnull(new_invoicdetailBase.new_mwfeesdays,0) as decimal(10,2)) dueamount
,cast(isnull(new_invoicdetailBase.new_gossifees,0) as decimal(10,2)) gossifees
,cast(isnull(new_invoicdetailBase.new_paymawaridfees,0) as decimal(10,2)) paymawaridfees
,cast(isnull(new_invoicdetailBase.new_foodallowance,0) as decimal(10,2)) fodalwnce
,cast(isnull(new_EmployeeBase.new_foodallowance,0) as decimal(10,2)) fodlwncx
,cast(isnull(new_EmployeeBase.new_TrasnAllowance,0) as decimal(10,2)) trnlwncx

,cast(cast(isnull(new_invoicdetailBase.new_hardworkallowance,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_missionallownce,0) as decimal(10,2))-cast(isnull(new_invoicdetailBase.new_HRADeduction,0) as decimal(10,2))+isnull((new_invoicdetailBase.new_nettoemployee+new_invoicdetailBase.new_mwfeesdays+new_invoicdetailBase.new_gossifees+cast((case when new_loanamount=0 then new_customerloan when new_loanamount is null then new_customerloan else new_loanamount end)   as decimal(10,2))),0) as decimal(10,2)) totalnetamount
,new_totalnetamount as netpay
,new_invoicdetailBase.new_notes as Notes
,new_EmployeeBase.EmailAddress EmailAddress
,new_invoiceBase.new_projectId projid
,SystemUser.FullName acnamx,SystemUser.internalemailaddress acmailx


FROM         new_invoiceBase INNER JOIN
                      new_invoicdetailBase ON new_invoiceBase.new_invoiceId = new_invoicdetailBase.new_invoiceId INNER JOIN
                      new_EmployeeBase ON new_invoicdetailBase.new_employeeId = new_EmployeeBase.new_EmployeeId INNER JOIN
                      new_projectBase ON new_invoiceBase.new_projectId = new_projectBase.new_projectId   
					  left outer join new_professionBase on new_professionBase.new_professionId=new_EmployeeBase.new_priceprofessionid  
					  left outer join new_Country on new_Country.new_countryid=new_EmployeeBase.new_nationalityId 
					  left outer join SystemUser on SystemUser.SystemUserid=   new_invoiceBase.ownerid

";
    string sqlQueryHeadOffice =
        @"SELECT    
ROW_NUMBER()OVER(ORDER BY new_EmployeeBase.new_IDNumber DESC) AS rowxx, 
 DateName( month , DateAdd( month , month(new_invoiceBase.new_todate) , -1 ) ) slipMonth,
CASE WHEN month(new_invoiceBase.new_todate)=1 THEN N'يناير' 
      WHEN month(new_invoiceBase.new_todate)=2 THEN N'فبراير' 
      WHEN month(new_invoiceBase.new_todate)=3 THEN N'مارس' 
      WHEN month(new_invoiceBase.new_todate)=4 THEN N'ابريل' 
      WHEN month(new_invoiceBase.new_todate)=5 THEN N'مايو' 
      WHEN month(new_invoiceBase.new_todate)=6 THEN N'يونيو' 
      WHEN month(new_invoiceBase.new_todate)=7 THEN N'يوليو' 
      WHEN month(new_invoiceBase.new_todate)=8 THEN N'أغسطس' 
      WHEN month(new_invoiceBase.new_todate)=9 THEN N'سبتمبر' 
      WHEN month(new_invoiceBase.new_todate)=10 THEN N'أكتوبر' 
      WHEN month(new_invoiceBase.new_todate)=11 THEN N'نوفمبر' 
      WHEN month(new_invoiceBase.new_todate)=12 THEN N'ديسمبر' END as slipMonthArabic,
year(new_invoiceBase.new_todate) slipYear,new_invoiceBase.new_todate AS تاريخالفاتورة,
new_EmployeeBase.new_name AS اسمالموظف,new_EmployeeBase.new_namearabic اسمالموظفعربي,  new_EmployeeBase.new_EmpIdNumber AS employeenumber,
left(new_professionBase.new_ProfessionEnglish,20) Profession,
convert(nvarchar(11),new_EmployeeBase.new_DateOfJoining,103) joindtex,
new_projectBase.new_nameenglish AS ProjectName,
cast(isnull(new_EmployeeBase.new_basicsalary,0) as decimal(10,2)) BasicSalary,
cast(isnull(new_EmployeeBase.new_TrasnAllowance,0) as decimal(10,2)) trnlwncx,
cast(isnull(new_EmployeeBase.new_foodAllowance,0) as decimal(10,2)) foodAllow,
cast(isnull(new_EmployeeBase.new_HousingAllowance,0) as decimal(10,2)) HousXRH,
cast(isnull(new_EmployeeBase.new_carAllowance,0) as decimal(10,2)) carAllow,
cast(isnull(new_EmployeeBase.new_mobileallowance,0) as decimal(10,2)) mobileAllow,
cast(isnull(new_EmployeeBase.new_OvertimeAllowance,0) as decimal(10,2)) fxdOverTime,
cast(isnull(new_EmployeeBase.new_OtherAllowance,0) as decimal(10,2)) otherallowance,
cast(isnull(new_EmployeeBase.new_totalSalary,0) as decimal(10,2)) totSalary,
CONVERT(int, new_invoicdetailBase.new_workingdays) workDays,
cast(isnull(new_invoicdetailBase.new_duesalary,0) as decimal(10,2)) AS dueSalary,
cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2)) AS ovrTimeAllow,
cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2)) AS othIncrease,
cast(isnull(new_invoicdetailBase.new_duesalary,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2)) as totCredit,
cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2)) abcded ,
cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))   emploan,
cast(isnull(new_invoicdetailBase.new_employeegossi,0) as decimal(10,2)) empGossi,
cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2)) otherDeduct,
cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_employeegossi,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2)) as totDebit,
(cast(isnull(new_invoicdetailBase.new_duesalary,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2)))-
(cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_employeegossi,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2))+
cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2))) netSalary
,new_EmployeeBase.EmailAddress EmailAddress


FROM         new_invoiceBase INNER JOIN
                      new_invoicdetailBase ON new_invoiceBase.new_invoiceId = new_invoicdetailBase.new_invoiceId INNER JOIN
                      new_EmployeeBase ON new_invoicdetailBase.new_employeeId = new_EmployeeBase.new_EmployeeId INNER JOIN
                      new_projectBase ON new_invoiceBase.new_projectId = new_projectBase.new_projectId   
					  left outer join new_professionBase on new_professionBase.new_professionId=new_EmployeeBase.new_professionid  
";

    protected void Page_Load(object sender, EventArgs e) { }

    protected void BtnDownloadVertas_Click(object sender, EventArgs e)
    {
        string FileText = File.ReadAllText(
            Server.MapPath("~/Templates/ProjectInvoice/Salary-SlipVertas.xml")
        );
        FileText = GlobalCode.ReplaceTempPassword(
            FileText,
            ConfigurationManager.AppSettings["MSWordhash"],
            ConfigurationManager.AppSettings["MSWordsalt"],
            ConfigurationManager.AppSettings["MSWordProviderType"],
            ConfigurationManager.AppSettings["MSWordAlgorithmSid"],
            Request.QueryString["UserID"]
        );
        #region old query split the slary
        /* old query split the salsary
        string sql = @"SELECT     new_EmployeeBase.new_name AS اسمالموظف, ISNULL(new_EmployeeBase.new_TrasnAllowance, 0) AS بدلالنقل,
                      new_invoicdetailBase.new_foodallowance AS بدلطعام, ISNULL(new_EmployeeBase.new_HousingAllowance, 0) AS بدلالسكن, 0 AS otherallowance,
                      new_invoicdetailBase.new_absenceDeductions AS الغياب, new_invoiceBase.new_invoicedate AS تاريخالفاتورة,
                      CASE WHEN new_HousingAllowance > 0 THEN ROUND((isnull(new_EmployeeBase.new_BasicSalary, 0)
                      + ISNULL(new_EmployeeBase.new_OverTimeAllowance, 0) + isnull(new_otherallowance, 0)), 0)
                      ELSE ROUND((isnull(new_EmployeeBase.new_totalSalary - ISNULL(new_EmployeeBase.new_foodallowance, 0), 0)), 0) END AS الراتبالاساسي,
                      new_invoicdetailBase.new_customerloan AS سلفةعميل, new_invoicdetailBase.new_loanamount AS سلفةموارد,
                      new_projectBase.new_nameenglish AS ProjectName,
                      new_invoicdetailBase.new_absenceDeductions + new_invoicdetailBase.new_loanamount + new_invoicdetailBase.new_customerloan AS
                       اجماليالخصم, new_invoicdetailBase.new_overtime AS قيمةالوقتالاضافي, new_invoicdetailBase.new_nettoemployee AS صافيالراتب,
                      new_invoicdetailBase.new_othercommission AS زياداتاخرى,
                      new_invoicdetailBase.new_overtime + new_invoicdetailBase.new_othercommission AS اجماليالاضافي,
                      new_EmployeeBase.new_EmpIdNumber as employeenumber,new_EmployeeBase.new_IDNumber as Iqamanumber
FROM         new_invoiceBase INNER JOIN
                      new_invoicdetailBase ON new_invoiceBase.new_invoiceId = new_invoicdetailBase.new_invoiceId INNER JOIN
                      new_EmployeeBase ON new_invoicdetailBase.new_employeeId = new_EmployeeBase.new_EmployeeId INNER JOIN
                      new_projectBase ON new_invoiceBase.new_projectId = new_projectBase.new_projectId

";
         * */
        #endregion

        string sql =
            @"SELECT    
new_EmployeeBase.new_name AS اسمالموظف
, 0 AS بدلالنقل
, 0 AS بدلطعام
, 0 AS بدلالسكن
, 0 AS otherallowance
, (ISNULL(new_invoicdetailBase.new_absenceDeductions,0)+ISNULL(new_invoicdetailBase.new_otherdeduction,0)) AS الغياب
, new_invoiceBase.new_todate AS تاريخالفاتورة
, new_invoicdetailBase.new_totalsalary AS الراتبالاساسي
, new_invoicdetailBase.new_customerloan AS سلفةعميل
, new_invoicdetailBase.new_loanamount AS سلفةموارد
, new_projectBase.new_nameenglish AS ProjectName
, new_invoicdetailBase.new_absenceDeductions + new_invoicdetailBase.new_loanamount + new_invoicdetailBase.new_customerloan+ new_invoicdetailBase.new_otherdeduction AS اجماليالخصم
, new_invoicdetailBase.new_overtime AS قيمةالوقتالاضافي
, new_invoicdetailBase.new_nettoemployee AS صافيالراتب
, new_invoicdetailBase.new_othercommission AS زياداتاخرى
, new_invoicdetailBase.new_overtime + new_invoicdetailBase.new_othercommission AS اجماليالاضافي

, new_EmployeeBase.new_EmpIdNumber AS employeenumber
, new_EmployeeBase.new_IDNumber AS Iqamanumber
, new_invoicdetailBase.new_workingdays
, new_invoiceBase.new_invoicedate as InvoiceDate,


new_projectbase.new_contractcode ContractNumber,
new_invoicdetailBase.new_invoicenumber InvoiceNumber
,replace(new_EmployeeBase.new_EmpIdNumber ,'ML-','') Empnumber
, new_EmployeeBase.new_idnumber IDNumber 
/*,left(new_EmployeeBase.new_name,20) اسمالموظف*/
,convert(varchar, dateadd(hh,3, new_invoicdetailBase.new_startworkdate), 103) StartDate
,new_Country.new_nameenglish Country
,left(new_professionBase.new_ProfessionEnglish,20) Profession
,cast(isnull(new_invoicdetailBase.new_workingdays,0) as int) Days
,cast(isnull(new_invoicdetailBase.new_gossifees,0) as decimal(10,2)) Gosi

,cast(isnull(new_EmployeeBase.new_basicsalary,0) as decimal(10,2)) BasicSalary
,cast(isnull(new_EmployeeBase.new_HousingAllowance,0) as decimal(10,2)) HRA
,cast(isnull(new_EmployeeBase.new_TrasnAllowance,0) as decimal(10,2)) TRA

/*,cast(isnull(new_EmployeeBase.new_totalSalary,0) as decimal(10,2)) Salary*/

,cast(isnull(new_invoicdetailBase.new_totalnetamount,0) as decimal(10,2)) Cust
,cast(isnull(new_invoicdetailBase.new_nettoemployee,0) as decimal(10,2)) duesalaries
,cast(isnull(new_invoicdetailBase.new_abscentdays,0) as decimal(10,2)) abscentdays


,cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2)) abcded /* الغياب*/
,cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))   emploan /* سلفةموارد*/
,cast(isnull(new_invoicdetailBase.new_customerloan,0) as decimal(10,2))   loanamount /* سلفةعميل */
,cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2)) OTP /* زياداتاخرى*/
,cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2)) overtime /* قيمةالوقتالاضافي*/
,cast(isnull(new_invoicdetailBase.new_totalsalary,0)as decimal(10,2)) xsal /*بركات*/
,cast(isnull(new_invoicdetailBase.new_hardworkallowance,0) as decimal(10,2)) projectallownce
,cast(isnull(new_invoicdetailBase.new_missionallownce,0) as decimal(10,2)) missionallownce
,cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2)) otherdeduction
,cast(isnull(new_invoicdetailBase.new_HRADeduction,0) as decimal(10,2)) HRRADeduction

/*,(cast(isnull(new_EmployeeBase.new_basicsalary,0) as decimal(10,2))+cast(isnull(new_EmployeeBase.new_HousingAllowance,0) as decimal(10,2))+cast(isnull(new_EmployeeBase.new_TrasnAllowance,0) as decimal(10,2)))-cast(isnull(new_invoicdetailBase.new_abscentdays,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_hardworkallowance,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_missionallownce,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2))-cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2))-cast(isnull(new_invoicdetailBase.new_HRADeduction,0) as decimal(10,2)) -cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))-cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2)) nettoemployee*/
,cast((
(cast(isnull(new_EmployeeBase.new_basicsalary,0) as decimal(10,2))
+cast(isnull(new_EmployeeBase.new_HousingAllowance,0) as decimal(10,2))
+cast(isnull(new_EmployeeBase.new_TrasnAllowance,0) as decimal(10,2)))
/30*cast(isnull(new_invoicdetailBase.new_workingdays,0) as int))-
-cast(isnull(new_invoicdetailBase.new_abscentdays,0) as decimal(10,2))
+cast(isnull(new_invoicdetailBase.new_overtime,0) as decimal(10,2))
+cast(isnull(new_invoicdetailBase.new_hardworkallowance,0) as decimal(10,2))
+cast(isnull(new_invoicdetailBase.new_missionallownce,0) as decimal(10,2))
+cast(isnull(new_invoicdetailBase.new_othercommission,0) as decimal(10,2))
-cast(isnull(new_invoicdetailBase.new_otherdeduction,0) as decimal(10,2))
-cast(isnull(new_invoicdetailBase.new_HRADeduction,0) as decimal(10,2)) 
-cast(isnull(new_invoicdetailBase.new_loanamount,0) as decimal(10,2))
-cast(isnull(new_invoicdetailBase.new_absencedeductions,0) as decimal(10,2))  as decimal(10,2)) nettoemployee
/*,cast(isnull(new_invoicdetailBase.new_nettoemployee,0) as decimal(10,2)) nettoemployee*/

,cast(isnull(new_invoicdetailBase.new_mwfeesdays,0) as decimal(10,2)) dueamount
,cast(isnull(new_invoicdetailBase.new_gossifees,0) as decimal(10,2)) gossifees
,cast(isnull(new_invoicdetailBase.new_paymawaridfees,0) as decimal(10,2)) paymawaridfees

,cast(cast(isnull(new_invoicdetailBase.new_hardworkallowance,0) as decimal(10,2))+cast(isnull(new_invoicdetailBase.new_missionallownce,0) as decimal(10,2))-cast(isnull(new_invoicdetailBase.new_HRADeduction,0) as decimal(10,2))+isnull((new_invoicdetailBase.new_nettoemployee+new_invoicdetailBase.new_mwfeesdays+new_invoicdetailBase.new_gossifees+cast((case when new_loanamount=0 then new_customerloan when new_loanamount is null then new_customerloan else new_loanamount end)   as decimal(10,2))),0) as decimal(10,2)) totalnetamount
,new_totalnetamount as netpay
,new_invoicdetailBase.new_notes as Notes

FROM         new_invoiceBase INNER JOIN
                      new_invoicdetailBase ON new_invoiceBase.new_invoiceId = new_invoicdetailBase.new_invoiceId INNER JOIN
                      new_EmployeeBase ON new_invoicdetailBase.new_employeeId = new_EmployeeBase.new_EmployeeId INNER JOIN
                      new_projectBase ON new_invoiceBase.new_projectId = new_projectBase.new_projectId   
					  left outer join new_professionBase on new_professionBase.new_professionId=new_EmployeeBase.new_priceprofessionid  
					  left outer join new_Country on new_Country.new_countryid=new_EmployeeBase.new_nationalityId        
                      

";

        try
        {
            var projectInvoiceId = Request.QueryString["id"];

            if (!string.IsNullOrEmpty(projectInvoiceId))
            {
                sql += " where new_invoiceBase.new_invoiceId = '@id'";
                sql = sql.Replace("@id", projectInvoiceId);
            }
            else
            {
                sql += "Where ";

                if (string.IsNullOrEmpty(fromdate.Text))
                {
                    sql += "YEAR(new_invoiceBase.new_fromdate) = YEAR(GETDATE())";
                }
                else
                {
                    sql += "new_invoiceBase.new_fromdate >= CONVERT(datetime, '" + fromdate.Text + "')";
                }



                if (string.IsNullOrEmpty(todate.Text))
                {
                    sql += " and YEAR(new_invoiceBase.new_todate) = YEAR(GETDATE())";
                }
                else
                {
                    sql += " and new_invoiceBase.new_todate >= CONVERT(datetime, '" + todate.Text + "')";
                }
            }

            DataSet myds = new DataSet();
        myds = CRMAccessDB.SelectQ(sql);
        DataTable dt = myds.Tables[0];
        DateTime now = DateTime.Now;

        for (int i = 0; i < dt.Rows.Count; i++)
        {
            now = Convert.ToDateTime(dt.Rows[i]["تاريخالفاتورة"].ToString());
            FileText = FileText.Replace("تاريخالفاتورة", now.ToString("MM/yyyy"));
        }

        FileText = InsertRepeatedRow(FileText, "اسمالموظف", dt);

        FileText = FileText.Replace("تاريخاليوم", DateTime.Now.ToString());

        //FileText = FileText.Replace("tafkeetEnlish", toword.ConvertToEnglish());
        Response.ClearContent();

        // LINE1: Add the file name and attachment, which will force the open/cance/save dialog to show, to the header
        string fileName = "Download.doc"; // dt.Rows[0]["اسمالموظف"] + ".doc";
        Response.AddHeader("Content-Disposition", "attachment; filename=" + fileName);

        // Add the file size into the response header
        //  Response.AddHeader("Content-Length", file.Length.ToString());

        // Set the ContentType
        Response.ContentType = "application/ms-word";

        // Write the file into the response (TransmitFile is for ASP.NET 2.0. In ASP.NET 1.1 you have to use WriteFile instead)
        Response.Write(FileText);

        // End the response
        Response.End();
    }
        catch (Exception ex)
        {
            Response.Write(ex.Message);

            // End the response
            Response.End();
        }
    }

    protected async Task BtnDownload_Click(object sender, EventArgs e)
{
    try
    {
        bool HeadOffice = false;
        string empIds = string.Empty;
        string FileText = File.ReadAllText(
            Server.MapPath("~/Templates/ProjectInvoice/Salary-Slip.xml")
        );
        FileText = GlobalCode.ReplaceTempPassword(
            FileText,
            ConfigurationManager.AppSettings["MSWordhash"],
            ConfigurationManager.AppSettings["MSWordsalt"],
            ConfigurationManager.AppSettings["MSWordProviderType"],
            ConfigurationManager.AppSettings["MSWordAlgorithmSid"],
            Request.QueryString["UserID"]
        );


            var projectInvoiceId = Request.QueryString["id"];

            if (!string.IsNullOrEmpty(projectInvoiceId))
            {
                salarySlipQuery += " where new_invoiceBase.new_invoiceId = '@id'";
                salarySlipQuery = salarySlipQuery.Replace("@id", projectInvoiceId);
            }
            else
            {
                salarySlipQuery += "Where ";

                if (string.IsNullOrEmpty(fromdate.Text))
                {
                    salarySlipQuery += "YEAR(new_invoiceBase.new_fromdate) = YEAR(GETDATE())";
                }
                else
                {
                    salarySlipQuery += "new_invoiceBase.new_fromdate >= CONVERT(datetime, '" + fromdate.Text + "')";
                }



                if (string.IsNullOrEmpty(todate.Text))
                {
                    salarySlipQuery += " and YEAR(new_invoiceBase.new_todate) = YEAR(GETDATE())";
                }
                else
                {
                    salarySlipQuery += " and new_invoiceBase.new_todate >= CONVERT(datetime, '" + todate.Text + "')";
                }
            }


            if (!string.IsNullOrEmpty(empsNo.Text))
        {
            empIds = GlobalCode.ConvertLinesToQuoteComma(empsNo.Text);
            salarySlipQuery +=
                "  and (  new_EmployeeBase.new_empidnumber in("
                + empIds
                + ") or new_EmployeeBase.new_borderno in("
                + empIds
                + ") or new_EmployeeBase.new_IDNumber in("
                + empIds
                + ") )";
        }
        DataSet myds = new DataSet();

        //
        myds = CRMAccessDB.SelectQ(salarySlipQuery);
        DataTable dt = myds.Tables[0];
        DateTime now = DateTime.Now;
        if (dt.Rows.Count > 0)
        {
            if (
                dt.Rows[0]["projid"].ToString().ToLower()
                == "C4FA09AF-383D-E711-80C9-000D3AB61E51".ToLower()
            ) //COGNIZANT. Project
            {
                FileText = await File.ReadAllTextAsync(
                    Server.MapPath("~/Templates/ProjectInvoice/Salary-Slip-COGNIZANT.xml")
                );
                FileText = GlobalCode.ReplaceTempPassword(
                    FileText,
                    ConfigurationManager.AppSettings["MSWordhash"],
                    ConfigurationManager.AppSettings["MSWordsalt"],
                    ConfigurationManager.AppSettings["MSWordProviderType"],
                    ConfigurationManager.AppSettings["MSWordAlgorithmSid"],
                    Request.QueryString["UserID"]
                );
            }

            //if (
            //    dt.Rows[0]["projid"].ToString().ToLower()
            //    == "5D1B5CBB-6307-E511-80C6-0050568B1DBC".ToLower()
            //) //head office. Project
            //{
            //    HeadOffice = true;
            //    FileText = File.ReadAllText(
            //        Server.MapPath("~/Templates/ProjectInvoice/Salary-Slip-HeadOffice.xml")
            //    );
            //    FileText = GlobalCode.ReplaceTempPassword(
            //        FileText,
            //        ConfigurationManager.AppSettings["MSWordhash"],
            //        ConfigurationManager.AppSettings["MSWordsalt"],
            //        ConfigurationManager.AppSettings["MSWordProviderType"],
            //        ConfigurationManager.AppSettings["MSWordAlgorithmSid"],
            //        Request.QueryString["UserID"]
            //    );




            //   var projectInvoiceId = Request.QueryString["id"];

            //    if (!string.IsNullOrEmpty(projectInvoiceId))
            //    {
            //        sqlQueryHeadOffice += " where new_invoiceBase.new_invoiceId = '@id'";
            //        sqlQueryHeadOffice = sqlQueryHeadOffice.Replace("@id", projectInvoiceId);
            //    }


            //    myds = new DataSet();
            //    myds = CRMAccessDB.SelectQ(sqlQueryHeadOffice);
            //    dt = myds.Tables[0];
            //}
        }

        if (!HeadOffice)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                now = Convert.ToDateTime(dt.Rows[i]["تاريخالفاتورة"].ToString());
                FileText = FileText.Replace("تاريخالفاتورة", now.ToString("MM/yyyy"));
            }

            FileText = FileText.Replace("تاريخاليوم", DateTime.Now.ToString());
        }
        FileText = InsertRepeatedRow(FileText, "rowxx", dt);
        //FileText = FileText.Replace("tafkeetEnlish", toword.ConvertToEnglish());
        Response.ClearContent();

        // LINE1: Add the file name and attachment, which will force the open/cance/save dialog to show, to the header
        string fileName = "Salary Slip.doc"; // dt.Rows[0]["اسمالموظف"] + ".doc";
        Response.AddHeader("Content-Disposition", "attachment; filename=" + fileName);

        GlobalCode.ConvertDocToPDF(FileText, fileName);

        // Add the file size into the response header
        //  Response.AddHeader("Content-Length", file.Length.ToString());

        // Set the ContentType
        Response.ContentType = "application/ms-word";

        // Write the file into the response (TransmitFile is for ASP.NET 2.0. In ASP.NET 1.1 you have to use WriteFile instead)
        Response.Write(FileText);

        // End the response
        Response.End();

    }
    catch (Exception ex)
    {
        Response.Write(ex.Message);

        // End the response
        Response.End();
    }
}

public static string InsertRepeatedRow(
    string xmlData,
    string WordInRowToReplace,
    System.Data.DataTable dtData
)
{
    XmlDocument doc = new XmlDocument();
    doc.LoadXml(xmlData);
    XmlNode trStart = null;
    XmlNode Table = null;
    XmlNode trEnd = null;
    XmlNodeList nodeList = doc.GetElementsByTagName("w:t");
    for (int i = 0; i < nodeList.Count; i++)
    {
        if (nodeList[i].InnerText.Contains(WordInRowToReplace))
        {
            trStart = GlobalCode.GetParent(nodeList[i], "w:tbl");
            Table = GlobalCode.GetParent(nodeList[i], "w:body");
        }
    }
    Table.RemoveChild(trStart);

    //  XmlDataDocument docstart = new XmlDataDocument();
    //docstart.LoadXml(fileStart);
    XmlNode currentNode = trStart;
    string replaceText = trStart.OuterXml;
    System.Text.StringBuilder rows = new System.Text.StringBuilder();
    string tempRow = "";
    for (int rowIndex = 0; rowIndex < dtData.Rows.Count; rowIndex++)
    {
        XmlNode node = trStart.CloneNode(true);
        for (int i = 0; i < dtData.Columns.Count; i++)
        {
            try
            {
                decimal value = 0;

                if (dtData.Columns[i].DataType == typeof(decimal))
                {
                    decimal.TryParse(dtData.Rows[rowIndex][i].ToString(), out value);
                    tempRow = node.InnerXml;
                    tempRow = tempRow.Replace(
                        "" + dtData.Columns[i].ColumnName,
                        value.ToString("0")
                    );
                    // tempRow = System.Web.HTTPUtility.HTMLDecode(tempRow);
                    node.InnerXml = tempRow;
                }
                else if (
                    dtData.Columns[i].DataType == typeof(DateTime)
                    && dtData.Rows[rowIndex][i].ToString() != ""
                )
                {
                    DateTime currentDate = new DateTime();
                    currentDate = (DateTime)dtData.Rows[rowIndex][i];
                    tempRow = tempRow.Replace(
                        "" + dtData.Columns[i].ColumnName,
                        currentDate.ToString("dd/MM/yyyy")
                    );
                }
                else
                {
                    tempRow = node.InnerXml;
                    tempRow = tempRow.Replace(
                        "" + dtData.Columns[i].ColumnName,
                        dtData.Rows[rowIndex][i]
                            .ToString()
                            .Replace(".000000000000", "")
                            .Replace(".0000000000", "")
                            .Replace(".00000000", "")
                            .Replace(".0000", "")
                            .Replace("'", "&apos;")
                            .Replace("\"", "&quot;")
                            .Replace(">", "&gt;")
                            .Replace("<", "&lt;")
                            .Replace("&", "&amp;")
                    );
                    // tempRow = System.Web.HTTPUtility.HTMLDecode(tempRow);
                    node.InnerXml = tempRow;
                }
            }
            catch (Exception ex) { }
        }

        Table.AppendChild(node);
    }

    return doc.OuterXml;
}

protected void btnUpdate_Click(object sender, EventArgs e)
{
    if (FileUpload1.HasFile)
    {
        ConvertToExcel.ExcelImport excel = new ConvertToExcel.ExcelImport();
        DataTable dt = ConvertToExcel.ExcelImport
            .ImportAnyExcel(FileUpload1.FileName, FileUpload1.PostedFile, true)
            .Tables[0];
        //  ColumnPair keyColumn= new ColumnPair("MawaridNumber","new_onlynumber");
        ColumnPair keyColumn = new ColumnPair("crmID", "new_onlynumber");
        List<ColumnPair> columns = new List<ColumnPair>();
        //columns.Add(new ColumnPair("EMAIL", "EmailAddress"));
        //int rows = excel.GenerateUpdateFromExcelByTempTable(dt, "new_employee", keyColumn, columns);

        columns.Add(new ColumnPair("EMAIL", "emailaddress"));
        columns.Add(new ColumnPair("Mobile", "new_mobileno"));
        DataView dvData = new DataView(dt);
        dvData.RowFilter = "Mobile is not null";
        DataTable NewDT = dvData.Table.Copy();
        int rows = excel.GenerateUpdateFromExcelByTempTable(
            NewDT,
            "new_employee",
            keyColumn,
            columns
        );

        Response.Write("Number of Updated rows is " + rows.ToString());
    }
    else
    {
        Response.Write("please , choose the file first");
    }
}

protected void btnSendToAll_Click(object sender, EventArgs e)
{
    bool HeadOffice = false;
    string empIds = string.Empty;
    string FileText = File.ReadAllText(
        Server.MapPath("~/Templates/ProjectInvoice/Salary-Slip.xml")
    );
    FileText = GlobalCode.ReplaceTempPassword(
        FileText,
        ConfigurationManager.AppSettings["MSWordhash"],
        ConfigurationManager.AppSettings["MSWordsalt"],
        ConfigurationManager.AppSettings["MSWordProviderType"],
        ConfigurationManager.AppSettings["MSWordAlgorithmSid"],
        Request.QueryString["UserID"]
    );



    var projectInvoiceId = Request.QueryString["id"];

    if (!string.IsNullOrEmpty(projectInvoiceId))
    {
        salarySlipQuery += " where new_invoiceBase.new_invoiceId = '@id'";
        salarySlipQuery = salarySlipQuery.Replace("@id", projectInvoiceId);
    }
    else
    {
        if (!string.IsNullOrEmpty(fromdate.Text))
        {
            if (!salarySlipQuery.ToLowerInvariant().Contains("where"))
                salarySlipQuery += "Where ";
            else
                salarySlipQuery += " and ";

            salarySlipQuery += "new_fromdate >= CONVERT(datetime, '" + fromdate.Text + "')";
        }

        if (!string.IsNullOrEmpty(todate.Text))
        {
            if (!salarySlipQuery.ToLowerInvariant().Contains("where"))
                salarySlipQuery += "Where ";
            else
                salarySlipQuery += " and ";

            salarySlipQuery += "new_todate <= CONVERT(datetime, '" + todate.Text + "')";
        }
    }


    //and  new_EmployeeBase.new_empidnumber='20004666'
    //new_empidnumber ,new_borderno,new_IDNumber

    if (!string.IsNullOrEmpty(empsNo.Text))
    {
        empIds = GlobalCode.ConvertLinesToQuoteComma(empsNo.Text);
        salarySlipQuery +=
            "  and (  new_EmployeeBase.new_empidnumber in("
            + empIds
            + ") or new_EmployeeBase.new_borderno in("
            + empIds
            + ") or new_EmployeeBase.new_IDNumber in("
            + empIds
            + ") )";
    }
    DataSet myds = new DataSet();
    myds = CRMAccessDB.SelectQ(salarySlipQuery);
    DataTable dt = myds.Tables[0];
    DateTime now = DateTime.Now;
    if (dt.Rows.Count > 0)
    {
        if (
            dt.Rows[0]["projid"].ToString().ToLower()
            == "C4FA09AF-383D-E711-80C9-000D3AB61E51".ToLower()
        )
        {
            FileText = File.ReadAllText(
                Server.MapPath("~/Templates/ProjectInvoice/Salary-Slip-COGNIZANT.xml")
            );
            FileText = GlobalCode.ReplaceTempPassword(
                FileText,
                ConfigurationManager.AppSettings["MSWordhash"],
                ConfigurationManager.AppSettings["MSWordsalt"],
                ConfigurationManager.AppSettings["MSWordProviderType"],
                ConfigurationManager.AppSettings["MSWordAlgorithmSid"],
                Request.QueryString["UserID"]
            );
        }

        //skip head office
        //if (
        //    dt.Rows[0]["projid"].ToString().ToLower()
        //    == "5D1B5CBB-6307-E511-80C6-0050568B1DBC".ToLower()
        //) //head office. Project
        //{
        //    HeadOffice = true;
        //    FileText = File.ReadAllText(
        //        Server.MapPath("~/Templates/ProjectInvoice/Salary-Slip-HeadOffice.xml")
        //    );
        //    FileText = GlobalCode.ReplaceTempPassword(
        //        FileText,
        //        ConfigurationManager.AppSettings["MSWordhash"],
        //        ConfigurationManager.AppSettings["MSWordsalt"],
        //        ConfigurationManager.AppSettings["MSWordProviderType"],
        //        ConfigurationManager.AppSettings["MSWordAlgorithmSid"],
        //        Request.QueryString["UserID"]
        //    );
        //    //sqlQueryHeadOffice += " where new_invoiceBase.new_invoiceId = '@id'";
        //    //sqlQueryHeadOffice = sqlQueryHeadOffice.Replace("@id", Request.QueryString["id"]);


        //    if (!string.IsNullOrEmpty(empsNo.Text))
        //    {
        //        empIds = GlobalCode.ConvertLinesToQuoteComma(empsNo.Text);
        //        sqlQueryHeadOffice +=
        //            "  and (  new_EmployeeBase.new_empidnumber in("
        //            + empIds
        //            + ") or new_EmployeeBase.new_borderno in("
        //            + empIds
        //            + ") or new_EmployeeBase.new_IDNumber in("
        //            + empIds
        //            + ") )";
        //    }
        //    myds = new DataSet();
        //    myds = CRMAccessDB.SelectQ(sqlQueryHeadOffice);
        //    dt = myds.Tables[0];
        //}

    }
    //for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
    for (int i = 0; i < dt.Rows.Count; i++)
    {
        string tempText = FileText;
        // now = Convert.ToDateTime(dt.Rows[rowIndex]["تاريخالفاتورة"].ToString());
        if (!HeadOffice)
            tempText = tempText.Replace("تاريخالفاتورة", now.ToString("MM/yyyy"));
        //for (int i = 0; i < dt.Columns.Count; i++)
        //{
        // tempText = tempText.Replace("" + dt.Columns[i].ColumnName, dt.Rows[rowIndex][i].ToString().Replace(".0000000000", "").Replace(".00000000", "").Replace(".0000", ""));
        string MailContent =
            @"
<div style='border: 1px solid #999; padding: 10px; margin-bottom: 20px;'><span
        style='font-size: 40px;float: left;line-height: 1;margin-right: 10px;color: #68adff;'>✉</span> 
   "
            + mailcontent.Text
            + @"
</div>

<div style='text-align:center'>
    <p><strong><u>SALARY SLIP</u></strong></p>
    <p><strong><u>Month</u></strong><u>:</u> "
            + dt.Rows[i]["slipMonth"]
            + @"
        <strong>|</strong><strong><u>Year</u></strong><u>:</u> "
            + dt.Rows[i]["slipYear"]
            + @"
    </p>
</div>
<table width='100%'
    style='border-collapse: collapse; border-spacing: 0; background-color: #eee; border: 1px solid #999;'
    bgcolor='#eee'>
    <tbody>

        <tr>
            <td width='20%' style='padding: 3px; background-color: rgb(149,179,215);' bgcolor='rgb(149,179,215)'>
                <p><strong>ABDAL EMP CODE:</strong></p>
            </td>
            <td width='30%' style='padding: 3px;'>
                <p><strong>"
            + dt.Rows[i]["employeenumber"]
            + @"</strong></p>
            </td>
            <td width='20%' style='padding: 3px; background-color: rgb(149,179,215);' bgcolor='rgb(149,179,215)'>
                <p><strong>Designation:</strong></p>
            </td>
            <td width='30%' style='padding: 3px;'>
                <p>"
            + dt.Rows[i]["ProfX"]
            + @"</p>
            </td>
        </tr>
        <tr>
            <td width='20%' style='padding: 3px; background-color: rgb(149,179,215);' bgcolor='rgb(149,179,215)'>
                <p><strong>EMP Name :</strong></p>
            </td>
            <td width='30%' style='padding: 3px;'>
                <p>"
            + dt.Rows[i]["اسمالموظف"]
            + @"</p>
            </td>
            <td width='20%' style='padding: 3px; background-color: rgb(149,179,215);' bgcolor='rgb(149,179,215)'>
                <p><strong>Currency</strong><strong>:</strong></p>
            </td>
            <td width='30%' style='padding: 3px;'>
                <p> SAR </p>
            </td>
        </tr>
        <tr>
            <td width='20%' style='padding: 3px; background-color: rgb(149,179,215);' bgcolor='rgb(149,179,215)'>
                <p><strong>PROJECT:</strong></p>
            </td>
            <td width='30%' style='padding: 3px;'>
                <p>"
            + dt.Rows[i]["اسمالمشروع"]
            + @"</p>
            </td>
            <td width='20%' style='padding: 3px; background-color: rgb(149,179,215);' bgcolor='rgb(149,179,215)'>
            </td>
            <td width='30%' style='padding: 3px;'>
                <p></p>
            </td>
        </tr>
    </tbody>
</table>
<br>
<table style='border-collapse: collapse; border-spacing: 0; background-color: #eee; border: 1px solid #999;'
    width='100%' bgcolor='#eee'>
    <tbody>
        <tr>
            <td width='20%' style='padding: 3px; background-color: rgb(149,179,215);' bgcolor='rgb(149,179,215)'>
                <p><strong>Salary Information: </strong></p>
            </td>
            <td width='20%' style='padding: 3px;'>
                <p><strong>Total Salary: </strong></p>
            </td>
            <td style='padding: 3px;'>
                <p>"
            + dt.Rows[i]["xsal"]
            + @"</p>
            </td>
        </tr>
    </tbody>
</table>
<br>
<table width='100%'
    style='border-collapse: collapse; border-spacing: 0; background-color: #eee; border: 1px solid #999;'
    class='bordered' bgcolor='#eee'>
    <tbody>
        <tr>
            <td style='padding: 3px; border: 1px solid #999; text-align: center; background-color: rgb(149,179,215);'
                colspan='2' width='50%' align='center' bgcolor='rgb(149,179,215)'>
                <p><strong>Credit</strong></p>
            </td>
            <td style='padding: 3px; border: 1px solid #999; text-align: center; background-color: rgb(149,179,215);'
                colspan='2' width='50%' align='center' bgcolor='rgb(149,179,215)'>
                <p><strong>Debit</strong></p>
            </td>
        </tr>
        <tr>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p><strong>Working Days:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["workDays"]), 2).ToString()
            + @"</p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p><strong>Absence</strong><strong>Amount:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["abcded"]), 2).ToString()
            + @"</p>
            </td>
        </tr>
        <tr>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p><strong>Due Salary:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["dusalry"]), 2).ToString()
            + @"</p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p><strong>Company Loans:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["emploan"]), 2).ToString()
            + @"</p>
            </td>
        </tr>
        <tr>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p><strong>Over-Time:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["ovrTimeAllow"]), 2).ToString()
            + @"</p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p><strong>Customer Loans:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["loanamount"]), 2).ToString()
            + @"</p>
            </td>
        </tr>
        <tr>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p><strong>Other Increase:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["othIncrease"]), 2).ToString()
            + @"</p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p><strong>Other Deductions:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999;'>
                <p>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["otherdeduction"]), 2).ToString()
            + @"</p>
            </td>
        </tr>
        <tr>
            <td width='25%' style='padding: 3px; border: 1px solid #999; background-color: rgb(149,179,215);'
                bgcolor='rgb(149,179,215)'>
                <p><strong>Total Credit:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999; background-color: rgb(149,179,215);'
                bgcolor='rgb(149,179,215)'>
                <p><strong>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["totCredit"]), 2).ToString()
            + @"</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999; background-color: rgb(149,179,215);'
                bgcolor='rgb(149,179,215)'>
                <p><strong>Total Debit:</strong></p>
            </td>
            <td width='25%' style='padding: 3px; border: 1px solid #999; background-color: rgb(149,179,215);'
                bgcolor='rgb(149,179,215)'>
                <p><strong>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["totDebit"]), 2).ToString()
            + @"</strong></p>
            </td>
        </tr>
    </tbody>
</table>
<br>
<table width='100%'
    style='border-collapse: collapse; border-spacing: 0; background-color: #eee; border: 1px solid #999;'
    bgcolor='#eee'>
    <tbody>
        <tr>
            <td style='text-align: center; padding: 15px;' width='100%' align='center'>
                <p><strong>Net Salary</strong> = Total Credit &ndash; Total Debit =
                    <strong>"
            + Math.Round(Convert.ToDecimal(dt.Rows[i]["netSalary"]), 2).ToString()
            + @"</strong></p>
            </td>
        </tr>
    </tbody>
</table>

<p style='padding: 5px;'>Generated by system.</p>

";

        string cc = "";
        //string mailTo = GlobalCode.GetTeamEmails("7d868e34-6315-eb11-a838-000d3abaded5"/*Timesheet Notification Team*/);
        //string mailTo = "m.hashish@excprotection.com";
        string mailTo = dt.Rows[i]["EmailAddress"].ToString();
        try
        {
            if (mailTo != "")
            {
                MailSender.SendEmail02(mailTo, cc, "Salary Slip", MailContent, true, "");
            }
        }
        catch (Exception ex) { }

        //  }

        #region OldCode3
        //string tempFile = System.IO.Path.GetTempPath() + "/SalaryPaySlip" + dt.Rows[rowIndex]["اسمالموظف"].ToString() + ".doc";
        //System.IO.File.WriteAllText(tempFile, tempText);

        //string EmailAddress = dt.Rows[rowIndex]["EmailAddress"].ToString();

        //tempFile = ConvertDocToPDF(tempText, "salaryslip-" + dt.Rows[rowIndex]["slipMonth"].ToString() + "-" + dt.Rows[rowIndex]["slipYear"].ToString()) + ".pdf";
        //try
        //{

        //    string Cogbody = "Greetings of the day !!! <br/>  It's our pleasure to serve you from Esad recruitment company,<br/>  Please find attached pay slip for your kind reference and record.<br/>  We are ready for your kind support at any time .<br/> Thank you,,";
        //    if (HeadOffice)
        //    {
        //        Cogbody = @"<div style='dir:rtl;text-align:right;color:#003a93;'>عزيزي منسوب شركة راحة الأسرة<br/><br>نبارك لك جهدك و ثمرة عطائك<br/>" +
        //            "مرفق لكم تفاصيل راتب شهر: " + dt.Rows[rowIndex]["slipMonthArabic"].ToString() + " " + dt.Rows[rowIndex]["slipYear"].ToString() +
        //            "<br/><br/> مع تمنياتنا لك بدوام التوفيق ،،<br/>اساد .. راحة بال</div><hr/>";
        //        Cogbody += @"<div style='color:#003a93;'>Dear Esad Employee<br/><br/>Thank you for your effort and good work<br/>" +
        //            "Attached for you Salary Payslip for month: " + dt.Rows[rowIndex]["slipMonth"].ToString() + " " + dt.Rows[rowIndex]["slipYear"].ToString() + "<br/><br/> Best Wishes,,</div>";
        //    }
        //    //MailSender.SendEmail02(EmailAddress, ccmail.Text, "Salary Payslip for " + now.ToString("MM/yyyy"), "Salary Payslip for " + now.ToString("MM/yyyy"), false, tempFile);
        //    if (!string.IsNullOrEmpty(EmailAddress))
        //        MailSender.SendEmail02(EmailAddress, ccmail.Text, "Salary Payslip for " + now.ToString("MM/yyyy"), Cogbody, false, tempFile);
        //    //   MailSender.SendEmail02("mgomaa@tamkeenhr.com", ccmail.Text, "Salary Payslip for " + now.ToString("MM/yyyy"), Cogbody, false, tempFile);

        //}
        //catch (Exception ex)
        //{
        //    throw ex;
        //}
        #endregion
    }
}

protected void Button1_Click(object sender, EventArgs e)
{
    Response.Redirect("EmailUpdate.xlsx");
}

public string ConvertDocToPDF(string FileText, string FileName)
{
    string PdfName = "";

    try
    {
        System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(
            HttpContext.Current.Server.MapPath("~")
        );
        dir = dir.Parent;
        dir = new System.IO.DirectoryInfo(dir.FullName + "/PDFFiles/");
        string filePath = @"C:\ISV\PDFFiles\" + DateTime.Now.Ticks.ToString() + FileName;
        string outPutFile = filePath.Replace(".doc", ".pdf").Replace(".xml", ".pdf");
        if (System.IO.File.Exists(filePath))
        {
            System.IO.File.Delete(filePath);
        }

        if (System.IO.File.Exists(outPutFile))
        {
            System.IO.File.Delete(outPutFile);
        }
        System.IO.File.WriteAllText(filePath, FileText);

        // Create an instance of Word.exe
        Microsoft.Office.Interop.Word._Application oWord =
            new Microsoft.Office.Interop.Word.Application();

        // Make this instance of word invisible (Can still see it in the taskmgr).
        oWord.Visible = false;

        // Interop requires objects.
        object oMissing = System.Reflection.Missing.Value;
        object isVisible = true;
        object readOnly = false;
        object oInput = filePath;
        object oOutput = outPutFile;
        PdfName = outPutFile;
        object oFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;

        // Load a document into our instance of word.exe
        Microsoft.Office.Interop.Word._Document oDoc = oWord.Documents.Open(
            ref oInput,
            ref oMissing,
            ref readOnly,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref isVisible,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing
        );

        // Make this document the active document.
        oDoc.Activate();

        // Save this document in Word 2003 format.
        oDoc.SaveAs(
            ref oOutput,
            ref oFormat,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing,
            ref oMissing
        );

        oWord.Quit(ref oMissing, ref oMissing, ref oMissing);

        System.IO.FileInfo file = new System.IO.FileInfo(outPutFile);
        return outPutFile;
    }
    catch (Exception ex)
    {
        return null;
    }
}
}
