//LoadEmpScript("/WebResources/new_mainLib.js");
/// <reference path="MainLib.js" />

function setProjectCostCenter() {
    debugger;
    if (getLookUpValue("new_projectid") != null) {
        var projectid = getLookUpValue("new_projectid");

        var projectresults = GetEntityByGuid("new_project", "new_projectId", projectid).results;
        if (projectresults != null && projectresults.length > 0) {
            var costcenterid = projectresults[0].new_costcenterid;
            SetLookUpResult("new_costcenterid", "new_costcenterid", projectresults[0]);
            SaveForm();
        }
    }


}
function GetDateDiff(startDate, endDate) {
    if (startDate && endDate) {
        var one_day = 1000 * 60 * 60 * 24;

        // Convert both dates to milliseconds
        var date1_ms = startDate.getTime();
        var date2_ms = endDate.getTime();

        // Calculate the difference in milliseconds
        var difference_ms = date2_ms - date1_ms;

        // Convert back to days and return
        return Math.round(difference_ms / one_day) + 1;
    }
}

function GetDateDiffmonth(startDate, endDate) {
    if (startDate && endDate) {
        var one_day = 1000 * 60 * 60 * 24;

        // Convert both dates to milliseconds
        var date1_ms = startDate.getTime();
        var date2_ms = endDate.getTime();

        // Calculate the difference in milliseconds
        var difference_ms = date2_ms - date1_ms;

        // Convert back to days and return
        return Math.round(difference_ms / one_day / 30);
    }
}
function LoadEmpScript(path) {
    var newscript = document.createElement('script');
    newscript.type = 'text/javascript';
    newscript.async = true;
    newscript.src = path;
    (document.getElementsByTagName('head')[0] || document.getElementsByTagName('body')[0]).appendChild(newscript);

}
function CalcEmpSalary() {
    if (!isCalculated && (checkCurrentUserInTeam("F6BAB5F2-C38D-E311-900B-00155D010308") || IsAdministrator())) {
        CalculateField("new_totalsalary", "new_basicsalary+new_housingallowance+new_trasnallowance+new_overtimeallowance+new_foodallowance+new_hardworkallowance+new_otherallowance+new_mobileallowance+new_carallowance+new_monthlyincentive");
        isCalculated = true;
    }
    //BindOnChange("new_basicsalary", OnChangeSalaryToSplit);
    //bindHijriDate("new_expirydatehijri", "new_expirydateg");

    //HidaSalarySection();
    BindCheckWithDate("new_passportstatus", 2, "new_passportdeliverydate", "", "", "");
    ChangeFieldBackGround("new_projectjoingdate", "rgb(247, 108, 108)");

}
var isCalculated = false;
function OnChangeSalaryToSplit() {

    //    if (GetValueAttribute("new_overtimeallowance") == null || GetValueAttribute("new_overtimeallowance") == 0) {
    //        if (confirm("هل تريد تقسيم الراتب بناء على توجيهات الإدارة؟")) {
    //            var totalSalary = GetValueAttribute("new_basicsalary");
    //            var basicSalary = totalSalary * 0.66;
    //            var OverTime = totalSalary * 0.11;
    //            var hardAllowance = totalSalary * 0.23;
    //            SetValueAttribute("new_overtimeallowance", OverTime);
    //            SetValueAttribute("new_otherallowance", hardAllowance);
    //            SetValueAttribute("new_basicsalary", basicSalary);
    //        }
    //    }

}



function HidaSalarySection() {

    var audit = document.getElementById("navAudit");
    audit.style.display = "none";

    alert(GetValueAttribute("new_employeetype"));
    if (GetValueAttribute("new_employeetype") == 1) {
        if (IsUserInRole("597C1869-0834-E311-92AC-00155D010303")) {
            SetTabVisibility("contactTab", true);
            var audit = document.getElementById("navAudit");
            audit.style.display = "block";

        }
        else {
            SetTabVisibility("contactTab", false);








        }
    }
    else {
        SetTabVisibility("contactTab", true);
        var audit = document.getElementById("navAudit");
        audit.style.display = "block";
    }







}

function onDrivingLicenceChanged() {
    var isReleased = GetValueAttribute("new_drivinglicence");
    if (isReleased) {
        SetTabVisibility("tab_driving", true);
        SetMandatory("new_dridinglicenceplace");
        SetMandatory("new_drivinglicencestartdate");
        SetMandatory("new_drivinglicenceduration");
        SetMandatory("new_drivinglicenceenddate");
    }
    else {
        SetTabVisibility("tab_driving", false);
        RemoveMandatory("new_dridinglicenceplace");
        RemoveMandatory("new_drivinglicencestartdate");
        RemoveMandatory("new_drivinglicenceduration");
        RemoveMandatory("new_drivinglicenceenddate");
    }
}

function onLicenceDurationChanged() {

    var startDate = GetValueAttribute("new_drivinglicencestartdate");
    var year = parseInt(startDate.getFullYear());
    var month = parseInt(startDate.getMonth());
    var date = parseInt(startDate.getDate());
    var hour = parseInt(startDate.getHours());
    var newDate = null;
    if (startDate != null) {
        var durationControlValue = GetValueAttribute("new_drivinglicenceduration");
        switch (durationControlValue) {
            case 1: //2 years
                newDate = new Date(year + 2, month, date, hour);
                break;
            case 2: //5 years
                newDate = new Date(year + 5, month, date, hour);
                break;
            case 3: //10 years
                newDate = new Date(year + 10, month, date, hour);
                break;
            default:
                break;

        }
        if (newDate != null)
            SetValueAttribute("new_drivinglicenceenddate", newDate);
    }



}

function onProfessionChange() {

    var professionID = getLookUpValue("new_professionid");
    if (professionID.toString() == '{49C7F260-292F-E311-B3FD-00155D010303}' || professionID.toString() == '{BFCEF260-292F-E311-B3FD-00155D010303}') {
        var passport = GetValueAttribute("new_passportnumber");
        if (passport != null) {
            SetMandatory("new_experience");
        }
        else {
            RemoveMandatory("new_experience");
        }
    }

}

function showCard() {
    var url = getISVWebUrl();
    var propstatus = GetValueAttribute("new_propstatus");
    var empType = GetValueAttribute("new_employeetype");
    var empidnumber = GetValueAttribute("new_empidnumber");

    if (checkCurrentUserInTeam("626F9499-3F37-E311-A96D-00155D010303"))

        if (((empType == 3) && ((propstatus == 4) || (propstatus == 10))) || (empidnumber == 'E06308') || (empidnumber == 'E05748') || (empidnumber == 'E05760') || (empidnumber == 'E05862')
        ) {
            //if ((empType == 3) ){


            //  url += "ShowCardIndv.aspx?id=" + GetFormID();
            url += "ShowCard.aspx?id=" + GetFormID();

        }
        else {
            alert("هده العاملة يوجد لديها إقامة ولا يمكن طباعة باركود لها");
            return;
        }
    openNewWindow(url);
}
var workerStatusOldValue = null;
function onEmployeeLoad() {


    setDisabled("new_empidnumber");

    debugger;
    //////////////////////////////////
    var EmpID = GetFormID();
    /* SetValueAttribute("statuscode", 100000000);
     SetValueAttribute("new_statuswithcustomer", 1);
     UpdateRecord("new_employee",      { "new_statuswithcustomer":{Value:1}  }, EmpID)*/
SetMandatory("new_sponseredby");

    /////////////////////////////////
    //SetDisabledWithSaving("new_availableforweb");
    //setDisabled("new_availableforweb");
    // SetMandatory("new_lawcity");
    // SetMandatory("new_housingbuilding");
    BindOnChange("new_projectid", setBusinessUnit);
    debugger;
    workerStatusOldValue = GetValueAttribute("statuscode");
    //SetDisabledWithSaving("statuscode");
    BindOnChange("statuscode", WorkerStatusChanged);
    //  CalculateField("new_balanceofannualleave", "new_meritedvacation-new_consumedvacations");
    if (IsAdministrator()) {
        //  SetVisible("new_oldprojectid", true);
    }
    //   alert(getLookUpValue('new_projectid'));
    // setProjectCostCenter();
    //    SetDisabledWithSaving("new_reservedaccount");
    //   SetDisabledWithSaving("new_reservationsnotes");
    //  setDisabled("new_meritedvacation");
    //setDisabled("new_balanceofannualleave");
    //    setDisabled("new_consumedvacations");
    // BindOnChange("new_projectid", setProjectCostCenter);
    //ChangeFieldBackGround("new_costcenterid", "yellow");
    //فريق الاطلاع على ATM PIN CODE عاملات
    // if (checkCurrentUserInTeam("0D5D9C15-EF48-E611-9420-0050569120F1") && GetValueAttribute("new_employeetype") == 3 && GetValueAttribute("new_sex") == 2) {
    //     //    Xrm.Page.getControl("new_atmpincode").setDisabled(true);
    // }
    // if (getLookUpValue("new_projectid") != '5D1B5CBB-6307-E511-80C6-0050568B1DBC') {
    //     // SetVisible('new_headofficeprofession', false);
    // }
    //by tamer, check with me
    //if (checkCurrentUserInTeam("A7928A4A-645B-E511-80E7-00505691216D") || checkCurrentUserInTeam('CB4EBD52-AB94-E811-8102-000D3AB61E51'))//فريق تعديل حاله الموظف الى تحت اجرارءت نقل الكفالة)
    //{
    //    setEnabled("statuscode");
    //}
    //if (checkCurrentUserInTeam("8D957E73-9D51-E611-9421-0050569120F1") && GetValueAttribute("new_employeetype") == 3) {
    //setEnabled("statuscode");
    //}
    //  EnableSection("SalarySection");
    // if (isHRAuthorized()) {
    //new_totalsalary", "new_basicsalary+new_housingallowance+new_trasnallowance+new_overtimeallowance+new_foodallowance+new_hardworkallowance+new_otherallowance+new_mobileallowance+new_carallowance+new_monthlyincentive
    //setEnabled("new_totalsalary");
    //setEnabled("new_basicsalary");
    //setEnabled("new_housingallowance");
    //setEnabled("new_trasnallowance");
    //setEnabled("new_overtimeallowance");
    //setEnabled("new_foodallowance");
    //setEnabled("new_hardworkallowance");
    //setEnabled("new_otherallowance");
    //setEnabled("new_mobileallowance");
    //setEnabled("new_carallowance");
    //setEnabled("new_monthlyincentive");
    //setEnabled("new_gossisalary");
    //setEnabled("new_gossihousing");

    // }

    //فريق تعديل الراتب فى شاشة الموظف
    if (checkCurrentUserInTeam("01978F3F-BA60-EC11-A831-000D3ABE20F8")) {
        EnableSection("SalarySection");
        setEnabled("new_gossisalary");
        setEnabled("new_gossihousing");
        setEnabled("new_gossiidnumber");
    } else {
        DisableSection("SalarySection");
        //   setDisabled("new_gossisalary");
        //   setDisabled("new_gossihousing");
        //     setDisabled("new_gossiidnumber");
    }

    var projectid = getLookUpValue("new_projectid");

    if (projectid) {

        // tawsel or headoffice
        var isHeadOffice = projectid == "{551CC644-1725-EC11-A82C-000D3ABE20F8}" || projectid == "{5D1B5CBB-6307-E511-80C6-0050568B1DBC}";
        //فريق تعديل الراتب فى شاشة الموظف للكيان الرئيسى
        var enableHOSalaryEdit = checkCurrentUserInTeam("9A8E702D-1192-EC11-A838-000D3ABE20F8");

        // is headoffice project
        if (isHeadOffice) {
            if (enableHOSalaryEdit) {
                EnableSection("SalarySection");
                setEnabled("new_gossisalary");
                setEnabled("new_gossihousing");
                setEnabled("new_gossiidnumber");
            } else {
                DisableSection("SalarySection");
            }

        }
    }

    //if (!checkCurrentUserInTeam("A7928A4A-645B-E511-80E7-00505691216D")) {
    //    //   فريق تعديل حالة الموظف
    //    /*      Xrm.Page.data.entity.attributes.forEach(function (attribute, index) {
    //              var control = Xrm.Page.getControl(attribute.getName());
    //              if (control) {
    //         control.setDisabled(true);
    //              }
    //          });*/

    //    //if (checkCurrentUserInTeam("F9A3F5F1-AACC-E411-A63B-005056866836")) {
    //    //    // فريق الإيواء
    //    //    //     setEnabled("new_deliverydevice");
    //    //    //    setEnabled("new_mobileno");
    //    //    //     SetMandatory("new_deliverydevice");
    //    //    //     SetMandatory("new_mobileno");
    //    //}
    //    // setEnabled("statuscode");
    //    setEnabled("new_projectid");
    //    //  setEnabled("new_branchprojectid");
    //    setEnabled("new_projectjoingdate");
    //}
    // if (checkCurrentUserInTeam("F9A3F5F1-AACC-E411-A63B-005056866836")) {
    // فريق الإيواء
    //Xrm.Page.data.entity.attributes.forEach(function (attribute, index) {    
    //    var control = Xrm.Page.getControl(attribute.getName());
    //    if (control) {
    //        control.setDisabled(true)
    //    }
    //});
    //        setEnabled("new_namearabic");
    //        setEnabled("new_mobileno");      
    //    }

    debugger;
    checkIqamaNumber("new_idnumber");
    checkUniquenessAjax("new_employeeextensionbase", "new_idnumber", "new_idnumber", "رقم الهوية");
    //hideFemaleData();
    //SetDisabledWithSaving("new_pirodcontract");
    /*   if (GetValueAttribute("new_pirodcontract") == null) {
           if (GetValueAttribute("new_contractstartdate") != null && GetValueAttribute("new_contractenddate") != null) {
   
               var months = GetDateDiffmonth(GetValueAttribute("new_contractstartdate"), GetValueAttribute("new_contractenddate"));
   
               SetValueAttribute("new_pirodcontract", months);
           }
       }*/
    //DisableSection("SalarySection");
    //Osama Business accountant || ALi HR || Sultan
    SetTabVisibility("contactTab", true);


    // BindOnChange("new_contractstartdate", SetContractEndDate);
    //   BindOnChange("new_bed", CheckBed);
    // todo uncomment again
   // SetDisabledWithSaving("new_empidnumber");
    //SetMandatory("new_idnumber");

    if (checkCurrentUserInTeam("4BD86E9F-5B49-E711-80CD-000D3AB61E51") && !IsAdministrator()) { //hr team
        SetMandatory("new_lawcontracttype");
        SetMandatory("new_religion");
        SetMandatory("new_idtype");
        SetMandatory("new_borderno");
        SetMandatory("new_sex");
        SetMandatory("new_maritalstatus");
        // SetMandatory("emailaddress");
        SetMandatory("new_directmanager");
        SetMandatory("new_atmstatus");
        SetMandatory("new_insurancecardstatus");
        SetMandatory("new_noticeperioddays");
        SetMandatory("new_expirydate");
        SetMandatory("new_dateofbirth");

        setMendatorySalaryInfo();
    }
    alertforbirthdate();
    setDisabled("statuscode");

    // فريق تعديل بيانات العاملة
    if (!checkCurrentUserInTeam("2D7E150C-A17C-EA11-A82A-000D3ABADED5")) {
        disableFormFields(true);
    }


    // فريق رؤية تكلفة الخدمة
    if (!checkCurrentUserInTeam("6702657D-6B56-EC11-A830-000D3ABE20F8")) {
        debugger;
        SetVisible("new_dueamountpermonth", false);
        SetVisible("new_custsalary", false);
        SetVisible("new_serviceprice", false);
        SetVisible("new_marketingprice", false);
    }

    //فريق تعديل حالة الموظف
    if (checkCurrentUserInTeam("A7928A4A-645B-E511-80E7-00505691216D")) {
        setEnabled("statuscode");
    }



}

function alertforbirthdate() {
    debugger;
    var birthdate = GetValueAttribute("new_dateofbirth");

    if (birthdate != null) {
        var year = birthdate.getFullYear();
        var month = birthdate.getMonth() + 1;
        var day = birthdate.getDate();
        var today = new Date();
        var resutlt = today.getFullYear() - year;
        if (today.getMonth() < month || (today.getMonth() == month && today.getDate() < day)) {
            resutlt--;
        }

        if (resutlt < 19) {
            alert("This employee is under 19 years old \n هذا الموظف عمره أقل من 19 عامًا");
        }
    }
}
function setMendatorySalaryInfo() {
    // Salary Info
    SetMandatory("new_contractstartdate");
    SetMandatory("new_annualleave");
    SetMandatory("new_contractenddate");
    SetMandatory("new_priodcontract");
    SetMandatory("new_basicsalary");
    SetMandatory("new_foodallowance");
    SetMandatory("new_housingallowance");
    SetMandatory("new_trasnallowance");
    SetMandatory("new_otherallowance");
    SetMandatory("new_overtimeallowance");
    SetMandatory("transactioncurrencyid");
    SetMandatory("new_pirodcontract");
    SetMandatory("new_mobileallowance");
    SetMandatory("new_carallowance");
    SetMandatory("new_monthlyincentive");
    SetMandatory("new_hardworkallowance");
    SetMandatory("new_totalsalary");
    SetMandatory("new_probationperioddays");
    SetMandatory("new_esbpenalty");
    // SetMandatory("new_insurancejoindate");
}

function WorkerStatusChanged() {
    if (workerStatusOldValue == GetValueAttribute("statuscode")) return;
    if (checkCurrentUserInTeam('CB4EBD52-AB94-E811-8102-000D3AB61E51'))//فريق تعديل حاله الموظف الى تحت اجرارءت نقل الكفالة

    {
        /*  if (GetValueAttribute("statuscode") != '279640027')//الحالة لا تساوى تحت اجراءات نقل الكفالة
          {
              alert("غير مسموح لك بتغير حالة الموظف الا لحالة تحت اجراءات نقل الكفالة");
              SetValueAttribute("statuscode", workerStatusOldValue);
              return;
          }*/
    }

}

function hideFemaleData() {
    if (!IsCreateForm()) {
        if (GetValueAttribute("new_employeetype") == 1 && GetValueAttribute("new_gender") == 2) {
            $("#new_mobileno").hide();
            $("#FormView1_imgResource").hide();
        }
    }

}

function OpenPrepare() {
    if (checkCurrentUserInTeam("E14E1B3C-6F91-E311-9C82-00155D010308")) {
        var PrjectOwner = GetEntityByGuid("new_project", "new_projectId", getLookUpValue("new_projectid")).results[0];
        if (CurrentUserID() == "{" + PrjectOwner.OwnerId.Id.toUpperCase() + "}") {
            var url = getISVWebUrl();
            url += "Emp/PrintEmpTemplate.aspx?id=" + GetFormID() + "&UserID=" + CurrentUserID();
            openNewWindow(url);
        } else {
            alert("You don't have a permission to do this procedure");
        }
    } else {
        var url = getISVWebUrl();
        url += "Emp/PrintEmpTemplate.aspx?id=" + GetFormID() + "&UserID=" + CurrentUserID();
        openNewWindow(url);
    }





}

function DefinitionSpeech() {

    if (CurrentUserID() != "{75C54007-4A5B-E311-AD4D-00155D010303}") {
        var url = getISVWebUrl();
        url += "Emp/PrintEmpDefinition.aspx?id=" + GetFormID() + "&UserID=" + CurrentUserID();
        openNewWindow(url);
    }
    else {
        var url = getISVWebUrl();
        url += "Emp/BarcodeInfo.aspx?id=" + GetFormID() + "&UserID=" + CurrentUserID();
        openNewWindow(url);
    }
}

function onIqamaPaymentType() {

    if (GetValueAttribute("new_iqamapaymenttype") == 2) {

        SetValueAttribute("new_iqamareceiveddate", "");
        SetValueAttribute("new_uploadiqmafees", 1);
        SetValueAttribute("new_isiqamapaid", false);
        SetValueAttribute("new_workpermitpaid", false);
        SetValueAttribute("new_isiqamagenerated", false);
        SetValueAttribute("new_iqamadelived", false);
        SetValueAttribute("new_iqamadelivereddate", "");
        SetValueAttribute("new_dataissue", "");
        SetValueAttribute("new_uploadsadad", 1);
        SetValueAttribute("new_iqamapaydate", "");
        SetValueAttribute("new_sadaddate", "");
        SetValueAttribute("new_iqamaissuedate", "");
        SetValueAttribute("new_sadadworkpermitno", "");
        SetValueAttribute("new_deliverdto", "");
        SetValueAttribute("new_iqamaplace", "");

    }
}

function BtnEmpAudit_Click() {
    var url = getISVWebUrl();
    url += "Emp/EmployeeAudit.aspx?id=" + GetFormID();
    openNewWindow(url);
}
function btnEmployeeCost() {
    //var url = getISVWebUrl();
    //url += "Emp/EmployeeCost.aspx?id=" + GetFormID();
    //openNewWindow(url);
    OpenISVWindow("Emp/EmployeeCost.aspx");

}
function employeeActions() {
    var url = getISVWebUrl();
    url += "Emp/EmpActions.aspx?id=" + GetFormID() + "&UID=" + CurrentUserID();
    openNewWindow(url);

}
var ISVURL = "http://" + window.location.hostname + ":8001/";
function OpenEmpServices() {
    // alert('ahmed');

    var url = ISVURL;
    url += "Emp/EmpServices.aspx?UID=" + Xrm.Page.context.getUserId();

    var name = "newWindow";
    var width = 800;
    var height = 600;
    var newWindowFeatures = "status=1;scrollbars=1";

    // CRM function to open a new window
    //window.open(url, "mywindow", "scrollbars=yes,resizable=yes,  width=800,height=600");
    openStdWin(url, name, width, height, newWindowFeatures);
}

function OnChangenew_contractstartdate() {

    if (GetValueAttribute("new_pirodcontract") == null) {
        if (GetValueAttribute("new_contractstartdate") != null && GetValueAttribute("new_contractenddate") != null) {

            var months = GetDateDiffmonth(GetValueAttribute("new_contractstartdate"), GetValueAttribute("new_contractenddate"));

            SetValueAttribute("new_pirodcontract", months);
        }
    }

}

function OnChangenew_contractenddate() {

    if (GetValueAttribute("new_pirodcontract") == null) {
        if (GetValueAttribute("new_contractstartdate") != null && GetValueAttribute("new_contractenddate") != null) {

            var months = GetDateDiffmonth(GetValueAttribute("new_contractstartdate"), GetValueAttribute("new_contractenddate"));

            SetValueAttribute("new_pirodcontract", months);
        }
    }

}


function OnChangepriodcontract() {

    if (GetValueAttribute("new_contractstartdate") != null && GetValueAttribute("new_contractenddate") != null) {

        var days = GetDateDiff(GetValueAttribute("new_contractstartdate"), GetValueAttribute("new_contractenddate"));

        SetValueAttribute("new_priodcontract", days);
    }

}

function IqamaRenewalDate() {
    if (GetValueAttribute("new_expirydatehijri") != null) {
        SetValueAttribute("new_iqamarenewalreceiveddate", new Date());
    } else {
        SetValueAttribute("new_iqamarenewalreceiveddate", null);
    }
}
function isHRAuthorized() {
    var isAurhorized = false;
    //فريق مشرف مشروع خلدا - الموارد
    if (checkCurrentUserInTeam("E14E1B3C-6F91-E311-9C82-00155D010308") || checkCurrentUserInTeam("9B553F98-398B-E411-AC67-005056866836")) {

        if (CurrentUserID() == getLookUpValue("ownerid")) {
            isAurhorized = true;
        }
    }



    //مشرف لمشرفي المشاريع 
    if (IsUserInRole("C909AD20-321C-46B1-8292-0C34FED6F036") || IsUserInRole("1837C5C0-9A8E-E511-80DC-0050568B1DBC") || IsUserInRole("779FAACD-6FF8-E411-80C2-0050568B1DBC")) {
        isAurhorized = true;
    }
    //مسئول النظام
    if (IsAdministrator() || IsUserInRole("46A3AACD-6FF8-E411-80C2-0050568B1DBC")) {
        isAurhorized = true;
        setEnabled("new_meritedvacation");
        setEnabled("new_balanceofannualleave");
        setEnabled("new_consumedvacations");
    }
    //مدير موارد بشرية
    if (IsUserInRole("597C1869-0834-E311-92AC-00155D010303")) {
        isAurhorized = true;
        setEnabled("new_meritedvacation");
        setEnabled("new_balanceofannualleave");
        setEnabled("new_consumedvacations");
    }
    //فريق مدراء الموارد البشرية
    if (checkCurrentUserInTeam("8AF34466-7ADE-E311-859D-00155D010308")) {
        isAurhorized = true;
    }

    if (isAurhorized) {
        SetTabVisibility("contactTab", true);
    }
    else {
        SetTabVisibility("contactTab", false);
    }
    return isAurhorized;

}
function ShowEmpFiles() {
    var url = getISVWebUrl();
    url += "Emp/EmployeeFiles.aspx?id=" + GetFormID();
    openNewWindow(url);
}

function SetContractEndDate() {
     if (GetValueAttribute("new_contractstartdate") == null) {
      SetValueAttribute("new_contractstartdate", GetValueAttribute("new_dateofjoining"));

  }
  if (GetValueAttribute("new_contractstartdate") != null) {
      debugger;

      var ContractDate = new Date(GetValueAttribute("new_contractstartdate"));

      debugger;
      var duration = 1; // year
     // if (getLookUpValue("new_projectid") != '{5D1B5CBB-6307-E511-80C6-0050568B1DBC}')//not head office

     if(GetValueAttribute("new_pirodcontract") == 24) 
          duration = 2;// two years


      var newDate =  new Date(ContractDate.setFullYear(ContractDate.getFullYear()+ duration));
          //newDate.setDate((ContractDate.getDate() + duration));

      SetValueAttribute("new_contractenddate", newDate);

  }
    
}

function setBusinessUnit() {
    if (getLookUpValue("new_projectid") != null) {
        var projectid = getLookUpValue("new_projectid");
        if (projectid == '5D1B5CBB-6307-E511-80C6-0050568B1DBC') {
            SetValueAttribute("new_employeetype", 1);
        }
        else if (projectid == 'F11E1049-3874-E711-80CF-000D3AB61E51') {
            SetValueAttribute("new_employeetype", 4);
        }
    }
}


//////////////////////////////////////////////////////////////////////////////////
//تحديث من تقرير الإقامات الجديدة
function UpdateFromTheNewIqamaReport() {

    if (checkCurrentUserInTeam('d53e0f02-51f0-e711-80e0-000d3ab61e51') || IsAdministrator()) {
        OpenISVWindow("Emp/UpdateNewIqama.aspx");
    }
    else {
        alert("ليس لديك صلاحية");
    }
}

//حديثى الدخول 
function NewlyEntered() {
    if (checkCurrentUserInTeam('d53e0f02-51f0-e711-80e0-000d3ab61e51') || IsAdministrator()) {
        OpenISVWindow("Emp/Muqueem.aspx");
    }
    else {
        alert("ليس لديك صلاحية");
    }
}


//تحديث الهروب والخروج من مقيم
function UpdateEscapeAndExitFromMuqeem() {
    if (checkCurrentUserInTeam('d53e0f02-51f0-e711-80e0-000d3ab61e51') || IsAdministrator()) {
        OpenISVWindow("Emp/MuqeemFinalExit.aspx");
    }
    else {
        alert("ليس لديك صلاحية");
    }
}

//تحديث من تقرير جميع المقيمين 
function UpdateFromTheOfAllResidentsReport() {
    if (checkCurrentUserInTeam('d53e0f02-51f0-e711-80e0-000d3ab61e51') || IsAdministrator()) {
        OpenISVWindow("Emp/Services/UpdateMuqeem.aspx");

    }
    else {
        alert("ليس لديك صلاحية");
    }
}

function CheckBed() {
    var bed = getLookUpValue("new_bed");
    if (bed == null) {
        if (confirm("فى حال مسح السرير سيتم اغلاق امر التسكين . استمرار فى الاجراء؟ - In case you remove the bed from employee you will close the labor housing order of the employee. Continue the procedure?")) {

            //  Process.callWorkflow("9B697463-DA5E-4CEE-B7E8-1FA2F790BC4C", Xrm.Page.data.entity.getId(),
            //function () {
            // alert("Done");
            // },
            //function () {
            //    alert("There is a problem");
            //  });

        }
        else {
            // var saveEvent = context.getEventArgs();
            // saveEvent.preventDefault();
        }
    }
}



function RefreshEmployeeVacations() {
    debugger;
    var EmpId = GetFormID();
    var url = getISVWebUrl() + "Emp/UpdateEmployeeVacationBalance.aspx?EmpId=" + EmpId + "&EntityName=new_employee"
    var res = CallISVResponseURL(url);
    res = JSON.parse(res);
    refreshPage();
}
function UpdateCandidateSkillsFromEmp() {
    debugger;
    if (getLookUpValue("new_candidateild") != null) {
        var candidateId = getLookUpValue("new_candidateild");

        var canCook = GetValueAttribute("new_cancook");
        var cleaning = GetValueAttribute("new_cleaning");
        var canCareOld = GetValueAttribute("new_cancareold");
        var canSpeakEnglish = GetValueAttribute("new_canspeakenglish");
        var weight = GetValueAttribute("new_weight") !=null ? GetValueAttribute("new_weight") : 0.00 ;
        var age = GetValueAttribute("new_age") !=null ? GetValueAttribute("new_age") : 0.00 ;
        var anotherSkill = GetValueAttribute("new_anotherskill");
        var canDoWithChildren = GetValueAttribute("new_candowithchildren");
        var canSpeakArabic = GetValueAttribute("new_canspeakarabic");
        var height = GetValueAttribute("new_tall") !=null ? GetValueAttribute("new_tall") : 0.00 ;

        var record = {
            "new_cancook": canCook, "new_cleaning": cleaning, "new_cancareold": canCareOld,
            "new_canspeakenglish": canSpeakEnglish, "new_weight": parseFloat(weight).toFixed(2), "new_age": parseFloat(age).toFixed(2),
            "new_anotherskill": anotherSkill, "new_candowithchildren": canDoWithChildren,
            "new_canspeakarabic": canSpeakArabic, "new_tall": parseFloat(height).toFixed(2)
        };
        UpdateRecord("new_candidate", record, candidateId);
    }
}

function OpenPrintSalary() {
        var url = getISVWebUrl();
        var empid = GetFormID();
        url += "Templates/ProjectInvoice/PrintSalary.aspx?empid=" + empid;
        openNewWindow(url);
}

//////////////////////////////////////////////////////////////////////////////