var timerForInvoiceDetailsGrid;

function onloadHeader() {
    BindOnChange('new_candelete', CanDeleteChanged);
    SetIframeSrc("IFRAME_Report", getISVWebUrl() + "Invoice/ShortInvoiceReport.aspx" + GetCurrentUrlParamters());

    SetDisabledWithSaving("new_invoicenumber");
    //فريق المسح 

    //546E6F82-CDB7-E711-80D4-000D3AB61E51|| IsAdministrator()
    //فريق اعتماد وتعديل الفواتير
    //d606cc44-fc25-e411-9dce-00505686037e



    //فريق إعداد الرواتب

    //if (checkCurrentUserInTeam("D2BC69A5-9F23-E411-9DCE-00505686037E")) {
       // setEnabled("new_generateempsalary");
    //} else {
       // setDisabled("new_generateempsalary");
    //}

    //فريق إعتماد الفواتير
    if (checkCurrentUserInTeam("D606CC44-FC25-E411-9DCE-00505686037E")) {
        setEnabled("new_approved");
    } else {
        setDisabled("new_approved");
    }

    //فريق مسير الرواتب
   // if (checkCurrentUserInTeam("D2BC69A5-9F23-E411-9DCE-00505686037E")) {
       // setEnabled("new_generateempsalary");
    //} else {
        //setDisabled("new_generateempsalary");
    //}

    //فريق دفع الرواتب
    if (checkCurrentUserInTeam("E205E6C0-9F23-E411-9DCE-00505686037E")) {
        setEnabled("new_paidsalary");
    } else {
        setDisabled("new_paidsalary");
    }

    // BindSubgridRefresh("invoicedetail", "new_invoicdetail", "new_invoiceId", "new_totalnetamount", "new_totalamount");
    SetDisabledWithSaving("new_totalamount");
    SetDisabledWithSaving("new_unappliedamount");

    if (IsCreateForm()) {
        SetValueAttribute("new_monthdays", 30);
    }

    //Set Workflow Controls
    //set mandatory for each

    setToggleMandatory("new_senttofinance", 1, "new_senttofinancedate");
    setToggleMandatory("new_revisedfromfinance", 1, "new_reviedfromfinancedate");
    setToggleMandatory("new_senttocustomer", 1, "new_senttocustomerdate");
    setToggleMandatory("new_partialcollect", 1, "new_partialcollectdate");
    setToggleMandatory("new_fullycollect", 1, "new_fullycollectdate");


    //set datevalue for each
    BindCheckWithDate("new_senttofinance", 1, "new_senttofinancedate", "statuscode", 279640000, 1);
    BindCheckWithDate("new_revisingfromfinance", 1, "", "statuscode", 279640001, 279640000);
    BindCheckWithDate("new_revisedfromfinance", 1, "new_reviedfromfinancedate", "statuscode", 279640005, 279640001);
    BindCheckWithDate("new_senttocustomer", 1, "new_senttocustomerdate", "statuscode", 279640002, 279640005);
    BindCheckWithDate("new_partialcollect", 1, "new_partialcollectdate", "statuscode", 279640003, 279640002);
    BindCheckWithDate("new_fullycollect", 1, "new_fullycollectdate", "statuscode", 279640004, 279640003);




    //set disabled process
    SetDisableWhen("new_fullycollect", 1, "new_fullycollectdate,new_partialcollect,new_partialcollectdate,new_senttocustomer,new_senttocustomerdate,new_revisedfromfinance,new_reviedfromfinancedate,new_revisingfromfinance,new_senttofinance,new_senttofinancedate");
    SetDisableWhen("new_partialcollect", 1, "new_partialcollectdate,new_senttocustomer,new_senttocustomerdate,new_revisedfromfinance,new_reviedfromfinancedate,new_revisingfromfinance,new_senttofinance,new_senttofinancedate");
    SetDisableWhen("new_senttocustomer", 1, "new_senttocustomerdate,new_revisedfromfinance,new_reviedfromfinancedate,new_revisingfromfinance,new_senttofinance,new_senttofinancedate");
    SetDisableWhen("new_revisedfromfinance", 1, "new_reviedfromfinancedate,new_revisingfromfinance,new_senttofinance,new_senttofinancedate");
    SetDisableWhen("new_revisingfromfinance", 1, "new_senttofinance,new_senttofinancedate");
    SetDisableWhen("new_senttofinance", 1, "new_senttofinancedate");
    if (!IsAdministrator()) {
        SetDisabledWithSaving("statuscode");
        SetDisabledWithSaving("new_workflowinvoice");

    }

    if (IsCreateForm()) {
        //SetStartEndDate(new Date());
    }

    BindOnChange("new_fromdate", function() { SetStartEndDate(GetValueAttribute("new_fromdate"), false); });

    AddSalaryView();
    if (GetValueAttribute("new_workflowinvoice") == 5) //accept invoice
    {
        disableFormFields(true);
    }
    if (checkCurrentUserInTeam("4CD9C675-303C-E811-80F3-000D3AB61E51")) { //تعديل الفواتير
        setEnabled("new_candelete");
    } else {
        setDisabled("new_candelete");
    }


    if (checkCurrentUserInTeam("1F33324D-CA3B-E811-80F3-000D3AB61E51")) //فريق تعديل خيار الارسال الى المالية 
    {
        setEnabled("new_senttofinance");
        setEnabled("new_salaryissendtofinance");
        setEnabled("new_offsetaccount");

    }

    // checkHOInvoice();
    checkProjectFees();
    setDisabled("new_invoiceproblems");
}

function checkProjectFees() {

    var ProjectID = getLookUpValue("new_projectid");

    //Individual
    if (ProjectID == "{E8E4C329-472F-E311-80B0-00155D010303}") {
     //   SetVisible("new_totalamount", false);
        //SetVisible("new_netsalarywithvat", false);
    } else {
        //SetVisible("new_totalamount", true);
        //SetVisible("new_netsalarywithvat", true);
    }
}

function onLoadDetail() {
    debugger;
    if (GetValueAttribute("new_othercommission") == null) {
        SetValueAttribute("new_othercommission", 0);
    }
    if (GetValueAttribute("new_salesprice") == null) {
        SetValueAttribute("new_salesprice", 0);
    }
    if (GetValueAttribute("new_otherdeduction") == null) {
        SetValueAttribute("new_otherdeduction", 0);
    }
    if (GetValueAttribute("new_overtime") == null) {
        SetValueAttribute("new_overtime", 0);
    }
    if (GetValueAttribute("new_otherdeduction") == null) {
        SetValueAttribute("new_otherdeduction", 0);
    }
    if (GetValueAttribute("new_absencedeductions") == null) {
        SetValueAttribute("new_absencedeductions", 0);
    }

    //Ahmed
    //الصافي المطلوب للموظف الفعلي من الموارد
    // NetEmployeeSalary=(TotalEmployeeSalary*WorkingDays/MonthDays)+Earnings-anydecuctions(AbscenceValue,OtherDeduction,MawaridLoan,CustomerLoan,internaldeduction)
    //Earning=new_overtime+new_othercommission
    //deductions =new_absencedeductions+new_otherdeduction
    //كل الخصومات على الموظف الداخلية والخارجية مع الزيادات 
    //CalculateField("new_nettoemployee", "(new_totalsalary*new_workingdays/new_monthdays)
    //+new_overtime+new_othercommission-new_otherdeduction-new_absencedeductions-new_loanamount-new_internaldeduction-//new_customerloan");
    //hani update
    // CalculateField("new_nettoemployee", "(new_totalsalary*new_workingdays/new_monthdays)+new_overtime+new_bonus+new_othercommission-new_otherdeduction-new_absencedeductions-new_loanamount-new_internaldeduction-new_customerloan+new_missionallownce+new_hardworkallowance-new_hradeduction-new_employeegossi");
    // //hani update

    // //CalculateField("new_salesprice", "new_invoicesalary+new_dueamount");
    // //hani update

    // //راتب الفاتورة
    // //راتب الفاتورة هو الراتب الذي سيظهر القيمة المستحقة للموظف للعميل
    // // NetEmployeeSalary=(TotalEmployeeSalary*WorkingDays/MonthDays)+Earnings-anydecuctions(AbscenceValue,OtherDeduction,CustomerLoan)
    // //Earning=new_overtime+new_othercommission
    // //deductions =new_absencedeductions+new_otherdeduction
    // //كل الخصومات والزيادات من العميل فقط وبالتالي نحذف سلفة الموارد وخصومات داخلية من المعادلة


    // CalculateField("new_invoicesalary", "(new_custsalary*new_workingdays/new_monthdays)+new_overtime+new_othercommission-new_otherdeduction-new_absencedeductions-new_customerloan+new_missionallownce+new_hardworkallowance-new_hradeduction");
    // //hani update
    // //الراتب المستحق
    // //قيمة عدد أيام العمل فقط
    // //بدون خصومات اوزيادات
    // CalculateField("new_duesalary", "(new_totalsalary*new_workingdays/new_monthdays)");


    // //CalculateField("new_totalnetamount", "(new_salesprice*new_mawariddays/new_monthdays)+new_overtime+new_othercommission-new_otherdeduction-new_absencedeductions+new_paymawaridfees-new_deductforcust+new_gossifees-(new_totalsalary*(new_mawariddays-new_workingdays)/new_monthdays)");

    // //new_netfees
    // //CalculateField("new_netfees", "(new_salesprice*new_mawariddays/new_monthdays)+new_overtime+new_othercommission-new_otherdeduction-new_absencedeductions+new_paymawaridfees-new_deductforcust+new_gossifees-(new_totalsalary*(new_mawariddays-new_workingdays)/new_monthdays)");
    // CalculateField("new_netfees", "(new_dueamount*new_mawariddays/new_monthdays)+new_gossifees+new_paymawaridfees+new_othermawaridearning-new_deductforcust-new_othercustdeduction");

    // //القيمة المرسلة للعميل يضاف إليه سلفة العميل 

    // CalculateField("new_totalnetamount", "new_netfees+new_invoicesalary");

    //Tamer

    //alert("new_nettoemployee:" + GetValueAttribute("new_nettoemployee") + "      " + "new_mwfeesdays:" + GetValueAttribute("new_mwfeesdays") + "      ")
    //alert("new_gossifees:" + GetValueAttribute("new_gossifees") + "      " + "new_paymawaridfees:" + GetValueAttribute("new_paymawaridfees") + "      ")
    //alert("new_customerloan:" + GetValueAttribute("new_customerloan") + "      ")
    // CalculateField("new_totalnetamount", "new_nettoemployee+new_mwfeesdays+new_gossifees+new_paymawaridfees-new_deductforcust-(new_totalsalary*(new_mawariddays-new_workingdays)/new_monthdays)-new_othercustdeduction+new_loanamount");
    debugger;
    var inv = GetEntityByGuid("new_invoice", "new_invoiceId", getLookUpValue("new_invoiceid"));

    if (inv != null && inv.results.length > 0 && (inv.results[0].new_WorkFlowInvoice.Value == 5 || inv.results[0].new_candelete == 0)) {
        disableFormFields(true);

    }

    if (checkCurrentUserInTeam("D606CC44-FC25-E411-9DCE-00505686037E")) {
        setEnabled("new_senttofinance");
    }
    setEnabled("new_uploadedtobank");
    setEnabled("new_uploadtobankdate");

}

function AddEmployees() {
    if (getLookUpValue("ownerid") == CurrentUserID() || checkCurrentUserInTeam("626F9499-3F37-E311-A96D-00155D010303")) {

        OpenISVWindow("Invoice/AddEmployeesToInvoice.aspx");
    } else {
        alert("you are not the owner of the invoice please refer to " + getLookUpDisplay("ownerid") + " to modify ");
    }
}

function onChangeStartDate() {
    var fromDate = GetValueAttribute("new_fromdate");
    if (fromDate != null) {
        SetValueAttribute("new_lastmonthdate", fromDate - 10);
    }
}

function CheckProjectInvoiceButtonsAuthorized() {
    if (!checkCurrentUserInTeam('4229EEDD-B1B8-E711-80D4-000D3AB61E51') && GetValueAttribute('new_sector') == '1') {
        alert('Sorry, You Are Not Authorized to open it !!');
        return false;
    }
    return true;
}

function SetStartEndDate(curDat, setFrom) {
    // var curDat = new Date();
    curDat = curDat.setHours(curDat.getHours() + 3)
    var month = curDat.getUTCMonth();
    var day = curDat.getUTCDate();
    var year = curDat.getUTCFullYear();
    var daysInMonth = new Date(year, month, 0).getDate();
    if (setFrom) {
        SetValueAttribute("new_fromdate", new Date(year, month, 1));
    }

    SetValueAttribute("new_todate", new Date(year, month, daysInMonth));
    SetValueAttribute("new_lastmonthdate", GetValueAttribute("new_fromdate") - 10);
}

function OpenProjectWindow() {

    //فريق مسئولو النظام
    //if (getLookUpValue("ownerid") == CurrentUserID() || checkCurrentUserInTeam("626F9499-3F37-E311-A96D-00155D010303")) {
    if (CheckProjectInvoiceButtonsAuthorized()) {
        OpenISVWindow("Invoice/AddEmployeesToInvoice.aspx");
    }

    //else
    //alert("you are not the owner of the invoice please refer to " + getLookUpDisplay("ownerid") + " to modify ");
}

function OpenPrintProjectInvoice() {
    OpenISVWindow("templates/ProjectInvoice/PrintProjectInvoice.aspx");
}

function OpenPrepareSalary() {
    if (CheckProjectInvoiceButtonsAuthorized()) {
        OpenISVWindow("Invoice/PrepareSalary.aspx");
    }
}

function OpenPrintSalary() {
    if (CheckProjectInvoiceButtonsAuthorized()) {
        OpenISVWindow("Templates/ProjectInvoice/PrintSalary.aspx");
    }
}

function AddSalaryView() {
    var control = document.getElementById("{A9338F4E-1872-E311-8F5A-00155D010303}");
    var salryIndv = document.getElementById("{085444E6-64C5-E311-853D-00155D010308}");
    //debugger;
    if (control == null) {
        setTimeout("AddSalaryView()", 1000);
        return;
    } else {
        if (GetValueAttribute("new_approved")) {
            control.style.display = "block";
            salryIndv.style.display = "block";
        } else {
            control.style.display = "none";
            salryIndv.style.display = "none";
        }


    }

}

function ValidateDate() {

    var FromDate = GetValueAttribute("new_fromdate");
    var ToDate = GetValueAttribute("new_todate");

    if (ToDate < FromDate) {
        alert("لابد أن يكون التاريخ أكبر من الحقل (من تاريخ) ,, أرجو تصحيح الخطأ");
        SetValueAttribute("new_todate", "");
    }

}

function DisableDetails() {
    var isenabled = false;
    if (getLookUpValue("ownerid") == CurrentUserID() || checkCurrentUserInTeam("626F9499-3F37-E311-A96D-00155D010303")) {

        isenabled = true;
    }
    if (checkCurrentUserInTeam("D2BC69A5-9F23-E411-9DCE-00505686037E")) {
        isenabled = true
    }

    if (isenabled) {
        disableFormFields(false);
    } else {
        disableFormFields(true);
    }
}

function PrintInvoice() {
        OpenSystemDoc();
}

function CanDeleteChanged() {
    if (GetValueAttribute('new_candelete') == 0) //False
    {
        SetValueAttribute('new_workflowinvoice', 5); //Accepted
    }
    if (GetValueAttribute('new_candelete') == 1) //False
    {
        SetValueAttribute('new_workflowinvoice', 3); //send to finance
    }
}

function checkHOInvoice() {
    debugger;
    var projectId = getLookUpValue("new_projectid");
    //if (projectId == '{5D1B5CBB-6307-E511-80C6-0050568B1DBC}') {
    //    loadHOInvoiceDetails();
    //  timerForInvoiceDetailsGrid = setInterval(function () { loadHOInvoiceDetails(); }, 1000);

}

function loadHOInvoiceDetails() {
    var detailsSubGridCtrl = Xrm.Page.getControl("invoicedetails");
    if (detailsSubGridCtrl != null) {
        // clearInterval(timerForInvoiceDetailsGrid);


        var detailsGrid = detailsSubGridCtrl.getGrid();
        var HoInvoiceId = Xrm.Page.data.entity.getId();
        var detailsGridSelector = detailsSubGridCtrl.getViewSelector();

        var url = 'https://' + window.location.hostname + ":8001/ViewManagment/Update.aspx?Id=" + HoInvoiceId + "&view=HeadofficeCustomerInvoiceDetails";
        CallISVResponseURL(url);

        var HoView = {
            entityType: 10104, // UserQuery
            id: "7267138F-61DC-EA11-A837-000D3A227AB4",
            name: "Headoffice Customer Invoice Details"
        };
        detailsGridSelector.setCurrentView(HoView);
        Xrm.Page.getControl('invoicedetails').getViewSelector().setCurrentView(HoView);


        //detailsGrid.control.Refresh();
    }
}


function sendSalariesToAX() {
    debugger;
    OpenISVWindow("AX/DynaimcAXSending.aspx");
}

function createFinalInvoice() {
    debugger;

    if (GetValueAttribute("new_invoicetype") == 100000002) {
        if (GetValueAttribute("new_approved") == true) {

            if (confirm("Are you sure you want to create a final invoice?")) {
                var url = getISVWebUrl();
                url += "Invoice/FinalProjectInvoice.aspx";
                url += GetCurrentUrlParamters();

                var InProgressOptions = {
                    id: "loading",
                    title: "Final Invoice is being created.....",
                    width: 650

                };
                var AlertBase = new Alert(InProgressOptions);
                AlertBase.showLoading();

                setTimeout(function () {

                    var json = CallISVResponseURL(url)
                    AlertBase.remove();
                    Xrm.Page.data.refresh();
                }, 1000);
            }

        } else {
            alert("You can't create a final invoice as customer didn't approve so.");
        }
    }
    else {
        let invoiceType = GetValueAttribute("new_invoicetype") == 100000000 ? "Final Invoice" : "Payroll Invoice";
        alert(`Can't create final invoice as that invoice is ${invoiceType}`);
    }

}