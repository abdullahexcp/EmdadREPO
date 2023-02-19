<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PrintSalary.aspx.cs" Inherits="Templates_ProjectInvoice_PrintSalary" Async="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>

    <style type="text/css">
        body {
            background-image: url(/Styles/AbdalLogo.png);
            background-repeat: no-repeat;
            background-attachment: fixed;
            background-position: center;
            background-size: cover;
        }

        .auto-style1 {
            width: 777px;
            text-align: right;
        }

        .auto-style2 {
            font-size: large;
        }

        .auto-style3 {
            width: 241px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div style="background-color: #015287; color: white; padding: .5px; margin-bottom: 10px">
            <h3 style="font-weight: bold; text-align: center">Employee Salary Slip </h3>
        </div>
        <br />
        <div class="container text-center">

            <asp:Button ID="BtnDownload" CssClass="btn btn-success btn-lg" runat="server" OnClick="BtnDownload_Click" Text="Download Salary Slip" />
            &nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Button ID="BtnDownloadVertas" CssClass="btn btn-success btn-lg" runat="server" OnClick="BtnDownloadVertas_Click" Text="Download Salary Slip Vertas" Visible="false" />

            <br />
            <br />
            <table class="table table-bordered table-striped">
                <tr>
                    <td colspan="2">
                        <div style="display: flex; flex-wrap: wrap;">
                            <div style="flex-basis: 50%;">
                                <label for="fromdate">From:</label>
                                <asp:TextBox runat="server" name="fromdate" ID="fromdate" textmode="Date" ></asp:TextBox>
                            </div>
                            <div style="flex-basis: 50%;">
                                <label for="todate">To:</label>
                                <asp:TextBox runat="server" name="todate" ID="todate"  textmode="Date"></asp:TextBox>
                            </div>
                        </div>
                    </td>
                </tr>

                <tr>
                    <td class="auto-style1">
                        <br />
                        <br />
                        <br />
                        Employee No:
                    </td>
                    <td class="text-left">
                        <asp:TextBox ID="empsNo" runat="server" Height="166px" TextMode="MultiLine" Width="371px"></asp:TextBox>

                    </td>
                </tr>
                <tr>
                    <td class="auto-style1">
                        <br />
                        <asp:Label ID="Label1" runat="server" Text="CC Mail:"></asp:Label>
                        <br />
                    </td>
                    <td class="text-left">


                        <asp:TextBox ID="ccmail" runat="server" class="form-control" placeholder="CC Mails" Width="347px"></asp:TextBox>


                    </td>
                </tr>
            </table>
            <table class="table table-bordered table-striped">
                <tr>
                    <td class="auto-style1" style="vertical-align: middle;">
                        <asp:Button ID="btnSendToAll" OnClick="btnSendToAll_Click" CssClass="btn btn-success btn-lg" runat="server" Text="Send Salary Slip To Employee" />
                    </td>
                    <td class="text-left">
                        <asp:TextBox ID="mailcontent" TextMode="multiline" Columns="50" Rows="5" runat="server" class="form-control" placeholder="Message" Height="166px" Width="347px"></asp:TextBox>
                    </td>
                </tr>
            </table>
        </div>

        <br />
        <hr />
        <div class="container text-center">

            <div class="text-left">
                <strong><span class="auto-style2">Update Emails
                </span></strong>
            </div>
            <br />
            <table class="table table-bordered table-striped">
                <tr>
                    <td colspan="2">
                        <asp:Button ID="Button1" runat="server" Text="Download Email Update Template" CssClass="btn btn-success btn-lg" OnClick="Button1_Click" />

                    </td>
                </tr>
                <tr>
                    <td colspan="2"></td>
                </tr>
                <tr>
                    <td class="auto-style3">
                        <div class="text-left">
                            Choose Files :
                        </div>
                        <asp:FileUpload ID="FileUpload1" runat="server" CssClass="" Width="" /></td>
                    <td class="text-left">
                        <asp:Button ID="btnUpdate" runat="server" OnClick="btnUpdate_Click" Text="Update" CssClass="btn btn-success btn-lg" Height="48px" />

                    </td>
                </tr>
            </table>

            &nbsp;&nbsp;&nbsp;&nbsp;
     
      
        </div>


    </form>
    <script>
        const fromDateInput = document.getElementById('fromdate');
        const toDateInput = document.getElementById('todate');

        fromDateInput.value = new Date(new Date().getFullYear(), 0, 2).toISOString().slice(0, 10);
        toDateInput.value = new Date(new Date().getFullYear(), 11, 32).toISOString().slice(0, 10);

    </script>
</body>

</html>
