<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="StyleSheet.css" rel="stylesheet" />
    <script>
        function getWidth() {
            var width = '<%=this.MyProperty%>';
            return width;
        }
    </script>
    <script src="JavaScript.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        Select some file!
        <asp:FileUpload ID="FileUpload1" runat="server" Width="303px" />
        <br />
        <br />
        Has Header?       
            <asp:RadioButtonList ID="rbHDR" runat="server" BorderStyle="None" BorderWidth="0px">
                <asp:ListItem Text="Yes" Value="Yes" Selected="True"></asp:ListItem>
                <asp:ListItem Text="No" Value="No"></asp:ListItem>
            </asp:RadioButtonList>
        <br />
        <asp:Label ID="Label3" runat="server" Text="Select a sheet!"></asp:Label>
        &nbsp;<asp:Button ID="Button2" runat="server" Height="22px" OnClick="Button2_Click" Text="Click here" />
        <asp:DropDownList ID="DropDownList1" runat="server" Height="22px" Width="173px">
        </asp:DropDownList>
        <br />
        <br />
        <asp:Button ID="PreButtonID" runat="server" OnClick="PreButton" Text="Preview" Width="150px" />
        &nbsp;<asp:Button ID="UploadID" runat="server" Text="Upload" OnClick="UploadID_Click" Width="150px" />
        <br />
        <br />
        <asp:Button ID="ViewDBID" runat="server" Text="View Database" OnClick="ViewDBID_Click" Width="150px" />
        &nbsp;<asp:Button ID="TruncateDBID" runat="server" Text="Truncate Database" OnClick="TruncateDBID_Click" Width="150px" />
        &nbsp;<asp:Button ID="Button1" runat="server" OnClick="Button1_Click1" Text="Test progress bar" Width="150px" />
        <br />
        <br />
        <asp:Label ID="Label2" runat="server" Text="Status!"></asp:Label>
        <br />
        <div style="width: 240px; height: 15px; border: 1px solid black;">
            <div id="myProgress">
                <div id="myBar">
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 0%
                <br />
                </div>
            </div>
            <br />
            <div style="height: 209px; width: 703px; overflow: auto;">
                <asp:GridView ID="GridView1" runat="server" AllowPaging="true" PageSize="10" Height="16px" Width="677px" OnPageIndexChanging="GridView1_PageIndexChanging" />
                <br />
            </div>
        </div>
    </form>
</body>
</html>
