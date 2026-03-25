<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="RCS_Print.aspx.cs" Inherits="RCS_DocumentPrint._RCS_Print" %>

<!DOC<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
    <script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
    <script type="text/javascript">
        $(function () {
            $(".datepicker").datepicker({ dateFormat: 'yymmdd' });
        });
    </script>
    <div>
        <asp:Label ID="Label1" runat="server" Text="Document type:" Width="125px"></asp:Label>
        <asp:DropDownList ID="_DropDownListDocType" runat="server" DataSourceID="XmlDataSource1"
            DataTextField="name" DataValueField="objectCode" Height="29px" Width="258px"
            OnSelectedIndexChanged="_DropDownListDocType_SelectedIndexChanged" AutoPostBack="True"
            CausesValidation="True" OnDataBinding="Page_Load" OnDataBound="Page_Load">
        </asp:DropDownList>
        <br />
        <br />
        <asp:PlaceHolder ID="_PlaceHolder1" runat="server"></asp:PlaceHolder>
        <br />
        <br />
        <asp:Button ID="_ButtonPrint" runat="server" Text="Print dokument" OnClick="_ButtonPrint_Click" />
        <br />
        <br />
    </div>
    <div>
        <asp:Label ID="_Label_Message" runat="server" Text=""></asp:Label>
    </div>
    <div>
        <asp:PlaceHolder ID="_PlaceHolder2" runat="server"></asp:PlaceHolder>
    </div>
    <asp:XmlDataSource ID="XmlDataSource1" runat="server" DataFile="~/PrintObjectsList.xml">
    </asp:XmlDataSource>
    </form>
</body>
</html>
