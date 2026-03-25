<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SendSMS.aspx.cs" Inherits="RCS_DocumentPrint.SendSMS" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Table ID="Table1" runat="server">
        <asp:TableRow>
        <asp:TableCell VerticalAlign="Top">
               <asp:Label ID="Label1" runat="server" Text=" Telefon nummer" Width="120px"></asp:Label>
        </asp:TableCell>
         <asp:TableCell VerticalAlign="Top">
         
          <asp:TextBox ID="TextBoxNumber" runat="server" Width="286px"></asp:TextBox>
         </asp:TableCell>
        
        </asp:TableRow>
         <asp:TableRow>
        <asp:TableCell VerticalAlign="Top">
              <asp:Label ID="Label3" runat="server" Text="Besked" Width="120px"></asp:Label>

        </asp:TableCell>
         <asp:TableCell VerticalAlign="Top">
         <asp:TextBox ID="TextBoxMessage" runat="server" TextMode="MultiLine" Height="188px" 
            Width="296px"></asp:TextBox>
         </asp:TableCell>
        
        </asp:TableRow>
       
       <asp:TableRow>
        <asp:TableCell VerticalAlign="Top">
           <asp:Button ID="Button1" runat="server" onclick="Button1_Click" 
            Text="Send SMS" />
        </asp:TableCell>
       </asp:TableRow>
       
          <asp:TableRow>
        <asp:TableCell VerticalAlign="Top">
        <asp:Label ID="Label2" runat="server" Text="Resultat" Width="120px"></asp:Label>
            </asp:TableCell>
              <asp:TableCell VerticalAlign="Top">
                  <asp:TextBox ID="TextBoxResult" runat="server" TextMode="MultiLine"
                  Height="188px" 
                  Width="296px"></asp:TextBox>
            </asp:TableCell>
       </asp:TableRow>
        </asp:Table>
    
     
    
        <asp:Button ID="Button2" runat="server" onclick="Button2_Click" 
            Text="Test SMS Functionality" />
    
     
    
    </div>
    </form>
</body>
</html>
