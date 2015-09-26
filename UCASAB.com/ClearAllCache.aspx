<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ClearAllCache.aspx.vb" Inherits="ClearAllCache" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:BulletedList ID="BulletedList1" runat="server"></asp:BulletedList>
    </div>
    <div>
        <asp:HyperLink runat="server" NavigateUrl="~/Default.aspx">Home</asp:HyperLink>
    </div>
    </form>
</body>
</html>
