<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" CodeFile="Calendar.aspx.vb" Inherits="Calendar" %>

<%@ Register Src="~/UserControls/ChairBoxLinks.ascx" TagPrefix="uc1" TagName="ChairBoxLinks" %>


<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

    <div id="PageTextHeading">
        <%= Semester.Description %> <%= Year(Today) %> Activities Calendar
    </div>
    <br />
    <div class="CalendarPage">
        <asp:PlaceHolder ID="ph1" runat="server"></asp:PlaceHolder>
    </div>    

    <br style="clear:both;" />

    <uc1:ChairBoxLinks runat="server" ID="ChairBoxLinks1" />
    <div style="clear:both;min-height:50px;"></div>

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="foot" Runat="Server">
</asp:Content>

