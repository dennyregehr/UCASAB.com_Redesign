<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" CodeFile="Event.aspx.vb" Inherits="SABEvent" %>

<%@ Register src="~/UserControls/EventDetailControl.ascx" tagname="EventDetailControl" tagprefix="uc1" %>
<%@ Register Src="~/UserControls/SocialMetaTags.ascx" TagName="SocialMetaTags" TagPrefix="uc2" %>
<%@ Register Src="~/UserControls/ChairBoxLinks.ascx" TagName="ChairBoxLinks" TagPrefix="uc3" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <uc2:SocialMetaTags ID="SocialMetaTags1" runat="server" />
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

    <div id="PageTextHeading">
        <asp:Label ID="lblEventName" runat="server"></asp:Label>
    </div>
    <div id="EventDetailContainer">
        <div style="clear:both;min-height:0;"></div>
        <uc1:EventDetailControl ID="EventDetailControl1" runat="server" />
        <div style="clear:both;min-height:0px;"></div>
    </div>

    <uc3:ChairBoxLinks ID="ChairBoxLinksControl1" runat="server"/>
    <div style="clear:both;min-height:50px;"></div>

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="foot" Runat="Server">
</asp:Content>

