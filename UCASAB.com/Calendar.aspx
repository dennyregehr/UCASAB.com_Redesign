<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" CodeFile="Calendar.aspx.vb" Inherits="Calendar" %>

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

    <div id="ChairBoxLinks">
        <div class="ChairBoxLink" onclick="window.location='music.aspx'" style="background-image:url(Images/Music170.png);"></div>
        <div class="ChairBoxLink" onclick="window.location='movies.aspx'" style="background-image:url(Images/Film170.png);"></div>
        <div class="ChairBoxLink" onclick="window.location='comedy.aspx'" style="background-image:url(Images/Comedy170.png);"></div>
        <div class="ChairBoxLink" onclick="window.location='popculture.aspx'" style="background-image:url(Images/pop-culture170.png);"></div>
        <div class="ChairBoxLink" onclick="window.location='novelty.aspx'" style="background-image:url(Images/novelty170.png);"></div>
    </div>
    <div style="clear:both;min-height:50px;"></div>

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="foot" Runat="Server">
</asp:Content>

