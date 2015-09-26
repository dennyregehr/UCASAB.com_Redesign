<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" CodeFile="ExecBoard.aspx.vb" Inherits="ExecBoard" %>

<%@ Register Src="~/UserControls/ChairBoxLinks.ascx" TagPrefix="uc1" TagName="ChairBoxLinks" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

    <div id="PageTextHeading">
        <%= Semester.Description %> <%= Year(Today) %> Exec Board
    </div>
    <br />
    <div class="ExecBoardPage">
        <%--<asp:PlaceHolder ID="ph1" runat="server"></asp:PlaceHolder>--%>
        <div class="ExecBoardMember">
            <img src="Images/scott.jackson.jpg" /><br />
            Scott Jackson<br/>
            President
        </div>
        <div class="ExecBoardMember">
            <img src="Images/jorge.hernandez.jpg" /><br />
            Jorge Hernandez<br/>
            Vice President
        </div>
        <div class="ExecBoardMember">
            <img src="Images/bertita.barrientos.jpg" /><br />
            Bertita Barrientos<br/>
            Media
        </div>
        <div class="ExecBoardMember">
            <img src="Images/kiera.smithton.jpg" /><br />
            Kiera Smithton<br/>
            Movies
        </div>
        <div class="ExecBoardMember">
            <img src="Images/emilia.barrick.jpg" /><br />
            Emilia Barrick<br/>
            Music
        </div>
        <div class="ExecBoardMember">
            <img src="Images/seth.wilson.jpg" /><br />
            Seth Wilson<br/>
            Comedy
        </div>
        <div class="ExecBoardMember">
            <img src="Images/dylan.kimery.jpg" /><br />
            Dylan Kimery<br/>
            Novelty
        </div>
        <div class="ExecBoardMember">
            <img src="Images/javier.hernandez.jpg" /><br />
            Javier Hernandez<br/>
            Pop Culture
        </div>
        <div class="ExecBoardMember">
            <img src="Images/jill.wulfenstein.jpg" /><br />
            Jill Wulfenstein<br/>
            Graduate Assistant
        </div>
    </div>    

    <br style="clear:both;" />

    <uc1:ChairBoxLinks runat="server" ID="ChairBoxLinks1" />
    <div style="clear:both;min-height:50px;"></div>

</asp:Content>
