<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" CodeFile="ExecBoard.aspx.vb" Inherits="ExecBoard" %>

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
            Brian Thompson<br/>
            <img src="Images/brianThompson.jpg" /><br />
            President
        </div>
        <div class="ExecBoardMember">
            Patrick Moore<br/>
            <img src="Images/patrickMoore.jpg" /><br />
            Vice President
        </div>
        <div class="ExecBoardMember">
            Pharon Williams<br/>
            <img src="Images/pharonWilliams.jpg" /><br />
            Music
        </div>
        <div class="ExecBoardMember">
            Scott Jackson<br/>
            <img src="Images/scottJackson.jpg" /><br />
            Comedy
        </div>
        <div class="ExecBoardMember">
            Tamelah Redden<br/>
            <img src="Images/tamelahRedden.jpg" /><br />
            Novelty
        </div>
        <div class="ExecBoardMember">
            Jorge Hernandez<br/>
            <img src="Images/jorgeHernandez.jpg" /><br />
            Pop Culture
        </div>
        <div class="ExecBoardMember">
            Sydney Spradlin<br/>
            <img src="Images/sydneySpradlin.jpg" /><br />
            Graduate Assistant
        </div>
        <div class="ExecBoardMember">
            Kendra Regehr<br/>
            <img src="Images/director1.jpg" /><br />
            Director
        </div>
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

