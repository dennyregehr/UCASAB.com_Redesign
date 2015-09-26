<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" CodeFile="Join.aspx.vb" Inherits="Join" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div id="PageTextHeading">
        Become the Fun - Join the Student Activities Board
    </div>
    <div id="JoinUsContainer">
        <h3 style="padding:0 20px;">
            If you are interested in joining SAB, <br />
            we would <em>LOVE</em> to meet you!
            <br /><br />
            You may pick up an application at either of these locations:
        </h3>
        <br />
        <div>
            <h3>
                Student Activities Board Office
                <br />
                in Student Center, Room 207
            </h3>
        </div>
        <div>
            <h3>
                Director of Student Activities Office
                <br />
                in Student Center, Room 206
            </h3>
        </div>
        <div style="border-color:#0086e0;">
            <h3>
                Or Just<br /><br />
                <asp:LinkButton ID="btnSabApplication" runat="server" Text="Download It Here" />
                <br /><br />
            </h3>
        </div>
        <br style="clear:both;" />
    </div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="foot" Runat="Server">
</asp:Content>

