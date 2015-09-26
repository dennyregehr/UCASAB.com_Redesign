<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" CodeFile="Music.aspx.vb" Inherits="Music" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <link rel="stylesheet" title="Standard" href="ContentFlow/contentflow.css" type="text/css" media="screen" />
    <script language="javascript" type="text/javascript" src="ContentFlow/contentflow.js" load="slideshow">
    </script> 
    <meta property="og:title" content="UCA Student Activities Board Music"/>
    <meta property="og:description"
          content="The University of Central Arkansas loves good music, and the Student Activities Board 
            is proud to bring high quality bands and artists to the UCA campus for the student body and the local community."/>
    <meta property="og:type" content="university"/>
    <meta property="og:image_src" content="http://www.ucasab.com/images/fbSABLogo.jpg"/>
    <meta property="og:image" content="http://www.ucasab.com/images/fbSABLogo.jpg"/>
    <meta property="og:url" content="http://www.ucasab.com/music.aspx"/>
    <meta property="og:site_name" content="University of Central Arkansas Student Activities Board"/>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div id="PageTextHeading">
        Music Events
        <fb:like layout="button_count" show_faces="false" fb_ref="music" colorscheme="light"></fb:like>
    </div>
    <asp:ListView ID="ListView1" runat="server" DataSourceID="ObjectDataSource1">
        <LayoutTemplate>
            <div id="myFantasticFlow" class="ContentFlow">
                <div class="loadIndicator"><div class="indicator"></div></div> 
                <div class="flow">
                    <div id="itemPlaceholder" runat="server">
                    </div>
                </div>
                <div class="globalCaption"></div><div class="scrollbar"><div class="slider"><div class="position"></div></div></div>
            </div>
        </LayoutTemplate>
        <ItemTemplate>
            <div class="item" href='<%# Eval("EventId","Event.aspx?id={0}") %>'>
                <asp:Image ID="Image1" runat="server" CssClass="content" ImageUrl='<%# Eval("ImageURL") %>' />
                <asp:Panel ID="Panel1" runat="server" CssClass="caption">
                    <asp:Literal ID="Literal1" runat="server" Text='<%# Eval("EventName") %>' /><br />
                    <asp:Literal ID="Literal2" runat="server" Text='<%# Eval("StartDate", "{0:dddd, M-d-yyyy}") %>' />
                </asp:Panel>
            </div>
        </ItemTemplate>
        <EmptyDataTemplate>
            <div style="width=100%;text-align:center;padding:80px 20px;color:#FFF;">
                There are no more events this semester.  Check back when the semester is getting ready to begin.
            </div>
        </EmptyDataTemplate>        
    </asp:ListView>
    <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" 
        SelectMethod="GetEventsByEventType"
        TypeName="EventUtil">
        <SelectParameters>
            <asp:Parameter Name="EventTypeId" DefaultValue="1" />
        </SelectParameters>
    </asp:ObjectDataSource>
    
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

