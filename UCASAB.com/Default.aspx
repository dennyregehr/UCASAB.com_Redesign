<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" 
        CodeFile="Default.aspx.vb" Inherits="_Default" EnableViewState="false" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <link rel="stylesheet" title="Standard" href="ContentFlow/contentflow.css" type="text/css" media="screen" />
    <script language="javascript" type="text/javascript" src="ContentFlow/contentflow.js" load="slideshow">
    </script>
    <meta property="og:title" content="University of Central Arkansas Student Activities Board - Fun Lives Here!"/>
    <meta property="og:description"
          content="Come to Where Fun Lives.  The University of Central Arkansas Student Activities Board brings
                entertainment to the UCA campus for the student body and the local community."/>
    <meta property="og:type" content="university"/>
    <meta property="og:image_src" content="http://www.ucasab.com/images/fbSABLogo.jpg"/>
    <meta property="og:image" content="http://www.ucasab.com/images/fbSABLogo.jpg"/>
    <meta property="og:url" content="http://www.ucasab.com/"/>
    <meta property="og:site_name" content="UCA Student Activities Board"/>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

    <div id="PageTextHeading">
        <%= Semester.Description %> <%= Year(Today) %> Activities
        <fb:like layout="button_count" show_faces="false" fb_ref="home" colorscheme="light"></fb:like>
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
            <div class="item" runat="server" href='<%# Eval("EventId","Event.aspx?id={0}") %>'>
                <asp:Image ID="EventImage" runat="server" 
                    CssClass="content" 
                    ImageUrl='<%# Eval("ImageURL") %>' />
                <asp:Panel runat="server" CssClass="caption">
                    <%# Eval("EventName") %><br />
                    <%# Eval("StartDate", "{0:dddd, M-d-yyyy}") %>
                </asp:Panel>
            </div>
        </ItemTemplate>
        <EmptyDataTemplate>
            <div style="width:100%;text-align:center;padding:80px 20px;color:#FFF;">
                There are no more events this semester.  Check back when the semester is getting ready to begin.
            </div>
        </EmptyDataTemplate>        
    </asp:ListView>
    <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" 
        SelectMethod="GetRemainingEvents"
        TypeName="EventUtil">
        <SelectParameters>
            <asp:Parameter Name="AddPreviousEventsAtEnd" DefaultValue="true" />
            <asp:Parameter Name="CommitteeEventsOnly" DefaultValue="true" />
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
