﻿<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" CodeFile="PopCulture.aspx.vb" Inherits="PopCulture" %>

<%@ Register Src="~/UserControls/ChairBoxLinks.ascx" TagPrefix="uc1" TagName="ChairBoxLinks" %>


<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <link rel="stylesheet" title="Standard" href="ContentFlow/contentflow.css" type="text/css" media="screen" />
    <script language="javascript" type="text/javascript" src="ContentFlow/contentflow.js" load="slideshow">
    </script> 
    <meta property="og:title" content="UCA Student Activities Board Pop Culture"/>
    <meta property="og:description"
          content="The University of Central Arkansas is in-the-know.  The Student Activities Board brings speakers ranging from the odd and peculiar to the bold and reckless to the dating gurus who think they've got it all figured out.  It's always different, and it's *always* fun!"/>
    <meta property="og:type" content="university"/>
    <meta property="og:image_src" content="http://www.ucasab.com/images/fbSABLogo.jpg"/>
    <meta property="og:image" content="http://www.ucasab.com/images/fbSABLogo.jpg"/>
    <meta property="og:url" content="http://www.ucasab.com/popculture.aspx"/>
    <meta property="og:site_name" content="University of Central Arkansas Student Activities Board"/>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div id="PageTextHeading">
        Pop Culture Speakers
        <fb:like layout="button_count" show_faces="false" fb_ref="popculture" colorscheme="light"></fb:like>
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
            <div style="width:100%;text-align:center;padding:80px 20px;color:#FFF;">
                There are no more events this semester.  Check back when the semester is getting ready to begin.
            </div>
        </EmptyDataTemplate>        
    </asp:ListView>
    <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" 
        SelectMethod="GetEventsByEventType"
        TypeName="EventUtil">
        <SelectParameters>
            <asp:Parameter Name="EventTypeId" DefaultValue="3" />
        </SelectParameters>
    </asp:ObjectDataSource>
    
    <uc1:ChairBoxLinks runat="server" ID="ChairBoxLinks1" />
    <div style="clear:both;min-height:50px;"></div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="foot" Runat="Server">
</asp:Content>

