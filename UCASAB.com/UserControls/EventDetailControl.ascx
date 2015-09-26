<%@ Control Language="VB" AutoEventWireup="false" CodeFile="EventDetailControl.ascx.vb"
    Inherits="UserControls_EventDetailControl" EnableViewState="false" %>

<asp:Panel ID="pnlEventImage" runat="server" Style="float: left; margin:20px;">
    <asp:Image ID="imgEventImage" runat="server" />    
</asp:Panel>
<div style="float:right;">
    <fb:like layout="button_count" show_faces="false" fb_ref="<%= EventDetailClass.EventId %>"></fb:like>
</div>
<div style="font-size:large;line-height:1.4em;margin:0 15px;">
    <h2>
    <asp:Label ID="lblEventName" runat="server"></asp:Label>
    </h2>
    <h2>
    <asp:Label ID="lblStartDate" runat="server"></asp:Label>
    <asp:Label ID="lblEndDate" runat="server"></asp:Label>
    </h2>
    <h2>
    <asp:Label ID="lblStartTime" runat="server"></asp:Label>
    <asp:Label ID="lblEndTime" runat="server"></asp:Label>
    </h2>
    <h3>
    <asp:Label ID="lblLocation" runat="server"></asp:Label>
    </h3>
    <h2>
    <asp:HyperLink ID="lnkWebsite" runat="server" Target="_blank" Text="Website"></asp:HyperLink>
    <asp:HyperLink ID="lnkVideo" runat="server" Target="_blank" Text="Video"></asp:HyperLink>
    <asp:HyperLink ID="lnkAudio1" runat="server" Target="_blank" Text="Audio(1)"></asp:HyperLink>
    <asp:HyperLink ID="lnkAudio2" runat="server" Target="_blank" Text="Audio(2)"></asp:HyperLink>
    </h2>
    <asp:Label ID="lblEventDescription" runat="server"></asp:Label>
    <asp:HyperLink ID="lnkEventDescriptionMore" runat="server" Text="[more]"></asp:HyperLink>
</div>
