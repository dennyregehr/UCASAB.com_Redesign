<%@ Page Title="" Language="VB" MasterPageFile="~/Masterpages/MasterPage.master" AutoEventWireup="false" CodeFile="Contact.aspx.vb" Inherits="Contact" %>
<%@ Register Assembly="MSCaptcha" Namespace="MSCaptcha" TagPrefix="cc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div id="PageTextHeading">
        Contact Us
    </div>
    <div id="ContactUsContainer">
        <h3>Message from the SAB Director:</h3>
        <p>
            <img src="Images/director1.jpg" alt="Director of UCA SAB" align="left" style="margin:0 15px;"/>
            We are here to serve the student body of UCA and it's greater area community.
            Anytime you have a question, or need more information about the services
            and events we provide, we are here to help.  Come on by one of our locations, 
            or shoot us an email.
            <br /><br />
            -Kendra Regehr, Director
        </p>
        <br style="clear:both;" />
        <div id="ContactUsEmailForm">
            <asp:Panel ID="pnlMessage" runat="server" 
                Visible="false" 
                CssClass="MailForm">
                <h2>
                    Thanks for your message.<br />
                    Your input makes SAB better for everyone!
                </h2>
            </asp:Panel>
            <asp:Panel ID="pnlErrorMessage" runat="server" 
                Visible="false"
                CssClass="MailForm">
                <asp:Label ID="lblErrorMessage" runat="server"></asp:Label>
            </asp:Panel>
            <asp:Panel ID="pnlForm" runat="server" CssClass="MailForm">
                <h3>SAB Email:</h3>
                <table align="center" border="0">
                    <tr>
                        <td>
                            Your Name *:
                        </td>
                        <td>
                            <asp:TextBox ID="txtName" runat="server" 
                                ClientIDMode="Static"
                                Columns="30"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            Your Email:
                        </td>
                        <td>
                            <asp:TextBox ID="txtEmail" runat="server"
                                ClientIDMode="Static"
                                Columns="30"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            Select an Executive Board Member:
                            <br />
                            <asp:DropDownList ID="lstRecipients" runat="server" EnableViewState="true">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td valign="top">
                            Message:
                        </td>
                        <td>
                            <asp:TextBox ID="txtBody" runat="server"
                                ClientIDMode="Static"
                                Columns="40"
                                Rows="10"
                                TextMode="MultiLine">
                            </asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <cc1:CaptchaControl ID="ccJoin" runat="server" 
                                CaptchaBackgroundNoise="low" 
                                CaptchaLength="5" 
                                CaptchaHeight="60" 
                                CaptchaWidth="200" 
                                CaptchaLineNoise="none" 
                                CaptchaMinTimeout="5" 
                                CaptchaMaxTimeout="240" />
                            <br />
                            Enter the text you see in the image above:
                            <br />
                            <asp:TextBox ID="txtCaptchaResponse" runat="server"></asp:TextBox>
                            <asp:Label ID="lblCaptchaError" runat="server"
                                Font-Bold="true"
                                Text="* Incorrect entry. Try again."
                                Visible="false"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <asp:Button ID="btnSubmit" runat="server"
                                Text="Send Now!"
                                OnClientClick="needToConfirm=false;" />
                        </td>
                    </tr>
                    <tr>
                        <td align="center" colspan="2">
                            * indicates a required field
                        </td>
                    </tr>
                </table>
                <br style="clear:both;" />
            </asp:Panel>
        </div>
        <div class="contactUsInfoBox">
            <dl>
                <dt>
            <h3>SAB Office:</h3>
                </dt>
                <dd>
            Student Center 207
                </dd>
            </dl>
        </div>
        <div class="contactUsInfoBox">
            <dl>
                <dt>
            <h3>SAB Director's Office:</h3>
                </dt>
                <dd>
            Kendra Regehr<br />
            Student Center 206<br /><br />
            Phone: 501.450.3235<br />
            Fax: 501.450.5874<br /><br />
            Mailing Address:<br />
            &nbsp;&nbsp;&nbsp;&nbsp;UCA Box 5101<br />
            &nbsp;&nbsp;&nbsp;&nbsp;Conway, AR 72035
                </dd>
            </dl>
        </div>
        <br style="clear:both;" />
    </div>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="foot" Runat="Server">
    <script type="text/javascript">
        var needToConfirm = true;
        var email = document.getElementById('txtEmail');
        var message = document.getElementById('txtBody');

        window.onbeforeunload = confirmExit;

        function confirmExit() 
        {
            if (needToConfirm) 
            {
                var name = document.getElementById('txtName');
                if (name.value != '' || email.value != '' || message.value != '')
                {
                    return "Are you sure you want to abandon your message to us?\nSelect \"Stay on this Page\" to finish your message.";
                }
            }
        }
    </script>
</asp:Content>

