<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeFile="Default.aspx.cs" Inherits="_Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <h2>Current Emails</h2>
    <asp:Repeater runat="server" ID="emails">
        <ItemTemplate>
            <div class="email">
            <h3><%# ((OpenPop.Mime.Message)Container.DataItem).ToMailMessage().Subject%></h3>
            <p><%# ((OpenPop.Mime.Message)Container.DataItem).ToMailMessage().Body%></p>
            </div>
        </ItemTemplate>
    </asp:Repeater>
</asp:Content>
