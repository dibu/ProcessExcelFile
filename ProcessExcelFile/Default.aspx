<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ProcessExcelFile._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:Button ID="btnProcessExcel" runat="server" Text="Procexx Excel" OnClick="btnProcessExcel_Click" />
    <asp:Button ID="btnFindParent" runat="server" OnClick="btnFindParent_Click" Text="Find Parent" />
    <asp:Button ID="btnProxessXML" runat="server" Text="XML" OnClick="btnProxessXML_Click" />
</asp:Content>
