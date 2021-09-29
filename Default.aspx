<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ExportToExcel._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="container">
        <div class="table-responsive">
            <asp:Literal ID="Results" runat="server" />
            <asp:Button Text="Export to Excel" CssClass="btn btn-sm btn-info" ID="btn_ExportExcel" runat="server" OnClick="btn_ExportExcel_Click" />
        </div>
    </div>

</asp:Content>
