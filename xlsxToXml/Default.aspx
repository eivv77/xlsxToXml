<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="xlsxToXml._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    

    <p>
        <br />
        <asp:FileUpload ID="file1" runat="server" />
        <asp:Button ID="Button1" runat="server" OnClick="btncnvrt_Click" Text="Button" />
        &nbsp;&nbsp;&nbsp;
        <asp:Button ID="Button2" runat="server" OnClick="btncnvrt_Click_without_Guarantees" Text="Button" />
        <asp:GridView ID="grdExcel" runat="server" OnSelectedIndexChanged="GridView1_SelectedIndexChanged">
        </asp:GridView>
    </p>

    

</asp:Content>
