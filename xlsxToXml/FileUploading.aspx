<%@ Page Title="File Upload" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="FileUploading.aspx.cs" Inherits="xlsxToXml.FileUploading" %>


<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
<body>
    <form id="form1">
        <div>
            <asp:Table ID="tblFileUpload" runat="server" HorizontalAlign="Center" BorderColor="Black" BorderStyle="Solid" BorderWidth="1">
                <asp:TableHeaderRow>
                    <asp:TableHeaderCell Font-Bold="true" Font-Size="X-Large" ColumnSpan="3" ForeColor="Blue">
                FILE UPLOADING<hr />
                    </asp:TableHeaderCell>
                </asp:TableHeaderRow>
                <asp:TableRow>
                    <asp:TableCell Font-Bold="true" VerticalAlign="Middle">
                        <asp:Label ID="lblSelectFile" runat="server" Text="Select File"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell VerticalAlign="Middle">
                        <asp:FileUpload ID="fuFileUpload" runat="server" ForeColor="YellowGreen" Font-Bold="true" />
                    </asp:TableCell>
                    <asp:TableCell VerticalAlign="Middle">
                        <asp:Button ID="btnUpload" runat="server" Text="Upload File" OnClick="btnUpload_Click" BackColor="YellowGreen" Font-Bold="true" Height="30px" />
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableFooterRow>
                    <asp:TableCell ColumnSpan="3">
                        <hr />
                        <asp:Label ID="lblMessage" runat="server" ForeColor="Green" Font-Bold="true"></asp:Label>
                    </asp:TableCell>
                </asp:TableFooterRow>
            </asp:Table>
        </div>
    </form>
</body>
</html>
    </asp:Content>