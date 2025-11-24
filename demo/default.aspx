<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="ali1982ReviewMailmerge.demo._default" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <p>
        <br />
    </p>
    <p>
        <asp:GridView ID="gvStudent" runat="server" AutoGenerateColumns="False" DataKeyNames="studentId" >
            <Columns>
                <asp:BoundField DataField="studentId" HeaderText="studentId" InsertVisible="False" ReadOnly="True" SortExpression="studentId" />
                <asp:BoundField DataField="studentName" HeaderText="studentName" SortExpression="studentName" />
                <asp:BoundField DataField="nId" HeaderText="nId" SortExpression="nId" />
                <asp:BoundField DataField="gpa" HeaderText="gpa" SortExpression="gpa" />
                <asp:BoundField DataField="major" HeaderText="major" SortExpression="major" />
                <asp:BoundField DataField="certificateDate" HeaderText="certificateDate" SortExpression="certificateDate" />
            </Columns>
        </asp:GridView>
    </p>
    <p>
        &nbsp;</p>
    <p>
        <asp:Label ID="lblOutput" runat="server"></asp:Label>
    </p>
    <p>
        <asp:Button ID="btnIssueCertificateViaArray" runat="server"  Text="(1) Issue Certificate ViaArray" OnClick="btnIssueCertificateViaArray_Click" />
&nbsp;
        <asp:Button ID="btnIssueCertificateViaDT" runat="server" Text="(2) IssueCertificateViaDT" OnClick="btnIssueCertificateViaDT_Click1" />
        <asp:Button ID="btnIssueCertificateViaExce" runat="server" OnClick="btnIssueCertificateViaExce_Click" Text="(3) IssueCertificateViaExcel" />
    </p>
    <p>
        &nbsp;</p>
    <p>
        In the Above buttons,&nbsp; I applied Mail merge in two different ways: </p>
    <p>
        (1)&nbsp; using Array </p>
    <p>
        (2)&nbsp; using DataTable </p>
    <p>
        &nbsp;</p>
    <p>
        Please review the code behind each of those buttons to learn how to reference and apply Mail merge.</p>
</asp:Content>
