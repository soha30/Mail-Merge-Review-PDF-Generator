<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="default2.aspx.cs" Inherits="ali1982ReviewMailmerge.demo.default2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <asp:Repeater ID="rptCertificates" runat="server">
    <ItemTemplate>
        <asp:HyperLink ID="lnkCertificate" runat="server" 
            NavigateUrl='<%# Eval("FileUrl") %>' 
            Text='<%# Eval("FileName") %>' 
            Target="_blank" />
    </ItemTemplate>
</asp:Repeater>

   
    <p>
        <asp:GridView ID="gvStudent" runat="server" AutoGenerateColumns="False" DataKeyNames="studentId">
            <Columns>
                <asp:BoundField DataField="studentId" HeaderText="studentId" InsertVisible="False" ReadOnly="True" SortExpression="studentId" />
                <asp:BoundField DataField="studentName" HeaderText="studentName" SortExpression="studentName" />
                <asp:BoundField DataField="gpa" HeaderText="gpa" SortExpression="gpa" />
                <asp:BoundField DataField="major" HeaderText="major" SortExpression="major" />
                <asp:BoundField DataField="certificateDate" HeaderText="certificateDate" SortExpression="certificateDate" />
                 <asp:BoundField DataField="email" HeaderText="email" SortExpression="email" />
           
            </Columns>
        </asp:GridView>
    </p>
    <p>
        <asp:Label ID="lblOutput" runat="server"></asp:Label>
    </p>
    <p>
        <asp:Button ID="btnIssueCertificateViaArray" runat="server" Text="(1) Issue Certificate ViaArray" OnClick="btnIssueCertificateViaArray_Click" Width="200px" />
        &nbsp;
        <asp:Button ID="btnIssueCertificateViaDT" runat="server" Text="(2) IssueCertificateViaDT" OnClick="btnIssueCertificateViaDT_Click1" Width="200px" />
        <asp:Button ID="btnIssueCertificateViaExce" runat="server" OnClick="btnIssueCertificateViaExce_Click" Text="(3) IssueCertificateViaExcel" Width="200px" />
        <asp:Button ID="btnSendCertificateByEmail" runat="server" OnClick="btnSendCertificateByEmail_Click" Text="(1-2) Send By Email" Width="200px" />
        <asp:Button ID="btnIssueCertificateViaEmail" runat="server" OnClick="btnIssueCertificateViaEmail_Click" Text="(4) IssueCertificateByEmail" Width="200px" />
    </p>
   <div> 
      
  
<asp:LinkButton ID="btnViewCertificate" runat="server" Text="View Certificates" OnClick="ShowCertificates" />



   </div>
    <asp:Repeater ID="Repeater1" runat="server" >
    <ItemTemplate>
        <asp:HyperLink ID="lnkCertificate" runat="server" 
            NavigateUrl='<%# Eval("FileUrl") %>' 
            Text='<%# Eval("FileName") %>' 
            Target="_blank" />
    </ItemTemplate>
</asp:Repeater>
    <hr />

    <h3>Add Watermark to Document</h3>
    <p>
        <label for="watermarkText">Watermark Text:</label>
        <asp:TextBox ID="lblWatermarkText" runat="server"></asp:TextBox>
        <br /><br />
        <label for="fontSize">Font Size:</label>
        <asp:TextBox ID="lblFontSize" runat="server"></asp:TextBox>
        <br /><br />
        <label for="fontColor">Font Color (e.g., Blue, Red):</label>
        <asp:TextBox ID="lblFontColor" runat="server"></asp:TextBox>
        <br /><br />
        <asp:Button ID="btnAddWatermark" runat="server" Text="Add Watermark" OnClick="btnAddWatermark_Click" />
    </p>
   

</asp:Content>

