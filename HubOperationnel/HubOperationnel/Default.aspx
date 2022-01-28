<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>




<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
      <div style="margin:100px 100px 100px 250px">
           <br /><br />
        <asp:Button ID="Button1" runat="server" Text="Charger les donnees" Height="59px" Width="301px" OnClick="LoadData" BackColor="LightSkyBlue" BorderColor="Black" BorderStyle="Solid" BorderWidth="5px" Font-Bold="True" Font-Size="Medium" ForeColor="Black" Visible="True" />
         <br /><br />
        <asp:Label ID="Label1" width="100px" runat="server" Text="Professeurs :" Font-Bold="True"></asp:Label> <asp:DropDownList ID="DropDownProfesseurs" runat="server" Width="180px" OnSelectedIndexChanged="LoadDropDownModule" AutoPostBack="True"></asp:DropDownList>
         <br /><br />
        <asp:Label ID="Label2" width="100px" runat="server" Text="Modules" Font-Bold="True">Modules :</asp:Label> <asp:DropDownList ID="DropDownListModules" runat="server" Width="180px"></asp:DropDownList>
         <br /><br />
        <asp:Label ID="Label4" width="100px" runat="server" Text="Promotions" Font-Bold="True">Promotions :</asp:Label> <asp:DropDownList ID="DropDownListPromotions" runat="server" Width="180px"></asp:DropDownList>
         <br /><br />
        <asp:Label ID="Label3" width="100px" runat="server" Text="Groupes" Font-Bold="True">Groupes :</asp:Label> <asp:DropDownList ID="DropDownListGroupes" runat="server" Width="180px"></asp:DropDownList>
         <br /><br />
         <br /><br />
            &emsp;<asp:Button ID="Button2" runat="server" Text="GenererPDF" OnClick="GeneratePDF" CssClass="btn" Width="100px" BackColor="#0066ff"/>&emsp;&emsp;&emsp;&emsp;&emsp;
            <asp:Button ID="Button3" runat="server" Text="GenererXML" OnClick="GenerateXML" CssClass="btn" Width="100px"  BackColor="#0066ff"/>
      </div>
        
</asp:Content>


