<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

     <br /><br />
    <asp:Button ID="Button1" runat="server" Text="Charger les donnees" Height="59px" Width="308px" OnClick="LoadData" BackColor="White" BorderColor="Black" BorderStyle="Solid" BorderWidth="5px" Font-Bold="True" Font-Size="Medium" ForeColor="Black" Visible="True" />
     <br /><br />
    <asp:Label ID="Label1" runat="server" Text="Professeurs :" Font-Bold="True"></asp:Label> <asp:DropDownList ID="DropDownProfesseurs" runat="server" Width="100px" OnSelectedIndexChanged="LoadDropDownModule" AutoPostBack="True"></asp:DropDownList>
     <br /><br />
    <asp:Label ID="Label2" runat="server" Text="Modules" Font-Bold="True">Modules :</asp:Label> <asp:DropDownList ID="DropDownListModules" runat="server" Width="100px"></asp:DropDownList>
     <br /><br />
    <asp:Label ID="Label4" runat="server" Text="Promotions" Font-Bold="True">Promotions :</asp:Label> <asp:DropDownList ID="DropDownListPromotions" runat="server" Width="100px"></asp:DropDownList>
     <br /><br />
    <asp:Label ID="Label3" runat="server" Text="Groupes" Font-Bold="True">Groupes :</asp:Label> <asp:DropDownList ID="DropDownListGroupes" runat="server" Width="100px"></asp:DropDownList>
     <br /><br />
   
        <asp:Button ID="Button2" runat="server" Text="ToPDF" OnClick="GeneratePDF"/>
        <asp:Button ID="Button3" runat="server" Text="ToXML" OnClick="GenerateXML"/>
  
</asp:Content>
