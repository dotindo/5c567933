<%@ Page Title="" Language="C#" MasterPageFile="~/Main.master" AutoEventWireup="true" CodeBehind="FormFileManager.aspx.cs" Inherits="DotMercy.custom.FormFileManager" %>

<asp:Content ID="content" ContentPlaceHolderID="MainContent" runat="server">

    <dx:ASPxFileManager ID="ASPxFileManager1" runat="server" Height="360px" >
        <Settings InitialFolder="Images\Product icons" RootFolder="~/custom/FileUpload" ThumbnailFolder="~/Content/FileManager/Thumbnails" />
        <Styles>
            <FolderContainer Width="190px">
            </FolderContainer>
        </Styles>
        <SettingsFileList>
            <ThumbnailsViewSettings ThumbnailSize="50px" />
        </SettingsFileList>
    </dx:ASPxFileManager>

</asp:Content>
