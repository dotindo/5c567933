<%@ Page Title="" Language="C#" MasterPageFile="~/Main.master" AutoEventWireup="true" CodeBehind="FormFileManager.aspx.cs" Inherits="DotMercy.custom.FormFileManager" %>

<asp:Content ID="content" ContentPlaceHolderID="MainContent" runat="server">

    <dx:ASPxFileManager ID="fileManager" runat="server" >
        <Settings AllowedFileExtensions=".xls,.xlsx" RootFolder="~/custom/FileUpload" />
        <SettingsEditing AllowCreate="true" AllowDelete="true" AllowDownload="true" AllowMove="true" AllowRename="true" />
        <ClientSideEvents CurrentFolderChanged="function(s, e) { DXEventMonitor.Trace(s, e, 'CurrentFolderChanged') }" ErrorOccurred="function(s, e) { DXEventMonitor.Trace(s, e, 'ErrorOccurred') }" FileDownloading="function(s, e) { DXEventMonitor.Trace(s, e, 'FileDownloading') }" FileUploaded="function(s, e) { DXEventMonitor.Trace(s, e, 'FileUploaded') }" FileUploading="function(s, e) { DXEventMonitor.Trace(s, e, 'FileUploading') }" FolderCreated="function(s, e) { DXEventMonitor.Trace(s, e, 'FolderCreated') }" FolderCreating="function(s, e) { DXEventMonitor.Trace(s, e, 'FolderCreating') }" Init="function(s, e) { DXEventMonitor.Trace(s, e, 'Init') }" ItemDeleted="function(s, e) { DXEventMonitor.Trace(s, e, 'ItemDeleted') }" ItemDeleting="function(s, e) { DXEventMonitor.Trace(s, e, 'ItemDeleting') }" ItemMoved="function(s, e) { DXEventMonitor.Trace(s, e, 'ItemMoved') }" ItemMoving="function(s, e) { DXEventMonitor.Trace(s, e, 'ItemMoving') }" ItemRenamed="function(s, e) { DXEventMonitor.Trace(s, e, 'ItemRenamed') }" ItemRenaming="function(s, e) { DXEventMonitor.Trace(s, e, 'ItemRenaming') }" SelectedFileChanged="function(s, e) { DXEventMonitor.Trace(s, e, 'SelectedFileChanged') }" SelectedFileOpened="function(s, e) { DXEventMonitor.Trace(s, e, 'SelectedFileOpened') }" />
    </dx:ASPxFileManager>

</asp:Content>
