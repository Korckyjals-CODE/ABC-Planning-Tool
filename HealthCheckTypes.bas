Attribute VB_Name = "HealthCheckTypes"
Option Explicit

' ===========================
' Data Structures for Issue Tracking
' ===========================

' Summary record for individual file/sheet health check results
Type IssueSummary
    filePath As String
    fileName As String
    sheetName As String
    WorkbookName As String
    IssueCount As Long
    Issues As Collection  ' Collection of issue detail strings
End Type

' Summary record for folder-level health check results
Type FolderSummary
    folderPath As String
    FolderName As String
    totalFiles As Long
    TotalIssues As Long
    FileSummaries As Collection  ' Collection of IssueSummary records
End Type

' Summary record for bimester-level health check results
Type BimesterSummary
    BimesterPath As String
    BimesterName As String
    TotalFolders As Long
    totalFiles As Long
    TotalIssues As Long
    FolderSummaries As Collection  ' Collection of FolderSummary records
End Type
