VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function GetWebFileNameTitle(dwUrl As String) As String
Dim lpos As Integer, I As Integer

    For I = 1 To Len(dwUrl)
        If Mid(dwUrl, I, 1) = "/" Then lpos = I
    Next
    
    If lpos <= 0 Then
        lpos = 0
        dwUrl = ""
        GetWebFileNameTitle = ""
        Exit Function
    Else
        GetWebFileNameTitle = Mid(dwUrl, lpos + 1, Len(dwUrl))
        lpos = 0
        dwUrl = ""
    End If
    
End Function

Function DoDownload(TheObject As Object)
Dim nUrl As String, nUrlTile As String
On Error Resume Next

    frmdload.lstV.ListItems.Clear
    
    frmdload.lstV.ColumnHeaders(1).Text = "Filename"
    frmdload.lstV.ColumnHeaders(2).Text = "Location"
    frmdload.lstV.ColumnHeaders(1).Width = 2099.906
    frmdload.lstV.ColumnHeaders(2).Width = (frmdload.lstV.Width - frmdload.lstV.ColumnHeaders(1).Width - 10)
    
    For I = 1 To TheObject.links.length - 1
       
        nUrl = TheObject.links(I)
        nUrlTile = GetWebFileNameTitle(TheObject.links(I))

        If Len(nUrlTile) > 0 Then
             iCount = iCount + 1
            frmdload.lstV.ListItems.Add , nUrl, nUrlTile, 1, 1
            frmdload.lstV.ListItems(iCount).SubItems(1) = nUrl
        End If
    Next
    
    frmdload.lblTitle.Caption = TheObject.Title
    frmdload.lblURL.Caption = TheObject.URL
    frmdload.lblFondUrls.Caption = "[" & I - 1 & "]"
    If Len(TheObject.referrer) <= 0 Then
        frmdload.lblRef = "no-referrer found"
    Else
        frmdload.lblRef = TheObject.referrer
    End If
    
    I = 0
    frmdload.Show vbModal
    
End Function
