VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDownload 
   Caption         =   "ダウンロード"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   OleObjectBlob   =   "frmDownload.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDownload_Click()

Dim pdfname, pwd, dl_pwd, dl_usr, url, dlPath

pdfname = cbxPdf.Value
pwd = TLookup(pdfname, "pdf", "pwd")
dl_pwd = TLookup(pdfname, "pdf", "dl_pwd")
dl_usr = TLookup(pdfname, "pdf", "dl_usr")
url = TLookup(pdfname, "pdf", "url")
dlPath = Application.GetSaveAsFilename(pdfname & pwd & ".pdf", "pdfファイル,*.pdf", , "保存ファイル名を指定してください")
If dlPath = "False" Then Exit Sub

Call mkXhr
Call dlUrlToFile(url, dlPath, dl_usr, dl_pwd)

MsgBox "終了しました"

End Sub
