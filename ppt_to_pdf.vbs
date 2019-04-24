On Error Resume Next
Const PpExportFormatPDF = 2
Set oWord = WScript.CreateObject("PowerPoint.Application")
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set fds=fso.GetFolder(".")
Set ffs=fds.Files
For Each ff In ffs
    If (LCase(Right(ff.Name,4))=".ppt" Or LCase(Right(ff.Name,4))="pptx" ) And Left(ff.Name,1)<>"~" Then
        Set oDoc=oWord.Presentations.Open(ff.Path)
        'odoc.ExportAsFixedFormat Left(ff.Path,InStrRev(ff.Path,"."))&"pdf",wdExportFormatPDF
		'Set pptfile = oWord.Presentations.Open(ff.path,InStrRev(ff.Path,"."))&"pdf",false,false,false)
		'pptfile.Saveas path & "\" & InStrRev(ff.Path,"."))&"pdf",32,false
		'Set pptfile = oWord.Presentations.Open((ff.path & "\" & InStrRev(ff.Path,".")) & "pdf",false,false,false)
		'Set pptfile = oWord.Presentations.Open(ff.path)
		oDoc.Saveas ff.path & ".pdf",32,false
		oWord.Quit
		'Set pptfile= nothing
		'oDoc.ExportAsFixedFormat
		If Err.Number Then
        MsgBox Err.Description
        End If
    End If
Next
