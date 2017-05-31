option explicit
dim fso
set fso=createObject("Scripting.FileSystemObject")
dim file,subfolder
dim fpath,sfpath,filepath
set fpath=fso.getFolder(".\list\")
' テキストファイル用
dim fileob
' エクセル用
dim excel
set excel=createObject("Excel.Application")
dim book,sheet
' サブフォルダ一覧
for each subfolder in fpath.subfolders
  set sfpath=fso.getFolder(".\list\"&subfolder.name)
  ' ファイル一覧
  for each file in sfpath.files
    set filepath=fso.getFile(sfpath&"\"&file.name)
    msgbox filepath
    ' ファイルをテキストファイルで開く
    set fileob=fso.OpenTextFile(filepath,2,false)
    if Err.Number=0 then
      fileob.WriteLine(file.name)
      fileob.Write(filepath)
      fileob.Close
    end if
  next
next
set sheet=Nothing
set book=Nothing
set excel=Nothing
set fpath=Nothing
set sfpath=Nothing
set filepath=Nothing
set fileob=Nothing
set fso=Nothing