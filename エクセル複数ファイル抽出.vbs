option explicit
dim fso
set fso=createObject("Scripting.FileSystemObject")
dim file,subfolder
dim fpath,sfpath,filepath
set fpath=fso.getFolder(".\list\")
' �e�L�X�g�t�@�C���p
dim fileob
' �G�N�Z���p
dim excel
set excel=createObject("Excel.Application")
dim book,sheet
' �T�u�t�H���_�ꗗ
for each subfolder in fpath.subfolders
  set sfpath=fso.getFolder(".\list\"&subfolder.name)
  ' �t�@�C���ꗗ
  for each file in sfpath.files
    set filepath=fso.getFile(sfpath&"\"&file.name)
    msgbox filepath
    ' �t�@�C�����e�L�X�g�t�@�C���ŊJ��
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