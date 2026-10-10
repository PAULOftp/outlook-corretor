' Cria os atalhos da A.P.IAv1 no Ambiente de Trabalho e no menu Iniciar (com o ícone)
Set sh = CreateObject("WScript.Shell")
pasta = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
For Each destino In Array(sh.SpecialFolders("Desktop"), sh.SpecialFolders("Programs"))
  Set l = sh.CreateShortcut(destino & "\A.P.IAv1.lnk")
  l.TargetPath = pasta & "\iniciar-apiav1.bat"
  l.WorkingDirectory = pasta
  l.IconLocation = pasta & "\apiav1.ico"
  l.WindowStyle = 7
  l.Description = "A.P.IAv1 - assistente pessoal do Paulo"
  l.Save
Next
MsgBox "Atalhos A.P.IAv1 criados no Ambiente de Trabalho e no menu Iniciar.", 64, "A.P.IAv1"
