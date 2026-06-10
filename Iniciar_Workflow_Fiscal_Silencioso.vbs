Set WshShell = CreateObject("WScript.Shell")
' Executa o arquivo .bat de forma silenciosa com a flag --silent (0 oculta a janela, False não aguarda a conclusão)
WshShell.Run "cmd.exe /c Iniciar_Workflow_Fiscal.bat --silent", 0, False
