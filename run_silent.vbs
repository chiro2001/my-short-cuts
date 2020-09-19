DIM objShell
set objShell=wscript.createObject("wscript.shell")
iReturn=objShell.Run("cmd /C run_silent.cmd", 0, TRUE)