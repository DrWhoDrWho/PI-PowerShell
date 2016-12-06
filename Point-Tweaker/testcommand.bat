@rem testcommand.bat -- batch file to test running as a scheduled task
@rem run task in directory containing this file
@rem cmd.exe /c testcommand.bat
@rem
@echo %date% %time% testcommand.bat running
@echo = = = = = = = = = = = = = = = = = = = = = = = = = = >> testcommand.txt
@echo %date% %time% testcommand.bat running >> testcommand.txt
@rem set >> testcommand.txt