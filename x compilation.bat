start/w setversion.exe

pyinstaller -w -F -i "logo.ico" Autodor.py

xcopy %CD%\*.xltx %CD%\dist /H /Y /C /R
xcopy %CD%\*.dotx %CD%\dist /H /Y /C /R
xcopy %CD%\*.ico %CD%\dist /H /Y /C /R
xcopy %CD%\*.ini %CD%\dist /H /Y /C /R

xcopy C:\vxvproj\tnnc-Autodor\AutodorApp\dist C:\vxvproj\tnnc-Autodor\ConsoleApp\ /H /Y /C /R
