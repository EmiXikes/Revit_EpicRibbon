echo Configuration: $(Configuration)
if $(Configuration) == Debug2022 goto 2022
if $(Configuration) == Debug2023 goto 2023

:2022
mkdir "C:\Users\User\source\repos\_Revit\Revit_EpicRibbon\bin\Apps\EpicWallBox"
mkdir "C:\Epic\RevitAddInsSetup\2022\Apps\EpicWallBox"
mkdir "C:\Epic\RevitAddIns\2022\Apps\EpicWallBox"
xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Users\User\source\repos\_Revit\Revit_EpicRibbon\bin\Apps\EpicWallBox" /Y /I /E
xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddIns\2022\Apps\EpicWallBox" /Y /I /E
xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddInsSetup\2022\Apps\EpicWallBox" /Y /I /E
goto exit

:2023
mkdir "C:\Users\User\source\repos\_Revit\Revit_EpicRibbon\bin\Apps\EpicWallBox"
mkdir "C:\Epic\RevitAddIns\2023\Apps\EpicWallBox"
mkdir "C:\Epic\RevitAddInsSetup\2023\Apps\EpicWallBox"
xcopy "$(ProjectDir)bin\Debug 2023\*.*" "C:\Users\User\source\repos\_Revit\Revit_EpicRibbon\bin\Apps\EpicWallBox" /Y /I /E
xcopy "$(ProjectDir)bin\Debug 2023\*.*" "C:\Epic\RevitAddIns\2023\Apps\EpicWallBox" /Y /I /E
xcopy "$(ProjectDir)bin\Debug 2023\*.*" "C:\Epic\RevitAddInsSetup\2023\Apps\EpicWallBox" /Y /I /E
goto exit

:exit

xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddIns\2022\Ribbon" /Y /I /E
xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddIns\2022\Ribbon" /Y /I /E

xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddIns\2023\Ribbon" /Y /I /E
xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddIns\2023\Ribbon" /Y /I /E

xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddInsSetup\2022\Ribbon" /Y /I /E
xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddInsSetup\2022\Ribbon" /Y /I /E

xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddInsSetup\2023\Ribbon" /Y /I /E
xcopy "$(ProjectDir)bin\Debug\*.*" "C:\Epic\RevitAddInsSetup\2023\Ribbon" /Y /I /E

:exit