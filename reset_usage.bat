@echo off
echo Resetting pic2ppt usage count...
echo.
echo User Profile: %USERPROFILE%
echo Temp Folder: %TEMP%
echo.

echo Deleting data files...
if exist "%TEMP%\.pic2ppt_data" (
    del /f /q "%TEMP%\.pic2ppt_data"
    echo [OK] %%TEMP%%\.pic2ppt_data deleted
) else (
    echo [Not Found] %%TEMP%%\.pic2ppt_data
)

if exist "%USERPROFILE%\.pic2ppt_cfg" (
    del /f /q "%USERPROFILE%\.pic2ppt_cfg"
    echo [OK] %%USERPROFILE%%\.pic2ppt_cfg deleted
) else (
    echo [Not Found] %%USERPROFILE%%\.pic2ppt_cfg
)

if exist "%USERPROFILE%\.config\pic2ppt\data.bin" (
    del /f /q "%USERPROFILE%\.config\pic2ppt\data.bin"
    echo [OK] %%USERPROFILE%%\.config\pic2ppt\data.bin deleted
) else (
    echo [Not Found] %%USERPROFILE%%\.config\pic2ppt\data.bin
)

echo.
echo Done! You now have 3 conversions remaining.
echo.
pause
