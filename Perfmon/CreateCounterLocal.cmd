Echo Off
@Echo This command expects 4 variables
@Echo 1. The name of the counter
@Echo 2. The output filename
@Echo 3. The Configuration File i.e. the one you created earlier


logman create counter %1 -f bincirc -max 200 -si 00:15:00 --v -o "c:\perflogs\%2" -cf "%userprofile%\desktop\%3.cfg"