 FCalc 2.0 Windows version
===========================
By Joseph T Realmuto and Okrasa Ghia, November 2005.

Developing some railcars I had problem with them being to powerfull despite
correct weight and power. In order to cure this I downloaded and used the
earlier version of FCalc 2.0. It helped my problem but I was dissapointed 
with the GUI. This lead to the program you have now downloaded.

The calculations are those from Joes earlier version only change is how
steam engines are handled. For these there is now a new input, weight on 
drivers, which is used to cover the internal friction of the engine.

The program can also read inputs from an eng/wag file and write the
calculated friction values to the file. For this is used the same dll-file 
as used by my other utilities.

Excel spreadsheets and other documentation is by Joe who deserves all credit 
for the friction calculations (please read his "Original Readme.txt").
My contribution is to wrap it in a Windows application.

 Installation
--------------
Installation is simple and does not even require you have MSTS installed 
on the computer. Only requirement is that you have the 
'Microsoft .Net Framework v1.1' or later, the name of the installer 
is 'dotnetfx.exe'. If you don't have it already it can be freely downloaded 
here: http://msdn.microsoft.com or you can get it from 'Windows Update'.
If you have a fairly new computer with Windows regularly updated chance
is you already have the framework installed.

The files included with this package can be placed anywhere you like.
Only requirement is that the exe files and dll are located in the same 
folder or the application will fail to locate the dll.
The FCalc.exe.config is a configuration file that can be used set units
used and if the Davis coefficients should be shown.

/Okrasa
okrasaghia@yahoo.com