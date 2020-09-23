Place the font in windows\fonts\
Place the Screensaver in windows\system32\
The agent is just an example picture for use to show the background imaging. Browse for it from the Configarations form, due to the size I could not include an example movie


The project is designed to be ran as a screensaver so it will not run in the IDE with its default settings. 

To alter the settings change the startup object from sub Main to FrmMain, this bypasses the check. 

But before compiling to be run as a screensaver change it back, the extension for the screensaver is not .exe but .scr

If you have XP, there is the manifest file included so it will have the XP style controls. Please include the manifest file in the same directory as the Screensaver itself. The default directory for the screensaver is \Windows\System32\


For more information about the screensaver or just the newest version please visit my site:

www.QuantumCoding.cjb.net