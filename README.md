# BR7-roblox-skybox-exporter-macro
Important info about this macro:
1. It is meant to be used on Bryce 7, specifically designed for the pro version so a **valid license** is expected for the use of this macro.
2. I am not responsible for any misuse, damages or violations of the Bryce 7 license.
3. You need to have AHK installed for this macro to work.

For this macro to work as intended, palettes should be set to default in Bryce 7 and the user should be on **1920x1080** resolution, **100%** scale and **landscape** display orientation for your display settings. 
The macro exports the files in 112.5 FOV and 100% scaling, may add features to the macro later on to modify the numbers yourself in the GUI given.

Before using the macro, for proper outcome the numbers regarding pan v and h are not changed and neither is the aspect ratio, to make sure they are proper have the pan values set at 0 and on a 1:1 aspect ratio for the document which is 512x512 resolution

There are two ways to run this macro
This macro can be run by the exe file itself, or the python script, both are provided allowing you to edit some parts of the python script to your own needs if necessary

If you are using the python script then you would need to have python installed and the libraries, make sure you have the libraries installed run the following command through pip:
``pip install ahk pywin32 psutil``
