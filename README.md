#ie-plugin
Browser Helper Object (BHO) sample for IE Plugin.

## Prerequisites
This project was written and built using Visual Studio 2015 on a Windows 7.

To make this work in IE (Tested in IE11):

1. Build the project in 'release' mode
2. Copy the dll (IEExtension.dll) to a folder within the 'Program Files' folder. This is so IE can access it.
3. Open the Visual Studio Command Prompt, browse to the folder where you dropped the IEExtension.dll file (see step 2)
	- To install run: regasm IEExtension.dll
	- To Unistall run: regasm /unregister IEExtension.dll