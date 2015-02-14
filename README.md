# office_genius
Repo for how to integrate Python and Excel using win32om, py2exe and Innosetup. This will allow you to invoke Python using VBA and create user defined functions (UDF's) in Excel that offload the data processing to Python. We then compile the Python code down to dll's so that we can distribute our code without requiring users to install Python or our modules. Lastly we use Inno Setup to package everything together to create a custom installer for our AddIN that will guide our users through an easy installation of our addin.

Steps:
0. Download basic_excel_object.py, exampleaddin.xlam, and office_genius_installer.iss to your local machine
1. Install win32com module and py2exe
2. Create a new _reg_clsid_ class attribute that represents the Windows Registry key for our Python class
2.5. Do this by running import pythoncom; pythoncom.CreateGuid() from within python to get a unique key and replacing the registry string in the github file
3. Register your python class as a COM object from the command line by running 'basic_excel_object --register'
4. Open exampleaddin.xlsm and verify that you can instantiate your python class from within VBA. Open the VBA editor by Alt + F11
4.5. Save your .xlsm file as a .xlam file so that it is ready to be installed as an addin
5. Unregister our Python class so that we can re-register it using Inno Setup and py2exe. Run 'basic_excel_object --unregister' from the commandline
6. Compile our Python code with pywexe by running the following from the commandline: 'setup.py py2exe'
6.5 THis will create both the build and dist directories on your local machine. dist will contain your final files that need to be distributed with your addin
7. Download Inno Setup from the internet
8. Open office_genius_installer.iss and modify the path for your dist directory and make sure all files are included. Ensure that your dll has the special 'regserver' command within office_genius_installer
9 Generate your .exe using InnoSetup. This will create the Output directory in the same location as your .iss script where your .exe will be located.
10. Run the installer to ensure that no errors are thrown durring installation. If there are no errors open Excel and you should be able to call the same methods.
