PasswordViewer

Password Viewer is VB application that organizes all of your passwords. Passwords are stored encrypted in Access database. The app provides regular windows interface as well as IE explorer bands. Explorer band can insert the password right into <INPUT> tag of the page.

It contains three major parts:
PasswordViewBus DLL that provides all functionality and data access.
PassworView EXE desktop app that uses DLL to view and edit database of passwords.
PassworView Explorer Band OCX that provides password viewer functionality in the Explorer Band.

Passwords and the rest is stored in Access database. Passwords are encrypted using login name/password pair. So if you forget your login or password you won't be able to get to your password list.
	
To run Password Viewer:
	Unzip archive
	Register PasswordViewBus.dll
	Compile PasswordView into EXE
	Run it

To run Password Viewer Explorer Band:
	Unzip archive
	Register PasswordViewBus.dll
	Run pwdviewbands.inf
	Open IE, View->Eplorer Bar->Password Viewer

Big thanks for Explorer Band example and type libs to Eduardo Morcillo (edanmo@geocities.com), visit his site at http://www.domaindlx.com/e_morcillo.
Thanks to Dan Appleman VBPJ article on Crypto API.