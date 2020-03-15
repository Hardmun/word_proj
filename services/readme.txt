Sourse:
a) https://gist.github.com/guillaumevincent/d8d94a0a44a7ec13def7f96bfb713d3f
b)

1) To install win. service using pyinstall
pyinstaller -F --icon=word.ico --hidden-import=win32timezone Wordsplit.py

Error starting service: The service did not respond to the start or control request in a timely fashion.
Solution:
This specific problem was solved by copying this file - pywintypes36.dll
From -> Python36\Lib\site-packages\pywin32_system32
To -> Python36\Lib\site-packages\win32

2) to install xlrd need to use two options:
    python3 -m pip install --user xlrd
    python3 -m pip install xlrd

3) Console Root->Component Services->Computers->My Computer->DCOM Config->Microsoft Word Document->Right
Click(Properties)->Identity Tab
Then select interactive user instead of launching user. By setting this MSWordwill be executed with the rights
of user that is currently logged on.

