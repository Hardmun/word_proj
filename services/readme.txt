Sourse:
a) https://gist.github.com/guillaumevincent/d8d94a0a44a7ec13def7f96bfb713d3f
b)

1) To install win. service using pyinstall
pyinstaller -F --hidden-import=win32timezone WindowsService.py

Error starting service: The service did not respond to the start or control request in a timely fashion.
Solution:
This specific problem was solved by copying this file - pywintypes36.dll
From -> Python36\Lib\site-packages\pywin32_system32
To -> Python36\Lib\site-packages\win32

