@echo off
copy systimeset.dll c:\windows\system32\systimeset.dll
regsvr32 systimeset.dll
