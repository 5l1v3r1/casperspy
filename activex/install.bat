@echo off
echo :: ---------------------------------------
echo :: Install and Register activex components
echo :: ---------------------------------------
pause
for %%f in (*.ocx *.dll) do regsvr32 /s %%f
echo :: -------
echo :: Succeed
echo :: -------

