for /D %%x in (%windir%\microsoft.net\framework\v4.0*.*) do set netdir=%%x
%netdir%\csc mailwrench.cs

