%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\installutil.exe NPOIHelper.WinService.exe
Net Start NPOIHelperPrintService
sc config NPOIHelperPrintService start= auto