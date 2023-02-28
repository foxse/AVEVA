# AVEVA
Aveva applications repository

PMLNet - PMLNet addin for loading UDETS into AVEVA Engineering and fill UDA attributes
Use this commands for run this addin:
IMPORT 'Aveva.Core.InstLoader'
USING NAMESPACE 'Aveva.Core.InstLoader'
!instLoader = object InstLoader()
!instLoader.Start()

NamingChecker - Console application for finding issues in the data register and generation errors report.
Use console to run this application with command:
NamingChecker.exe "pathToSourceFile"