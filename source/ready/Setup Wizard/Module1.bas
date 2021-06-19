Attribute VB_Name = "Module1"
' ***      *** ***   *****  ***   *******    *******
'  ***    ***  ***   *****  ***  ***   ***   ***  ****
'   ***  ***   ***   *** ** ***  ***   ***   ***   ****
'    ******    ***   ***  *****  ***   ***   ***  ****
'     ****     ***   ***   ****   *******    *******
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Programmer : VINOD KOTIYA
'  B.E. (Information Technology)
'  Semester V
'  University Institute of Technology
'  Rajeev Gandhi Prodyogiki Vishwavidyalaya Bhopal.
'  Address: S-2 ShreeMaya Apartment Sector-B/363
'           Sarvdharm Colony Bhopal-42 (India)
'  Email: vinodkotiya24@rediffmail.com
'  Web : http://vinodkotiya.tripod.com
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Date of Starting:Wednesday,06,Aug 2003, 10:59:20 AM
'  Completion Date :Monday,05,Aug 2003, 11:43:33 AM
'  Associated Projects:1. Main Installer 2. VIN Uninstaller
'
'  First Modification : 10-aug-2003
'                       Debugging feature in compilation
'     window was added.
'  Second Modification : 15-aug-2003
'                       Settings option added.
'  Third Modification : 24-aug-2003
'                       Path creation bug was fixed by
'     CreatePath algorithm.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public AppName As String
Public Version As String
Public Company As String
Public OutputDir As String  'where setup to be created
Public InstallDir As String  'where installation files to be installed
Public WelMessage As Boolean '0 for text 1 for image
Public DispMessage As Boolean '0 for text 1 for image
Public EndMessage As Boolean '0 for text 1 for image
Public WelImage As String 'store welcome image relattive file name
Public DispImage As String 'store left display image relattive file name
Public EndImage As String 'store left display image relattive file name
Public artificialDelay As Double 'contain compile delay time
Public isCompiled As Boolean 'true after compilation . fase when any changees made

Public SystemFiles As String  'store dll files path to be installed. init at compilation

Public SysAd As Boolean 'true to display user name
Public SysCompany As Boolean 'true to display user company
Public RegCode As Boolean 'true to enter reg code
Public colDll As New Collection 'contain original location of dll files
