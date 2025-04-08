LoadFunctionLibrary "..\Libraries\CommonPage.qfl"
LoadFunctionLibrary "..\Libraries\HomePage.qfl"
LoadFunctionLibrary "..\Libraries\LoginPage.qfl"
LoadFunctionLibrary "..\Libraries\GeneralFunctions.qfl"

initializeCommonPage()
initializeHomePage()
initializeLoginPage()

L_UserMenu.Click
L_SignOut.Click

CleanHomePage()
CleanCommonPage()
CleanLoginPage()

