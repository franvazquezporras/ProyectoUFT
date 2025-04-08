﻿LoadFunctionLibrary "..\Libraries\CommonPage.qfl"
LoadFunctionLibrary "..\Libraries\HomePage.qfl"
LoadFunctionLibrary "..\Libraries\LoginPage.qfl"
LoadFunctionLibrary "..\Libraries\GeneralFunctions.qfl"

Call Setup("chrome", "https://advantageonlineshopping.com")
initializeCommonPage()
initializeHomePage()
initializeLoginPage()

Call WaitExt(2)
L_UserMenu.Click
Call LoginWithAdmin()
Call WaitExt(1)
L_UserMenu.Click
L_SignOut.Click

CleanHomePage()
CleanCommonPage()
CleanLoginPage()
