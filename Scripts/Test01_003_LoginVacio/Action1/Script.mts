LoadFunctionLibrary "..\Libraries\CommonPage.qfl"
LoadFunctionLibrary "..\Libraries\HomePage.qfl"
LoadFunctionLibrary "..\Libraries\LoginPage.qfl"
LoadFunctionLibrary "..\Libraries\GeneralFunctions.qfl"

Call Setup("chrome", "https://advantageonlineshopping.com")
initializeCommonPage()
initializeHomePage()
initializeLoginPage()

Call WaitExt(2)
L_UserMenu.Click
Call LoginIncorrecto("","")
Call WaitExt(1)
Call CheckVisibility(W_EmptyUserText,30)
Call CheckVisibility(W_EmptyPasswordText,30)

CleanHomePage()
CleanCommonPage()
CleanLoginPage()
