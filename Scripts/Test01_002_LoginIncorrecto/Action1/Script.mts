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
Call LoginIncorrecto(DataTable("USERNAME",dtLocalSheet),DataTable("PASSWORD",dtLocalSheet))
Call WaitExt(1)
B_SignInLogin.Click
Call CheckVisibility(W_IncorrectLoginText,30)

CleanHomePage()
CleanCommonPage()
CleanLoginPage()
