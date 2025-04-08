LoadFunctionLibrary "..\Libraries\CommonPage.qfl"
LoadFunctionLibrary "..\Libraries\HomePage.qfl"
LoadFunctionLibrary "..\Libraries\LoginPage.qfl"
LoadFunctionLibrary "..\Libraries\RegisterPage.qfl"
LoadFunctionLibrary "..\Libraries\GeneralFunctions.qfl"

Call Setup("chrome", "https://advantageonlineshopping.com")
initializeCommonPage()
initializeHomePage()
initializeLoginPage()
InicializeRegisterPage()

Call WaitExt(2)
L_UserMenu.Click
Call WaitExt(1)
L_CreateNewAccountLogin.Click
Call WaitExt(2)
Call CrearUsuarioDatos()

CleanHomePage()
CleanCommonPage()
CleanLoginPage()	
CleanRegisterPage()
