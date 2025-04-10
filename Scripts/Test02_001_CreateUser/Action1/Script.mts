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

If CheckVisibility(L_CreateNewAccountLogin,5) Then
	Call WaitExt(1)
	L_CreateNewAccountLogin.Click
End If

Call WaitExt(2)
Call CrearUsuarioDatos()
Call WaitExt(1)
Call CheckVisibility(W_UserLogged,5)

CleanHomePage()
CleanCommonPage()
CleanLoginPage()	
CleanRegisterPage()
