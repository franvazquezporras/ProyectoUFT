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
Call WaitExt(1)
CrearUsuarioDatos()
Call WaitExt(1)

CleanRegisterPage()
CleanHomePage()
CleanCommonPage()
CleanLoginPage()




Function CrearUsuarioDatos()
	E_UsernameRegister.Set(DataTable("nombreUsuario",dtLocalSheet))
	E_EmailRegister.Set(DataTable("correoElectronico",dtLocalSheet))
	E_PasswordRegister.SetSecure(DataTable("password",dtLocalSheet))
	E_ConfirmPasswordRegister.SetSecure(DataTable("password",dtLocalSheet))
	E_FirstNameRegister.Set(DataTable("nombre",dtLocalSheet))
	E_LastNameRegister.Set(DataTable("apellido",dtLocalSheet))
	E_PhoneNumberRegister.Set(DataTable("telefono",dtLocalSheet))
	D_CountryRegister.Select(DataTable("pais",dtLocalSheet))
	E_CityRegister.Set(DataTable("ciudad",dtLocalSheet))
	E_AddressRegister.Set(DataTable("calle",dtLocalSheet))
	E_StateRegister.Set(DataTable("ciudad",dtLocalSheet))
	E_PostalCodeRegister.Set(DataTable("codigoPostal",dtLocalSheet))
	C_AllowOfferPromotion.Set("ON")
	C_IAgree.Set("ON")
	Call WaitExt(2)
	B_Register.Click
End Function
