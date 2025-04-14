LoadFunctionLibrary "..\Libraries\CommonPage.qfl"
LoadFunctionLibrary "..\Libraries\HomePage.qfl"
LoadFunctionLibrary "..\Libraries\LoginPage.qfl"
LoadFunctionLibrary "..\Libraries\GeneralFunctions.qfl"
LoadFunctionLibrary "..\Libraries\ShoppingCartPage.qfl"

' --- Inicialización de páginas ---
initializeCommonPage()
initializeHomePage()
initializeLoginPage()
InicializeRegisterPage()
InitializeCommonPage()

L_ShoppingCart.Click

Call EmptyCart()


' --- Limpieza de páginas ---
CleanHomePage()
CleanCommonPage()
CleanLoginPage()
InitializeCommonPage()
