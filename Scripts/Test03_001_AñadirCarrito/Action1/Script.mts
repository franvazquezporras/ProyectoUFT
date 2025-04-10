LoadFunctionLibrary "..\Libraries\CommonPage.qfl"
LoadFunctionLibrary "..\Libraries\HomePage.qfl"
LoadFunctionLibrary "..\Libraries\LoginPage.qfl"
LoadFunctionLibrary "..\Libraries\RegisterPage.qfl"
LoadFunctionLibrary "..\Libraries\SpeakersPage.qfl"
LoadFunctionLibrary "..\Libraries\GeneralFunctions.qfl"

' --- Inicialización de páginas ---
initializeCommonPage()
initializeHomePage()
initializeLoginPage()
InicializeRegisterPage()

' --- Acceder a la categoría de speakers ---
L_SpeakersCategoryImg.Click
Call WaitExt(2)

' --- Cargar página de productos ---
InitializeSpeakersPage()
Call WaitExt(2)

If IsArray(W_SpeakerList) Then
    Dim total, cantidad, i, indexRandom
    total = UBound(W_SpeakerList)

    If total >= 0 Then
        ' Número aleatorio de productos a añadir (mínimo 1)
        cantidad = Int((total + 1) * Rnd) + 1

        For i = 1 To cantidad
            indexRandom = Int((total + 1) * Rnd) ' Índice aleatorio
            W_SpeakerList(indexRandom).Click
            Call WaitExt(3)

            B_AddToCart.Click
            Call WaitExt(1)

            ' Volver a la lista
            L_Speakers.Click
            Call WaitExt(2)

            ' Re-inicializar productos tras navegar
            InitializeSpeakersPage()
        Next

        MsgBox "Se agregaron " & cantidad & " producto(s) aleatorios al carrito."
    Else
        Reporter.ReportEvent micFail, "Error", "No se encontraron productos en la lista."
    End If
Else
    Reporter.ReportEvent micFail, "Error", "La lista de productos no está inicializada."
End If

' --- Limpieza de páginas ---
CleanHomePage()
CleanCommonPage()
CleanLoginPage()
CleanRegisterPage()
CleanSpeakersPage()
