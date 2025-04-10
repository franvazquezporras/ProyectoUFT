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

' --- Validar que haya al menos 2 productos ---
If IsArray(W_SpeakerList) Then
    totalProductos = UBound(W_SpeakerList) + 1

    If totalProductos > 0 Then
        Randomize
        Dim cantidadAgregar
        cantidadAgregar = Int((totalProductos) * Rnd) + 1

        MsgBox "Se van a agregar " & cantidadAgregar & " producto(s) al carrito."

        For i = 1 To cantidadAgregar
            indexRandom = Int(totalProductos * Rnd)
            W_SpeakerList(indexRandom).Click
            Call WaitExt(3)

            ' Reasignar botón tras cada navegación
            Dim detailPage
            Set detailPage = Browser("micclass:=Browser").Page("micclass:=page")

            If detailPage.WebButton("name:=ADD TO CART").Exist(3) Then
                detailPage.WebButton("name:=ADD TO CART").Click
                Call WaitExt(1)
            Else
                Reporter.ReportEvent micWarning, "Botón no encontrado", "No se encontró el botón ADD TO CART para el producto índice " & indexRandom
            End If

            ' Volver a la lista de productos
            L_Speakers.Click
            Call WaitExt(2)

            ' Volver a cargar la lista
            InitializeSpeakersPage()
            Call WaitExt(2)
        Next

    Else
        Reporter.ReportEvent micFail, "Error", "No hay productos disponibles"
    End If
Else
    Reporter.ReportEvent micFail, "Error", "W_SpeakerList no fue inicializado correctamente"
End If
' --- Limpieza de páginas ---
CleanHomePage()
CleanCommonPage()
CleanLoginPage()
CleanRegisterPage()
CleanSpeakersPage()
