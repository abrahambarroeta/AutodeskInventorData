Attribute VB_Name = "CerrarVentana"
Sub CerrarVentana()
    ' Simula la combinación de teclas Ctrl + F4
    ' Luego de esto, debemos de añadir un atajo de teclado en la configuracion de Inventor
    
    Dim objApp As Object
    Set objApp = ThisApplication

    ' Enviar la combinación de teclas Ctrl+F4 (cerrar el documento activo)
    objApp.ActiveDocument.Close False
    
End Sub

    



