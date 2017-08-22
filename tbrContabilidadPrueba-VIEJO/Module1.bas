Attribute VB_Name = "Module1"
Dim AP As String

Public Sub Main()
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    frmUltimos.AbrirDatos AP + "Ctas.mdb", "zuliani"
End Sub

