Option Strict On
Option Explicit On

Imports System
Imports System.Runtime.InteropServices
Imports Preactor
Imports Preactor.Interop.PreactorObject

<ComVisible(True)> _
<Microsoft.VisualBasic.ComClass("91257c38-d9ae-4fc6-ad11-4e3ef66386a2", "ea25b172-9a2d-493f-9d75-09cc8a8699a1")> _
Public Class Preactor_Extensions
    Public Function Run(ByRef preactorComObject As PreactorObj, ByRef pespComObject As Object) As Integer

        Dim preactor As IPreactor = PreactorFactory.CreatePreactorObject(preactorComObject)

        Return 0
    End Function
End Class
