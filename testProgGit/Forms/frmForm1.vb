Imports Inventor

Public Class frmForm1

    Public _inventorApplication As Global.Inventor.Application


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim inventorRunning As Boolean = GetInventorApplication()
        'If Not inventorRunning Then
        If inventorRunning = False Then
            MsgBox("Inventor is nog niet gestart" & vbCr & "Start Inventor op en start dit programma nogmaals op" & vbCr & vbCr, MsgBoxStyle.Critical, "Cadac_2017")
            Return
        End If
        'PLAATS HIER COMMANDO's DIE UITGEVOERD MOETE WORDEN OP STARTEN VAN FORMULIER

        'Controle of er wel een document in inventor is geopend
        If _inventorApplication.Documents.Count = 0 Then
            MsgBox("Geen document geopend. Er dient een partdocument geopend te zijn.")
            Exit Sub 'Stoppen met programma
        End If


        'Controle of huidig document in inventor wel een part is
        Dim oDoc As Inventor.Document
        oDoc = _inventorApplication.ActiveDocument

        If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            MsgBox("Er dient een partdocument geopend te zijn.")
            Exit Sub 'Stoppen met programma
        End If



    End Sub

    Private Function GetInventorApplication() As Boolean

        Try
            _inventorApplication = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function


End Class
