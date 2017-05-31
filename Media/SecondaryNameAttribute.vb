Public Class SecondaryNameAttribute

    Inherits Attribute

    Private SecondaryName As String

    Sub New(SecondaryName As String)
        Me.SecondaryName = SecondaryName
    End Sub

    ReadOnly Property Name As String
        Get
            Return SecondaryName
        End Get
    End Property

End Class
