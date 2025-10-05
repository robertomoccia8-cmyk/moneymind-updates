Public Class Transazione
    Public Property ID As Integer
    Public Property Data As Date
    Public Property Importo As Decimal
    Public Property Descrizione As String
    Public Property Causale As String
    Public Property MacroCategoria As String
    Public Property Categoria As String
    Public Property Necessita As String
    Public Property Frequenza As String
    Public Property Stagionalita As String
    Public Property DataInserimento As DateTime
    Public Property DataModifica As DateTime

    ''' <summary>
    ''' MacroCategoria con fallback a "Non Classificata" se vuota
    ''' </summary>
    Public ReadOnly Property MacroCategoriaDisplay As String
        Get
            Return If(String.IsNullOrWhiteSpace(MacroCategoria), "Non Classificata", MacroCategoria)
        End Get
    End Property

    ''' <summary>
    ''' Categoria con fallback a "Non Classificata" se vuota
    ''' </summary>
    Public ReadOnly Property CategoriaDisplay As String
        Get
            Return If(String.IsNullOrWhiteSpace(Categoria), "Non Classificata", Categoria)
        End Get
    End Property

    Public Sub New()
        ' Valori di default
        ID = 0
        Data = Date.Today
        Importo = 0
        Descrizione = ""
        Causale = ""
        MacroCategoria = ""
        Categoria = ""
        Necessita = ""
        Frequenza = ""
        Stagionalita = ""
        DataInserimento = DateTime.Now
        DataModifica = DateTime.Now
    End Sub

    Public Sub New(descrizione As String, importo As Decimal, data As Date)
        Me.New()
        Me.Descrizione = descrizione
        Me.Importo = importo
        Me.Data = data
    End Sub
End Class
