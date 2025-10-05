Imports System.ComponentModel

''' <summary>
''' ViewModel per le transazioni con supporto per data binding WPF
''' </summary>
Public Class TransazioneViewModel
    Implements INotifyPropertyChanged

#Region "Fields"
    
    Private _transazione As Transazione
    
#End Region

#Region "Constructor"
    
    Public Sub New(transazione As Transazione)
        _transazione = If(transazione, New Transazione())
    End Sub
    
    Public Sub New()
        _transazione = New Transazione()
    End Sub

#End Region

#Region "Properties"
    
    Public Property Id As Integer
        Get
            Return _transazione.ID
        End Get
        Set(value As Integer)
            If _transazione.ID <> value Then
                _transazione.ID = value
                OnPropertyChanged()
            End If
        End Set
    End Property
    
    Public Property DataTransazione As Date
        Get
            Return _transazione.Data
        End Get
        Set(value As Date)
            If _transazione.Data <> value Then
                _transazione.Data = value
                OnPropertyChanged()
                OnPropertyChanged(NameOf(DataFormattata))
                OnPropertyChanged(NameOf(MeseAnno))
            End If
        End Set
    End Property
    
    Public Property Importo As Decimal
        Get
            Return _transazione.Importo
        End Get
        Set(value As Decimal)
            If _transazione.Importo <> value Then
                _transazione.Importo = value
                OnPropertyChanged()
                OnPropertyChanged(NameOf(ImportoFormatted))
                OnPropertyChanged(NameOf(ImportoFormattedAbs))
                OnPropertyChanged(NameOf(IsPositive))
                OnPropertyChanged(NameOf(IsNegative))
                OnPropertyChanged(NameOf(TipoTransazione))
            End If
        End Set
    End Property
    
    Public Property Descrizione As String
        Get
            Return _transazione.Descrizione
        End Get
        Set(value As String)
            If _transazione.Descrizione <> value Then
                _transazione.Descrizione = value
                OnPropertyChanged()
            End If
        End Set
    End Property
    
    Public Property MacroCategoria As String
        Get
            Return _transazione.MacroCategoria
        End Get
        Set(value As String)
            If _transazione.MacroCategoria <> value Then
                _transazione.MacroCategoria = value
                OnPropertyChanged()
                OnPropertyChanged(NameOf(IsClassificata))
                OnPropertyChanged(NameOf(MacroCategoriaDisplay))
            End If
        End Set
    End Property

    Public Property Categoria As String
        Get
            Return _transazione.Categoria
        End Get
        Set(value As String)
            If _transazione.Categoria <> value Then
                _transazione.Categoria = value
                OnPropertyChanged()
                OnPropertyChanged(NameOf(IsClassificata))
                OnPropertyChanged(NameOf(CategoriaDisplay))
            End If
        End Set
    End Property
    
    Public Property Necessita As String
        Get
            Return _transazione.Necessita
        End Get
        Set(value As String)
            If _transazione.Necessita <> value Then
                _transazione.Necessita = value
                OnPropertyChanged()
            End If
        End Set
    End Property
    
    Public Property Frequenza As String
        Get
            Return _transazione.Frequenza
        End Get
        Set(value As String)
            If _transazione.Frequenza <> value Then
                _transazione.Frequenza = value
                OnPropertyChanged()
            End If
        End Set
    End Property
    
    Public Property Stagionalita As String
        Get
            Return _transazione.Stagionalita
        End Get
        Set(value As String)
            If _transazione.Stagionalita <> value Then
                _transazione.Stagionalita = value
                OnPropertyChanged()
            End If
        End Set
    End Property

#End Region

#Region "Computed Properties for UI"
    
    ''' <summary>
    ''' Importo formattato per la visualizzazione
    ''' </summary>
    Public ReadOnly Property ImportoFormatted As String
        Get
            Return Importo.ToString("C2", New System.Globalization.CultureInfo("it-IT"))
        End Get
    End Property
    
    ''' <summary>
    ''' Importo formattato in valore assoluto
    ''' </summary>
    Public ReadOnly Property ImportoFormattedAbs As String
        Get
            Return Math.Abs(Importo).ToString("C2", New System.Globalization.CultureInfo("it-IT"))
        End Get
    End Property
    
    ''' <summary>
    ''' Data formattata per la visualizzazione
    ''' </summary>
    Public ReadOnly Property DataFormattata As String
        Get
            Return DataTransazione.ToString("dd/MM/yyyy")
        End Get
    End Property
    
    ''' <summary>
    ''' Data formattata con giorno della settimana
    ''' </summary>
    Public ReadOnly Property DataCompletaFormattata As String
        Get
            Dim giornoSettimana = DataTransazione.ToString("dddd", New System.Globalization.CultureInfo("it-IT"))
            Return $"{giornoSettimana} {DataTransazione:dd/MM/yyyy}"
        End Get
    End Property
    
    ''' <summary>
    ''' Mese e anno della transazione
    ''' </summary>
    Public ReadOnly Property MeseAnno As String
        Get
            Return DataTransazione.ToString("MMMM yyyy", New System.Globalization.CultureInfo("it-IT"))
        End Get
    End Property
    
    ''' <summary>
    ''' Indica se l'importo è positivo (entrata)
    ''' </summary>
    Public ReadOnly Property IsPositive As Boolean
        Get
            Return Importo > 0
        End Get
    End Property
    
    ''' <summary>
    ''' Indica se l'importo è negativo (uscita)
    ''' </summary>
    Public ReadOnly Property IsNegative As Boolean
        Get
            Return Importo < 0
        End Get
    End Property
    
    ''' <summary>
    ''' Tipo di transazione per display
    ''' </summary>
    Public ReadOnly Property TipoTransazione As String
        Get
            Return If(IsPositive, "Entrata", "Uscita")
        End Get
    End Property
    
    ''' <summary>
    ''' Indica se la transazione è stata classificata
    ''' </summary>
    Public ReadOnly Property IsClassificata As Boolean
        Get
            Return Not String.IsNullOrEmpty(MacroCategoria) AndAlso
                   MacroCategoria <> "Non Classificata" AndAlso
                   Not String.IsNullOrEmpty(Categoria) AndAlso
                   Categoria <> "Non Classificata"
        End Get
    End Property
    
    ''' <summary>
    ''' Status della classificazione
    ''' </summary>
    Public ReadOnly Property StatusClassificazione As String
        Get
            Return If(IsClassificata, "✅ Classificata", "⚠️ Da classificare")
        End Get
    End Property
    
    ''' <summary>
    ''' Icona per il tipo di transazione
    ''' </summary>
    Public ReadOnly Property IconaTipo As String
        Get
            Return If(IsPositive, "💰", "💸")
        End Get
    End Property
    
    ''' <summary>
    ''' Colore per l'importo basato sul tipo
    ''' </summary>
    Public ReadOnly Property ColoreTipo As String
        Get
            Return If(IsPositive, "#28A745", "#DC3545") ' Green for positive, red for negative
        End Get
    End Property
    
    ''' <summary>
    ''' Descrizione troncata per liste
    ''' </summary>
    Public ReadOnly Property DescrizioneTroncata As String
        Get
            If String.IsNullOrEmpty(Descrizione) Then Return ""
            Return If(Descrizione.Length > 50, Descrizione.Substring(0, 47) & "...", Descrizione)
        End Get
    End Property

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

#End Region

#Region "Methods"
    
    ''' <summary>
    ''' Converte il ViewModel in model per il database
    ''' </summary>
    Public Function ToTransazione() As Transazione
        Return _transazione
    End Function
    
    ''' <summary>
    ''' Aggiorna il ViewModel da un model
    ''' </summary>
    Public Sub UpdateFromTransazione(transazione As Transazione)
        If transazione IsNot Nothing Then
            _transazione = transazione
            OnPropertyChanged(Nothing) ' Aggiorna tutte le proprietà
        End If
    End Sub
    
    ''' <summary>
    ''' Clona il ViewModel
    ''' </summary>
    Public Function Clone() As TransazioneViewModel
        Dim clonedTransazione As New Transazione() With {
            .ID = Me.Id,
            .Data = Me.DataTransazione,
            .Importo = Me.Importo,
            .Descrizione = Me.Descrizione,
            .MacroCategoria = Me.MacroCategoria,
            .Categoria = Me.Categoria,
            .Necessita = Me.Necessita,
            .Frequenza = Me.Frequenza,
            .Stagionalita = Me.Stagionalita
        }
        
        Return New TransazioneViewModel(clonedTransazione)
    End Function
    
    ''' <summary>
    ''' Verifica se la transazione corrisponde ai criteri di ricerca
    ''' </summary>
    Public Function MatchesSearch(searchTerm As String) As Boolean
        If String.IsNullOrWhiteSpace(searchTerm) Then Return True
        
        Dim searchLower = searchTerm.ToLowerInvariant()
        
        Return (Descrizione?.ToLowerInvariant().Contains(searchLower) = True) OrElse
               (MacroCategoria?.ToLowerInvariant().Contains(searchLower) = True) OrElse
               (Categoria?.ToLowerInvariant().Contains(searchLower) = True) OrElse
               (ImportoFormatted.ToLowerInvariant().Contains(searchLower)) OrElse
               (DataFormattata.Contains(searchTerm))
    End Function
    
    ''' <summary>
    ''' Verifica se la transazione appartiene al periodo specificato
    ''' </summary>
    Public Function IsInPeriod(startDate As Date, endDate As Date) As Boolean
        Return DataTransazione.Date >= startDate.Date AndAlso DataTransazione.Date <= endDate.Date
    End Function
    
    ''' <summary>
    ''' Verifica se la transazione appartiene al mese/anno specificato
    ''' </summary>
    Public Function IsInMonth(year As Integer, month As Integer) As Boolean
        Return DataTransazione.Year = year AndAlso DataTransazione.Month = month
    End Function

#End Region

#Region "INotifyPropertyChanged Implementation"
    
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    
    Protected Overridable Sub OnPropertyChanged(Optional propertyName As String = "")
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

#End Region

#Region "Overrides"
    
    Public Overrides Function ToString() As String
        Return $"{DataFormattata} - {ImportoFormatted} - {Descrizione}"
    End Function
    
    Public Overrides Function Equals(obj As Object) As Boolean
        Dim other = TryCast(obj, TransazioneViewModel)
        If other Is Nothing Then Return False
        
        Return Me.Id = other.Id AndAlso Me.Id > 0 ' Solo per transazioni salvate
    End Function
    
    Public Overrides Function GetHashCode() As Integer
        Return Id.GetHashCode()
    End Function

#End Region

End Class