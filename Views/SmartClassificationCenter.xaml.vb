Imports System.Collections.ObjectModel
Imports System.Data.SQLite
Imports System.Threading.Tasks
Imports System.Windows.Media
Imports System.Windows.Input
Imports Microsoft.Extensions.Logging.Abstractions
Imports MoneyMind.Services
Imports System.Diagnostics
Imports System.Text.Json
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.ComponentModel
Imports System.Linq
Imports System.Text

Public Class TransactionPreview
    Implements INotifyPropertyChanged

    Public Property ID As Integer
    Public Property Descrizione As String
    Public Property ImportoFormatted As String
    Public Property ParolaChiave As String

    Private _macroCategoria As String = ""
    Public Property MacroCategoria As String
        Get
            Return _macroCategoria
        End Get
        Set(value As String)
            If _macroCategoria <> value Then
                _macroCategoria = value
                OnPropertyChanged(NameOf(MacroCategoria))
            End If
        End Set
    End Property

    Private _categoria As String = ""
    Public Property Categoria As String
        Get
            Return _categoria
        End Get
        Set(value As String)
            If _categoria <> value Then
                _categoria = value
                OnPropertyChanged(NameOf(Categoria))
            End If
        End Set
    End Property

    Public Property Necessita As String
    Public Property Frequenza As String
    Public Property Stagionalita As String
    Public Property ConfidenzaFormatted As String
    Public Property MetodoClassificazione As String
    Public Property OriginalTransaction As Transazione
    
    Private _isSelected As Boolean = True
    Public Property IsSelected As Boolean
        Get
            Return _isSelected
        End Get
        Set(value As Boolean)
            If _isSelected <> value Then
                _isSelected = value
                OnPropertyChanged(NameOf(IsSelected))
                
                ' Notifica il cambio di selezione alla finestra principale
                RaiseEvent SelectionChanged(Me)
            End If
        End Set
    End Property
    
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Event SelectionChanged(item As TransactionPreview)
    
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
End Class

Public Class ManualTransactionView
    Public Property ID As Integer
    Public Property Descrizione As String
    Public Property ImportoFormatted As String
    Public Property ImportoColor As String
    Public Property DataFormatted As String
    Public Property ParolaChiaveEstratta As String
    Public Property OriginalTransaction As Transazione

    Public Sub New(transaction As Transazione, Optional parolaChiave As String = "")
        ID = transaction.ID
        Descrizione = transaction.Descrizione
        ImportoFormatted = transaction.Importo.ToString("C2")
        ImportoColor = If(transaction.Importo >= 0, "#4CAF50", "#F44336")
        DataFormatted = transaction.Data.ToString("dd/MM/yyyy")
        OriginalTransaction = transaction
        ParolaChiaveEstratta = parolaChiave
    End Sub
End Class

Public Class SmartClassificationCenter

    ' === EVENTI ===
    Public Event TransactionsClassified()

    ' === SERVIZI CORE ===
    Private _classificatore As MoneyMind.Services.ClassificatoreTransazioni
    Private _gptClassificatore As MoneyMind.Services.GptClassificatoreTransazioni
    Private _transazioni As List(Of Transazione)
    Private _logger As ILogger = NullLogger.Instance

    ' === CONSIGLI DINAMICI ===
    Private _smartTips As List(Of String)
    Private _currentTipIndex As Integer = 0

    ' === STATISTICHE ===
    Private _lastClassificationResults As New Dictionary(Of String, Integer)

    ' === ANTEPRIMA CLASSIFICAZIONI ===
    Private _previewResults As New ObservableCollection(Of TransactionPreview)
    Private _requiresResetBeforeApproval As Boolean = False ' Flag per "All" scope

    ' === CLASSIFICAZIONE MANUALE ===
    Private _manualTransactions As New ObservableCollection(Of ManualTransactionView)
    Private _selectedManualTransaction As ManualTransactionView

    Public Sub New()
        InitializeComponent()

        ' Inizializzazione servizi
        _classificatore = New MoneyMind.Services.ClassificatoreTransazioni(_logger)
        _gptClassificatore = New MoneyMind.Services.GptClassificatoreTransazioni()
        _transazioni = New List(Of Transazione)()

        ' Inizializzazione consigli
        InitializeSmartTips()

        ' Setup iniziale UI
        InitializeUIAsync()
        
    End Sub

    Public Sub New(transazioniDaClassificare As List(Of Transazione))
        Me.New()
        _transazioni = transazioniDaClassificare
    End Sub

    ''' <summary>
    ''' Inizializzazione consigli dinamici per guidare l'utente
    ''' </summary>
    Private Sub InitializeSmartTips()
        _smartTips = New List(Of String) From {
            "💡 Prima di iniziare, assicurati di avere almeno 10-20 transazioni da classificare per ottenere risultati significativi.",
            "🚀 La modalità 'AI Potenziata' combina pattern esistenti + intelligenza artificiale per risultati ottimali.",
            "⚙️ Configura le API OpenAI e Google Places per sfruttare al massimo l'intelligenza artificiale.",
            "📊 Controlla sempre il dashboard per monitorare i progressi e la qualità delle classificazioni.",
            "🔄 Il sistema impara automaticamente: più lo usi, più diventa preciso nel tempo.",
            "💾 I nuovi pattern vengono salvati automaticamente per migliorare le future classificazioni.",
            "🧪 Usa il 'Test Singola' per provare il sistema su descrizioni specifiche prima di elaborazioni massive.",
            "📈 Un tasso di successo sopra l'80% indica un sistema ben ottimizzato per le tue transazioni.",
            "🎯 Pattern con peso alto (8-10) hanno priorità durante la classificazione automatica.",
            "🔍 Se vedi molte classificazioni 'Altro', considera di aggiungere pattern personalizzati per migliorare l'accuratezza."
        }
        ShowCurrentTip()
    End Sub

    ''' <summary>
    ''' Setup iniziale dell'interfaccia utente
    ''' </summary>
    Private Async Sub InitializeUIAsync()
        Try
            ' Carica dati iniziali
            Await LoadTransactions()
            Await UpdateDashboardStats()

            ' Verifica configurazione AI
            CheckAIConfiguration()

            ' Mostra wizard se primo avvio
            ShowWizardIfNeeded()

        Catch ex As Exception
            ShowError("Errore durante l'inizializzazione", ex.Message)
        End Try
    End Sub

#Region "=== GESTIONE DATI ==="

    ''' <summary>
    ''' Carica tutte le transazioni dal database
    ''' </summary>
    Private Async Function LoadTransactions() As Task
        Try
            ' Assicurati che il database sia inizializzato
            DatabaseManager.InitializeDatabase()

            Dim connectionString = DatabaseManager.GetConnectionString()
            If String.IsNullOrEmpty(connectionString) Then
                Throw New InvalidOperationException("Database non inizializzato correttamente")
            End If

            Dim repository As New MoneyMind.DataAccess.TransazioneRepository(connectionString)
            Dim logger = NullLogger.Instance
            Dim transactionService As New MoneyMind.Business.TransazioneService(repository, logger)
            Dim transactions = Await transactionService.GetAllTransazioniAsync()

            _transazioni = transactions.ToList()

        Catch ex As Exception
            _transazioni = New List(Of Transazione)
            Throw New Exception("Impossibile caricare le transazioni", ex)
        End Try
    End Function

    ''' <summary>
    ''' Aggiorna statistiche dashboard
    ''' </summary>
    Private Async Function UpdateDashboardStats() As Task
        Try
            Await LoadTransactions() ' Ricarica dati freschi

            Dim totalCount = _transazioni.Count
            Dim classifiedCount = _transazioni.Where(Function(t) Not String.IsNullOrEmpty(t.MacroCategoria) AndAlso t.MacroCategoria <> "Non Classificata").Count()
            Dim unclassifiedCount = totalCount - classifiedCount

            ' Aggiornamento UI 
            TxtTotalTransactions.Text = totalCount.ToString()
            TxtClassifiedTransactions.Text = classifiedCount.ToString()
            TxtUnclassifiedTransactions.Text = unclassifiedCount.ToString()

            ' Calcolo percentuale
            Dim classificationRate = If(totalCount > 0, (classifiedCount / totalCount) * 100, 0)
            TxtClassificationRate.Text = $"{classificationRate:F1}%"

            ' Aggiornamento progress bar
            ProgressBarOverall.Value = classificationRate

            ' Colori dinamici
            UpdateProgressBarColor(classificationRate)

        Catch ex As Exception
            _logger?.LogError(ex, "Errore aggiornamento statistiche dashboard")
        End Try
    End Function

    ''' <summary>
    ''' Aggiorna colore progress bar basato su performance
    ''' </summary>
    Private Sub UpdateProgressBarColor(rate As Double)
        Dim brush As SolidColorBrush

        If rate >= 80 Then
            brush = New SolidColorBrush(Color.FromRgb(76, 175, 80))   ' Verde
        ElseIf rate >= 60 Then
            brush = New SolidColorBrush(Color.FromRgb(255, 152, 0))   ' Arancione  
        Else
            brush = New SolidColorBrush(Color.FromRgb(244, 67, 54))   ' Rosso
        End If

        ProgressBarOverall.Foreground = brush
    End Sub

#End Region

#Region "=== CONFIGURAZIONE AI ==="

    ''' <summary>
    ''' Verifica stato configurazione AI e aggiorna UI
    ''' </summary>
    Private Sub CheckAIConfiguration()
        Try
            Dim openAIConfigured = Not String.IsNullOrEmpty(MoneyMind.Services.GestoreApiKeys.CaricaChiaveApi("OpenAI"))
            Dim googleConfigured = Not String.IsNullOrEmpty(MoneyMind.Services.GestoreApiKeys.CaricaChiaveApi("GooglePlaces"))

            If openAIConfigured AndAlso googleConfigured Then
                ' Configurazione completa
                StatusAI.Background = New SolidColorBrush(Color.FromRgb(76, 175, 80))
                TxtStatusAI.Text = "CONFIGURATA"
                AIWarningPanel.Visibility = Visibility.Collapsed
            ElseIf openAIConfigured OrElse googleConfigured Then
                ' Configurazione parziale
                StatusAI.Background = New SolidColorBrush(Color.FromRgb(255, 152, 0))
                TxtStatusAI.Text = "PARZIALE"
                AIWarningPanel.Visibility = Visibility.Visible
            Else
                ' Non configurata
                StatusAI.Background = New SolidColorBrush(Color.FromRgb(244, 67, 54))
                TxtStatusAI.Text = "NON CONFIGURATA"
                AIWarningPanel.Visibility = Visibility.Visible
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Errore verifica configurazione AI")
        End Try
    End Sub

    ''' <summary>
    ''' Mostra finestra configurazione OpenAI
    ''' </summary>
    ''' <summary>
    ''' Apre la finestra impostazioni per configurare le API keys
    ''' </summary>
    Private Sub BtnOpenSettings_Click(sender As Object, e As RoutedEventArgs)
        Try
            ' Apri la finestra impostazioni
            Dim settingsWindow As New SettingsWindow()
            settingsWindow.Show()

            ' Dopo la chiusura delle impostazioni, ricontrolla la configurazione
            CheckAIConfiguration()

        Catch ex As Exception
            ShowError("Errore apertura impostazioni", ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Mostra dialogo quando le API keys sono mancanti e offre di aprire le impostazioni
    ''' </summary>
    Private Async Function ShowAPIKeysMissingDialog() As Task
        Try
            Await Dispatcher.BeginInvoke(Sub()
                                             Dim result = MessageBox.Show(
                                                 "⚠️ API Keys Non Configurate" & vbCrLf & vbCrLf &
                                                 "Per utilizzare la classificazione 'AI Enhanced' sono necessarie almeno una delle seguenti API keys:" & vbCrLf & vbCrLf &
                                                 "🧠 OpenAI (ChatGPT) - Per l'analisi intelligente" & vbCrLf &
                                                 "🗺️ Google Places - Per identificare negozi e servizi" & vbCrLf & vbCrLf &
                                                 "💡 Senza API keys verrà utilizzata la classificazione offline." & vbCrLf & vbCrLf &
                                                 "Vuoi aprire le impostazioni per configurare le API keys?",
                                                 "Configurazione API Keys",
                                                 MessageBoxButton.YesNo,
                                                 MessageBoxImage.Information)

                                             If result = MessageBoxResult.Yes Then
                                                 ' Apri impostazioni
                                                 Dim settingsWindow As New SettingsWindow()
                                                 settingsWindow.Show()
                                                 CheckAIConfiguration()
                                             End If
                                         End Sub)
        Catch ex As Exception
            _logger?.LogError(ex, "Error showing API keys dialog")
        End Try
    End Function

    ''' <summary>
    ''' Mostra dialogo per errori specifici delle API keys
    ''' </summary>
    Private Async Function ShowAPIKeyErrorDialog(apiName As String, errorMessage As String) As Task
        Try
            Await Dispatcher.BeginInvoke(Sub()
                                             Dim icon As String = If(apiName = "OpenAI", "🧠", "🗺️")
                                             Dim result = MessageBox.Show(
                                                 $"❌ Problema con {icon} {apiName}" & vbCrLf & vbCrLf &
                                                 errorMessage & vbCrLf & vbCrLf &
                                                 "💡 Soluzioni possibili:" & vbCrLf &
                                                 "• Verifica che la chiave API sia corretta" & vbCrLf &
                                                 "• Controlla che non sia scaduta" & vbCrLf &
                                                 "• Verifica i limiti di utilizzo" & vbCrLf & vbCrLf &
                                                 "Vuoi aprire le impostazioni per correggere la configurazione?",
                                                 $"Errore {apiName}",
                                                 MessageBoxButton.YesNo,
                                                 MessageBoxImage.Warning)

                                             If result = MessageBoxResult.Yes Then
                                                 ' Apri impostazioni
                                                 Dim settingsWindow As New SettingsWindow()
                                                 settingsWindow.Show()
                                                 CheckAIConfiguration()
                                             End If
                                         End Sub)
        Catch ex As Exception
            _logger?.LogError(ex, "Error showing API key error dialog")
        End Try
    End Function

    ''' <summary>
    ''' Struttura per risultato validazione API keys
    ''' </summary>
    Private Structure APIValidationResult
        Public IsValid As Boolean
        Public Title As String
        Public Message As String
    End Structure

    ''' <summary>
    ''' Valida velocemente le API keys configurate
    ''' </summary>
    Private Async Function ValidateAPIKeysAsync() As Task(Of APIValidationResult)
        Try
            Debug.WriteLine($"DEBUG: === ValidateAPIKeysAsync START ===")
            Dim openAIKey = MoneyMind.Services.GestoreApiKeys.CaricaChiaveApi("OpenAI")
            Dim googleKey = MoneyMind.Services.GestoreApiKeys.CaricaChiaveApi("GooglePlaces")

            Debug.WriteLine($"DEBUG: OpenAI Key: {If(String.IsNullOrEmpty(openAIKey), "VUOTA", $"Presente ({openAIKey.Length} caratteri)")}")
            Debug.WriteLine($"DEBUG: Google Key: {If(String.IsNullOrEmpty(googleKey), "VUOTA", $"Presente ({googleKey.Length} caratteri)")}")

            ' Se nessuna chiave configurata
            If String.IsNullOrEmpty(openAIKey) AndAlso String.IsNullOrEmpty(googleKey) Then
                Debug.WriteLine($"DEBUG: Nessuna API key configurata")
                Return New APIValidationResult With {
                    .IsValid = False,
                    .Title = "API Keys Mancanti",
                    .Message = "Il metodo 'AI Enhanced' richiede almeno una API key:" & vbCrLf &
                              "🧠 OpenAI (ChatGPT) - Per l'analisi intelligente" & vbCrLf &
                              "🗺️ Google Places - Per identificare negozi e servizi"
                }
            End If

            ' Test veloce validità chiavi (con timeout breve)
            Debug.WriteLine($"DEBUG: Inizio test validità chiavi...")
            Dim validationTasks As New List(Of Task(Of Boolean))
            Dim errorMessages As New List(Of String)

            If Not String.IsNullOrEmpty(openAIKey) Then
                Debug.WriteLine($"DEBUG: Aggiungendo test OpenAI...")
                validationTasks.Add(QuickValidateOpenAI(openAIKey))
            End If
            If Not String.IsNullOrEmpty(googleKey) Then
                Debug.WriteLine($"DEBUG: Aggiungendo test Google Places...")
                validationTasks.Add(QuickValidateGooglePlaces(googleKey))
            End If

            Debug.WriteLine($"DEBUG: Totale test da eseguire: {validationTasks.Count}")

            ' Attendi tutti i test (con timeout di 5 secondi)
            Try
                Debug.WriteLine($"DEBUG: Eseguendo test in parallelo...")
                Dim results = Await Task.WhenAll(validationTasks).ConfigureAwait(False)
                Debug.WriteLine($"DEBUG: Test completati, risultati: [{String.Join(", ", results)}]")

                ' Se almeno una API funziona, è OK
                If results.Any(Function(r) r) Then
                    Debug.WriteLine($"DEBUG: ValidateAPIKeysAsync - SUCCESSO (almeno una API valida)")
                    Return New APIValidationResult With {.IsValid = True}
                End If

                ' Tutte le API hanno fallito
                Dim failedMessage = ""
                If Not String.IsNullOrEmpty(openAIKey) Then failedMessage += "🧠 OpenAI: Chiave non valida" & vbCrLf
                If Not String.IsNullOrEmpty(googleKey) Then failedMessage += "🗺️ Google Places: Chiave non valida"

                Return New APIValidationResult With {
                    .IsValid = False,
                    .Title = "API Keys Non Valide",
                    .Message = "Le API keys configurate non sono funzionanti:" & vbCrLf & failedMessage
                }

            Catch ex As TimeoutException
                Return New APIValidationResult With {
                    .IsValid = False,
                    .Title = "Timeout Validazione",
                    .Message = "Impossibile validare le API keys (timeout di rete)." & vbCrLf &
                              "Verifica la connessione internet."
                }
            End Try

        Catch ex As Exception
            Return New APIValidationResult With {
                .IsValid = False,
                .Title = "Errore Validazione",
                .Message = $"Errore durante la validazione: {ex.Message}"
            }
        End Try
    End Function

    ''' <summary>
    ''' Test veloce validità chiave OpenAI (timeout 3 secondi)
    ''' </summary>
    Private Async Function QuickValidateOpenAI(apiKey As String) As Task(Of Boolean)
        Try
            Debug.WriteLine($"DEBUG: QuickValidateOpenAI - Inizio test con chiave: {apiKey.Substring(0, Math.Min(10, apiKey.Length))}...")
            Using client As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(3)}
                client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)

                ' Test minimo: lista modelli (più veloce di chat completion)
                Debug.WriteLine($"DEBUG: QuickValidateOpenAI - Chiamata GET a /v1/models...")
                Dim response = Await client.GetAsync("https://api.openai.com/v1/models")
                Debug.WriteLine($"DEBUG: QuickValidateOpenAI - Risposta: {response.StatusCode}")
                Dim isValid = response.IsSuccessStatusCode
                Debug.WriteLine($"DEBUG: QuickValidateOpenAI - Risultato: {isValid}")
                Return isValid
            End Using
        Catch ex As Exception
            Debug.WriteLine($"DEBUG: QuickValidateOpenAI - ECCEZIONE: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Test veloce validità chiave Google Places (timeout 3 secondi)
    ''' </summary>
    Private Async Function QuickValidateGooglePlaces(apiKey As String) As Task(Of Boolean)
        Try
            Debug.WriteLine($"DEBUG: QuickValidateGooglePlaces - Inizio test con chiave: {apiKey.Substring(0, Math.Min(10, apiKey.Length))}...")
            Using client As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(3)}
                ' Test minimo: ricerca vuota per verificare autenticazione
                Dim url = $"https://places.googleapis.com/v1/places:searchText"
                Dim requestBody = New With {.textQuery = "test"}
                Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(requestBody)
                Dim content = New StringContent(json, Encoding.UTF8, "application/json")

                client.DefaultRequestHeaders.Add("X-Goog-Api-Key", apiKey)
                client.DefaultRequestHeaders.Add("X-Goog-FieldMask", "places.displayName")

                Debug.WriteLine($"DEBUG: QuickValidateGooglePlaces - Chiamata POST a Places API...")
                Dim response = Await client.PostAsync(url, content)
                Debug.WriteLine($"DEBUG: QuickValidateGooglePlaces - Risposta: {response.StatusCode}")
                ' Anche se la ricerca fallisce, se l'autenticazione è OK non dovremmo avere 401
                Dim isValid = response.StatusCode <> HttpStatusCode.Unauthorized
                Debug.WriteLine($"DEBUG: QuickValidateGooglePlaces - Risultato: {isValid}")
                Return isValid
            End Using
        Catch ex As Exception
            Debug.WriteLine($"DEBUG: QuickValidateGooglePlaces - ECCEZIONE: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Mostra finestra configurazione Google Places
    ''' </summary>

#End Region

#Region "=== CLASSIFICAZIONE PRINCIPALE ==="

    ''' <summary>
    ''' Avvia classificazione completa
    ''' </summary>
    Private Async Sub BtnStartClassification_Click(sender As Object, e As RoutedEventArgs)
        ' Mostra dialog di conferma con riepilogo delle scelte
        If Await ShowClassificationConfirmationDialog() Then
            Await ExecuteClassificationWithSelectedOptions()
        End If
    End Sub

    ''' <summary>
    ''' Mostra dialog di conferma con riepilogo delle scelte utente
    ''' </summary>
    Private Async Function ShowClassificationConfirmationDialog() As Task(Of Boolean)
        Try
            Dim metodo = GetSelectedClassificationMethod()
            Dim scope = GetSelectedClassificationScope()
            Dim metodoDesc = GetMethodDisplayDescription()
            Dim scopeDesc = GetScopeDisplayDescription()

            Debug.WriteLine($"DEBUG: ===== INIZIO CLASSIFICAZIONE =====")
            Debug.WriteLine($"DEBUG: Metodo selezionato: '{metodo}' ({metodoDesc})")
            Debug.WriteLine($"DEBUG: Scope selezionato: '{scope}' ({scopeDesc})")

            ' CONTROLLO API KEYS - Prima di tutto verifica se il metodo selezionato necessita di API keys
            If metodo = "AIEnhanced" Then
                Debug.WriteLine($"DEBUG: Metodo AIEnhanced selezionato - controllo API keys...")
                Dim apiValidationResult = Await ValidateAPIKeysAsync()

                If Not apiValidationResult.IsValid Then
                    ' API keys mancanti o non valide
                    Dim apiResult = MessageBox.Show(
                        $"⚠️ {apiValidationResult.Title}" & vbCrLf & vbCrLf &
                        apiValidationResult.Message & vbCrLf & vbCrLf &
                        "💡 Soluzioni:" & vbCrLf &
                        "• Configura chiavi API valide nelle Impostazioni" & vbCrLf &
                        "• Usa il metodo 'Pattern Only' per classificazione offline" & vbCrLf & vbCrLf &
                        "Vuoi aprire le Impostazioni per correggere la configurazione?",
                        "Problema API Keys",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning)

                    If apiResult = MessageBoxResult.Yes Then
                        ' Apri impostazioni
                        Dim settingsWindow As New SettingsWindow()
                        settingsWindow.Show()
                        CheckAIConfiguration()
                    End If

                    Return False ' Non procedere con la classificazione
                End If
            End If

            ' Costruisci messaggio di conferma
            Dim message As New StringBuilder()
            message.AppendLine("🎯 CONFERMA CLASSIFICAZIONE")
            message.AppendLine()
            message.AppendLine("Le tue scelte:")
            message.AppendLine($"📋 Metodo: {metodoDesc}")
            message.AppendLine($"🔍 Scope: {scopeDesc}")
            message.AppendLine()

            ' Aggiungi avvisi specifici
            If metodo = "ManualOnly" Then
                message.AppendLine("⚠️ Verrà aperto il pannello di classificazione manuale.")
            ElseIf scope = "Single" Then
                message.AppendLine("🎯 Verrà aperto il pannello per testare una singola transazione.")
            ElseIf scope = "All" Then
                message.AppendLine("⚠️ ATTENZIONE: Tutte le classificazioni esistenti verranno rimosse!")
                message.AppendLine("   Questa operazione potrebbe richiedere diversi minuti.")
            Else
                message.AppendLine("✅ Verranno classificate solo le transazioni non ancora elaborate.")
            End If

            message.AppendLine()
            message.AppendLine("Vuoi procedere con la classificazione?")

            Dim result = MessageBox.Show(
                message.ToString(),
                "Conferma Classificazione",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question)

            Return result = MessageBoxResult.Yes

        Catch ex As Exception
            _logger?.LogError(ex, "Error showing classification confirmation dialog")
            ShowError("Errore", "Errore durante la preparazione della classificazione: " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Esegue classificazione basata su metodo e scope selezionati dall'utente
    ''' </summary>
    Private Async Function ExecuteClassificationWithSelectedOptions() As Task
        Try
            Dim metodo = GetSelectedClassificationMethod()
            Dim scope = GetSelectedClassificationScope()

            Debug.WriteLine($"DEBUG: === ExecuteClassificationWithSelectedOptions ===")
            Debug.WriteLine($"DEBUG: Metodo: '{metodo}', Scope: '{scope}'")

            ' Se è selezionato ManualOnly, apri pannello manuale indipendentemente dallo scope
            If metodo = "ManualOnly" Then
                Debug.WriteLine($"DEBUG: Metodo ManualOnly - aprendo pannello manuale...")
                Await OpenManualClassificationPanel()
                Return
            End If

            ' Esegui classificazione basata su scope
            Debug.WriteLine($"DEBUG: Esecuzione scope '{scope}'...")
            Select Case scope
                Case "All"
                    Debug.WriteLine($"DEBUG: SCOPE ALL - Riclassificazione di tutte le transazioni...")
                    Await ExecuteReclassifyAllWithMethod(metodo)
                Case "Unmatched"
                    Debug.WriteLine($"DEBUG: SCOPE UNMATCHED - Classificazione transazioni non classificate...")
                    Await ExecuteClassifyUnmatchedWithMethod(metodo)
                Case "Single"
                    Debug.WriteLine($"DEBUG: SCOPE SINGLE - Classificazione singola transazione...")
                    Await OpenSingleClassificationPanel(metodo)
                Case "Test"
                    Debug.WriteLine($"DEBUG: SCOPE TEST - Apertura pannello test transazione...")
                    Await OpenTestTransactionPanel(metodo)
                Case Else
                    Debug.WriteLine($"DEBUG: SCOPE DEFAULT - Fallback a Unmatched...")
                    Await ExecuteClassifyUnmatchedWithMethod(metodo) ' Default fallback
            End Select

            Debug.WriteLine($"DEBUG: === Fine ExecuteClassificationWithSelectedOptions ===")

        Catch ex As Exception
            _logger?.LogError(ex, "Error executing classification with selected options")
            ShowError("Errore Classificazione", "Errore durante la classificazione: " & ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Riclassifica tutte le transazioni con metodo specificato
    ''' </summary>
    Private Async Function ExecuteReclassifyAllWithMethod(metodo As String) As Task
        Dim startTime = DateTime.Now

        Try
            Debug.WriteLine($"DEBUG: === ExecuteReclassifyAllWithMethod START ===")
            Debug.WriteLine($"DEBUG: Metodo: '{metodo}'")

            ' Reset statistiche AI se necessario
            If metodo = "AIEnhanced" AndAlso _gptClassificatore IsNot Nothing Then
                _gptClassificatore.ResetSessionStats()
                Debug.WriteLine($"DEBUG: Statistiche AI resettate per nuova sessione")
            End If

            ' Setup UI
            ProgressPanel.Visibility = Visibility.Visible
            ResultsPanel.Visibility = Visibility.Collapsed
            PreviewPanel.Visibility = Visibility.Collapsed

            ' Carica tutte le transazioni (anche quelle già classificate)
            Debug.WriteLine($"DEBUG: Caricamento di tutte le transazioni...")
            Await LoadAllTransactions()
            Debug.WriteLine($"DEBUG: Transazioni caricate: {_transazioni.Count}")

            If _transazioni.Count = 0 Then
                Debug.WriteLine($"DEBUG: Nessuna transazione trovata, uscita anticipata")
                ShowInfo("Nessuna transazione trovata nel database.")
                Return
            End If

            ' NOTA: Reset classificazioni spostato all'approvazione dell'anteprima
            ' per evitare perdita dati se l'utente non approva
            _requiresResetBeforeApproval = True
            Debug.WriteLine($"DEBUG: Reset classificazioni rimandata all'approvazione utente")

            ' Prepara lista risultati anteprima
            _previewResults.Clear()
            Dim tempResults As New List(Of TransactionPreview)
            Debug.WriteLine($"DEBUG: Lista anteprima pulita, inizio classificazione...")

            LblProgressStatus.Text = $"Riclassificazione in corso... (0/{_transazioni.Count})"

            ' Riclassifica ogni transazione con il metodo specificato
            For i As Integer = 0 To _transazioni.Count - 1
                Dim transazione = _transazioni(i)

                ' Aggiorna progress
                Dim percentuale = ((i + 1) / _transazioni.Count) * 100
                ProgressBarClassification.Value = percentuale
                LblProgressStatus.Text = $"Riclassificazione in corso... ({i + 1}/{_transazioni.Count})"
                LblProgressDetail.Text = $"Elaborando: {TruncateString(transazione.Descrizione, 40)}"

                ' Esegui classificazione basata sul metodo specificato
                Dim risultato = Await ExecuteClassificationByMethod(metodo, transazione)

                ' Aggiungi alla anteprima se valido
                If risultato IsNot Nothing AndAlso Not String.IsNullOrEmpty(risultato.MacroCategoria) Then
                    tempResults.Add(New TransactionPreview With {
                        .ID = transazione.ID,
                        .Descrizione = TruncateString(transazione.Descrizione, 80),
                        .ImportoFormatted = transazione.Importo.ToString("C2"),
                        .ParolaChiave = If(Not String.IsNullOrEmpty(risultato.ParolaChiave), risultato.ParolaChiave, If(_gptClassificatore IsNot Nothing, _gptClassificatore.EstraiParolaChiaveDaDescrizione(transazione.Descrizione), "")),
                        .MacroCategoria = risultato.MacroCategoria,
                        .Categoria = risultato.Categoria,
                        .Necessita = risultato.Necessita,
                        .Frequenza = risultato.Frequenza,
                        .Stagionalita = risultato.Stagionalita,
                        .ConfidenzaFormatted = $"{risultato.PunteggioConfidenza:P1}",
                        .MetodoClassificazione = GetDisplayNameFromMethod(metodo, risultato.Motivazione),
                        .OriginalTransaction = transazione,
                        .IsSelected = True
                    })
                End If

                ' Pausa per aggiornamento UI
                Await Task.Delay(30)
            Next

            ' Trasferisci risultati alla ObservableCollection sul thread UI
            Dispatcher.Invoke(Sub()
                For Each result In tempResults
                    _previewResults.Add(result)
                Next
            End Sub)

            ' Mostra anteprima
            Dim elapsed = DateTime.Now - startTime
            LblProgressStatus.Text = $"Riclassificazione completata! {_previewResults.Count} classificazioni proposte"
            LblProgressDetail.Text = $"Tempo di analisi: {elapsed.TotalSeconds:F1} secondi"

            ' Mostra riepilogo costi AI se applicabile
            If metodo = "AIEnhanced" AndAlso _gptClassificatore IsNot Nothing AndAlso _gptClassificatore.HasSessionUsage() Then
                Dim costSummary = _gptClassificatore.GetSessionSummary()
                Debug.WriteLine($"DEBUG: === RIEPILOGO COSTI AI ===")
                Debug.WriteLine(costSummary)
                ShowAICostSummary(costSummary)
            End If

            If _previewResults.Count > 0 Then
                ShowPreviewResults()
            Else
                ShowInfo("Nessuna transazione è stata classificata con successo.")
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error in ExecuteReclassifyAllWithMethod")
            ShowError("Errore Riclassificazione", "Errore durante la riclassificazione: " & ex.Message)
        Finally
            ProgressBarClassification.Value = 100
        End Try
    End Function

    ''' <summary>
    ''' Classifica solo transazioni non classificate con metodo specificato
    ''' </summary>
    Private Async Function ExecuteClassifyUnmatchedWithMethod(metodo As String) As Task
        Dim startTime = DateTime.Now

        Try
            Debug.WriteLine($"DEBUG: === ExecuteClassifyUnmatchedWithMethod START ===")
            Debug.WriteLine($"DEBUG: Metodo: '{metodo}'")

            ' Questo non è "All scope", quindi nessun reset necessario
            _requiresResetBeforeApproval = False

            ' Reset statistiche AI se necessario
            If metodo = "AIEnhanced" AndAlso _gptClassificatore IsNot Nothing Then
                _gptClassificatore.ResetSessionStats()
                Debug.WriteLine($"DEBUG: Statistiche AI resettate per nuova sessione")
            End If

            ' Setup UI
            ProgressPanel.Visibility = Visibility.Visible
            ResultsPanel.Visibility = Visibility.Collapsed
            PreviewPanel.Visibility = Visibility.Collapsed

            ' Carica solo transazioni non classificate
            Debug.WriteLine($"DEBUG: Caricamento transazioni non classificate...")
            Await LoadUnmatchedTransactions()
            Debug.WriteLine($"DEBUG: Transazioni non classificate caricate: {_transazioni.Count}")

            If _transazioni.Count = 0 Then
                Debug.WriteLine($"DEBUG: Tutte le transazioni già classificate, uscita anticipata")
                ShowInfo("Tutte le transazioni sono già classificate!")
                Return
            End If

            LblProgressStatus.Text = $"Classificazione transazioni non classificate... (0/{_transazioni.Count})"

            ' Filtra solo le transazioni non classificate
            Dim transazioniDaClassificare = _transazioni.Where(
                Function(t) String.IsNullOrEmpty(t.MacroCategoria) OrElse t.MacroCategoria = "Non Classificata"
            ).ToList()

            Debug.WriteLine($"DEBUG: Transazioni non classificate caricate: {transazioniDaClassificare.Count}")

            If transazioniDaClassificare.Count = 0 Then
                ShowInfo("Tutte le transazioni sono già classificate!")
                Return
            End If

            ' Prepara lista risultati anteprima
            _previewResults.Clear()
            Dim tempResults As New List(Of TransactionPreview)
            Debug.WriteLine($"DEBUG: Lista anteprima pulita, inizio classificazione...")

            ' Classifica ogni transazione non classificata con il metodo specificato
            For i As Integer = 0 To transazioniDaClassificare.Count - 1
                Dim transazione = transazioniDaClassificare(i)

                ' Aggiorna progress
                Dim percentuale = ((i + 1) / transazioniDaClassificare.Count) * 100
                ProgressBarClassification.Value = percentuale
                LblProgressStatus.Text = $"Analisi in corso... ({i + 1}/{transazioniDaClassificare.Count})"
                LblProgressDetail.Text = $"Elaborando: {TruncateString(transazione.Descrizione, 40)}"

                ' Esegui classificazione
                Dim risultato As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione = Nothing

                Debug.WriteLine($"DEBUG: 🎬 INIZIO CHIAMATA METODO '{metodo}' per ID: {transazione.ID}")

                Select Case metodo
                    Case "PatternOnly"
                        risultato = _classificatore.ClassificaTransazione(transazione)
                        Debug.WriteLine($"DEBUG: 🎯 PATTERN ONLY COMPLETATO per ID: {transazione.ID}")
                    Case "AIEnhanced"
                        risultato = Await ExecuteAIEnhancedClassification(transazione)
                        Debug.WriteLine($"DEBUG: 🤖 AI ENHANCED COMPLETATO per ID: {transazione.ID} - Risultato IsNothing: {risultato Is Nothing}")
                        If risultato IsNot Nothing Then
                            Debug.WriteLine($"DEBUG: 🤖 AI ENHANCED RESULT - MacroCategoria: '{risultato.MacroCategoria}', Categoria: '{risultato.Categoria}', Motivazione: '{risultato.Motivazione}'")
                        End If
                    Case "OfflineOnly"
                        risultato = ExecuteOfflineClassification(transazione)
                        Debug.WriteLine($"DEBUG: 📴 OFFLINE ONLY COMPLETATO per ID: {transazione.ID}")
                    Case "ManualOnly"
                        ' Questo non dovrebbe mai essere raggiunto, ma fallback offline
                        risultato = ExecuteOfflineClassification(transazione)
                        risultato.Motivazione = "Fallback da modalità manuale"
                        Debug.WriteLine($"DEBUG: ✋ MANUAL ONLY COMPLETATO per ID: {transazione.ID}")
                End Select

                Debug.WriteLine($"DEBUG: 🏁 FINE CHIAMATA METODO '{metodo}' per ID: {transazione.ID} - Passando al controllo anteprima...")

                ' DEBUG: Controllo risultato per anteprima UNMATCHED
                Debug.WriteLine($"DEBUG: CONTROLLO ANTEPRIMA UNMATCHED - ID: {transazione.ID}, Descrizione: '{TruncateString(transazione.Descrizione, 50)}'")
                Debug.WriteLine($"DEBUG: CONTROLLO ANTEPRIMA UNMATCHED - Risultato IsNothing: {risultato Is Nothing}")
                If risultato IsNot Nothing Then
                    Debug.WriteLine($"DEBUG: CONTROLLO ANTEPRIMA UNMATCHED - MacroCategoria: '{risultato.MacroCategoria}', IsEmpty: {String.IsNullOrEmpty(risultato.MacroCategoria)}")
                    Debug.WriteLine($"DEBUG: CONTROLLO ANTEPRIMA UNMATCHED - Categoria: '{risultato.Categoria}', ParolaChiave: '{risultato.ParolaChiave}'")
                End If

                ' Aggiungi alla anteprima se valido
                If risultato IsNot Nothing AndAlso Not String.IsNullOrEmpty(risultato.MacroCategoria) Then
                    Debug.WriteLine($"DEBUG: ✅ AGGIUNGENDO A ANTEPRIMA UNMATCHED - ID: {transazione.ID}, MacroCategoria: '{risultato.MacroCategoria}', Categoria: '{risultato.Categoria}'")

                    tempResults.Add(New TransactionPreview With {
                        .ID = transazione.ID,
                        .Descrizione = TruncateString(transazione.Descrizione, 80),
                        .ImportoFormatted = transazione.Importo.ToString("C2"),
                        .ParolaChiave = If(Not String.IsNullOrEmpty(risultato.ParolaChiave), risultato.ParolaChiave, If(_gptClassificatore IsNot Nothing, _gptClassificatore.EstraiParolaChiaveDaDescrizione(transazione.Descrizione), "")),
                        .MacroCategoria = risultato.MacroCategoria,
                        .Categoria = risultato.Categoria,
                        .Necessita = risultato.Necessita,
                        .Frequenza = risultato.Frequenza,
                        .Stagionalita = risultato.Stagionalita,
                        .ConfidenzaFormatted = $"{risultato.PunteggioConfidenza:P1}",
                        .MetodoClassificazione = GetDisplayNameFromMethod(metodo, risultato.Motivazione),
                        .OriginalTransaction = transazione,
                        .IsSelected = True
                    })
                    Debug.WriteLine($"DEBUG: ✅ AGGIUNTA CONFERMATA UNMATCHED - tempResults.Count ora: {tempResults.Count}")
                Else
                    Debug.WriteLine($"DEBUG: ❌ NON AGGIUNGENDO A ANTEPRIMA UNMATCHED - ID: {transazione.ID}, Motivo: risultato={risultato Is Nothing}, MacroCategoria vuota={If(risultato IsNot Nothing, String.IsNullOrEmpty(risultato.MacroCategoria).ToString(), "N/A")}")
                End If

                ' Pausa per aggiornamento UI
                Await Task.Delay(30)
            Next

            Debug.WriteLine($"DEBUG: 🔚 FINE CICLO UNMATCHED - tempResults.Count totale: {tempResults.Count}")

            ' Trasferisci risultati alla ObservableCollection sul thread UI
            Dispatcher.Invoke(Sub()
                For Each result In tempResults
                    _previewResults.Add(result)
                Next
            End Sub)

            Debug.WriteLine($"DEBUG: 🎯 TRASFERIMENTO COMPLETATO - _previewResults.Count: {_previewResults.Count}")

            ' Mostra anteprima
            Dim elapsed = DateTime.Now - startTime
            LblProgressStatus.Text = $"Classificazione completata! {_previewResults.Count} nuove classificazioni proposte"
            LblProgressDetail.Text = $"Tempo di analisi: {elapsed.TotalSeconds:F1} secondi"

            ' Mostra riepilogo costi AI se applicabile
            If metodo = "AIEnhanced" AndAlso _gptClassificatore IsNot Nothing AndAlso _gptClassificatore.HasSessionUsage() Then
                Dim costSummary = _gptClassificatore.GetSessionSummary()
                Debug.WriteLine($"DEBUG: === RIEPILOGO COSTI AI ===")
                Debug.WriteLine(costSummary)
                ShowAICostSummary(costSummary)
            End If

            If _previewResults.Count > 0 Then
                ShowPreviewResults()
            Else
                ShowInfo("Nessuna transazione non classificata è stata classificata con successo.")
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error in ExecuteClassifyUnmatchedWithMethod")
            ShowError("Errore Classificazione", "Errore durante la classificazione: " & ex.Message)
        Finally
            ProgressBarClassification.Value = 100
        End Try
    End Function

    ''' <summary>
    ''' Esegue classificazione singola transazione con metodo specificato
    ''' </summary>
    Private Async Function ExecuteClassificationByMethod(metodo As String, transazione As Transazione) As Task(Of MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione)
        Select Case metodo
            Case "PatternOnly"
                Return _classificatore.ClassificaTransazione(transazione)
            Case "AIEnhanced"
                Return Await ExecuteAIEnhancedClassification(transazione)
            Case "OfflineOnly"
                Return ExecuteOfflineClassification(transazione)
            Case "ManualOnly"
                ' Questo non dovrebbe mai essere raggiunto, ma fallback offline
                Dim risultato = ExecuteOfflineClassification(transazione)
                risultato.Motivazione = "Fallback da modalità manuale"
                Return risultato
            Case Else
                Return Await ExecuteAIEnhancedClassification(transazione) ' Default fallback
        End Select
    End Function

    ''' <summary>
    ''' Riclassifica tutte le transazioni (anche quelle già classificate)
    ''' </summary>
    Private Async Sub BtnReclassifyAll_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim result = MessageBox.Show(
                "Sei sicuro di voler riclassificare TUTTE le transazioni?" & vbCrLf & vbCrLf &
                "Questa operazione:" & vbCrLf &
                "• Rimuoverà le classificazioni esistenti" & vbCrLf &
                "• Riapplicherà i pattern aggiornati" & vbCrLf &
                "• Potrebbe modificare classificazioni già corrette" & vbCrLf & vbCrLf &
                "L'operazione potrebbe richiedere diversi minuti.",
                "Conferma Riclassificazione Totale",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning)

            If result = MessageBoxResult.Yes Then
                Await ExecuteReclassifyAll()
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error in reclassify all button")
            ShowError("Errore Riclassificazione", "Errore durante l'avvio della riclassificazione: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Classifica solo le transazioni non ancora classificate
    ''' </summary>
    Private Async Sub BtnClassifyUnmatched_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim result = MessageBox.Show(
                "Vuoi classificare tutte le transazioni senza match?" & vbCrLf & vbCrLf &
                "Questa operazione classificherà solo le transazioni che:" & vbCrLf &
                "• Non hanno MacroCategoria assegnata" & vbCrLf &
                "• Non hanno Categoria assegnata" & vbCrLf &
                "• Sono rimaste senza classificazione" & vbCrLf & vbCrLf &
                "Le classificazioni esistenti non verranno modificate.",
                "Conferma Classificazione Non Classificate",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question)

            If result = MessageBoxResult.Yes Then
                Await ExecuteClassifyUnmatched()
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error in classify unmatched button")
            ShowError("Errore Classificazione", "Errore durante l'avvio della classificazione: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Apre la sezione di classificazione manuale
    ''' </summary>
    Private Async Sub BtnManualClassification_Click(sender As Object, e As RoutedEventArgs)
        Try
            ' Nascondi altri pannelli
            ProgressPanel.Visibility = Visibility.Collapsed
            ResultsPanel.Visibility = Visibility.Collapsed
            PreviewPanel.Visibility = Visibility.Collapsed

            ' Mostra pannello manuale
            ManualClassificationPanel.Visibility = Visibility.Visible

            ' Carica transazioni non classificate
            Await LoadManualTransactions()

            ' Popola combo
            PopulateManualComboBoxes()

            ' Scroll automatico verso il pannello manuale dopo un breve delay per permettere il rendering
            Await Task.Delay(200)
            ScrollToManualPanel()

        Catch ex As Exception
            _logger?.LogError(ex, "Error opening manual classification")
            ShowError("Errore", "Errore durante l'apertura della classificazione manuale: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Apre il pannello di classificazione manuale (usato quando è selezionato ManualOnly)
    ''' </summary>
    Private Async Function OpenManualClassificationPanel() As Task
        Try
            ' Nascondi altri pannelli
            ProgressPanel.Visibility = Visibility.Collapsed
            ResultsPanel.Visibility = Visibility.Collapsed
            PreviewPanel.Visibility = Visibility.Collapsed

            ' Mostra pannello manuale
            ManualClassificationPanel.Visibility = Visibility.Visible

            ' Carica transazioni non classificate
            Await LoadManualTransactions()

            ' Popola combo
            PopulateManualComboBoxes()

            ' Scroll automatico verso il pannello manuale dopo un breve delay per permettere il rendering
            Await Task.Delay(200)
            ScrollToManualPanel()

        Catch ex As Exception
            _logger?.LogError(ex, "Error opening manual classification panel")
            ShowError("Errore", "Errore durante l'apertura della classificazione manuale: " & ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Riclassifica tutte le transazioni (rimuove classificazioni esistenti e riapplica pattern)
    ''' </summary>
    Private Async Function ExecuteReclassifyAll() As Task
        Dim startTime = DateTime.Now

        Try
            ' Setup UI
            ProgressPanel.Visibility = Visibility.Visible
            ResultsPanel.Visibility = Visibility.Collapsed
            PreviewPanel.Visibility = Visibility.Collapsed

            ' Carica tutte le transazioni (anche quelle già classificate)
            Await LoadAllTransactions()

            If _transazioni.Count = 0 Then
                ShowInfo("Nessuna transazione trovata nel database.")
                Return
            End If

            ' Reset classificazioni esistenti nel database
            LblProgressStatus.Text = "Rimozione classificazioni esistenti..."
            Await ResetAllClassifications()

            ' Prepara lista risultati anteprima
            _previewResults.Clear()
            Dim tempResults As New List(Of TransactionPreview)
            Dim metodo = GetSelectedClassificationMethod()

            ' Se è selezionato ManualOnly, apri pannello manuale invece di classificazione automatica
            If metodo = "ManualOnly" Then
                Await OpenManualClassificationPanel()
                Return
            End If

            LblProgressStatus.Text = $"Riclassificazione in corso... (0/{_transazioni.Count})"

            ' Riclassifica ogni transazione
            For i As Integer = 0 To _transazioni.Count - 1
                Dim transazione = _transazioni(i)

                ' Aggiorna progress
                Dim percentuale = ((i + 1) / _transazioni.Count) * 100
                ProgressBarClassification.Value = percentuale
                LblProgressStatus.Text = $"Riclassificazione in corso... ({i + 1}/{_transazioni.Count})"
                LblProgressDetail.Text = $"Elaborando: {TruncateString(transazione.Descrizione, 40)}"

                ' Esegui classificazione basata sul metodo selezionato
                Dim risultato As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione = Nothing

                Select Case metodo
                    Case "PatternOnly"
                        risultato = _classificatore.ClassificaTransazione(transazione)
                    Case "AIEnhanced"
                        risultato = Await ExecuteAIEnhancedClassification(transazione)
                    Case "OfflineOnly"
                        risultato = ExecuteOfflineClassification(transazione)
                    Case "ManualOnly"
                        ' Questo non dovrebbe mai essere raggiunto, ma fallback offline
                        risultato = ExecuteOfflineClassification(transazione)
                        risultato.Motivazione = "Fallback da modalità manuale"
                End Select

                ' Aggiungi alla anteprima se valido
                If risultato IsNot Nothing AndAlso Not String.IsNullOrEmpty(risultato.MacroCategoria) Then
                    tempResults.Add(New TransactionPreview With {
                        .ID = transazione.ID,
                        .Descrizione = TruncateString(transazione.Descrizione, 80),
                        .ImportoFormatted = transazione.Importo.ToString("C2"),
                        .ParolaChiave = If(Not String.IsNullOrEmpty(risultato.ParolaChiave), risultato.ParolaChiave, If(_gptClassificatore IsNot Nothing, _gptClassificatore.EstraiParolaChiaveDaDescrizione(transazione.Descrizione), "")),
                        .MacroCategoria = risultato.MacroCategoria,
                        .Categoria = risultato.Categoria,
                        .Necessita = risultato.Necessita,
                        .Frequenza = risultato.Frequenza,
                        .Stagionalita = risultato.Stagionalita,
                        .ConfidenzaFormatted = $"{risultato.PunteggioConfidenza:P1}",
                        .MetodoClassificazione = GetDisplayNameFromMethod(metodo, risultato.Motivazione),
                        .OriginalTransaction = transazione,
                        .IsSelected = True
                    })
                End If

                ' Pausa per aggiornamento UI
                Await Task.Delay(30)
            Next

            ' Mostra anteprima
            Dim elapsed = DateTime.Now - startTime
            ' Trasferisci risultati alla ObservableCollection sul thread UI
            Dispatcher.Invoke(Sub()
                For Each result In tempResults
                    _previewResults.Add(result)
                Next
            End Sub)

            LblProgressStatus.Text = $"Riclassificazione completata! {_previewResults.Count} classificazioni proposte"
            LblProgressDetail.Text = $"Tempo di analisi: {elapsed.TotalSeconds:F1} secondi"

            If _previewResults.Count > 0 Then
                ShowPreviewResults()
            Else
                ShowInfo("Nessuna transazione è stata classificata con il metodo selezionato.")
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error in ExecuteReclassifyAll")
            ShowError("Errore Riclassificazione", "Errore durante la riclassificazione: " & ex.Message)
        Finally
            ProgressBarClassification.Value = 100
        End Try
    End Function

    ''' <summary>
    ''' Classifica solo le transazioni senza classificazione esistente
    ''' </summary>
    Private Async Function ExecuteClassifyUnmatched() As Task
        Dim startTime = DateTime.Now

        Try
            ' Setup UI
            ProgressPanel.Visibility = Visibility.Visible
            ResultsPanel.Visibility = Visibility.Collapsed
            PreviewPanel.Visibility = Visibility.Collapsed

            ' Carica solo transazioni non classificate
            Await LoadUnmatchedTransactions()

            If _transazioni.Count = 0 Then
                ShowInfo("Tutte le transazioni sono già classificate!")
                Return
            End If

            LblProgressStatus.Text = $"Classificazione transazioni non classificate... (0/{_transazioni.Count})"

            ' Prepara lista risultati anteprima
            _previewResults.Clear()
            Dim tempResults As New List(Of TransactionPreview)
            Dim metodo = GetSelectedClassificationMethod()

            ' Se è selezionato ManualOnly, apri pannello manuale invece di classificazione automatica
            If metodo = "ManualOnly" Then
                Await OpenManualClassificationPanel()
                Return
            End If

            ' Classifica ogni transazione non classificata
            For i As Integer = 0 To _transazioni.Count - 1
                Dim transazione = _transazioni(i)

                ' Aggiorna progress
                Dim percentuale = ((i + 1) / _transazioni.Count) * 100
                ProgressBarClassification.Value = percentuale
                LblProgressStatus.Text = $"Classificazione in corso... ({i + 1}/{_transazioni.Count})"
                LblProgressDetail.Text = $"Elaborando: {TruncateString(transazione.Descrizione, 40)}"

                ' Esegui classificazione basata sul metodo selezionato
                Dim risultato As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione = Nothing

                Select Case metodo
                    Case "PatternOnly"
                        risultato = _classificatore.ClassificaTransazione(transazione)
                    Case "AIEnhanced"
                        risultato = Await ExecuteAIEnhancedClassification(transazione)
                    Case "OfflineOnly"
                        risultato = ExecuteOfflineClassification(transazione)
                    Case "ManualOnly"
                        ' Questo non dovrebbe mai essere raggiunto, ma fallback offline
                        risultato = ExecuteOfflineClassification(transazione)
                        risultato.Motivazione = "Fallback da modalità manuale"
                End Select

                ' Aggiungi alla anteprima se valido
                If risultato IsNot Nothing AndAlso Not String.IsNullOrEmpty(risultato.MacroCategoria) Then
                    tempResults.Add(New TransactionPreview With {
                        .ID = transazione.ID,
                        .Descrizione = TruncateString(transazione.Descrizione, 80),
                        .ImportoFormatted = transazione.Importo.ToString("C2"),
                        .ParolaChiave = If(Not String.IsNullOrEmpty(risultato.ParolaChiave), risultato.ParolaChiave, If(_gptClassificatore IsNot Nothing, _gptClassificatore.EstraiParolaChiaveDaDescrizione(transazione.Descrizione), "")),
                        .MacroCategoria = risultato.MacroCategoria,
                        .Categoria = risultato.Categoria,
                        .Necessita = risultato.Necessita,
                        .Frequenza = risultato.Frequenza,
                        .Stagionalita = risultato.Stagionalita,
                        .ConfidenzaFormatted = $"{risultato.PunteggioConfidenza:P1}",
                        .MetodoClassificazione = GetDisplayNameFromMethod(metodo, risultato.Motivazione),
                        .OriginalTransaction = transazione,
                        .IsSelected = True
                    })
                End If

                ' Pausa per aggiornamento UI
                Await Task.Delay(30)
            Next

            ' Trasferisci risultati alla ObservableCollection sul thread UI
            Dispatcher.Invoke(Sub()
                For Each result In tempResults
                    _previewResults.Add(result)
                Next
            End Sub)

            ' Mostra anteprima
            Dim elapsed = DateTime.Now - startTime
            LblProgressStatus.Text = $"Classificazione completata! {_previewResults.Count} nuove classificazioni proposte"
            LblProgressDetail.Text = $"Tempo di analisi: {elapsed.TotalSeconds:F1} secondi"

            If _previewResults.Count > 0 Then
                ShowPreviewResults()
            Else
                ShowInfo("Nessuna nuova transazione è stata classificata con il metodo selezionato.")
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error in ExecuteClassifyUnmatched")
            ShowError("Errore Classificazione", "Errore durante la classificazione: " & ex.Message)
        Finally
            ProgressBarClassification.Value = 100
        End Try
    End Function

    ''' <summary>
    ''' Esegue classificazione completa con anteprima e approvazione utente
    ''' </summary>
    Private Async Function ExecuteFullClassification() As Task
        Dim startTime = DateTime.Now

        Try
            ' Setup UI
            ProgressPanel.Visibility = Visibility.Visible
            BtnStartClassification.IsEnabled = False
            ProgressBarClassification.Value = 0
            LblProgressStatus.Text = "Inizializzazione classificatore..."

            ' Ricarica transazioni
            Await LoadTransactions()

            ' Filtra transazioni da classificare
            Dim transazioniDaClassificare = _transazioni.Where(
                Function(t) String.IsNullOrEmpty(t.MacroCategoria) OrElse t.MacroCategoria = "Non Classificata"
            ).ToList()

            If transazioniDaClassificare.Count = 0 Then
                ShowInfo("Tutte le transazioni sono già classificate!")
                Return
            End If

            LblProgressStatus.Text = $"Analisi in corso... (0/{transazioniDaClassificare.Count})"

            ' Prepara lista risultati anteprima
            _previewResults.Clear()
            Dim tempResults As New List(Of TransactionPreview)
            Dim metodo = GetSelectedClassificationMethod()

            ' Se è selezionato ManualOnly, apri pannello manuale invece di classificazione automatica
            If metodo = "ManualOnly" Then
                Await OpenManualClassificationPanel()
                Return
            End If

            ' Classificazione per ogni transazione (solo anteprima)
            For i As Integer = 0 To transazioniDaClassificare.Count - 1
                Dim transazione = transazioniDaClassificare(i)

                ' Aggiorna progress
                Dim percentuale = ((i + 1) / transazioniDaClassificare.Count) * 100
                ProgressBarClassification.Value = percentuale
                LblProgressStatus.Text = $"Analisi in corso... ({i + 1}/{transazioniDaClassificare.Count})"
                LblProgressDetail.Text = $"Elaborando: {TruncateString(transazione.Descrizione, 40)}"

                ' Esegui classificazione
                Dim risultato As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione = Nothing

                Debug.WriteLine($"DEBUG: 🎬 INIZIO CHIAMATA METODO '{metodo}' per ID: {transazione.ID}")

                Select Case metodo
                    Case "PatternOnly"
                        risultato = _classificatore.ClassificaTransazione(transazione)
                        Debug.WriteLine($"DEBUG: 🎯 PATTERN ONLY COMPLETATO per ID: {transazione.ID}")
                    Case "AIEnhanced"
                        risultato = Await ExecuteAIEnhancedClassification(transazione)
                        Debug.WriteLine($"DEBUG: 🤖 AI ENHANCED COMPLETATO per ID: {transazione.ID} - Risultato IsNothing: {risultato Is Nothing}")
                        If risultato IsNot Nothing Then
                            Debug.WriteLine($"DEBUG: 🤖 AI ENHANCED RESULT - MacroCategoria: '{risultato.MacroCategoria}', Categoria: '{risultato.Categoria}', Motivazione: '{risultato.Motivazione}'")
                        End If
                    Case "OfflineOnly"
                        risultato = ExecuteOfflineClassification(transazione)
                        Debug.WriteLine($"DEBUG: 📴 OFFLINE ONLY COMPLETATO per ID: {transazione.ID}")
                    Case "ManualOnly"
                        ' Questo non dovrebbe mai essere raggiunto, ma fallback offline
                        risultato = ExecuteOfflineClassification(transazione)
                        risultato.Motivazione = "Fallback da modalità manuale"
                        Debug.WriteLine($"DEBUG: ✋ MANUAL ONLY COMPLETATO per ID: {transazione.ID}")
                End Select

                Debug.WriteLine($"DEBUG: 🏁 FINE CHIAMATA METODO '{metodo}' per ID: {transazione.ID} - Passando al controllo anteprima...")

                ' DEBUG: Controllo risultato per anteprima UNMATCHED
                Debug.WriteLine($"DEBUG: CONTROLLO ANTEPRIMA UNMATCHED - ID: {transazione.ID}, Descrizione: '{TruncateString(transazione.Descrizione, 50)}'")
                Debug.WriteLine($"DEBUG: CONTROLLO ANTEPRIMA UNMATCHED - Risultato IsNothing: {risultato Is Nothing}")
                If risultato IsNot Nothing Then
                    Debug.WriteLine($"DEBUG: CONTROLLO ANTEPRIMA UNMATCHED - MacroCategoria: '{risultato.MacroCategoria}', IsEmpty: {String.IsNullOrEmpty(risultato.MacroCategoria)}")
                    Debug.WriteLine($"DEBUG: CONTROLLO ANTEPRIMA UNMATCHED - Categoria: '{risultato.Categoria}', ParolaChiave: '{risultato.ParolaChiave}'")
                End If

                ' Aggiungi alla anteprima se valido
                If risultato IsNot Nothing AndAlso Not String.IsNullOrEmpty(risultato.MacroCategoria) Then
                    Debug.WriteLine($"DEBUG: ✅ AGGIUNGENDO A ANTEPRIMA UNMATCHED - ID: {transazione.ID}, MacroCategoria: '{risultato.MacroCategoria}', Categoria: '{risultato.Categoria}'")

                    tempResults.Add(New TransactionPreview With {
                        .ID = transazione.ID,
                        .Descrizione = TruncateString(transazione.Descrizione, 80),
                        .ImportoFormatted = transazione.Importo.ToString("C2"),
                        .ParolaChiave = If(Not String.IsNullOrEmpty(risultato.ParolaChiave), risultato.ParolaChiave, If(_gptClassificatore IsNot Nothing, _gptClassificatore.EstraiParolaChiaveDaDescrizione(transazione.Descrizione), "")),
                        .MacroCategoria = risultato.MacroCategoria,
                        .Categoria = risultato.Categoria,
                        .Necessita = risultato.Necessita,
                        .Frequenza = risultato.Frequenza,
                        .Stagionalita = risultato.Stagionalita,
                        .ConfidenzaFormatted = $"{risultato.PunteggioConfidenza:P1}",
                        .MetodoClassificazione = GetDisplayNameFromMethod(metodo, risultato.Motivazione),
                        .OriginalTransaction = transazione,
                        .IsSelected = True
                    })
                    Debug.WriteLine($"DEBUG: ✅ AGGIUNTA CONFERMATA UNMATCHED - tempResults.Count ora: {tempResults.Count}")
                Else
                    Debug.WriteLine($"DEBUG: ❌ NON AGGIUNGENDO A ANTEPRIMA UNMATCHED - ID: {transazione.ID}, Motivo: risultato={risultato Is Nothing}, MacroCategoria vuota={If(risultato IsNot Nothing, String.IsNullOrEmpty(risultato.MacroCategoria).ToString(), "N/A")}")
                End If

                ' Pausa per aggiornamento UI
                Await Task.Delay(30)
            Next

            Debug.WriteLine($"DEBUG: 🔚 FINE CICLO UNMATCHED - tempResults.Count totale: {tempResults.Count}")

            ' Trasferisci risultati alla ObservableCollection sul thread UI
            Dispatcher.Invoke(Sub()
                For Each result In tempResults
                    _previewResults.Add(result)
                Next
            End Sub)

            Debug.WriteLine($"DEBUG: 🎯 TRASFERIMENTO COMPLETATO - _previewResults.Count: {_previewResults.Count}")

            ' Mostra anteprima
            Dim elapsed = DateTime.Now - startTime
            LblProgressStatus.Text = $"Analisi completata! {_previewResults.Count} classificazioni proposte"
            LblProgressDetail.Text = $"Tempo di analisi: {elapsed.TotalSeconds:F1} secondi"

            If _previewResults.Count > 0 Then
                ShowPreviewResults()
            Else
                ShowInfo("Nessuna transazione è stata classificata con il metodo selezionato.")
            End If

        Catch ex As Exception
            ShowError("Errore durante l'analisi", ex.Message)
        Finally
            ' Reset UI
            BtnStartClassification.IsEnabled = True
        End Try
    End Function

    ''' <summary>
    ''' Classificazione AI Enhanced (Pattern + AI + Google)
    ''' </summary>
    Private Async Function ExecuteAIEnhancedClassification(transazione As Transazione, Optional isTestMode As Boolean = False) As Task(Of MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione)
        Try
            Debug.WriteLine($"DEBUG: === INIZIO ExecuteAIEnhancedClassification per transazione ID: {transazione.ID} ===")
            Debug.WriteLine($"DEBUG: Descrizione: '{transazione.Descrizione}', Importo: {transazione.Importo}")
            Debug.WriteLine($"DEBUG: Modalità test: {isTestMode}")

            ' Nota: Il controllo delle API keys viene fatto all'inizio della classificazione batch
            ' Non è necessario ripeterlo qui per ogni transazione

            ' Step 0a: PRIORITÀ MASSIMA - Controllo aziende europee di trasporto
            Debug.WriteLine($"DEBUG: Step 0a - Controllo aziende europee di trasporto...")
            Dim riconoscimentoEuropeo = MoneyMind.Services.ClassificatoreTransazioni.RiconosciAziendaTrasporto(transazione.Descrizione)
            If riconoscimentoEuropeo.PunteggioConfidenza > 0 Then
                Debug.WriteLine($"DEBUG: 🎯 AZIENDA EUROPEA RICONOSCIUTA - MacroCategoria: '{riconoscimentoEuropeo.MacroCategoria}', Categoria: '{riconoscimentoEuropeo.Categoria}', ParolaChiave: '{riconoscimentoEuropeo.ParolaChiave}', Confidenza: {riconoscimentoEuropeo.PunteggioConfidenza:P1}")
                riconoscimentoEuropeo.Motivazione = "Azienda europea di trasporto riconosciuta (priorità massima)"
                Debug.WriteLine($"DEBUG: 🚀 RITORNANDO RICONOSCIMENTO EUROPEO - Motivazione impostata: '{riconoscimentoEuropeo.Motivazione}'")
                Return riconoscimentoEuropeo
            Else
                Debug.WriteLine($"DEBUG: ❌ NESSUNA AZIENDA EUROPEA RICONOSCIUTA")
            End If

            ' Step 0a-bis: CLASSIFICAZIONE SEMANTICA AVANZATA
            Debug.WriteLine($"DEBUG: Step 0a-bis - Controllo classificazione semantica avanzata...")
            Dim classificazioneSemantica = MoneyMind.Services.ClassificatoreTransazioni.ClassificazioneSemanticaAvanzata(transazione.Descrizione)
            If classificazioneSemantica.PunteggioConfidenza > 0 Then
                Debug.WriteLine($"DEBUG: 🧠 CLASSIFICAZIONE SEMANTICA RIUSCITA - MacroCategoria: '{classificazioneSemantica.MacroCategoria}', Categoria: '{classificazioneSemantica.Categoria}', ParolaChiave: '{classificazioneSemantica.ParolaChiave}', Confidenza: {classificazioneSemantica.PunteggioConfidenza:P1}")
                classificazioneSemantica.Motivazione = "Classificazione semantica avanzata (alta precisione)"
                Debug.WriteLine($"DEBUG: 🚀 RITORNANDO CLASSIFICAZIONE SEMANTICA - Motivazione impostata: '{classificazioneSemantica.Motivazione}'")
                Return classificazioneSemantica
            Else
                Debug.WriteLine($"DEBUG: ❌ NESSUNA CLASSIFICAZIONE SEMANTICA APPLICABILE")
            End If

            ' Step 0b: Controllo pre-classificazione per business specifici
            Debug.WriteLine($"DEBUG: Step 0b - Controllo pre-classificazione business specifici...")
            Dim risultatoPreClassifica = ClassifySpecificBusinessTypes(transazione.Descrizione)
            If risultatoPreClassifica IsNot Nothing Then
                Debug.WriteLine($"DEBUG: 🎯 PRE-CLASSIFICAZIONE RIUSCITA - MacroCategoria: '{risultatoPreClassifica.MacroCategoria}', Categoria: '{risultatoPreClassifica.Categoria}', ParolaChiave: '{risultatoPreClassifica.ParolaChiave}', Confidenza: {risultatoPreClassifica.PunteggioConfidenza:P1}")
                risultatoPreClassifica.Motivazione = "Classificato tramite riconoscimento business specifico"
                Debug.WriteLine($"DEBUG: 🚀 RITORNANDO PRE-CLASSIFICAZIONE - Motivazione impostata: '{risultatoPreClassifica.Motivazione}'")
                Return risultatoPreClassifica
            Else
                Debug.WriteLine($"DEBUG: ❌ PRE-CLASSIFICAZIONE FALLITA - Nessun business type riconosciuto")
            End If

            ' Step 1: Tentativo pattern tradizionale
            Debug.WriteLine($"DEBUG: Step 1 - Tentativo pattern tradizionale...")
            Dim risultatoPattern = _classificatore.ClassificaTransazione(transazione)
            Debug.WriteLine($"DEBUG: Pattern result - IsValid: {risultatoPattern.IsValid}, Confidenza: {risultatoPattern.PunteggioConfidenza}, MacroCategoria: '{risultatoPattern.MacroCategoria}', Categoria: '{risultatoPattern.Categoria}'")

            ' IMPORTANTE: Non considerare "Non Classificata" come un successo, anche se ha confidenza alta
            Dim isNonClassificata = (risultatoPattern.MacroCategoria = "Non Classificata" OrElse risultatoPattern.Categoria = "Non Classificata")
            Debug.WriteLine($"DEBUG: Pattern è 'Non Classificata': {isNonClassificata}")

            ' LOGICA AI ENHANCED SEMPLIFICATA:
            ' - In modalità test: sempre AI
            ' - In modalità normale: se il metodo è AI Enhanced, prioritizza sempre AI → Google → Pattern
            ' - Solo in casi di pattern molto sicuri (>= 0.8) e senza ambiguità, usa direttamente i pattern

            Dim requiresAIValidation As Boolean
            requiresAIValidation = Me.RequiresAIValidation(transazione.Descrizione, risultatoPattern)

            ' Controllo avanzato per ottimizzare l'uso dell'AI
            Dim shouldUseAI As Boolean = False

            ' Step 1: Pattern con alta confidenza - usa direttamente
            If risultatoPattern.PunteggioConfidenza >= 0.85 AndAlso Not isNonClassificata AndAlso Not isTestMode Then
                Debug.WriteLine($"DEBUG: ⚡ PATTERN ECCELLENTE (confidenza >= 85%) - saltando AI per performance")
                shouldUseAI = False
            ' Step 2: Pattern con buona confidenza ma richiede validazione
            ElseIf risultatoPattern.PunteggioConfidenza >= 0.7 AndAlso requiresAIValidation AndAlso Not isTestMode Then
                Debug.WriteLine($"DEBUG: ⚠️ PATTERN BUONO ma keyword ambigua - validazione AI necessaria")
                shouldUseAI = True
            ' Step 3: Pattern basso o non classificato - sempre AI
            ElseIf isNonClassificata OrElse risultatoPattern.PunteggioConfidenza < 0.7 OrElse isTestMode Then
                Debug.WriteLine($"DEBUG: 🤖 PATTERN DEBOLE/INESISTENTE - AI necessaria")
                shouldUseAI = True
            End If

            Debug.WriteLine($"DEBUG: 📊 DECISION MATRIX:")
            Debug.WriteLine($"DEBUG:   - Confidenza Pattern: {risultatoPattern.PunteggioConfidenza:P1}")
            Debug.WriteLine($"DEBUG:   - È Non Classificata: {isNonClassificata}")
            Debug.WriteLine($"DEBUG:   - Richiede Validazione AI: {requiresAIValidation}")
            Debug.WriteLine($"DEBUG:   - Modalità Test: {isTestMode}")
            Debug.WriteLine($"DEBUG:   - DECISIONE: Usa AI = {shouldUseAI}")

            If Not shouldUseAI Then
                Debug.WriteLine($"DEBUG: ⚡ OTTIMIZZAZIONE: Pattern sicuro utilizzato senza AI (risparmio di tempo e costi)")
                risultatoPattern.Motivazione = "Classificato tramite pattern esistente (alta confidenza - ottimizzato)"
                Return risultatoPattern
            End If

            Debug.WriteLine($"DEBUG: 🚀 PROCEDEANDO CON AI per classificazione ottimizzata...")

            ' Controllo ulteriore: verifica se questa descrizione è molto simile a qualcosa già classificato recentemente
            Dim descrizioneTruncata = TruncateString(transazione.Descrizione, 50)
            Debug.WriteLine($"DEBUG: 🔍 VERIFICA CACHE: Controllo similarità per '{descrizioneTruncata}'...")

            ' Step 2: Tentativo AI (GPT)
            Debug.WriteLine($"DEBUG: Step 2 - Tentativo AI (GPT)...")
            Dim openAIKey = MoneyMind.Services.GestoreApiKeys.CaricaChiaveApi("OpenAI")
            Debug.WriteLine($"DEBUG: OpenAI Key caricata: {If(String.IsNullOrEmpty(openAIKey), "VUOTA", $"Presente ({openAIKey.Length} caratteri, inizia con: '{openAIKey.Substring(0, Math.Min(10, openAIKey.Length))}...')")}")

            If Not String.IsNullOrEmpty(openAIKey) Then
                Try
                    Debug.WriteLine($"DEBUG: Chiamata a _gptClassificatore.AnalizzaTransazione...")
                    ' In modalità test, non includere esempi storici per evitare influenze da pattern esistenti
                    Dim suggerimentoAI = Await _gptClassificatore.AnalizzaTransazione(transazione.Descrizione, transazione.Importo, includiEsempiStorici:=Not isTestMode)
                    Debug.WriteLine($"DEBUG: Risposta AI ricevuta - IsNull: {suggerimentoAI Is Nothing}")

                    If suggerimentoAI IsNot Nothing Then
                        Debug.WriteLine($"DEBUG: AI Response Details:")
                        Debug.WriteLine($"DEBUG:   - IsValid: {suggerimentoAI.IsValid}")
                        Debug.WriteLine($"DEBUG:   - MacroCategoria: '{suggerimentoAI.MacroCategoria}'")
                        Debug.WriteLine($"DEBUG:   - Categoria: '{suggerimentoAI.Categoria}'")
                        Debug.WriteLine($"DEBUG:   - Motivazione: '{suggerimentoAI.Motivazione}'")
                        Debug.WriteLine($"DEBUG:   - ParolaChiave: '{suggerimentoAI.ParolaChiave}'")
                        Debug.WriteLine($"DEBUG:   - Confidenza: {suggerimentoAI.Confidenza}")

                        If suggerimentoAI.IsValid Then
                            Debug.WriteLine($"DEBUG: ✅ AI RIUSCITA - Validando accuratezza risultato...")

                            ' CONTROLLO AVANZATO ACCURATEZZA AI
                            If ValidateAIResult(suggerimentoAI, transazione.Descrizione) Then
                                Debug.WriteLine($"DEBUG: 🎯 RISULTATO AI VALIDATO - Alta qualità confermata")

                                Dim risultatoAI = ConvertSuggestionToResult(suggerimentoAI)
                                risultatoAI.Motivazione = "Classificato tramite AI (validato per accuratezza)"

                                ' Aumenta confidenza se la validazione è positiva
                                If risultatoAI.PunteggioConfidenza < 0.9 Then
                                    risultatoAI.PunteggioConfidenza = Math.Min(risultatoAI.PunteggioConfidenza + 0.1, 0.95)
                                    Debug.WriteLine($"DEBUG: 📈 BOOST CONFIDENZA: {risultatoAI.PunteggioConfidenza:P1} (per validazione positiva)")
                                End If

                                ' Auto-save pattern se abilitato E non in modalità test
                                If ChkAutoSavePatterns.IsChecked AndAlso Not isTestMode Then
                                    Debug.WriteLine($"DEBUG: Auto-save pattern abilitato (non in test mode), salvando...")
                                    Await SaveNewPattern(suggerimentoAI)
                                ElseIf isTestMode Then
                                    Debug.WriteLine($"DEBUG: Modalità test attiva - pattern NON salvato per evitare creazioni indesiderate")
                                End If

                                Debug.WriteLine($"DEBUG: 🚀 AI Enhanced COMPLETATO con successo (VALIDATO)")
                                Return risultatoAI
                            Else
                                Debug.WriteLine($"DEBUG: ⚠️ RISULTATO AI NON VALIDATO - Fallback a Google Places")
                                ' Continua con Google Places come fallback
                            End If
                        ElseIf Not suggerimentoAI.IsValid AndAlso
                               (suggerimentoAI.MacroCategoria = "Non Classificata" OrElse suggerimentoAI.Categoria = "Non Classificata") Then
                            ' L'AI ha fallito ma ha restituito "Non Classificata" - mostra questo risultato
                            Debug.WriteLine($"DEBUG: AI fallita, restituendo 'Non Classificata' per preview: {suggerimentoAI.Motivazione}")
                            Dim risultatoFallito = ConvertSuggestionToResult(suggerimentoAI)
                            risultatoFallito.Motivazione = $"AI non disponibile: {suggerimentoAI.Motivazione}"
                            Return risultatoFallito
                        ElseIf Not String.IsNullOrEmpty(suggerimentoAI.Motivazione) Then
                            Debug.WriteLine($"DEBUG: AI ha restituito errore specifico: {suggerimentoAI.Motivazione}")
                            ' L'AI ha restituito un errore specifico - mostralo all'utente
                            If suggerimentoAI.Motivazione.Contains("non valida") OrElse suggerimentoAI.Motivazione.Contains("scaduta") Then
                                Debug.WriteLine($"DEBUG: Mostrando dialog errore API key...")
                                Await ShowAPIKeyErrorDialog("OpenAI", suggerimentoAI.Motivazione)
                            End If
                        End If
                    Else
                        Debug.WriteLine($"DEBUG: AI ha restituito NULL - possibile errore nella chiamata")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"DEBUG: ECCEZIONE in AI classification: {ex.Message}")
                    Debug.WriteLine($"DEBUG: Stack trace: {ex.StackTrace}")
                    _logger?.LogWarning(ex, "AI classification failed, trying Google fallback")
                End Try
            Else
                Debug.WriteLine($"DEBUG: OpenAI Key VUOTA - saltando tentativo AI")
            End If

            ' Step 3: Tentativo Google Places
            Debug.WriteLine($"DEBUG: Step 3 - Tentativo Google Places...")
            Dim googleKey = MoneyMind.Services.GestoreApiKeys.CaricaChiaveApi("GooglePlaces")

            If Not String.IsNullOrEmpty(googleKey) AndAlso ShouldUseGooglePlaces(transazione.Descrizione) Then
                Try
                    Debug.WriteLine($"DEBUG: Chiamata a Google Places per business specifico...")
                    Dim suggerimentoGoogle = Await _gptClassificatore.RicercaWebPerDescrizioneAsync(transazione.Descrizione)

                    If suggerimentoGoogle IsNot Nothing AndAlso suggerimentoGoogle.IsValid Then
                        Debug.WriteLine($"DEBUG: GOOGLE PLACES RIUSCITO - MacroCategoria: '{suggerimentoGoogle.MacroCategoria}', Categoria: '{suggerimentoGoogle.Categoria}'")
                        Dim risultatoGoogle = ConvertSuggestionToResult(suggerimentoGoogle)
                        risultatoGoogle.Motivazione = "Classificato tramite Google Places"
                        Return risultatoGoogle
                    Else
                        Debug.WriteLine($"DEBUG: Google Places fallito o risultato non valido")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"DEBUG: ECCEZIONE in Google Places: {ex.Message}")
                    _logger?.LogWarning(ex, "Google Places classification failed")
                End Try
            Else
                Debug.WriteLine($"DEBUG: Google Places saltato - Key: {If(String.IsNullOrEmpty(googleKey), "VUOTA", "PRESENTE")}, ShouldUse: {ShouldUseGooglePlaces(transazione.Descrizione)}")
            End If

            ' Step 4: Fallback intelligente
            Debug.WriteLine($"DEBUG: Step 4 - Fallback intelligente...")

            ' Se avevamo un pattern decente (confidenza >= 0.5) ma non perfetto, usiamolo come fallback
            If risultatoPattern.IsValid AndAlso risultatoPattern.PunteggioConfidenza >= 0.5 AndAlso Not isNonClassificata Then
                Debug.WriteLine($"DEBUG: AI e Google falliti ma ho pattern decente (confidenza >= 0.5), lo uso come fallback")
                risultatoPattern.Motivazione = "Classificato tramite pattern (fallback da AI/Google)"
                Return risultatoPattern
            End If

            ' Altrimenti, classificazione offline
            Debug.WriteLine($"DEBUG: Nessun metodo utilizzabile, uso fallback offline...")
            Dim risultatoOffline = ExecuteOfflineClassification(transazione)
            Debug.WriteLine($"DEBUG: Offline result - MacroCategoria: '{risultatoOffline.MacroCategoria}', Categoria: '{risultatoOffline.Categoria}', Motivazione: '{risultatoOffline.Motivazione}'")
            Debug.WriteLine($"DEBUG: === FINE ExecuteAIEnhancedClassification ===")
            Return risultatoOffline

        Catch ex As Exception
            Debug.WriteLine($"DEBUG: ECCEZIONE GENERALE in ExecuteAIEnhancedClassification: {ex.Message}")
            Debug.WriteLine($"DEBUG: Stack trace: {ex.StackTrace}")
            _logger?.LogError(ex, "Error in AI Enhanced classification")
            Debug.WriteLine($"DEBUG: Fallback finale a offline classification...")
            Return ExecuteOfflineClassification(transazione)
        End Try
    End Function

    ''' <summary>
    ''' Classificazione offline intelligente
    ''' </summary>
    Private Function ExecuteOfflineClassification(transazione As Transazione) As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione
        Try
            Dim risultato As New MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione()

            ' Estrazione parola chiave intelligente
            Dim parolaChiave = EstrayParolaChiave(transazione.Descrizione)

            ' Classificazione di fallback basata su parole chiave
            Dim categoria = ClassifyByKeywords(parolaChiave, transazione.Descrizione)

            risultato.PatternUsato = parolaChiave
            risultato.MacroCategoria = categoria.MacroCategoria
            risultato.Categoria = categoria.Categoria
            risultato.Necessita = categoria.Necessita
            risultato.Frequenza = categoria.Frequenza
            risultato.Stagionalita = categoria.Stagionalita
            risultato.PunteggioConfidenza = 0.6
            risultato.LivelloConfidenza = "Media"
            risultato.ColoreConfidenza = "#FF9800"
            risultato.Motivazione = "Classificazione offline intelligente"
            risultato.IsValid = True

            Return risultato

        Catch ex As Exception
            _logger?.LogError(ex, "Error in offline classification")
            Return CreateMinimalFallback(transazione.Descrizione)
        End Try
    End Function

    ''' <summary>
    ''' Applica risultato classificazione alla transazione
    ''' </summary>
    Private Async Function ApplyClassificationResult(transazione As Transazione, risultato As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione) As Task
        Try
            ' Aggiorna oggetto transazione
            transazione.MacroCategoria = risultato.MacroCategoria
            transazione.Categoria = risultato.Categoria
            transazione.Necessita = risultato.Necessita
            transazione.Frequenza = risultato.Frequenza
            transazione.Stagionalita = risultato.Stagionalita

            ' Salva nel database
            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                Dim updateQuery = "UPDATE Transazioni SET 
                                  MacroCategoria = @macro, Categoria = @cat, 
                                  Necessita = @nec, Frequenza = @freq, 
                                  Stagionalita = @stag 
                                  WHERE ID = @id"

                Using cmd As New SQLiteCommand(updateQuery, connection)
                    cmd.Parameters.AddWithValue("@macro", risultato.MacroCategoria)
                    cmd.Parameters.AddWithValue("@cat", risultato.Categoria)
                    cmd.Parameters.AddWithValue("@nec", risultato.Necessita)
                    cmd.Parameters.AddWithValue("@freq", risultato.Frequenza)
                    cmd.Parameters.AddWithValue("@stag", risultato.Stagionalita)
                    cmd.Parameters.AddWithValue("@id", transazione.ID)

                    Await Task.Run(Sub() cmd.ExecuteNonQuery())
                End Using
            End Using

        Catch ex As Exception
            _logger?.LogError(ex, "Error applying classification result")
        End Try
    End Function


#End Region

#Region "=== TEST SINGOLA TRANSAZIONE ==="

    ''' <summary>
    ''' Apre il pannello per il test di una singola transazione
    ''' </summary>
    Private Async Function OpenTestTransactionPanel(metodo As String) As Task
        Try
            Debug.WriteLine($"DEBUG: === OpenTestTransactionPanel START ===")
            Debug.WriteLine($"DEBUG: Metodo per test transazione: '{metodo}'")

            ' Nascondi altri pannelli
            ProgressPanel.Visibility = Visibility.Collapsed
            ResultsPanel.Visibility = Visibility.Collapsed
            PreviewPanel.Visibility = Visibility.Collapsed
            SingleClassificationPanel.Visibility = Visibility.Collapsed

            ' Mostra pannello test
            SingleTestPanel.Visibility = Visibility.Visible
            SingleTestResult.Visibility = Visibility.Collapsed

            ' Memorizza il metodo selezionato per uso nel test
            _selectedMethodForTest = metodo
            Debug.WriteLine($"DEBUG: Metodo memorizzato per test: '{_selectedMethodForTest}'")

            ' Scroll automatico verso il pannello
            Await Task.Delay(200)
            ScrollToPanel(SingleTestPanel)
            Debug.WriteLine($"DEBUG: OpenTestTransactionPanel completato")

        Catch ex As Exception
            _logger?.LogError(ex, "Error opening test transaction panel")
            ShowError("Errore", "Errore durante l'apertura del pannello test: " & ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Chiude pannello test singola transazione
    ''' </summary>
    Private Sub BtnCloseSingleTest_Click(sender As Object, e As RoutedEventArgs)
        SingleTestPanel.Visibility = Visibility.Collapsed
    End Sub

    ''' <summary>
    ''' Esegue test su singola transazione
    ''' </summary>
    Private Async Sub BtnRunSingleTest_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim descrizione = TxtTestDescription.Text.Trim()
            Dim importoText = TxtTestAmount.Text.Trim().Replace(".", ",")

            If String.IsNullOrEmpty(descrizione) Then
                ShowWarning("Inserisci una descrizione per il test")
                Return
            End If

            Dim importo As Decimal = 0
            If Not Decimal.TryParse(importoText, importo) Then
                importo = 0
            End If

            ' Crea transazione test
            Dim transazioneTest As New Transazione With {
                .ID = -1,
                .Descrizione = descrizione,
                .Importo = importo,
                .Data = DateTime.Today
            }

            ' Esegui classificazione
            Dim startTime = DateTime.Now
            Dim metodo = _selectedMethodForTest
            Dim risultato As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione = Nothing
            Dim suggerimentoAI As MoneyMind.Services.GptClassificatoreTransazioni.SuggerimentoClassificazione = Nothing

            BtnRunSingleTest.IsEnabled = False
            BtnRunSingleTest.Content = "🔄 Elaborazione..."

            Select Case metodo
                Case "PatternOnly"
                    risultato = _classificatore.ClassificaTransazione(transazioneTest)
                Case "AIEnhanced"
                    ' Per AIEnhanced, ottieni anche il suggerimento AI originale per le informazioni aggiuntive
                    If _gptClassificatore IsNot Nothing Then
                        suggerimentoAI = Await _gptClassificatore.AnalizzaTransazione(transazioneTest.Descrizione, transazioneTest.Importo, includiEsempiStorici:=False)
                    End If
                    risultato = Await ExecuteAIEnhancedClassification(transazioneTest, isTestMode:=True)
                Case "OfflineOnly"
                    risultato = ExecuteOfflineClassification(transazioneTest)
                Case "ManualOnly"
                    ' Per test singola, non possiamo aprire pannello manuale, usiamo offline invece
                    risultato = ExecuteOfflineClassification(transazioneTest)
                    risultato.Motivazione = "Test con modalità manuale (usando fallback offline)"
            End Select

            Dim elapsed = DateTime.Now - startTime

            ' Mostra risultati
            If risultato IsNot Nothing Then
                TxtTestPattern.Text = risultato.PatternUsato

                ' Mostra la parola chiave estratta per debugging
                Dim parolaChiave = ""
                If Not String.IsNullOrEmpty(transazioneTest.Descrizione) Then
                    ' Usa la stessa funzione che abbiamo migliorato nel GptClassificatore
                    If _gptClassificatore IsNot Nothing Then
                        parolaChiave = _gptClassificatore.EstraiParolaChiaveDaDescrizione(transazioneTest.Descrizione)
                    End If
                End If
                TxtTestParolaChiave.Text = If(String.IsNullOrEmpty(parolaChiave), "-", parolaChiave)

                TxtTestMacro.Text = risultato.MacroCategoria
                TxtTestCategory.Text = risultato.Categoria
                TxtTestConfidence.Text = $"{risultato.PunteggioConfidenza:P1}"
                TxtTestConfidenceLevel.Text = risultato.LivelloConfidenza.ToUpper()
                ConfidenceBadge.Background = New SolidColorBrush(ColorConverter.ConvertFromString(risultato.ColoreConfidenza))

                ' Mostra il metodo usato in modo più chiaro
                TxtTestMethod.Text = GetMethodDisplayDescription(metodo)
                TxtTestTime.Text = $"{elapsed.TotalMilliseconds:F0}ms"

                ' Mostra informazioni AI aggiuntive se disponibili e se l'AI ha riconosciuto una società
                If suggerimentoAI IsNot Nothing AndAlso suggerimentoAI.IsValid Then
                    ' Mostra sempre la motivazione per i risultati AI
                    TxtTestMotivazione.Text = If(String.IsNullOrEmpty(suggerimentoAI.Motivazione), "-", suggerimentoAI.Motivazione)

                    ' Mostra sommario e descrizione attività SOLO se l'AI ha riconosciuto una società
                    If Not String.IsNullOrEmpty(suggerimentoAI.NomeSocieta) AndAlso
                       Not String.IsNullOrEmpty(suggerimentoAI.SommarioSocieta) Then

                        LblTestSummary.Visibility = Visibility.Visible
                        TxtTestSummary.Visibility = Visibility.Visible
                        TxtTestSummary.Text = suggerimentoAI.SommarioSocieta

                        LblTestActivity.Visibility = Visibility.Visible
                        TxtTestActivity.Visibility = Visibility.Visible
                        TxtTestActivity.Text = If(String.IsNullOrEmpty(suggerimentoAI.DescrizioneAttivita), "-", suggerimentoAI.DescrizioneAttivita)
                    Else
                        ' Nascondi i campi società se non riconosciuta
                        LblTestSummary.Visibility = Visibility.Collapsed
                        TxtTestSummary.Visibility = Visibility.Collapsed
                        LblTestActivity.Visibility = Visibility.Collapsed
                        TxtTestActivity.Visibility = Visibility.Collapsed
                    End If
                Else
                    ' Mostra solo la motivazione standard per metodi non-AI
                    TxtTestMotivazione.Text = risultato.Motivazione
                    ' Nascondi i campi società
                    LblTestSummary.Visibility = Visibility.Collapsed
                    TxtTestSummary.Visibility = Visibility.Collapsed
                    LblTestActivity.Visibility = Visibility.Collapsed
                    TxtTestActivity.Visibility = Visibility.Collapsed
                End If

                ' Controlla se ci sono errori di configurazione Google Places da mostrare
                If risultato.Motivazione.Contains("[AVVISO: Google Places") Then
                    ' Mostra dialog di avviso per errori Google Places
                    ShowWarning("⚠️ Problema Configurazione Google Places" & vbCrLf & vbCrLf &
                              "Google Places API non è configurata correttamente." & vbCrLf &
                              "Vai in Impostazioni → API Keys per risolvere il problema." & vbCrLf & vbCrLf &
                              "La classificazione continua con metodi alternativi.")
                End If

                SingleTestResult.Visibility = Visibility.Visible
            Else
                ShowWarning("Impossibile classificare la transazione con i metodi selezionati")
            End If

        Catch ex As Exception
            ShowError("Errore durante il test", ex.Message)
        Finally
            BtnRunSingleTest.IsEnabled = True
            BtnRunSingleTest.Content = "🚀 Testa"
        End Try
    End Sub

    ''' <summary>
    ''' Apre il pannello per classificazione singola transazione dal database
    ''' </summary>
    Private Async Function OpenSingleClassificationPanel(metodo As String) As Task
        Try
            Debug.WriteLine($"DEBUG: === OpenSingleClassificationPanel START ===")
            Debug.WriteLine($"DEBUG: Metodo per classificazione singola: '{metodo}'")

            ' Nascondi altri pannelli
            ProgressPanel.Visibility = Visibility.Collapsed
            ResultsPanel.Visibility = Visibility.Collapsed
            PreviewPanel.Visibility = Visibility.Collapsed
            SingleTestPanel.Visibility = Visibility.Collapsed

            ' Mostra pannello classificazione singola
            SingleClassificationPanel.Visibility = Visibility.Visible
            SingleClassificationResult.Visibility = Visibility.Collapsed

            ' Memorizza il metodo selezionato per uso successivo
            _selectedMethodForSingle = metodo
            Debug.WriteLine($"DEBUG: Metodo memorizzato per uso futuro: '{_selectedMethodForSingle}'")

            ' Aggiorna info metodo nel pannello
            TxtSingleMethodInfo.Text = $"Metodo: {GetMethodDisplayDescription(metodo)}"

            ' Carica transazioni non classificate
            Debug.WriteLine($"DEBUG: Caricamento transazioni per pannello singolo...")
            Await LoadUnclassifiedTransactionsForSinglePanel()
            Debug.WriteLine($"DEBUG: OpenSingleClassificationPanel completato")

            ' Scroll automatico verso il pannello
            Await Task.Delay(200)
            ScrollToPanel(SingleClassificationPanel)

        Catch ex As Exception
            _logger?.LogError(ex, "Error opening single classification panel")
            ShowError("Errore", "Errore durante l'apertura del pannello classificazione singola: " & ex.Message)
        End Try
    End Function

    Private _selectedMethodForSingle As String = "AIEnhanced"
    Private _selectedMethodForTest As String = "AIEnhanced"
    Private _singleTransactionsList As New List(Of ManualTransactionView)
    Private _selectedSingleTransaction As ManualTransactionView
    Private _currentClassificationResult As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione

    ''' <summary>
    ''' Carica transazioni non classificate per il pannello singola classificazione
    ''' </summary>
    Private Async Function LoadUnclassifiedTransactionsForSinglePanel() As Task
        Try
            _singleTransactionsList.Clear()

            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                ' Usa la stessa query del metodo funzionante ma con LIMIT
                Dim query = "SELECT * FROM Transazioni
                           WHERE (MacroCategoria IS NULL OR MacroCategoria = '' OR MacroCategoria = 'Non Classificata')
                           ORDER BY Data DESC LIMIT 50"

                Using cmd As New SQLiteCommand(query, connection)
                    Using reader = Await Task.Run(Function() cmd.ExecuteReader())
                        While reader.Read()
                            ' Usa la stessa logica del metodo LoadUnmatchedTransactions che funziona
                            Dim transazione As New Transazione With {
                                .ID = Convert.ToInt32(reader("ID")),
                                .Descrizione = reader("Descrizione").ToString(),
                                .Importo = Convert.ToDecimal(reader("Importo")),
                                .Data = Convert.ToDateTime(reader("Data")),
                                .MacroCategoria = If(reader("MacroCategoria") Is DBNull.Value, String.Empty, reader("MacroCategoria").ToString()),
                                .Categoria = If(reader("Categoria") Is DBNull.Value, String.Empty, reader("Categoria").ToString()),
                                .Necessita = String.Empty,
                                .Frequenza = String.Empty,
                                .Stagionalita = String.Empty
                            }

                            ' Estrai parola chiave per la visualizzazione
                            Dim parolaChiave = If(_gptClassificatore IsNot Nothing,
                                                  _gptClassificatore.EstraiParolaChiaveDaDescrizione(transazione.Descrizione),
                                                  "")
                            _singleTransactionsList.Add(New ManualTransactionView(transazione, parolaChiave))
                        End While
                    End Using
                End Using
            End Using

            ' Aggiorna DataGrid
            DgSingleTransactions.ItemsSource = _singleTransactionsList
            TxtSingleTransactionCount.Text = $"({_singleTransactionsList.Count} trovate)"

            ' Abilita/disabilita controlli
            BtnClassifySelectedSingle.IsEnabled = False

        Catch ex As Exception
            _logger?.LogError(ex, "Error loading unclassified transactions for single panel")
            ShowError("Errore", "Errore durante il caricamento delle transazioni: " & ex.Message)
        End Try
    End Function


    ''' <summary>
    ''' Aggiorna badge confidenza con colore appropriato
    ''' </summary>
    Private Sub UpdateConfidenceBadge(badge As Border, textBlock As TextBlock, confidenza As Double)
        If confidenza >= 0.8 Then
            badge.Background = New SolidColorBrush(Color.FromRgb(76, 175, 80)) ' Verde
            textBlock.Text = "ALTA"
        ElseIf confidenza >= 0.6 Then
            badge.Background = New SolidColorBrush(Color.FromRgb(255, 193, 7)) ' Giallo
            textBlock.Text = "MEDIA"
        Else
            badge.Background = New SolidColorBrush(Color.FromRgb(244, 67, 54)) ' Rosso
            textBlock.Text = "BASSA"
        End If
    End Sub

    ''' <summary>
    ''' Gestisce click su "Salva come Pattern"
    ''' </summary>
    Private Async Sub BtnSaveSinglePattern_Click(sender As Object, e As RoutedEventArgs)
        Try
            If _currentClassificationResult Is Nothing OrElse _selectedSingleTransaction Is Nothing Then
                ShowError("Errore", "Nessun risultato di classificazione disponibile.")
                Return
            End If

            Dim result = MessageBox.Show(
                $"Vuoi salvare questo risultato come nuovo pattern?" & vbCrLf & vbCrLf &
                $"Transazione: {_selectedSingleTransaction.Descrizione}" & vbCrLf &
                $"MacroCategoria: {_currentClassificationResult.MacroCategoria}" & vbCrLf &
                $"Categoria: {_currentClassificationResult.Categoria}" & vbCrLf &
                $"Confidenza: {_currentClassificationResult.PunteggioConfidenza:P1}",
                "Salva Pattern",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question)

            If result = MessageBoxResult.Yes Then
                ' Salva pattern utilizzando il metodo esistente
                Await SalvaPatternDaRisultato(_selectedSingleTransaction.OriginalTransaction, _currentClassificationResult)
                ShowInfo("Pattern salvato con successo!")

                ' Chiudi il pannello risultati
                SingleClassificationResult.Visibility = Visibility.Collapsed
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error saving single pattern")
            ShowError("Errore", "Errore durante il salvataggio del pattern: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Gestisce click su "Chiudi" nel pannello risultati singola
    ''' </summary>
    Private Sub BtnCloseSingleResult_Click(sender As Object, e As RoutedEventArgs)
        SingleClassificationResult.Visibility = Visibility.Collapsed
    End Sub

#End Region

#Region "=== HELPER METHODS ==="

    ''' <summary>
    ''' Carica tutte le transazioni dal database (anche quelle già classificate)
    ''' </summary>
    Private Async Function LoadAllTransactions() As Task
        Try
            _transazioni.Clear()

            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                Dim query = "SELECT * FROM Transazioni ORDER BY Data DESC"

                Using cmd As New SQLiteCommand(query, connection)
                    Using reader = Await Task.Run(Function() cmd.ExecuteReader())
                        While reader.Read()
                            Dim transazione As New Transazione With {
                                .ID = Convert.ToInt32(reader("ID")),
                                .Descrizione = reader("Descrizione").ToString(),
                                .Importo = Convert.ToDecimal(reader("Importo")),
                                .Data = Convert.ToDateTime(reader("Data")),
                                .MacroCategoria = If(IsDBNull(reader("MacroCategoria")), String.Empty, reader("MacroCategoria").ToString()),
                                .Categoria = If(IsDBNull(reader("Categoria")), String.Empty, reader("Categoria").ToString()),
                                .Necessita = If(IsDBNull(reader("Necessita")), String.Empty, reader("Necessita").ToString()),
                                .Frequenza = If(IsDBNull(reader("Frequenza")), String.Empty, reader("Frequenza").ToString()),
                                .Stagionalita = If(IsDBNull(reader("Stagionalita")), String.Empty, reader("Stagionalita").ToString())
                            }
                            _transazioni.Add(transazione)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            _logger?.LogError(ex, "Error loading all transactions")
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Carica solo le transazioni non ancora classificate
    ''' </summary>
    Private Async Function LoadUnmatchedTransactions() As Task
        Try
            _transazioni.Clear()

            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                Dim query = "SELECT * FROM Transazioni 
                           WHERE (MacroCategoria IS NULL OR MacroCategoria = '' OR MacroCategoria = 'Non Classificata') 
                           ORDER BY Data DESC"

                Using cmd As New SQLiteCommand(query, connection)
                    Using reader = Await Task.Run(Function() cmd.ExecuteReader())
                        While reader.Read()
                            Dim transazione As New Transazione With {
                                .ID = Convert.ToInt32(reader("ID")),
                                .Descrizione = reader("Descrizione").ToString(),
                                .Importo = Convert.ToDecimal(reader("Importo")),
                                .Data = Convert.ToDateTime(reader("Data")),
                                .MacroCategoria = If(IsDBNull(reader("MacroCategoria")), String.Empty, reader("MacroCategoria").ToString()),
                                .Categoria = If(IsDBNull(reader("Categoria")), String.Empty, reader("Categoria").ToString()),
                                .Necessita = If(IsDBNull(reader("Necessita")), String.Empty, reader("Necessita").ToString()),
                                .Frequenza = If(IsDBNull(reader("Frequenza")), String.Empty, reader("Frequenza").ToString()),
                                .Stagionalita = If(IsDBNull(reader("Stagionalita")), String.Empty, reader("Stagionalita").ToString())
                            }
                            _transazioni.Add(transazione)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            _logger?.LogError(ex, "Error loading unmatched transactions")
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Reset tutte le classificazioni esistenti nel database
    ''' </summary>
    Private Async Function ResetAllClassifications() As Task
        Try
            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                Dim updateQuery = "UPDATE Transazioni SET 
                                 MacroCategoria = NULL, 
                                 Categoria = NULL, 
                                 Necessita = NULL, 
                                 Frequenza = NULL, 
                                 Stagionalita = NULL"

                Using cmd As New SQLiteCommand(updateQuery, connection)
                    Await Task.Run(Sub() cmd.ExecuteNonQuery())
                End Using
            End Using

            ' Aggiorna anche gli oggetti in memoria
            For Each transazione In _transazioni
                transazione.MacroCategoria = String.Empty
                transazione.Categoria = String.Empty
                transazione.Necessita = String.Empty
                transazione.Frequenza = String.Empty
                transazione.Stagionalita = String.Empty
            Next

        Catch ex As Exception
            _logger?.LogError(ex, "Error resetting classifications")
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Carica transazioni per classificazione manuale
    ''' </summary>
    Private Async Function LoadManualTransactions() As Task
        Try
            ' CRITICAL FIX: Usa Dispatcher per operazioni thread-safe su ObservableCollection
            Await Dispatcher.InvokeAsync(Sub()
                ' Disconnetti temporaneamente l'ItemsSource per evitare race condition
                ManualTransactionsList.ItemsSource = Nothing
                _manualTransactions.Clear()
            End Sub)

            ' Carica dati dal database (operazione in background)
            Dim tempTransactions As New List(Of ManualTransactionView)

            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                Await connection.OpenAsync()

                Dim query = "SELECT * FROM Transazioni
                           WHERE (MacroCategoria IS NULL OR MacroCategoria = '' OR MacroCategoria = 'Non Classificata')
                           ORDER BY Data DESC"

                Using cmd As New SQLiteCommand(query, connection)
                    Using reader = Await cmd.ExecuteReaderAsync()
                        While Await reader.ReadAsync()
                            Dim transazione As New Transazione With {
                                .ID = Convert.ToInt32(reader("ID")),
                                .Descrizione = reader("Descrizione").ToString(),
                                .Importo = Convert.ToDecimal(reader("Importo")),
                                .Data = Convert.ToDateTime(reader("Data")),
                                .MacroCategoria = String.Empty,
                                .Categoria = String.Empty,
                                .Necessita = String.Empty,
                                .Frequenza = String.Empty,
                                .Stagionalita = String.Empty
                            }
                            ' Estrai parola chiave per la visualizzazione
                            Dim parolaChiave = If(_gptClassificatore IsNot Nothing,
                                                  _gptClassificatore.EstraiParolaChiaveDaDescrizione(transazione.Descrizione),
                                                  "")
                            tempTransactions.Add(New ManualTransactionView(transazione, parolaChiave))
                        End While
                    End Using
                End Using
            End Using

            ' CRITICAL FIX: Aggiorna la collezione sul thread UI
            Await Dispatcher.InvokeAsync(Sub()
                ' Popola la ObservableCollection
                For Each item In tempTransactions
                    _manualTransactions.Add(item)
                Next

                ' Ricollega l'ItemsSource SOLO DOPO aver popolato i dati
                ManualTransactionsList.ItemsSource = _manualTransactions
                TxtManualCount.Text = $"{_manualTransactions.Count} da classificare"
            End Sub)

        Catch ex As Exception
            _logger?.LogError(ex, "Error loading manual transactions")
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Popola le ComboBox per la classificazione manuale
    ''' </summary>
    Private Sub PopulateManualComboBoxes()
        Try
            ' Popola MacroCategoria
            Dim macroCategorie = GetDistinctMacroCategories()
            CmbMacroCategoria.ItemsSource = macroCategorie

            ' Popola Categoria (tutte inizialmente)
            Dim categorie = GetDistinctCategories()
            CmbCategoria.ItemsSource = categorie

            ' Aggiungi event handler per cascade
            RemoveHandler CmbMacroCategoria.SelectionChanged, AddressOf CmbMacroCategoria_SelectionChanged
            AddHandler CmbMacroCategoria.SelectionChanged, AddressOf CmbMacroCategoria_SelectionChanged

        Catch ex As Exception
            _logger?.LogError(ex, "Error populating manual comboboxes")
        End Try
    End Sub

    ''' <summary>
    ''' Gestisce cascade MacroCategoria -> Categoria
    ''' </summary>
    Private Sub CmbMacroCategoria_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Try
            Dim combo = TryCast(sender, ComboBox)
            If combo IsNot Nothing AndAlso combo.SelectedItem IsNot Nothing Then
                Dim selectedMacro = combo.SelectedItem.ToString()

                ' Aggiorna Categoria ComboBox con solo categorie della MacroCategoria selezionata
                Dim categorieFiltered = GetDistinctCategories(selectedMacro)
                CmbCategoria.ItemsSource = categorieFiltered
                CmbCategoria.SelectedIndex = -1 ' Reset selezione

                _logger?.LogInformation($"Filtered categories for {selectedMacro}: {categorieFiltered.Count} items")
            End If
        Catch ex As Exception
            _logger?.LogError(ex, "Error in MacroCategoria cascade")
        End Try
    End Sub

    ''' <summary>
    ''' Verifica se MacroCategoria e Categoria esistono nel database
    ''' </summary>
    Private Async Function CheckCategoriesExist(macroCategoria As String, categoria As String) As Task(Of (macroExists As Boolean, categoriaExists As Boolean))
        Try
            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                ' Controlla MacroCategoria
                Dim macroQuery = "SELECT COUNT(*) FROM Pattern WHERE MacroCategoria = @macro"
                Using macroCmd As New SQLiteCommand(macroQuery, connection)
                    macroCmd.Parameters.AddWithValue("@macro", macroCategoria)
                    Dim macroExists = Convert.ToInt32(Await Task.Run(Function() macroCmd.ExecuteScalar())) > 0

                    ' Controlla Categoria
                    Dim catQuery = "SELECT COUNT(*) FROM Pattern WHERE Categoria = @cat"
                    Using catCmd As New SQLiteCommand(catQuery, connection)
                        catCmd.Parameters.AddWithValue("@cat", categoria)
                        Dim catExists = Convert.ToInt32(Await Task.Run(Function() catCmd.ExecuteScalar())) > 0

                        Return (macroExists, catExists)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            _logger?.LogError(ex, "Error checking categories existence")
            Return (False, False)
        End Try
    End Function

    ''' <summary>
    ''' Crea un nuovo pattern dal form manuale
    ''' </summary>
    Private Async Function CreateManualPattern() As Task
        Try
            Dim macroCategoria = CmbMacroCategoria.Text.Trim()
            Dim categoria = CmbCategoria.Text.Trim()

            ' Controlla se le categorie esistono già
            Dim categoriesCheck = Await CheckCategoriesExist(macroCategoria, categoria)
            Dim macroExists = categoriesCheck.macroExists
            Dim categoriaExists = categoriesCheck.categoriaExists

            ' Costruisci messaggio di notifica se necessario
            Dim notificationMessage = ""
            If Not macroExists AndAlso Not categoriaExists Then
                notificationMessage = $"✨ Nuova MacroCategoria '{macroCategoria}' e Categoria '{categoria}' create!"
            ElseIf Not macroExists Then
                notificationMessage = $"✨ Nuova MacroCategoria '{macroCategoria}' creata!"
            ElseIf Not categoriaExists Then
                notificationMessage = $"✨ Nuova Categoria '{categoria}' creata nella MacroCategoria esistente!"
            End If

            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                ' Verifica se il pattern esiste già
                Dim checkQuery = "SELECT COUNT(*) FROM Pattern WHERE Parola = @parola"
                Using checkCmd As New SQLiteCommand(checkQuery, connection)
                    checkCmd.Parameters.AddWithValue("@parola", TxtKeyword.Text.Trim())
                    Dim exists = Convert.ToInt32(Await Task.Run(Function() checkCmd.ExecuteScalar())) > 0

                    If exists Then
                        ' Aggiorna pattern esistente
                        Dim updateQuery = "UPDATE Pattern SET 
                                         MacroCategoria = @macro, Categoria = @cat,
                                         Necessita = @nec, Frequenza = @freq,
                                         Stagionalita = @stag, Peso = Peso + 1,
                                         Fonte = @fonte
                                         WHERE Parola = @parola"

                        Using updateCmd As New SQLiteCommand(updateQuery, connection)
                            updateCmd.Parameters.AddWithValue("@parola", TxtKeyword.Text.Trim())
                            updateCmd.Parameters.AddWithValue("@macro", macroCategoria)
                            updateCmd.Parameters.AddWithValue("@cat", categoria)
                            updateCmd.Parameters.AddWithValue("@nec", GetComboBoxText(CmbNecessita))
                            updateCmd.Parameters.AddWithValue("@freq", GetComboBoxText(CmbFrequenza))
                            updateCmd.Parameters.AddWithValue("@stag", GetComboBoxText(CmbStagionalita))
                            updateCmd.Parameters.AddWithValue("@fonte", "Manuale")

                            Await Task.Run(Sub() updateCmd.ExecuteNonQuery())
                        End Using
                    Else
                        ' Inserisci nuovo pattern
                        Dim insertQuery = "INSERT INTO Pattern
                                         (Parola, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita, Peso, Fonte)
                                         VALUES (@parola, @macro, @cat, @nec, @freq, @stag, 5, @fonte)"

                        Using insertCmd As New SQLiteCommand(insertQuery, connection)
                            insertCmd.Parameters.AddWithValue("@parola", TxtKeyword.Text.Trim())
                            insertCmd.Parameters.AddWithValue("@macro", macroCategoria)
                            insertCmd.Parameters.AddWithValue("@cat", categoria)
                            insertCmd.Parameters.AddWithValue("@nec", GetComboBoxText(CmbNecessita))
                            insertCmd.Parameters.AddWithValue("@freq", GetComboBoxText(CmbFrequenza))
                            insertCmd.Parameters.AddWithValue("@stag", GetComboBoxText(CmbStagionalita))
                            insertCmd.Parameters.AddWithValue("@fonte", "Manuale")

                            Await Task.Run(Sub() insertCmd.ExecuteNonQuery())
                        End Using
                    End If
                End Using
            End Using

            ' Mostra notifica se necessario
            If Not String.IsNullOrEmpty(notificationMessage) Then
                ShowInfo(notificationMessage & $" Pattern '{TxtKeyword.Text.Trim()}' salvato con successo.")
            End If

            ' Aggiorna ComboBoxes per riflettere le nuove categorie
            Await RefreshManualComboBoxes()

            ' Notifica altri componenti della creazione di nuovi pattern
            Services.PatternNotificationService.Instance.NotifyPatternsChanged()

        Catch ex As Exception
            _logger?.LogError(ex, "Error creating manual pattern")
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Aggiorna le ComboBox dopo la creazione di un pattern
    ''' </summary>
    Private Async Function RefreshManualComboBoxes() As Task
        Try
            Await Task.Run(Sub()
                               Dispatcher.Invoke(Sub()
                                                     ' Salva selezioni correnti
                                                     Dim currentMacro = CmbMacroCategoria.Text
                                                     Dim currentCat = CmbCategoria.Text

                                                     ' Ricarica dati
                                                     PopulateManualComboBoxes()

                                                     ' Ripristina selezioni se ancora valide
                                                     If Not String.IsNullOrEmpty(currentMacro) Then
                                                         CmbMacroCategoria.Text = currentMacro
                                                     End If
                                                     If Not String.IsNullOrEmpty(currentCat) Then
                                                         CmbCategoria.Text = currentCat
                                                     End If
                                                 End Sub)
                           End Sub)
        Catch ex As Exception
            _logger?.LogError(ex, "Error refreshing manual comboboxes")
        End Try
    End Function

    ''' <summary>
    ''' Applica classificazione manuale alla transazione selezionata
    ''' </summary>
    Private Async Function ApplyManualClassification() As Task
        Try
            If _selectedManualTransaction Is Nothing Then Return

            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                Dim updateQuery = "UPDATE Transazioni SET 
                                 MacroCategoria = @macro, Categoria = @cat,
                                 Necessita = @nec, Frequenza = @freq,
                                 Stagionalita = @stag
                                 WHERE ID = @id"

                Using cmd As New SQLiteCommand(updateQuery, connection)
                    cmd.Parameters.AddWithValue("@macro", CmbMacroCategoria.Text.Trim())
                    cmd.Parameters.AddWithValue("@cat", CmbCategoria.Text.Trim())
                    cmd.Parameters.AddWithValue("@nec", GetComboBoxText(CmbNecessita))
                    cmd.Parameters.AddWithValue("@freq", GetComboBoxText(CmbFrequenza))
                    cmd.Parameters.AddWithValue("@stag", GetComboBoxText(CmbStagionalita))
                    cmd.Parameters.AddWithValue("@id", _selectedManualTransaction.ID)

                    Await Task.Run(Sub() cmd.ExecuteNonQuery())
                End Using
            End Using

        Catch ex As Exception
            _logger?.LogError(ex, "Error applying manual classification")
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Ottiene testo da ComboBox (item o text)
    ''' </summary>
    Private Function GetComboBoxText(combo As ComboBox) As String
        If combo.SelectedItem IsNot Nothing Then
            Dim item = TryCast(combo.SelectedItem, ComboBoxItem)
            If item IsNot Nothing AndAlso item.Content IsNot Nothing Then
                Return item.Content.ToString()
            Else
                Return If(combo.Text, String.Empty)
            End If
        Else
            Return If(combo.Text, String.Empty)
        End If
    End Function

    ''' <summary>
    ''' Ottiene metodo di classificazione selezionato
    ''' </summary>
    Private Function GetSelectedClassificationMethod() As String
        If RadioPatternOnly.IsChecked Then Return "PatternOnly"
        If RadioAIEnhanced.IsChecked Then Return "AIEnhanced"
        If RadioOfflineOnly.IsChecked Then Return "OfflineOnly"
        If RadioManualOnly.IsChecked Then Return "ManualOnly"
        Return "AIEnhanced" ' Default
    End Function

    ''' <summary>
    ''' Ottiene scope di classificazione selezionato
    ''' </summary>
    Private Function GetSelectedClassificationScope() As String
        Debug.WriteLine($"DEBUG: GetSelectedClassificationScope - Controllo RadioButton:")
        Debug.WriteLine($"DEBUG: RadioScopeAll.IsChecked = {RadioScopeAll?.IsChecked}")
        Debug.WriteLine($"DEBUG: RadioScopeUnmatched.IsChecked = {RadioScopeUnmatched?.IsChecked}")
        Debug.WriteLine($"DEBUG: RadioScopeSingle.IsChecked = {RadioScopeSingle?.IsChecked}")
        Debug.WriteLine($"DEBUG: RadioScopeTest.IsChecked = {RadioScopeTest?.IsChecked}")

        If RadioScopeAll?.IsChecked = True Then
            Debug.WriteLine($"DEBUG: Scope selezionato: 'All'")
            Return "All"
        End If
        If RadioScopeUnmatched?.IsChecked = True Then
            Debug.WriteLine($"DEBUG: Scope selezionato: 'Unmatched'")
            Return "Unmatched"
        End If
        If RadioScopeSingle?.IsChecked = True Then
            Debug.WriteLine($"DEBUG: Scope selezionato: 'Single'")
            Return "Single"
        End If
        If RadioScopeTest?.IsChecked = True Then
            Debug.WriteLine($"DEBUG: Scope selezionato: 'Test'")
            Return "Test"
        End If

        Debug.WriteLine($"DEBUG: Nessun scope selezionato, usando default 'Unmatched'")
        Return "Unmatched" ' Default
    End Function

    ''' <summary>
    ''' Ottiene descrizione user-friendly del metodo selezionato
    ''' </summary>
    Private Function GetMethodDisplayDescription() As String
        Return GetMethodDisplayDescription(GetSelectedClassificationMethod())
    End Function

    ''' <summary>
    ''' Ottiene descrizione user-friendly per un metodo specifico
    ''' </summary>
    Private Function GetMethodDisplayDescription(metodo As String) As String
        Select Case metodo
            Case "PatternOnly"
                Return "Solo Pattern Esistenti"
            Case "AIEnhanced"
                Return "AI Potenziata"
            Case "OfflineOnly"
                Return "Solo Modalità Offline"
            Case "ManualOnly"
                Return "Classificazione Manuale"
            Case Else
                Return "Metodo sconosciuto"
        End Select
    End Function

    ''' <summary>
    ''' Ottiene descrizione user-friendly dello scope selezionato
    ''' </summary>
    Private Function GetScopeDisplayDescription() As String
        Select Case GetSelectedClassificationScope()
            Case "All"
                Return "Tutte le Transazioni - rimuove classificazioni esistenti e riapplica il metodo a tutte"
            Case "Unmatched"
                Return "Solo Non Classificate - classifica solo le transazioni senza MacroCategoria o con 'Non Classificata'"
            Case "Single"
                Return "Singola Transazione - apre il pannello per selezionare e classificare una singola transazione dal database"
            Case "Test"
                Return "Test Singola Transazione - testa una descrizione personalizzata senza creare pattern o modificare il database"
            Case Else
                Return "Scope sconosciuto"
        End Select
    End Function

    ''' <summary>
    ''' Converte suggerimento AI in risultato classificazione
    ''' </summary>
    Private Function ConvertSuggestionToResult(suggestion As MoneyMind.Services.GptClassificatoreTransazioni.SuggerimentoClassificazione) As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione
        Debug.WriteLine($"DEBUG: === ConvertSuggestionToResult ===")
        Debug.WriteLine($"DEBUG: Input Suggestion:")
        Debug.WriteLine($"DEBUG:   - MacroCategoria: '{suggestion.MacroCategoria}'")
        Debug.WriteLine($"DEBUG:   - Categoria: '{suggestion.Categoria}'")
        Debug.WriteLine($"DEBUG:   - Necessita: '{suggestion.Necessita}'")
        Debug.WriteLine($"DEBUG:   - Frequenza: '{suggestion.Frequenza}'")
        Debug.WriteLine($"DEBUG:   - Stagionalita: '{suggestion.Stagionalita}'")
        Debug.WriteLine($"DEBUG:   - ParolaChiave: '{suggestion.ParolaChiave}'")
        Debug.WriteLine($"DEBUG:   - Confidenza: {suggestion.Confidenza}")
        Debug.WriteLine($"DEBUG:   - IsValid: {suggestion.IsValid}")
        Debug.WriteLine($"DEBUG:   - Motivazione: '{suggestion.Motivazione}'")

        Dim result = New MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione With {
            .MacroCategoria = suggestion.MacroCategoria,
            .Categoria = suggestion.Categoria,
            .Necessita = suggestion.Necessita,
            .Frequenza = suggestion.Frequenza,
            .Stagionalita = suggestion.Stagionalita,
            .PatternUsato = suggestion.ParolaChiave,
            .ParolaChiave = suggestion.ParolaChiave,
            .PunteggioConfidenza = suggestion.Confidenza / 100.0,
            .LivelloConfidenza = If(suggestion.Confidenza >= 80, "Alta", If(suggestion.Confidenza >= 60, "Media", "Bassa")),
            .ColoreConfidenza = If(suggestion.Confidenza >= 80, "#4CAF50", If(suggestion.Confidenza >= 60, "#FF9800", "#F44336")),
            .Motivazione = suggestion.Motivazione,
            .IsValid = suggestion.IsValid
        }

        Debug.WriteLine($"DEBUG: Output Result:")
        Debug.WriteLine($"DEBUG:   - MacroCategoria: '{result.MacroCategoria}'")
        Debug.WriteLine($"DEBUG:   - Categoria: '{result.Categoria}'")
        Debug.WriteLine($"DEBUG:   - PunteggioConfidenza: {result.PunteggioConfidenza}")
        Debug.WriteLine($"DEBUG:   - IsValid: {result.IsValid}")
        Debug.WriteLine($"DEBUG:   - Motivazione: '{result.Motivazione}'")
        Debug.WriteLine($"DEBUG: === Fine ConvertSuggestionToResult ===")

        Return result
    End Function

    ''' <summary>
    ''' Estrae parola chiave da descrizione
    ''' </summary>
    Private Function EstrayParolaChiave(descrizione As String) As String
        ' Implementazione semplificata - usa logica del vecchio sistema
        Return MoneyMind.Services.ClassificatoreTransazioni.NormalizzaTesto(descrizione).Split(" "c).FirstOrDefault()
    End Function

    ''' <summary>
    ''' Classificazione basata su parole chiave
    ''' </summary>
    Private Function ClassifyByKeywords(parolaChiave As String, descrizione As String) As (MacroCategoria As String, Categoria As String, Necessita As String, Frequenza As String, Stagionalita As String)
        ' Implementazione base con regole euristiche
        Dim desc = descrizione.ToUpperInvariant()

        ' Commercio
        If desc.Contains("ESSELUNGA") OrElse desc.Contains("CONAD") OrElse desc.Contains("SUPERMERCATO") Then
            Return ("Alimentari", "Spesa quotidiana", "Essenziale", "Ricorrente", "Annuale")
        End If

        ' Carburante
        If desc.Contains("BENZINA") OrElse desc.Contains("CARBURANTE") OrElse desc.Contains("SHELL") OrElse desc.Contains("ENI") Then
            Return ("Trasporti", "Carburante", "Essenziale", "Ricorrente", "Annuale")
        End If

        ' Ristorazione
        If desc.Contains("RISTORANTE") OrElse desc.Contains("PIZZERIA") OrElse desc.Contains("BAR") Then
            Return ("Ristorazione", "Ristoranti", "Non essenziale", "Occasionale", "Annuale")
        End If

        ' Farmacia
        If desc.Contains("FARMACIA") OrElse desc.Contains("PARAFARMACIA") Then
            Return ("Salute", "Medicinali", "Essenziale", "Occasionale", "Annuale")
        End If

        ' Default - Non Classificata
        Return ("Non Classificata", "Non Classificata", "Base", "Occasionale", "Annuale")
    End Function

    ''' <summary>
    ''' Crea fallback minimale garantito
    ''' </summary>
    Private Function CreateMinimalFallback(descrizione As String) As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione
        Return New MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione With {
            .MacroCategoria = "Non Classificata",
            .Categoria = "Non Classificata",
            .Necessita = "Base",
            .Frequenza = "Occasionale",
            .Stagionalita = "Annuale",
            .PatternUsato = EstrayParolaChiave(descrizione),
            .ParolaChiave = EstrayParolaChiave(descrizione),
            .PunteggioConfidenza = 0.3,
            .LivelloConfidenza = "Bassa",
            .ColoreConfidenza = "#F44336",
            .Motivazione = "Classificazione di emergenza",
            .IsValid = True
        }
    End Function

    ''' <summary>
    ''' Salva nuovo pattern scoperto
    ''' </summary>
    Private Async Function SaveNewPattern(suggestion As MoneyMind.Services.GptClassificatoreTransazioni.SuggerimentoClassificazione) As Task
        Try
            Debug.WriteLine($"DEBUG: === SaveNewPattern START ===")
            Debug.WriteLine($"DEBUG: ParolaChiave: '{suggestion.ParolaChiave}', MacroCategoria: '{suggestion.MacroCategoria}', Categoria: '{suggestion.Categoria}'")

            ' Non salvare pattern per classificazioni "Non Classificata"
            If suggestion.MacroCategoria = "Non Classificata" OrElse suggestion.Categoria = "Non Classificata" Then
                Debug.WriteLine($"DEBUG: Pattern 'Non Classificata' non salvato")
                Return
            End If
            
            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                ' Verifica esistenza
                Debug.WriteLine($"DEBUG: Verifica esistenza pattern per '{suggestion.ParolaChiave}'...")
                Dim checkQuery = "SELECT COUNT(*) FROM Pattern WHERE Parola = @parola"
                Using checkCmd As New SQLiteCommand(checkQuery, connection)
                    checkCmd.Parameters.AddWithValue("@parola", suggestion.ParolaChiave)
                    Dim exists = Convert.ToInt32(Await Task.Run(Function() checkCmd.ExecuteScalar())) > 0
                    Debug.WriteLine($"DEBUG: Pattern esistente: {exists}")

                    If Not exists Then
                        Debug.WriteLine($"DEBUG: Inserimento nuovo pattern nel database...")
                        ' Inserimento nuovo pattern
                        Dim insertQuery = "INSERT INTO Pattern 
                                         (Parola, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita, Peso, Fonte) 
                                         VALUES (@parola, @macro, @cat, @nec, @freq, @stag, @peso, @fonte)"

                        Using insertCmd As New SQLiteCommand(insertQuery, connection)
                            insertCmd.Parameters.AddWithValue("@parola", suggestion.ParolaChiave)
                            insertCmd.Parameters.AddWithValue("@macro", suggestion.MacroCategoria)
                            insertCmd.Parameters.AddWithValue("@cat", suggestion.Categoria)
                            insertCmd.Parameters.AddWithValue("@nec", suggestion.Necessita)
                            insertCmd.Parameters.AddWithValue("@freq", suggestion.Frequenza)
                            insertCmd.Parameters.AddWithValue("@stag", suggestion.Stagionalita)
                            insertCmd.Parameters.AddWithValue("@peso", suggestion.Peso)
                            insertCmd.Parameters.AddWithValue("@fonte", "AI")

                            Await Task.Run(Sub() insertCmd.ExecuteNonQuery())
                            Debug.WriteLine($"DEBUG: Pattern inserito con successo nel database")

                            ' Notifica la creazione di nuovi pattern
                            Debug.WriteLine($"DEBUG: Notifica aggiornamento pattern...")
                            Services.PatternNotificationService.Instance.NotifyPatternsChanged()
                            Debug.WriteLine($"DEBUG: Notifica pattern inviata")
                        End Using
                    Else
                        Debug.WriteLine($"DEBUG: Pattern già esistente, aggiornamento saltato")
                    End If
                End Using
            End Using
            Debug.WriteLine($"DEBUG: === SaveNewPattern COMPLETATO ===")

        Catch ex As Exception
            _logger?.LogError(ex, "Error saving new pattern")
        End Try
    End Function

    ''' <summary>
    ''' Mostra risultati classificazione
    ''' </summary>
    Private Sub ShowClassificationResults(processed As Integer, successful As Integer, elapsed As TimeSpan)
        TxtProcessed.Text = processed.ToString()
        TxtSuccessful.Text = successful.ToString()

        Dim successRate = If(processed > 0, (successful / processed) * 100, 0)
        TxtResultSummary.Text = $"Tasso di successo: {successRate:F1}% • Tempo: {elapsed.TotalSeconds:F1}s • Media: {If(processed > 0, elapsed.TotalMilliseconds / processed, 0):F0}ms/transazione"

        ResultsPanel.Visibility = Visibility.Visible
    End Sub

    ''' <summary>
    ''' Tronca stringa per UI
    ''' </summary>
    Private Function TruncateString(text As String, maxLength As Integer) As String
        If String.IsNullOrEmpty(text) OrElse text.Length <= maxLength Then
            Return text
        End If
        Return text.Substring(0, maxLength - 3) & "..."
    End Function

    ''' <summary>
    ''' Ottiene nome metodo per display
    ''' </summary>
    Private Function GetMethodDisplayName(motivazione As String) As String
        If String.IsNullOrEmpty(motivazione) Then Return "Sconosciuto"

        ' Controlli più specifici per identificare il metodo corretto
        If motivazione.Contains("intelligenza artificiale") OrElse motivazione.Contains("AI") Then Return "🤖 AI"
        If motivazione.Contains("Google Places") Then Return "🌐 Google"
        If motivazione.Contains("pattern esistente") Then Return "🎯 Pattern"
        If motivazione.Contains("offline") OrElse motivazione.Contains("Offline") Then Return "🔧 Offline"
        If motivazione.Contains("manuale") OrElse motivazione.Contains("Manuale") Then Return "✋ Manuale"
        If motivazione.Contains("fallback") Then Return "🔄 Fallback"

        Return "❓ Sistema"
    End Function

    ''' <summary>
    ''' Ottiene il nome di visualizzazione basato sul metodo selezionato e la motivazione
    ''' </summary>
    Private Function GetDisplayNameFromMethod(metodo As String, motivazione As String) As String
        ' Per AI Enhanced, mostra cosa ha effettivamente funzionato
        If metodo = "AIEnhanced" Then
            If String.IsNullOrEmpty(motivazione) Then
                Return "🔧 Sistema"
            ElseIf motivazione.Contains("riconoscimento business specifico") Then
                Return "🏢 Business"
            ElseIf motivazione.Contains("intelligenza artificiale") OrElse motivazione.Contains("AI") Then
                Return "🤖 AI"
            ElseIf motivazione.Contains("pattern") Then
                Return "🎯 Pattern"
            ElseIf motivazione.Contains("Google Places") Then
                Return "🌐 Google"
            Else
                Return "🔧 Sistema"
            End If
        ElseIf metodo = "PatternOnly" Then
            Return "🎯 Pattern Only"
        ElseIf metodo = "ManualOnly" Then
            Return "✋ Manuale"
        Else
            ' Fallback al metodo originale per compatibilità
            Return GetMethodDisplayName(motivazione)
        End If
    End Function

    ''' <summary>
    ''' Valida chiave API OpenAI (formato più flessibile)
    ''' </summary>
    Private Function IsValidOpenAIKey(apiKey As String) As Boolean
        If String.IsNullOrWhiteSpace(apiKey) Then Return False

        ' Rimuovi spazi
        apiKey = apiKey.Trim()

        ' Verifica lunghezza minima (le chiavi OpenAI sono generalmente lunghe)
        If apiKey.Length < 20 Then Return False

        ' Verifica che contenga solo caratteri alfanumerici, trattini e underscore
        If Not System.Text.RegularExpressions.Regex.IsMatch(apiKey, "^[a-zA-Z0-9_-]+$") Then Return False

        ' Accetta chiavi che iniziano con sk- (formato classico) o altri formati OpenAI
        Return apiKey.StartsWith("sk-") OrElse
               apiKey.StartsWith("org-") OrElse
               (apiKey.Length >= 40 AndAlso apiKey.All(Function(c) Char.IsLetterOrDigit(c) OrElse c = "-"c OrElse c = "_"c))
    End Function

    ''' <summary>
    ''' Mostra risultati anteprima classificazione
    ''' </summary>
    Private Sub ShowPreviewResults()
        Try
            ' Configura ComboBox columns con opzioni disponibili
            SetupPreviewComboBoxes()

            ' Registra event handler per aggiornamento contatore
            For Each item In _previewResults
                AddHandler item.SelectionChanged, AddressOf OnItemSelectionChanged
            Next

            ' Configura DataGrid
            PreviewDataGrid.ItemsSource = _previewResults

            ' Nascondi pannelli precedenti
            ProgressPanel.Visibility = Visibility.Collapsed
            ResultsPanel.Visibility = Visibility.Collapsed

            ' Mostra pannello anteprima
            PreviewPanel.Visibility = Visibility.Visible

            ' Auto-scroll all'anteprima per attirare l'attenzione dell'utente
            ScrollToPreviewPanel()

            ' Aggiorna contatori nell'header
            TxtPreviewCount.Text = $"{_previewResults.Count} transazioni"

            ' Aggiorna contatore selezioni
            UpdateSelectionCount()

        Catch ex As Exception
            _logger?.LogError(ex, "Error showing preview results")
            ShowError("Errore Anteprima", "Errore durante la visualizzazione dell'anteprima: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Configura ComboBox options per anteprima
    ''' </summary>
    Private Sub SetupPreviewComboBoxes()
        Try
            ' Ottieni liste distinte da database
            Dim macroCategorie = GetDistinctMacroCategories()
            Dim categorie = GetDistinctCategories()

            ' Trova ComboBox columns
            For Each column In PreviewDataGrid.Columns
                If TypeOf column Is DataGridComboBoxColumn Then
                    Dim comboColumn = CType(column, DataGridComboBoxColumn)

                    If comboColumn.Header.ToString() = "MacroCategoria" Then
                        comboColumn.ItemsSource = macroCategorie
                    ElseIf comboColumn.Header.ToString() = "Categoria" Then
                        comboColumn.ItemsSource = categorie
                    End If
                End If
            Next

        Catch ex As Exception
            _logger?.LogError(ex, "Error setting up preview comboboxes")
        End Try
    End Sub

    ''' <summary>
    ''' Ottieni macro categorie distinte
    ''' </summary>
    Private Function GetDistinctMacroCategories() As List(Of String)
        Try
            ' Usa il nuovo servizio con cache trasparente
            Return MoneyMind.Services.CacheService.GetOrSet("macrocategories_smart_center", 
                Function() MoneyMind.Services.PatternService.GetDistinctMacroCategories(), 5)
        Catch ex As Exception
            MoneyMind.Services.LoggingService.LogError("GetDistinctMacroCategories", ex)
            _logger?.LogError(ex, "Error getting macro categories")
            
            ' Fallback al metodo originale in caso di errore
            Return GetDistinctMacroCategoriesLegacy()
        End Try
    End Function
    
    ''' <summary>
    ''' Fallback method - mantiene logica originale per sicurezza
    ''' </summary>
    Private Function GetDistinctMacroCategoriesLegacy() As List(Of String)
        Dim result As New List(Of String)()

        Try
            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()
                ' Solo da Pattern per consistenza con PatternManagerWindow
                Dim query = "SELECT DISTINCT MacroCategoria FROM Pattern 
                           WHERE MacroCategoria IS NOT NULL AND MacroCategoria != '' 
                           ORDER BY MacroCategoria COLLATE NOCASE"

                Using cmd As New SQLiteCommand(query, connection)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim categoria = reader("MacroCategoria").ToString()
                            If Not String.IsNullOrEmpty(categoria) Then
                                result.Add(categoria)
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MoneyMind.Services.LoggingService.LogError("GetDistinctMacroCategoriesLegacy", ex)
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Ottieni categorie distinte
    ''' </summary>
    Private Function GetDistinctCategories(Optional macroCategoria As String = Nothing) As List(Of String)
        Try
            ' Usa i nuovi servizi con cache
            If String.IsNullOrEmpty(macroCategoria) Then
                Return MoneyMind.Services.CacheService.GetOrSet("categories_smart_center", 
                    Function() MoneyMind.Services.PatternService.GetDistinctCategories(), 5)
            Else
                ' Cache specifica per MacroCategoria
                Dim cacheKey = $"categories_for_{macroCategoria.ToLowerInvariant()}_smart_center"
                Return MoneyMind.Services.CacheService.GetOrSet(cacheKey, 
                    Function() MoneyMind.Services.PatternService.GetCategoriesForMacroAsync(macroCategoria).Result, 10)
            End If
        Catch ex As Exception
            MoneyMind.Services.LoggingService.LogError("GetDistinctCategories", ex, $"MacroCategoria: {macroCategoria}")
            _logger?.LogError(ex, "Error getting categories")
            
            ' Fallback al metodo originale
            Return GetDistinctCategoriesLegacy(macroCategoria)
        End Try
    End Function
    
    ''' <summary>
    ''' Fallback method - mantiene logica originale per sicurezza
    ''' </summary>
    Private Function GetDistinctCategoriesLegacy(Optional macroCategoria As String = Nothing) As List(Of String)
        Dim result As New List(Of String)()

        Try
            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()
                Dim query As String

                If String.IsNullOrEmpty(macroCategoria) Then
                    ' Tutte le categorie se nessuna MacroCategoria selezionata
                    query = "SELECT DISTINCT Categoria FROM Pattern 
                           WHERE Categoria IS NOT NULL AND Categoria != '' 
                           ORDER BY Categoria COLLATE NOCASE"
                Else
                    ' Solo categorie della MacroCategoria selezionata
                    query = "SELECT DISTINCT Categoria FROM Pattern 
                           WHERE MacroCategoria = @macro AND Categoria IS NOT NULL AND Categoria != '' 
                           ORDER BY Categoria COLLATE NOCASE"
                End If

                Using cmd As New SQLiteCommand(query, connection)
                    If Not String.IsNullOrEmpty(macroCategoria) Then
                        cmd.Parameters.AddWithValue("@macro", macroCategoria)
                    End If

                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim categoria = reader("Categoria").ToString()
                            If Not String.IsNullOrEmpty(categoria) Then
                                result.Add(categoria)
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MoneyMind.Services.LoggingService.LogError("GetDistinctCategoriesLegacy", ex, $"MacroCategoria: {macroCategoria}")
        End Try

        Return result
    End Function

#End Region

#Region "=== UI EVENTS ==="

    ''' <summary>
    ''' Aggiorna statistiche dashboard
    ''' </summary>
    Private Async Sub BtnRefreshStats_Click(sender As Object, e As RoutedEventArgs)
        Await UpdateDashboardStats()
    End Sub

    ''' <summary>
    ''' Gestione pattern - Apre PatternManagerWindow
    ''' </summary>
    Private Sub BtnViewPatterns_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim patternWindow As New PatternManagerWindow() With {
                .Owner = Me
            }
            patternWindow.Show()
        Catch ex As Exception
            ShowError("Errore apertura gestione pattern", ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Nasconde wizard
    ''' </summary>
    Private Sub BtnHideWizard_Click(sender As Object, e As RoutedEventArgs)
        WizardPanel.Visibility = Visibility.Collapsed
        BtnToggleWizard.Visibility = Visibility.Visible
    End Sub

    ''' <summary>
    ''' Toggle wizard visibility
    ''' </summary>
    Private Sub BtnToggleWizard_Click(sender As Object, e As RoutedEventArgs)
        If WizardPanel.Visibility = Visibility.Visible Then
            WizardPanel.Visibility = Visibility.Collapsed
            BtnToggleWizard.Content = "🧙‍♂️ Mostra Consigli"
        Else
            WizardPanel.Visibility = Visibility.Visible
            BtnToggleWizard.Content = "🧙‍♂️ Nascondi Consigli"
        End If
    End Sub

    ''' <summary>
    ''' Approva tutte le classificazioni nell'anteprima
    ''' </summary>
    Private Async Sub BtnApproveAll_Click(sender As Object, e As RoutedEventArgs)
        Try
            BtnApproveAll.IsEnabled = False

            ' Esegui reset classificazioni esistenti se necessario (All scope)
            If _requiresResetBeforeApproval Then
                LblPreviewStatus.Text = "Rimozione classificazioni esistenti..."
                Debug.WriteLine($"DEBUG: Esecuzione reset classificazioni prima dell'approvazione")
                Await ResetAllClassifications()
                _requiresResetBeforeApproval = False ' Reset flag
            End If

            LblPreviewStatus.Text = "Applicando classificazioni..."

            Dim successful As Integer = 0
            Dim startTime = DateTime.Now

            For Each preview In _previewResults.Where(Function(p) p.IsSelected)
                Try
                    Await ApplyPreviewResult(preview)
                    successful += 1
                Catch ex As Exception
                    _logger?.LogError(ex, $"Error applying classification for transaction {preview.ID}")
                End Try

                ' Update progress
                LblPreviewStatus.Text = $"Applicando... ({successful}/{_previewResults.Where(Function(p) p.IsSelected).Count()})"
                Await Task.Delay(10)
            Next

            Dim elapsed = DateTime.Now - startTime

            ' Nascondi anteprima e mostra risultati
            PreviewPanel.Visibility = Visibility.Collapsed
            ShowClassificationResults(_previewResults.Count, successful, elapsed)

            ' Aggiorna dashboard
            Await UpdateDashboardStats()

            ' Notifica MainWindow per aggiornare la lista transazioni
            Debug.WriteLine($"DEBUG: Sollevando evento TransactionsClassified per aggiornare MainWindow...")
            RaiseEvent TransactionsClassified()

        Catch ex As Exception
            _logger?.LogError(ex, "Error in approve all")
            ShowError("Errore Approvazione", "Errore durante l'applicazione delle classificazioni: " & ex.Message)
        Finally
            BtnApproveAll.IsEnabled = True
        End Try
    End Sub

    ''' <summary>
    ''' Approva solo classificazioni selezionate
    ''' </summary>
    Private Async Sub BtnApproveSelected_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim selectedItems = _previewResults.Where(Function(p) p.IsSelected).ToList()
            If selectedItems.Count = 0 Then
                ShowInfo("Seleziona almeno una transazione da approvare.")
                Return
            End If

            BtnApproveSelected.IsEnabled = False

            ' Esegui reset classificazioni esistenti se necessario (All scope)
            If _requiresResetBeforeApproval Then
                LblPreviewStatus.Text = "Rimozione classificazioni esistenti..."
                Debug.WriteLine($"DEBUG: Esecuzione reset classificazioni prima dell'approvazione selezionata")
                Await ResetAllClassifications()
                _requiresResetBeforeApproval = False ' Reset flag
            End If

            LblPreviewStatus.Text = "Applicando classificazioni selezionate..."

            Dim successful As Integer = 0
            Dim startTime = DateTime.Now

            For Each preview In selectedItems
                Try
                    Await ApplyPreviewResult(preview)
                    successful += 1
                Catch ex As Exception
                    _logger?.LogError(ex, $"Error applying classification for transaction {preview.ID}")
                End Try

                ' Update progress
                LblPreviewStatus.Text = $"Applicando... ({successful}/{selectedItems.Count})"
                Await Task.Delay(10)
            Next

            Dim elapsed = DateTime.Now - startTime

            ' Nascondi anteprima e mostra risultati
            PreviewPanel.Visibility = Visibility.Collapsed
            ShowClassificationResults(selectedItems.Count, successful, elapsed)

            ' Aggiorna dashboard
            Await UpdateDashboardStats()

            ' Notifica MainWindow per aggiornare la lista transazioni
            Debug.WriteLine($"DEBUG: Sollevando evento TransactionsClassified per aggiornare MainWindow...")
            RaiseEvent TransactionsClassified()

        Catch ex As Exception
            _logger?.LogError(ex, "Error in approve selected")
            ShowError("Errore Approvazione", "Errore durante l'applicazione delle classificazioni selezionate: " & ex.Message)
        Finally
            BtnApproveSelected.IsEnabled = True
        End Try
    End Sub

    ''' <summary>
    ''' Annulla anteprima
    ''' </summary>
    Private Sub BtnCancelPreview_Click(sender As Object, e As RoutedEventArgs)
        PreviewPanel.Visibility = Visibility.Collapsed
        _previewResults.Clear()

        ' Reset flag se necessario (così le classificazioni esistenti rimangono intatte)
        If _requiresResetBeforeApproval Then
            _requiresResetBeforeApproval = False
            Debug.WriteLine($"DEBUG: Anteprima annullata - classificazioni esistenti preservate")
        End If

        LblProgressStatus.Text = "Anteprima annullata."
    End Sub

    ''' <summary>
    ''' Chiudi anteprima
    ''' </summary>
    Private Sub BtnClosePreview_Click(sender As Object, e As RoutedEventArgs)
        PreviewPanel.Visibility = Visibility.Collapsed
    End Sub

    ''' <summary>
    ''' Applica risultato classificazione da anteprima
    ''' </summary>
    Private Async Function ApplyPreviewResult(preview As TransactionPreview) As Task
        Try
            ' Trova transazione originale
            Dim transazione = preview.OriginalTransaction
            If transazione Is Nothing Then
                Debug.WriteLine($"DEBUG: ApplyPreviewResult - Transazione originale non trovata per ID {preview.ID}")
                Return
            End If

            Debug.WriteLine($"DEBUG: ApplyPreviewResult - Applicazione classificazione per ID {preview.ID}: MacroCategoria='{preview.MacroCategoria}', Categoria='{preview.Categoria}'")

            ' Aggiorna oggetto transazione
            transazione.MacroCategoria = preview.MacroCategoria
            transazione.Categoria = preview.Categoria

            ' Salva nel database
            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                connection.Open()

                Dim updateQuery = "UPDATE Transazioni SET
                                  MacroCategoria = @macro, Categoria = @cat
                                  WHERE ID = @id"

                Using cmd As New SQLiteCommand(updateQuery, connection)
                    cmd.Parameters.AddWithValue("@macro", preview.MacroCategoria)
                    cmd.Parameters.AddWithValue("@cat", preview.Categoria)
                    cmd.Parameters.AddWithValue("@id", preview.ID)

                    Dim rowsAffected = Await Task.Run(Function() cmd.ExecuteNonQuery())
                    Debug.WriteLine($"DEBUG: ApplyPreviewResult - Update completato per ID {preview.ID}, righe interessate: {rowsAffected}")
                End Using
            End Using

        Catch ex As Exception
            _logger?.LogError(ex, "Error applying preview result")
            Throw
        End Try
    End Function

    ''' <summary>
    ''' Prossimo consiglio
    ''' </summary>
    Private Sub BtnNextTip_Click(sender As Object, e As RoutedEventArgs)
        _currentTipIndex = (_currentTipIndex + 1) Mod _smartTips.Count
        ShowCurrentTip()
    End Sub

    ''' <summary>
    ''' Mostra consiglio corrente
    ''' </summary>
    Private Sub ShowCurrentTip()
        If _smartTips IsNot Nothing AndAlso _smartTips.Count > 0 Then
            TxtCurrentTip.Text = _smartTips(_currentTipIndex)
        End If
    End Sub

    ''' <summary>
    ''' Mostra wizard se necessario
    ''' </summary>
    Private Sub ShowWizardIfNeeded()
        ' Mostra wizard se poche transazioni o prima volta
        If _transazioni.Count < 10 Then
            WizardPanel.Visibility = Visibility.Visible
        End If
    End Sub

    ''' <summary>
    ''' Minimizza finestra
    ''' </summary>
    Private Sub BtnMinimize_Click(sender As Object, e As RoutedEventArgs)
        Me.WindowState = WindowState.Minimized
    End Sub

    ''' <summary>
    ''' Chiude finestra
    ''' </summary>
    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    ''' <summary>
    ''' Drag della finestra
    ''' </summary>
    Private Sub Window_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        Try
            If e.LeftButton = MouseButtonState.Pressed Then
                Me.DragMove()
            End If
        Catch
            ' Ignora errori di drag
        End Try
    End Sub

    ''' <summary>
    ''' Selezione transazione nella lista manuale
    ''' </summary>
    Private Sub ManualTransactionsList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Try
            Dim selectedItem = TryCast(ManualTransactionsList.SelectedItem, ManualTransactionView)
            If selectedItem IsNot Nothing Then
                _selectedManualTransaction = selectedItem

                ' Mostra dettagli transazione
                SelectedTransactionPanel.Visibility = Visibility.Visible
                SelectedTransactionDesc.Text = selectedItem.Descrizione
                SelectedTransactionAmount.Text = $"{selectedItem.ImportoFormatted} - {selectedItem.DataFormatted}"

                ' Abilita form e popola parola chiave
                PatternCreationForm.IsEnabled = True
                TxtKeyword.Text = EstrayParolaChiave(selectedItem.Descrizione)

                ' Reset form
                CmbMacroCategoria.SelectedItem = Nothing
                CmbCategoria.SelectedItem = Nothing
                CmbNecessita.SelectedIndex = 0
                CmbFrequenza.SelectedIndex = 0
                CmbStagionalita.SelectedIndex = 0
            Else
                SelectedTransactionPanel.Visibility = Visibility.Collapsed
                PatternCreationForm.IsEnabled = False
                _selectedManualTransaction = Nothing
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error in manual transaction selection")
        End Try
    End Sub

    ''' <summary>
    ''' Aggiorna lista transazioni manuali
    ''' </summary>
    Private Async Sub BtnRefreshManualList_Click(sender As Object, e As RoutedEventArgs)
        Try
            Await LoadManualTransactions()
            ShowInfo($"Lista aggiornata: {_manualTransactions.Count} transazioni da classificare")
        Catch ex As Exception
            ShowError("Errore", "Errore durante l'aggiornamento della lista: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Chiude classificazione manuale
    ''' </summary>
    Private Sub BtnCloseManual_Click(sender As Object, e As RoutedEventArgs)
        ' CRITICAL FIX: Disconnetti ItemsSource prima di pulire la collezione
        ManualTransactionsList.ItemsSource = Nothing
        ManualClassificationPanel.Visibility = Visibility.Collapsed
        _manualTransactions.Clear()
        _selectedManualTransaction = Nothing
    End Sub

    ''' <summary>
    ''' Crea pattern e applica classificazione
    ''' </summary>
    Private Async Sub BtnCreateAndApply_Click(sender As Object, e As RoutedEventArgs)
        Try
            If _selectedManualTransaction Is Nothing Then
                ShowWarning("Seleziona prima una transazione!")
                Return
            End If

            ' Valida input
            If String.IsNullOrWhiteSpace(TxtKeyword.Text) Then
                ShowWarning("Inserisci una parola chiave!")
                TxtKeyword.Focus()
                Return
            End If

            If String.IsNullOrWhiteSpace(CmbMacroCategoria.Text) Then
                ShowWarning("Inserisci una MacroCategoria!")
                CmbMacroCategoria.Focus()
                Return
            End If

            If String.IsNullOrWhiteSpace(CmbCategoria.Text) Then
                ShowWarning("Inserisci una Categoria!")
                CmbCategoria.Focus()
                Return
            End If

            BtnCreateAndApply.IsEnabled = False
            BtnCreateAndApply.Content = "🔄 Creazione..."

            ' Crea pattern nel database
            Await CreateManualPattern()

            ' Applica classificazione alla transazione
            Await ApplyManualClassification()

            ' CRITICAL FIX: Rimuovi dalla ObservableCollection in modo thread-safe
            Await Dispatcher.InvokeAsync(Sub()
                _manualTransactions.Remove(_selectedManualTransaction)
                TxtManualCount.Text = $"{_manualTransactions.Count} da classificare"
            End Sub)

            ' Reset selection
            _selectedManualTransaction = Nothing
            SelectedTransactionPanel.Visibility = Visibility.Collapsed
            PatternCreationForm.IsEnabled = False

            ' Aggiorna dashboard
            Await UpdateDashboardStats()

            ' Notifica MainWindow per aggiornare la lista transazioni
            Debug.WriteLine($"DEBUG: Sollevando evento TransactionsClassified per aggiornare MainWindow...")
            RaiseEvent TransactionsClassified()

            ShowSuccess($"Pattern creato e classificazione applicata! Rimangono {_manualTransactions.Count} transazioni.")

        Catch ex As Exception
            _logger?.LogError(ex, "Error creating pattern and applying classification")
            ShowError("Errore", "Errore durante la creazione del pattern: " & ex.Message)
        Finally
            BtnCreateAndApply.IsEnabled = True
            BtnCreateAndApply.Content = "✅ Crea Pattern e Applica"
        End Try
    End Sub
    
#End Region
    
#Region "=== MESSAGGI UTENTE ==="
    
    Private Sub ShowSuccess(message As String)
        MessageBox.Show(message, "✅ Successo", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub
    
    Private Sub ShowInfo(message As String)
        MessageBox.Show(message, "ℹ️ Informazione", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub
    
    Private Sub ShowWarning(message As String)
        MessageBox.Show(message, "⚠️ Attenzione", MessageBoxButton.OK, MessageBoxImage.Warning)
    End Sub
    
    Private Sub ShowError(title As String, message As String)
        MessageBox.Show(message, $"❌ {title}", MessageBoxButton.OK, MessageBoxImage.Error)
    End Sub

    ''' <summary>
    ''' Mostra il riepilogo dei costi AI
    ''' </summary>
    Private Sub ShowAICostSummary(costSummary As String)
        Try
            ' Mostra il riepilogo in una finestra di dialogo dedicata
            Dim result = MessageBox.Show(
                costSummary & vbCrLf & vbCrLf &
                "💡 NOTA: I costi mostrati sono calcolati sui prezzi pubblici di OpenAI." & vbCrLf &
                "Per informazioni sul credito rimanente, controlla il tuo account OpenAI." & vbCrLf & vbCrLf &
                "Vuoi aprire la dashboard OpenAI per controllare il tuo saldo?",
                "💰 Riepilogo Costi AI",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information)

            ' Se l'utente vuole aprire la dashboard OpenAI
            If result = MessageBoxResult.Yes Then
                Try
                    Process.Start(New ProcessStartInfo("https://platform.openai.com/usage") With {.UseShellExecute = True})
                Catch ex As Exception
                    Debug.WriteLine($"DEBUG: Errore apertura browser: {ex.Message}")
                    ' Fallback: copia URL negli appunti
                    Try
                        Clipboard.SetText("https://platform.openai.com/usage")
                        ShowInfo("URL copiato negli appunti: https://platform.openai.com/usage")
                    Catch
                        ShowInfo("Vai manualmente su: https://platform.openai.com/usage")
                    End Try
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine($"DEBUG: Errore visualizzazione riepilogo costi: {ex.Message}")
            ShowError("Errore", "Errore nella visualizzazione del riepilogo costi: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Effettua scroll automatico alla sezione anteprima
    ''' </summary>
    Private Sub ScrollToPreviewPanel()
        Try
            ' Usa Dispatcher per assicurarsi che l'UI sia aggiornata prima dello scroll
            Dispatcher.BeginInvoke(Sub()
                Try
                    ' BringIntoView forza il controllo a essere visibile nell'area di scroll
                    PreviewPanel.BringIntoView()
                    Debug.WriteLine("DEBUG: Auto-scroll alla sezione anteprima completato")
                Catch ex As Exception
                    Debug.WriteLine($"DEBUG: Errore auto-scroll anteprima: {ex.Message}")
                End Try
            End Sub, System.Windows.Threading.DispatcherPriority.Loaded)

        Catch ex As Exception
            Debug.WriteLine($"DEBUG: Errore programmazione auto-scroll: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Mostra un messaggio di stato temporaneo
    ''' </summary>
    Private Sub ShowStatusMessage(message As String, Optional isSuccess As Boolean = False)
        Try
            ' Usa il meccanismo di notifica esistente o MessageBox se necessario
            If isSuccess Then
                ShowSuccess(message)
            Else
                ShowInfo(message)
            End If
        Catch ex As Exception
            ' Fallback silenzioso
            Debug.WriteLine($"Error showing status message: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Effettua scroll automatico verso il pannello di classificazione manuale
    ''' </summary>
    Private Sub ScrollToManualPanel()
        ScrollToPanel(ManualClassificationPanel)
    End Sub
    
    ''' <summary>
    ''' Scroll automatico generico per qualsiasi pannello
    ''' </summary>
    Private Sub ScrollToPanel(panel As FrameworkElement)
        Try
            ' Metodo semplice e affidabile: usa BringIntoView
            If panel IsNot Nothing Then
                panel.BringIntoView()
                
                ' Log per debug
                MoneyMind.Services.LoggingService.LogInfo("ScrollToPanel", $"Scroll automatico verso pannello {panel.Name}")
            End If
        Catch ex As Exception
            ' Fallback silenzioso
            MoneyMind.Services.LoggingService.LogWarning("ScrollToPanel", "Errore scroll automatico", ex.Message)
            Debug.WriteLine($"Error scrolling to panel: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Seleziona tutte le transazioni nell'anteprima
    ''' </summary>
    Private Sub BtnSelectAll_Click(sender As Object, e As RoutedEventArgs)
        Try
            If _previewResults IsNot Nothing Then
                For Each item In _previewResults
                    item.IsSelected = True
                Next
                
                ' Aggiorna la visualizzazione
                PreviewDataGrid.Items.Refresh()
                UpdateSelectionCount()
                
                MoneyMind.Services.LoggingService.LogInfo("SelectAll", $"Selezionate tutte {_previewResults.Count} transazioni")
            End If
        Catch ex As Exception
            MoneyMind.Services.LoggingService.LogError("SelectAll", ex)
            ShowError("Errore", "Errore durante la selezione di tutte le transazioni: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Deseleziona tutte le transazioni nell'anteprima
    ''' </summary>
    Private Sub BtnDeselectAll_Click(sender As Object, e As RoutedEventArgs)
        Try
            If _previewResults IsNot Nothing Then
                For Each item In _previewResults
                    item.IsSelected = False
                Next
                
                ' Aggiorna la visualizzazione
                PreviewDataGrid.Items.Refresh()
                UpdateSelectionCount()
                
                MoneyMind.Services.LoggingService.LogInfo("DeselectAll", "Deselezionate tutte le transazioni")
            End If
        Catch ex As Exception
            MoneyMind.Services.LoggingService.LogError("DeselectAll", ex)
            ShowError("Errore", "Errore durante la deselezione di tutte le transazioni: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Aggiorna il contatore delle transazioni selezionate
    ''' </summary>
    Private Sub UpdateSelectionCount()
        Try
            If _previewResults IsNot Nothing AndAlso LblSelectionCount IsNot Nothing Then
                Dim selectedCount As Integer = 0
                Dim totalCount = _previewResults.Count
                
                For Each item In _previewResults
                    If item.IsSelected Then
                        selectedCount += 1
                    End If
                Next
                
                LblSelectionCount.Text = $"Selezionate: {selectedCount}/{totalCount}"
                
                ' Abilita/disabilita pulsante approvazione selezionate
                If BtnApproveSelected IsNot Nothing Then
                    BtnApproveSelected.IsEnabled = selectedCount > 0
                End If
            End If
        Catch ex As Exception
            MoneyMind.Services.LoggingService.LogError("UpdateSelectionCount", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Gestisce il cambio di selezione di una transazione
    ''' </summary>
    Private Sub OnItemSelectionChanged(item As TransactionPreview)
        Try
            ' Aggiorna il contatore in tempo reale
            UpdateSelectionCount()
        Catch ex As Exception
            MoneyMind.Services.LoggingService.LogError("OnItemSelectionChanged", ex)
        End Try
    End Sub

#Region "Single Transaction Panel Event Handlers"

    ''' <summary>
    ''' Gestisce selezione transazione nel pannello singola classificazione
    ''' </summary>
    Private Sub DgSingleTransactions_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Try
            Dim selectedItem = TryCast(DgSingleTransactions.SelectedItem, ManualTransactionView)
            _selectedSingleTransaction = selectedItem

            BtnClassifySelectedSingle.IsEnabled = selectedItem IsNot Nothing

            ' Nasconde risultato precedente quando si cambia selezione
            If selectedItem IsNot Nothing Then
                SingleClassificationResult.Visibility = Visibility.Collapsed
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error in DgSingleTransactions_SelectionChanged")
        End Try
    End Sub

    ''' <summary>
    ''' Ricarica lista transazioni non classificate
    ''' </summary>
    Private Async Sub BtnRefreshSingleList_Click(sender As Object, e As RoutedEventArgs)
        Try
            BtnRefreshSingleList.IsEnabled = False
            Await LoadUnclassifiedTransactionsForSinglePanel()
        Catch ex As Exception
            _logger?.LogError(ex, "Error refreshing single transactions list")
            ShowError("Errore", "Errore durante il refresh: " & ex.Message)
        Finally
            BtnRefreshSingleList.IsEnabled = True
        End Try
    End Sub

    ''' <summary>
    ''' Classifica transazione selezionata e mostra anteprima
    ''' </summary>
    Private Async Sub BtnClassifySelectedSingle_Click(sender As Object, e As RoutedEventArgs)
        Try
            If _selectedSingleTransaction Is Nothing Then
                ShowError("Errore", "Selezionare una transazione dalla lista.")
                Return
            End If

            BtnClassifySelectedSingle.IsEnabled = False
            BtnClassifySelectedSingle.Content = "🔄 Classificazione in corso..."

            Dim startTime = DateTime.Now

            ' Esegui classificazione con il metodo selezionato
            Dim risultato = Await ExecuteClassificationByMethod(_selectedMethodForSingle, _selectedSingleTransaction.OriginalTransaction)

            Dim elapsed = DateTime.Now - startTime

            ' Mostra risultati nell'anteprima
            If risultato IsNot Nothing Then
                ' Aggiorna info transazione
                TxtSingleTransactionDetails.Text = $"{_selectedSingleTransaction.Descrizione} ({_selectedSingleTransaction.ImportoFormatted})"

                ' Aggiorna risultati classificazione
                TxtSingleMacro.Text = risultato.MacroCategoria
                TxtSingleCategory.Text = risultato.Categoria
                TxtSingleConfidence.Text = $"{risultato.PunteggioConfidenza:P1}"
                TxtSingleMethod.Text = GetMethodDisplayName(risultato.Motivazione)
                TxtSingleTime.Text = $"{elapsed.TotalSeconds:F1}s"
                TxtSingleMotivation.Text = risultato.Motivazione

                ' Aggiorna badge confidenza
                UpdateConfidenceBadge(SingleConfidenceBadge, TxtSingleConfidenceLevel, risultato.PunteggioConfidenza)

                ' Memorizza risultato per approvazione
                _currentClassificationResult = risultato

                ' Mostra anteprima
                SingleClassificationResult.Visibility = Visibility.Visible
            Else
                ShowError("Errore", "Impossibile classificare la transazione selezionata.")
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error classifying selected single transaction")
            ShowError("Errore", "Errore durante la classificazione: " & ex.Message)
        Finally
            BtnClassifySelectedSingle.IsEnabled = True
            BtnClassifySelectedSingle.Content = "🚀 Classifica Selezionata"
        End Try
    End Sub

    ''' <summary>
    ''' Applica classificazione alla transazione selezionata
    ''' </summary>
    Private Async Sub BtnApplySingleClassification_Click(sender As Object, e As RoutedEventArgs)
        Try
            If _selectedSingleTransaction Is Nothing OrElse _currentClassificationResult Is Nothing Then
                ShowError("Errore", "Nessuna classificazione da applicare.")
                Return
            End If

            Dim result = MessageBox.Show(
                $"Applicare questa classificazione alla transazione?" & vbCrLf & vbCrLf &
                $"Transazione: {_selectedSingleTransaction.Descrizione}" & vbCrLf &
                $"MacroCategoria: {_currentClassificationResult.MacroCategoria}" & vbCrLf &
                $"Categoria: {_currentClassificationResult.Categoria}" & vbCrLf &
                $"Confidenza: {_currentClassificationResult.PunteggioConfidenza:P1}",
                "Conferma Applicazione",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question)

            If result = MessageBoxResult.Yes Then
                ' Applica classificazione al database
                Await ApplySingleClassificationToDatabase()

                ' Refresh lista per rimuovere la transazione classificata
                Await LoadUnclassifiedTransactionsForSinglePanel()

                ' Nasconde risultato
                SingleClassificationResult.Visibility = Visibility.Collapsed

                ShowInfo("Classificazione applicata con successo!")
            End If

        Catch ex As Exception
            _logger?.LogError(ex, "Error applying single classification")
            ShowError("Errore", "Errore durante l'applicazione: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Applica classificazione al database
    ''' </summary>
    Private Async Function ApplySingleClassificationToDatabase() As Task
        Dim query = "UPDATE Transazioni
                    SET MacroCategoria = @MacroCategoria,
                        Categoria = @Categoria,
                        Necessita = @Necessita,
                        Frequenza = @Frequenza,
                        Stagionalita = @Stagionalita,
                        DataModifica = @DataModifica
                    WHERE ID = @ID"

        Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
            Await connection.OpenAsync()
            Using cmd As New SQLiteCommand(query, connection)
                cmd.Parameters.AddWithValue("@MacroCategoria", _currentClassificationResult.MacroCategoria)
                cmd.Parameters.AddWithValue("@Categoria", _currentClassificationResult.Categoria)
                cmd.Parameters.AddWithValue("@Necessita", _currentClassificationResult.Necessita)
                cmd.Parameters.AddWithValue("@Frequenza", _currentClassificationResult.Frequenza)
                cmd.Parameters.AddWithValue("@Stagionalita", _currentClassificationResult.Stagionalita)
                cmd.Parameters.AddWithValue("@DataModifica", DateTime.Now)
                cmd.Parameters.AddWithValue("@ID", _selectedSingleTransaction.ID)

                Await cmd.ExecuteNonQueryAsync()
            End Using
        End Using

        ' Salva automaticamente pattern se l'opzione è attiva
        If ChkAutoSavePatterns?.IsChecked = True Then
            Await SalvaPatternDaRisultato(_selectedSingleTransaction.OriginalTransaction, _currentClassificationResult)
        End If
    End Function

    ''' <summary>
    ''' Scarta risultato classificazione
    ''' </summary>
    Private Sub BtnDiscardSingleResult_Click(sender As Object, e As RoutedEventArgs)
        SingleClassificationResult.Visibility = Visibility.Collapsed
        _currentClassificationResult = Nothing
    End Sub

    ''' <summary>
    ''' Annulla pannello classificazione singola
    ''' </summary>
    Private Sub BtnCancelSingleClassification_Click(sender As Object, e As RoutedEventArgs)
        SingleClassificationPanel.Visibility = Visibility.Collapsed
        SingleClassificationResult.Visibility = Visibility.Collapsed
        _selectedSingleTransaction = Nothing
        _currentClassificationResult = Nothing
    End Sub

    ''' <summary>
    ''' Salva pattern dal risultato di classificazione
    ''' </summary>
    Private Async Function SalvaPatternDaRisultato(transazione As Transazione, risultato As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione) As Task
        Try
            If risultato Is Nothing OrElse String.IsNullOrEmpty(risultato.MacroCategoria) Then
                Return
            End If

            ' Estrai parola chiave dalla descrizione
            Dim parolaChiave = EstrayParolaChiave(transazione.Descrizione)
            If String.IsNullOrEmpty(parolaChiave) Then
                Return
            End If

            ' Controlla se pattern esiste già
            Dim existingPattern = Await GetPatternByParola(parolaChiave)
            If existingPattern IsNot Nothing Then
                Return ' Pattern già esistente
            End If

            ' Inserisce nuovo pattern
            Dim query = "INSERT INTO Pattern (Parola, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita, Peso)
                        VALUES (@Parola, @MacroCategoria, @Categoria, @Necessita, @Frequenza, @Stagionalita, 5)"

            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                Await connection.OpenAsync()
                Using cmd As New SQLiteCommand(query, connection)
                    cmd.Parameters.AddWithValue("@Parola", parolaChiave)
                    cmd.Parameters.AddWithValue("@MacroCategoria", risultato.MacroCategoria)
                    cmd.Parameters.AddWithValue("@Categoria", risultato.Categoria)
                    cmd.Parameters.AddWithValue("@Necessita", risultato.Necessita)
                    cmd.Parameters.AddWithValue("@Frequenza", risultato.Frequenza)
                    cmd.Parameters.AddWithValue("@Stagionalita", risultato.Stagionalita)

                    Await cmd.ExecuteNonQueryAsync()
                End Using
            End Using

        Catch ex As Exception
            _logger?.LogError(ex, "Error saving pattern from result")
        End Try
    End Function

    ''' <summary>
    ''' Ottiene pattern esistente per parola chiave
    ''' </summary>
    Private Async Function GetPatternByParola(parolaChiave As String) As Task(Of Object)
        Try
            Dim query = "SELECT COUNT(*) FROM Pattern WHERE Parola = @Parola"
            Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
                Await connection.OpenAsync()
                Using cmd As New SQLiteCommand(query, connection)
                    cmd.Parameters.AddWithValue("@Parola", parolaChiave)
                    Dim count = Convert.ToInt32(Await cmd.ExecuteScalarAsync())
                    Return If(count > 0, New Object(), Nothing)
                End Using
            End Using
        Catch ex As Exception
            _logger?.LogError(ex, "Error getting pattern by parola")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Determina se una descrizione potrebbe essere un ristorante basandosi su pattern linguistici
    ''' </summary>
    Private Function IsLikelyRestaurant(descrizione As String) As Boolean
        If String.IsNullOrWhiteSpace(descrizione) Then
            Return False
        End If

        Dim desc = descrizione.ToUpperInvariant().Trim()

        ' Pattern tipici dei ristoranti/bar/locali
        Dim restaurantKeywords = {
            "BAR", "RISTORANTE", "TRATTORIA", "OSTERIA", "PIZZERIA", "TAVERNA",
            "BRACERIA", "PANINOTECA", "GELATERIA", "PASTICCERIA", "CAFFETTERIA",
            "LOCANDA", "BISTROT", "PUB", "WINE", "ENOTECA", "GRILL"
        }

        ' Pattern strutturali tipici dei ristoranti
        Dim restaurantPatterns = {
            "^(LA|IL|I|LE|GLI|L')\s+[A-ZÀ-ÖØ-Þ]{3,}", ' Articoli + nome (LA TAVERNA, I QUADRI, etc.)
            "^[A-ZÀ-ÖØ-Þ]{3,}\s+(E|&)\s+[A-ZÀ-ÖØ-Þ]", ' Nome + E + iniziale (PESCE E P.)
            "^[A-ZÀ-ÖØ-Þ]{2,}\s+[A-ZÀ-ÖØ-Þ]{2,}\s+(LAGO|VIA|PIAZZA|CORSO)", ' Nome + località
            "^(DA|DALLA|DAL)\s+[A-ZÀ-ÖØ-Þ]{3,}", ' DA MARIO, DALLA NONNA
            "'S\s" ' MARIO'S, etc.
        }

        ' Controllo diretto per parole chiave
        For Each keyword In restaurantKeywords
            If desc.Contains(keyword) Then
                Debug.WriteLine($"DEBUG: IsLikelyRestaurant - Trovata keyword: '{keyword}'")
                Return True
            End If
        Next

        ' Controllo pattern strutturali
        For Each pattern In restaurantPatterns
            If Regex.IsMatch(desc, pattern) Then
                Debug.WriteLine($"DEBUG: IsLikelyRestaurant - Match pattern: '{pattern}'")
                Return True
            End If
        Next

        ' Pattern più specifici basati su importi e località tipiche
        If desc.Contains("LAGO") AndAlso (desc.StartsWith("I ") OrElse desc.StartsWith("LA ") OrElse desc.StartsWith("IL ")) Then
            Debug.WriteLine($"DEBUG: IsLikelyRestaurant - Pattern lago + articolo: '{desc}'")
            Return True
        End If

        Debug.WriteLine($"DEBUG: IsLikelyRestaurant - Non riconosciuto come ristorante: '{desc}'")
        Return False
    End Function

    ''' <summary>
    ''' Determina se una transazione richiede validazione AI per parole ambigue o incoerenze evidenti
    ''' </summary>
    Private Function RequiresAIValidation(descrizione As String, classificazione As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione) As Boolean
        If String.IsNullOrWhiteSpace(descrizione) OrElse classificazione Is Nothing Then
            Return False
        End If

        Dim upperDesc = descrizione.ToUpper()

        ' Parole ambigue che richiedono sempre validazione AI
        Dim ambiguousWords = {
            "SALOON", "HOTEL", "ALBERGO", "BAR", "CENTRO", "PALACE", "RESORT",
            "GRAND", "ELITE", "ROYAL", "PLAZA", "CLUB", "LOUNGE", "SPA",
            "BEAUTY", "HAIR", "NAILS", "WELLNESS", "FITNESS", "GYM"
        }

        ' Controlla presenza di parole ambigue
        For Each word In ambiguousWords
            If upperDesc.Contains(word) Then
                Debug.WriteLine($"DEBUG: RequiresAIValidation - Trovata parola ambigua: '{word}'")
                Return True
            End If
        Next

        ' Controlla incoerenze evidenti specifiche
        ' Hotel classificati come ristorazione
        If (upperDesc.Contains("HOTEL") OrElse upperDesc.Contains("ALBERGO")) AndAlso
           classificazione.Categoria = "Ristorazione" Then
            Debug.WriteLine($"DEBUG: RequiresAIValidation - Hotel mal classificato come ristorazione")
            Return True
        End If

        ' Saloon/barbieri classificati come ristorazione
        If upperDesc.Contains("SALOON") AndAlso
           classificazione.MacroCategoria = "Alimentari e Ristorazione" Then
            Debug.WriteLine($"DEBUG: RequiresAIValidation - Saloon mal classificato come ristorazione")
            Return True
        End If

        ' Business con nomi propri che potrebbero essere mal classificati
        If HasSpecificBusinessName(upperDesc) Then
            Debug.WriteLine($"DEBUG: RequiresAIValidation - Business con nome specifico richiede controllo")
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Determina se la descrizione contiene un nome di business specifico
    ''' </summary>
    Private Function HasSpecificBusinessName(upperDesc As String) As Boolean
        ' Pattern per nomi di business specifici (non generici)
        Dim businessPatterns = {
            "^[A-Z]{2,}\s+[A-Z']{2,}\s+[A-Z]{2,}", ' ES: GENNY D'AURIA SALOON
            "^[A-Z]{3,}\s+[A-Z]{3,}\s+[A-Z]{3,}", ' ES: GRAND HOTEL ELITE
            "[A-Z]{2,}'[A-Z]\s+[A-Z]{2,}", ' Nomi con apostrofi
            "^[A-Z]{2,}\s+[A-Z]{2,}\s+\d+" ' Nomi con numeri
        }

        For Each pattern In businessPatterns
            If Regex.IsMatch(upperDesc, pattern) Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Valida la qualità e accuratezza del risultato AI
    ''' </summary>
    Private Function ValidateAIResult(suggerimento As MoneyMind.Services.GptClassificatoreTransazioni.SuggerimentoClassificazione, descrizione As String) As Boolean
        Try
            Debug.WriteLine($"DEBUG: 🔍 VALIDAZIONE AI - Descrizione: '{TruncateString(descrizione, 50)}'")

            ' Controllo 1: Verificare che MacroCategoria e Categoria non siano vuote
            If String.IsNullOrEmpty(suggerimento.MacroCategoria) OrElse String.IsNullOrEmpty(suggerimento.Categoria) Then
                Debug.WriteLine($"DEBUG: ❌ VALIDAZIONE FALLITA - MacroCategoria o Categoria vuote")
                Return False
            End If

            ' Controllo 2: Verificare che non sia "Non Classificata" (a meno che non sia davvero non classificabile)
            If suggerimento.MacroCategoria = "Non Classificata" AndAlso suggerimento.Categoria = "Non Classificata" Then
                Debug.WriteLine($"DEBUG: ❌ VALIDAZIONE FALLITA - Risultato è 'Non Classificata'")
                Return False
            End If

            ' Controllo 3: Verificare che ParolaChiave abbia senso rispetto alla descrizione
            If Not String.IsNullOrEmpty(suggerimento.ParolaChiave) Then
                Dim descUpper = descrizione.ToUpper()
                Dim keywordUpper = suggerimento.ParolaChiave.ToUpper()

                ' La parola chiave dovrebbe apparire nella descrizione o essere correlata
                If Not descUpper.Contains(keywordUpper) Then
                    Debug.WriteLine($"DEBUG: ⚠️ VALIDAZIONE PARZIALE - ParolaChiave '{suggerimento.ParolaChiave}' non trovata in descrizione")
                    ' Non fallire completamente, ma segnala
                End If
            End If

            ' Controllo 4: Verificare confidenza minima
            If suggerimento.Confidenza < 0.5 Then
                Debug.WriteLine($"DEBUG: ❌ VALIDAZIONE FALLITA - Confidenza troppo bassa: {suggerimento.Confidenza:P1}")
                Return False
            End If

            ' Controllo 5: Verificare combinazioni logiche di categorie
            If Not ValidateCategoryLogic(suggerimento.MacroCategoria, suggerimento.Categoria) Then
                Debug.WriteLine($"DEBUG: ❌ VALIDAZIONE FALLITA - Combinazione categorie non valida")
                Return False
            End If

            Debug.WriteLine($"DEBUG: ✅ VALIDAZIONE SUPERATA - Risultato AI di alta qualità")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"DEBUG: ❌ ERRORE VALIDAZIONE: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Verifica che la combinazione MacroCategoria-Categoria sia logicamente corretta
    ''' </summary>
    Private Function ValidateCategoryLogic(macroCategoria As String, categoria As String) As Boolean
        ' Definisce combinazioni valide note
        Dim combinazioniValide As New Dictionary(Of String, String()) From {
            {"Viaggi e Trasporti", {"Alloggio", "Trasporti", "Carburante", "Viaggi"}},
            {"Salute e Benessere", {"Parrucchiere/Barbiere", "Estetica", "Sport e Fitness", "Farmacia", "Visite Mediche"}},
            {"Spesa e Alimentari", {"Alimentari", "Supermercato", "Bevande"}},
            {"Ristorazione", {"Ristorante", "Bar/Caffè", "Fast Food", "Pizzeria"}},
            {"Casa e Famiglia", {"Utenze", "Manutenzione", "Arredamento", "Elettrodomestici"}},
            {"Stipendio e Entrate", {"Stipendio", "Bonus", "Rimborsi", "Altre Entrate"}}
        }

        If combinazioniValide.ContainsKey(macroCategoria) Then
            Return combinazioniValide(macroCategoria).Contains(categoria)
        End If

        ' Se non è nelle combinazioni note, considera valida (per nuove categorie)
        Return True
    End Function

    ''' <summary>
    ''' Classifica business specifici (hotel, barbieri, ecc.) con alta precisione
    ''' </summary>
    Private Function ClassifySpecificBusinessTypes(descrizione As String) As MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione
        Debug.WriteLine($"DEBUG: 🔍 ClassifySpecificBusinessTypes INIZIO - Descrizione: '{TruncateString(descrizione, 60)}'")

        If String.IsNullOrWhiteSpace(descrizione) Then
            Debug.WriteLine($"DEBUG: ❌ ClassifySpecificBusinessTypes - Descrizione vuota o null")
            Return Nothing
        End If

        Dim upperDesc = descrizione.ToUpper()

        ' Hotel e alloggi - Alta priorità
        If upperDesc.Contains("HOTEL") OrElse upperDesc.Contains("ALBERGO") OrElse
           upperDesc.Contains("RESORT") OrElse upperDesc.Contains("B&B") OrElse
           upperDesc.Contains("PENSIONE") OrElse upperDesc.Contains("MOTEL") Then
            Dim parolaChiave = ExtractBusinessNameFromDescription(upperDesc)
            Debug.WriteLine($"DEBUG: ✅ ClassifySpecificBusinessTypes - Riconosciuto HOTEL/ALLOGGIO - ParolaChiave: '{parolaChiave}'")
            Return New MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione With {
                .MacroCategoria = "Viaggi e Trasporti",
                .Categoria = "Alloggio",
                .Necessita = "Variabile",
                .Frequenza = "Saltuaria",
                .Stagionalita = "Variabile",
                .PunteggioConfidenza = 0.95,
                .LivelloConfidenza = "Alta",
                .ColoreConfidenza = "#4CAF50",
                .PatternUsato = "Hotel/Alloggio Detection",
                .ParolaChiave = parolaChiave,
                .IsValid = True
            }
        End If

        ' Barbieri e parrucchieri
        If upperDesc.Contains("SALOON") OrElse upperDesc.Contains("BARBIERE") OrElse
           upperDesc.Contains("PARRUCCHIERE") OrElse upperDesc.Contains("HAIR") OrElse
           upperDesc.Contains("BEAUTY") OrElse upperDesc.Contains("SALON") Then
            Dim parolaChiaveBarbiere = ExtractBusinessNameFromDescription(upperDesc)
            Debug.WriteLine($"DEBUG: ✅ ClassifySpecificBusinessTypes - Riconosciuto BARBIERE/PARRUCCHIERE - ParolaChiave: '{parolaChiaveBarbiere}'")
            Return New MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione With {
                .MacroCategoria = "Salute e Benessere",
                .Categoria = "Parrucchiere/Barbiere",
                .Necessita = "Accessoria",
                .Frequenza = "Periodica",
                .Stagionalita = "Tutto l'anno",
                .PunteggioConfidenza = 0.9,
                .LivelloConfidenza = "Alta",
                .ColoreConfidenza = "#4CAF50",
                .PatternUsato = "Barbiere/Parrucchiere Detection",
                .ParolaChiave = parolaChiaveBarbiere,
                .IsValid = True
            }
        End If

        ' Centri estetici e SPA (pattern più specifici per evitare false positive)
        If (upperDesc.Contains(" SPA ") OrElse upperDesc.Contains("SPA ") OrElse upperDesc.EndsWith(" SPA")) OrElse
           upperDesc.Contains("WELLNESS") OrElse upperDesc.Contains("ESTETICA") OrElse
           upperDesc.Contains("NAILS") OrElse upperDesc.Contains("CENTRO ESTETICO") Then
            Dim parolaChiaveEstetica = ExtractBusinessNameFromDescription(upperDesc)
            Debug.WriteLine($"DEBUG: ✅ ClassifySpecificBusinessTypes - Riconosciuto CENTRO ESTETICO - ParolaChiave: '{parolaChiaveEstetica}'")
            Return New MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione With {
                .MacroCategoria = "Salute e Benessere",
                .Categoria = "Estetica",
                .Necessita = "Accessoria",
                .Frequenza = "Periodica",
                .Stagionalita = "Tutto l'anno",
                .PunteggioConfidenza = 0.9,
                .LivelloConfidenza = "Alta",
                .ColoreConfidenza = "#4CAF50",
                .PatternUsato = "Centro Estetico Detection",
                .ParolaChiave = parolaChiaveEstetica,
                .IsValid = True
            }
        End If

        ' Palestre e fitness
        If upperDesc.Contains("GYM") OrElse upperDesc.Contains("FITNESS") OrElse
           upperDesc.Contains("PALESTRA") OrElse upperDesc.Contains("SPORT") Then
            Dim parolaChiaveFitness = ExtractBusinessNameFromDescription(upperDesc)
            Debug.WriteLine($"DEBUG: ✅ ClassifySpecificBusinessTypes - Riconosciuto PALESTRA/FITNESS - ParolaChiave: '{parolaChiaveFitness}'")
            Return New MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione With {
                .MacroCategoria = "Salute e Benessere",
                .Categoria = "Sport e Fitness",
                .Necessita = "Accessoria",
                .Frequenza = "Periodica",
                .Stagionalita = "Tutto l'anno",
                .PunteggioConfidenza = 0.9,
                .LivelloConfidenza = "Alta",
                .ColoreConfidenza = "#4CAF50",
                .PatternUsato = "Palestra/Fitness Detection",
                .ParolaChiave = parolaChiaveFitness,
                .IsValid = True
            }
        End If

        ' Farmacie e sanitario
        If upperDesc.Contains("FARMACIA") OrElse upperDesc.Contains("PHARMACY") OrElse
           upperDesc.Contains("OSPEDALE") OrElse upperDesc.Contains("CLINICA") Then
            Dim parolaChiaveFarmacia = ExtractBusinessNameFromDescription(upperDesc)
            Debug.WriteLine($"DEBUG: ✅ ClassifySpecificBusinessTypes - Riconosciuto FARMACIA/SANITARIO - ParolaChiave: '{parolaChiaveFarmacia}'")
            Return New MoneyMind.Services.ClassificatoreTransazioni.RisultatoClassificazione With {
                .MacroCategoria = "Salute e Benessere",
                .Categoria = "Farmacia",
                .Necessita = "Primaria",
                .Frequenza = "Al bisogno",
                .Stagionalita = "Tutto l'anno",
                .PunteggioConfidenza = 0.95,
                .LivelloConfidenza = "Alta",
                .ColoreConfidenza = "#4CAF50",
                .PatternUsato = "Farmacia/Sanitario Detection",
                .ParolaChiave = parolaChiaveFarmacia,
                .IsValid = True
            }
        End If

        Debug.WriteLine($"DEBUG: ❌ ClassifySpecificBusinessTypes - Nessun business type riconosciuto per: '{TruncateString(descrizione, 60)}'")
        Return Nothing
    End Function

    ''' <summary>
    ''' Determina se una descrizione dovrebbe essere analizzata con Google Places
    ''' </summary>
    Private Function ShouldUseGooglePlaces(descrizione As String) As Boolean
        If String.IsNullOrWhiteSpace(descrizione) Then
            Return False
        End If

        Dim upperDesc = descrizione.ToUpper()

        ' Business con nomi specifici che beneficiano da Google Places
        Dim businessTypeKeywords = {
            "HOTEL", "ALBERGO", "RESORT", "PENSIONE", "MOTEL",
            "SALOON", "BARBIERE", "PARRUCCHIERE", "SALON", "HAIR", "BEAUTY",
            "RISTORANTE", "TRATTORIA", "PIZZERIA", "BAR", "CAFFÈ", "PUB",
            "FARMACIA", "PHARMACY", "CLINICA", "OSPEDALE",
            "PALESTRA", "GYM", "FITNESS", "SPA", "WELLNESS",
            "CENTRO", "PALACE", "GRAND", "ELITE", "ROYAL", "PLAZA"
        }

        ' Controlla presenza di keyword business
        For Each keyword In businessTypeKeywords
            If upperDesc.Contains(keyword) Then
                Debug.WriteLine($"DEBUG: ShouldUseGooglePlaces - Trovata keyword business: '{keyword}'")
                Return True
            End If
        Next

        ' Business con nomi propri (non generici)
        If HasSpecificBusinessName(upperDesc) Then
            Debug.WriteLine($"DEBUG: ShouldUseGooglePlaces - Business con nome specifico")
            Return True
        End If

        ' Pattern tipici di transazioni POS con nomi di business
        If upperDesc.Contains("POS") AndAlso HasBusinessNamePattern(upperDesc) Then
            Debug.WriteLine($"DEBUG: ShouldUseGooglePlaces - Transazione POS con nome business")
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Rileva pattern tipici di nomi di business nelle transazioni POS
    ''' </summary>
    Private Function HasBusinessNamePattern(upperDesc As String) As Boolean
        ' Pattern per identificare nomi di business in transazioni POS
        Dim businessNamePatterns = {
            "[A-Z]{3,}\s+[A-Z']{3,}\s+[A-Z]{3,}", ' ES: GENNY D'AURIA SALOON
            "[A-Z]{4,}\s+[A-Z]{4,}\s+[A-Z]{4,}", ' ES: GRAND HOTEL ELITE
            "[A-Z]{3,}\s+[A-Z]{3,}\s+\d+", ' Nome + numeri
            "[A-Z']+\s+[A-Z']+\s+[A-Z']+", ' Nomi con apostrofi multipli
            "^[A-Z]{2,}\s+[A-Z]{2,}\s+[A-Z]{2,}.*[A-Z]{3,}$" ' Pattern business complessi
        }

        For Each pattern In businessNamePatterns
            If Regex.IsMatch(upperDesc, pattern) Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Estrae il nome del business dalla descrizione della transazione (ottimizzato per riutilizzo)
    ''' </summary>
    Private Function ExtractBusinessNameFromDescription(upperDesc As String) As String
        Try
            Debug.WriteLine($"DEBUG: 🔍 ESTRAZIONE PAROLA CHIAVE - Input: '{TruncateString(upperDesc, 60)}'")

            Dim cleaned = upperDesc.Trim()

            ' STEP 1: Estrazione specifica per transazioni POS italiane
            If cleaned.Contains("OPERAZIONE POS EUROZONA DEL") Then
                ' Pattern migliorato per catturare solo il nome del business
                ' Esempio: "Operazione POS Eurozona Del 10.09.25 18:18 Carta *1250 GENNY D'AURIA SALOON . SAN VITALIANO ITA"
                ' Cattura: "GENNY D'AURIA SALOON"

                Dim posMatch = Regex.Match(cleaned, "CARTA \*\d+\s+(.+?)\s*\.\s*[\w\s]*ITA", RegexOptions.IgnoreCase)
                If posMatch.Success Then
                    cleaned = posMatch.Groups(1).Value.Trim()
                    Debug.WriteLine($"DEBUG: ✅ MATCH POS - Estratto: '{cleaned}'")
                Else
                    ' Fallback pattern più semplice
                    Dim fallbackMatch = Regex.Match(cleaned, "CARTA \*\d+\s+(.+?)(?:\s+\d+|\s+\.|$)", RegexOptions.IgnoreCase)
                    If fallbackMatch.Success Then
                        cleaned = fallbackMatch.Groups(1).Value.Trim()
                        Debug.WriteLine($"DEBUG: ⚠️ FALLBACK POS - Estratto: '{cleaned}'")
                    End If
                End If
            End If

            ' STEP 2: Pulizia nome business dai suffissi geografici comuni
            Dim locationSuffixes = {"ITA", "ITALIA", "ROMA", "MILANO", "NAPOLI", "TORINO", "BOLOGNA", "MODENA", "SAN VITALIANO", "FIRENZE", "VENEZIA", "PALERMO"}
            For Each suffix In locationSuffixes
                If cleaned.EndsWith(" " & suffix, StringComparison.OrdinalIgnoreCase) Then
                    cleaned = cleaned.Substring(0, cleaned.Length - suffix.Length - 1).Trim()
                    Debug.WriteLine($"DEBUG: 🗺️ RIMOSSO SUFFISSO GEOGRAFICO: '{suffix}' - Risultato: '{cleaned}'")
                End If
            Next

            ' STEP 3: Rimuovi numeri finali (indirizzi civici)
            cleaned = Regex.Replace(cleaned, "\s+\d+(\s+\d+)*$", "").Trim()

            ' STEP 4: Rimuovi caratteri finali di punteggiatura
            cleaned = cleaned.TrimEnd("."c, ","c, ";"c, ":"c, " "c)

            ' STEP 5: Normalizza spazi multipli
            cleaned = Regex.Replace(cleaned, "\s+", " ").Trim()

            ' STEP 6: Controllo qualità risultato
            If String.IsNullOrEmpty(cleaned) OrElse cleaned.Length < 3 Then
                Debug.WriteLine($"DEBUG: ❌ RISULTATO TROPPO CORTO - Fallback a estrazione automatica")

                ' Fallback intelligente
                If _gptClassificatore IsNot Nothing Then
                    cleaned = _gptClassificatore.EstraiParolaChiaveDaDescrizione(upperDesc)
                Else
                    cleaned = EstrayParolaChiave(upperDesc)
                End If
            End If

            ' STEP 7: Capitalizzazione corretta (solo prima lettera di ogni parola)
            cleaned = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(cleaned.ToLower())

            Debug.WriteLine($"DEBUG: 🎯 PAROLA CHIAVE FINALE: '{cleaned}'")
            Return cleaned

        Catch ex As Exception
            Debug.WriteLine($"DEBUG: ❌ ERRORE ESTRAZIONE PAROLA CHIAVE: {ex.Message}")
            Return upperDesc ' Fallback sicuro
        End Try
    End Function

#End Region


#End Region

End Class