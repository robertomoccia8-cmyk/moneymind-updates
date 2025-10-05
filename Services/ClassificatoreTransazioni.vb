Imports System.Data.SQLite
Imports System.Text.RegularExpressions
Imports Microsoft.Extensions.Logging

Namespace MoneyMind.Services
    ''' <summary>
    ''' Sistema avanzato di classificazione automatica delle transazioni
    ''' Utilizza pattern matching intelligente con punteggio di confidenza
    ''' </summary>
    Public Class ClassificatoreTransazioni
        Private ReadOnly _logger As ILogger
        Private Const SogliaMatch As Double = 0.5
        Private Const SogliaAlta As Double = 0.8
        Private Const SogliaMedia As Double = 0.65

        Public Sub New(logger As ILogger)
            _logger = logger
        End Sub

        ''' <summary>
        ''' Risultato della classificazione con informazioni dettagliate
        ''' </summary>
        Public Class RisultatoClassificazione
            Public Property MacroCategoria As String = ""
            Public Property Categoria As String = ""
            Public Property Necessita As String = ""
            Public Property Frequenza As String = ""
            Public Property Stagionalita As String = ""
            Public Property PunteggioConfidenza As Double
            Public Property PatternUsato As String = ""
            Public Property LivelloConfidenza As String = ""
            Public Property ColoreConfidenza As String = "#FF9800" ' Default arancione
            Public Property Motivazione As String = ""
            Public Property IsValid As Boolean = True
            Public Property ParolaChiave As String = ""

            Public ReadOnly Property DescrizioneConfidenza As String
                Get
                    If PunteggioConfidenza >= SogliaAlta Then
                        Return "Classificazione molto sicura"
                    ElseIf PunteggioConfidenza >= SogliaMedia Then
                        Return "Classificazione abbastanza sicura"
                    Else
                        Return "Classificazione incerta - verifica manuale consigliata"
                    End If
                End Get
            End Property
        End Class

        ''' <summary>
        ''' Normalizza il testo per il matching, gestendo abbreviazioni e caratteri speciali
        ''' </summary>
        Public Shared Function NormalizzaTesto(testo As String) As String
            If String.IsNullOrEmpty(testo) Then Return ""

            Dim normalizzato As String = testo.ToLowerInvariant().Trim()

            ' Rimuove caratteri speciali mantenendo lettere, numeri e spazi
            normalizzato = Regex.Replace(normalizzato, "[^\w\s]", " ")

            ' Unifica spazi multipli
            normalizzato = Regex.Replace(normalizzato, "\s+", " ").Trim()

            ' Espande abbreviazioni comuni
            normalizzato = EspandiAbbreviazioni(normalizzato)

            Return normalizzato
        End Function

        ''' <summary>
        ''' Espande abbreviazioni comuni italiane per migliorare il matching
        ''' </summary>
        Private Shared Function EspandiAbbreviazioni(testo As String) As String
            Dim abbreviazioni As New Dictionary(Of String, String) From {
                {"srl", "societa responsabilita limitata"},
                {"spa", "societa per azioni"},
                {"snc", "societa nome collettivo"},
                {"sas", "societa accomandita semplice"},
                {"coop", "cooperativa"},
                {"superm", "supermercato"},
                {"farm", "farmacia"},
                {"rist", "ristorante"},
                {"tratt", "trattoria"},
                {"pizz", "pizzeria"},
                {"bar", "bar"},
                {"tab", "tabaccheria"},
                {"gior", "giornale"},
                {"benz", "benzina"},
                {"carb", "carburante"},
                {"autost", "autostrada"},
                {"parch", "parcheggio"},
                {"aut", "automobile"},
                {"cons", "consorzio"},
                {"serv", "servizio"},
                {"uff", "ufficio"},
                {"amm", "amministrazione"},
                {"com", "comune"},
                {"prov", "provincia"},
                {"reg", "regione"},
                {"min", "ministero"},
                {"univ", "universita"},
                {"osp", "ospedale"},
                {"asl", "azienda sanitaria locale"},
                {"inps", "istituto nazionale previdenza sociale"},
                {"inail", "istituto nazionale infortuni lavoro"},
                {"tel", "telefono"},
                {"mob", "mobile"},
                {"int", "internet"},
                {"tv", "televisione"},
                {"gas", "gas"},
                {"ele", "elettrica"},
                {"elet", "elettrica"},
                {"acq", "acqua"},
                {"rif", "rifiuti"},
                {"spaz", "spazzatura"},
                {"oebb", "osterreichische bundesbahnen ferrovie austriache"},
                {"trenitalia", "trenitalia ferrovie"},
                {"sncf", "societe nationale chemins fer francais"},
                {"db", "deutsche bahn ferrovie tedesche"},
                {"renfe", "red nacional ferrocarriles espanoles"},
                {"blue star", "blue star ferries traghetti grecia"},
                {"minoan", "minoan lines traghetti"},
                {"anek", "anek lines traghetti"},
                {"superfast", "superfast ferries traghetti"},
                {"grimaldi", "grimaldi lines traghetti"},
                {"aspit", "autostrade pedemontana lombarda pedaggio"},
                {"autostrade", "autostrade pedaggio casello"},
                {"tangenziali", "tangenziali pedaggio"},
                {"q8", "q8 stazione servizio carburante"},
                {"agip", "agip eni stazione servizio"},
                {"esso", "esso exxon stazione servizio"},
                {"shell", "shell stazione servizio"},
                {"total", "totalenergies stazione servizio"},
                {"tamoil", "tamoil stazione servizio"},
                {"ip", "ip italiana petroli stazione servizio"}
            }

            ' FIX: Usa Regex con word boundary per evitare sostituzioni parziali
            ' (es. "mensa" non deve matchare "en" → "energia")
            Dim risultato As String = testo
            For Each kvp In abbreviazioni
                ' Match solo parole intere (\b = word boundary)
                Dim pattern As String = "\b" & Regex.Escape(kvp.Key) & "\b"
                risultato = Regex.Replace(risultato, pattern, kvp.Value, RegexOptions.IgnoreCase)
            Next

            Return risultato
        End Function

        ''' <summary>
        ''' Riconosce aziende di trasporto europee per pre-classificazione accurata
        ''' </summary>
        Public Shared Function RiconosciAziendaTrasporto(descrizione As String) As RisultatoClassificazione
            Dim testoNorm = NormalizzaTesto(descrizione)
            Dim risultato As New RisultatoClassificazione()

            ' SISTEMA DI RICONOSCIMENTO SEMANTICO AVANZATO
            ' Verifica che non sia una classificazione errata
            If ContieneEsclusioni(testoNorm) Then
                Return risultato ' Return vuoto se contiene esclusioni
            End If

            ' Ferrovie europee
            If testoNorm.Contains("oebb") OrElse testoNorm.Contains("osterreichische bundesbahnen") Then
                risultato.MacroCategoria = "Trasporti"
                risultato.Categoria = "Treni"
                risultato.Necessita = "Utile"
                risultato.Frequenza = "Occasionale"
                risultato.Stagionalita = "Tutto l'Anno"
                risultato.PunteggioConfidenza = 0.95
                risultato.PatternUsato = "OEBB Austrian Railways"
                risultato.ParolaChiave = "OEBB"
                Return risultato
            End If

            ' Traghetti greci
            If testoNorm.Contains("blue star") Then
                risultato.MacroCategoria = "Trasporti"
                risultato.Categoria = "Traghetti"
                risultato.Necessita = "Utile"
                risultato.Frequenza = "Occasionale"
                risultato.Stagionalita = "Estate"
                risultato.PunteggioConfidenza = 0.95
                risultato.PatternUsato = "Blue Star Ferries"
                risultato.ParolaChiave = "Blue Star"
                Return risultato
            End If

            ' Autostrade italiane
            If testoNorm.Contains("aspit") OrElse testoNorm.Contains("pedaggio") Then
                risultato.MacroCategoria = "Trasporti"
                risultato.Categoria = "Pedaggi Autostradali"
                risultato.Necessita = "Essenziale"
                risultato.Frequenza = "Frequente"
                risultato.Stagionalita = "Tutto l'Anno"
                risultato.PunteggioConfidenza = 0.9
                risultato.PatternUsato = "Highway Tolls"
                risultato.ParolaChiave = "ASPIT"
                Return risultato
            End If

            ' Stazioni di servizio - PATTERN PRECISI per evitare false positive
            ' IP: deve essere IP isolato, non parte di altre parole
            If (testoNorm.Contains(" ip ") OrElse testoNorm.Contains("distributore ip") OrElse
                testoNorm.EndsWith(" ip") OrElse testoNorm.StartsWith("ip ")) AndAlso
                Not ContieneEsclusioni(testoNorm) Then
                risultato.MacroCategoria = "Trasporti"
                risultato.Categoria = "Carburante"
                risultato.Necessita = "Essenziale"
                risultato.Frequenza = "Frequente"
                risultato.Stagionalita = "Tutto l'Anno"
                risultato.PunteggioConfidenza = 0.9
                risultato.PatternUsato = "IP Gas Station"
                risultato.ParolaChiave = "IP"
                Return risultato
            End If

            ' Altri brand carburanti - pattern più sicuri
            Dim fuelBrands As String() = {"q8", "agip", "esso", "shell", "total", "tamoil"}
            For Each brand In fuelBrands
                If testoNorm.Contains(brand) AndAlso Not ContieneEsclusioni(testoNorm) Then
                    risultato.MacroCategoria = "Trasporti"
                    risultato.Categoria = "Carburante"
                    risultato.Necessita = "Essenziale"
                    risultato.Frequenza = "Frequente"
                    risultato.Stagionalita = "Tutto l'Anno"
                    risultato.PunteggioConfidenza = 0.9
                    risultato.PatternUsato = $"{brand.ToUpper()} Fuel Station"
                    risultato.ParolaChiave = brand.ToUpper()
                    Return risultato
                End If
            Next

            ' Nessun riconoscimento
            risultato.PunteggioConfidenza = 0
            Return risultato
        End Function

        ''' <summary>
        ''' Calcola la similarità tra pattern e testo basandosi sulle parole comuni
        ''' </summary>
        Public Shared Function CalcolaSimilaritaParole(pattern As String, testo As String) As Double
            Dim parolePattern() As String = NormalizzaTesto(pattern).Split(" "c, StringSplitOptions.RemoveEmptyEntries)
            Dim testoNormalizzato As String = NormalizzaTesto(testo)

            If parolePattern.Length = 0 Then Return 0.0

            Dim paroleMatchate As Integer = 0
            For Each parola In parolePattern
                If parola.Length > 2 AndAlso testoNormalizzato.Contains(parola) Then
                    paroleMatchate += 1
                End If
            Next

            Return CDbl(paroleMatchate) / CDbl(parolePattern.Length)
        End Function

        ''' <summary>
        ''' Calcola il punteggio di match combinando match esatto e similarità parole
        ''' </summary>
        Public Shared Function CalcolaPunteggioMatch(descrizioneTransazione As String, pattern As String, peso As Integer) As Double
            Dim testoNorm As String = NormalizzaTesto(descrizioneTransazione)
            Dim patternNorm As String = NormalizzaTesto(pattern)

            Dim punteggio As Double = 0.0
            Dim hasExactMatch As Boolean = False

            ' Match esatto (60% del punteggio base)
            If testoNorm.Contains(patternNorm) Then
                punteggio += 0.6
                hasExactMatch = True
            End If

            ' Similarità parole (40% del punteggio base)
            Dim similParole As Double = CalcolaSimilaritaParole(pattern, descrizioneTransazione)
            punteggio += 0.4 * similParole

            ' Moltiplicatore peso (massimo 2x)
            Dim moltiplicatorePeso As Double = Math.Min(CDbl(peso) / 5.0, 2.0)
            Dim punteggioFinale As Double = punteggio * moltiplicatorePeso

            Return punteggioFinale
        End Function

        ''' <summary>
        ''' Classifica una singola transazione utilizzando i pattern del database
        ''' </summary>
        Public Function ClassificaTransazione(transazione As Transazione) As RisultatoClassificazione
            Dim risultato As New RisultatoClassificazione()

            Try
                _logger?.LogInformation("Inizio classificazione per transazione: {Descrizione}", transazione.Descrizione)

                ' PRIORITÀ 1: Controlla prima le aziende di trasporto europee
                Dim riconoscimentoTrasporto = RiconosciAziendaTrasporto(transazione.Descrizione)
                If riconoscimentoTrasporto.PunteggioConfidenza > 0 Then
                    ' Imposta livello e colore confidenza per riconoscimento europeo
                    If riconoscimentoTrasporto.PunteggioConfidenza >= 0.9 Then
                        riconoscimentoTrasporto.LivelloConfidenza = "Molto Alta"
                        riconoscimentoTrasporto.ColoreConfidenza = "#2E7D32" ' Verde scuro per alta sicurezza
                    ElseIf riconoscimentoTrasporto.PunteggioConfidenza >= SogliaAlta Then
                        riconoscimentoTrasporto.LivelloConfidenza = "Alta"
                        riconoscimentoTrasporto.ColoreConfidenza = "#4CAF50" ' Verde
                    End If

                    riconoscimentoTrasporto.Motivazione = $"Azienda europea riconosciuta: {riconoscimentoTrasporto.PatternUsato}"

                    _logger?.LogInformation("Azienda trasporto riconosciuta: {Pattern} con confidenza {Confidenza:P}",
                                          riconoscimentoTrasporto.PatternUsato, riconoscimentoTrasporto.PunteggioConfidenza)
                    Return riconoscimentoTrasporto
                End If

                Using connection As New SQLiteConnection(GlobalDatabaseManager.GetConnectionString())
                    connection.Open()

                    ' Pattern table NON ha più Necessita, Frequenza, Stagionalita (sono solo in Transazioni)
                    Dim query As String = "
                        SELECT Parola, MacroCategoria, Categoria,
                               COALESCE(Peso, 5) as Peso
                        FROM Pattern
                        WHERE MacroCategoria IS NOT NULL AND MacroCategoria != ''
                          AND Categoria IS NOT NULL AND Categoria != ''
                        ORDER BY Peso DESC"

                    Using cmd As New SQLiteCommand(query, connection)
                        Using reader As SQLiteDataReader = cmd.ExecuteReader()
                            Dim migliorePunteggio As Double = 0.0

                            While reader.Read()
                                Dim pattern As String = reader("Parola").ToString()
                                Dim peso As Integer = Convert.ToInt32(reader("Peso"))
                                Dim macroCategoria As String = reader("MacroCategoria").ToString()
                                Dim categoria As String = reader("Categoria").ToString()

                                Dim punteggio As Double = CalcolaPunteggioMatch(transazione.Descrizione, pattern, peso)

                                If punteggio > migliorePunteggio AndAlso punteggio >= SogliaMatch Then
                                    migliorePunteggio = punteggio
                                    risultato.PatternUsato = pattern
                                    risultato.PunteggioConfidenza = punteggio
                                    risultato.MacroCategoria = macroCategoria
                                    risultato.Categoria = categoria
                                    ' Necessita/Frequenza/Stagionalita non sono più in Pattern, lasciali vuoti
                                    risultato.Necessita = String.Empty
                                    risultato.Frequenza = String.Empty
                                    risultato.Stagionalita = String.Empty
                                End If
                            End While
                        End Using
                    End Using
                End Using

                ' Imposta livello e colore confidenza
                If risultato.PunteggioConfidenza >= SogliaAlta Then
                    risultato.LivelloConfidenza = "Alta"
                    risultato.ColoreConfidenza = "#4CAF50" ' Verde
                ElseIf risultato.PunteggioConfidenza >= SogliaMedia Then
                    risultato.LivelloConfidenza = "Media"
                    risultato.ColoreConfidenza = "#FF9800" ' Arancione
                Else
                    risultato.LivelloConfidenza = "Bassa"
                    risultato.ColoreConfidenza = "#F44336" ' Rosso
                End If

                If risultato.PunteggioConfidenza > 0 Then
                    _logger?.LogInformation("Classificazione completata con confidenza {Confidenza:P}: {MacroCategoria} > {Categoria}",
                                          risultato.PunteggioConfidenza, risultato.MacroCategoria, risultato.Categoria)
                Else
                    _logger?.LogWarning("Nessun pattern trovato per la transazione: {Descrizione}", transazione.Descrizione)
                End If

            Catch ex As Exception
                _logger?.LogError(ex, "Errore durante la classificazione della transazione")
                risultato.PunteggioConfidenza = 0
            End Try

            Return risultato
        End Function

        ''' <summary>
        ''' Applica la classificazione alla transazione modificando le sue proprietà
        ''' </summary>
        Public Sub ApplicaClassificazione(transazione As Transazione, classificazione As RisultatoClassificazione)
            transazione.MacroCategoria = classificazione.MacroCategoria
            transazione.Categoria = classificazione.Categoria
            transazione.Necessita = classificazione.Necessita
            transazione.Frequenza = classificazione.Frequenza
            transazione.Stagionalita = classificazione.Stagionalita

            _logger?.LogInformation("Classificazione applicata alla transazione ID {ID}", transazione.ID)
        End Sub

        ''' <summary>
        ''' Classifica automaticamente tutte le transazioni non classificate
        ''' </summary>
        Public Function ClassificaTutteLeTransazioni() As Integer
            Dim transazioniClassificate As Integer = 0

            Try
                Using connection As New SQLiteConnection(GlobalDatabaseManager.GetConnectionString())
                    connection.Open()

                    ' Carica tutte le transazioni non classificate
                    Dim transazioniDaClassificare As New List(Of Transazione)
                    Dim queryTransazioni = "SELECT * FROM Transazioni WHERE MacroCategoria IS NULL OR MacroCategoria = '' OR MacroCategoria = 'Non Classificata'"

                    Using cmd As New SQLiteCommand(queryTransazioni, connection)
                        Using reader = cmd.ExecuteReader()
                            While reader.Read()
                                transazioniDaClassificare.Add(New Transazione With {
                                    .ID = Convert.ToInt32(reader("ID")),
                                    .Data = Convert.ToDateTime(reader("Data")),
                                    .Importo = Convert.ToDecimal(reader("Importo")),
                                    .Descrizione = reader("Descrizione").ToString(),
                                    .Causale = If(reader("Causale") IsNot DBNull.Value, reader("Causale").ToString(), ""),
                                    .MacroCategoria = If(reader("MacroCategoria") IsNot DBNull.Value, reader("MacroCategoria").ToString(), ""),
                                    .Categoria = If(reader("Categoria") IsNot DBNull.Value, reader("Categoria").ToString(), "")
                                })
                            End While
                        End Using
                    End Using

                    ' Classifica ogni transazione
                    For Each transazione In transazioniDaClassificare
                        Dim risultato = ClassificaTransazione(transazione)

                        ' Verifica se la classificazione è riuscita (ha MacroCategoria e punteggio > 0)
                        If Not String.IsNullOrEmpty(risultato.MacroCategoria) AndAlso risultato.PunteggioConfidenza > 0 Then
                            ' Aggiorna la transazione nel database
                            Dim updateQuery = "UPDATE Transazioni SET MacroCategoria = @macro, Categoria = @cat, Necessita = @nec, Frequenza = @freq, Stagionalita = @stag WHERE ID = @id"

                            Using updateCmd As New SQLiteCommand(updateQuery, connection)
                                updateCmd.Parameters.AddWithValue("@macro", risultato.MacroCategoria)
                                updateCmd.Parameters.AddWithValue("@cat", risultato.Categoria)
                                updateCmd.Parameters.AddWithValue("@nec", risultato.Necessita)
                                updateCmd.Parameters.AddWithValue("@freq", risultato.Frequenza)
                                updateCmd.Parameters.AddWithValue("@stag", risultato.Stagionalita)
                                updateCmd.Parameters.AddWithValue("@id", transazione.ID)

                                updateCmd.ExecuteNonQuery()
                                transazioniClassificate += 1
                            End Using
                        End If
                    Next
                End Using

            Catch ex As Exception
                _logger?.LogError(ex, "Errore durante la classificazione automatica delle transazioni")
            End Try

            Return transazioniClassificate
        End Function

        ''' <summary>
        ''' Test automatico delle correzioni AI implementate
        ''' </summary>
        Public Shared Function TestCorrenzioniAI() As String
            Dim report As New System.Text.StringBuilder()
            report.AppendLine("🧪 RAPPORTO TEST CORREZIONI AI")
            report.AppendLine("=" & New String("="c, 40))

            Dim testCases() As (descrizione As String, expectedMacro As String, expectedCat As String) = {
                ("OEBB PERSONENVERKEHR AG", "Trasporti", "Treni"),
                ("Q8 - VIA DEL MARINO", "Trasporti", "Carburante"),
                ("ASPIT - PEDAGGIO", "Trasporti", "Pedaggi Autostradali"),
                ("Blue Star Ferries", "Trasporti", "Traghetti"),
                ("AGIP STAZIONE SERVIZIO", "Trasporti", "Carburante"),
                ("SHELL CARBURANTE", "Trasporti", "Carburante")
            }

            Dim successi As Integer = 0
            Dim totali As Integer = testCases.Length

            For Each testCase In testCases
                Dim risultato = RiconosciAziendaTrasporto(testCase.descrizione)
                Dim successo = risultato.MacroCategoria = testCase.expectedMacro AndAlso
                              risultato.Categoria = testCase.expectedCat AndAlso
                              risultato.PunteggioConfidenza > 0.8

                If successo Then
                    successi += 1
                    report.AppendLine($"✅ {testCase.descrizione}")
                    report.AppendLine($"   {risultato.MacroCategoria} > {risultato.Categoria} ({risultato.PunteggioConfidenza:P})")
                Else
                    report.AppendLine($"❌ {testCase.descrizione}")
                    report.AppendLine($"   Atteso: {testCase.expectedMacro} > {testCase.expectedCat}")
                    report.AppendLine($"   Ottenuto: {risultato.MacroCategoria} > {risultato.Categoria} ({risultato.PunteggioConfidenza:P})")
                End If
                report.AppendLine()
            Next

            report.AppendLine($"📊 RISULTATI: {successi}/{totali} ({successi * 100 \ totali}%)")

            If successi = totali Then
                report.AppendLine("🎉 OTTIMO! Tutte le correzioni funzionano perfettamente!")
            ElseIf successi >= totali * 0.8 Then
                report.AppendLine("✅ Bene! La maggior parte delle correzioni funziona.")
            Else
                report.AppendLine("⚠️ Attenzione! Sono necessarie ulteriori correzioni.")
            End If

            Return report.ToString()
        End Function

        ''' <summary>
        ''' Sistema di esclusioni semantiche per evitare false positive
        ''' </summary>
        Private Shared Function ContieneEsclusioni(testoNormalizzato As String) As Boolean
            ' Lista di esclusioni per evitare classificazioni errate
            Dim esclusioni As String() = {
                "iperalfa", "ipercoop", "ipermercato", "iper", "principe", "filippo",
                "slip", "flip", "grip", "ship", "trip", "clip", "whip",
                "squad", "quite", "question", "queue",
                "autogrill", "bar ", "hotel", "ristorante", "pizzeria", "trattoria",
                "supermercato", "farmacia", "ospedale", "banca", "posta"
            }

            For Each esclusione In esclusioni
                If testoNormalizzato.Contains(esclusione) Then
                    Return True
                End If
            Next

            Return False
        End Function

        ''' <summary>
        ''' Sistema di riconoscimento semantico intelligente per migliorare la precisione
        ''' </summary>
        Public Shared Function ClassificazioneSemanticaAvanzata(descrizione As String) As RisultatoClassificazione
            Dim testoNorm = NormalizzaTesto(descrizione)
            Dim risultato As New RisultatoClassificazione()

            ' 🇮🇹 REGOLA 1: Riconoscimento Officine e Motori - TUTTO IN ITALIANO
            If testoNorm.Contains("motori") OrElse testoNorm.Contains("officina") OrElse
               testoNorm.Contains("autoricambi") OrElse testoNorm.Contains("car shop") Then
                risultato.MacroCategoria = "Trasporti"
                risultato.Categoria = "Manutenzione Auto"
                risultato.Necessita = "Occasionale"
                risultato.Frequenza = "Al Bisogno"
                risultato.Stagionalita = "Tutto l'Anno"
                risultato.PunteggioConfidenza = 0.95
                risultato.PatternUsato = "Officina/Autoricambi"
                risultato.ParolaChiave = EstrarreNomeAzienda(descrizione)
                Return risultato
            End If

            ' 🇮🇹 REGOLA 2: Riconoscimento Parcheggi - TUTTO IN ITALIANO
            If testoNorm.Contains("parcheggio") OrElse testoNorm.Contains("parking") Then
                risultato.MacroCategoria = "Trasporti"
                risultato.Categoria = "Parcheggi"
                risultato.Necessita = "Occasionale"
                risultato.Frequenza = "Variabile"
                risultato.Stagionalita = "Tutto l'Anno"
                risultato.PunteggioConfidenza = 0.95
                risultato.PatternUsato = "Rilevamento Parcheggio"
                risultato.ParolaChiave = EstrarreNomeAzienda(descrizione)
                Return risultato
            End If

            ' 🇮🇹 REGOLA 3: Riconoscimento Stazioni di Servizio - TUTTO IN ITALIANO
            If (testoNorm.Contains("stazione") AndAlso
                (testoNorm.Contains("ttp") OrElse testoNorm.Contains("agip") OrElse testoNorm.Contains("eni"))) Then
                risultato.MacroCategoria = "Trasporti"
                risultato.Categoria = "Carburante"
                risultato.Necessita = "Essenziale"
                risultato.Frequenza = "Frequente"
                risultato.Stagionalita = "Tutto l'Anno"
                risultato.PunteggioConfidenza = 0.9
                risultato.PatternUsato = "Contesto Stazione Servizio"
                risultato.ParolaChiave = EstrarreNomeAzienda(descrizione)
                Return risultato
            End If

            ' 🇮🇹 REGOLA 4: Riconoscimento Supermercati - TUTTO IN ITALIANO
            If testoNorm.Contains("conad") OrElse testoNorm.Contains("coop") OrElse
               testoNorm.Contains("iperalfa") OrElse testoNorm.Contains("ipercoop") OrElse
               testoNorm.Contains("esselunga") OrElse testoNorm.Contains("carrefour") OrElse
               testoNorm.Contains("ipermercato") OrElse testoNorm.Contains("supermarket") OrElse
               testoNorm.Contains("tigre") OrElse testoNorm.Contains("penny") OrElse
               testoNorm.Contains("lidl") OrElse testoNorm.Contains("eurospin") Then
                risultato.MacroCategoria = "Acquisti"
                risultato.Categoria = "Alimentari e Casa"
                risultato.Necessita = "Essenziale"
                risultato.Frequenza = "Frequente"
                risultato.Stagionalita = "Tutto l'Anno"
                risultato.PunteggioConfidenza = 0.95
                risultato.PatternUsato = "Catena Supermercati"
                risultato.ParolaChiave = EstrarreNomeAzienda(descrizione)
                Return risultato
            End If

            ' 🇮🇹 REGOLA 5: Riconoscimento Aziende Zootecniche/Agricole - TUTTO IN ITALIANO
            If testoNorm.Contains("zootecnica") OrElse testoNorm.Contains("agri") OrElse
               testoNorm.Contains("agricola") OrElse testoNorm.Contains("farm") Then
                risultato.MacroCategoria = "Acquisti"
                risultato.Categoria = "Prodotti Agricoli"
                risultato.Necessita = "Occasionale"
                risultato.Frequenza = "Stagionale"
                risultato.Stagionalita = "Variabile"
                risultato.PunteggioConfidenza = 0.9
                risultato.PatternUsato = "Business Agricolo"
                risultato.ParolaChiave = EstrarreNomeAzienda(descrizione)
                Return risultato
            End If

            Return risultato ' Vuoto se non trova nulla
        End Function

        ''' <summary>
        ''' Estrae il nome dell'azienda in modo intelligente
        ''' </summary>
        Private Shared Function EstrarreNomeAzienda(descrizione As String) As String
            ' Pattern per POS Eurozona
            Dim posMatch = System.Text.RegularExpressions.Regex.Match(descrizione,
                "Carta \*\d+\s+(.+?)\s+[A-Z]{3}$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            If posMatch.Success Then
                Return posMatch.Groups(1).Value.Trim()
            End If

            ' Pattern per PagoBancomat
            Dim bancoMatch = System.Text.RegularExpressions.Regex.Match(descrizione,
                "DEL \d{2}\.\d{2}\.\d{2} \d{2}:\d{2}\s+(.+?)\s+CARTA",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            If bancoMatch.Success Then
                Return bancoMatch.Groups(1).Value.Trim()
            End If

            ' Fallback: estrai parole significative
            Dim parole = descrizione.Split(" "c).Where(Function(p) p.Length > 2 AndAlso Not Char.IsDigit(p(0))).Take(3)
            Return String.Join(" ", parole)
        End Function
    End Class
End Namespace