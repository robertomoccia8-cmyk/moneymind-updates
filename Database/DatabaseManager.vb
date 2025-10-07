Imports System.Data.SQLite
Imports System.Data
Imports System.IO
Imports System.Reflection
Imports System.Windows
Imports System.Configuration

Public Class DatabaseManager
    Private Shared ConnectionString As String = ""
    Private Shared ReadOnly logLock As New Object()

    ' Helper per scrivere nel log in modo thread-safe con lock
    Private Shared Sub AppendToLog(logPath As String, message As String)
        Try
            SyncLock logLock
                Using writer As New StreamWriter(logPath, append:=True, System.Text.Encoding.UTF8)
                    writer.WriteLine(message)
                End Using
            End SyncLock
        Catch
            ' Ignora errori di log per non bloccare l'operazione principale
        End Try
    End Sub

    ' Inizializza il database e le tabelle
    Public Shared Sub InitializeDatabase()
        Try
            ' Verifica protezione password prima di continuare
            If PasswordService.IsPasswordProtected() Then
                If Not PasswordService.ShowPasswordDialog() Then
                    Throw New UnauthorizedAccessException("Password non corretta o accesso negato")
                End If
            End If

            ' Inizializza variabili comuni usate sia nel path setup che nel retry loop
            Dim logPath = Path.Combine(Path.GetTempPath(), "MoneyMind_DatabasePath.log")
            Dim databasePath As String = Nothing

            ' Se il ConnectionString è già stato impostato (es. da SetCurrentConto), non sovrascriverlo
            If Not String.IsNullOrEmpty(ConnectionString) Then
                'Debug.WriteLine("DatabaseManager: ConnectionString già impostato, skip GetDatabasePath")
                ' Estrai il database path dal ConnectionString per logging
                Try
                    Dim connBuilder As New System.Data.SQLite.SQLiteConnectionStringBuilder(ConnectionString)
                    databasePath = connBuilder.DataSource
                Catch
                    databasePath = "Unknown"
                End Try
                AppendToLog(logPath, $"ConnectionString già impostato, path: '{databasePath}'{Environment.NewLine}")
                ' Usa il ConnectionString esistente senza modificarlo
                GoTo SkipPathSetup
            End If

            ' Ottieni percorso database dalle impostazioni o usa default
            databasePath = GetDatabasePath()

            ' Log debug per vedere cosa succede
            AppendToLog(logPath, $"DatabasePath ottenuto: '{databasePath}'{Environment.NewLine}")

            ' Verifica che il percorso non sia null o vuoto
            If String.IsNullOrEmpty(databasePath) Then
                Throw New Exception("Percorso database non valido (null o vuoto)")
            End If

            Dim dataFolder As String = Path.GetDirectoryName(databasePath)
            AppendToLog(logPath, $"DataFolder calcolato: '{dataFolder}'{Environment.NewLine}")

            ' Verifica che dataFolder non sia null
            If String.IsNullOrEmpty(dataFolder) Then
                ' Fallback: usa la directory dell'exe
                dataFolder = AppDomain.CurrentDomain.BaseDirectory
                databasePath = Path.Combine(dataFolder, "MoneyMind.db")
                AppendToLog(logPath, $"DataFolder era null, fallback a: '{dataFolder}'{Environment.NewLine}")
                AppendToLog(logPath, $"DatabasePath aggiornato a: '{databasePath}'{Environment.NewLine}")
            End If

            If Not Directory.Exists(dataFolder) Then
                Directory.CreateDirectory(dataFolder)
                AppendToLog(logPath, $"Cartella creata: '{dataFolder}'{Environment.NewLine}")
            End If

            ' Normalizza il percorso per evitare problemi con SQLite
            databasePath = Path.GetFullPath(databasePath)
            AppendToLog(logPath, $"DatabasePath normalizzato: '{databasePath}'{Environment.NewLine}")

            ' Connection string con parametri per gestire meglio la concorrenza
            ' Busy Timeout=5000 per aspettare 5 secondi se il DB è locked
            ' Pooling=true per riutilizzare le connessioni
            ' NOTA: Non impostiamo Journal Mode=WAL qui per evitare lock, lo faremo dopo nella connessione
            ConnectionString = $"Data Source={databasePath};Version=3;Busy Timeout=5000;Pooling=true;Max Pool Size=100;"
            AppendToLog(logPath, $"ConnectionString con Busy Timeout: '{ConnectionString}'{Environment.NewLine}")

            ' Test del path prima di creare la connessione
            Try
                Dim testDir = Path.GetDirectoryName(databasePath)
                If Not Directory.Exists(testDir) Then
                    Directory.CreateDirectory(testDir)
                    AppendToLog(logPath, $"Directory creata: '{testDir}'{Environment.NewLine}")
                End If

                ' Test scrittura nel percorso
                Dim testFile = Path.Combine(testDir, "test_write.tmp")
                File.WriteAllText(testFile, "test")
                File.Delete(testFile)
                AppendToLog(logPath, $"Test scrittura: OK{Environment.NewLine}")

            Catch ex As Exception
                AppendToLog(logPath, $"ERRORE test percorso: {ex.Message}{Environment.NewLine}")
                ' Fallback a path diverso se non può scrivere
                Dim fallbackPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                If String.IsNullOrEmpty(fallbackPath) Then
                    fallbackPath = Path.Combine(Environment.CurrentDirectory, "Data")
                End If
                databasePath = Path.Combine(fallbackPath, "MoneyMind", "MoneyMind.db")
                Directory.CreateDirectory(Path.GetDirectoryName(databasePath))
                ConnectionString = $"Data Source={databasePath};Version=3;Busy Timeout=5000;Pooling=true;Max Pool Size=100;"
                AppendToLog(logPath, $"Fallback path: '{databasePath}'{Environment.NewLine}")
                AppendToLog(logPath, $"Fallback ConnectionString: '{ConnectionString}'{Environment.NewLine}")
            End Try

            AppendToLog(logPath, $"Tentativo connessione SQLite...{Environment.NewLine}")

SkipPathSetup:
            ' STRATEGIA MULTI-RETRY per evitare bug SQLite interno
            Dim connection As SQLiteConnection = Nothing
            Dim connectionSuccess As Boolean = False

            For attempt As Integer = 1 To 3
                Try
                    AppendToLog(logPath, $"Tentativo #{attempt} connessione...{Environment.NewLine}")
                    connection = New SQLiteConnection(ConnectionString)
                    connection.Open()
                    AppendToLog(logPath, $"SUCCESSO tentativo #{attempt}!{Environment.NewLine}")
                    connectionSuccess = True
                    Exit For

                Catch ex As Exception When ex.Message.Contains("path1")
                    AppendToLog(logPath, $"Tentativo #{attempt} fallito con errore path1: {ex.Message}{Environment.NewLine}")

                    ' Prova connection string ancora più semplice
                    If attempt = 2 Then
                        ConnectionString = $"Data Source={databasePath};Busy Timeout=5000"
                        AppendToLog(logPath, $"Retry con ConnectionString BASE: '{ConnectionString}'{Environment.NewLine}")
                    End If

                    ' Ultimo tentativo con path diverso
                    If attempt = 3 Then
                        Dim appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
                        If String.IsNullOrEmpty(appDataPath) Then
                            appDataPath = Path.Combine(Environment.CurrentDirectory, "Data")
                        End If
                        databasePath = Path.Combine(appDataPath, "MoneyMind", "MoneyMind.db")
                        Directory.CreateDirectory(Path.GetDirectoryName(databasePath))
                        ConnectionString = $"Data Source={databasePath};Busy Timeout=5000"
                        AppendToLog(logPath, $"Ultimo tentativo con path AppData: '{ConnectionString}'{Environment.NewLine}")
                    End If

                Catch ex As Exception
                    AppendToLog(logPath, $"Tentativo #{attempt} fallito con altro errore: {ex.Message}{Environment.NewLine}")
                    Throw ' Re-lancia altri errori
                End Try
            Next

            If Not connectionSuccess Then
                Throw New Exception("Impossibile stabilire connessione SQLite dopo 3 tentativi")
            End If

            Using connection
                ' Connessione già aperta nel retry loop

                ' Imposta WAL mode per migliorare la concorrenza (solo se non è già impostato)
                Try
                    Using cmdWal As New SQLiteCommand("PRAGMA journal_mode=WAL", connection)
                        cmdWal.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    ' Ignora errori se WAL è già impostato o se il DB è locked
                    AppendToLog(logPath, $"WAL mode già impostato o non disponibile: {ex.Message}{Environment.NewLine}")
                End Try

                ' Tabella Transazioni
                Dim createTransQuery As String = "
                CREATE TABLE IF NOT EXISTS Transazioni (
                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                    Data DATE NOT NULL,
                    Importo DECIMAL(10,2) NOT NULL,
                    Descrizione TEXT NOT NULL,
                    Causale TEXT DEFAULT '',
                    MacroCategoria TEXT DEFAULT '',
                    Categoria TEXT DEFAULT '',
                    Necessita TEXT DEFAULT '',
                    Frequenza TEXT DEFAULT '',
                    Stagionalita TEXT DEFAULT '',
                    DataInserimento DATETIME DEFAULT CURRENT_TIMESTAMP,
                    DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
                );"

                Using cmd As New SQLiteCommand(createTransQuery, connection)
                    cmd.ExecuteNonQuery()
                End Using

                ' Tabella Pattern semplificata
                Dim createPatternQuery As String = "
                CREATE TABLE IF NOT EXISTS Pattern (
                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                    Parola TEXT NOT NULL UNIQUE,
                    MacroCategoria TEXT NOT NULL DEFAULT '',
                    Categoria TEXT NOT NULL DEFAULT '',
                    Necessita TEXT NOT NULL DEFAULT '',
                    Frequenza TEXT NOT NULL DEFAULT '',
                    Stagionalita TEXT NOT NULL DEFAULT ''
                );"

                Using cmd As New SQLiteCommand(createPatternQuery, connection)
                    cmd.ExecuteNonQuery()
                End Using

                ' Tabella MacroCategorie per gestire MacroCategorie vuote
                Dim createMacroCategorieQuery As String = "
                CREATE TABLE IF NOT EXISTS MacroCategorie (
                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                    Nome TEXT NOT NULL UNIQUE,
                    DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP
                );"

                Using cmd As New SQLiteCommand(createMacroCategorieQuery, connection)
                    cmd.ExecuteNonQuery()
                End Using

                ' Tabella Budget per il tracking dei budget per categoria
                Dim createBudgetQuery As String = "
                CREATE TABLE IF NOT EXISTS Budget (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Categoria TEXT NOT NULL,
                    MacroCategoria TEXT,
                    BudgetMensile DECIMAL(10,2) NOT NULL,
                    Mese INTEGER NOT NULL,
                    Anno INTEGER NOT NULL,
                    DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                    DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(Categoria, Anno, Mese)
                );"

                Using cmd As New SQLiteCommand(createBudgetQuery, connection)
                    cmd.ExecuteNonQuery()
                End Using

                ' Tabella Obiettivi per obiettivi di risparmio
                Dim createObiettiviQuery As String = "
                CREATE TABLE IF NOT EXISTS Obiettivi (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Nome TEXT NOT NULL,
                    Descrizione TEXT,
                    ImportoTarget DECIMAL(10,2) NOT NULL,
                    ImportoCorrente DECIMAL(10,2) DEFAULT 0,
                    DataInizio DATETIME NOT NULL,
                    DataScadenza DATETIME,
                    Completato BOOLEAN DEFAULT 0,
                    Colore TEXT DEFAULT '#4CAF50',
                    DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                    DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
                );"

                Using cmd As New SQLiteCommand(createObiettiviQuery, connection)
                    cmd.ExecuteNonQuery()
                End Using

                ' Aggiorna struttura Pattern: aggiunge colonne Fonte e Peso se mancanti
                AggiornaStrutturaPattern(connection)

                ' Esempi iniziali (solo se vuota)
                InserisciPatternEsempio(connection)
                
                ' *** OTTIMIZZAZIONI PERFORMANCE - Crea indici per query più veloci ***
                CreaIndiciPerformance(connection)
                
                ' *** INIZIALIZZA TABELLA API KEYS ***
                InizializzaTabellaApiKeys(connection)

                ' *** INIZIALIZZA TABELLA IMPOSTAZIONI CONTO CORRENTE ***
                InizializzaTabellaImpostazioniConto(connection)
            End Using

            ' *** MIGRAZIONE AUTOMATICA - Chiamata dopo l'inizializzazione base ***
            MigrazioneConfigurazioneStipendi()
            
            ' *** MIGRAZIONE DATI - Importa dati dal database completo se esiste ***
            MigraDatiDaVecchioDatabase()

        Catch ex As Exception
            ' Log dettagliato per debug
            Try
                Dim logPath = Path.Combine(Path.GetTempPath(), "MoneyMind_DatabaseError.log")
                AppendToLog(logPath, $"{DateTime.Now}: {ex.Message}{Environment.NewLine}{ex.StackTrace}{Environment.NewLine}{Environment.NewLine}")
            Catch
                ' Ignora errori di logging
            End Try
            MessageBox.Show("Errore inizializzazione database: " & ex.Message & Environment.NewLine & Environment.NewLine & "Dettagli salvati in: " & Path.Combine(Path.GetTempPath(), "MoneyMind_DatabaseError.log"), "Errore", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' *** MIGRAZIONE AUTOMATICA - Controlla e crea le tabelle stipendi se non esistono ***
    Public Shared Sub MigrazioneConfigurazioneStipendi()
        Try
            Using conn As New SQLiteConnection(GetConnectionString())
                conn.Open()

                ' Controlla se le tabelle stipendi esistono
                Dim sqlCheck As String = "SELECT name FROM sqlite_master WHERE type='table' AND name='ConfigurazioneStipendi'"
                Using cmd As New SQLiteCommand(sqlCheck, conn)
                    Dim exists = cmd.ExecuteScalar()

                    If exists Is Nothing Then
                        ' Prima volta - crea tabelle e configurazione di default
                        CreaTabellConfigurazioneStipendi()

                        ' Messaggio informativo per utenti esistenti
                        MessageBox.Show("Benvenuto nel nuovo sistema di configurazione stipendi!" & vbCrLf &
                                       "È stata creata una configurazione predefinita:" & vbCrLf &
                                       "- Giorno stipendio: 23" & vbCrLf &
                                       "- Weekend: anticipa al venerdì" & vbCrLf & vbCrLf &
                                       "Puoi personalizzarla dal menu Impostazioni.",
                                       "Aggiornamento Sistema", MessageBoxButton.OK, MessageBoxImage.Information)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore durante la migrazione del sistema stipendi: " & ex.Message, "Errore Migrazione", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' Ottieni percorso database dalle impostazioni o usa default
    Private Shared Function GetDatabasePath() As String
        Dim logPath = Path.Combine(Path.GetTempPath(), "MoneyMind_DatabasePath.log")

        Try
            AppendToLog(logPath, $"{DateTime.Now}: Inizio GetDatabasePath{Environment.NewLine}")

            ' Prova a leggere dalle impostazioni
            Dim configPath As String = ConfigurationManager.AppSettings("DatabasePath")
            AppendToLog(logPath, $"ConfigPath da AppSettings: '{configPath}'{Environment.NewLine}")

            If Not String.IsNullOrEmpty(configPath) Then
                ' Se è un percorso relativo, rendilo assoluto
                If Not Path.IsPathRooted(configPath) Then
                    Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
                    AppendToLog(logPath, $"BaseDirectory: '{baseDir}'{Environment.NewLine}")
                    configPath = Path.Combine(baseDir, configPath)
                End If
                AppendToLog(logPath, $"ConfigPath finale: '{configPath}'{Environment.NewLine}")
                Return configPath
            End If
        Catch ex As Exception
            AppendToLog(logPath, $"ERRORE lettura configurazione: {ex.Message}{Environment.NewLine}")
            Debug.WriteLine($"Errore lettura configurazione database: {ex.Message}")
        End Try

        ' Database nella cartella dell'applicazione (come app portable)
        Try
            ' Usa la cartella dell'eseguibile corrente
            Dim appDir As String = AppDomain.CurrentDomain.BaseDirectory
            Dim dataDir As String = Path.Combine(appDir, "Data")
            Dim defaultPath As String = Path.Combine(dataDir, "MoneyMind.db")

            AppendToLog(logPath, $"APP DIR: '{appDir}'{Environment.NewLine}")
            AppendToLog(logPath, $"DATA DIR: '{dataDir}'{Environment.NewLine}")
            AppendToLog(logPath, $"DATABASE PATH: '{defaultPath}'{Environment.NewLine}")

            ' Crea directory se non esiste
            Directory.CreateDirectory(dataDir)

            Return defaultPath
        Catch ex As Exception
            ' Ultimate fallback per situazioni estreme
            AppendToLog(logPath, $"ERRORE fallback: {ex.Message}{Environment.NewLine}")
            Debug.WriteLine($"Errore creazione percorso database: {ex.Message}")
            Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            If String.IsNullOrEmpty(localAppData) Then
                localAppData = Path.Combine(Environment.CurrentDirectory, "Data")
            End If
            Dim ultimatePath = Path.Combine(localAppData, "MoneyMind", "Data", "MoneyMind.db")
            AppendToLog(logPath, $"Ultimate fallback: '{ultimatePath}'{Environment.NewLine}")
            Return ultimatePath
        End Try
    End Function

    ' Migra dati dal database completo esistente
    Public Shared Sub MigraDatiDaVecchioDatabase()
        Try
            ' Percorso del database completo esistente
            Dim oldDbPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "MoneyMind_completo.db")
            
            If Not File.Exists(oldDbPath) Then
                ' Debug.WriteLine("Database completo non trovato, migrazione non necessaria")
                Return
            End If
            
            ' Verifica se il nuovo database è vuoto
            Using newConn As New SQLiteConnection(GetConnectionString())
                newConn.Open()
                Using cmd As New SQLiteCommand("SELECT COUNT(*) FROM Transazioni", newConn)
                    Dim count As Integer = CInt(cmd.ExecuteScalar())
                    If count > 0 Then
                        Debug.WriteLine("Il nuovo database contiene già dati, migrazione saltata")
                        Return
                    End If
                End Using
            End Using
            
            ' Connessione al vecchio database
            Dim oldConnectionString As String = $"Data Source={oldDbPath};Version=3;"
            
            Using oldConn As New SQLiteConnection(oldConnectionString)
            Using newConn As New SQLiteConnection(GetConnectionString())
                oldConn.Open()
                newConn.Open()
                
                Using trans = newConn.BeginTransaction()
                    ' Migra Transazioni
                    Dim sqlSelectTrans As String = "SELECT Data, Importo, Descrizione, Causale, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita FROM Transazioni ORDER BY Data"
                    Using cmdSelect As New SQLiteCommand(sqlSelectTrans, oldConn)
                    Using reader = cmdSelect.ExecuteReader()
                        While reader.Read()
                            Dim sqlInsert As String = "INSERT INTO Transazioni (Data, Importo, Descrizione, Causale, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita) VALUES (@data, @importo, @desc, @causale, @macro, @cat, @nec, @freq, @stag)"
                            Using cmdInsert As New SQLiteCommand(sqlInsert, newConn, trans)
                                cmdInsert.Parameters.AddWithValue("@data", reader("Data"))
                                cmdInsert.Parameters.AddWithValue("@importo", reader("Importo"))
                                cmdInsert.Parameters.AddWithValue("@desc", reader("Descrizione"))
                                cmdInsert.Parameters.AddWithValue("@causale", If(reader("Causale"), ""))
                                cmdInsert.Parameters.AddWithValue("@macro", If(reader("MacroCategoria"), ""))
                                cmdInsert.Parameters.AddWithValue("@cat", If(reader("Categoria"), ""))
                                cmdInsert.Parameters.AddWithValue("@nec", If(reader("Necessita"), ""))
                                cmdInsert.Parameters.AddWithValue("@freq", If(reader("Frequenza"), ""))
                                cmdInsert.Parameters.AddWithValue("@stag", If(reader("Stagionalita"), ""))
                                cmdInsert.ExecuteNonQuery()
                            End Using
                        End While
                    End Using
                    End Using
                    
                    ' Migra Pattern (se esistono nel vecchio database)
                    Try
                        Dim sqlSelectPattern As String = "SELECT Parola, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita FROM Pattern"
                        Using cmdSelect As New SQLiteCommand(sqlSelectPattern, oldConn)
                        Using reader = cmdSelect.ExecuteReader()
                            While reader.Read()
                                Dim sqlInsert As String = "INSERT OR REPLACE INTO Pattern (Parola, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita, Fonte, Peso) VALUES (@parola, @macro, @cat, @nec, @freq, @stag, 'Migrazione', 8)"
                                Using cmdInsert As New SQLiteCommand(sqlInsert, newConn, trans)
                                    cmdInsert.Parameters.AddWithValue("@parola", reader("Parola"))
                                    cmdInsert.Parameters.AddWithValue("@macro", If(reader("MacroCategoria"), ""))
                                    cmdInsert.Parameters.AddWithValue("@cat", If(reader("Categoria"), ""))
                                    cmdInsert.Parameters.AddWithValue("@nec", If(reader("Necessita"), ""))
                                    cmdInsert.Parameters.AddWithValue("@freq", If(reader("Frequenza"), ""))
                                    cmdInsert.Parameters.AddWithValue("@stag", If(reader("Stagionalita"), ""))
                                    cmdInsert.ExecuteNonQuery()
                                End Using
                            End While
                        End Using
                        End Using
                    Catch ex As Exception
                        Debug.WriteLine($"Tabella Pattern non trovata nel vecchio database: {ex.Message}")
                    End Try
                    
                    trans.Commit()
                End Using
            End Using
            End Using
            
            ' Backup del vecchio database
            Dim backupPath As String = oldDbPath.Replace(".db", "_backup.db")
            If Not File.Exists(backupPath) Then
                File.Copy(oldDbPath, backupPath)
            End If
            
            MessageBox.Show($"Migrazione completata con successo!" & vbCrLf & 
                          $"Backup creato: {backupPath}", 
                          "Migrazione Database", MessageBoxButton.OK, MessageBoxImage.Information)
                          
        Catch ex As Exception
            MessageBox.Show($"Errore durante la migrazione: {ex.Message}", "Errore Migrazione", MessageBoxButton.OK, MessageBoxImage.Error)
            Debug.WriteLine($"Errore migrazione database: {ex.Message}")
        End Try
    End Sub

    ' Crea le tabelle per la configurazione stipendi
    Public Shared Sub CreaTabellConfigurazioneStipendi()
        Using conn As New SQLiteConnection(GetConnectionString())
            conn.Open()

            ' Tabella configurazione principale
            Dim sqlConfig As String = "
            CREATE TABLE IF NOT EXISTS ConfigurazioneStipendi (
                Id INTEGER PRIMARY KEY,
                GiornoDefault INTEGER NOT NULL DEFAULT 23,
                RegoleWeekend TEXT NOT NULL DEFAULT 'ANTICIPA',
                DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
            )"

            ' Tabella eccezioni mensili
            Dim sqlEccezioni As String = "
            CREATE TABLE IF NOT EXISTS EccezioniStipendi (
                Id INTEGER PRIMARY KEY,
                Mese INTEGER NOT NULL,
                GiornoSpeciale INTEGER NOT NULL,
                Descrizione TEXT,
                Attivo INTEGER DEFAULT 1,
                UNIQUE(Mese)
            )"

            Using cmd As New SQLiteCommand(sqlConfig, conn)
                cmd.ExecuteNonQuery()
            End Using

            Using cmd As New SQLiteCommand(sqlEccezioni, conn)
                cmd.ExecuteNonQuery()
            End Using

            ' Inserisci configurazione predefinita se non esiste
            Dim sqlCheck As String = "SELECT COUNT(*) FROM ConfigurazioneStipendi"
            Using cmd As New SQLiteCommand(sqlCheck, conn)
                If Convert.ToInt32(cmd.ExecuteScalar()) = 0 Then
                    Dim sqlInsert As String = "INSERT INTO ConfigurazioneStipendi (GiornoDefault, RegoleWeekend) VALUES (23, 'ANTICIPA')"
                    Using cmdInsert As New SQLiteCommand(sqlInsert, conn)
                        cmdInsert.ExecuteNonQuery()
                    End Using

                End If
            End Using
        End Using
    End Sub

    ' Test di validazione configurazione stipendi
    ' TODO: Implementare in WPF quando il modulo stipendi sarà disponibile
    Public Shared Sub ValidaConfigurazioneStipendi()
        Try
            ' Configurazione stipendi non ancora implementata in WPF
            ' Dim config = GestoreStipendi.CaricaConfigurazione()
            ' Dim annoTest = Date.Today.Year

            ' For mese As Integer = 1 To 12
            '     Dim payDate = GestoreStipendi.CalcolaPayDate(annoTest, mese)
            '     Dim periodo = GestoreStipendi.CalcolaPeriodoStipendiale(annoTest, mese)
            
            '     ' Verifica che non ci siano sovrapposizioni
            '     If mese < 12 Then
            '         Dim periodoSuccessivo = GestoreStipendi.CalcolaPeriodoStipendiale(annoTest, mese + 1)
            '         If periodo.DataFine >= periodoSuccessivo.DataInizio Then
            '             Throw New Exception($"Sovrapposizione rilevata tra {mese} e {mese + 1}")
            '         End If
            '     End If
            ' Next

            Debug.WriteLine("Configurazione stipendi validata con successo")

        Catch ex As Exception
            MessageBox.Show($"Errore nella configurazione stipendi: {ex.Message}", "Errore Configurazione", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    '
    'Rimuove la colonna SottoCategoria dallo schema del database ***
    '
    Public Shared Sub RimuoviSottoCategoriaSchema()
        Try
            Using conn As New SQLiteConnection(GetConnectionString())
                conn.Open()
                Using trans = conn.BeginTransaction()

                    ' 1. Crea nuove tabelle senza SottoCategoria
                    Dim sqlNuoveTabelle As String = "
                    CREATE TABLE IF NOT EXISTS Pattern_Temp (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        Parola TEXT NOT NULL,
                        MacroCategoria TEXT NOT NULL,
                        Categoria TEXT NOT NULL,
                        Necessita TEXT,
                        Frequenza TEXT,
                        Stagionalita TEXT,
                        Fonte TEXT DEFAULT 'Manuale',
                        Peso INTEGER DEFAULT 5
                    );
                    
                    CREATE TABLE IF NOT EXISTS Transazioni_Temp (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        Data DATE NOT NULL,
                        Importo DECIMAL(10,2) NOT NULL,
                        Descrizione TEXT NOT NULL,
                        Causale TEXT,
                        MacroCategoria TEXT,
                        Categoria TEXT,
                        Necessita TEXT,
                        Frequenza TEXT,
                        Stagionalita TEXT
                    );"

                    Using cmd As New SQLiteCommand(sqlNuoveTabelle, conn, trans)
                        cmd.ExecuteNonQuery()
                    End Using

                    ' 2. Copia dati dalle tabelle originali (senza SottoCategoria)
                    Dim sqlCopiaPattern As String = "
                    INSERT INTO Pattern_Temp (ID, Parola, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita, Fonte, Peso)
                    SELECT ID, Parola, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita, Fonte, Peso 
                    FROM Pattern"

                    Using cmd As New SQLiteCommand(sqlCopiaPattern, conn, trans)
                        cmd.ExecuteNonQuery()
                    End Using

                    Dim sqlCopiaTransazioni As String = "
                    INSERT INTO Transazioni_Temp (ID, Data, Importo, Descrizione, Causale, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita)
                    SELECT ID, Data, Importo, Descrizione, Causale, MacroCategoria, Categoria, Necessita, Frequenza, Stagionalita 
                    FROM Transazioni"

                    Using cmd As New SQLiteCommand(sqlCopiaTransazioni, conn, trans)
                        cmd.ExecuteNonQuery()
                    End Using

                    ' 3. Elimina tabelle originali
                    Using cmd As New SQLiteCommand("DROP TABLE Pattern", conn, trans)
                        cmd.ExecuteNonQuery()
                    End Using
                    Using cmd As New SQLiteCommand("DROP TABLE Transazioni", conn, trans)
                        cmd.ExecuteNonQuery()
                    End Using

                    ' 4. Rinomina tabelle temporanee
                    Using cmd As New SQLiteCommand("ALTER TABLE Pattern_Temp RENAME TO Pattern", conn, trans)
                        cmd.ExecuteNonQuery()
                    End Using
                    Using cmd As New SQLiteCommand("ALTER TABLE Transazioni_Temp RENAME TO Transazioni", conn, trans)
                        cmd.ExecuteNonQuery()
                    End Using

                    trans.Commit()
                End Using
            End Using

            'MessageBox.Show("Schema database aggiornato: SottoCategoria rimossa con successo!", "Migrazione Completata", MessageBoxButton.OK, MessageBoxImage.Information)

        Catch ex As Exception
            MessageBox.Show($"Errore durante la migrazione database: {ex.Message}", "Errore Migrazione", MessageBoxButton.OK, MessageBoxImage.Error)
            Throw
        End Try
    End Sub

    ' Verifica e aggiunge colonne mancanti nella tabella Pattern
    Private Shared Sub AggiornaStrutturaPattern(connection As SQLiteConnection)
        Dim cols As New List(Of String)
        Using cmd As New SQLiteCommand("PRAGMA table_info(Pattern)", connection)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    cols.Add(reader("name").ToString().ToLower())
                End While
            End Using
        End Using

        If Not cols.Contains("fonte") Then
            Using cmd As New SQLiteCommand("ALTER TABLE Pattern ADD COLUMN Fonte TEXT DEFAULT ''", connection)
                cmd.ExecuteNonQuery()
            End Using
        End If

        If Not cols.Contains("peso") Then
            Using cmd As New SQLiteCommand("ALTER TABLE Pattern ADD COLUMN Peso INTEGER DEFAULT 5", connection)
                cmd.ExecuteNonQuery()
            End Using
        End If
        
        If Not cols.Contains("utilizzi") Then
            Using cmd As New SQLiteCommand("ALTER TABLE Pattern ADD COLUMN Utilizzi INTEGER DEFAULT 0", connection)
                cmd.ExecuteNonQuery()
            End Using
        End If
        
        If Not cols.Contains("ultimouso") Then
            Using cmd As New SQLiteCommand("ALTER TABLE Pattern ADD COLUMN UltimoUso DATETIME", connection)
                cmd.ExecuteNonQuery()
            End Using
        End If
    End Sub

    ' Inserisce pattern di esempio se la tabella è vuota
    Private Shared Sub InserisciPatternEsempio(connection As SQLiteConnection)
        Using countCmd As New SQLiteCommand("SELECT COUNT(*) FROM Pattern", connection)
            Dim cnt As Integer = CInt(countCmd.ExecuteScalar())
            If cnt > 0 Then Return
        End Using

        Dim patternsBase As New List(Of (String, String, String, String, String, String, Integer)) From {
            ("Supermercato", "Alimentari", "Spesa quotidiana", "SUPERMERCATO", "Essenziale", "Ricorrente", 10),
            ("Farmacia", "Salute", "Medicinali", "FARMACIA", "Essenziale", "Occasionale", 10),
            ("Benzina", "Trasporti", "Carburante", "BENZINA", "Essenziale", "Ricorrente", 10),
            ("Amazon", "Shopping", "E-commerce", "AMAZON", "Utile", "Occasionale", 9)
        }

        For Each p In patternsBase
            Dim insertQuery As String = "
            INSERT OR IGNORE INTO Pattern
              (MacroCategoria, Categoria, Parola, Necessita, Frequenza, Stagionalita, Peso, Fonte)
            VALUES
              (@macro, @cat, @parola, @nec, @freq, @stag, @peso, '');"

            Using cmd As New SQLiteCommand(insertQuery, connection)
                cmd.Parameters.AddWithValue("@macro", p.Item1)
                cmd.Parameters.AddWithValue("@cat", p.Item2)
                'cmd.Parameters.AddWithValue("@sotto", p.Item3)
                cmd.Parameters.AddWithValue("@parola", p.Item4)
                cmd.Parameters.AddWithValue("@nec", p.Item5)
                cmd.Parameters.AddWithValue("@freq", p.Item6)
                cmd.Parameters.AddWithValue("@stag", p.Item5) ' usa Necessita per Stagionalita
                cmd.Parameters.AddWithValue("@peso", p.Item7)
                cmd.ExecuteNonQuery()
            End Using
        Next
    End Sub

    ' *** OTTIMIZZAZIONI PERFORMANCE ***
    ' Crea indici critici per migliorare le performance delle query più frequenti
    Private Shared Sub CreaIndiciPerformance(connection As SQLiteConnection)
        Try
            Dim indici As String() = {
                "CREATE INDEX IF NOT EXISTS idx_transazioni_data ON Transazioni(Data DESC)",
                "CREATE INDEX IF NOT EXISTS idx_transazioni_categoria ON Transazioni(MacroCategoria, Categoria)",
                "CREATE INDEX IF NOT EXISTS idx_transazioni_importo ON Transazioni(Importo)",
                "CREATE INDEX IF NOT EXISTS idx_transazioni_descrizione ON Transazioni(Descrizione)",
                "CREATE INDEX IF NOT EXISTS idx_transazioni_data_categoria ON Transazioni(Data, MacroCategoria)",
                "CREATE INDEX IF NOT EXISTS idx_transazioni_anno_mese ON Transazioni(strftime('%Y-%m', Data))",
                "CREATE INDEX IF NOT EXISTS idx_pattern_parola ON Pattern(Parola)",
                "CREATE INDEX IF NOT EXISTS idx_pattern_categoria ON Pattern(MacroCategoria, Categoria)",
                "CREATE INDEX IF NOT EXISTS idx_pattern_peso ON Pattern(Peso DESC)",
                "CREATE INDEX IF NOT EXISTS idx_pattern_utilizzi ON Pattern(UltimoUso DESC)",
                "CREATE INDEX IF NOT EXISTS idx_pattern_macro_only ON Pattern(MacroCategoria)",
                "CREATE INDEX IF NOT EXISTS idx_transazioni_descrizione_lower ON Transazioni(LOWER(Descrizione))",
                "CREATE INDEX IF NOT EXISTS idx_pattern_parola_lower ON Pattern(LOWER(Parola))",
                "CREATE INDEX IF NOT EXISTS idx_macrocategorie_nome ON MacroCategorie(Nome COLLATE NOCASE)",
                "CREATE INDEX IF NOT EXISTS idx_transazioni_classificate ON Transazioni(MacroCategoria) WHERE MacroCategoria != '' AND MacroCategoria IS NOT NULL"
            }

            For Each indice In indici
                Using cmd As New SQLiteCommand(indice, connection)
                    cmd.ExecuteNonQuery()
                End Using
            Next

            ' Crea trigger per integrità dati
            CreaTriggerIntegrita(connection)

            ' Ottimizza il database dopo la creazione degli indici
            Using cmd As New SQLiteCommand("ANALYZE", connection)
                cmd.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            ' Log dell'errore ma non interrompe l'inizializzazione
            Debug.WriteLine($"Avviso: errore creazione indici performance: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Crea trigger per mantenere integrità dati
    ''' </summary>
    Private Shared Sub CreaTriggerIntegrita(connection As SQLiteConnection)
        Try
            ' Per ora disabilitiamo i trigger per evitare problemi di sintassi
            ' Saranno implementati in futuro quando sarà necessario
            ' Debug.WriteLine("Trigger integrità temporaneamente disabilitati")
            
            ' TODO: Implementare trigger quando sarà necessario per integrità dati
            ' I trigger avrebbero dovuto:
            ' 1. Auto-inserire MacroCategorie quando si creano Pattern
            ' 2. Aggiornare timestamp UltimoUso dei Pattern quando usati
            ' 3. Aggiornare timestamp DataModifica delle Transazioni

        Catch ex As Exception
            Debug.WriteLine($"Avviso: errore creazione trigger integrità: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Inizializza la tabella per le chiavi API
    ''' </summary>
    Private Shared Sub InizializzaTabellaApiKeys(connection As SQLiteConnection)
        Try
            Dim createApiKeysQuery = "
                CREATE TABLE IF NOT EXISTS ApiKeys (
                    Nome TEXT PRIMARY KEY,
                    ChiaveCriptata TEXT NOT NULL,
                    DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                    DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
                )"
            
            Using cmd As New SQLiteCommand(createApiKeysQuery, connection)
                cmd.ExecuteNonQuery()
            End Using
            
        Catch ex As Exception
            ' Log error ma non bloccare
            Debug.WriteLine($"Errore inizializzazione tabella ApiKeys: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Inizializza la tabella per le impostazioni del conto corrente
    ''' </summary>
    Private Shared Sub InizializzaTabellaImpostazioniConto(connection As SQLiteConnection)
        Try
            Dim createImpostazioniQuery = "
                CREATE TABLE IF NOT EXISTS ImpostazioniConto (
                    Chiave TEXT PRIMARY KEY,
                    Valore TEXT NOT NULL,
                    DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                    DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
                )"

            Using cmd As New SQLiteCommand(createImpostazioniQuery, connection)
                cmd.ExecuteNonQuery()
            End Using

            ' Inizializza saldo iniziale se non esiste
            Dim checkSaldoQuery = "SELECT COUNT(*) FROM ImpostazioniConto WHERE Chiave = 'SaldoInizialeContoCorrente'"
            Using checkCmd As New SQLiteCommand(checkSaldoQuery, connection)
                Dim count = Convert.ToInt32(checkCmd.ExecuteScalar())
                If count = 0 Then
                    Dim insertSaldoQuery = "INSERT INTO ImpostazioniConto (Chiave, Valore) VALUES ('SaldoInizialeContoCorrente', '0')"
                    Using insertCmd As New SQLiteCommand(insertSaldoQuery, connection)
                        insertCmd.ExecuteNonQuery()
                    End Using
                End If
            End Using

            ' Inizializza data punto zero se non esiste
            Dim checkDataQuery = "SELECT COUNT(*) FROM ImpostazioniConto WHERE Chiave = 'DataPuntoZeroSaldo'"
            Using checkCmd As New SQLiteCommand(checkDataQuery, connection)
                Dim count = Convert.ToInt32(checkCmd.ExecuteScalar())
                If count = 0 Then
                    ' Se non esiste, la impostiamo a una data molto antica (1900) per includere tutte le transazioni
                    Dim insertDataQuery = "INSERT INTO ImpostazioniConto (Chiave, Valore) VALUES ('DataPuntoZeroSaldo', '1900-01-01 00:00:00')"
                    Using insertCmd As New SQLiteCommand(insertDataQuery, connection)
                        insertCmd.ExecuteNonQuery()
                    End Using
                End If
            End Using

        Catch ex As Exception
            ' Log error ma non bloccare
            Debug.WriteLine($"Errore inizializzazione tabella ImpostazioniConto: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Salva un'impostazione del conto corrente nel database
    ''' </summary>
    Public Shared Sub SalvaImpostazioneConto(chiave As String, valore As String)
        Try
            Using connection As New SQLiteConnection(GetConnectionString())
                connection.Open()

                Dim query = "INSERT OR REPLACE INTO ImpostazioniConto (Chiave, Valore, DataModifica) VALUES (@chiave, @valore, CURRENT_TIMESTAMP)"
                Using cmd As New SQLiteCommand(query, connection)
                    cmd.Parameters.AddWithValue("@chiave", chiave)
                    cmd.Parameters.AddWithValue("@valore", valore)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"Errore salvataggio impostazione {chiave}: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Carica un'impostazione del conto corrente dal database
    ''' </summary>
    Public Shared Function CaricaImpostazioneConto(chiave As String, Optional valoreDefault As String = "") As String
        Try
            ' Assicurati che il database sia inizializzato
            InitializeDatabase()

            Using connection As New SQLiteConnection(GetConnectionString())
                connection.Open()

                Dim query = "SELECT Valore FROM ImpostazioniConto WHERE Chiave = @chiave"
                Using cmd As New SQLiteCommand(query, connection)
                    cmd.Parameters.AddWithValue("@chiave", chiave)
                    Dim risultato = cmd.ExecuteScalar()

                    If risultato IsNot Nothing Then
                        Return risultato.ToString()
                    Else
                        Return valoreDefault
                    End If
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"Errore caricamento impostazione {chiave}: {ex.Message}")
            Return valoreDefault
        End Try
    End Function

    ''' <summary>
    ''' Salva il saldo iniziale del conto corrente SENZA modificare il punto zero
    ''' </summary>
    Public Shared Sub SalvaSaldoInizialeContoCorrente(saldoIniziale As Decimal)
        SalvaImpostazioneConto("SaldoInizialeContoCorrente", saldoIniziale.ToString("F2"))
    End Sub

    ''' <summary>
    ''' Salva il saldo iniziale E crea un nuovo punto zero temporale (solo dalle impostazioni)
    ''' </summary>
    Public Shared Sub SalvaSaldoInizialeConPuntoZero(saldoIniziale As Decimal)
        SalvaImpostazioneConto("SaldoInizialeContoCorrente", saldoIniziale.ToString("F2"))

        ' Trova l'ID dell'ultima transazione esistente e salva come punto zero
        Dim ultimoId As Integer = 0
        Try
            Using connection As New SQLiteConnection(GetConnectionString())
                connection.Open()
                Dim query = "SELECT COALESCE(MAX(ID), 0) FROM Transazioni"
                Using command As New SQLiteCommand(query, connection)
                    ultimoId = Convert.ToInt32(command.ExecuteScalar())
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"Errore nel recupero ultimo ID: {ex.Message}")
        End Try

        ' Salva l'ID come punto zero (tutte le transazioni con ID > questo saranno contate)
        SalvaImpostazioneConto("IdPuntoZeroSaldo", ultimoId.ToString())
    End Sub

    ''' <summary>
    ''' Carica il saldo iniziale del conto corrente
    ''' </summary>
    Public Shared Function CaricaSaldoInizialeContoCorrente() As Decimal
        Dim valoreStringa = CaricaImpostazioneConto("SaldoInizialeContoCorrente", "0")
        Dim saldoIniziale As Decimal
        If Decimal.TryParse(valoreStringa, saldoIniziale) Then
            Return saldoIniziale
        Else
            Return 0D
        End If
    End Function

    ''' <summary>
    ''' Carica la data del punto zero per il saldo
    ''' </summary>
    Public Shared Function CaricaDataPuntoZeroSaldo() As DateTime
        Dim valoreStringa = CaricaImpostazioneConto("DataPuntoZeroSaldo", "1900-01-01 00:00:00")
        Dim dataPuntoZero As DateTime
        If DateTime.TryParse(valoreStringa, dataPuntoZero) Then
            Return dataPuntoZero
        Else
            Return New DateTime(1900, 1, 1) ' Default molto antico
        End If
    End Function

    ''' <summary>
    ''' Carica l'ID del punto zero per il saldo (tutte le transazioni con ID > questo vengono contate)
    ''' </summary>
    Public Shared Function CaricaIdPuntoZeroSaldo() As Integer
        Dim valoreStringa = CaricaImpostazioneConto("IdPuntoZeroSaldo", "0")
        Dim idPuntoZero As Integer
        If Integer.TryParse(valoreStringa, idPuntoZero) Then
            Return idPuntoZero
        Else
            Return 0 ' Default: conta tutte le transazioni
        End If
    End Function

    ' Proprietà per ottenere la connection string
    Public Shared ReadOnly Property GetConnectionString As String
        Get
            Return ConnectionString
        End Get
    End Property

    ''' <summary>
    ''' Imposta il conto corrente attivo (per sistema multi-conto)
    ''' Cambia dinamicamente il ConnectionString verso il database del conto specificato
    ''' </summary>
    Public Shared Sub SetCurrentConto(percorsoDatabase As String)
        If String.IsNullOrEmpty(percorsoDatabase) Then
            Throw New ArgumentException("Percorso database non valido", NameOf(percorsoDatabase))
        End If

        If Not File.Exists(percorsoDatabase) Then
            Throw New FileNotFoundException("Database conto non trovato", percorsoDatabase)
        End If

        ' Aggiorna connection string
        ConnectionString = $"Data Source={percorsoDatabase};Version=3;Busy Timeout=5000;Pooling=true;Max Pool Size=100;"

        Debug.WriteLine($"DatabaseManager: Conto corrente cambiato -> {Path.GetFileName(percorsoDatabase)}")
    End Sub

    ''' <summary>
    ''' Crea le tabelle standard in una connessione esistente (usato da ContoManager)
    ''' </summary>
    Public Shared Sub CreateTablesIfNotExist(connection As SQLiteConnection)
        If connection Is Nothing OrElse connection.State <> System.Data.ConnectionState.Open Then
            Throw New InvalidOperationException("Connessione database non valida")
        End If

        ' Tabella Transazioni
        Dim createTransQuery As String = "
            CREATE TABLE IF NOT EXISTS Transazioni (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Data DATE NOT NULL,
                Importo DECIMAL(10,2) NOT NULL,
                Descrizione TEXT NOT NULL,
                Causale TEXT DEFAULT '',
                MacroCategoria TEXT DEFAULT '',
                Categoria TEXT DEFAULT '',
                Necessita TEXT DEFAULT '',
                Frequenza TEXT DEFAULT '',
                Stagionalita TEXT DEFAULT '',
                DataInserimento DATETIME DEFAULT CURRENT_TIMESTAMP,
                DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
            );"

        Using cmd As New SQLiteCommand(createTransQuery, connection)
            cmd.ExecuteNonQuery()
        End Using

        ' Tabella Budget
        Dim createBudgetQuery As String = "
            CREATE TABLE IF NOT EXISTS Budget (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Categoria TEXT NOT NULL,
                MacroCategoria TEXT,
                Limite DECIMAL(10,2),
                Mese INTEGER,
                Anno INTEGER,
                UNIQUE(Categoria, MacroCategoria, Mese, Anno)
            );"

        Using cmd As New SQLiteCommand(createBudgetQuery, connection)
            cmd.ExecuteNonQuery()
        End Using

        ' Tabella Obiettivi
        Dim createObiettiviQuery As String = "
            CREATE TABLE IF NOT EXISTS Obiettivi (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Nome TEXT NOT NULL,
                Descrizione TEXT,
                Importo DECIMAL(10,2),
                ImportoRaggiunto DECIMAL(10,2) DEFAULT 0,
                DataInizio DATE,
                DataScadenza DATE,
                Completato INTEGER DEFAULT 0
            );"

        Using cmd As New SQLiteCommand(createObiettiviQuery, connection)
            cmd.ExecuteNonQuery()
        End Using

        ' Tabella ConfigurazioneStipendi
        Dim createConfigStipendiQuery As String = "
            CREATE TABLE IF NOT EXISTS ConfigurazioneStipendi (
                Id INTEGER PRIMARY KEY,
                GiornoDefault INTEGER NOT NULL DEFAULT 23,
                RegoleWeekend TEXT NOT NULL DEFAULT 'ANTICIPA',
                DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
            );"

        Using cmd As New SQLiteCommand(createConfigStipendiQuery, connection)
            cmd.ExecuteNonQuery()
        End Using

        ' Tabella EccezioniStipendi
        Dim createEccezioniQuery As String = "
            CREATE TABLE IF NOT EXISTS EccezioniStipendi (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Mese INTEGER NOT NULL,
                GiornoSpeciale INTEGER NOT NULL,
                Motivo TEXT,
                Attivo INTEGER DEFAULT 1,
                UNIQUE(Mese)
            );"

        Using cmd As New SQLiteCommand(createEccezioniQuery, connection)
            cmd.ExecuteNonQuery()
        End Using

        ' Tabella ImpostazioniConto
        Dim createImpostazioniQuery As String = "
            CREATE TABLE IF NOT EXISTS ImpostazioniConto (
                Chiave TEXT PRIMARY KEY,
                Valore TEXT NOT NULL,
                DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
            );"

        Using cmd As New SQLiteCommand(createImpostazioniQuery, connection)
            cmd.ExecuteNonQuery()
        End Using

        ' Tabella MacroCategorie (per compatibilità futura)
        Dim createMacroCategorieQuery As String = "
            CREATE TABLE IF NOT EXISTS MacroCategorie (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Nome TEXT NOT NULL UNIQUE,
                DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP
            );"

        Using cmd As New SQLiteCommand(createMacroCategorieQuery, connection)
            cmd.ExecuteNonQuery()
        End Using

        ' Indici per performance
        Dim indexes = New List(Of String) From {
            "CREATE INDEX IF NOT EXISTS idx_trans_data ON Transazioni(Data DESC);",
            "CREATE INDEX IF NOT EXISTS idx_trans_macro ON Transazioni(MacroCategoria);",
            "CREATE INDEX IF NOT EXISTS idx_trans_categoria ON Transazioni(Categoria);",
            "CREATE INDEX IF NOT EXISTS idx_budget_periodo ON Budget(Anno, Mese);"
        }

        For Each indexSql In indexes
            Using cmd As New SQLiteCommand(indexSql, connection)
                cmd.ExecuteNonQuery()
            End Using
        Next

        Debug.WriteLine("Tabelle database conto create con successo")
    End Sub

End Class
