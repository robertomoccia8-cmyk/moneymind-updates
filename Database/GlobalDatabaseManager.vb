Imports System.Data.SQLite
Imports System.IO
Imports MoneyMind

''' <summary>
''' Manager per il database globale condiviso tra tutti i conti correnti
''' Gestisce: ContiCorrenti, Pattern globali, API Keys, MacroCategorie, Impostazioni globali
''' Database: %APPDATA%\MoneyMind\MoneyMind_Global.db
''' </summary>
Public Class GlobalDatabaseManager
    Private Shared _connectionString As String = ""
    Private Shared _isInitialized As Boolean = False
    Private Shared ReadOnly _lockObject As New Object()

    ''' <summary>
    ''' Ottiene il connection string del database globale
    ''' </summary>
    Public Shared Function GetConnectionString() As String
        If String.IsNullOrEmpty(_connectionString) Then
            InitializeGlobalDatabase()
        End If
        Return _connectionString
    End Function

    ''' <summary>
    ''' Ottiene il percorso del database globale
    ''' </summary>
    Public Shared Function GetDatabasePath() As String
        ' Prova ApplicationData, se fallisce usa LocalApplicationData come fallback
        Dim appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

        ' Se ApplicationData è null o vuoto, prova LocalApplicationData
        If String.IsNullOrEmpty(appDataPath) Then
            appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        End If

        ' Se ancora null, usa directory corrente come ultima risorsa
        If String.IsNullOrEmpty(appDataPath) Then
            appDataPath = Path.Combine(Environment.CurrentDirectory, "Data")
        End If

        Dim moneyMindFolder = Path.Combine(appDataPath, "MoneyMind")

        If Not Directory.Exists(moneyMindFolder) Then
            Directory.CreateDirectory(moneyMindFolder)
        End If

        Return Path.Combine(moneyMindFolder, "MoneyMind_Global.db")
    End Function

    ''' <summary>
    ''' Inizializza il database globale e crea tutte le tabelle necessarie
    ''' </summary>
    Public Shared Sub InitializeGlobalDatabase()
        SyncLock _lockObject
            If _isInitialized Then Return

            Try
                Dim dbPath = GetDatabasePath()
                _connectionString = $"Data Source={dbPath};Version=3;Busy Timeout=5000;Pooling=true;Max Pool Size=100;"

                Using conn As New SQLiteConnection(_connectionString)
                    conn.Open()

                    ' Abilita WAL mode per migliori performance concorrenti
                    Using cmd As New SQLiteCommand("PRAGMA journal_mode=WAL;", conn)
                        cmd.ExecuteNonQuery()
                    End Using

                    ' Crea tutte le tabelle
                    CreaTabellaContiCorrenti(conn)
                    CreaTabellaPattern(conn)
                    CreaTabellaPatternPersonalizzati(conn)
                    CreaTabellaMacroCategorie(conn)
                    CreaTabellaApiKeys(conn)
                    CreaTabellaImpostazioniGlobali(conn)
                    CreaTabellaLogEventi(conn)

                    ' Inizializza dati default
                    InizializzaDatiDefault(conn)
                End Using

                _isInitialized = True

            Catch ex As Exception
                Throw New Exception($"Errore inizializzazione database globale: {ex.Message}", ex)
            End Try
        End SyncLock
    End Sub

#Region "Creazione Tabelle"

    Private Shared Sub CreaTabellaContiCorrenti(conn As SQLiteConnection)
        Dim sql = "
            CREATE TABLE IF NOT EXISTS ContiCorrenti (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                NomeFile TEXT NOT NULL UNIQUE,
                NomeVisualizzato TEXT NOT NULL,
                Icona TEXT,
                Colore TEXT,
                Ordinamento INTEGER DEFAULT 0,
                DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                UltimoAccesso DATETIME,
                Attivo INTEGER DEFAULT 1,
                Note TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_conti_attivi ON ContiCorrenti(Attivo);
            CREATE INDEX IF NOT EXISTS idx_conti_ordinamento ON ContiCorrenti(Ordinamento);
        "

        Using cmd As New SQLiteCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub CreaTabellaPattern(conn As SQLiteConnection)
        Dim sql = "
            CREATE TABLE IF NOT EXISTS Pattern (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Parola TEXT NOT NULL UNIQUE COLLATE NOCASE,
                MacroCategoria TEXT NOT NULL DEFAULT '',
                Categoria TEXT NOT NULL DEFAULT '',
                TipoFonte TEXT DEFAULT '',
                Peso INTEGER DEFAULT 5,
                DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                UltimoUtilizzo DATETIME,
                NumeroUtilizzi INTEGER DEFAULT 0
            );

            CREATE INDEX IF NOT EXISTS idx_pattern_parola ON Pattern(Parola COLLATE NOCASE);
            CREATE INDEX IF NOT EXISTS idx_pattern_macrocategoria ON Pattern(MacroCategoria);
            CREATE INDEX IF NOT EXISTS idx_pattern_peso ON Pattern(Peso DESC);
        "

        Using cmd As New SQLiteCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub CreaTabellaPatternPersonalizzati(conn As SQLiteConnection)
        Dim sql = "
            CREATE TABLE IF NOT EXISTS PatternPersonalizzati (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                ContoId INTEGER NOT NULL,
                Parola TEXT NOT NULL COLLATE NOCASE,
                MacroCategoria TEXT NOT NULL,
                Categoria TEXT NOT NULL,
                Peso INTEGER DEFAULT 5,
                SovrascriviGlobale INTEGER DEFAULT 0,
                DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(ContoId, Parola)
            );

            CREATE INDEX IF NOT EXISTS idx_pattern_pers_conto ON PatternPersonalizzati(ContoId);
            CREATE INDEX IF NOT EXISTS idx_pattern_pers_parola ON PatternPersonalizzati(Parola COLLATE NOCASE);
        "

        Using cmd As New SQLiteCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub CreaTabellaMacroCategorie(conn As SQLiteConnection)
        Dim sql = "
            CREATE TABLE IF NOT EXISTS MacroCategorie (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Nome TEXT NOT NULL UNIQUE,
                Icona TEXT,
                Colore TEXT,
                Descrizione TEXT,
                DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE INDEX IF NOT EXISTS idx_macro_nome ON MacroCategorie(Nome);
        "

        Using cmd As New SQLiteCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub CreaTabellaApiKeys(conn As SQLiteConnection)
        Dim sql = "
            CREATE TABLE IF NOT EXISTS ApiKeys (
                Nome TEXT PRIMARY KEY,
                ChiaveCriptata TEXT NOT NULL,
                DataCreazione DATETIME DEFAULT CURRENT_TIMESTAMP,
                UltimoUtilizzo DATETIME,
                DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
            );
        "

        Using cmd As New SQLiteCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub CreaTabellaImpostazioniGlobali(conn As SQLiteConnection)
        Dim sql = "
            CREATE TABLE IF NOT EXISTS ImpostazioniGlobali (
                Chiave TEXT PRIMARY KEY,
                Valore TEXT NOT NULL,
                Tipo TEXT DEFAULT 'string',
                Descrizione TEXT,
                DataModifica DATETIME DEFAULT CURRENT_TIMESTAMP
            );
        "

        Using cmd As New SQLiteCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub CreaTabellaLogEventi(conn As SQLiteConnection)
        Dim sql = "
            CREATE TABLE IF NOT EXISTS LogEventi (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                ContoId INTEGER,
                TipoEvento TEXT NOT NULL,
                Descrizione TEXT,
                Dettagli TEXT,
                DataOra DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE INDEX IF NOT EXISTS idx_log_conto ON LogEventi(ContoId);
            CREATE INDEX IF NOT EXISTS idx_log_tipo ON LogEventi(TipoEvento);
            CREATE INDEX IF NOT EXISTS idx_log_data ON LogEventi(DataOra DESC);
        "

        Using cmd As New SQLiteCommand(sql, conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

#End Region

#Region "Inizializzazione Dati Default"

    Private Shared Sub InizializzaDatiDefault(conn As SQLiteConnection)
        ' Inserisce le 12 MacroCategorie standard se la tabella è vuota
        Dim checkSql = "SELECT COUNT(*) FROM MacroCategorie"
        Using cmd As New SQLiteCommand(checkSql, conn)
            Dim count = Convert.ToInt32(cmd.ExecuteScalar())
            If count = 0 Then
                InserisciMacroCategorie(conn)
            End If
        End Using

        ' Inserisce impostazioni globali di default
        InserisciImpostazioniDefault(conn)
    End Sub

    Private Shared Sub InserisciMacroCategorie(conn As SQLiteConnection)
        ' Le 12 MacroCategorie standard ottimizzate (parole singole)
        Dim macroCategorie = New Dictionary(Of String, (Icona As String, Colore As String)) From {
            {"Casa", ("🏠", "#6f42c1")},
            {"Finanze", ("💰", "#28a745")},
            {"Mobilità", ("🚗", "#17a2b8")},
            {"Ristorazione", ("🍽️", "#fd7e14")},
            {"Salute", ("⚕️", "#dc3545")},
            {"Servizi", ("🔧", "#6c757d")},
            {"Shopping", ("🛍️", "#e83e8c")},
            {"Spesa", ("🛒", "#20c997")},
            {"Svago", ("🎭", "#ffc107")},
            {"Tasse", ("📋", "#343a40")},
            {"Trasporti", ("🚆", "#007bff")},
            {"Viaggi", ("✈️", "#17a2b8")}
        }

        Dim sql = "INSERT OR IGNORE INTO MacroCategorie (Nome, Icona, Colore) VALUES (@Nome, @Icona, @Colore)"

        For Each kvp In macroCategorie
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Nome", kvp.Key)
                cmd.Parameters.AddWithValue("@Icona", kvp.Value.Icona)
                cmd.Parameters.AddWithValue("@Colore", kvp.Value.Colore)
                cmd.ExecuteNonQuery()
            End Using
        Next
    End Sub

    Private Shared Sub InserisciImpostazioniDefault(conn As SQLiteConnection)
        ' Impostazioni di default
        Dim impostazioni = New Dictionary(Of String, (Valore As String, Tipo As String, Descrizione As String)) From {
            {"UltimoContoSelezionato", ("0", "int", "ID ultimo conto selezionato dall'utente")},
            {"TemaScuro", ("false", "bool", "Abilita tema scuro UI")},
            {"LinguaInterfaccia", ("it-IT", "string", "Lingua interfaccia utente")},
            {"VersioneDatabase", (MoneyMind.VersionManager.CURRENT_VERSION, "string", "Versione database globale")},
            {"PrimoAvvio", ("true", "bool", "Flag primo avvio applicazione")}
        }

        Dim sql = "INSERT OR IGNORE INTO ImpostazioniGlobali (Chiave, Valore, Tipo, Descrizione) VALUES (@Chiave, @Valore, @Tipo, @Desc)"

        For Each kvp In impostazioni
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Chiave", kvp.Key)
                cmd.Parameters.AddWithValue("@Valore", kvp.Value.Valore)
                cmd.Parameters.AddWithValue("@Tipo", kvp.Value.Tipo)
                cmd.Parameters.AddWithValue("@Desc", kvp.Value.Descrizione)
                cmd.ExecuteNonQuery()
            End Using
        Next
    End Sub

#End Region

#Region "Utility Methods"

    ''' <summary>
    ''' Verifica se il database globale esiste
    ''' </summary>
    Public Shared Function DatabaseExists() As Boolean
        Return File.Exists(GetDatabasePath())
    End Function

    ''' <summary>
    ''' Resetta il flag di inizializzazione (per test o re-init)
    ''' </summary>
    Public Shared Sub Reset()
        _isInitialized = False
        _connectionString = String.Empty
    End Sub

    ''' <summary>
    ''' Logga un evento nel database globale
    ''' </summary>
    Public Shared Sub LogEvento(contoId As Integer?, tipoEvento As String, descrizione As String, Optional dettagli As String = Nothing)
        Try
            Using conn As New SQLiteConnection(GetConnectionString())
                conn.Open()

                Dim sql = "INSERT INTO LogEventi (ContoId, TipoEvento, Descrizione, Dettagli) VALUES (@ContoId, @Tipo, @Desc, @Det)"
                Using cmd As New SQLiteCommand(sql, conn)
                    If contoId.HasValue Then
                        cmd.Parameters.AddWithValue("@ContoId", contoId.Value)
                    Else
                        cmd.Parameters.AddWithValue("@ContoId", DBNull.Value)
                    End If
                    cmd.Parameters.AddWithValue("@Tipo", tipoEvento)
                    cmd.Parameters.AddWithValue("@Desc", descrizione)
                    cmd.Parameters.AddWithValue("@Det", If(dettagli, DBNull.Value))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            ' Ignora errori di log per non bloccare operazioni principali
            Debug.WriteLine($"Errore log evento: {ex.Message}")
        End Try
    End Sub

#End Region
End Class
