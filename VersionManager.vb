Imports System.Reflection

Namespace MoneyMind

    ''' <summary>
    ''' Gestore centralizzato della versione dell'applicazione.
    ''' MODIFICA SOLO QUESTO FILE PER CAMBIARE LA VERSIONE!
    ''' </summary>
    Public Class VersionManager

        ''' <summary>
        ''' ⚠️ VERSIONE CORRENTE DELL'APPLICAZIONE - MODIFICA SOLO QUESTO VALORE ⚠️
        ''' Formato: Major.Minor.Patch (es. 1.4.12)
        ''' </summary>
        Public Const CURRENT_VERSION As String = "1.4.14"

        ''' <summary>
        ''' Ottiene la versione corrente dell'applicazione
        ''' </summary>
        Public Shared Function GetVersion() As String
            Return CURRENT_VERSION
        End Function

        ''' <summary>
        ''' Ottiene la versione corrente con formato Assembly (x.x.x.0)
        ''' </summary>
        Public Shared Function GetAssemblyVersion() As String
            Return CURRENT_VERSION & ".0"
        End Function

        ''' <summary>
        ''' Ottiene la versione dall'assembly se disponibile, altrimenti usa la costante
        ''' </summary>
        Public Shared Function GetVersionFromAssembly() As String
            Try
                Dim currentAssembly As Assembly = Assembly.GetExecutingAssembly()
                Dim version As Version = currentAssembly.GetName().Version
                If version IsNot Nothing Then
                    Return $"{version.Major}.{version.Minor}.{version.Build}"
                End If
            Catch ex As Exception
                Debug.WriteLine($"Errore lettura versione da assembly: {ex.Message}")
            End Try

            Return CURRENT_VERSION
        End Function

        ''' <summary>
        ''' Confronta la versione corrente con un'altra versione
        ''' </summary>
        ''' <returns>-1 se corrente è minore, 0 se uguale, 1 se corrente è maggiore</returns>
        Public Shared Function CompareVersion(otherVersion As String) As Integer
            Try
                ' Normalizza le stringhe di versione (rimuovi 'v' se presente)
                Dim currentVersionString = CURRENT_VERSION.TrimStart("v"c)
                Dim otherVersionString = If(otherVersion, "").TrimStart("v"c)

                If String.IsNullOrWhiteSpace(otherVersionString) Then
                    Debug.WriteLine("Versione da confrontare è vuota")
                    Return 0
                End If

                Dim current As Version = Nothing
                Dim other As Version = Nothing

                ' Usa TryParse invece di costruttore diretto per evitare FormatException
                If Not Version.TryParse(currentVersionString, current) Then
                    Debug.WriteLine($"Impossibile parsare versione corrente: '{currentVersionString}'")
                    Return 0
                End If

                If Not Version.TryParse(otherVersionString, other) Then
                    Debug.WriteLine($"Impossibile parsare versione da confrontare: '{otherVersionString}'")
                    Return 0
                End If

                Return current.CompareTo(other)
            Catch ex As Exception
                Debug.WriteLine($"Errore confronto versioni: {ex.Message}")
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' Verifica se è disponibile un aggiornamento
        ''' </summary>
        Public Shared Function IsUpdateAvailable(latestVersion As String) As Boolean
            Return CompareVersion(latestVersion) < 0
        End Function

    End Class

End Namespace
