Imports System.Data.SQLite
Imports System.Linq
Imports MoneyMind.Models
Imports MoneyMind.Services

''' <summary>
''' Servizio per la gestione di budget e obiettivi di risparmio
''' Tutti i calcoli sono basati sui PERIODI STIPENDIALI dal GestorePeriodi
''' </summary>
Public Class BudgetService

#Region "Budget Management"

    ''' <summary>
    ''' Salva o aggiorna un budget per una categoria
    ''' NOTA: Il budget è per PERIODO STIPENDIALE, non per mese calendario
    ''' </summary>
    Public Shared Sub SalvaBudget(budget As BudgetItem)
        Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
            connection.Open()

            Dim query As String = "
                INSERT OR REPLACE INTO Budget
                (Id, Categoria, MacroCategoria, Limite, Mese, Anno)
                VALUES
                (@Id, @Categoria, @MacroCategoria, @Limite, @Mese, @Anno)"

            Using command As New SQLiteCommand(query, connection)
                If budget.Id = 0 Then
                    command.CommandText = "
                        INSERT INTO Budget
                        (Categoria, MacroCategoria, Limite, Mese, Anno)
                        VALUES
                        (@Categoria, @MacroCategoria, @Limite, @Mese, @Anno);
                        SELECT last_insert_rowid();"

                    command.Parameters.AddWithValue("@Categoria", budget.Categoria)
                    command.Parameters.AddWithValue("@MacroCategoria", If(budget.MacroCategoria, ""))
                    command.Parameters.AddWithValue("@Limite", budget.BudgetMensile)
                    command.Parameters.AddWithValue("@Mese", budget.Mese)
                    command.Parameters.AddWithValue("@Anno", budget.Anno)

                    budget.Id = Convert.ToInt32(command.ExecuteScalar())
                Else
                    command.Parameters.AddWithValue("@Id", budget.Id)
                    command.Parameters.AddWithValue("@Categoria", budget.Categoria)
                    command.Parameters.AddWithValue("@MacroCategoria", If(budget.MacroCategoria, ""))
                    command.Parameters.AddWithValue("@Limite", budget.BudgetMensile)
                    command.Parameters.AddWithValue("@Mese", budget.Mese)
                    command.Parameters.AddWithValue("@Anno", budget.Anno)

                    command.ExecuteNonQuery()
                End If
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Carica tutti i budget per un periodo stipendiale specifico
    ''' </summary>
    Public Shared Function CaricaBudget(anno As Integer, mese As Integer) As List(Of BudgetItem)
        Dim budgets As New List(Of BudgetItem)

        Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
            connection.Open()

            Dim query As String = "
                SELECT Id, Categoria, MacroCategoria, Limite, Mese, Anno
                FROM Budget
                WHERE Anno = @Anno AND Mese = @Mese"

            Using command As New SQLiteCommand(query, connection)
                command.Parameters.AddWithValue("@Anno", anno)
                command.Parameters.AddWithValue("@Mese", mese)

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        budgets.Add(New BudgetItem With {
                            .Id = reader.GetInt32(0),
                            .Categoria = reader.GetString(1),
                            .MacroCategoria = If(reader.IsDBNull(2), "", reader.GetString(2)),
                            .BudgetMensile = If(reader.IsDBNull(3), 0, reader.GetDecimal(3)),
                            .Mese = reader.GetInt32(4),
                            .Anno = reader.GetInt32(5),
                            .DataCreazione = DateTime.Now,
                            .DataModifica = DateTime.Now
                        })
                    End While
                End Using
            End Using
        End Using

        Return budgets
    End Function

    ''' <summary>
    ''' Carica i budget per il PERIODO STIPENDIALE CORRENTE
    ''' Usa GestorePeriodi per ottenere il periodo corrente
    ''' </summary>
    Public Shared Function CaricaBudgetPeriodoCorrente() As List(Of BudgetItem)
        Dim periodoCorrente = GestorePeriodi.GetPeriodoCorrente()
        Return CaricaBudget(periodoCorrente.Anno, periodoCorrente.Mese)
    End Function

    ''' <summary>
    ''' Calcola la spesa corrente per ogni budget basandosi sulle transazioni del PERIODO STIPENDIALE
    ''' </summary>
    Public Shared Sub AggiornaSpeseBudget(budgets As List(Of BudgetItem), transazioni As IEnumerable(Of TransazioneViewModel))
        If budgets Is Nothing OrElse budgets.Count = 0 Then Return

        ' Usa il periodo stipendiale corrente
        Dim periodoCorrente = GestorePeriodi.GetPeriodoCorrente()

        ' Filtra transazioni nel periodo stipendiale corrente
        Dim transazioniPeriodo = transazioni.Where(Function(t)
                                                        Return t.DataTransazione >= periodoCorrente.DataInizio AndAlso
                                                               t.DataTransazione <= periodoCorrente.DataFine AndAlso
                                                               t.Importo < 0 ' Solo uscite
                                                    End Function).ToList()

        ' Calcola spesa per ogni budget
        For Each budget In budgets
            Dim spesa = transazioniPeriodo.Where(Function(t) t.Categoria = budget.Categoria).
                                          Sum(Function(t) Math.Abs(t.Importo))
            budget.SpesoCorrente = spesa
        Next
    End Sub

    ''' <summary>
    ''' Elimina un budget
    ''' </summary>
    Public Shared Sub EliminaBudget(budgetId As Integer)
        Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
            connection.Open()

            Dim query As String = "DELETE FROM Budget WHERE Id = @Id"

            Using command As New SQLiteCommand(query, connection)
                command.Parameters.AddWithValue("@Id", budgetId)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

#End Region

#Region "Obiettivi Management"

    ''' <summary>
    ''' Salva o aggiorna un obiettivo di risparmio
    ''' </summary>
    Public Shared Sub SalvaObiettivo(obiettivo As ObiettivoRisparmio)
        Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
            connection.Open()

            If obiettivo.Id = 0 Then
                Dim query As String = "
                    INSERT INTO Obiettivi
                    (Nome, Descrizione, Importo, ImportoRaggiunto, DataInizio, DataScadenza, Completato)
                    VALUES
                    (@Nome, @Descrizione, @Importo, @ImportoRaggiunto, @DataInizio, @DataScadenza, @Completato);
                    SELECT last_insert_rowid();"

                Using command As New SQLiteCommand(query, connection)
                    command.Parameters.AddWithValue("@Nome", obiettivo.Nome)
                    command.Parameters.AddWithValue("@Descrizione", If(obiettivo.Descrizione, ""))
                    command.Parameters.AddWithValue("@Importo", obiettivo.ImportoTarget)
                    command.Parameters.AddWithValue("@ImportoRaggiunto", obiettivo.ImportoCorrente)
                    command.Parameters.AddWithValue("@DataInizio", obiettivo.DataInizio)
                    command.Parameters.AddWithValue("@DataScadenza", If(obiettivo.DataScadenza, DBNull.Value))
                    command.Parameters.AddWithValue("@Completato", If(obiettivo.Completato, 1, 0))

                    obiettivo.Id = Convert.ToInt32(command.ExecuteScalar())
                End Using
            Else
                Dim query As String = "
                    UPDATE Obiettivi SET
                    Nome = @Nome,
                    Descrizione = @Descrizione,
                    Importo = @Importo,
                    ImportoRaggiunto = @ImportoRaggiunto,
                    DataInizio = @DataInizio,
                    DataScadenza = @DataScadenza,
                    Completato = @Completato
                    WHERE Id = @Id"

                Using command As New SQLiteCommand(query, connection)
                    command.Parameters.AddWithValue("@Id", obiettivo.Id)
                    command.Parameters.AddWithValue("@Nome", obiettivo.Nome)
                    command.Parameters.AddWithValue("@Descrizione", If(obiettivo.Descrizione, ""))
                    command.Parameters.AddWithValue("@Importo", obiettivo.ImportoTarget)
                    command.Parameters.AddWithValue("@ImportoRaggiunto", obiettivo.ImportoCorrente)
                    command.Parameters.AddWithValue("@DataInizio", obiettivo.DataInizio)
                    command.Parameters.AddWithValue("@DataScadenza", If(obiettivo.DataScadenza, DBNull.Value))
                    command.Parameters.AddWithValue("@Completato", If(obiettivo.Completato, 1, 0))

                    command.ExecuteNonQuery()
                End Using
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Carica tutti gli obiettivi
    ''' </summary>
    Public Shared Function CaricaObiettivi() As List(Of ObiettivoRisparmio)
        Dim obiettivi As New List(Of ObiettivoRisparmio)

        Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
            connection.Open()

            Dim query As String = "
                SELECT Id, Nome, Descrizione, Importo, ImportoRaggiunto, DataInizio, DataScadenza, Completato
                FROM Obiettivi
                ORDER BY Completato ASC, DataScadenza ASC"

            Using command As New SQLiteCommand(query, connection)
                Using reader = command.ExecuteReader()
                    While reader.Read()
                        obiettivi.Add(New ObiettivoRisparmio With {
                            .Id = reader.GetInt32(0),
                            .Nome = reader.GetString(1),
                            .Descrizione = If(reader.IsDBNull(2), "", reader.GetString(2)),
                            .ImportoTarget = If(reader.IsDBNull(3), 0, reader.GetDecimal(3)),
                            .ImportoCorrente = If(reader.IsDBNull(4), 0, reader.GetDecimal(4)),
                            .DataInizio = If(reader.IsDBNull(5), DateTime.Now, reader.GetDateTime(5)),
                            .DataScadenza = If(reader.IsDBNull(6), Nothing, reader.GetDateTime(6)),
                            .Completato = If(reader.IsDBNull(7), False, reader.GetInt32(7) = 1),
                            .Colore = "#4CAF50",
                            .DataCreazione = DateTime.Now,
                            .DataModifica = DateTime.Now
                        })
                    End While
                End Using
            End Using
        End Using

        Return obiettivi
    End Function

    ''' <summary>
    ''' Elimina un obiettivo
    ''' </summary>
    Public Shared Sub EliminaObiettivo(obiettivoId As Integer)
        Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
            connection.Open()

            Dim query As String = "DELETE FROM Obiettivi WHERE Id = @Id"

            Using command As New SQLiteCommand(query, connection)
                command.Parameters.AddWithValue("@Id", obiettivoId)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Aggiorna il progresso di un obiettivo
    ''' </summary>
    Public Shared Sub AggiornaProgressoObiettivo(obiettivoId As Integer, importoCorrente As Decimal)
        Using connection As New SQLiteConnection(DatabaseManager.GetConnectionString())
            connection.Open()

            Dim query As String = "
                UPDATE Obiettivi SET
                ImportoRaggiunto = @ImportoRaggiunto,
                Completato = CASE WHEN @ImportoRaggiunto >= Importo THEN 1 ELSE 0 END
                WHERE Id = @Id"

            Using command As New SQLiteCommand(query, connection)
                command.Parameters.AddWithValue("@Id", obiettivoId)
                command.Parameters.AddWithValue("@ImportoRaggiunto", importoCorrente)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

#End Region

End Class
