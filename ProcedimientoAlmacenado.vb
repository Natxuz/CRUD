Imports System.Data.OleDb
Imports System.Text

Module procedimiento
    Private mCadenaConexion As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=HeMi;Data Source=MITIESTOPC\SERVIDOR2014"
    Private mConexion As New OleDbConnection
    Private mError As New StringBuilder
    Private mTransaccion As OleDbTransaction

    Sub main()
        Console.WriteLine("calculo")
        'Console.WriteLine(prueba(8).ToString)
        Console.WriteLine(pruebaTexto(2).ToString)
        Console.ReadLine()
    End Sub

    Private Function prueba(ByVal numero As Integer) As Integer
        Using cmdComando As New OleDbCommand("CalculadoraSuma", mConexion)
            cmdComando.CommandType = CommandType.StoredProcedure
            cmdComando.Parameters.AddWithValue("@dato", numero)
            cmdComando.Parameters.Add(New OleDbParameter("@respuesta", SqlDbType.Int))
            cmdComando.Parameters("@respuesta").Direction = ParameterDirection.Output
            Dim señal As Boolean = Open()
            If señal = True Then
                Return 0
            Else
                Call cmdComando.ExecuteNonQuery()
                Call Close()
                Return CInt(cmdComando.Parameters("@respuesta").Value)
            End If
        End Using
    End Function

    Private Function pruebaTexto(ByVal numero As Integer) As String
        Using cmdComando As New OleDbCommand("CalculadoraSuma", mConexion)
            cmdComando.CommandType = CommandType.StoredProcedure
            cmdComando.Parameters.AddWithValue("@dato", numero)
            cmdComando.Parameters.Add(New OleDbParameter("@respuesta", OleDbType.VarChar, 200))
            cmdComando.Parameters("@respuesta").Direction = ParameterDirection.Output
            'Dim señal As Boolean = Open()
            If True = Open() Then
                Return String.Empty
            Else
                Call cmdComando.ExecuteNonQuery()
                Call Close()
                Return cmdComando.Parameters("@respuesta").Value.ToString
            End If
        End Using
    End Function

#Region "Metodos privados"
    ''' <summary>
    ''' Abre la conexión con la base de datos
    ''' </summary>
    ''' <returns>Señal con el exito o fracaso de la operación</returns>
    Public Function Open() As Boolean
        Dim bError As Boolean = False
        Try
            If mConexion.State = ConnectionState.Closed Then
                mConexion.ConnectionString = mCadenaConexion
                mConexion.Open()
                Call ExecuteNoQuery("SET DATEFORMAT dmy")
            End If
        Catch ex As Exception
            bError = True
            mError.Append(Err.Number & ": " & Err.Description)
        End Try
        Return bError
    End Function

    ''' <summary>
    ''' Cierra la conexión con la base de datos
    ''' </summary>
    ''' <returns>Señal con el exito o fracaso de la operación</returns>
    Public Function Close() As Boolean
        Dim bError As Boolean = False
        Try
            If mConexion.State = ConnectionState.Open Then mConexion.Close()
        Catch ex As Exception
            bError = True
            mError.Append(Err.Number & ": " & Err.Description)
        End Try
        Return bError
    End Function

    ''' <summary>
    ''' Ejecura una consulta
    ''' </summary>
    ''' <param name="query">Consulta</param>
    ''' <returns></returns>
    Public Function ExecuteNoQuery(ByVal query As String) As Integer
        Dim cmdComando As New OleDbCommand
        Dim oRespuesta As Int32
        cmdComando.Connection = mConexion
        cmdComando.Transaction = mTransaccion
        cmdComando.CommandText = query
        oRespuesta = cmdComando.ExecuteNonQuery
        cmdComando = Nothing
        Return oRespuesta
    End Function
#End Region

#Region "Cacluladora suma"
    '    CREATE PROCEDURE CalculadoraSuma
    '	@dato integer,
    '	@respuesta integer output	
    'AS
    'BEGIN
    '	Set NOCOUNT On;
    '    Set @respuesta = @dato + 2
    'End
#End Region

#Region "Calculadora suma 2"
    '    ALTER PROCEDURE [dbo].[CalculadoraSuma]
    '	@dato integer,
    '	@respuesta varchar(200) output	
    'AS
    'BEGIN
    '	Set NOCOUNT On;
    '    Set @respuesta = 'La respuesta es ' +  cast(@dato + 2 as varchar(100))
    'End
#End Region
End Module