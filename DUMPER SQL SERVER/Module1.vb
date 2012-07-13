'Inclusion d'un namespace
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System
Imports System.Text
Imports Microsoft.VisualBasic.Strings
Imports System.Text.RegularExpressions








Module Module1


    Sub Main()
        Dim originalForegroundColor As ConsoleColor = Console.ForegroundColor
        Dim serveur, id, base, requete, sortie As String
        Console.Clear()
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("DUMPER SQL SERVER")
        Console.ForegroundColor = originalForegroundColor
        'Console.WriteLine(My.Application.CommandLineArgs.Count)
        If My.Application.CommandLineArgs.Count > 0 Then
            serveur = My.Application.CommandLineArgs(0)
            id = My.Application.CommandLineArgs(1)
            base = My.Application.CommandLineArgs(2)
            requete = My.Application.CommandLineArgs(3)
            sortie = My.Application.CommandLineArgs(4)
            Console.WriteLine("Serveur : " & serveur & " ID : " & id & " BASE : " & base & " requete : " & requete & " sortie : " & sortie)
            requetes(serveur, id, base, requete, sortie)
        Else
            Console.WriteLine("Veuillez indiquer des paramètres")
            Console.WriteLine("ex : ""server"" ""sa@pwd"" ""bddname"" ""select * from table;"" ""output.csv"" ")            
        End If
        End
    End Sub

    Public Sub requetes(ByVal serveur As String, ByVal id As String, ByVal base As String, ByVal requete As String, ByVal sortie As String)
        Dim ident = Split(id, "@")
        Dim strConnexion As String = "server=" & serveur & ";Database=" & base & ";User ID=" & ident(0) & ";Password=" & ident(1) & ";Trusted_Connection=False"
        Dim strRequete As String = requete
        Dim archivo As StreamWriter
        Dim i As Integer = 0
        Dim ligne, result As String
        Dim separator As String = ";"
        Dim counter = 0
        Dim strBuilder As New System.Text.StringBuilder
        archivo = File.CreateText(sortie)
        Try
            Dim source As String = UCase(requete)
            Dim find As String = ".*FROM"
            Dim findw As String = "WHERE*."
            result = Regex.Replace(source, find, "", RegexOptions.IgnoreCase)
            result = Regex.Replace(result, findw, "", RegexOptions.IgnoreCase)
            findw = "ORDER*.*"
            result = Regex.Replace(result, findw, "", RegexOptions.IgnoreCase)
            findw = "GROUP*.*"
            result = Regex.Replace(result, findw, "", RegexOptions.IgnoreCase)
            result = Replace(result, ";", "")
            result = Trim(result)
            'Console.WriteLine(result)
            Dim pb As New ConsoleProgressBar(getmax(result, strConnexion))
            Dim oConnection As New SqlConnection(strConnexion)
            Dim oCommand As New SqlCommand(strRequete, oConnection)
            oConnection.Open()
            Dim oReader As SqlDataReader = oCommand.ExecuteReader()
            Do
                'Console.WriteLine(ControlChars.Tab + "{0}" + ControlChars.Tab + "{1}", oReader.GetName(0), oReader.GetName(1))
                While oReader.Read()
                    strBuilder.Append("")
                    For i = 0 To (oReader.FieldCount - 1)
                        strBuilder.Append(oReader.GetValue(i) & separator)
                    Next
                    counter = counter + 1
                    pb.Update(counter)
                    'System.Threading.Thread.Sleep(10) ' Sleep for 1 second
                    ligne = strBuilder.ToString
                    archivo.WriteLine(ligne)
                    strBuilder.Remove(0, strBuilder.Length)

                    'Console.WriteLine(ControlChars.Tab + "{0}" + ControlChars.Tab + "{1}", oReader.GetValue(0), oReader.GetValue(1))
                End While
            Loop While oReader.NextResult()


            oReader.Close()
            oConnection.Close()
        Catch e As Exception
            Console.WriteLine(("L'erreur suivante a été rencontrée :" + e.Message))
            MsgBox((("L'erreur suivante a été rencontrée :" + e.Message)))
        End Try
        archivo.Close()
    End Sub 'Main    
    Public Function getmax(ByVal table As String, ByVal strConnexion As String) As Integer
        Dim strRequete As String = "SELECT COUNT(*) FROM " & table & ";"
        Dim oConnection As New SqlConnection(strConnexion)
        Dim oCommand As New SqlCommand(strRequete, oConnection)
        oConnection.Open()
        Dim oReader As SqlDataReader = oCommand.ExecuteReader()
        'Debug.Print(oReader.GetValue(0))
        While oReader.Read()
            getmax = oReader.GetValue(0)
        End While
        oReader.Close()
        oConnection.Close()
    End Function

End Module
