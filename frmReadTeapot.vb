Option Strict On
Option Explicit On
Imports System.IO
Imports System.Net
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class frmReadTeapot
   Private Sub frmReadTeapot_Load(sender As Object, e As EventArgs) Handles MyBase.Load
      Dim request As HttpWebRequest
      Dim response As HttpWebResponse = Nothing
      Dim reader As StreamReader
      Dim RawResponse As String
      Dim jO As Object

      Try

         request = DirectCast(WebRequest.Create(
            "https://swd.weatherflow.com/swd/rest/observations/station/21592?api_key=20c70eae-e62f-4d3b-b3a4-8586e90f3ac8"),
            HttpWebRequest)
         response = DirectCast(request.GetResponse(), HttpWebResponse)
         reader = New StreamReader(response.GetResponseStream())

         RawResponse = reader.ReadToEnd()
         RawResponse = RawResponse

         jO = JObject.Parse(RawResponse)

      Catch ex As Exception
         Console.WriteLine(ex.ToString)
         MsgBox(ex.ToString)
      Finally
         If Not response Is Nothing Then response.Close()
      End Try

   End Sub
End Class
