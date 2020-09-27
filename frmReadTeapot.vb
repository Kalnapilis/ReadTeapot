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
      Dim jsonObject As Newtonsoft.Json.Linq.JObject
      Dim jsonArray As JArray

      Try

         request = DirectCast(WebRequest.Create(
            "https://swd.weatherflow.com/swd/rest/observations/station/21592?api_key=20c70eae-e62f-4d3b-b3a4-8586e90f3ac8"),
            HttpWebRequest)
         response = DirectCast(request.GetResponse(), HttpWebResponse)
         reader = New StreamReader(response.GetResponseStream())

         RawResponse = reader.ReadToEnd()
         RawResponse = RawResponse

         jsonObject = Newtonsoft.Json.Linq.JObject.Parse(RawResponse)
         jsonArray = CType(jsonObject("obs"), JArray)
         'string jsonFormatted = JValue.Parse(json).ToString(Formatting.Indented)

         Dim o1 As JObject 'https://www.newtonsoft.com/json/help/html/QueryJson.htm
         o1 = jsonObject("status")
         Dim sm As String
         sm = o1("status_message")

         For Each item As JObject In jsonArray
            Debug.WriteLine(item.SelectToken("Last").ToString)
         Next

      Catch ex As Exception
         Console.WriteLine(ex.ToString)
         MsgBox(ex.ToString)
      Finally
         If Not response Is Nothing Then response.Close()
      End Try

   End Sub
End Class
