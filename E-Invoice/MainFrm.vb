Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports System.Text
Imports System.IO
Imports E_Invoice.SilentWebModule
Imports HtmlAgilityPack
Imports System.Threading

Public Class MainFrm
    Public cookies As New CookieContainer()
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                    (ByVal hwnd As IntPtr, _
                     ByVal wMsg As Integer, _
                     ByVal wParam As IntPtr, _
                     ByVal lParam As Byte()) _
                     As Integer
    Public Const EM_SETCUEBANNER As Integer = &H1501
    Dim E_InvoiceThread As Thread
    Dim ThreadCheck As Boolean = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            E_InvoiceThread = New Thread(AddressOf Me.E_InvoiceBackground)
            E_InvoiceThread.Start()
        Catch ex As Exception

        End Try
    End Sub
    Sub E_InvoiceBackground()
        Try
            ThreadCheck = True
            Button1.Enabled = False
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False

            Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
            Dim response As HttpWebResponse = HttpWebResponseUtility.CreateGetHttpResponse(TextBox2.Text, Nothing, Nothing, cookies)
            Dim reader As StreamReader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            Dim respHTML As String = reader.ReadToEnd()

            If respHTML.Contains("系統錯誤") Then
                MsgBox("會員載具歸戶網址輸入錯誤 !", MsgBoxStyle.Critical)
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox1.Enabled = True
                TextBox2.Enabled = True
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                CatchCaptcha()
                ThreadCheck = False
                response.Close()
                Exit Sub
            End If

            Dim doc As New HtmlDocument()
            doc.LoadHtml(respHTML)
            Dim node As HtmlNode = doc.DocumentNode

            Dim __VIEWSTATE = WebUtility.HtmlDecode(node.SelectNodes("//*[@id=""__VIEWSTATE""]")(0).Attributes.Item("value").Value)
            Dim __EVENTVALIDATION = WebUtility.HtmlDecode(node.SelectNodes("//*[@id=""__EVENTVALIDATION""]")(0).Attributes.Item("value").Value)
            Dim txtMemAccount = WebUtility.HtmlDecode(node.SelectNodes("//*[@id=""txtMemAccount""]")(0).Attributes.Item("value").Value)
            Dim txtCarryType = WebUtility.HtmlDecode(node.SelectNodes("//*[@id=""txtCarryType""]")(0).Attributes.Item("value").Value)

            parameters.Add("__VIEWSTATE", System.Uri.EscapeDataString(__VIEWSTATE))
            parameters.Add("__EVENTVALIDATION", System.Uri.EscapeDataString(__EVENTVALIDATION))
            parameters.Add("txtMemAccount", System.Uri.EscapeDataString(txtMemAccount))
            parameters.Add("txtCarryType", System.Uri.EscapeDataString(txtCarryType))
            parameters.Add("txtCheckCode", TextBox1.Text)
            parameters.Add("_btnQuery", System.Uri.EscapeDataString("確定歸戶"))
            response = HttpWebResponseUtility.CreatePostHttpResponse("https://www.bpscm.com.tw/ASP/APMember/MemberSameOwnerToMOFWeb.aspx?c=" & System.Uri.EscapeDataString(TextBox2.Text.Split("?")(1).Substring(2)), parameters, Nothing, Nothing, Encoding.UTF8, cookies)

            reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            respHTML = reader.ReadToEnd()
            doc.LoadHtml(respHTML)
            node = doc.DocumentNode
            Dim card_ban = Nothing
            Dim card_no1 = Nothing
            Dim card_no2 = Nothing
            Dim card_type = Nothing
            Dim back_url = Nothing
            Dim token = Nothing

            If respHTML.Contains("驗證碼錯誤") Then
                MsgBox("驗證碼錯誤", MsgBoxStyle.Exclamation)
                TextBox1.Clear()
                TextBox1.Enabled = True
                TextBox2.Enabled = True
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                CatchCaptcha()
                ThreadCheck = False
                response.Close()
                Exit Sub
            Else
                card_ban = WebUtility.HtmlDecode(node.SelectNodes("//*[@name=""card_ban""]")(0).Attributes.Item("value").Value)
                card_no1 = WebUtility.HtmlDecode(node.SelectNodes("//*[@name=""card_no1""]")(0).Attributes.Item("value").Value)
                card_no2 = WebUtility.HtmlDecode(node.SelectNodes("//*[@name=""card_no2""]")(0).Attributes.Item("value").Value)
                card_type = WebUtility.HtmlDecode(node.SelectNodes("//*[@name=""card_type""]")(0).Attributes.Item("value").Value)
                back_url = WebUtility.HtmlDecode(node.SelectNodes("//*[@name=""back_url""]")(0).Attributes.Item("value").Value)
                token = WebUtility.HtmlDecode(node.SelectNodes("//*[@name=""token""]")(0).Attributes.Item("value").Value)
            End If

            parameters.Clear()
            parameters.Add("card_ban", System.Uri.EscapeDataString(card_ban))
            parameters.Add("card_no1", System.Uri.EscapeDataString(card_no1))
            parameters.Add("card_no2", System.Uri.EscapeDataString(card_no2))
            parameters.Add("card_type", System.Uri.EscapeDataString(card_type))
            parameters.Add("back_url", System.Uri.EscapeDataString(back_url))
            parameters.Add("token", System.Uri.EscapeDataString(token))
            response = HttpWebResponseUtility.CreatePostHttpResponse("https://www.einvoice.nat.gov.tw/APMEMBERVAN/membercardlogin", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
            reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            respHTML = reader.ReadToEnd()
            response.Close()

            doc.LoadHtml(respHTML)
            node = doc.DocumentNode

            Dim CSRT = Nothing
            For i = 0 To respHTML.Split(vbNewLine).Count - 1
                If respHTML.Split(vbNewLine)(i).Contains("CSRT") Then
                    CSRT = respHTML.Split(vbNewLine)(i).Substring(respHTML.Split(vbNewLine)(i).IndexOf("var pv = '") + 10, respHTML.Split(vbNewLine)(i).IndexOf("';") - respHTML.Split(vbNewLine)(i).IndexOf("var pv = '") - 10)
                    Exit For
                End If
            Next

            parameters.Add("vo.contactName", "")
            parameters.Add("moicaSerialNo", "")
            parameters.Add("card_code", System.Uri.EscapeDataString(node.SelectNodes("//*[@name=""card_code""]")(0).Attributes.Item("value").Value))
            parameters.Add("CSRT", CSRT)
            response = HttpWebResponseUtility.CreatePostHttpResponse("https://www.einvoice.nat.gov.tw/APMEMBERVAN/membercardlogin!publicCarrierLogin", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
            response.Close()

            parameters.Add("mobile", System.Uri.EscapeDataString(TextBox3.Text))
            parameters.Add("verifyCode", System.Uri.EscapeDataString(TextBox4.Text))
            response = HttpWebResponseUtility.CreatePostHttpResponse("https://www.einvoice.nat.gov.tw/APMEMBERVAN/membercardlogin!queryIDNByPublicCarrier", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
            reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            respHTML = reader.ReadToEnd()
            response.Close()

            If respHTML.Contains("手機或驗證碼錯誤") Then
                MsgBox("手機或驗證碼錯誤", MsgBoxStyle.Exclamation)
                TextBox1.Clear()
                TextBox1.Enabled = True
                TextBox2.Enabled = True
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                CatchCaptcha()
                ThreadCheck = False
                response.Close()
                Exit Sub
            End If

            doc.LoadHtml(respHTML)
            node = doc.DocumentNode

            parameters.Add("ind", "")
            parameters.Add("cardName", System.Uri.EscapeDataString(WebUtility.HtmlDecode(node.SelectNodes("//*[@name=""cardName""]")(0).Attributes.Item("value").Value)))
            parameters.Add("idn", System.Uri.EscapeDataString(WebUtility.HtmlDecode(node.SelectNodes("//*[@name=""idn""]")(0).Attributes.Item("value").Value)))
            parameters.Add("carrierMode", System.Uri.EscapeDataString(WebUtility.HtmlDecode(node.SelectNodes("//*[@name=""carrierMode""]")(0).Attributes.Item("value").Value)))
            response = HttpWebResponseUtility.CreatePostHttpResponse("https://www.einvoice.nat.gov.tw/APMEMBERVAN/membercardlogin!carrierMgtComplete", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
            reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            respHTML = reader.ReadToEnd()

            For i = 0 To respHTML.Split(vbNewLine).Count - 1
                If respHTML.Split(vbNewLine)(i).Contains("alert") Then
                    MsgBox(respHTML.Split(vbNewLine)(i).Substring(respHTML.Split(vbNewLine)(i).IndexOf("alert('") + 7, respHTML.Split(vbNewLine)(i).IndexOf("');") - respHTML.Split(vbNewLine)(i).IndexOf("alert('") - 7))
                    TextBox1.Clear()
                    TextBox2.Clear()
                    TextBox1.Enabled = True
                    TextBox2.Enabled = True
                    TextBox3.Enabled = True
                    TextBox4.Enabled = True
                    CatchCaptcha()
                    response.Close()
                    ThreadCheck = False
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox("暫時無法使用 , 請稍後再試 !", MsgBoxStyle.Exclamation)
            TextBox1.Clear()
            TextBox1.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            ThreadCheck = False
            CatchCaptcha()
        End Try
    End Sub
    Sub CatchCaptcha()
        Try
            Dim response As HttpWebResponse = HttpWebResponseUtility.CreateGetHttpResponse("https://www.bpscm.com.tw/ASP/APMember/Image.aspx", Nothing, Nothing, cookies)
            Dim reader As Stream = response.GetResponseStream()
            Dim captcha As Bitmap = New Bitmap(reader)
            CaptchaPic.Image = captcha
            response.Close()
            Exit Sub
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MainFrm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            E_InvoiceThread.Abort()
        Catch ex As Exception

        End Try
        End
    End Sub

    Private Sub MainFrm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Form.CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub MainFrm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        SendMessage(TextBox2.Handle, _
                     EM_SETCUEBANNER, _
                     IntPtr.Zero, _
                     System.Text.Encoding.Unicode.GetBytes("會員載具歸戶網址"))
        SendMessage(TextBox3.Handle, _
                     EM_SETCUEBANNER, _
                     IntPtr.Zero, _
                     System.Text.Encoding.Unicode.GetBytes("手機號碼"))
        SendMessage(TextBox4.Handle, _
                     EM_SETCUEBANNER, _
                     IntPtr.Zero, _
                     System.Text.Encoding.Unicode.GetBytes("驗證碼"))
        SendMessage(TextBox1.Handle, _
                     EM_SETCUEBANNER, _
                     IntPtr.Zero, _
                     System.Text.Encoding.Unicode.GetBytes("圖片驗證碼"))
        CatchCaptcha()

        ' Set up ToolTip. 
        ToolTip.AutoPopDelay = 5000
        ToolTip.InitialDelay = 1
        ToolTip.ReshowDelay = 1
        ToolTip.ShowAlways = True
        ToolTip.UseAnimation = True
        ToolTip.UseFading = True

        ' Set up the ToolTip text for the Button and Checkbox. 
        ToolTip.SetToolTip(Me.TextBox2, "Ex : https://www.bpscm.com.tw/ASP/APMember/MemberSameOwnerToMOFWeb.aspx?c=*")
    End Sub

    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged
        If ThreadCheck = True Then : Exit Sub : End If

        If TextBox1.Text.Length <> 5 Or Not TextBox2.Text.Contains("https://www.bpscm.com.tw/ASP/APMember/MemberSameOwnerToMOFWeb.aspx?c=") Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub
End Class

Namespace SilentWebModule
    ''' <summary>
    ''' 有關HTTP請求的模組
    ''' </summary>
    Public Class HttpWebResponseUtility
        Private Shared ReadOnly DefaultUserAgent As String = "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/33.0.1750.146 Safari/537.36"
        ''' <summary>
        ''' 創建GET方式的HTTP請求
        ''' </summary>
        ''' <param name="url">請求的URL</param>
        ''' <param name="timeout">請求的超時時間</param>
        ''' <param name="userAgent">請求的客戶端瀏覽器信息，可以為空</param>
        ''' <param name="cookies">隨同HTTP請求發送的Cookie信息，如果不需要身分驗證可以為空</param>
        ''' <returns></returns>
        Public Shared Function CreateGetHttpResponse(url As String, timeout As System.Nullable(Of Integer), userAgent As String, cookies As CookieContainer) As HttpWebResponse
            If String.IsNullOrEmpty(url) Then
                Throw New ArgumentNullException("url")
            End If
            Dim request As HttpWebRequest = TryCast(WebRequest.Create(url), HttpWebRequest)
            request.Method = "GET"
            request.KeepAlive = True
            request.UserAgent = DefaultUserAgent
            If Not String.IsNullOrEmpty(userAgent) Then
                request.UserAgent = userAgent
            End If
            If timeout.HasValue Then
                request.Timeout = timeout.Value
            End If
            If cookies IsNot Nothing Then
                request.CookieContainer = cookies
                'request.CookieContainer = New CookieContainer()
                'request.CookieContainer.Add(cookies)
            End If
            Return TryCast(request.GetResponse(), HttpWebResponse)
        End Function
        ''' <summary>
        ''' 創建POST方式的HTTP請求
        ''' </summary>
        ''' <param name="url">請求的URL</param>
        ''' <param name="parameters">隨同請求POST的參數名稱及參數值字典</param>
        ''' <param name="timeout">請求的超時時間</param>
        ''' <param name="userAgent">請求的客戶端瀏覽器信息，可以為空</param>
        ''' <param name="requestEncoding">發送HTTP請求時所用的編碼</param>
        ''' <param name="cookies">隨同HTTP請求發送的Cookie信息，如果不需要身分驗證可以為空</param>
        ''' <returns></returns>
        Public Shared Function CreatePostHttpResponse(url As String, parameters As IDictionary(Of String, String), timeout As System.Nullable(Of Integer), userAgent As String, requestEncoding As Encoding, cookies As CookieContainer) As HttpWebResponse
            If String.IsNullOrEmpty(url) Then
                Throw New ArgumentNullException("url")
            End If
            If requestEncoding Is Nothing Then
                Throw New ArgumentNullException("requestEncoding")
            End If
            Dim request As HttpWebRequest = Nothing
            '如果是發送HTTPS請求
            If url.StartsWith("https", StringComparison.OrdinalIgnoreCase) Then
                ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
                request = TryCast(WebRequest.Create(url), HttpWebRequest)
                request.ProtocolVersion = HttpVersion.Version11
            Else
                request = TryCast(WebRequest.Create(url), HttpWebRequest)
            End If
            request.Method = "POST"
            request.KeepAlive = True
            request.ContentType = "application/x-www-form-urlencoded"


            If Not String.IsNullOrEmpty(userAgent) Then
                request.UserAgent = userAgent
            Else
                request.UserAgent = DefaultUserAgent
            End If

            If timeout.HasValue Then
                request.Timeout = timeout.Value
            End If
            If cookies IsNot Nothing Then
                request.CookieContainer = cookies
                'request.CookieContainer = New CookieContainer()
                'request.CookieContainer.Add(cookies)
            End If
            '如果需要POST數據
            If Not (parameters Is Nothing OrElse parameters.Count = 0) Then
                Dim buffer As New StringBuilder()
                Dim i As Integer = 0
                For Each key As String In parameters.Keys
                    If i > 0 Then
                        buffer.AppendFormat("&{0}={1}", key, parameters(key))
                    Else
                        buffer.AppendFormat("{0}={1}", key, parameters(key))
                    End If
                    i += 1
                Next
                Dim data As Byte() = requestEncoding.GetBytes(buffer.ToString())
                Using stream As Stream = request.GetRequestStream()
                    stream.Write(data, 0, data.Length)
                End Using
            End If
            Return TryCast(request.GetResponse(), HttpWebResponse)

        End Function

        Private Shared Function CheckValidationResult(sender As Object, certificate As X509Certificate, chain As X509Chain, errors As SslPolicyErrors) As Boolean
            Return True
            '總是接受
        End Function
    End Class
End Namespace
