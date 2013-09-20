Module Main

    Function UrlEncodeEx(ByVal str As String)
        Dim UnreservedChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.~"
        Dim sb As New Text.StringBuilder
        Dim bytes = System.Text.Encoding.UTF8.GetBytes(str)
        For Each b In bytes
            If UnreservedChars.IndexOf(System.Convert.ToChar(b)) <> -1 Then
                sb.Append(System.Convert.ToChar(b))
            Else
                sb.AppendFormat("%{0:X2}", b)
            End If
        Next
        Return sb.ToString()
    End Function

    Function Get_HMACSHA1_Signature(ByVal key As String, ByVal src As String) As String
        Dim hs As New System.Security.Cryptography.HMACSHA1(System.Text.Encoding.ASCII.GetBytes(key))
        Dim hash = hs.ComputeHash(System.Text.Encoding.ASCII.GetBytes(src))
        Return System.Convert.ToBase64String(hash)
    End Function

    Function Get_BaseStringPair(ByRef valpair As SortedDictionary(Of String, String)) As String
        Dim sd As New SortedDictionary(Of String, String)(valpair)
        Dim values = ""
        For Each p In sd
            values += p.Key & "=" & p.Value & "&"
        Next
        values = values.Remove(values.Length - 1)
        Return values
    End Function

    Function Get_Header_Authorization(ByRef valpair As SortedDictionary(Of String, String)) As String
        Dim sd As New SortedDictionary(Of String, String)(valpair)
        Dim values = ""
        For Each p In sd
            values += p.Key & "=""" & p.Value & """, "
        Next
        values = values.Remove(values.Length - 2)
        Return values
    End Function

    Function Get_OAuth_Signature(ByRef req As Net.HttpWebRequest, ByRef valpair As SortedDictionary(Of String, String), ByVal cs As String, ByVal ts As String) As String
        Dim method = UrlEncodeEx(req.Method)
        Dim url = UrlEncodeEx(req.RequestUri.OriginalString)
        Dim values = Get_BaseStringPair(valpair)

        Dim encodedvalues = UrlEncodeEx(values)

        Dim target = System.String.Format("{0}&{1}&{2}", method, url, encodedvalues)
        Dim key = UrlEncodeEx(cs) & "&" & UrlEncodeEx(ts)

        Return UrlEncodeEx(Get_HMACSHA1_Signature(key, target))
    End Function

    Function Get_UNIXTime() As Integer
        Dim elapsedtime As TimeSpan = System.DateTime.Now.ToUniversalTime() - New DateTime(1970, 1, 1, 0, 0, 0, 0)
        Return elapsedtime.TotalSeconds
    End Function

    Function Get_RequestToken(ByVal ckey As String, ByVal cskey As String) As Dictionary(Of String, String)
        Dim timestamp As String = Get_UNIXTime()
        Dim nonce As String = System.DateTime.Now.Millisecond
        Dim signature_method = "HMAC-SHA1"
        Dim values As New SortedDictionary(Of String, String)
        values.Add("oauth_consumer_key", ckey)
        values.Add("oauth_timestamp", timestamp)
        values.Add("oauth_nonce", nonce)
        values.Add("oauth_signature_method", signature_method)
        values.Add("oauth_version", "1.0")

        Dim requrl = "https://api.twitter.com/oauth/request_token"
        Dim hreq = Net.HttpWebRequest.Create(requrl)
        hreq.Method = "POST"
        Dim signature = Get_OAuth_Signature(hreq, values, cskey, "")
        values.Add("oauth_signature", signature)

        Dim kvpair = Get_Header_Authorization(values)

        hreq.Headers.Add("Authorization", "OAuth " & kvpair)

        Dim hres = hreq.GetResponse()
        Dim strm = hres.GetResponseStream()
        Dim strmreader As New IO.StreamReader(strm)
        Dim result = strmreader.ReadToEnd()
        strmreader.Close()
        Return Get_Dictionary(result)
    End Function

    Function Get_Authorization_URL(ByVal rkey As String) As String
        Return "https://api.twitter.com/oauth/authorize?oauth_token=" & rkey
    End Function

    Function Get_Dictionary(ByVal str As String) As Dictionary(Of String, String)
        Dim splited = str.Split("&")
        Dim dic As New Dictionary(Of String, String)
        For Each kv As String In splited
            Dim kvsplited = kv.Split("=")
            dic.Add(kvsplited(0), kvsplited(1))
        Next
        Return dic
    End Function

    Function Get_AccessToken_with_PIN(ByVal ckey As String, ByVal cskey As String, ByVal rkey As String, ByVal rskey As String, ByVal pin As String) As Dictionary(Of String, String)
        Dim values As New SortedDictionary(Of String, String)
        values("oauth_consumer_key") = ckey
        values("oauth_nonce") = System.DateTime.Now.Millisecond
        values("oauth_signature_method") = "HMAC-SHA1"
        values("oauth_timestamp") = Get_UNIXTime()
        values("oauth_version") = "1.0"
        values("oauth_token") = rkey
        values("oauth_verifier") = pin

        Dim hreq = Net.HttpWebRequest.Create("https://api.twitter.com/oauth/access_token")
        hreq.Method = "POST"
        Dim signature = Get_OAuth_Signature(hreq, values, cskey, rskey)

        values("oauth_signature") = signature
        hreq.Headers.Add("Authorization", "OAuth " & Get_Header_Authorization(values))

        Dim hres = hreq.GetResponse()
        Dim sr As New IO.StreamReader(hres.GetResponseStream())
        Dim result = sr.ReadToEnd()
        sr.Close()
        Return Get_Dictionary(result)
    End Function

    Function Send_Status(ByVal ckey As String, ByVal cskey As String, ByVal akey As String, ByVal askey As String, ByVal status As String, Optional ByVal id As String = Nothing) As Boolean
        Dim values As New SortedDictionary(Of String, String)
        If id IsNot Nothing Then
            values("in_reply_to_status_id") = id
        End If
        values("oauth_consumer_key") = ckey
        values("oauth_nonce") = System.DateTime.Now.Millisecond
        values("oauth_signature_method") = "HMAC-SHA1"
        values("oauth_timestamp") = Get_UNIXTime()
        values("oauth_version") = "1.0"
        values("oauth_token") = akey
        values("status") = UrlEncodeEx(status)

        Dim hreq = Net.HttpWebRequest.Create("https://api.twitter.com/1.1/statuses/update.json")
        hreq.Method = "POST"
        Dim signature = Get_OAuth_Signature(hreq, values, cskey, askey)
        values("oauth_signature") = signature

        Dim content = Get_BaseStringPair(values)
        hreq.ContentType = "application/x-www-form-urlencoded"
        hreq.ContentLength = content.Length
        Dim reqstrm = hreq.GetRequestStream()
        Dim sw = New IO.StreamWriter(reqstrm)
        sw.Write(content)
        sw.Close()

        Try
            Dim hres = hreq.GetResponse()
            Dim sr As New IO.StreamReader(hres.GetResponseStream())
            Dim result = sr.ReadToEnd()
            sr.Close()

            Dim jdc = Newtonsoft.Json.Linq.JObject.Parse(result)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Function Get_UserStream(ByVal ckey As String, ByVal cskey As String, ByVal akey As String, ByVal askey As String) As IO.Stream
        Dim values As New SortedDictionary(Of String, String)
        values("oauth_consumer_key") = ckey
        values("oauth_nonce") = System.DateTime.Now.Millisecond
        values("oauth_signature_method") = "HMAC-SHA1"
        values("oauth_timestamp") = Get_UNIXTime()
        values("oauth_version") = "1.0"
        values("oauth_token") = akey

        Dim hreq = Net.HttpWebRequest.Create("https://userstream.twitter.com/1.1/user.json")
        hreq.Method = "POST"
        Dim signature = Get_OAuth_Signature(hreq, values, cskey, askey)
        values("oauth_signature") = signature

        Dim content = Get_BaseStringPair(values)
        hreq.ContentType = "application/x-www-form-urlencoded"
        hreq.ContentLength = content.Length
        Dim reqstrm = hreq.GetRequestStream()
        Dim sw = New IO.StreamWriter(reqstrm)
        sw.Write(content)
        sw.Close()

        Dim hres = hreq.GetResponse()
        Dim strm = hres.GetResponseStream()

        Return strm
    End Function

    Sub Connect_UserStream(ByVal ckey As String, ByVal cskey As String, ByVal akey As String, ByVal askey As String, ByVal act As Action(Of Newtonsoft.Json.Linq.JObject))

        Dim strm = Get_UserStream(ckey, cskey, akey, askey)

        Dim buffer(1024) As Byte
        Dim strbuilder As New Text.StringBuilder

        While True
            Try
                Dim length = strm.Read(buffer, 0, 1024)
                Dim strbuffer As String = System.Text.Encoding.UTF8.GetChars(buffer, 0, length)
                If strbuffer.Contains(vbCrLf) Then
                    Dim pos As Integer = strbuffer.IndexOf(vbCr)
                    strbuilder.Append(strbuffer.Substring(0, pos))

                    If Not String.IsNullOrEmpty(strbuilder.ToString) AndAlso Not String.IsNullOrWhiteSpace(strbuilder.ToString) Then
                        Dim jobj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(strbuilder.ToString())
                        act(jobj)
                    End If

                    strbuilder.Remove(0, strbuilder.Length)
                    strbuilder.Append(strbuffer.Substring(pos + 2, strbuffer.Length - (pos + 2)))
                Else
                    strbuilder.Append(strbuffer)
                End If
            Catch e As IO.IOException
                strm = Get_UserStream(ckey, cskey, akey, askey)
            End Try
        End While

    End Sub

    Function Get_AccountSettings(ByVal ckey As String, ByVal cskey As String, ByVal akey As String, ByVal askey As String) As Newtonsoft.Json.Linq.JObject
        Dim values As New SortedDictionary(Of String, String)
        values("oauth_consumer_key") = ckey
        values("oauth_nonce") = DateTime.Now.Millisecond
        values("oauth_signature_method") = "HMAC-SHA1"
        values("oauth_timestamp") = Get_UNIXTime()
        values("oauth_version") = "1.0"
        values("oauth_token") = akey

        Dim hreq As Net.HttpWebRequest = Net.HttpWebRequest.Create("https://api.twitter.com/1.1/account/settings.json")
        hreq.Method = "POST"
        Dim signature = Get_OAuth_Signature(hreq, values, cskey, askey)
        values("oauth_signature") = signature

        Dim content = Get_BaseStringPair(values)
        hreq.ContentType = "application/x-www-form-urlencoded"
        'hreq.Headers("Content-Length") = content.Length
        hreq.ContentLength = content.Length

        Dim reqstrm = New IO.StreamWriter(hreq.GetRequestStream())
        reqstrm.Write(content)
        reqstrm.Close()


        Dim hres As Net.HttpWebResponse = hreq.GetResponse()
        Dim sr As New IO.StreamReader(hres.GetResponseStream())
        Dim result = sr.ReadToEnd()
        sr.Close()
        Dim jdc = Newtonsoft.Json.Linq.JObject.Parse(result)

        Return jdc
    End Function

    Function Get_BusInfoString(ByVal dt As DateTime, ByVal max As Integer) As String
        Dim str As String = ""
        Dim buslist = GetBusInfo(dt, max)
        str += vbCrLf + Now.ToString("MM月dd日 dddd HH時mm分 (JST)") + vbCrLf
        If buslist.Count = 0 Then
            str += "今日のバスはもうありません。"
        Else
            For Each b In buslist
                str += String.Format("{0:00}:{1:00} {2} {3}" + vbCrLf, b.Leaving.Hours, b.Leaving.Minutes, b.BusNumber, b.Destination)
            Next
        End If
        Return str
    End Function

    Sub Main()
        Dim ckey As String = ""
        Dim cskey As String = ""
        Dim akey As String = ""
        Dim askey As String = ""

        LoadTimetable("timetable.xml")

        Dim tweetlog As New Logger("tweet.log")
        Dim actlog As New Logger("action.log")

        If Not IO.File.Exists("settings.txt") Then
            IO.File.Create("settings.txt").Close()
        End If

        Dim sr As New IO.StreamReader("settings.txt")
        Dim settings = Get_Dictionary(sr.ReadToEnd())
        sr.Close()
        If Not settings.ContainsKey("ConsumerKey") Then
            Console.WriteLine("Please put 'ConsumerKey' and 'ConsumerSecretKey' into settings.txt.")
            Return
        End If
        ckey = settings("ConsumerKey")
        cskey = settings("ConsumerSecretKey")
        If settings.ContainsKey("AccessKey") Then
            akey = settings("AccessKey")
            askey = settings("AccessSecretKey")
        End If


        If (String.IsNullOrEmpty(akey)) Then
            Dim reqdir = Get_RequestToken(ckey, cskey)

            Dim url = Get_Authorization_URL(reqdir("oauth_token"))

            Console.WriteLine(url)
            Console.Write("PIN:")
            Dim PIN = Console.ReadLine()

            Dim actokenpair = Get_AccessToken_with_PIN(ckey, cskey, reqdir("oauth_token"), reqdir("oauth_token_secret"), PIN)

            akey = actokenpair("oauth_token")
            askey = actokenpair("oauth_token_secret")

            Dim sw As New IO.StreamWriter("settings.txt", False)
            sw.Write(String.Format("AccessKey={0}&AccessSecretKey={1}&ConsumerKey={2}&ConsumerSecretKey={3}", akey, askey, ckey, cskey))
            sw.Close()
        End If

        Dim scrname As String = Get_AccountSettings(ckey, cskey, akey, askey)("screen_name")

        Dim th As New System.Threading.Thread(
            Sub()
                Connect_UserStream(ckey, cskey, akey, askey,
                                   Sub(jobj As Newtonsoft.Json.Linq.JObject)
                                       If jobj("created_at") IsNot Nothing And jobj("event") Is Nothing Then
                                           tweetlog.WriteLog(
                                               String.Format("{0}[{1}] {2} ({3})",
                                                             jobj("user")("screen_name"),
                                                             jobj("user")("name"),
                                                             jobj("text").ToString().Replace(vbLf, "[LF]"),
                                                             jobj("created_at")
                                                             )
                                                         )
                                           If Text.RegularExpressions.Regex.IsMatch(jobj("text").ToString(), "^@" + scrname) Then
                                               Dim str As String = "@" + jobj("user")("screen_name").ToString + " "
                                               str += Get_BusInfoString(Now, 5)
                                               Send_Status(ckey, cskey, akey, askey, str, jobj("id"))
                                               actlog.WriteLog(String.Format("Send: {0} (in_reply_to: {1})", str.Replace(vbCrLf, "[CRLF]"), jobj("id")))
                                           End If
                                       End If
                                   End Sub)
            End Sub)

        For i As Integer = 0 To 13
            Dim tt As New IntervalTask
            tt.act = (Sub()
                          Dim str = Get_BusInfoString(Now, 5)
                          Send_Status(ckey, cskey, akey, askey, str)
                      End Sub)
            tt.begintime = New DateTime(2013, 8, 20, 14 + i / 2, 30 * (i Mod 2), 0)
            tt.interval = New TimeSpan(24, 0, 0)
            AddIntervalTask(tt)
        Next

        th.Start()

        While th.IsAlive
            DoTask()
            Threading.Thread.Sleep(1)
        End While
    End Sub

End Module
