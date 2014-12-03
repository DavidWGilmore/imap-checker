Imports ActiveUp.Net.Mail
Imports System


Module Main

    Sub Main()
        Dim imap = New Imap4Client
        Dim inbox As Mailbox
        Dim ids() As Integer
        Dim msg As Message
        Dim imap_server As String
        Dim imap_username As String
        Dim imap_password As String
        Dim imap_port As String
        Dim use_ssl As Boolean
        Dim debug_level As Integer
        Dim msg_from As String
        Dim subject As String
        Dim debug_msg As String
        Dim search_phrase As String
        Dim delete_msg As Boolean

        Try
            imap_server = My.Settings.IMAP_server
            imap_username = My.Settings.IMAP_user
            imap_password = My.Settings.IMAP_password
            imap_port = My.Settings.port
            use_ssl = My.Settings.use_ssl
            debug_level = My.Settings.debug_level
            msg_from = My.Settings.from
            subject = My.Settings.subject
            delete_msg = My.Settings.delete_msg


            If debug_level > 0 Then
                debug_msg = "Connecting to " & imap_server & " on port " & imap_port & " via "
                If use_ssl = True Then
                    debug_msg = debug_msg & "SSL"
                Else
                    debug_msg = debug_msg & "non SSL"
                End If
                Console.WriteLine(debug_msg)
            End If

            If use_ssl = True Then
                imap.ConnectSsl(imap_server, imap_port)
            Else
                imap.Connect(imap_server, imap_port)
            End If

            If debug_level > 0 Then
                Console.WriteLine("Connected")
            End If
            imap.LoginFast(imap_username, imap_password)
            If debug_level > 0 Then
                Console.WriteLine("Logged In")
            End If
            imap.Command("capability")
            inbox = imap.SelectMailbox("inbox")

            If debug_level > 0 Then
                Console.WriteLine("Inbox Selected")
            End If

            search_phrase = "UNSEEN"
            If msg_from <> "" Then
                search_phrase = search_phrase & " From """ & msg_from & """"
            End If
            If subject <> "" Then
                search_phrase = search_phrase & " Subject """ & subject & """"
            End If

            If debug_level > 0 Then
                Console.WriteLine("Searching for " & search_phrase)
            End If
            ids = inbox.Search(search_phrase)

            If debug_level > 0 Then
                Console.WriteLine(ids.Count.ToString & " unread messages")
            End If

            msg = inbox.Fetch.MessageObject(ids.Last)

            If debug_level > 0 Then
                Console.WriteLine("Subject: " & msg.Subject.ToString & " Sender: " & msg.From.ToString & msg.ReceivedDate.ToString)
            End If

            If delete_msg = True Then
                If debug_level > 0 Then
                    Console.WriteLine("Deleting message " & ids.Last.ToString)
                End If
                DeleteMessage(ids.Last.ToString, imap)
                If debug_level > 0 Then
                    Console.WriteLine("Message Deleted")
                End If
            End If
            Console.WriteLine(msg.ReceivedDate.ToLocalTime)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        
    End Sub

    Sub DeleteMessage(msgid As String, imap As Imap4Client)
        imap.Command("store " & msgid & " +flags \deleted")
        imap.Expunge()
    End Sub
End Module
