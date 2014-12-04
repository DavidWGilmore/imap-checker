Imports ActiveUp.Net.Mail
Imports System


Module Main
    'When developing the application, having the debugging lines was useful to see where the code was getting hung up. For instance
    'sometimes the authentication is slow, sometimes searching for messages is slow. Enabling debugging will show you where the bottleneck is

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

            'All settings are pulled from the app.config file. See comments in the file for allowed vaules and what each setting does
            imap_server = My.Settings.IMAP_server
            imap_username = My.Settings.IMAP_user
            imap_password = My.Settings.IMAP_password
            imap_port = My.Settings.port
            use_ssl = My.Settings.use_ssl
            debug_level = My.Settings.debug_level
            msg_from = My.Settings.from
            subject = My.Settings.subject
            delete_msg = My.Settings.delete_msg

            'If debugging is enabled, show the user which IMAP server we are connecting to and with what type of connection
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

            'If debugging is enabled, tell the user we are now connected
            If debug_level > 0 Then
                Console.WriteLine("Connected")
            End If

            'On some servers imap.login was very slow. imap.loginfast provided better results
            imap.LoginFast(imap_username, imap_password)

            'Let the yser know we have logged in
            If debug_level > 0 Then
                Console.WriteLine("Logged In")
            End If


            imap.Command("capability")
            inbox = imap.SelectMailbox("inbox")

            'Let the user know we have found the Inbox folder
            If debug_level > 0 Then
                Console.WriteLine("Inbox Selected")
            End If

            'Here we build the search phrase to return the message we want. UNSEEN returns all unread messages. If msg_from is blank, it will return messages from everyone
            'If te subject is blank, it will return any unread message regardless of subject
            search_phrase = "UNSEEN"
            If msg_from <> "" Then
                search_phrase = search_phrase & " From """ & msg_from & """"
            End If
            If subject <> "" Then
                search_phrase = search_phrase & " Subject """ & subject & """"
            End If

            'Show the user the search phrase we are using to find their message
            If debug_level > 0 Then
                Console.WriteLine("Searching for " & search_phrase)
            End If

            'Returns a list of messages matching the search criteria
            ids = inbox.Search(search_phrase)

            'Tell the user how many messages we found
            If debug_level > 0 Then
                Console.WriteLine(ids.Count.ToString & " unread messages")
            End If

            'We only want the most recent message
            msg = inbox.Fetch.MessageObject(ids.Last)

            'Show the user information about the newest message we found
            If debug_level > 0 Then
                Console.WriteLine("Subject: " & msg.Subject.ToString & " Sender: " & msg.From.ToString & msg.ReceivedDate.ToString)
            End If

            'Check to see if the user wants to delete the message we found
            If delete_msg = True Then

                'Show the user the index number of the message we are deleting
                If debug_level > 0 Then
                    Console.WriteLine("Deleting message " & ids.Last.ToString)
                End If
                'This is the line that actually deletes the message

                DeleteMessage(ids.Last.ToString, imap)

                If debug_level > 0 Then
                    Console.WriteLine("Message Deleted")
                End If
            End If

            'Return the date/time stamp (in Local Time) of the message we found.
            Console.WriteLine(msg.ReceivedDate.ToLocalTime)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        
    End Sub

    Sub DeleteMessage(msgid As String, imap As Imap4Client)
        'IMAP does not have a Delete function built in. In order to achieve this, we need to flag that we want to delete the message and then Expunge the folder
        'During research, I found some people using a _move to Trashfolder_ ype approach, but the name of the folder, and how 
        'it is addressed varies based on a number of factors including the mail server type and the language of the system. Setting the flag proved
        'much more consistent
        imap.Command("store " & msgid & " +flags \deleted")
        imap.Expunge()
    End Sub
End Module
