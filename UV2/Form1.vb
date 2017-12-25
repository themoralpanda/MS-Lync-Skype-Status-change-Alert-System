Imports System
Imports model = Microsoft.Lync.Model
Imports System.Net
Imports System.Web
Imports System.IO
Imports System.Text

'The below application will provide following features.
' 1. Connect to Lync
' 2. To enable searching for contacts
' 3. To tag a person and provide a target status
' 4. To trigger an SMS to the specified number if the target person reaches the respective tagged status

' authors: Vigneshwar Venkatachalapathi, Ajeeth Kumar
' Email: Vigneshwar_V@infosys.com, Ajeeth_kumar@infosys.com
' Emp#: 647830, 647821
' Copyright: Infosys Ltd.

Public Class LyncSMS
    'Importing Kernel32.dll for using the clock within VB.net
	Declare Sub Sleep Lib "kernel32.dll" (ByVal Milliseconds As Integer)
#Region "Declarations"
    Dim lyncClient As model.LyncClient = model.LyncClient.GetClient 'Gets the Lync client
    Dim gI As Integer = Nothing
    Dim dict As New Dictionary(Of String, String)
    Dim contact As model.Contact = lyncClient.Self.Contact 'Get the Lync contact interface
    Dim resultList As ArrayList
    Dim resultsCount As Integer = Nothing
    Dim targetFixed As String = Nothing
    Dim check As DialogResult
    Dim taglist1 As ArrayList = New ArrayList
    Dim taglist2 As ArrayList = New ArrayList
	
	'Callbacks for the event driven operations
	Dim callback As AsyncCallback
    Dim sChangeCallback As AsyncCallback
#End Region

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load 'Main Form load event
        Label1.ForeColor = Color.White
        Label1.Text = lyncClient.Self.Contact.GetContactInformation(model.ContactInformationType.DisplayName)
        Dim myStatus As String = contact.GetContactInformation(model.ContactInformationType.ActivityId)
        ComboBox1.DropDownHeight = ComboBox1.Height * 10
        Label3.BackColor = Color.White
        
		Select Case myStatus 'case validation for Lync status specific colors
            Case "Away"
                Label3.ForeColor = Color.Yellow
                Label3.Text = myStatus
            Case "Free"
                Label3.ForeColor = Color.Green
                Label3.Text = "Available"
            Case "DoNotDisturb"
                Label3.ForeColor = Color.Red
                Label3.Text = myStatus
            Case "BeRightBack"
                Label3.ForeColor = Color.YellowGreen
                Label3.Text = myStatus
        End Select
        Me.AutoSize = True
        GroupBox1.AutoSize = True
    End Sub

    Private Sub Search(key As String)  'Module for searching contacts
        MsgBox("Searching for contacts on" & key)
        callback = AddressOf searchHandler
        Me.TopMost = False
        lyncClient.ContactManager.BeginSearch(key, callback, Nothing)
    End Sub

    Private Sub tagName(ByVal key As String) 'Module for tagging a particular person
        sChangeCallback = AddressOf searchHandlerSpecific
        lyncClient.ContactManager.BeginSearch(key, sChangeCallback, Nothing)
    End Sub

    Private Sub searchHandlerSpecific(result As IAsyncResult) 'Callback for Tagging
		'this function checks the status change for every 5 seconds and sends an SMS if an status of the tagged person changes.
        Dim res As model.SearchResults = lyncClient.ContactManager.EndSearch(result)
        If res.Contacts.Count = 1 Then
            Do Until String.Compare(res.Contacts.Item(0).GetContactInformation(model.ContactInformationType.Activity).ToString, dict.Item(res.Contacts.Item(0).GetContactInformation(model.ContactInformationType.DisplayName).ToString())) = 0
                Sleep(5000) 'To check the status change for every 5 seconds.  
            Loop
            MsgBox("The status has changed to " & " " & dict.Item(res.Contacts.Item(0).GetContactInformation(model.ContactInformationType.DisplayName).ToString()))
            SendSMS(res.Contacts.Item(0).GetContactInformation(model.ContactInformationType.Activity).ToString)
        Else
            MsgBox("some problem with the search")
        End If
    End Sub

    Private Sub searchHandler(result As IAsyncResult) 'Callback for populating the contact search list.
        'This function search for the contact name based on the key specified and update the list of contacts in the combo box.
		Dim res As model.SearchResults = lyncClient.ContactManager.EndSearch(result)
        If res.Contacts.Count > 0 Then          
            resultsCount = res.Contacts.Count
            For Each c As model.Contact In res.Contacts
                UpdateComboBox(c.GetContactInformation(model.ContactInformationType.DisplayName).ToString(), c.GetContactInformation(model.ContactInformationType.Activity).ToString())                
            Next
        Else
            MsgBox("some problem")
        End If
    End Sub
	
    Private Sub UpdateComboBox(ByVal contact As String, ByVal status As String) 'Module for updating combo box with the result list of names that matches the key
        If Me.InvokeRequired Then
            Dim args() As String = {contact, status}
            Me.Invoke(New Action(Of String, String)(AddressOf UpdateComboBox), args)
            Return
        End If
        ComboBox1.Items.Add(contact & "- " & status)
    End Sub


    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
	'Search button event handler
        If String.IsNullOrEmpty(TextBox1.Text) = False Then
            ComboBox1.Items.Clear()
            Dim key = TextBox1.Text
            TextBox1.Enabled = False

            Search(key)

            TextBox1.Enabled = True
            MsgBox("please check.. process over !")
            Dim length = 0
            For Each item As String In ComboBox1.Items
                length = Math.Max(length, item.Length)
            Next
            ComboBox1.Width = length * CType(Me.ComboBox1.Font.SizeInPoints, Integer)

        Else
            MsgBox("Please give some key to serch buddy! ")
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If String.IsNullOrEmpty(targetFixed) = True Then
            targetFixed = ComboBox1.SelectedItem
        Else
            check = MessageBox.Show("Have you selected " & ComboBox1.SelectedItem & " as the target", "Selected ?", MessageBoxButtons.YesNo)
            If check = Windows.Forms.DialogResult.Yes Then
                Dim label As Label = New Label
                Dim caption = ComboBox1.SelectedItem.ToString.Split(" ").GetValue(0) & " " & ComboBox1.SelectedItem.ToString.Split(" ").GetValue(1)
                label.Text = caption
                taglist1.Add(ComboBox1.SelectedItem.ToString)
                Dim myfont As New Font("Sans Serif", 10, FontStyle.Regular)
                label.Font = myfont
				
                Dim cbox As ComboBox = New ComboBox
                cbox.Items.Add("Available")
                cbox.Items.Add("Away")
                cbox.Items.Add("Busy")               
                cbox.SelectedIndex = 1
                AddHandler cbox.SelectedIndexChanged, AddressOf Me.cboxHandler
                AddHandler cbox.SelectionChangeCommitted, AddressOf Me.cboxHandler2

                Dim closebutton As Button = New Button
                TableLayoutPanel1.Controls.Add(closebutton)
                TableLayoutPanel1.Controls.Add(cbox)
                TableLayoutPanel1.Controls.Add(label)
                Me.Refresh()

            End If
        End If
    End Sub

    Private Sub cboxHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
        taglist2.Add(CType(CType(sender, System.Windows.Forms.ComboBox).SelectedItem, String))
    End Sub
    Private Sub cboxHandler2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If String.IsNullOrEmpty(CType(CType(sender, System.Windows.Forms.ComboBox).SelectedItem, String)) = True Then
            MsgBox("please choose the target status")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim temp As String = Nothing
        If taglist2.Count <> 0 And taglist1.Count = taglist2.Count Then
            Dim t1 As Array = taglist1.ToArray
            Dim t2 As Array = taglist2.ToArray

            For i As Integer = 0 To taglist1.Count - 1
                gI = i
                dict.Add(t1(i).Split("-")(0), t2(i))
                tagName(CType(t1(i).Split("-")(0), String))
            Next
        Else
            MsgBox("some problem")
        End If
    End Sub

    Private Sub SendSMS(ByVal msg As String) 'Function to send an SMS once the status changes.
        Dim client As WebClient = New WebClient
        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)")
        client.QueryString.Add("user", "a********")
        client.QueryString.Add("password", "*************")
        client.QueryString.Add("api_id", "*****")
        client.QueryString.Add("to", "+919**989***")
        client.QueryString.Add("text", "StatusHasBeenChanged " & msg)
        Dim baseurl As String = "http://api.clickatell.com/http/sendmsg"
        Dim data As Stream = client.OpenRead(baseurl)
        Dim reader As StreamReader = New StreamReader(data)
        Dim s As String = reader.ReadToEnd()
        MsgBox(s)
        data.Close()
        reader.Close()
    End Sub
End Class
