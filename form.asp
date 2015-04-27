<%

' Retrieve values from form fields and set as variables
formname = Request.Form("name")
formemail = Request.Form("email")
formquery = Request.Form("query")

' Create the AspEmail message object
Set Mail = Server.CreateObject("Persits.MailSender")

' Set the from Name and E-Mail address using values retrieved from the form
Mail.From = formemail
Mail.FromName = formname

' Add the e-mail recipient address - replace values within the quotes with your own
Mail.AddAddress "you@yoursite.co.uk"

' Set the subject for the e-mail
Mail.Subject = "Form submitted from web site"

' Create a string called bodytxt and build it line by line using values from the form
Bodytxt = "Details of Form submission :" & VbCrLf & VbCrLf
Bodytxt = Bodytxt & "Contact Name : " & formname & VbCrLf
Bodytxt = Bodytxt & "E-Mail Address : " & formemail & VbCrLf
Bodytxt = Bodytxt & "Query Entered : " & formquery

' Set body text for the e-mail to the Bodytxt string we built
Mail.Body = Bodytxt

' The mail server requires that we authenticate so supply username and password
Mail.Username = "me@mysite.co.uk"
Mail.Password = "password"

' The e-mail is now ready to go, we just need to specify the server and send
Mail.Host = "smtp.dotnetted.co.uk"
Mail.Send

' Mail is sent - tidy up and delete the AspEmail message object
Set Mail = Nothing

%>