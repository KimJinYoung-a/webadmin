<%
Option Explicit

Dim MailDbConn
Set MailDbConn = Server.CreateObject("ADODB.Connection")
MailDbConn.Open "DSN=ThunderDB"

Dim strSQL
Dim MailTitles : MailTitles = "타이틀"
Dim MailConts : MailConts = "콘텐츠 "
Dim SenderMail : SenderMail = "mailzine@10x10.co.kr"
Dim SenderNm : SenderNm = "텐바이텐"
Dim ReceiverNm : ReceiverNm ="허진원"
Dim MailType : MailType ="22"
Dim strQuery : strQuery ="kobula@10x10.co.kr" 
Dim EmailDataType : EmailDataType = "string"
Dim DB_ID : DB_ID=""
  
strSQL= strSQL &_

	" INSERT INTO event_dbevent ( " &_
	" 	title, content " &_
	" 	, sender, sender_alias ,receiver_alias " &_
	"	, content_type, event_id, user_info " &_
	"	, email_insert_type, wasSended, email_data_type, email_sql, db_id) " &_
	" VALUES ( "&_
	" 	'" & MailTitles & "' , '" & replace(MailConts,"'","") & "' " &_
	" 	,'" & SenderMail & "' , '" & SenderNm & "' , '" & ReceiverNm & "' " &_
	" 	,'text/html', '" & MailType & "', '"& strQuery & "' " &_
	" 	,'new', 'X', '"& EmailDataType &"', '', ''" &_
	" ) "
	
MailDbConn.execute(strSQL)		
MailDbConn.close
set MailDbConn = Nothing

response.write strSQL
%>