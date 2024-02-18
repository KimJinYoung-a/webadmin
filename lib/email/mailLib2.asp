<%
'###########################################################
' Description : ���Ϲ߼�
' History : ������ ����
'			2020.09.28 �ѿ�� ����(���Ϸ� ����)
'###########################################################

CLASS MailCls

	dim MailTitles		'���� ����
	dim MailConts		'���� ���� 			(text/html)
	dim SenderMail		'���� �߼��� �ּ� 	(customer@10x10.co.kr,mailzine@10x10.co.kr)
	dim SenderNm		'���� �߼����̸� 	(�ٹ�����)

	dim MailType		'���ø� ��ȣ 		([4],5,6,7,8,9)

	dim ReceiverNm		'���� ������ �̸� 	($1)
	dim ReceiverMail	'���� ������ �ּ� 	(xxxx@aaa.com..)


	dim AddrType				'���ϼ��� ��� (event,userid)
	dim arrUserId 				'AddrType ="userid" �ϰ�� ���

	dim AddrString				'�����ּ� ������ ���� ����
	dim EvtCode,EvtGroupCode 	'AddrType ="event" �ϰ�� ���
	dim MailerMailGubun		' ���Ϸ� �ڵ����� ��ȣ

	dim strQuery 		'�̸��� ���� ���� ����
	dim EmailDataType	'�̸��� ���� ���� ��� (Enum : string - ���� �Է�,sql - ���� �̿�)
	Dim DB_ID 			'�������� ��񿬰� ��ȣ - ���� (�Ǽ���- 4 ; �׽�Ʈ- 5)


	Private Sub Class_Initialize()
		EvtCode =0
		EvtGroupCode =0
		EmailDataType = "sql"
		MailType = 5
		MailerMailGubun = 2		' ���Ϸ� �ڵ����� ��ȣ �⺻�߼� 2��

		IF application("Svr_Info")="Dev" THEN
			DB_ID = "5" '//(�Ǽ���- 4 ; �׽�Ʈ- 5)
		ELSE
			DB_ID = "4"
		END IF
		SenderMail	= "mailzine@10x10.co.kr"
		SenderNm	= "�ٹ�����"

	End Sub

	Private Sub Class_Terminate()

	End Sub

	'// mailer ���� �ּ� ��ȯ
	Public Function fnMakeMailerQuery()

		dim tmpAddrType , tmpString
		dim tmpVar

		dim tmpQuery

		On Error Resume Next

		tmpAddrType = AddrType
		tmpString = AddrString

		tmpVar = fnReArr(tmpString,",")

		IF tmpAddrType = "userid" THEN '// ���̵� �迭 �Է�

			IF tmpVar = "" Then
				Err.Number = Err.Number - 1
			End IF

			tmpVar = replace(tmpVar,",","','")
			tmpVar = "'" & tmpVar & "'"

            ''���̸��Ϸ� ������ ���̻� ��� ����.
			''tmpQuery = " SELECT UMail, UName FROM db_user.dbo.vw_UserMailList WHERE Uid in (" & tmpVar & ")"
			Err.Number = Err.Number - 1  ''������

		ELSEIF tmpAddrType ="string" Then '// �̸� �ּ� �ϳ��� ó��

			IF ReceiverMail="" Then
				Err.Number = Err.Number - 1
			End IF

			EmailDataType ="string"

			tmpQuery = ReceiverMail

		ELSEIF tmpAddrType ="array" Then '// �̸� & �ּ� �Է� ��Ģ�Է�
			'// �۾�����(�����߰�)
			tmpQuery=""
		ELSE
			'// �۾�����
			tmpQuery=""
		End IF

		fnMakeMailerQuery = tmpQuery

	End Function

	'// cdo ���� �ּ� ��ȯ

	Public Function fnMakeCdoQuery(byref iArr)

		dim tmpAddrType , tmpString

		dim tmpVar , tmpVar2 , intLp

		dim tmpQuery,tmpArrList()

		On Error Resume Next

		tmpAddrType = AddrType
		tmpString = AddrString

		tmpVar = fnReArr(tmpString,",")

		IF tmpAddrType = "userid" THEN '// ���̵� �迭 �Է�

			IF tmpVar = "" Then
				Err.Number = Err.Number - 1
			End IF

			tmpVar = replace(tmpVar,",","','")
			tmpVar = "'" & tmpVar & "'"

			tmpQuery = " SELECT UMail, UName FROM db_user.dbo.vw_UserMailList WHERE Uid in (" & tmpVar & ")"

			rsget.open tmpQuery , dbget , 2
			IF not rsget.eof Then
				iArr = rsget.getRows()
			End IF
			rsget.close

		ELSEIF tmpAddrType ="string" Then '// �̸� �ּ� �ϳ��� ó��

			IF ReceiverMail="" Then
				Err.Number = Err.Number - 1
			End IF

			Redim iArr(1,0)
			iArr(0,0) = ReceiverMail
			iArr(1,0) = ReceiverNm

		ELSEIF tmpAddrType ="array" Then '// �̸� & �ּ� �Է� ��Ģ�Է�
			IF tmpVar = "" Then
				Err.Number = Err.Number - 1
			End IF
			tmpVar = fnReArr(tmpVar,",")
			tmpVar = Split(tmpVar,",")

			IF isArray(tmpVar) Then

				Redim iArr(1,Ubound(tmpVar))
				For intLp=0 To Ubound(tmpVar)
					tmpVar2 = tmpVar(intLp)

					IF instr(tmpVar2,"$")>0 Then
						iArr(0,intLp) = Left(tmpVar2,instr(tmpVar2,"$")-1)
						iArr(1,intLp) = Right(tmpVar2,len(tmpVar2)-instr(tmpVar2,"$"))
					ELSE
						iArr(0,intLp) = tmpVar2
						iArr(1,intLp) = ""
					End IF
				Next
			End IF
		End IF
		IF Err.Number=0 Then
			fnMakeCdoQuery = 0
		ELSE
			fnMakeCdoQuery = -1
		End IF


	End Function

	Public Function fnReArr(byval strVar,byval strChk)

		'// �����ڷ� �Ѿ�� �� strChk üũ�� ��ȯ
		'// ��ȯ�� ���� "," �� ���е�

		dim tmpVar , tmpArrVar , intLp

		IF strVar="" or strChk="" Then '�Ѿ�� �� üũ (���ų� �߸��� ���� �Ѿ���� ������)
			Exit Function
		ELSE
			tmpArrVar = trim(strVar)
			tmpArrVar = split(tmpArrVar,strChk)

			IF Ubound(tmpArrVar) < 0 Then Exit Function

			For intLp=0 to Ubound(tmpArrVar)
				IF tmpArrVar(intLp)<>"" Then
					tmpVar = (tmpVar & tmpArrVar(intLp) & ",")
				END IF
			Next
			tmpVar = Left(tmpVar,Len(tmpVar)-1)
		END IF
		fnReArr = tmpVar

	End Function

	'//+++	���� ���ø� �ҷ����� 	+++//	' 2017.12.20 �ѿ��
	Public Function getMailTemplate()
		dim mFileNm, dfPath, fso,ffso,fnHTML
		dim mailheader, mailfooter

		'/* ���� ���� */
		'// MailType - 5 �̻� ���� ��� (�����ڿ� ����/���� ����! ��.�Ѥ� )
		IF MailType ="5" Then '// �Ļ�������� ��� ����
			mFileNm =""
		ELSEIF MailType="6" Then 		'// �ֹ�����
			mFileNm ="mail_a01.htm"
		ELSEIF MailType ="7" Then '// ����Ȯ��
			mFileNm ="mail_a02.htm"
		ELSEIF MailType ="8" Then '// ������
			'mFileNm = "mail_delivery2011.htm"
			mFileNm ="mail_delivery2017.html"
		ELSEIF MailType ="9" Then '// �������ڵ���Ҿȳ�
			mFileNm ="mail_a04.htm"

		ELSEIF MailType ="10" Then '// ��ŸCS���߼�
			mFileNm ="mail_b01.htm"
		ELSEIF MailType ="11" Then '// �ֹ����(ȯ�Ҿȳ�)
			mFileNm ="mail_b02.htm"
		ELSEIF MailType ="12" Then '// ��ǰ����
			mFileNm ="mail_b03.htm"
		ELSEIF MailType ="13" Then '// ��ǰ�Ϸ�(ȯ�Ҿȳ�)
			mFileNm ="mail_b04.htm"
		ELSEIF MailType ="14" Then '// ȯ��/ī����ҿϷ�
			mFileNm ="mail_b05.htm"

		ELSEIF MailType ="15" Then '// 1:1��� �亯
			'mFileNm ="mail_c01.htm"
			mFileNm ="mail_c01_new.html"
		ELSEIF MailType ="16" Then '// ��ǰQ&A �亯
			mFileNm ="mail_c02.htm"
		ELSEIF MailType ="17" Then '// �Ϲ� ���� ����
			'mFileNm ="mail_d01.htm"
			mFileNm ="mail_d01.html"
		ELSEIF MailType ="18" Then '// ��ǰ���ۼ��ȳ�
			mFileNm ="mail_d02.htm"
		ELSEIF MailType ="19" Then '// ȸ����ް���
			mFileNm ="mail_d03.htm"
		ELSEIF MailType ="20" Then '// �̺�Ʈ��÷����
			mFileNm ="mail_d06.htm"
		ELSEIF MailType ="21" Then '// ��й�ȣ��߼۸���
			mFileNm ="mail_d07.htm"
		ELSEIF MailType ="22" Then '// �����������
			'mFileNm ="mail_misend.htm"
			mFileNm ="email_misend.html"
		End IF

		IF MailType<>"5" and mFileNm="" Then
			response.write "���ø� �ҷ����� ����"
			Exit Function
		End IF

		'//�Ǽ�,�׼�����
		IF application("Svr_Info")="Dev" THEN
			'dfPath = "C:\testweb\admin2009scm\lib\email\mailtemplate" 		'// �׼�(scm)
			dfPath = Server.MapPath("\lib\email\mailtemplate")
		ELSE
		    dfPath = Server.MapPath("\lib\email\mailtemplate")
			''dfPath = "E:\home\cube1010\admin2009scm\lib\email\mailtemplate" 	'// �Ǽ�(scm)
		END IF

		'/* ���� �ҷ����� */
		IF mFileNm<>"" Then
			Set fso = server.CreateObject("Scripting.FileSystemObject")
				IF fso.FileExists(dfPath & "\" & mFileNm) then
					set ffso = fso.OpenTextFile(dfPath & "\" & mFileNm,1)
					fnHTML = ffso.ReadAll
					ffso.close
					set ffso = nothing
				ELSE
					fnHTML = ""
				End IF
			Set fso = nothing
		End IF

		'/�ű� ������ �������� ���ø� ����� Ǫ�Ϳ� ������ �и���. ���� �ٸ��ǵ鵵 ������� ���� �и��ϰ�, �� �Ϸ� �Ǹ� �б�ó�� ����.
		IF MailType ="17" or MailType ="22" Then
	        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
	        Set fso = Server.CreateObject("Scripting.FileSystemObject")
	        dfPath = server.mappath("\lib\email")

	        mFileNm = dfPath&"\\email_header_1.html"

	        Set ffso = fso.OpenTextFile(mFileNm,1)
	        mailheader = ffso.readall	' ���
		
	        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
	        Set fso = Server.CreateObject("Scripting.FileSystemObject")
	        dfPath = server.mappath("/lib/email")

	        mFileNm = dfPath&"\\email_footer_1.html"

	        Set ffso = fso.OpenTextFile(mFileNm,1)
	        mailfooter = ffso.readall	' Ǫ��

			fnHTML = mailheader & fnHTML & mailfooter
		End IF

		getMailTemplate = fnHTML
	End Function

    '//+++	TMS���Ϸ� ���Ϲ߼�	' 2020.09.29 �ѿ�� ����
    Public Function Send_TMSMailer()
        Dim sqlStr

		'response.write MailerMailGubun & "<br>"
		'response.write replace(ReceiverMail,"'","") & "<br>"
		'response.write replace(MailTitles,"'","") & "<br>"
		'response.write newhtml2db(MailConts) & "<br>"
		'response.end

        IF (AddrType<>"string") or (ReceiverMail="") Then '// �̸� �ּ� �ϳ��� ó��
		    Err.Number = Err.Number - 1
        ENd IF

        sqlStr =  " exec db_cs.dbo.usp_TEN_TMS_SendAutoMail '"&replace(ReceiverMail,"'","")&"','','"&replace(MailTitles,"'","")&"','"&newhtml2db(MailConts)&"',"& MailerMailGubun &""
        dbget.Execute sqlStr
    end Function

    '//+++	EMS ���� ���Ϸ� ���Ϲ߼� 2014/04/28	+++//
    Public Function Send_Mailer()
        Dim sqlStr

		'response.write MailerMailGubun & "<br>"
		'response.write replace(ReceiverMail,"'","") & "<br>"
		'response.write replace(MailTitles,"'","") & "<br>"
		'response.write newhtml2db(MailConts) & "<br>"
		'response.end

        IF (AddrType<>"string") or (ReceiverMail="") Then '// �̸� �ּ� �ϳ��� ó��
		    Err.Number = Err.Number - 1
        ENd IF

        sqlStr =  " exec db_cs.[dbo].[sp_Ten_SendAutoMail_Amailer] '"&replace(ReceiverMail,"'","")&"','','"&replace(MailTitles,"'","")&"','"&newhtml2db(MailConts)&"',"& MailerMailGubun &""
        dbget.Execute sqlStr
    end Function

	'//+++	������� ���Ϲ߼� 	+++//
	Public Function Send_Mailer_OLD_THUNDER()

		On Error Resume Next

		Dim MailDbConn
		Set MailDbConn = Server.CreateObject("ADODB.Connection")
			MailDbConn.Open "DSN=ThunderDB"

		Dim strSQL

		strQuery = fnMakeMailerQuery()

		IF strQuery="" Then
			response.write "����ڰ� �������� �ʽ��ϴ�"
		End IF

		strQuery = replace(strQuery,"'","''")

'response.write MailConts
'dbget.close()	:	response.End
		strSQL= strSQL &_

			" INSERT INTO event_dbevent ( " &_
			" 	title, content " &_
			" 	, sender, sender_alias ,receiver_alias " &_
			"	, content_type, event_id, user_info " &_
			"	, email_insert_type, wasSended, email_data_type, email_sql, db_id) " &_
			" VALUES ( "&_
			" 	'" & MailTitles & "' , '" & newhtml2db(MailConts) & "' " &_
			" 	,'" & SenderMail & "' , '" & SenderNm & "' , '" & ReceiverNm & "' " &_
			" 	,'text/html', '" & MailType & "', '"& strQuery & "' " &_
			" 	,'new', 'X', '"& EmailDataType &"', '" & strQuery & "', '"&DB_ID&"'" &_
			" ) "
''response.write strSQL

		MailDbConn.beginTrans
		MailDbConn.execute(strSQL)


		IF Err.Number =0 THEN
			MailDbConn.CommitTrans
			response.write "���� �߼� ����_Mailer<br>"
		ELSE
			MailDbConn.RollBackTrans
			response.write "���� �߼� ����_Mailer<br>"
		END IF

		MailDbConn.close

        On Error Goto 0
	End Function

	'//+++	�ܺ� ���� ���Ϲ߼� 	+++//

	Public Function Send_CDO()
		dim ArrMailList,intP,ret

		ret = fnMakeCdoQuery(ArrMailList)

		IF ret < 0 Then
			response.write "�ּ� ó�� ����"
			Exit Function
		End IF

		dim cdoMessage,cdoConfig

		'On Error Resume Next

		IF isArray(ArrMailList) Then
			For intP=0 To Ubound(ArrMailList,2)
				Set cdoConfig = Server.CreateObject("CDO.Configuration")
				'-> ���� ���ٹ���� �����մϴ�
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
				'-> ���� �ּҸ� �����մϴ�(dns or ip)-(localhost or 110.93.128.94)
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="110.93.128.94"
				'-> ������ ��Ʈ��ȣ�� �����մϴ�
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'-> ���ӽõ��� ���ѽð��� �����մϴ�
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 5
				'-> SMTP ���� ��������� �����մϴ�
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				'-> SMTP ������ ������ ID�� �Է��մϴ�
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"
				'-> SMTP ������ ������ ��ȣ�� �Է��մϴ�
				cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"
				cdoConfig.Fields.Update

				Set cdoMessage = CreateObject("CDO.Message")
				Set cdoMessage.Configuration = cdoConfig


				cdoMessage.To 		= ArrMailList(1,intP) &"<"& ArrMailList(0,intP) &">"
				cdoMessage.From 	= SenderNm &"<"& SenderMail &">"
				cdoMessage.SubJect 	= MailTitles

				'���� ������ �ؽ�Ʈ�� ��� cdoMessage.TextBody, html�� ��� cdoMessage.HTMLBody
				cdoMessage.HTMLBody	= MailConts

				cdoMessage.BodyPart.Charset="ks_c_5601-1987"		'// �ѱ��� ���ؼ� �� �־� �־�� �մϴ�.
				cdoMessage.HTMLBodyPart.Charset="ks_c_5601-1987"	'// �ѱ��� ���ؼ� �� �־� �־�� �մϴ�.
				'cdoMessage.BodyPart.Charset="utf-8"		'// �ѱ��� ���ؼ� �� �־� �־�� �մϴ�.
				'cdoMessage.HTMLBodyPart.Charset="utf-8"	'// �ѱ��� ���ؼ� �� �־� �־�� �մϴ�.

				IF (application("Svr_Info")	= "Dev") then
				    if ((InStr(ArrMailList(0,intP),"10x10.co.kr")>0) or (ArrMailList(0,intP)="kobula@nate.com") or (ArrMailList(0,intP)="tozzinet@daum.net")) then
    		            cdoMessage.Send
                    end if
                ELSE
				    cdoMessage.Send
                END IF

				Set cdoMessage = nothing
				Set cdoConfig = nothing
			Next
		End IF

		IF Err.Number =0 THEN
			response.write "���� �߼� ����_Send_CDO<br>"
		ELSE
			response.write "���� �߼� ����_Send_CDO<br>"
		END IF

	End Function


	'//+++	���� ���� ���Ϲ߼� 	+++//
	Public Function Send_CDONT()

	 	dim ArrMailList,intP,ret

		IF ReceiverMail="" THEN
			Exit Function
		END IF

		ret = fnMakeCdoQuery(ArrMailList)

		IF ret < 0 Then
			response.write "�ּ� ó�� ����"
			Exit Function
		End IF

		dim oCDONT

        'On Error Resume Next

		IF isArray(ArrMailList) Then
			For intP=0 To Ubound(ArrMailList,2)
				Set oCDONT=Server.CreateObject("CDONTS.NewMail")
				oCDONT.to 		= ArrMailList(1,intP) &"<"& ArrMailList(0,intP) &">"
				oCDONT.from 	= SenderNm &"<"& SenderMail &">"
				oCDONT.subject 	= MailTitles
				'html style
				oCDONT.bodyformat = 0
				oCDONT.mailformat = 0

				oCDONT.body = MailConts
				oCDONT.send
				Set oCDONT = Nothing
			Next
		End IF

		IF Err.Number =0 THEN
			MailDbConn.CommitTrans
			response.write "���� �߼� ����_Send_CDONT<br>"
		ELSE
			MailDbConn.RollBackTrans
			response.write "���� �߼� ����_Send_CDONT<br>"
		END IF
	End Function

End CLASS

'// �ܼ� ���� �߼� ����
Function fnSendMail(mailto,title,contents)

	Dim objMail

	Set objMail = New MailCls

	objMail.AddrType="string"
	objMail.ReceiverMail = mailto
	objMail.MailTitles = title
	objMail.MailConts = contents
	objMail.Send_CDO()

	Set objMail = Nothing

End Function

Function fnSendMail_Mailer(mailto,title,contents)

	Dim objMail

	Set objMail = New MailCls

	objMail.AddrType="string"
	objMail.ReceiverMail = mailto
	objMail.MailTitles = title
	objMail.MailConts = contents
	objMail.MailerMailGubun = 2		' ���Ϸ� �ڵ����� ��ȣ
	objMail.Send_TMSMailer()		'TMS���Ϸ�
	'objMail.Send_Mailer()

	Set objMail = Nothing
End Function


%>
