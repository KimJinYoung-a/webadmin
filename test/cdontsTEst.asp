<%

sub sendmail(mailfrom, mailto, mailtitle, mailcontent)
        dim mailobject

        set mailobject=server.createobject("CDONTS.NewMail")
        mailobject.from = mailfrom
        mailobject.to = mailto
        mailobject.subject = mailtitle

        'html style
        mailobject.bodyformat = 0
        mailobject.mailformat = 0

        mailobject.body = mailcontent
        mailobject.send
        set mailobject = nothing
end sub

sub SendMailCDO(mailfrom, mailto, mailtitle, mailcontent)
        dim mailobject
        dim cdoMessage,cdoConfig
        
    ''On Error Resume Next    
        Set cdoConfig = CreateObject("CDO.Configuration")

		'-> ���� ���ٹ���� �����մϴ�
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)

		'-> ���� �ּҸ� �����մϴ�
		if (application("Svr_Info")	= "Dev") then
		    cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "61.252.133.2"
		else
    	    cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "110.93.128.95"
        end if
    
		'-> ������ ��Ʈ��ȣ�� �����մϴ�
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

		'-> ���ӽõ��� ���ѽð��� �����մϴ�
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10

		'-> SMTP ���� ��������� �����մϴ�
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

		'-> SMTP ������ ������ ID�� �Է��մϴ�
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"

		'-> SMTP ������ ������ ��ȣ�� �Է��մϴ�
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"

		cdoConfig.Fields.Update

		Set cdoMessage = CreateObject("CDO.Message")

		Set cdoMessage.Configuration = cdoConfig

		cdoMessage.To 				= mailto
		cdoMessage.From 			= mailfrom
		cdoMessage.SubJect 	= mailtitle
		'���� ������ �ؽ�Ʈ�� ��� cdoMessage.TextBody, html�� ��� cdoMessage.HTMLBody
		cdoMessage.HTMLBody	= mailcontent

		cdoMessage.BodyPart.Charset="ks_c_5601-1987"         '/// �ѱ��� ���ؼ� �� �־� �־�� �մϴ�.
        cdoMessage.HTMLBodyPart.Charset="ks_c_5601-1987"     '/// �ѱ��� ���ؼ� �� �־� �־�� �մϴ�.
        
        if (application("Svr_Info")	= "Dev") then
            ''�׽�Ʈ ȯ��
    		''if ((InStr(mailto,"10x10.co.kr")>0) or (mailto="archilee@shinbiro.com")) then
    		    cdoMessage.Send
           '' end if
        else
		    cdoMessage.Send
		end if

		Set cdoMessage = nothing
		Set cdoConfig = nothing
		
	''On Error Goto 0	

end sub

'call sendmail("mailzine@10x10.co.kr","kobula@10x10.co.kr","Ÿ��Ʋ","������")

'call SendMailCDO("mailzine@10x10.co.kr","kobula@10x10.co.kr","Ÿ��Ʋ1","������1")
%>