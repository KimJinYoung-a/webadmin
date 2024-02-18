<%
'###########################################################
' Description : cs���� �̸��� �߼� ���� �Լ�
' History : �̻� ����
'			2017.12.21 �ѿ�� ����
'###########################################################

Function SendCsActionMail(id)

    dim oCsAction,strMailHTML,strMailTitle
	Set oCsAction = New CsActionMailCls
	strMailHTML = oCsAction.makeMailTemplate(id)

	if (oCsAction.FDivCD = "A008") then
		strMailTitle = "[�ٹ�����] "& oCsAction.FCustomerName & "���� ["& oCsAction.GetAsDivCDName &"] ó���� "& oCsAction.FCurrStateName &" �Ǿ����ϴ�."
	else
		strMailTitle = "[�ٹ�����] "& oCsAction.FCustomerName & "�Բ��� ��û�Ͻ� ["& oCsAction.GetAsDivCDName &"] ó���� "& oCsAction.FCurrStateName &" �Ǿ����ϴ�."
	end if


	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	IF oCsAction.FBuyEmail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oCsAction.FCustomerName
		oMail.ReceiverMail	= oCsAction.FBuyEmail
		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

	'// �׽�Ʈ
'	set oMail = New MailCls
'
'	IF oCsAction.FBuyEmail<>"" THEN
'		oMail.MailTitles	= strMailTitle
'		oMail.SenderNm		= "�ٹ�����"
'		oMail.SenderMail	= "customer@10x10.co.kr"
'		oMail.AddrType		= "string"
'		oMail.ReceiverNm	= oCsAction.FCustomerName
'		oMail.ReceiverMail	= "headab@naver.com"    ''oCsAction.FBuyEmail
'		oMail.MailConts 	= strMailHTML
'		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
'		oMail.Send_TMSMailer()		'TMS���Ϸ�
'		'oMail.Send_Mailer()
'
'	End IF
'
'	SET oMail = nothing

    Set oCsAction = Nothing

End Function

Function SendCsActionMail_GiftCard(id)

    dim oCsAction,strMailHTML,strMailTitle
	Set oCsAction = New CsActionMailCls
	strMailHTML = oCsAction.makeMailTemplate_GiftCard(id)

	if (oCsAction.FDivCD = "A008") then
		strMailTitle = "[�ٹ�����]"& oCsAction.FCustomerName & "���� ["& oCsAction.GetAsDivCDName &"] ó���� "& oCsAction.FCurrStateName &" �Ǿ����ϴ�."
	else
		strMailTitle = "[�ٹ�����]"& oCsAction.FCustomerName & "�Բ��� ��û�Ͻ� ["& oCsAction.GetAsDivCDName &"] ó���� "& oCsAction.FCurrStateName &" �Ǿ����ϴ�."
	end if

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	IF oCsAction.FBuyEmail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oCsAction.FCustomerName
		oMail.ReceiverMail	= oCsAction.FBuyEmail
		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

	'// �׽�Ʈ
'	set oMail = New MailCls
'
'	IF oCsAction.FBuyEmail<>"" THEN
'		oMail.MailTitles	= strMailTitle
'		oMail.SenderNm		= "�ٹ�����"
'		oMail.SenderMail	= "customer@10x10.co.kr"
'		oMail.AddrType		= "string"
'		oMail.ReceiverNm	= oCsAction.FCustomerName
'		oMail.ReceiverMail	= "headab@naver.com"    ''oCsAction.FBuyEmail
'		oMail.MailConts 	= strMailHTML
'		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
'		oMail.Send_TMSMailer()		'TMS���Ϸ�
'		'oMail.Send_Mailer()
'
'	End IF
'
'	SET oMail = nothing

    Set oCsAction = Nothing

End Function

function ReSendCsActionMail(id, iForceCurrState, iForceBuyEmail)
    dim oCsAction,strMailHTML,strMailTitle
	Set oCsAction = New CsActionMailCls
	if (iForceCurrState<>"") then
        oCsAction.FRectForceCurrState = iForceCurrState
    end if

    if (iForceBuyEmail<>"") then
        oCsAction.FRectForceBuyEmail = iForceBuyEmail
    end if

	strMailHTML = oCsAction.makeMailTemplate(id)

	if (oCsAction.FDivCD = "A008") then
		strMailTitle = "[�ٹ�����]"& oCsAction.FCustomerName & "���� ["& oCsAction.GetAsDivCDName &"] ó���� "& oCsAction.FCurrStateName &" �Ǿ����ϴ�."
	else
		strMailTitle = "[�ٹ�����]"& oCsAction.FCustomerName & "�Բ��� ��û�Ͻ� ["& oCsAction.GetAsDivCDName &"] ó���� "& oCsAction.FCurrStateName &" �Ǿ����ϴ�."
	end if

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	IF oCsAction.FBuyEmail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oCsAction.FCustomerName
		oMail.ReceiverMail	= oCsAction.FBuyEmail
		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
	End IF

	SET oMail = nothing

	'// �׽�Ʈ
''	set oMail = New MailCls
''
''	IF oCsAction.FBuyEmail<>"" THEN
''		oMail.MailTitles	= strMailTitle
''		oMail.SenderNm		= "�ٹ�����"
''		oMail.SenderMail	= "customer@10x10.co.kr"
''		oMail.AddrType		= "string"
''		oMail.ReceiverNm	= oCsAction.FCustomerName
''		oMail.ReceiverMail	= "headab@naver.com"    ''oCsAction.FBuyEmail
''		oMail.MailConts 	= strMailHTML
'		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
'		oMail.Send_TMSMailer()		'TMS���Ϸ�
''		'oMail.Send_Mailer()
''
''	End IF
''
''	SET oMail = nothing

    Set oCsAction = Nothing
end function

'/�������� : [CS]��۰���>>�������Ʈ NEW, ��ü���� : �ֹ� ���� > ����� ����Ʈ	' 2017.12.20 �ѿ��
function SendMiChulgoMailWithMessage(idx, mailmessage)
    ''require /lib/classes/cscenter/oldmisendcls.asp
	''require /lib/classes/order/upchebeasongcls.asp
    dim oneMisend
    dim strMailHTML,strMailTitle, contentsHtml
	strMailHTML = ""
	strMailTitle = "[�ٹ�����] ��� ���� �ȳ������Դϴ�."

    set oneMisend = new COldMiSend
    oneMisend.FRectDetailIDx = idx
	oneMisend.FRectForMail = "Y"
    oneMisend.getOneOldMisendItem

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oneMisend.FOneItem.FBuyname
		oMail.ReceiverMail	= oneMisend.FOneItem.FBuyEmail
		oMail.MailType = "22"
		strMailHTML = oMail.getMailTemplate
		''parsing
		strMailHTML = replace(strMailHTML,":mailtitle:", "������� �ȳ�����")		' �̸�������

		strMailHTML = replace(strMailHTML,":ORDERSERIAL:",oneMisend.FOneItem.Forderserial)
		strMailHTML = replace(strMailHTML,":ITEMIMAGEURL:",oneMisend.FOneItem.Fsmallimage)
		strMailHTML = replace(strMailHTML,":ITEMID:",oneMisend.FOneItem.Fitemid)
		strMailHTML = replace(strMailHTML,":ITEMNAME:",oneMisend.FOneItem.Fitemname)

		if oneMisend.FOneItem.Fitemoptionname<>"" then
			strMailHTML = replace(strMailHTML,":ITEMOPTIONNAME:","["&oneMisend.FOneItem.Fitemoptionname&"]")
		else
			strMailHTML = replace(strMailHTML,":ITEMOPTIONNAME:","")
		end if

		strMailHTML = replace(strMailHTML,":ITEMCNT:",oneMisend.FOneItem.Fitemcnt)
		strMailHTML = replace(strMailHTML,":COMPANYNAME:",oneMisend.FOneItem.getDlvCompanyName)
		strMailHTML = replace(strMailHTML,":MAYSENDDATE:",oneMisend.FOneItem.FMisendipgodate)

		if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
	        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�� ������ �ش� �Ǹ��ڰ� ���Բ� �����帮�� �����Դϴ�.<br>*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		else
	        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		end if

		if (oneMisend.FOneItem.FMisendipgodate<>"") then
			contentsHtml = nl2br(mailmessage)
			contentsHtml = Replace(contentsHtml, "\n", "<br>")

			if (GetMichulgoMailTitleString(oneMisend.FOneItem.FMisendReason) <> "") then
				oMail.MailTitles = GetMichulgoMailTitleString(oneMisend.FOneItem.FMisendReason)
			end if
		end if
		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)

		oMail.MailConts 	= strMailHTML
		'response.write strMailHTML
		'response.end
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()
		'oMail.Send_Mailer()
		'oMail.Send_CDO
	End IF

    ''�޸� ����.
    contentsHtml = replace(contentsHtml,"�߼ۿ�����","�߼ۿ�����("&oneMisend.FOneItem.FMisendipgodate&")")
	Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)

	SET oMail = nothing
	set oneMisend = Nothing
end function

' [��ü]�Խ��� �亯 �ۼ� �̸��� �߼�	' 2020-03-12 �ѿ�� ����
Function SendUpheBoardMail(idx)
	dim boardqna

	if idx="" or isnull(idx) then exit Function

	set boardqna = New CUpcheQnADetail
		boardqna.FRectIdx = idx
		boardqna.read()

    dim oCsAction,strMailHTML,strMailTitle
	Set oCsAction = New CsActionMailCls
	strMailHTML = oCsAction.makeMailTemplate_UpheBoard(idx)
	strMailTitle = "[�ٹ�����] "& boardqna.Fusername &"�� ��ü�Խ��ǿ� �����Ͻ� �ۿ� �亯�� ��� �Ǿ����ϴ�."

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	IF boardqna.femail<>"" THEN
		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= boardqna.Fusername
		oMail.ReceiverMail	= boardqna.femail
		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 13		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()		'EMS���Ϸ�
		''oMail.Send_CDO()
	End IF

	SET oMail = nothing
	set boardqna = nothing
    Set oCsAction = Nothing
End Function

' �����ϴµ�		'/ [CS]������>>[CS]��ó��CS����Ʈ
function SendMiChulgoMail_CS(csdetailidx)
    ''require /lib/classes/cscenter/cs_mifinishcls.asp
    dim oneMisend
    dim strMailHTML,strMailTitle, contentsHtml
	strMailHTML = ""
	strMailTitle = "[�ٹ�����] CS��� ���� �ȳ������Դϴ�."

    set oneMisend = new CCSMifinishMaster
        oneMisend.FRectCSDetailIDx = csdetailidx
        oneMisend.getOneMifinishItem

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oneMisend.FOneItem.FBuyname
		oMail.ReceiverMail	= oneMisend.FOneItem.FBuyEmail
		oMail.MailType = "22"
		strMailHTML = oMail.getMailTemplate
		''parsing
		strMailHTML = replace(strMailHTML,":mailtitle:", "������� �ȳ�����")		' �̸�������

		strMailHTML = replace(strMailHTML,":ORDERSERIAL:",oneMisend.FOneItem.Forderserial)
		strMailHTML = replace(strMailHTML,":ITEMIMAGEURL:",oneMisend.FOneItem.Fsmallimage)
		strMailHTML = replace(strMailHTML,":ITEMID:",oneMisend.FOneItem.Fitemid)
		strMailHTML = replace(strMailHTML,":ITEMNAME:",oneMisend.FOneItem.Fitemname)

		if oneMisend.FOneItem.Fitemoptionname<>"" then
			strMailHTML = replace(strMailHTML,":ITEMOPTIONNAME:","["&oneMisend.FOneItem.Fitemoptionname&"]")
		else
			strMailHTML = replace(strMailHTML,":ITEMOPTIONNAME:","")
		end if

		strMailHTML = replace(strMailHTML,":ITEMCNT:",oneMisend.FOneItem.FRegItemNo)
		strMailHTML = replace(strMailHTML,":COMPANYNAME:",oneMisend.FOneItem.getDlvCompanyName)
		strMailHTML = replace(strMailHTML,":MAYSENDDATE:",oneMisend.FOneItem.FMifinishipgodate)

		if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
		    if (oneMisend.FOneItem.FMifinishReason<>"07") then
		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�� ������ �ش� �Ǹ��ڰ� ���Բ� �����帮�� �����Դϴ�.<br>*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		    else
		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�� ������ �ش� �Ǹ��ڰ� ���Բ� �����帮�� �����Դϴ�.")
		    end if
		else
		    if (oneMisend.FOneItem.FMifinishReason<>"07") then
		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		    else
		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","")
		    end if
		end if

		if (oneMisend.FOneItem.FMifinishipgodate<>"") then
    		if (oneMisend.FOneItem.FMifinishReason="03") then
    		    if (oneMisend.FOneItem.getMifinishDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.   ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� ��û�Ͻ� ��ǰ�� �߼��� ������ �����Դϴ�.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			else
        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			end if
        			contentsHtml = contentsHtml & "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
        			contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
        			contentsHtml = contentsHtml & "���ο� ������ �帰 �� �������� ����帮��, ������� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.<br>"
    		    else
        		    contentsHtml = "�ȳ��ϼ���.   ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� ��û�Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>"
        			contentsHtml = contentsHtml & "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
        			contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"

    		    end if
    		elseif (oneMisend.FOneItem.FMifinishReason="02") then
    		    if (oneMisend.FOneItem.getMifinishDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� ��û�Ͻ� ��ǰ�� �ֹ� �� ����(����)�Ǵ� ��ǰ����<br>"
        			contentsHtml = contentsHtml & "�Ϲݻ�ǰ�� �޸� �ֹ�����(����)�� �Ⱓ�� �ҿ�Ǵ� ��ǰ�Դϴ�.<br>"
        			contentsHtml = contentsHtml & "�Ʒ��� ���� �߼ۿ������� �ȳ��ص帮����,<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			else
        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			end if
    		    else
        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� ��û�Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>"
        			contentsHtml = contentsHtml & "�Ʒ��� ���� �߼ۿ������� �ȳ��� �帳�ϴ�.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			else
        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			end if
    			end if
    	    elseif (oneMisend.FOneItem.FMifinishReason="04") then
    	        oMail.MailTitles = "[�ٹ�����] CS��� ���� �ȳ������Դϴ�."

    		    if (oneMisend.FOneItem.getMifinishDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
                    contentsHtml = contentsHtml & "���Բ��� ��û�Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>�����ۻ�ǰ</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"


    		    else
        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
                    contentsHtml = contentsHtml & "���Բ��� ��û�Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>�����ۻ�ǰ</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"


    			end if
    		elseif (oneMisend.FOneItem.FMifinishReason="07") then
    	        oMail.MailTitles = "[�ٹ�����] CS��� ���� �ȳ������Դϴ�."

    		    if (oneMisend.FOneItem.getMifinishDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
                    contentsHtml = contentsHtml & "���Բ��� ��û�Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>���������</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"


    		    else
        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
                    contentsHtml = contentsHtml & "���Բ��� ��û�Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>���������</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"

    			end if
    		end if
		end if
		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)

		oMail.MailConts 	= strMailHTML
		'response.write strMailHTML
		'response.end
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
		'oMail.Send_CDO
	End IF

    ''�޸� ����.
    contentsHtml = replace(contentsHtml,"�߼ۿ�����","�߼ۿ�����("&oneMisend.FOneItem.FMifinishipgodate&")")
	call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)

	SET oMail = nothing
	set oneMisend = Nothing
end function

'/�������� ���� �������ʿ� ����. ���� ���� SendMiChulgoMailWithMessage ��ǰ� ����ϰ� ���ļ� ����Ұ�. ���ø��� ��������.
function SendMiChulgoMail_off(detailidx)
    dim oneMisend ,strMailHTML,strMailTitle, contentsHtml

	strMailHTML = ""
	strMailTitle = "[�ٹ�����] ��� ���� �ȳ������Դϴ�."

    set oneMisend = new cupchebeasong_list
    oneMisend.FRectDetailIDx = detailidx
    oneMisend.fOneOldMisendItem

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oneMisend.FOneItem.FBuyname
		oMail.ReceiverMail	= oneMisend.FOneItem.FBuyEmail
		oMail.MailType = "22"
		strMailHTML = oMail.getMailTemplate

		strMailHTML = replace(strMailHTML,":ORDERSERIAL:",oneMisend.FOneItem.forderno)
		strMailHTML = replace(strMailHTML,":ITEMIMAGEURL:","")	'/oneMisend.FOneItem.Fsmallimage
		strMailHTML = replace(strMailHTML,":ITEMID:",oneMisend.FOneItem.Fitemid)
		strMailHTML = replace(strMailHTML,":ITEMNAME:",oneMisend.FOneItem.Fitemname)
		strMailHTML = replace(strMailHTML,":ITEMCNT:",oneMisend.FOneItem.Fitemno)
		strMailHTML = replace(strMailHTML,":COMPANYNAME:",oneMisend.FOneItem.getDlvCompanyName)
		strMailHTML = replace(strMailHTML,":MAYSENDDATE:",oneMisend.FOneItem.FMisendipgodate)

		if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
		    strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�� ������ �ش� �Ǹ��ڰ� ���Բ� �����帮�� �����Դϴ�.<br>*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		else
		    strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		end if

		if (oneMisend.FOneItem.FMisendipgodate<>"") then
    		if (oneMisend.FOneItem.FMisendReason="03") then
    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.   ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� �߼��� ������ �����Դϴ�.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			else
        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			end if
        			contentsHtml = contentsHtml & "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
        			contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
        			contentsHtml = contentsHtml & "���ο� ������ �帰 �� �������� ����帮��, ������� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.<br>"
    		    else
        		    contentsHtml = "�ȳ��ϼ���.   ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>"
        			contentsHtml = contentsHtml & "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
        			contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
    		    end if
    		elseif (oneMisend.FOneItem.FMisendReason="02") then
    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� �ֹ� �� ����(����)�Ǵ� ��ǰ����<br>"
        			contentsHtml = contentsHtml & "�Ϲݻ�ǰ�� �޸� �ֹ�����(����)�� �Ⱓ�� �ҿ�Ǵ� ��ǰ�Դϴ�.<br>"
        			contentsHtml = contentsHtml & "�Ʒ��� ���� �߼ۿ������� �ȳ��ص帮����,<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			else
        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			end if
    		    else
        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>"
        			contentsHtml = contentsHtml & "�Ʒ��� ���� �߼ۿ������� �ȳ��� �帳�ϴ�.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			else
        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			end if
    			end if
    	    elseif (oneMisend.FOneItem.FMisendReason="04") then
    	        oMail.MailTitles = "[�ٹ�����] ��� ���� �ȳ������Դϴ�."

    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
                    contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>�����ۻ�ǰ</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
    		    else
        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
                    contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>�����ۻ�ǰ</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
    			end if
    		end if
		end if
		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)

		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
		'oMail.Send_CDO
	End IF

    ''�޸� ����.
    'contentsHtml = replace(contentsHtml,"�߼ۿ�����","�߼ۿ�����("&oneMisend.FOneItem.FMisendipgodate&")")
	'call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)

	SET oMail = nothing
	set oneMisend = Nothing
end function

CLASS CsActionMailCls

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	Dim FAsID
	Dim FDivCD
	Dim FGubun01
	Dim FGubun02

	Dim FDivCDName
	Dim FGubun01Name
	Dim FGubun02Name

	Dim FOrderSerial
	Dim FCustomerName
	Dim FUserid
	Dim FBuyHP
	Dim FBuyEmail
	Dim FWriteUser
	Dim FFinishUser
	Dim FTitle
	Dim FContents_jupsu
	Dim FContents_finish
	Dim FCurrstate
	Dim FCurrstateName
	Dim FRegDate
	Dim FFinishDate

	Dim FDeleteyn
	Dim FExtSiteName

	Dim FOpenTitle
	Dim FOpenContents

	Dim FSiteGubun

	Dim FSongjangDiv
	Dim FSongjangNo
	Dim FSongjangDivName

	Dim FRequireUpche
	Dim FMakerid

	Dim FAdd_upchejungsanDeliveryPay
	Dim FAdd_upchejungsanCause

	Dim FOrgSubTotalPrice
	Dim FOrgItemCostSum
	Dim FOrgBeasongPay
	Dim FOrgMileageSum
	Dim FOrgCouponSum
	Dim FOrgAllatDiscountSum

	Dim FRefundRequire
	Dim FRefundResult
	Dim FReturnMethod

	Dim FRefundMileageSum
	Dim FRefundCouponSum
	Dim FAllatSubTractSum

	Dim FRefundItemCostSum
	Dim FRefundBeasongPay
	Dim FRefundDeliveryPay
	Dim FRefundAdjustPay
	Dim FCancelTotal

	Dim FReturnName
	Dim FReturnPhone
	Dim FReturnHP
	Dim FReturnZipCode
	Dim FReturnZipAddr
	Dim FReturnEtcAddr


	Dim FReBankName
	Dim FReBankAccount
	Dim FReBankOwnerName

	Dim Fencmethod
	Dim FdecAccount

	Dim FPayGateTid

	Dim FPayGateResultTid
	Dim FPayGateResultMsg

	Dim FReturnMethodName

	Dim FReqName
	Dim FReqPhone
	Dim FReqHP
	Dim FReqZipcode
	Dim FReqZipAddr
	Dim FReqEtcAddr
	Dim FReqEtcStr
    Dim FInfoHtml

    Dim FupcheReturnSongjangDivName
    Dim FupcheReturnSongjangDivTel

	Dim FSendDate

	Dim FResultCount

    Dim FRectForceCurrState     ''���°� ���� ����.
    Dim FRectForceBuyEmail      ''�̸��� ��������.

    Dim Faccountdiv      ''2016/08/05 �߰�
    Dim Fpggubun         ''2016/08/05 �߰�

 	public function GetAsDivCDName()
        GetAsDivCDName = db2html(FDivCDName)
	end function

	public function GetCauseDetailString()
        GetCauseDetailString = Fgubun02Name
    end function

	Public Sub GetOneCSASMaster(FRectCsAsID)
		dim tmpZipCode, tmpaddress1, tmpaddress2
			tmpZipCode="11154"
			tmpaddress1="��⵵ ��õ�� ������ ����������2�� 83"
			tmpaddress2="�ٹ����� ��������"

		dim strSQL
		strSQL =" SELECT TOP 1 " &_
				" 	A.ID ,A.DivCD ,A.Gubun01 ,A.Gubun02 ,A.OrderSerial ,A.CustomerName ,A.Userid ,A.WriteUser ,A.FinishUser " &_
				"	,A.Title ,A.Contents_Jupsu ,A.Contents_Finish ,A.CurrState ,A.RegDate ,A.FinishDate ,A.Deleteyn ,A.ExtSiteName "&_
				"	,A.OpenTitle ,A.OpenContents ,A.RequireUpche ,A.Makerid ,A.SongjangDiv ,A.SongjangNo ,A.SiteGubun "&_
				"	,(SELECT TOP 1 divname FROM db_order.dbo.tbl_songjang_div WHERE divcd=A.SongjangDiv) AS SongjangDivName " &_
				" 	,o.BuyHp,o.BuyEmail " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.divCD) as divcdname " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.gubun01) as gubun01name " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.gubun02) as gubun02name "
		IF (FRectForceCurrState<>"") then
		    strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd='"&FRectForceCurrState&"') as currstatename "
        ELSE
            strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.currstate) as currstatename "
        END IF

		strSQL = strSQL & " 	,IsNULL(J.add_upchejungsandeliverypay,0) as add_upchejungsandeliverypay , J.add_upchejungsancause " &_

				" 	,r.OrgSubTotalPrice,r.OrgItemCostSum,r.OrgBeasongPay,r.OrgMileageSum,r.OrgCouponSum,r.OrgAllatDiscountSum "&_
				" 	,IsNULL(r.RefundRequire,0) as RefundRequire ,isNULL(r.RefundResult,0) as RefundResult "&_
				"	,r.ReturnMethod,r.RefundMileageSum,r.RefundCouponSum,r.AllatSubTractSum "&_
				"	,r.RefundItemCostSum,r.RefundBeasongPay,r.RefundDeliveryPay,r.RefundAdjustPay,r.CancelTotal "&_
				" 	,r.RebankName ,r.RebankAccount ,r.RebankOwnerName ,r.PayGateTid " &_
				"   ,r.encmethod " &_
				"   , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_cs.dbo.uf_DecAcctPH1(r.encaccount), '') WHEN r.encmethod='AE2' THEN IsNull(db_cs.dbo.uf_DecAcctAES256(r.encaccount), '') ELSE '' END) as decaccount " &_
				" 	,r.paygateresultTid,r.PayGateResultMsg " &_
				" 	,(SELECT top 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=r.returnmethod and comm_group='Z090') as ReturnMethodName " &_

				" 	,IsNULL(D.ReqName,o.reqname) as ReqName ,IsNULL(D.ReqPhone,o.reqphone) as ReqPhone ,IsNULL(D.ReqHP,o.reqhp) as ReqHP " &_
				" 	,IsNULL(D.ReqZipcode,o.reqzipcode) as ReqZipcode ,IsNULL(D.ReqZipAddr,o.reqzipaddr) as ReqZipAddr ,IsNULL(D.ReqEtcAddr,o.reqaddress) as ReqEtcAddr ,IsNULL(D.ReqEtcStr,'') as ReqEtcStr " &_
				" 	,isNull(p.company_name,'(��)�ٹ�����') as ReturnName ,isNull(p.deliver_phone,'1644-6030') as ReturnPhone ,isNull(p.deliver_hp,'') as ReturnHP "&_
				" 	,isNull(p.return_zipcode,'"& tmpZipCode &"') as ReturnZipCode ,isNull(p.return_address,'"& tmpaddress1 &"') as ReturnZipAddr ,isNull(p.return_address2,'"& tmpaddress2 &"') as ReturnEtcAddr "&_
                " 	,isNull((SELECT TOP 1 divname FROM db_order.dbo.tbl_songjang_div WHERE divcd=p.defaultsongjangdiv),'') as upcheReturnSongjangDivName "&_
                " 	,isNull((SELECT TOP 1 tel FROM db_order.dbo.tbl_songjang_div WHERE divcd=p.defaultsongjangdiv),'') as upcheReturnSongjangDivTel "&_
                "   ,isNULL(o.accountdiv,'') as accountdiv, isNULL(o.pggubun,'') as pggubun"&_

				" FROM [db_cs].[dbo].tbl_new_as_list A " &_
				" LEFT JOIN db_order.dbo.tbl_order_master o " &_
				" 	on A.orderserial=o.orderserial " &_
				" LEFT JOIN [db_cs].[dbo].tbl_as_upcheAddjungsan J " &_
				" 	on A.id=J.asid " &_
				" LEFT JOIN [db_cs].[dbo].tbl_as_refund_info r " &_
				" 	on A.id=r.asid " &_
				" LEFT JOIN [db_cs].[dbo].tbl_new_as_delivery d " &_
				" 	on A.id = d.asid " &_
				" LEFT JOIN [db_partner].[dbo].tbl_partner p " &_
				" 	on A.makerid= p.id " &_
				" WHERE A.id=" & CStr(FRectCsAsID)

			rsget.Open strSQL, dbget, 1

	        FResultCount = rsget.RecordCount

	        if  not rsget.EOF  then
	        	'//GetOneCSASMaster
				FAsID		= rsget("ID")
				FDivCD	= rsget("divCD")
				FGubun01	= rsget("gubun01")
				FGubun02	= rsget("gubun02")

				FDivCDName	= rsget("divcdname")
				FGubun01Name	= rsget("gubun01name")
				FGubun02Name	= rsget("gubun02name")

				FOrderSerial	= rsget("orderserial")
				FCustomerName	= rsget("customername")
				FUserid	= rsget("userid")
				FWriteUser	= rsget("writeuser")
				FFinishUser	= rsget("finishuser")
				FBuyHP		= rsget("BuyHP")
				FBuyEmail	= rsget("BuyEmail")

				if (FRectForceBuyEmail<>"") then
				    FBuyEmail = FRectForceBuyEmail
				end if

				FTitle	= rsget("title")
				FContents_jupsu	= rsget("contents_jupsu")
				FContents_finish	= rsget("contents_finish")

				IF (FRectForceCurrState<>"") then  ''���°� ���� ���� (���� ��߼۽� ���.)
				    FCurrState = FRectForceCurrState
				ELSE
    				FCurrState	= rsget("currstate")
    			END IF
				FCurrStateName	= db2html(rsget("currstatename"))
				FRegDate	= rsget("regdate")
				FFinishDate	= rsget("finishdate")

				FDeleteyn	= rsget("Deleteyn")
				FExtSiteName	= rsget("ExtSiteName")

				FOpenTitle	= rsget("OpenTitle")
				FOpenContents	= rsget("OpenContents")

				FSiteGubun	= rsget("SiteGubun")

				FSongjangDiv	= rsget("SongjangDiv")
				FSongjangNo	= rsget("SongjangNo")
				FSongjangDivName = rsget("SongjangDivName")
				FRequireUpche	= rsget("RequireUpche")
				FMakerid	= rsget("Makerid")

				FAdd_upchejungsanDeliveryPay	= rsget("Add_upchejungsanDeliveryPay")
				FAdd_upchejungsanCause	= rsget("Add_upchejungsanCause")

				'//GetOneRefundInfo
				FOrgSubTotalPrice	= rsget("OrgSubTotalPrice")
				FOrgItemCostSum	= rsget("OrgItemCostSum")
				FOrgBeasongPay	= rsget("OrgBeasongPay")
				FOrgMileageSum	= rsget("OrgMileageSum")
				FOrgCouponSum	= rsget("OrgCouponSum")
				FOrgAllatDiscountSum	= rsget("OrgAllatDiscountSum")
				FRefundRequire	= rsget("RefundRequire")
				FRefundResult	= rsget("RefundResult")
				FReturnMethod	= rsget("ReturnMethod")
				FRefundMileageSum	= rsget("RefundMileageSum")
				FRefundCouponSum	= rsget("RefundCouponSum")
				FRefundItemCostSum	= rsget("RefundItemCostSum")
				FRefundBeasongPay	= rsget("RefundBeasongPay")
				FRefundDeliveryPay	= rsget("RefundDeliveryPay")
				FRefundAdjustPay	= rsget("RefundAdjustPay")

				FAllatSubTractSum	= rsget("AllatSubTractSum")
				FCancelTotal	= rsget("CancelTotal")

				FReBankName	= rsget("ReBankName")
				FReBankAccount	= rsget("ReBankAccount")
				Fencmethod      = rsget("encmethod")
				FdecAccount      = rsget("decAccount")
				IF (Fencmethod="PH1") then FReBankAccount=FdecAccount
				IF (Fencmethod="AE2") then FReBankAccount=FdecAccount

				FReBankOwnerName	= rsget("ReBankOwnerName")
				FPayGateTid	= rsget("PayGateTid")

				FPayGateResultTid	= rsget("PayGateResultTid")
				FPayGateResultMsg	= rsget("PayGateResultMsg")

				FReturnMethodName	= rsget("ReturnMethodName")

				'//GetReturnAddress
				FReturnName	= rsget("ReturnName")
				FReturnPhone	= rsget("ReturnPhone")
				FReturnHP	= rsget("ReturnHP")
				FReturnZipCode	= rsget("ReturnZipCode")
				FReturnZipAddr	= rsget("ReturnZipAddr")
				FReturnEtcAddr	= rsget("ReturnEtcAddr")

				FReqName	= rsget("ReqName")
				FReqPhone	= rsget("ReqPhone")
				FReqHP		= rsget("ReqHP")
				FReqZipcode	= rsget("ReqZipcode")
				FReqZipAddr	= rsget("ReqZipAddr")
				FReqEtcAddr	= rsget("ReqEtcAddr")
				FReqEtcStr	= rsget("ReqEtcStr")

                FupcheReturnSongjangDivName = db2html(rsget("upcheReturnSongjangDivName"))
                FupcheReturnSongjangDivTel  = db2html(rsget("upcheReturnSongjangDivTel"))

                Faccountdiv = rsget("accountdiv")
                Fpggubun    = rsget("pggubun")

                if (Fpggubun="NP") then
                    if (FReturnMethod="R100" or FReturnMethod="R120" or FReturnMethod="R020" or FReturnMethod="R022") then
                        FReturnMethodName = "���̹��������"
                    end if
                end if
			END IF
		rsget.close

		''��Ÿ �ȳ� ����
		if (FDivCD<>"") and ((FCurrState="B001") or (FCurrState="B007")) then
		    strSQL = " SELECT TOP 1 IsNULL(infoHtml,'') as infoHtml from db_cs.dbo.tbl_cs_comm_div_info"
		    strSQL = strSQL + " where div_comm_cd='" + FDivCD + "'"
		    strSQL = strSQL + " and state_comm_cd='" + FCurrState + "'"

		    rsget.Open strSQL, dbget, 1
		    if  not rsget.EOF  then
		        FInfoHtml = db2Html(rsget("infoHtml"))
		    end if
		    rsget.Close
		end if
	End Sub

	Public Sub GetOneCSASMaster_GiftCard(FRectCsAsID)

		dim strSQL
		strSQL =" SELECT TOP 1 " &_
				" 	A.ID ,A.DivCD ,A.Gubun01 ,A.Gubun02 ,A.OrderSerial ,A.CustomerName ,A.Userid ,A.WriteUser ,A.FinishUser " &_
				"	,A.Title ,A.Contents_Jupsu ,A.Contents_Finish ,A.CurrState ,A.RegDate ,A.FinishDate ,A.Deleteyn ,A.ExtSiteName "&_
				"	,A.OpenTitle ,A.OpenContents ,A.RequireUpche ,A.Makerid ,A.SongjangDiv ,A.SongjangNo ,A.SiteGubun "&_
				"	,(SELECT TOP 1 divname FROM db_order.dbo.tbl_songjang_div WHERE divcd=A.SongjangDiv) AS SongjangDivName " &_
				" 	,o.BuyHp,o.BuyEmail " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.divCD) as divcdname " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.gubun01) as gubun01name " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.gubun02) as gubun02name "
		IF (FRectForceCurrState<>"") then
		    strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd='"&FRectForceCurrState&"') as currstatename "
        ELSE
            strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.currstate) as currstatename "
        END IF


		strSQL = strSQL & " 	,r.OrgSubTotalPrice,r.OrgItemCostSum,r.OrgBeasongPay,r.OrgMileageSum,r.OrgCouponSum,r.OrgAllatDiscountSum "&_
				" 	,IsNULL(r.RefundRequire,0) as RefundRequire ,isNULL(r.RefundResult,0) as RefundResult "&_
				"	,r.ReturnMethod,r.RefundMileageSum,r.RefundCouponSum,r.AllatSubTractSum "&_
				"	,r.RefundItemCostSum,r.RefundBeasongPay,r.RefundDeliveryPay,r.RefundAdjustPay,r.CancelTotal "&_
				" 	,r.RebankName ,r.RebankAccount ,r.RebankOwnerName ,r.PayGateTid " &_
				"   ,r.encmethod " &_
				"   , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_cs.dbo.uf_DecAcctPH1(r.encaccount), '') WHEN r.encmethod='AE2' THEN IsNull(db_cs.dbo.uf_DecAcctAES256(r.encaccount), '') ELSE '' END) as decaccount " &_
				" 	,r.paygateresultTid,r.PayGateResultMsg " &_
				" 	,(SELECT top 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=r.returnmethod and comm_group='Z090') as ReturnMethodName " &_

				" FROM [db_cs].[dbo].tbl_new_as_list A " &_
				" LEFT JOIN db_order.dbo.tbl_giftcard_order o " &_
				" 	on A.orderserial=o.giftorderserial " &_
				" LEFT JOIN [db_cs].[dbo].tbl_as_refund_info r " &_
				" 	on A.id=r.asid " &_
				" LEFT JOIN [db_cs].[dbo].tbl_new_as_delivery d " &_
				" 	on A.id = d.asid " &_
				" WHERE A.id=" & CStr(FRectCsAsID)

			rsget.Open strSQL, dbget, 1

	        FResultCount = rsget.RecordCount

	        if  not rsget.EOF  then
	        	'//GetOneCSASMaster
				FAsID		= rsget("ID")
				FDivCD	= rsget("divCD")
				FGubun01	= rsget("gubun01")
				FGubun02	= rsget("gubun02")

				FDivCDName	= rsget("divcdname")
				FGubun01Name	= rsget("gubun01name")
				FGubun02Name	= rsget("gubun02name")

				FOrderSerial	= rsget("orderserial")
				FCustomerName	= rsget("customername")
				FUserid	= rsget("userid")
				FWriteUser	= rsget("writeuser")
				FFinishUser	= rsget("finishuser")
				FBuyHP		= rsget("BuyHP")
				FBuyEmail	= rsget("BuyEmail")

				if (FRectForceBuyEmail<>"") then
				    FBuyEmail = FRectForceBuyEmail
				end if

				FTitle	= rsget("title")
				FContents_jupsu	= rsget("contents_jupsu")
				FContents_finish	= rsget("contents_finish")

				IF (FRectForceCurrState<>"") then  ''���°� ���� ���� (���� ��߼۽� ���.)
				    FCurrState = FRectForceCurrState
				ELSE
    				FCurrState	= rsget("currstate")
    			END IF
				FCurrStateName	= db2html(rsget("currstatename"))
				FRegDate	= rsget("regdate")
				FFinishDate	= rsget("finishdate")

				FDeleteyn	= rsget("Deleteyn")
				FExtSiteName	= rsget("ExtSiteName")

				FOpenTitle	= rsget("OpenTitle")
				FOpenContents	= rsget("OpenContents")

				FSiteGubun	= rsget("SiteGubun")

				FSongjangDiv	= rsget("SongjangDiv")
				FSongjangNo	= rsget("SongjangNo")
				FSongjangDivName = rsget("SongjangDivName")
				FRequireUpche	= rsget("RequireUpche")
				FMakerid	= rsget("Makerid")

				'FAdd_upchejungsanDeliveryPay	= rsget("Add_upchejungsanDeliveryPay")
				'FAdd_upchejungsanCause	= rsget("Add_upchejungsanCause")

				'//GetOneRefundInfo
				FOrgSubTotalPrice	= rsget("OrgSubTotalPrice")
				FOrgItemCostSum	= rsget("OrgItemCostSum")
				FOrgBeasongPay	= rsget("OrgBeasongPay")
				FOrgMileageSum	= rsget("OrgMileageSum")
				FOrgCouponSum	= rsget("OrgCouponSum")
				FOrgAllatDiscountSum	= rsget("OrgAllatDiscountSum")
				FRefundRequire	= rsget("RefundRequire")
				FRefundResult	= rsget("RefundResult")
				FReturnMethod	= rsget("ReturnMethod")
				FRefundMileageSum	= rsget("RefundMileageSum")
				FRefundCouponSum	= rsget("RefundCouponSum")
				FRefundItemCostSum	= rsget("RefundItemCostSum")
				FRefundBeasongPay	= rsget("RefundBeasongPay")
				FRefundDeliveryPay	= rsget("RefundDeliveryPay")
				FRefundAdjustPay	= rsget("RefundAdjustPay")

				FAllatSubTractSum	= rsget("AllatSubTractSum")
				FCancelTotal	= rsget("CancelTotal")

				FReBankName	= rsget("ReBankName")
				FReBankAccount	= rsget("ReBankAccount")
				Fencmethod      = rsget("encmethod")
				FdecAccount      = rsget("decAccount")
				IF (Fencmethod="PH1") then FReBankAccount=FdecAccount
				IF (Fencmethod="AE2") then FReBankAccount=FdecAccount

				FReBankOwnerName	= rsget("ReBankOwnerName")
				FPayGateTid	= rsget("PayGateTid")

				FPayGateResultTid	= rsget("PayGateResultTid")
				FPayGateResultMsg	= rsget("PayGateResultMsg")

				FReturnMethodName	= rsget("ReturnMethodName")

				'//GetReturnAddress
				'FReturnName	= rsget("ReturnName")
				'FReturnPhone	= rsget("ReturnPhone")
				'FReturnHP	= rsget("ReturnHP")
				'FReturnZipCode	= rsget("ReturnZipCode")
				'FReturnZipAddr	= rsget("ReturnZipAddr")
				'FReturnEtcAddr	= rsget("ReturnEtcAddr")

				'FReqName	= rsget("ReqName")
				'FReqPhone	= rsget("ReqPhone")
				'FReqHP		= rsget("ReqHP")
				'FReqZipcode	= rsget("ReqZipcode")
				'FReqZipAddr	= rsget("ReqZipAddr")
				'FReqEtcAddr	= rsget("ReqEtcAddr")
				'FReqEtcStr	= rsget("ReqEtcStr")

                'FupcheReturnSongjangDivName = db2html(rsget("upcheReturnSongjangDivName"))
                'FupcheReturnSongjangDivTel  = db2html(rsget("upcheReturnSongjangDivTel"))
			END IF
		rsget.close

		''��Ÿ �ȳ� ����
		if (FDivCD<>"") and ((FCurrState="B001") or (FCurrState="B007")) then
		    strSQL = " SELECT TOP 1 IsNULL(infoHtml,'') as infoHtml from db_cs.dbo.tbl_cs_comm_div_info"
		    strSQL = strSQL + " where div_comm_cd='" + FDivCD + "'"
		    strSQL = strSQL + " and state_comm_cd='" + FCurrState + "'"

		    rsget.Open strSQL, dbget, 1
		    if  not rsget.EOF  then
		        FInfoHtml = db2Html(rsget("infoHtml"))
		    end if
		    rsget.Close
		end if
	End Sub

	''// ���� ��� �̹���
	Public Function getMailHeadImage()
		dim tmpImg
		IF FDivCD="A000" Then '// �±�ȯ���
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_exchange.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_exchange_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A001" Then '// ������߼�
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_resend.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_resend_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A002" Then '// ���񽺹߼�
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_send_service.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_send_service_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A003" Then '// ȯ�ҿ�û
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_refund.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_refund_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A004" Then '// ��ǰ����(��)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_return.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_return_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A007" Then '// �ſ�/��ü���
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_payment_cancel.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_payment_cancel_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A008" Then '// �ֹ����
			IF FCurrState="B001" Then
				'tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_order_cancel.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_order_cancel_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A010" Then '// ȸ����û(��)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_prd_recovery.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_prd_recovery_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A011" Then '// �±�ȯȸ��(��)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_cancel_prd_recovery.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_cancel_prd_recovery_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A900" Then '// �ֹ���������
			IF FCurrState="B001" Then
				'tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_change_order.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_change_order_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSE

		END IF
		getMailHeadImage = tmpImg
	End Function

	''// ���� ����	'2017.12.19 �ѿ�� ����
	Public Function getMailHeadtitle()
		dim tmptitle
		IF FDivCD="A000" Then '// �±�ȯ���
			IF FCurrState="B001" Then
				tmptitle = "��ȯ��� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "��ȯ��� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A001" Then '// ������߼�
			IF FCurrState="B001" Then
				tmptitle = "������߼� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "������߼� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A002" Then '// ���񽺹߼�
			IF FCurrState="B001" Then
				tmptitle = "���� �߼� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "���� �߼� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A003" Then '// ȯ�ҿ�û
			IF FCurrState="B001" Then
				tmptitle = "ȯ�� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "ȯ�� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A004" Then '// ��ǰ����(��)
			IF FCurrState="B001" Then
				tmptitle = "��ǰ ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "��ǰ �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A007" Then '// �ſ�/��ü���
			IF FCurrState="B001" Then
				tmptitle = "������� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "������� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A008" Then '// �ֹ����
			IF FCurrState="B001" Then
				'tmptitle = "�ֹ���� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "�ֹ���� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A010" Then '// ȸ����û(��)
			IF FCurrState="B001" Then
				tmptitle = "��ǰȸ�� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "��ǰȸ�� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A011" Then '// �±�ȯȸ��(��)
			IF FCurrState="B001" Then
				tmptitle = "��ȯ��ǰ ȸ�� ���� ����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "��ȯ��ǰ ȸ�� �Ϸ� ����"
			End IF
		ELSEIF FDivCD="A900" Then '// �ֹ���������
			IF FCurrState="B001" Then
				'tmptitle = "�ֹ��������� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "�ֹ��������� �Ϸ� �ȳ�����"
			End IF
		ELSE

		END IF
		getMailHeadtitle = tmptitle
	End Function

	'// ��Ÿ �ȳ�����		'/2017.12.19 �ѿ��
	Public Function getEtcNotice()
		dim tmpHTML

        getEtcNotice = ""

        if (Trim(FInfoHtml)="") then Exit function

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""����"", sans-serif; text-align:left;'>��Ÿ �ȳ� ����</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:10px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""����"", sans-serif; color:#707070; vertical-align:top; text-align:left;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:630px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									"& FInfoHtml &"" & vbcrlf
		tmpHTML=tmpHTML&"								</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		getEtcNotice = tmpHTML
	End Function

	''// �ù� ���� ��������		'/2017.12.19 �ѿ��
	Function getDlvInfo()
		dim tmpHTML
		tmpHTML=""

        if (IsNULL(FSongjangNo)) or (FSongjangNo="") then Exit function

		'A001(������߼�), A011(�±�ȯȸ��(��)), A000(�±�ȯ���), A010(ȸ����û(��)), A002(���񽺹߼�), A004(��ǰ����(��))
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>�ù�����</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf

			IF FSongjangNo<>"" then
				tmpHTML=tmpHTML&"									<span style='margin:0; text-align:left; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070;'>"& FSongjangDivName &"</span>"& vbcrlf
				tmpHTML=tmpHTML&"									&nbsp;&nbsp;<a href='"& DeliverDivTrace(Trim(FSongjangDiv)) & FSongjangNo &"' target='_blank' style='margin:0; padding:0; font-size:12px; color:#dd5555; font-size:11px; line-height:18px; color:#0066cc; text-align:left;'>"& FSongjangNo &"</a>" & vbcrlf
			ELSE
				tmpHTML=tmpHTML&"									<span style='margin:0; text-align:left; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070;'>�ù������� ��ϵ��� �ʾҽ��ϴ�.</span>" & vbcrlf
			END IF

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		END IF

		getDlvInfo =  tmpHTML
	END Function

	'// ó�� ��� ��������		'/2017.12.19 �ѿ��
	Function getFinishResult()
		dim tmpHTML
		tmpHTML=""

		'A001(������߼�), A011(�±�ȯȸ��(��)), A010(ȸ����û(��)), A003(ȯ�ҿ�û), A002(���񽺹߼�), A004(��ǰ����(��)), A000(�±�ȯ���)
		IF FCurrState="B007" THEN
		    ''ó�� ������ ������..
		    if (FOpenContents="") then
		        if (FDivCD="A000") then
		            FOpenContents = "�±�ȯ��ǰ ���Ϸ�"
		        elseif (FDivCD="A001") then
		            FOpenContents = "������ǰ ���Ϸ�"
		        elseif (FDivCD="A002") then
		            FOpenContents = "��ǰ ���Ϸ�"
		        elseif (FDivCD="A003") then

		        elseif (FDivCD="A004") then
		            FOpenContents = "��ǰ ��ǰ(ȸ��)�Ϸ�" '' / ȯ�ҵ��"

		        elseif (FDivCD="A010") then
		            FOpenContents = "��ǰ ȸ���Ϸ�" '' / ȯ�ҵ��"
		        elseif (FDivCD="A011") then
		            FOpenContents = "�±�ȯ��ǰ ȸ���Ϸ�"
		        else

		        end if
		    end if

			tmpHTML=tmpHTML&"	<tr>" & vbcrlf
			tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
			tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<th style='margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; text-align:left; font-family:dotum, ""����"", sans-serif; text-align:left;'>ó����� ����<th>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
			tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>ó���Ϸ���</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& FFinishDate &"</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			IF (Trim(FOpenContents)<>"") then
				tmpHTML=tmpHTML&"							<tr>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>ó������</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:16px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& nl2br(FOpenContents) &"</td>" & vbcrlf
				tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			end IF

			''// �ù����� ��������
			tmpHTML=tmpHTML& getDlvInfo()

			tmpHTML=tmpHTML&"						</table>" & vbcrlf
			tmpHTML=tmpHTML&"					</td>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"			</table>" & vbcrlf
			tmpHTML=tmpHTML&"		</td>" & vbcrlf
			tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		END IF
		getFinishResult = tmpHTML
	END Function

	''//ȯ������ ��������		'/2017.12.19 �ѿ��
	Function getRefundInfo()
		dim tmpHTML
		tmpHTML=""

		' A008(�ֹ����), A010(ȸ����û(��)), A007(�ſ�/��ü���), A003(ȯ�ҿ�û), A004(��ǰ����(��))
		IF FDivCD="A003" or FDivCD="A004" or FDivCD="A007" or FDivCD="A008" or FDivCD="A010" THEN
		    ''ȯ�Ҿ�0�̸� return
		    if (FRefundRequire=0) then Exit function

		    ''����Ȯ�� ȯ�� ���� ����
		    if (FReturnMethod="R007") then
		        if (Len(Replace(FReBankAccount,"-",""))<7) then
    		        FReBankName = ""
    		        FReBankAccount = "����Ȯ�ο��"
    		        FReBankOwnerName =""
    		    else
    		        FReBankAccount = Left(FReBankAccount,Len(Trim(FReBankAccount))-3) + "***"
    		    end if
		    end if

			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>ȯ�ҿ�����</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FormatNumber(FRefundRequire,0) &" ��" & vbcrlf

			'��ۺ����� �ȳ��� ���ظ� ������ ǥ�þ��ϰ� ����
			'if (FRefundDeliveryPay<>0) then
			'			tmpHTML=tmpHTML&"									(��ۺ����� : " & FormatNumber(FRefundDeliveryPay+Frefundbeasongpay,0) &")" & vbcrlf
			'end if

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>ȯ������(����)</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FReturnMethodName &"&nbsp;&nbsp;" & vbcrlf

			IF (FReturnMethod="R007") THEN
				tmpHTML=tmpHTML&"									"& FReBankName &"&nbsp;&nbsp; " & vbcrlf
				tmpHTML=tmpHTML&"									"& FReBankAccount &"&nbsp;&nbsp; " & vbcrlf
				tmpHTML=tmpHTML&"									"& AstarUserName(FReBankOwnerName) &" " & vbcrlf
			ELSEIF (FReturnMethod="R900") THEN
				tmpHTML=tmpHTML&"									(�������̵� : "& FUserID &") " & vbcrlf
			ELSEIF (FReturnMethod="R100") or (FReturnMethod="R550") or (FReturnMethod="R560") or (FReturnMethod="R120") or (FReturnMethod="R020") or (FReturnMethod="R080") THEN
				if (Left(FPayGateTid,6)="IniTec") and (FCurrState="B007") and (FReturnMethod<>"R120") then
					tmpHTML=tmpHTML&"									<a target='_blank' href=https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid="& FPayGateTid &"&noMethod=1>[������ǥ���]</a> " & vbcrlf
				end if
				if (FReturnMethod = "R550") or (FReturnMethod = "R560") then
					tmpHTML=tmpHTML&"									������/����Ƽ�� �� ����� ���������� �߱޵Ǹ�, ��ǰ���Žÿ��� ����Ұ��Դϴ�." & vbcrlf
				end if
			END IF

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		END IF
		getRefundInfo = tmpHTML
	END Function

	''//��ü �ּ� ��������		'/2017.12.19 �ѿ��
	Function getReturnInfo()
		dim tmpHTML
		tmpHTML=""

		' A011(�±�ȯȸ��(��)), A010(ȸ����û(��)), A004(��ǰ����(��))
		IF FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>��ǰȸ���ּ�</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FReturnName &" &nbsp; &nbsp;"& FReturnPhone &"<br />["& FReturnZipCode &"] "& FReturnZipAddr &"&nbsp;"& FReturnEtcAddr &"" & vbcrlf
			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			if (FReturnName<>"(��)�ٹ�����") and (FupcheReturnSongjangDivName<>"") and (Left(FupcheReturnSongjangDivTel,1)="1" or Left(FupcheReturnSongjangDivTel,1)="0") then
				tmpHTML=tmpHTML&"							<tr>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>�̿��ù��</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
				tmpHTML=tmpHTML&"									"& FupcheReturnSongjangDivName &"<br />�ù�翬��ó : "& FupcheReturnSongjangDivTel &"" & vbcrlf
				tmpHTML=tmpHTML&"								</td>" & vbcrlf
				tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			END IF
		END IF

		getReturnInfo = tmpHTML
	END Function

	''//���ּ� ��������		'/2017.12.19 �ѿ��
	Function getReqInfo()
		dim tmpHTML
		tmpHTML=""

		'A001(������߼�), A011(�±�ȯȸ��(��)), A000(�±�ȯ���), A010(ȸ����û(��)), A002(���񽺹߼�)
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A010" THEN 'or FDivCD="A011"
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>���ּ�</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& AstarUserName(trim(FReqName)) &" ���� &nbsp; &nbsp;"& AstarPhoneNumber(trim(FReqPhone)) &" / "& AstarPhoneNumber(trim(FReqHP)) &" <br />["& printUserId(trim(FReqZipcode), 2, "*") &"] "& printUserId(trim(FReqZipAddr), 2, "*") &"&nbsp;(���ϻ���)" & vbcrlf		' FReqEtcAddr
			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		END IF
		getReqInfo = tmpHTML
	END Function

	''//���� ��ǰ ���� ��������		'/2017.12.19 �ѿ��
	Function getAsItemLIst()
		dim tmpHTML
		dim OCsDetail,i

		tmpHTML = ""

		'A001(������߼�), A008(�ֹ����), A011(�±�ȯȸ��(��)), A000(�±�ȯ���), A010(ȸ����û(��)), A002(���񽺹߼�), A004(��ǰ����(��))
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A004" or FDivCD="A008" or FDivCD="A010" or FDivCD="A011" THEN
			Set OCsDetail = New CCSASList
			OCsDetail.FRectCsAsID = FAsID
			IF FResultCount>0 THEN
				OCsDetail.GetCsDetailList
			END IF

			if (OCsDetail.FresultCount<1) then Exit function

			tmpHTML=tmpHTML&"	<tr>" & vbcrlf
			tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
			tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""����"", sans-serif; text-align:left; color:#000;'>���� ��ǰ ����</th>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
			tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:12px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:50px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, ""����"", sans-serif; text-align:center; color:#707070; font-size:12px; line-height:12px;'>��ǰ</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:100px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>��ǰ�ڵ�</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:295px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>��ǰ��[�ɼ�]</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:right; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>�ǸŰ���</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:25px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>&nbsp;</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>����</th>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			IF OCsDetail.FresultCount>0 Then
				FOR i=0 TO OCsDetail.FResultCount-1
				    IF (OCsDetail.FItemList(i).Fitemid<>0) or (OCsDetail.FItemList(i).Fitemcost<>0) then
						tmpHTML=tmpHTML&"							<tr>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:50px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea;'>" & vbcrlf
						tmpHTML=tmpHTML&"									<img src='"& OCsDetail.FItemList(i).FSmallImage &"' alt='' />" & vbcrlf
						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf
						tmpHTML=tmpHTML&"									"& OCsDetail.FItemList(i).Fitemid &"" & vbcrlf
						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf

						IF (OCsDetail.FItemList(i).Fitemid=0) Then
							tmpHTML=tmpHTML&"									��ۺ�" & vbcrlf
						ELSE
							tmpHTML=tmpHTML&"									"& OCsDetail.FItemList(i).Fitemname &"" & vbcrlf
						END IF
						if ( OCsDetail.FItemList(i).Fitemoptionname <>"") then
							tmpHTML=tmpHTML&"									["& OCsDetail.FItemList(i).Fitemoptionname &"]" & vbcrlf
						END IF

						tmpHTML=tmpHTML&"								</td>" & vbcrlf

						tmpHTML=tmpHTML&"								<td style='width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, ""����"", sans-serif;'>" & vbcrlf

						IF (OCsDetail.FItemList(i).FdiscountAssingedCost<>0) and (OCsDetail.FItemList(i).Fitemcost>OCsDetail.FItemList(i).FdiscountAssingedCost) then
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:6px 0; font-size:11px; line-height:16px; color:#707070; font-family:dotum, ""����"", sans-serif; text-decoration:line-through; text-align:right;'>"& FormatNumber(OCsDetail.FItemList(i).Fitemcost,0) & "��</span><br />" & vbcrlf
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, ""����"", sans-serif; text-align:right;'>" & FormatNumber(OCsDetail.FItemList(i).FdiscountAssingedCost,0) &"��</span>" & vbcrlf
						ELSE
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, ""����"", sans-serif; text-align:right;'>"& FormatNumber(OCsDetail.FItemList(i).Fitemcost,0) &"��</span>" & vbcrlf
						END IF

						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>"& OCsDetail.FItemList(i).Fregitemno &"</td>" & vbcrlf
						tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			        END IF
				NEXT
			END IF

			tmpHTML=tmpHTML&"						</table>" & vbcrlf
			tmpHTML=tmpHTML&"					</td>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"			</table>" & vbcrlf
			tmpHTML=tmpHTML&"		</td>" & vbcrlf
			tmpHTML=tmpHTML&"	</tr>" & vbcrlf

			Set OCsDetail= nothing
		END IF
		getAsItemLIst = tmpHTML
	END Function

	''// ���� �⺻ ���� ��������		'/2017.12.19 �ѿ��
	Function getAsInfo()
		dim tmpHTML

		tmpHTML = ""
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>������</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& FRegDate &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>�����ڵ�</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& FAsID &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>�ֹ���ȣ</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& FOrderSerial &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>��������</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									"& FTitle &" <a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp?orderserial=" & FOrderSerial & "' target='_blank'>[�󼼳�������]</a>" & vbcrlf
		tmpHTML=tmpHTML&"								</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>��������</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& GetCauseDetailString &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf

		getAsInfo =tmpHTML
	END Function

	'// SMS������
	Public Function sendSMS(byval ipHp, byval ipText)
		dim tmpSms,strSQL
		dim RcvHp,RcvMsg

		'// ���� �Էµ� ���� ������� �ڵ� ����
		IF ipHp<>"" THEN
			RcvHp=ipHp
		ELSE
			RcvHp=FBuyHP
		END IF

		IF ipText<>"" THEN
			RcvMsg=ipText
		ELSE
			RcvMsg="[�ٹ�����] ��û�Ͻ� ["& GetAsDivCDName &"] ó���� "& FCurrStateName &" �Ǿ����ϴ�."
		END IF

		On Error Resume Next

		dbget.beginTrans

		IF RcvHp<>"" and not isnull(RcvHp) THEN
			strSQL = "INSERT INTO [db_sms].[ismsuser].em_tran (tran_phone, tran_callback, tran_status, tran_date, tran_msg)" &vbcrlf
			strSQL = strSQL & "VALUES('"& RcvHp &"','1644-6030','1',getdate(),'" & db2html(RcvMsg) & "')"
			dbget.execute(strSQL)
		END IF

		IF Err.Number = 0 Then
			dbget.commitTrans
			response.write "SMS �߼� - �Ϸ�"
			Exit Function
		ELSE
			dbget.RollBackTrans
			response.write "SMS �߼� - ����"
			Exit Function
		EnD IF

	End Function

	'// �̸��� ���ø� �����ͼ� ����°ɷ� ����.	2017.12.18 �ѿ�� ����
	Function makeMailTemplate(id)
		dim tmpHTML, fs, dirPath, fileName, objFile, mailheader, mailfooter

		Call GetOneCSASMaster(id) '// �� ����

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_header_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailheader = objFile.readall	' ���

		tmpHTML=mailheader

		'A001(������߼�), A008(�ֹ����), A000(�±�ȯ���)
		tmpHTML=tmpHTML&" <table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='height:253px; text-align:center;'>"& getMailHeadImage &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:20px 0 45px; font-size:12px; line-height:22px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:center;'>" & vbcrlf

		if (FDivCD = "A008") then
			tmpHTML=tmpHTML&"			������ "& GetAsDivCDName &"�� ���������� ó���� "& FCurrStateName &" �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		elseif (FDivCD = "A000") then
			tmpHTML=tmpHTML&"			������ ��û�Ͻ� "& GetAsDivCDName &"�� ���������� ó���� "& FCurrStateName &" �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		else
			tmpHTML=tmpHTML&"			������ ��û�Ͻ� "& GetAsDivCDName &"�� ���������� ó���� "& FCurrStateName &" �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		end if

		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:0 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%; background:#f8f8f8; text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:297px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""����"", sans-serif; text-align:right;'>�ֹ���ȣ :</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:7px; padding:34px 0;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:331px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""����"", sans-serif; color:#dd5555; text-align:left; letter-spacing:-1px;'>" & vbcrlf
		tmpHTML=tmpHTML&"						"& FOrderSerial &"" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// ó����� ��������
		tmpHTML=tmpHTML& getFinishResult()

		''// ���� ��ǰ ���� ��������
		tmpHTML=tmpHTML& getAsItemLIst()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""����"", sans-serif; text-align:left; color:#000;'>���� ����</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf

		''// ���� �⺻ ���� ��������
		tmpHTML=tmpHTML& getAsInfo()

		''// ���ּ� ��������
		tmpHTML=tmpHTML& getReqInfo()

		''// ��ü�ּ� ��������
		tmpHTML=tmpHTML& getReturnInfo()

		''// ȯ������ ��������
		tmpHTML=tmpHTML& getRefundInfo()

		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// ��Ÿ �ȳ�����
		tmpHTML=tmpHTML& getEtcNotice()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 104px; margin:0;text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_receiption_info.png' alt='���� ���� �󼼺���' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_go_shopping.png' alt='�ٹ����� �����ϱ�' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:25px 0; border-top:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:center;'>������ ��� ���� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"</table>" & vbcrlf

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' Ǫ��

		tmpHTML=tmpHTML&mailfooter

		tmpHTML = replace(tmpHTML,":mailtitle:", getMailHeadtitle())		' �̸�������

		makeMailTemplate = tmpHTML
	End Function

	'/ �̸��� ���ø� �����ͼ� ����°ɷ� ����.	2017.12.18 �ѿ�� ����
	Function makeMailTemplate_GiftCard(id)
		dim tmpHTML, fs, dirPath, fileName, objFile, mailheader, mailfooter

		Call GetOneCSASMaster_GiftCard(id) '// �� ����

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_header_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailheader = objFile.readall	' ���

		tmpHTML=mailheader

		'A001(������߼�), A008(�ֹ����), A000(�±�ȯ���)
		tmpHTML=tmpHTML&" <table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='height:253px; text-align:center;'>"& getMailHeadImage &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:20px 0 45px; font-size:12px; line-height:22px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:center;'>" & vbcrlf

		if (FDivCD = "A008") then
			tmpHTML=tmpHTML&"			������ "& GetAsDivCDName &"�� ���������� ó���� "& FCurrStateName &" �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		elseif (FDivCD = "A000") then
			tmpHTML=tmpHTML&"			������ ��û�Ͻ� "& GetAsDivCDName &"�� ���������� ó���� "& FCurrStateName &" �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		else
			tmpHTML=tmpHTML&"			������ ��û�Ͻ� "& GetAsDivCDName &"�� ���������� ó���� "& FCurrStateName &" �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		end if

		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:0 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%; background:#f8f8f8; text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:297px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""����"", sans-serif; text-align:right;'>�ֹ���ȣ :</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:7px; padding:34px 0;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:331px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""����"", sans-serif; color:#dd5555; text-align:left; letter-spacing:-1px;'>" & vbcrlf
		tmpHTML=tmpHTML&"						"& FOrderSerial &"" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// ó����� ��������
		tmpHTML=tmpHTML& getFinishResult()

		''// ���� ��ǰ ���� ��������
		'tmpHTML=tmpHTML& getAsItemLIst()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""����"", sans-serif; text-align:left; color:#000;'>���� ����</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf

		''// ���� �⺻ ���� ��������
		tmpHTML=tmpHTML& getAsInfo()

		''// ���ּ� ��������
		'tmpHTML=tmpHTML& getReqInfo()

		''// ��ü�ּ� ��������
		'tmpHTML=tmpHTML& getReturnInfo()

		''// ȯ������ ��������
		tmpHTML=tmpHTML& getRefundInfo()

		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// ��Ÿ �ȳ�����
		'tmpHTML=tmpHTML& getEtcNotice()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 104px; margin:0;text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_receiption_info.png' alt='���� ���� �󼼺���' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_go_shopping.png' alt='�ٹ����� �����ϱ�' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:25px 0; border-top:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:center;'>������ ��� ���� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"</table>" & vbcrlf

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' Ǫ��

		tmpHTML=tmpHTML&mailfooter

		tmpHTML = replace(tmpHTML,":mailtitle:", getMailHeadtitle())		' �̸�������

		makeMailTemplate_GiftCard = tmpHTML
	End Function

	'// �̸��� ���ø� �����ͼ� ����°ɷ� ����.	2020.03.12 �ѿ�� ����
	Function makeMailTemplate_UpheBoard(id)
		dim tmpHTML, fs, dirPath, fileName, objFile, mailheader, mailfooter, boardqna

		set boardqna = New CUpcheQnADetail
			boardqna.FRectIdx = idx
			boardqna.read()

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_header_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailheader = objFile.readall	' ���

		tmpHTML=mailheader

		tmpHTML=tmpHTML&" <table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:20px 0 45px; font-size:12px; line-height:22px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML & boardqna.Fusername &"�� ��ü�Խ��ǿ� �����Ͻ� ������ ���������� ó���� �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""����"", sans-serif; text-align:left;'>��������</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:10px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""����"", sans-serif; color:#707070; vertical-align:top; text-align:left;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:630px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									<strong>��������</strong><br>"& nl2br(boardqna.Ftitle) &"" & vbcrlf
		tmpHTML=tmpHTML&"									<Br><br><strong>���ǳ���</strong><br>"& nl2br(boardqna.Fcontents) &"" & vbcrlf
		tmpHTML=tmpHTML&"								</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""����"", sans-serif; text-align:left;'>ó�����</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:10px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""����"", sans-serif; color:#707070; vertical-align:top; text-align:left;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:630px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									<strong>�亯����</strong><br>"& nl2br(boardqna.Freplytitle) &"" & vbcrlf
		tmpHTML=tmpHTML&"									<Br><br><strong>�亯����</strong><br>"& nl2br(boardqna.Freplycontents) &"" & vbcrlf
		tmpHTML=tmpHTML&"								</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:25px 0; border-top:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:center;'>������ ��� ���� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"</table>" & vbcrlf

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' Ǫ��

		tmpHTML=tmpHTML&mailfooter

		tmpHTML = replace(tmpHTML,":mailtitle:", getMailHeadtitle())		' �̸�������

		makeMailTemplate_UpheBoard = tmpHTML
		set boardqna=nothing
	End Function
End Class

'/ �Ⱦ��µ��ѵ�? ������.
'function SendMiChulgoMail(idx)
'    ''require /lib/classes/cscenter/oldmisendcls.asp
'    dim oneMisend
'    dim strMailHTML,strMailTitle, contentsHtml
'	strMailHTML = ""
'	strMailTitle = "[�ٹ�����] ��� ���� �ȳ������Դϴ�."
'
'    set oneMisend = new COldMiSend
'    oneMisend.FRectDetailIDx = idx
'    oneMisend.getOneOldMisendItem
'
'	'//=======  ���� �߼� =========/
'	dim oMail
'	dim MailHTML
'
'	set oMail = New MailCls         '' mailLib2
'
'	IF oneMisend.FOneItem.Fbuyemail<>"" THEN
'
'		oMail.MailTitles	= strMailTitle
'		oMail.SenderNm		= "�ٹ�����"
'		oMail.SenderMail	= "customer@10x10.co.kr"
'		oMail.AddrType		= "string"
'		oMail.ReceiverNm	= oneMisend.FOneItem.FBuyname
'		oMail.ReceiverMail	= oneMisend.FOneItem.FBuyEmail
'		oMail.MailType = "22"
'		strMailHTML = oMail.getMailTemplate
'		''parsing
'		strMailHTML = replace(strMailHTML,":ORDERSERIAL:",oneMisend.FOneItem.Forderserial)
'		strMailHTML = replace(strMailHTML,":ITEMIMAGEURL:",oneMisend.FOneItem.Fsmallimage)
'		strMailHTML = replace(strMailHTML,":ITEMID:",oneMisend.FOneItem.Fitemid)
'		strMailHTML = replace(strMailHTML,":ITEMNAME:",oneMisend.FOneItem.Fitemname)
'		strMailHTML = replace(strMailHTML,":ITEMCNT:",oneMisend.FOneItem.Fitemcnt)
'		strMailHTML = replace(strMailHTML,":COMPANYNAME:",oneMisend.FOneItem.getDlvCompanyName)
'		strMailHTML = replace(strMailHTML,":MAYSENDDATE:",oneMisend.FOneItem.FMisendipgodate)
'
'		if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
'		    if (oneMisend.FOneItem.FMisendReason<>"07") then
'		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�� ������ �ش� �Ǹ��ڰ� ���Բ� �����帮�� �����Դϴ�.<br>*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
'		    else
'		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�� ������ �ش� �Ǹ��ڰ� ���Բ� �����帮�� �����Դϴ�.")
'		    end if
'		else
'		    if (oneMisend.FOneItem.FMisendReason<>"07") then
'		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
'		    else
'		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","")
'		    end if
'		end if
'
'		if (oneMisend.FOneItem.FMisendipgodate<>"") then
'    		if (oneMisend.FOneItem.FMisendReason="03") then
'    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
'    		        contentsHtml = "�ȳ��ϼ���.   ����<br>"
'        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� �߼��� ������ �����Դϴ�.<br>"
'        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
'        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
'        			else
'        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
'        			end if
'        			contentsHtml = contentsHtml & "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
'        			contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
'        			contentsHtml = contentsHtml & "���ο� ������ �帰 �� �������� ����帮��, ������� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.<br>"
'    		    else
'        		    contentsHtml = "�ȳ��ϼ���.   ����<br>"
'        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>"
'        			contentsHtml = contentsHtml & "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
'        			contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
'
'    		    end if
'    		elseif (oneMisend.FOneItem.FMisendReason="02") then
'    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
'    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
'        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� �ֹ� �� ����(����)�Ǵ� ��ǰ����<br>"
'        			contentsHtml = contentsHtml & "�Ϲݻ�ǰ�� �޸� �ֹ�����(����)�� �Ⱓ�� �ҿ�Ǵ� ��ǰ�Դϴ�.<br>"
'        			contentsHtml = contentsHtml & "�Ʒ��� ���� �߼ۿ������� �ȳ��ص帮����,<br>"
'        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
'        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
'        			else
'        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
'        			end if
'    		    else
'        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
'        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>"
'        			contentsHtml = contentsHtml & "�Ʒ��� ���� �߼ۿ������� �ȳ��� �帳�ϴ�.<br>"
'        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
'        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
'        			else
'        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
'        			end if
'    			end if
'    	    elseif (oneMisend.FOneItem.FMisendReason="04") then
'    	        oMail.MailTitles = "[�ٹ�����] ��� ���� �ȳ������Դϴ�."
'
'    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
'    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
'                    contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
'                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>�����ۻ�ǰ</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
'                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
'                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
'    		    else
'        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
'                    contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
'                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>�����ۻ�ǰ</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
'                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
'                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
'    			end if
'    		elseif (oneMisend.FOneItem.FMisendReason="07") then
'    	        oMail.MailTitles = "[�ٹ�����] ��� ���� �ȳ������Դϴ�."
'
'    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
'    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
'                    contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
'                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>���������</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
'                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
'                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
'    		    else
'        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
'                    contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
'                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>���������</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
'                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
'                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
'    			end if
'    		end if
'		end if
'		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)
'
'		oMail.MailConts 	= strMailHTML
'
'		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
'		oMail.Send_TMSMailer()		'TMS���Ϸ�
'		'oMail.Send_Mailer()
'		oMail.Send_CDO
'	End IF
'
'    ''�޸� ����.
'    contentsHtml = replace(contentsHtml,"�߼ۿ�����","�߼ۿ�����("&oneMisend.FOneItem.FMisendipgodate&")")
'	call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)
'
'	SET oMail = nothing
'	set oneMisend = Nothing
'end function

%>
