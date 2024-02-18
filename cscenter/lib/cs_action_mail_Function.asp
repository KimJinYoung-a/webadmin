<%
'###########################################################
' Description : cs센터 이메일 발송 공통 함수
' History : 이상구 생성
'			2017.12.21 한용민 수정
'###########################################################

Function SendCsActionMail(id)

    dim oCsAction,strMailHTML,strMailTitle
	Set oCsAction = New CsActionMailCls
	strMailHTML = oCsAction.makeMailTemplate(id)

	if (oCsAction.FDivCD = "A008") then
		strMailTitle = "[텐바이텐] "& oCsAction.FCustomerName & "님의 ["& oCsAction.GetAsDivCDName &"] 처리가 "& oCsAction.FCurrStateName &" 되었습니다."
	else
		strMailTitle = "[텐바이텐] "& oCsAction.FCustomerName & "님께서 요청하신 ["& oCsAction.GetAsDivCDName &"] 처리가 "& oCsAction.FCurrStateName &" 되었습니다."
	end if


	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	IF oCsAction.FBuyEmail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "텐바이텐"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oCsAction.FCustomerName
		oMail.ReceiverMail	= oCsAction.FBuyEmail
		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()		'TMS메일러
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

	'// 테스트
'	set oMail = New MailCls
'
'	IF oCsAction.FBuyEmail<>"" THEN
'		oMail.MailTitles	= strMailTitle
'		oMail.SenderNm		= "텐바이텐"
'		oMail.SenderMail	= "customer@10x10.co.kr"
'		oMail.AddrType		= "string"
'		oMail.ReceiverNm	= oCsAction.FCustomerName
'		oMail.ReceiverMail	= "headab@naver.com"    ''oCsAction.FBuyEmail
'		oMail.MailConts 	= strMailHTML
'		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
'		oMail.Send_TMSMailer()		'TMS메일러
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
		strMailTitle = "[텐바이텐]"& oCsAction.FCustomerName & "님의 ["& oCsAction.GetAsDivCDName &"] 처리가 "& oCsAction.FCurrStateName &" 되었습니다."
	else
		strMailTitle = "[텐바이텐]"& oCsAction.FCustomerName & "님께서 요청하신 ["& oCsAction.GetAsDivCDName &"] 처리가 "& oCsAction.FCurrStateName &" 되었습니다."
	end if

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	IF oCsAction.FBuyEmail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "텐바이텐"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oCsAction.FCustomerName
		oMail.ReceiverMail	= oCsAction.FBuyEmail
		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()		'TMS메일러
		'oMail.Send_Mailer()
		''oMail.Send_CDO()

	End IF

	SET oMail = nothing

	'// 테스트
'	set oMail = New MailCls
'
'	IF oCsAction.FBuyEmail<>"" THEN
'		oMail.MailTitles	= strMailTitle
'		oMail.SenderNm		= "텐바이텐"
'		oMail.SenderMail	= "customer@10x10.co.kr"
'		oMail.AddrType		= "string"
'		oMail.ReceiverNm	= oCsAction.FCustomerName
'		oMail.ReceiverMail	= "headab@naver.com"    ''oCsAction.FBuyEmail
'		oMail.MailConts 	= strMailHTML
'		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
'		oMail.Send_TMSMailer()		'TMS메일러
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
		strMailTitle = "[텐바이텐]"& oCsAction.FCustomerName & "님의 ["& oCsAction.GetAsDivCDName &"] 처리가 "& oCsAction.FCurrStateName &" 되었습니다."
	else
		strMailTitle = "[텐바이텐]"& oCsAction.FCustomerName & "님께서 요청하신 ["& oCsAction.GetAsDivCDName &"] 처리가 "& oCsAction.FCurrStateName &" 되었습니다."
	end if

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	IF oCsAction.FBuyEmail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "텐바이텐"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oCsAction.FCustomerName
		oMail.ReceiverMail	= oCsAction.FBuyEmail
		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()		'TMS메일러
		'oMail.Send_Mailer()
	End IF

	SET oMail = nothing

	'// 테스트
''	set oMail = New MailCls
''
''	IF oCsAction.FBuyEmail<>"" THEN
''		oMail.MailTitles	= strMailTitle
''		oMail.SenderNm		= "텐바이텐"
''		oMail.SenderMail	= "customer@10x10.co.kr"
''		oMail.AddrType		= "string"
''		oMail.ReceiverNm	= oCsAction.FCustomerName
''		oMail.ReceiverMail	= "headab@naver.com"    ''oCsAction.FBuyEmail
''		oMail.MailConts 	= strMailHTML
'		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
'		oMail.Send_TMSMailer()		'TMS메일러
''		'oMail.Send_Mailer()
''
''	End IF
''
''	SET oMail = nothing

    Set oCsAction = Nothing
end function

'/직원어드민 : [CS]배송관리>>미출고리스트 NEW, 업체어드민 : 주문 관리 > 미출고 리스트	' 2017.12.20 한용민
function SendMiChulgoMailWithMessage(idx, mailmessage)
    ''require /lib/classes/cscenter/oldmisendcls.asp
	''require /lib/classes/order/upchebeasongcls.asp
    dim oneMisend
    dim strMailHTML,strMailTitle, contentsHtml
	strMailHTML = ""
	strMailTitle = "[텐바이텐] 출고 지연 안내메일입니다."

    set oneMisend = new COldMiSend
    oneMisend.FRectDetailIDx = idx
	oneMisend.FRectForMail = "Y"
    oneMisend.getOneOldMisendItem

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "텐바이텐"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oneMisend.FOneItem.FBuyname
		oMail.ReceiverMail	= oneMisend.FOneItem.FBuyEmail
		oMail.MailType = "22"
		strMailHTML = oMail.getMailTemplate
		''parsing
		strMailHTML = replace(strMailHTML,":mailtitle:", "출고지연 안내메일")		' 이메일제목

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
	        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*본 메일은 해당 판매자가 고객님께 보내드리는 메일입니다.<br>*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
		else
	        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
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
		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()
		'oMail.Send_Mailer()
		'oMail.Send_CDO
	End IF

    ''메모에 저장.
    contentsHtml = replace(contentsHtml,"발송예정일","발송예정일("&oneMisend.FOneItem.FMisendipgodate&")")
	Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)

	SET oMail = nothing
	set oneMisend = Nothing
end function

' [업체]게시판 답변 작성 이메일 발송	' 2020-03-12 한용민 생성
Function SendUpheBoardMail(idx)
	dim boardqna

	if idx="" or isnull(idx) then exit Function

	set boardqna = New CUpcheQnADetail
		boardqna.FRectIdx = idx
		boardqna.read()

    dim oCsAction,strMailHTML,strMailTitle
	Set oCsAction = New CsActionMailCls
	strMailHTML = oCsAction.makeMailTemplate_UpheBoard(idx)
	strMailTitle = "[텐바이텐] "& boardqna.Fusername &"님 업체게시판에 문의하신 글에 답변이 등록 되었습니다."

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	IF boardqna.femail<>"" THEN
		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "텐바이텐"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= boardqna.Fusername
		oMail.ReceiverMail	= boardqna.femail
		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 13		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()		'TMS메일러
		'oMail.Send_Mailer()		'EMS메일러
		''oMail.Send_CDO()
	End IF

	SET oMail = nothing
	set boardqna = nothing
    Set oCsAction = Nothing
End Function

' 사용안하는듯		'/ [CS]고객센터>>[CS]미처리CS리스트
function SendMiChulgoMail_CS(csdetailidx)
    ''require /lib/classes/cscenter/cs_mifinishcls.asp
    dim oneMisend
    dim strMailHTML,strMailTitle, contentsHtml
	strMailHTML = ""
	strMailTitle = "[텐바이텐] CS출고 지연 안내메일입니다."

    set oneMisend = new CCSMifinishMaster
        oneMisend.FRectCSDetailIDx = csdetailidx
        oneMisend.getOneMifinishItem

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "텐바이텐"
		'oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oneMisend.FOneItem.FBuyname
		oMail.ReceiverMail	= oneMisend.FOneItem.FBuyEmail
		oMail.MailType = "22"
		strMailHTML = oMail.getMailTemplate
		''parsing
		strMailHTML = replace(strMailHTML,":mailtitle:", "출고지연 안내메일")		' 이메일제목

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
		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*본 메일은 해당 판매자가 고객님께 보내드리는 메일입니다.<br>*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
		    else
		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*본 메일은 해당 판매자가 고객님께 보내드리는 메일입니다.")
		    end if
		else
		    if (oneMisend.FOneItem.FMifinishReason<>"07") then
		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
		    else
		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","")
		    end if
		end if

		if (oneMisend.FOneItem.FMifinishipgodate<>"") then
    		if (oneMisend.FOneItem.FMifinishReason="03") then
    		    if (oneMisend.FOneItem.getMifinishDPlusDate>1) then
    		        contentsHtml = "안녕하세요.   고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 요청하신 상품이 발송이 지연될 예정입니다.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			else
        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			end if
        			contentsHtml = contentsHtml & "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>"
        			contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
        			contentsHtml = contentsHtml & "쇼핑에 불편을 드린 점 진심으로 사과드리며, 기분좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.<br>"
    		    else
        		    contentsHtml = "안녕하세요.   고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 요청하신 상품의 출고안내 메일입니다.<br>"
        			contentsHtml = contentsHtml & "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>"
        			contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"

    		    end if
    		elseif (oneMisend.FOneItem.FMifinishReason="02") then
    		    if (oneMisend.FOneItem.getMifinishDPlusDate>1) then
    		        contentsHtml = "안녕하세요.  고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 요청하신 상품은 주문 후 제작(수입)되는 상품으로<br>"
        			contentsHtml = contentsHtml & "일반상품과 달리 주문제작(수입)에 기간이 소요되는 상품입니다.<br>"
        			contentsHtml = contentsHtml & "아래와 같이 발송예정일을 안내해드리오니,<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			else
        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			end if
    		    else
        		    contentsHtml = "안녕하세요.  고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 요청하신 상품의 출고안내 메일입니다.<br>"
        			contentsHtml = contentsHtml & "아래와 같이 발송예정일을 안내해 드립니다.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			else
        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			end if
    			end if
    	    elseif (oneMisend.FOneItem.FMifinishReason="04") then
    	        oMail.MailTitles = "[텐바이텐] CS출고 예정 안내메일입니다."

    		    if (oneMisend.FOneItem.getMifinishDPlusDate>1) then
    		        contentsHtml = "안녕하세요.  고객님<br>"
                    contentsHtml = contentsHtml & "고객님께서 요청하신 상품의 출고안내메일입니다.<br>"
                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>예약배송상품</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"


    		    else
        		    contentsHtml = "안녕하세요.  고객님<br>"
                    contentsHtml = contentsHtml & "고객님께서 요청하신 상품의 출고안내메일입니다.<br>"
                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>예약배송상품</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"


    			end if
    		elseif (oneMisend.FOneItem.FMifinishReason="07") then
    	        oMail.MailTitles = "[텐바이텐] CS출고 예정 안내메일입니다."

    		    if (oneMisend.FOneItem.getMifinishDPlusDate>1) then
    		        contentsHtml = "안녕하세요.  고객님<br>"
                    contentsHtml = contentsHtml & "고객님께서 요청하신 상품의 출고안내메일입니다.<br>"
                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>고객지정배송</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"


    		    else
        		    contentsHtml = "안녕하세요.  고객님<br>"
                    contentsHtml = contentsHtml & "고객님께서 요청하신 상품의 출고안내메일입니다.<br>"
                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>고객지정배송</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"

    			end if
    		end if
		end if
		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)

		oMail.MailConts 	= strMailHTML
		'response.write strMailHTML
		'response.end
		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()		'TMS메일러
		'oMail.Send_Mailer()
		'oMail.Send_CDO
	End IF

    ''메모에 저장.
    contentsHtml = replace(contentsHtml,"발송예정일","발송예정일("&oneMisend.FOneItem.FMifinishipgodate&")")
	call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)

	SET oMail = nothing
	set oneMisend = Nothing
end function

'/오프라인 매장 매장배송쪽에 쓰임. 차후 사용시 SendMiChulgoMailWithMessage 펑션과 비슷하게 고쳐서 사용할것. 탬플릿이 옛날꺼임.
function SendMiChulgoMail_off(detailidx)
    dim oneMisend ,strMailHTML,strMailTitle, contentsHtml

	strMailHTML = ""
	strMailTitle = "[텐바이텐] 출고 지연 안내메일입니다."

    set oneMisend = new cupchebeasong_list
    oneMisend.FRectDetailIDx = detailidx
    oneMisend.fOneOldMisendItem

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "텐바이텐"
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
		    strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*본 메일은 해당 판매자가 고객님께 보내드리는 메일입니다.<br>*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
		else
		    strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
		end if

		if (oneMisend.FOneItem.FMisendipgodate<>"") then
    		if (oneMisend.FOneItem.FMisendReason="03") then
    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "안녕하세요.   고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품이 발송이 지연될 예정입니다.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			else
        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			end if
        			contentsHtml = contentsHtml & "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>"
        			contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
        			contentsHtml = contentsHtml & "쇼핑에 불편을 드린 점 진심으로 사과드리며, 기분좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.<br>"
    		    else
        		    contentsHtml = "안녕하세요.   고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내 메일입니다.<br>"
        			contentsHtml = contentsHtml & "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>"
        			contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
    		    end if
    		elseif (oneMisend.FOneItem.FMisendReason="02") then
    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "안녕하세요.  고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품은 주문 후 제작(수입)되는 상품으로<br>"
        			contentsHtml = contentsHtml & "일반상품과 달리 주문제작(수입)에 기간이 소요되는 상품입니다.<br>"
        			contentsHtml = contentsHtml & "아래와 같이 발송예정일을 안내해드리오니,<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			else
        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			end if
    		    else
        		    contentsHtml = "안녕하세요.  고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내 메일입니다.<br>"
        			contentsHtml = contentsHtml & "아래와 같이 발송예정일을 안내해 드립니다.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			else
        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			end if
    			end if
    	    elseif (oneMisend.FOneItem.FMisendReason="04") then
    	        oMail.MailTitles = "[텐바이텐] 출고 예정 안내메일입니다."

    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "안녕하세요.  고객님<br>"
                    contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내메일입니다.<br>"
                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>예약배송상품</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
    		    else
        		    contentsHtml = "안녕하세요.  고객님<br>"
                    contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내메일입니다.<br>"
                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>예약배송상품</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
    			end if
    		end if
		end if
		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)

		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
		oMail.Send_TMSMailer()		'TMS메일러
		'oMail.Send_Mailer()
		'oMail.Send_CDO
	End IF

    ''메모에 저장.
    'contentsHtml = replace(contentsHtml,"발송예정일","발송예정일("&oneMisend.FOneItem.FMisendipgodate&")")
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

    Dim FRectForceCurrState     ''상태값 강제 지정.
    Dim FRectForceBuyEmail      ''이메일 강제지정.

    Dim Faccountdiv      ''2016/08/05 추가
    Dim Fpggubun         ''2016/08/05 추가

 	public function GetAsDivCDName()
        GetAsDivCDName = db2html(FDivCDName)
	end function

	public function GetCauseDetailString()
        GetCauseDetailString = Fgubun02Name
    end function

	Public Sub GetOneCSASMaster(FRectCsAsID)
		dim tmpZipCode, tmpaddress1, tmpaddress2
			tmpZipCode="11154"
			tmpaddress1="경기도 포천시 군내면 용정경제로2길 83"
			tmpaddress2="텐바이텐 물류센터"

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
				" 	,isNull(p.company_name,'(주)텐바이텐') as ReturnName ,isNull(p.deliver_phone,'1644-6030') as ReturnPhone ,isNull(p.deliver_hp,'') as ReturnHP "&_
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

				IF (FRectForceCurrState<>"") then  ''상태값 강제 지정 (메일 재발송시 사용.)
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
                        FReturnMethodName = "네이버페이취소"
                    end if
                end if
			END IF
		rsget.close

		''기타 안내 사항
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

				IF (FRectForceCurrState<>"") then  ''상태값 강제 지정 (메일 재발송시 사용.)
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

		''기타 안내 사항
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

	''// 메일 헤더 이미지
	Public Function getMailHeadImage()
		dim tmpImg
		IF FDivCD="A000" Then '// 맞교환출고
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_exchange.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_exchange_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A001" Then '// 누락재발송
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_resend.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_resend_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A002" Then '// 서비스발송
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_send_service.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_send_service_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A003" Then '// 환불요청
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_refund.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_refund_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A004" Then '// 반품접수(업)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_return.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_return_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A007" Then '// 신용/이체취소
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_payment_cancel.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_payment_cancel_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A008" Then '// 주문취소
			IF FCurrState="B001" Then
				'tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_order_cancel.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_order_cancel_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A010" Then '// 회수신청(텐)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_prd_recovery.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_prd_recovery_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A011" Then '// 맞교환회수(텐)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_cancel_prd_recovery.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_cancel_prd_recovery_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A900" Then '// 주문내역변경
			IF FCurrState="B001" Then
				'tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_change_order.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_change_order_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSE

		END IF
		getMailHeadImage = tmpImg
	End Function

	''// 메일 제목	'2017.12.19 한용민 생성
	Public Function getMailHeadtitle()
		dim tmptitle
		IF FDivCD="A000" Then '// 맞교환출고
			IF FCurrState="B001" Then
				tmptitle = "교환출고 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "교환출고 완료 안내메일"
			End IF
		ELSEIF FDivCD="A001" Then '// 누락재발송
			IF FCurrState="B001" Then
				tmptitle = "누락재발송 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "누락재발송 완료 안내메일"
			End IF
		ELSEIF FDivCD="A002" Then '// 서비스발송
			IF FCurrState="B001" Then
				tmptitle = "서비스 발송 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "서비스 발송 완료 안내메일"
			End IF
		ELSEIF FDivCD="A003" Then '// 환불요청
			IF FCurrState="B001" Then
				tmptitle = "환불 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "환불 완료 안내메일"
			End IF
		ELSEIF FDivCD="A004" Then '// 반품접수(업)
			IF FCurrState="B001" Then
				tmptitle = "반품 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "반품 완료 안내메일"
			End IF
		ELSEIF FDivCD="A007" Then '// 신용/이체취소
			IF FCurrState="B001" Then
				tmptitle = "결제취소 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "결제취소 완료 안내메일"
			End IF
		ELSEIF FDivCD="A008" Then '// 주문취소
			IF FCurrState="B001" Then
				'tmptitle = "주문취소 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "주문취소 완료 안내메일"
			End IF
		ELSEIF FDivCD="A010" Then '// 회수신청(텐)
			IF FCurrState="B001" Then
				tmptitle = "상품회수 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "상품회수 완료 안내메일"
			End IF
		ELSEIF FDivCD="A011" Then '// 맞교환회수(텐)
			IF FCurrState="B001" Then
				tmptitle = "교환상품 회수 접수 메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "교환상품 회수 완료 메일"
			End IF
		ELSEIF FDivCD="A900" Then '// 주문내역변경
			IF FCurrState="B001" Then
				'tmptitle = "주문내역변경 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "주문내역변경 완료 안내메일"
			End IF
		ELSE

		END IF
		getMailHeadtitle = tmptitle
	End Function

	'// 기타 안내사항		'/2017.12.19 한용민
	Public Function getEtcNotice()
		dim tmpHTML

        getEtcNotice = ""

        if (Trim(FInfoHtml)="") then Exit function

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; text-align:left;'>기타 안내 정보</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:10px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""돋움"", sans-serif; color:#707070; vertical-align:top; text-align:left;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:630px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
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

	''// 택배 정보 가져오기		'/2017.12.19 한용민
	Function getDlvInfo()
		dim tmpHTML
		tmpHTML=""

        if (IsNULL(FSongjangNo)) or (FSongjangNo="") then Exit function

		'A001(누락재발송), A011(맞교환회수(텐)), A000(맞교환출고), A010(회수신청(텐)), A002(서비스발송), A004(반품접수(업))
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>택배정보</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf

			IF FSongjangNo<>"" then
				tmpHTML=tmpHTML&"									<span style='margin:0; text-align:left; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>"& FSongjangDivName &"</span>"& vbcrlf
				tmpHTML=tmpHTML&"									&nbsp;&nbsp;<a href='"& DeliverDivTrace(Trim(FSongjangDiv)) & FSongjangNo &"' target='_blank' style='margin:0; padding:0; font-size:12px; color:#dd5555; font-size:11px; line-height:18px; color:#0066cc; text-align:left;'>"& FSongjangNo &"</a>" & vbcrlf
			ELSE
				tmpHTML=tmpHTML&"									<span style='margin:0; text-align:left; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>택배정보가 등록되지 않았습니다.</span>" & vbcrlf
			END IF

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		END IF

		getDlvInfo =  tmpHTML
	END Function

	'// 처리 결과 가져오기		'/2017.12.19 한용민
	Function getFinishResult()
		dim tmpHTML
		tmpHTML=""

		'A001(누락재발송), A011(맞교환회수(텐)), A010(회수신청(텐)), A003(환불요청), A002(서비스발송), A004(반품접수(업)), A000(맞교환출고)
		IF FCurrState="B007" THEN
		    ''처리 내역이 없을때..
		    if (FOpenContents="") then
		        if (FDivCD="A000") then
		            FOpenContents = "맞교환상품 출고완료"
		        elseif (FDivCD="A001") then
		            FOpenContents = "누락상품 출고완료"
		        elseif (FDivCD="A002") then
		            FOpenContents = "상품 출고완료"
		        elseif (FDivCD="A003") then

		        elseif (FDivCD="A004") then
		            FOpenContents = "상품 반품(회수)완료" '' / 환불등록"

		        elseif (FDivCD="A010") then
		            FOpenContents = "상품 회수완료" '' / 환불등록"
		        elseif (FDivCD="A011") then
		            FOpenContents = "맞교환상품 회수완료"
		        else

		        end if
		    end if

			tmpHTML=tmpHTML&"	<tr>" & vbcrlf
			tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
			tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<th style='margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; text-align:left; font-family:dotum, ""돋움"", sans-serif; text-align:left;'>처리결과 정보<th>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
			tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>처리완료일</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& FFinishDate &"</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			IF (Trim(FOpenContents)<>"") then
				tmpHTML=tmpHTML&"							<tr>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>처리내용</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:16px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& nl2br(FOpenContents) &"</td>" & vbcrlf
				tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			end IF

			''// 택배정보 가져오기
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

	''//환불정보 가져오기		'/2017.12.19 한용민
	Function getRefundInfo()
		dim tmpHTML
		tmpHTML=""

		' A008(주문취소), A010(회수신청(텐)), A007(신용/이체취소), A003(환불요청), A004(반품접수(업))
		IF FDivCD="A003" or FDivCD="A004" or FDivCD="A007" or FDivCD="A008" or FDivCD="A010" THEN
		    ''환불액0이면 return
		    if (FRefundRequire=0) then Exit function

		    ''부정확한 환불 정보 제거
		    if (FReturnMethod="R007") then
		        if (Len(Replace(FReBankAccount,"-",""))<7) then
    		        FReBankName = ""
    		        FReBankAccount = "계좌확인요망"
    		        FReBankOwnerName =""
    		    else
    		        FReBankAccount = Left(FReBankAccount,Len(Trim(FReBankAccount))-3) + "***"
    		    end if
		    end if

			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>환불예정액</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FormatNumber(FRefundRequire,0) &" 원" & vbcrlf

			'배송비차감 안내가 오해를 일으켜 표시안하게 수정
			'if (FRefundDeliveryPay<>0) then
			'			tmpHTML=tmpHTML&"									(배송비차감 : " & FormatNumber(FRefundDeliveryPay+Frefundbeasongpay,0) &")" & vbcrlf
			'end if

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>환불정보(계좌)</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FReturnMethodName &"&nbsp;&nbsp;" & vbcrlf

			IF (FReturnMethod="R007") THEN
				tmpHTML=tmpHTML&"									"& FReBankName &"&nbsp;&nbsp; " & vbcrlf
				tmpHTML=tmpHTML&"									"& FReBankAccount &"&nbsp;&nbsp; " & vbcrlf
				tmpHTML=tmpHTML&"									"& AstarUserName(FReBankOwnerName) &" " & vbcrlf
			ELSEIF (FReturnMethod="R900") THEN
				tmpHTML=tmpHTML&"									(적립아이디 : "& FUserID &") " & vbcrlf
			ELSEIF (FReturnMethod="R100") or (FReturnMethod="R550") or (FReturnMethod="R560") or (FReturnMethod="R120") or (FReturnMethod="R020") or (FReturnMethod="R080") THEN
				if (Left(FPayGateTid,6)="IniTec") and (FCurrState="B007") and (FReturnMethod<>"R120") then
					tmpHTML=tmpHTML&"									<a target='_blank' href=https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid="& FPayGateTid &"&noMethod=1>[매출전표출력]</a> " & vbcrlf
				end if
				if (FReturnMethod = "R550") or (FReturnMethod = "R560") then
					tmpHTML=tmpHTML&"									기프팅/기프티콘 은 발행시 증빙서류가 발급되며, 상품구매시에는 발행불가입니다." & vbcrlf
				end if
			END IF

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		END IF
		getRefundInfo = tmpHTML
	END Function

	''//업체 주소 가져오기		'/2017.12.19 한용민
	Function getReturnInfo()
		dim tmpHTML
		tmpHTML=""

		' A011(맞교환회수(텐)), A010(회수신청(텐)), A004(반품접수(업))
		IF FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>반품회수주소</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FReturnName &" &nbsp; &nbsp;"& FReturnPhone &"<br />["& FReturnZipCode &"] "& FReturnZipAddr &"&nbsp;"& FReturnEtcAddr &"" & vbcrlf
			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			if (FReturnName<>"(주)텐바이텐") and (FupcheReturnSongjangDivName<>"") and (Left(FupcheReturnSongjangDivTel,1)="1" or Left(FupcheReturnSongjangDivTel,1)="0") then
				tmpHTML=tmpHTML&"							<tr>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>이용택배사</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
				tmpHTML=tmpHTML&"									"& FupcheReturnSongjangDivName &"<br />택배사연락처 : "& FupcheReturnSongjangDivTel &"" & vbcrlf
				tmpHTML=tmpHTML&"								</td>" & vbcrlf
				tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			END IF
		END IF

		getReturnInfo = tmpHTML
	END Function

	''//고객주소 가져오기		'/2017.12.19 한용민
	Function getReqInfo()
		dim tmpHTML
		tmpHTML=""

		'A001(누락재발송), A011(맞교환회수(텐)), A000(맞교환출고), A010(회수신청(텐)), A002(서비스발송)
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A010" THEN 'or FDivCD="A011"
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>고객주소</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& AstarUserName(trim(FReqName)) &" 고객님 &nbsp; &nbsp;"& AstarPhoneNumber(trim(FReqPhone)) &" / "& AstarPhoneNumber(trim(FReqHP)) &" <br />["& printUserId(trim(FReqZipcode), 2, "*") &"] "& printUserId(trim(FReqZipAddr), 2, "*") &"&nbsp;(이하생략)" & vbcrlf		' FReqEtcAddr
			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		END IF
		getReqInfo = tmpHTML
	END Function

	''//접수 상품 정보 가져오기		'/2017.12.19 한용민
	Function getAsItemLIst()
		dim tmpHTML
		dim OCsDetail,i

		tmpHTML = ""

		'A001(누락재발송), A008(주문취소), A011(맞교환회수(텐)), A000(맞교환출고), A010(회수신청(텐)), A002(서비스발송), A004(반품접수(업))
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
			tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; text-align:left; color:#000;'>접수 상품 정보</th>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
			tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:12px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:50px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, ""돋움"", sans-serif; text-align:center; color:#707070; font-size:12px; line-height:12px;'>상품</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:100px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>상품코드</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:295px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>상품명[옵션]</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:right; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>판매가격</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:25px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>&nbsp;</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>수량</th>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			IF OCsDetail.FresultCount>0 Then
				FOR i=0 TO OCsDetail.FResultCount-1
				    IF (OCsDetail.FItemList(i).Fitemid<>0) or (OCsDetail.FItemList(i).Fitemcost<>0) then
						tmpHTML=tmpHTML&"							<tr>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:50px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea;'>" & vbcrlf
						tmpHTML=tmpHTML&"									<img src='"& OCsDetail.FItemList(i).FSmallImage &"' alt='' />" & vbcrlf
						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf
						tmpHTML=tmpHTML&"									"& OCsDetail.FItemList(i).Fitemid &"" & vbcrlf
						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf

						IF (OCsDetail.FItemList(i).Fitemid=0) Then
							tmpHTML=tmpHTML&"									배송비" & vbcrlf
						ELSE
							tmpHTML=tmpHTML&"									"& OCsDetail.FItemList(i).Fitemname &"" & vbcrlf
						END IF
						if ( OCsDetail.FItemList(i).Fitemoptionname <>"") then
							tmpHTML=tmpHTML&"									["& OCsDetail.FItemList(i).Fitemoptionname &"]" & vbcrlf
						END IF

						tmpHTML=tmpHTML&"								</td>" & vbcrlf

						tmpHTML=tmpHTML&"								<td style='width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, ""돋움"", sans-serif;'>" & vbcrlf

						IF (OCsDetail.FItemList(i).FdiscountAssingedCost<>0) and (OCsDetail.FItemList(i).Fitemcost>OCsDetail.FItemList(i).FdiscountAssingedCost) then
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:6px 0; font-size:11px; line-height:16px; color:#707070; font-family:dotum, ""돋움"", sans-serif; text-decoration:line-through; text-align:right;'>"& FormatNumber(OCsDetail.FItemList(i).Fitemcost,0) & "원</span><br />" & vbcrlf
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, ""돋움"", sans-serif; text-align:right;'>" & FormatNumber(OCsDetail.FItemList(i).FdiscountAssingedCost,0) &"원</span>" & vbcrlf
						ELSE
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, ""돋움"", sans-serif; text-align:right;'>"& FormatNumber(OCsDetail.FItemList(i).Fitemcost,0) &"원</span>" & vbcrlf
						END IF

						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>"& OCsDetail.FItemList(i).Fregitemno &"</td>" & vbcrlf
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

	''// 접수 기본 내용 가져오기		'/2017.12.19 한용민
	Function getAsInfo()
		dim tmpHTML

		tmpHTML = ""
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>접수일</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& FRegDate &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>서비스코드</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& FAsID &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>주문번호</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& FOrderSerial &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>접수내용</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									"& FTitle &" <a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp?orderserial=" & FOrderSerial & "' target='_blank'>[상세내역보기]</a>" & vbcrlf
		tmpHTML=tmpHTML&"								</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>접수사유</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& GetCauseDetailString &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf

		getAsInfo =tmpHTML
	END Function

	'// SMS보내기
	Public Function sendSMS(byval ipHp, byval ipText)
		dim tmpSms,strSQL
		dim RcvHp,RcvMsg

		'// 직적 입력된 정보 없을경우 자동 생성
		IF ipHp<>"" THEN
			RcvHp=ipHp
		ELSE
			RcvHp=FBuyHP
		END IF

		IF ipText<>"" THEN
			RcvMsg=ipText
		ELSE
			RcvMsg="[텐바이텐] 요청하신 ["& GetAsDivCDName &"] 처리가 "& FCurrStateName &" 되었습니다."
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
			response.write "SMS 발송 - 완료"
			Exit Function
		ELSE
			dbget.RollBackTrans
			response.write "SMS 발송 - 실패"
			Exit Function
		EnD IF

	End Function

	'// 이메일 탬플릿 가져와서 만드는걸로 생성.	2017.12.18 한용민 생성
	Function makeMailTemplate(id)
		dim tmpHTML, fs, dirPath, fileName, objFile, mailheader, mailfooter

		Call GetOneCSASMaster(id) '// 값 세팅

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_header_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailheader = objFile.readall	' 헤더

		tmpHTML=mailheader

		'A001(누락재발송), A008(주문취소), A000(맞교환출고)
		tmpHTML=tmpHTML&" <table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='height:253px; text-align:center;'>"& getMailHeadImage &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:20px 0 45px; font-size:12px; line-height:22px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:center;'>" & vbcrlf

		if (FDivCD = "A008") then
			tmpHTML=tmpHTML&"			고객님의 "& GetAsDivCDName &"이 정상적으로 처리가 "& FCurrStateName &" 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		elseif (FDivCD = "A000") then
			tmpHTML=tmpHTML&"			고객님이 요청하신 "& GetAsDivCDName &"가 정상적으로 처리가 "& FCurrStateName &" 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		else
			tmpHTML=tmpHTML&"			고객님이 요청하신 "& GetAsDivCDName &"이 정상적으로 처리가 "& FCurrStateName &" 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		end if

		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:0 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%; background:#f8f8f8; text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:297px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""돋움"", sans-serif; text-align:right;'>주문번호 :</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:7px; padding:34px 0;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:331px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""돋움"", sans-serif; color:#dd5555; text-align:left; letter-spacing:-1px;'>" & vbcrlf
		tmpHTML=tmpHTML&"						"& FOrderSerial &"" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// 처리결과 가져오기
		tmpHTML=tmpHTML& getFinishResult()

		''// 접수 상품 정보 가져오기
		tmpHTML=tmpHTML& getAsItemLIst()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; text-align:left; color:#000;'>접수 정보</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf

		''// 접수 기본 내용 가져오기
		tmpHTML=tmpHTML& getAsInfo()

		''// 고객주소 가져오기
		tmpHTML=tmpHTML& getReqInfo()

		''// 업체주소 가져오기
		tmpHTML=tmpHTML& getReturnInfo()

		''// 환불정보 가져오기
		tmpHTML=tmpHTML& getRefundInfo()

		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// 기타 안내사항
		tmpHTML=tmpHTML& getEtcNotice()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 104px; margin:0;text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_receiption_info.png' alt='접수 정보 상세보기' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_go_shopping.png' alt='텐바이텐 쇼핑하기' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:25px 0; border-top:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:center;'>끝까지 기분 좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"</table>" & vbcrlf

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' 푸터

		tmpHTML=tmpHTML&mailfooter

		tmpHTML = replace(tmpHTML,":mailtitle:", getMailHeadtitle())		' 이메일제목

		makeMailTemplate = tmpHTML
	End Function

	'/ 이메일 탬플릿 가져와서 만드는걸로 생성.	2017.12.18 한용민 생성
	Function makeMailTemplate_GiftCard(id)
		dim tmpHTML, fs, dirPath, fileName, objFile, mailheader, mailfooter

		Call GetOneCSASMaster_GiftCard(id) '// 값 세팅

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_header_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailheader = objFile.readall	' 헤더

		tmpHTML=mailheader

		'A001(누락재발송), A008(주문취소), A000(맞교환출고)
		tmpHTML=tmpHTML&" <table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='height:253px; text-align:center;'>"& getMailHeadImage &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:20px 0 45px; font-size:12px; line-height:22px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:center;'>" & vbcrlf

		if (FDivCD = "A008") then
			tmpHTML=tmpHTML&"			고객님의 "& GetAsDivCDName &"이 정상적으로 처리가 "& FCurrStateName &" 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		elseif (FDivCD = "A000") then
			tmpHTML=tmpHTML&"			고객님이 요청하신 "& GetAsDivCDName &"가 정상적으로 처리가 "& FCurrStateName &" 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		else
			tmpHTML=tmpHTML&"			고객님이 요청하신 "& GetAsDivCDName &"이 정상적으로 처리가 "& FCurrStateName &" 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		end if

		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:0 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%; background:#f8f8f8; text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:297px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""돋움"", sans-serif; text-align:right;'>주문번호 :</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:7px; padding:34px 0;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:331px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""돋움"", sans-serif; color:#dd5555; text-align:left; letter-spacing:-1px;'>" & vbcrlf
		tmpHTML=tmpHTML&"						"& FOrderSerial &"" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// 처리결과 가져오기
		tmpHTML=tmpHTML& getFinishResult()

		''// 접수 상품 정보 가져오기
		'tmpHTML=tmpHTML& getAsItemLIst()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; text-align:left; color:#000;'>접수 정보</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf

		''// 접수 기본 내용 가져오기
		tmpHTML=tmpHTML& getAsInfo()

		''// 고객주소 가져오기
		'tmpHTML=tmpHTML& getReqInfo()

		''// 업체주소 가져오기
		'tmpHTML=tmpHTML& getReturnInfo()

		''// 환불정보 가져오기
		tmpHTML=tmpHTML& getRefundInfo()

		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// 기타 안내사항
		'tmpHTML=tmpHTML& getEtcNotice()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 104px; margin:0;text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_receiption_info.png' alt='접수 정보 상세보기' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_go_shopping.png' alt='텐바이텐 쇼핑하기' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:25px 0; border-top:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:center;'>끝까지 기분 좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"</table>" & vbcrlf

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' 푸터

		tmpHTML=tmpHTML&mailfooter

		tmpHTML = replace(tmpHTML,":mailtitle:", getMailHeadtitle())		' 이메일제목

		makeMailTemplate_GiftCard = tmpHTML
	End Function

	'// 이메일 탬플릿 가져와서 만드는걸로 생성.	2020.03.12 한용민 생성
	Function makeMailTemplate_UpheBoard(id)
		dim tmpHTML, fs, dirPath, fileName, objFile, mailheader, mailfooter, boardqna

		set boardqna = New CUpcheQnADetail
			boardqna.FRectIdx = idx
			boardqna.read()

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_header_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailheader = objFile.readall	' 헤더

		tmpHTML=mailheader

		tmpHTML=tmpHTML&" <table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:20px 0 45px; font-size:12px; line-height:22px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML & boardqna.Fusername &"님 업체게시판에 문의하신 내용이 정상적으로 처리가 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; text-align:left;'>접수내용</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:10px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""돋움"", sans-serif; color:#707070; vertical-align:top; text-align:left;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:630px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									<strong>문의제목</strong><br>"& nl2br(boardqna.Ftitle) &"" & vbcrlf
		tmpHTML=tmpHTML&"									<Br><br><strong>문의내용</strong><br>"& nl2br(boardqna.Fcontents) &"" & vbcrlf
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
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; text-align:left;'>처리결과</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:10px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""돋움"", sans-serif; color:#707070; vertical-align:top; text-align:left;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:630px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									<strong>답변제목</strong><br>"& nl2br(boardqna.Freplytitle) &"" & vbcrlf
		tmpHTML=tmpHTML&"									<Br><br><strong>답변내용</strong><br>"& nl2br(boardqna.Freplycontents) &"" & vbcrlf
		tmpHTML=tmpHTML&"								</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:25px 0; border-top:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:center;'>끝까지 기분 좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"</table>" & vbcrlf

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' 푸터

		tmpHTML=tmpHTML&mailfooter

		tmpHTML = replace(tmpHTML,":mailtitle:", getMailHeadtitle())		' 이메일제목

		makeMailTemplate_UpheBoard = tmpHTML
		set boardqna=nothing
	End Function
End Class

'/ 안쓰는듯한데? 사용안함.
'function SendMiChulgoMail(idx)
'    ''require /lib/classes/cscenter/oldmisendcls.asp
'    dim oneMisend
'    dim strMailHTML,strMailTitle, contentsHtml
'	strMailHTML = ""
'	strMailTitle = "[텐바이텐] 출고 지연 안내메일입니다."
'
'    set oneMisend = new COldMiSend
'    oneMisend.FRectDetailIDx = idx
'    oneMisend.getOneOldMisendItem
'
'	'//=======  메일 발송 =========/
'	dim oMail
'	dim MailHTML
'
'	set oMail = New MailCls         '' mailLib2
'
'	IF oneMisend.FOneItem.Fbuyemail<>"" THEN
'
'		oMail.MailTitles	= strMailTitle
'		oMail.SenderNm		= "텐바이텐"
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
'		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*본 메일은 해당 판매자가 고객님께 보내드리는 메일입니다.<br>*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
'		    else
'		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*본 메일은 해당 판매자가 고객님께 보내드리는 메일입니다.")
'		    end if
'		else
'		    if (oneMisend.FOneItem.FMisendReason<>"07") then
'		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
'		    else
'		        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","")
'		    end if
'		end if
'
'		if (oneMisend.FOneItem.FMisendipgodate<>"") then
'    		if (oneMisend.FOneItem.FMisendReason="03") then
'    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
'    		        contentsHtml = "안녕하세요.   고객님<br>"
'        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품이 발송이 지연될 예정입니다.<br>"
'        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
'        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
'        			else
'        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
'        			end if
'        			contentsHtml = contentsHtml & "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>"
'        			contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
'        			contentsHtml = contentsHtml & "쇼핑에 불편을 드린 점 진심으로 사과드리며, 기분좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.<br>"
'    		    else
'        		    contentsHtml = "안녕하세요.   고객님<br>"
'        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내 메일입니다.<br>"
'        			contentsHtml = contentsHtml & "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>"
'        			contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
'
'    		    end if
'    		elseif (oneMisend.FOneItem.FMisendReason="02") then
'    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
'    		        contentsHtml = "안녕하세요.  고객님<br>"
'        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품은 주문 후 제작(수입)되는 상품으로<br>"
'        			contentsHtml = contentsHtml & "일반상품과 달리 주문제작(수입)에 기간이 소요되는 상품입니다.<br>"
'        			contentsHtml = contentsHtml & "아래와 같이 발송예정일을 안내해드리오니,<br>"
'        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
'        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
'        			else
'        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
'        			end if
'    		    else
'        		    contentsHtml = "안녕하세요.  고객님<br>"
'        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내 메일입니다.<br>"
'        			contentsHtml = contentsHtml & "아래와 같이 발송예정일을 안내해 드립니다.<br>"
'        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
'        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
'        			else
'        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
'        			end if
'    			end if
'    	    elseif (oneMisend.FOneItem.FMisendReason="04") then
'    	        oMail.MailTitles = "[텐바이텐] 출고 예정 안내메일입니다."
'
'    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
'    		        contentsHtml = "안녕하세요.  고객님<br>"
'                    contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내메일입니다.<br>"
'                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>예약배송상품</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
'                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
'                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
'    		    else
'        		    contentsHtml = "안녕하세요.  고객님<br>"
'                    contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내메일입니다.<br>"
'                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>예약배송상품</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
'                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
'                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
'    			end if
'    		elseif (oneMisend.FOneItem.FMisendReason="07") then
'    	        oMail.MailTitles = "[텐바이텐] 출고 예정 안내메일입니다."
'
'    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
'    		        contentsHtml = "안녕하세요.  고객님<br>"
'                    contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내메일입니다.<br>"
'                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>고객지정배송</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
'                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
'                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
'    		    else
'        		    contentsHtml = "안녕하세요.  고객님<br>"
'                    contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내메일입니다.<br>"
'                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>고객지정배송</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
'                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
'                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
'    			end if
'    		end if
'		end if
'		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)
'
'		oMail.MailConts 	= strMailHTML
'
'		oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
'		oMail.Send_TMSMailer()		'TMS메일러
'		'oMail.Send_Mailer()
'		oMail.Send_CDO
'	End IF
'
'    ''메모에 저장.
'    contentsHtml = replace(contentsHtml,"발송예정일","발송예정일("&oneMisend.FOneItem.FMisendipgodate&")")
'	call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)
'
'	SET oMail = nothing
'	set oneMisend = Nothing
'end function

%>
