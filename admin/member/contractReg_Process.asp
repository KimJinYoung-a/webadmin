<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 계약 관리
' Hieditor : 2009.04.07 서동석 생성
'			 2010.05.26 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->

<%
dim mode, makerid, contractType, contractID , contractState , contractEtcContetns , onoffgubun
dim mailfrom, mailto, mailtitle, mailcontent, innerContents ,CurrState,NextState, sendOpenMail
	mode            = request("mode")
	makerid         = request("makerid")
	contractType    = request("contractType")
	contractID      = request("contractID")
	contractEtcContetns = request("contractEtcContetns")
	CurrState       = request("CurrState")
	NextState       = request("NextState")
	sendOpenMail    = request("sendOpenMail")

dim sqlStr , objItem, contractExists , contractContents ,contractNo, contractName, HtmlcontractEtcContetns
dim bufStr, refer
dim ocontract
refer = request.ServerVariables("HTTP_REFERER")

'//신규등록
if (mode="regContract") then

    sqlStr = "select contractContents, contractName ,onoffgubun" +vbcrlf
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
    sqlStr = sqlStr & " where contractType=" & contractType
    
    'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        contractContents = db2Html(rsget("contractContents"))
        contractName = db2Html(rsget("contractName"))
        onoffgubun = rsget("onoffgubun")
    end if
    rsget.Close
	
	'//온라인 계약서만 기존 등록되어 있는 계약서 있는지 Check
	if onoffgubun = "ON" then    
	    sqlStr = "select count(contractID) as cnt from db_partner.dbo.tbl_partner_contract"
	    sqlStr = sqlStr & " where makerid='" & makerid & "'"
	    sqlStr = sqlStr & " and contractType=" & contractType
	    sqlStr = sqlStr & " and contractState>=0"
	    sqlStr = sqlStr & " and contractState<7"
	    rsget.Open sqlStr,dbget,1
	    if Not rsget.Eof then
	        contractExists = rsget("cnt")>0
	    end if
	    rsget.Close
	    
	    if (contractExists) then
	        response.write "<script>alert('이미 진행중인 같은 종류의 계약건이 있습니다.\n기존계약건 삭제또는 완료후 등록가능합니다.');history.back();</script>"
	        dbget.close()	:	response.End
	    end if
	end if
	
    sqlStr = " select * from db_partner.dbo.tbl_partner_contract where 1=0"
    rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	    rsget("makerid")            = makerid       
	    rsget("contractType")       = contractType
        rsget("contractState")      = 0
        rsget("contractName")       = Newhtml2db(contractName)
        rsget("contractEtcContetns")= contractEtcContetns
        rsget("reguserid")          = session("ssBctID")

	rsget.update
	    contractID = rsget("contractID")
	rsget.close
   
    For Each objItem In Request.Form
        ''response.write objItem & "," & Request.Form(objItem) & "<br>"
        if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
            sqlStr = " insert into db_partner.dbo.tbl_partner_contractDetail"
            sqlStr = sqlStr & " (contractID, detailKey, detailValue)"
            sqlStr = sqlStr & " values("
            sqlStr = sqlStr & " " & contractID
            sqlStr = sqlStr & " ,'" & objItem & "'"
            sqlStr = sqlStr & " ,'" & Newhtml2db(Request.Form(objItem)) & "'"
            sqlStr = sqlStr & " )"
            
            dbget.Execute sqlStr
            
            if (objItem="$$CONTRACT_DATE$$") then
                bufStr  = Request.Form(objItem)
                bufStr  = Left(bufStr,4) & "년" & Mid(bufStr,6,2) & "월" & Mid(bufStr,9,2) & "일"
                contractContents = Replace(contractContents,objItem,bufStr)
            else
                contractContents = Replace(contractContents,objItem,Request.Form(objItem))
            end if
            
            if (objItem="$$CONTRACT_DATE$$") then contractNo=Request.Form(objItem)
        end if
    Next
    
    ''기타계약사항
    if Trim(contractEtcContetns)<>"" then
        HtmlcontractEtcContetns = "<p style='margin:0cm;margin-bottom:.0001pt;text-align:justify;text-justify:"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "inter-ideograph;punctuation-wrap:simple;word-break:break-hangul'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<b><span style='font-size:11.0pt;font-family:굴림;color:windowtext;layout-grid-mode:line'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "- 기타계약사항</span></b></p>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<br>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<p class=MsoNormal style='margin-left:5.0pt'><span style='font-size:11.0pt;font-family:굴림;color:windowtext;layout-grid-mode:line'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & replace(contractEtcContetns,VbCrlf,"<br>")
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "</span></p>"
        
        contractContents = Replace(contractContents,"$$ContractEtcContetns$$",HtmlcontractEtcContetns)
    else
        contractContents = Replace(contractContents,"$$ContractEtcContetns$$","")
    end if
    
    ''계약서 번호 생성. YYYYMMDD(계약일)-contractType-contractID
    contractNo = Replace(contractNo,"-","") & "-" & contractType & "-" & contractID
    contractContents = Replace(contractContents,"$$CONTRACT_NO$$",contractNo)

    sqlStr = " update db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
    sqlStr = sqlStr & " ,contractNo='" & contractNo & "'"
    sqlStr = sqlStr & " where contractID=" & contractID
    
    dbget.Execute sqlStr
        
    response.write "<script>alert('등록 되었습니다.\n\n확인 하신후 업체오픈(발송) 으로 변경하시기 바랍니다.');</script>"
    response.write "<script>location.replace('/admin/member/contractReg.asp?makerid="& makerid & "&ContractID=" & ContractID & "');</script>"
    dbget.close()	:	response.End

'//수정
elseif (mode="editContract") then
    ''수정 가능상태 Check
    sqlStr = "select contractType, contractState from db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " where  contractID=" & contractID
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        contractState   = rsget("contractState")
        contractType    = rsget("contractType")
        contractExists = (contractState>=3)
    end if
    rsget.Close
    
    if (contractExists) then
        response.write "<script>alert('수정 가능 상태가 아닙니다.\n관리자 문의 요망.');history.back();</script>"
        dbget.close()	:	response.End
    end if
    
    
    sqlStr = "select t.contractContents, t.contractName from "
    sqlStr = sqlStr & " db_partner.dbo.tbl_partner_contractType t,"
    sqlStr = sqlStr & " db_partner.dbo.tbl_partner_contract c"
    sqlStr = sqlStr & " where c.contractID=" & contractID
    sqlStr = sqlStr & " and c.contractType=t.contractType"
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        contractContents = db2Html(rsget("contractContents"))
        contractName = db2Html(rsget("contractName"))
    end if
    rsget.Close
    
    
    
    sqlStr = " update db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " set contractEtcContetns='" & Newhtml2db(contractEtcContetns) & "'"
    sqlStr = sqlStr & " where contractID=" & contractID
    
    dbget.Execute sqlStr
    
    For Each objItem In Request.Form
        ''response.write objItem & "," & Request.Form(objItem) & "<br>"
        if (Left(objItem,2)="$$") and (Right(objItem,2)="$$") then
            sqlStr = " update db_partner.dbo.tbl_partner_contractDetail"
            sqlStr = sqlStr & " set detailValue='" & Newhtml2db(Request.Form(objItem)) & "'"
            sqlStr = sqlStr & " where contractID=" & contractID
            sqlStr = sqlStr & " and detailKey='" & objItem & "'"
            
            dbget.Execute sqlStr
            
            if (objItem="$$CONTRACT_DATE$$") then
                bufStr  = Request.Form(objItem)
                bufStr  = Left(bufStr,4) & "년" & Mid(bufStr,6,2) & "월" & Mid(bufStr,9,2) & "일"
                contractContents = Replace(contractContents,objItem,bufStr)
            else
                contractContents = Replace(contractContents,objItem,Request.Form(objItem))
            end if
            
            if (objItem="$$CONTRACT_DATE$$") then contractNo=Request.Form(objItem)
        end if
    Next
    
    ''기타계약사항
    if Trim(contractEtcContetns)<>"" then
        HtmlcontractEtcContetns = "<p style='margin:0cm;margin-bottom:.0001pt;text-align:justify;text-justify:"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "inter-ideograph;punctuation-wrap:simple;word-break:break-hangul'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<b><span style='font-size:11.0pt;font-family:굴림;color:windowtext;layout-grid-mode:line'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "- 기타계약사항</span></b></p>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<br>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "<p class=MsoNormal style='margin-left:5.0pt'><span style='font-size:11.0pt;font-family:굴림;color:windowtext;layout-grid-mode:line'>"
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & replace(contractEtcContetns,VbCrlf,"<br>")
        HtmlcontractEtcContetns = HtmlcontractEtcContetns & "</span></p>"
        
        contractContents = Replace(contractContents,"$$ContractEtcContetns$$",HtmlcontractEtcContetns)
    else
        contractContents = Replace(contractContents,"$$ContractEtcContetns$$","")
    end if
    
    ''계약서 번호 생성. YYYYMMDD(계약일)-contractType-contractID
    contractNo = Replace(contractNo,"-","") & "-" & contractType & "-" & contractID
    contractContents = Replace(contractContents,"$$CONTRACT_NO$$",contractNo)

    sqlStr = " update db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " set contractContents='" & Newhtml2db(contractContents) & "'"
    sqlStr = sqlStr & " ,contractNo='" & contractNo & "'"
    sqlStr = sqlStr & " ,contractName='" & Newhtml2db(contractName) & "'"
    if (contractState=-2) then
        sqlStr = sqlStr & " ,contractState=0"
    end if
    sqlStr = sqlStr & " where contractID=" & contractID
    
    dbget.Execute sqlStr
    
    response.write "<script>alert('수정 되었습니다.\n\n확인 하신후 재발송 하시기 바랍니다.');</script>"
    response.write "<script>opener.location.reload(); location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
    
elseif (mode="stateChange") then
    ''CurrState, NextState

    set ocontract = new CPartnerContract
	    ocontract.FRectContractID = ContractID
	    ocontract.FRectMakerid = makerid
	    ocontract.getOneContract
	        
	    contractName = ocontract.FOneItem.FcontractName
	    contractNo   = ocontract.FOneItem.FcontractNo
	    contractType = ocontract.FOneItem.FcontractType
           
    ''온오프 구분인듯
    sqlStr = "select contractContents, contractName ,onoffgubun" +vbcrlf
    sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
    sqlStr = sqlStr & " where contractType=" & contractType
    
    'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        onoffgubun = rsget("onoffgubun")
    end if
    rsget.Close
         
    ''오픈인경우 , 메일발송 Check 되어 있으면 
    if (NextState="1") and (sendOpenMail="on") then
        sqlStr = "select IsNULL(p.email,'') as email from [db_partner].[dbo].tbl_partner p,"
        sqlStr = sqlStr & " db_partner.dbo.tbl_partner_contract c"
        sqlStr = sqlStr & " where c.contractID=" & contractID
        sqlStr = sqlStr & " and c.makerid=p.id"
        
        'response.write sqlStr &"<br>"
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            mailto = db2Html(rsget("email"))
        end if
        rsget.Close
        
        mailto = Trim(mailto)

        if (mailto="") or (InStr(mailto,"@")<0) or (Len(mailto)<8) then
            response.write "<script>alert('업체 담당자 E메일 주소가 유효하지 않습니다.\n브랜드정보에서 E메일 수정 후 사용하시기 바랍니다.');</script>"
            response.write "<script>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if
        
        sqlStr = "select IsNULL(p.usermail,'') as email from db_partner.dbo.tbl_user_tenbyten p,"
        sqlStr = sqlStr & " db_partner.dbo.tbl_partner_contract c"
        sqlStr = sqlStr & " where c.contractID=" & contractID
        sqlStr = sqlStr & " and c.reguserid=p.userid"
        
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            mailfrom = db2Html(rsget("email"))
        end if
        rsget.Close
        
        mailfrom = Trim(mailfrom)
        
        if (mailfrom="") or (InStr(mailfrom,"@")<0) or (Len(mailfrom)<8) then
            response.write "<script>alert('10x10 담당자 E메일 주소가 유효하지 않습니다.마이 정보에서 E메일 수정 후 사용하시기 바랍니다.');</script>"
            response.write "<script>location.replace('" & refer & "');</script>"
            dbget.close()	:	response.End
        end if
             
        mailtitle = "[텐바이텐]업체 계약서가 오픈 되었습니다."
        innerContents = ""
        
        '' 공지에 대한 내용 넣을것
        innerContents = innerContents & " 안녕하세요" & "<br>"
        innerContents = innerContents & "(주)텐바이텐과 좋은 인연으로 만나게 되어 반갑습니다." & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "아래와 같이 계약이 진행되오니 " & "<br>"
        innerContents = innerContents & "계약서 진행사항을 꼼꼼히 읽어주신 후 " & "<br>"
        innerContents = innerContents & "일정에 맞추어 계약서를 우편으로 발송해 주시면 감사하겠습니다." & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "계약서 명 : " & contractName  & "<br>"
        innerContents = innerContents & "계약서 번호 : " & contractNo  & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        
        if onoffgubun = "ON" then
	        innerContents = innerContents & "▶ 계약서 진행대상 : 온라인 입점 전 브랜드 " & "<br>"
	        innerContents = innerContents & "- 오프라인에만 진행하는 브랜드는 대상에서 제외됩니다. " & "<br>"
	        innerContents = innerContents & "(오프라인 입점시에는 오프라인 담당자가 개별적으로 보내드립니다.)" & "<br>"
    	else
       		innerContents = innerContents & "▶ 계약서 진행대상 : 오프라인 입점 전 브랜드 " & "<br>"    	
    	end if
        
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "▶ 계약서 다운방법 " & "<br>"
        innerContents = innerContents & "- 아래 [계약서 다운로드] 클릭하여 다운받은 후 내용확인 및 기재사항을 기재해주세요!!" & "<br>"
        innerContents = innerContents & " (다시 다운로드 받으시려면 어드민 로그인 후 우측 상단 [업체계약서 다운로드]를 이용하여 주세요 )" & "<br>"
        innerContents = innerContents & "- 계약서 날인 하는 방법은 [계약flow 다운로드] 다운받으신 후 그 방법대로 해주시기 바랍니다." 
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<a href='" & manageUrl & "/designer/company/popContract.asp?ContractID=" & ContractID & "' target='_blank'><b><font color=blue>[업체계약서 다운로드]</font></b></a>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<a href='" & manageUrl & "/designer/company/contractflow.ppt' target='_blank'><b><font color=blue>[계약flow 다운로드]</font></b></a>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"

        innerContents = innerContents & "▶ 필수 확인사항 (반드시 두번 세번 확인해주세요) " & "<br>"
        innerContents = innerContents & "수수료, 결제일 이 두가지는 맞는 지 꼭 확인해주셔야 합니다." & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "▶ 업체측에서의 필수 기재사항 (꼭!! 직접 기재하셔야 할 부분)" & "<br>"
        innerContents = innerContents & "- 표지(첫장)의 계약담당자 기재 : 협력업체의 대표이사 또는 계약을 실제로 진행하시는 담당자 성함" & "<br>"

        if (contractType=5) then
			innerContents = innerContents & "- 배송책임자 성함" & "<br>"
        end if

        innerContents = innerContents & "- 마지막 장의 '을'의 대표이사 주민등록번호 및 주소 기재 : 사업자등록증의 대표자 주민번호 및 주소여야 합니다." & "<br>"
        innerContents = innerContents & "- 법인사업자의 경우 마지막장의 대표이사 주민번호 및 주소를 생략하셔도 되며, '갑'사업자 정보 '을' 사업자 정보만 있으면 됩니다." & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "▶ 진행절차 : " & "<br>"
        
        innerContents = innerContents & "① 계약서다운로드" & "<br>"
        innerContents = innerContents & "② 협력업체에서 계약서 확인후 날인 / 2부 우편발송 " & "<br>"
        innerContents = innerContents & "③ 텐바이텐에서 계약서 우편 수령확인" & "<br>"
        innerContents = innerContents & "④ 텐바이텐에서 협력업체로 계약서 1부 발송 / 계약완료" & "<br>"
        
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "▶ 계약서 보내시는 곳" & "<br>"
        
        if onoffgubun = "ON" then     
            innerContents = innerContents & "주소 : (03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐 " & "<br>"
            innerContents = innerContents & "담당자 : " & ocontract.FoneItem.Fusername  & "<br>"
	        innerContents = innerContents & "tel : " & ocontract.FoneItem.Finterphoneno  & " (내선 " & ocontract.FoneItem.Fextension & ") / 직통 : "& ocontract.FoneItem.Fdirect070 &"<br>"
	        innerContents = innerContents & "fax : 02-2179-9244 <br>"	            
        
        '/온라인 계약에 오프라인이 묻어 가는 경우라.. 유동적 당담자 지정 불가능.. 박아넣음
        else
            innerContents = innerContents & "주소 : (03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐 오프라인 사무실 " & "<br>"
            innerContents = innerContents & "담당자 : 이요한 대리<br>"
	        innerContents = innerContents & "tel : 02-554-2033 (내선 222) / 직통 : 070-7515-5422<br>"
            innerContents = innerContents & "fax : 02-2179-9058 <br>"
            innerContents = innerContents & "mail: john6136@10x10.co.kr<br>"
        end if

        innerContents = innerContents & "<br>"    
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "▶ 우편발송시 함께 보내셔야 할 서류" & "<br>"
        innerContents = innerContents & "- 날인된 계약서 2부" & "<br>"
        innerContents = innerContents & "- 결제통장 사본" & "<br>"
        innerContents = innerContents & "- 사업자 등록증 사본" & "<br>"
        innerContents = innerContents & "- 인감증명서 원본 (계약서에 날인한 도장)" & "<br>"
        
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "▶ 기 타 " & "<br>"
        innerContents = innerContents & "텐바이텐 내에 진행하는 브랜드 아이디가 2개 이상일 경우 " & "<br>"
        innerContents = innerContents & "계약서는 각 브랜드 아이디마다 작성을 해주셔야 하며, " & "<br>"
        innerContents = innerContents & "관련서류(사업자등록증,인감증명서,결제통장)은 1부만 주셔도 됩니다. " & "<br>"
        innerContents = innerContents & "어드민 아이디와 패스워드를 부여받지 않으신 경우 담당엠디에게 연락해 주시기 바랍니다. " & "<br>"
               
        innerContents = innerContents & "<br>"
        innerContents = innerContents & "<br>"
        
        innerContents = innerContents & "▶ 계약서 내용 상의 궁금한 점은 각 담당MD에게 문의 하시기 바랍니다."
        
        innerContents = innerContents & "<br>"

        mailcontent = "<html>"
    	mailcontent = mailcontent + "<head>"
    	mailcontent = mailcontent + "<title>텐바이텐 업체 계약서 메일</title>"
    	mailcontent = mailcontent + "<meta http-equiv='Content-Type' content='text/html; charset=euc-kr'>"
    	mailcontent = mailcontent + "<style>"
    	mailcontent = mailcontent + ".text {"
    	mailcontent = mailcontent + ""
    	mailcontent = mailcontent + "font-family: 'Verdana', 'Arial', 'Helvetica', 'sans-serif';"
    	mailcontent = mailcontent + "font-size: 12px;"
    	mailcontent = mailcontent + "line-height: 130%;"
    	mailcontent = mailcontent + "color: #333333;"
    	mailcontent = mailcontent + "}"
    	mailcontent = mailcontent + "</style>"
    	mailcontent = mailcontent + "</head>"
    	mailcontent = mailcontent + "<body bgcolor=#FFFFFF leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
    	mailcontent = mailcontent + "<table width=573 border=0 cellpadding=0 cellspacing=0>"
    	mailcontent = mailcontent + "<tr>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_01.gif' width='45' height='114'></td>"
    	mailcontent = mailcontent + "<td> <a href='http://www.10x10.co.kr' target='_blank'><img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_02.gif' width='479' height='114' border='0'></a></td>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_03.gif' width='49' height='114'></td>"
    	mailcontent = mailcontent + "</tr>"
    	mailcontent = mailcontent + "<tr>"
    	mailcontent = mailcontent + "<td background='http://www.10x10.co.kr/apps/mail_form/images/mail_form_04.gif'></td>"
    	mailcontent = mailcontent + "<td bgcolor='#F7F7F7' align='center'>"
    	mailcontent = mailcontent + "<table border='0' cellpadding='0' cellspacing='0' height='200' width='90%'>"
    	mailcontent = mailcontent + "<tr>"
    	mailcontent = mailcontent + "<td class='text'>"
    	mailcontent = mailcontent + innerContents
    	mailcontent = mailcontent + "</td>"
    	mailcontent = mailcontent + "</tr>"
    	mailcontent = mailcontent + "</table>"
    	mailcontent = mailcontent + "</td>"
    	mailcontent = mailcontent + "<td background='http://www.10x10.co.kr/apps/mail_form/images/mail_form_06.gif'></td>"
    	mailcontent = mailcontent + "</tr>"
    	mailcontent = mailcontent + "<tr>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_07.gif' width='45' height='107'></td>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_10.gif' width='479' height='107'></td>"
    	mailcontent = mailcontent + "<td>"
    	mailcontent = mailcontent + "<img src='http://www.10x10.co.kr/apps/mail_form/images/mail_form_09.gif' width=49 height='107'></td>"
    	mailcontent = mailcontent + "</tr>"
    	mailcontent = mailcontent + "</table>"
    	mailcontent = mailcontent + "</body>"
    	mailcontent = mailcontent + "</html>"
    	
        Call SendMail(mailfrom, mailto, mailtitle, mailcontent)
        
        response.write "<script>alert('(" & mailfrom & ")메일로 업체 담당자에게 이메일(" & mailto & ")을 발송 하였습니다.');</script>"
    
    end if
    set ocontract = Nothing
    
    sqlStr = " update db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " set contractState=" & NextState & ""
    if (NextState="7") then
        sqlStr = sqlStr & " ,finishdate=getdate()"
    elseif (NextState="0") then
        sqlStr = sqlStr & " ,confirmdate=NULL"
    end if
    sqlStr = sqlStr & " where contractID=" & contractID
    
    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
    
    
    response.write "<script>alert('상태 정보를 수정 하였습니다.');</script>"
    response.write "<script>opener.location.reload(); location.replace('" & refer & "');</script>"
    dbget.close()	:	response.End
    
else
    response.write "<script>alert('정의되지 않았습니다. - " & mode & "');</script>"
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->