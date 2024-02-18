<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- include virtual="/admin/offshop/shopcscenter/cscenter_mail_Function_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->
<%
dim mode, modeflag2, divcd, reguserid, ipkumdiv ,title, orderno, contents_jupsu
dim finishuser, contents_finish ,requireupche, requiremakerid, ForceReturnByTen ,detailitemlist
dim opentitle, opencontents ,newasid ,isCsMailSend ,IsAllCancel ,CancelValidResultMessage
dim sqlStr, i ,ScanErr ,ResultMsg, ReturnUrl, EtcStr ,ProceedFinish ,returnmethod ,masteridxtmp
dim oordermaster ,buf_requiremakerid , masteridx , csmasteridx , cancelorderno ,GC_IsOLDOrder
dim reqname ,reqphone ,reqhp , reqzipcode ,reqzipaddr ,reqaddress ,comment ,reqemail
	masteridx        = requestCheckVar(request.Form("masteridx"),10)
	csmasteridx   = requestCheckVar(request.Form("csmasteridx"),10)
	mode        = requestCheckVar(request.Form("mode"),32)
	modeflag2   = requestCheckVar(request.Form("modeflag2"),10)
	divcd       = requestCheckVar(request.Form("divcd")	,4)
	ipkumdiv    = requestCheckVar(request.Form("ipkumdiv"),10)
	reguserid   = session("ssbctid")
	finishuser  = reguserid
	title       = requestCheckVar(html2DB(request.Form("title")),128)
	orderno = requestCheckVar(request.Form("orderno"),16)
	contents_jupsu  = requestCheckVar(html2DB(request.Form("contents_jupsu")),800)
	detailitemlist  = html2db(request.Form("detailitemlist"))
	contents_finish = requestCheckVar(html2DB(request.Form("contents_finish")),800)
	requireupche = requestCheckVar(request.Form("requireupche"),1)
	requiremakerid = requestCheckVar(request.Form("requiremakerid"),32)
	ForceReturnByTen = requestCheckVar(request.Form("ForceReturnByTen"),32)
	buf_requiremakerid  = requestCheckVar(request.Form("buf_requiremakerid"),32)
	isCsMailSend = requestCheckVar((request.Form("csmailsend")="on"),10)
	cancelorderno       = requestCheckVar(request.Form("cancelorderno"),16)
	reqname       = requestCheckVar(request.Form("reqname"),32)
	reqphone       = requestCheckVar(request.Form("reqphone"),32)
	reqhp       = requestCheckVar(request.Form("reqhp"),32)
	reqzipcode       = requestCheckVar(request.Form("reqzipcode"),7)
	reqzipaddr       = requestCheckVar(request.Form("reqzipaddr"),128)
	reqaddress       = requestCheckVar(request.Form("reqaddress"),255)
	comment       = request.Form("comment")
	reqemail       = requestCheckVar(request.Form("reqemail"),128)
	
	newasid = -1		
	if (returnmethod="") then returnmethod="R000"
	ScanErr = ""
	ProceedFinish = False

''주문 마스타
set oordermaster = new COrder
	oordermaster.FRectmasteridx = masteridx	
	oordermaster.fQuickSearchOrderMaster

'response.write "mode : " & mode & "<br>"

'/cs접수
if (mode="regcsas") then
	'response.write "divcd : " & divcd & "<br>"

    'CS 접수 - 업체a/s
	if (divcd="A030") then

        dbget.beginTrans
 
		'a/s 의 경우 업체a/s 와..   업체a/s(매장회수)가 있슴..  둘다 브랜드id 지정 해야함
        if (requiremakerid<>"") then

			'' CS Master 접수 ''html2db 사용하지 말것.
	        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)
	   
			'CS Detail 접수(관련 상품정보)
			Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx ,"Y")
		        	
        	'/업체 a/s 접수 업체배송변경 requireupche : Y
            call RegCSMasterAddUpche_off(csmasteridx, requiremakerid)

			'배송지 등록(고객주소나 매장주소)
			call Regdelivery_off(csmasteridx, reqname ,reqphone ,reqhp ,reqemail,reqzipcode ,reqzipaddr ,reqaddress ,comment)
			
        	'업체a/s(매장회수)
            '' CS Master 접수 ''html2db 사용하지 말것.
            newasid = RegCSMaster_off("A031", orderno, reguserid, "업체A/S(매장회수)", contents_jupsu, masteridx)
			
			'CS Detail 접수(관련 상품정보)			
            Call AddCSDetailByArrStr_off(detailitemlist, newasid, orderno ,masteridx ,"N")
			
			'/맞교환 회수 접수 매장배송 변경 requiremaejang : Y
        	call RegCSMasterAddmaejang_off(newasid, requiremakerid)

			'배송지 등록(업체주소)
			if isarray(Getpartnerdeliverinfo_off(requiremakerid,"")) then
				call Regdelivery_off(newasid, Getpartnerdeliverinfo_off(requiremakerid,"")(2,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(3,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(4,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(5,0),Getpartnerdeliverinfo_off(requiremakerid,"")(6,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(7,0) ,Getpartnerdeliverinfo_off(requiremakerid,"")(8,0) ,comment)			
			end if

			ResultMsg = "업체A/S 접수 및 업체A/S(매장회수) 접수 완료"
        end if

        ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script type='text/javascript'>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"            
            dbget.close()	:	response.End
        End If

    else
        ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
        response.write "<script type='text/javascript'>alert('" + ResultMsg + "');</script>"
        response.write "<script type='text/javascript'>history.back();</script>"
        dbget.close()	:	response.End
    end if

''접수 내역 수정
elseif (mode="editcsas") then

    dbget.beginTrans
	
	'' CS master 수정
	Call EditCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu ,csmasteridx)

    '' CS Detail 수정
    Call EditCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno)

	'배송지 등록(고객주소나 매장주소)
	call Regdelivery_off(csmasteridx, reqname ,reqphone ,reqhp ,reqemail,reqzipcode ,reqzipaddr ,reqaddress ,comment)
			
    ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd + "&mode=editreginfo"
		
    If (Err.Number = 0) and (ScanErr="") Then
        dbget.CommitTrans
    
    	ResultMsg = ResultMsg + "OK"
    Else
        dbget.RollBackTrans
        response.write "<script type='text/javascript'>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"        
        dbget.close()	:	response.End
    End If

'CS 접수 내역 완료처리
elseif (mode="finishcsas") then	    
	'response.write "divcd : " & divcd & "<br>"
			
	'CS 접수 내역 완료처리 - 맞교환 출고 / 누락 / 서비스 발송 / 기타 /  출고시 유의사항	/	업체a/s /	업체a/s(매장회수)
    if  (divcd="A000") or (divcd="A001") or (divcd="A002") or (divcd="A009") or (divcd="A006") or (divcd="A005") or (divcd="A700") or (divcd="A030") or (divcd="A031") then
	
		dbget.beginTrans
		
		Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)
		
		ResultMsg = "처리 완료"
		ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
		
		If (Err.Number = 0) and (ScanErr="") Then
		    dbget.CommitTrans
		Else
		    dbget.RollBackTrans
		    response.write "<script type='text/javascript'>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"	        
		    dbget.close()	:	response.End
		End If

	else
        ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
        response.write "<script type='text/javascript'>alert('" + ResultMsg + "');</script>"
        response.write "<script type='text/javascript'>history.back();</script>"
        dbget.close()	:	response.End
    end if

'' 업체 처리완료 => 접수상태로변경
elseif (mode="upcheconfirm2jupsu") then
	    
    sqlStr = " select top 1 currstate from db_shop.dbo.tbl_shopjumun_cs_master"
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)
	
	'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if not(rsget("currstate")="B006" or rsget("currstate")="B008") then
	            ResultMsg = "업체처리완료나 매장처리완료 상태가 아닙니다. 수정 불가"                
	        end if
		else
		    ResultMsg = "코드없음. 수정 불가"
		end if
	rsget.Close

    if (ResultMsg="") then
        sqlStr = " update db_shop.dbo.tbl_shopjumun_cs_master" + VbCrlf
        sqlStr = sqlStr + "set currstate='B001'" + VbCrlf
        sqlStr = sqlStr + ",contents_jupsu='" + (contents_jupsu) + "'" + VbCrlf
        sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)

		'response.write sqlStr &"<br>"        
        dbget.Execute sqlStr

        ResultMsg = "처리 완료"
        ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
    else
        response.write "<script type='text/javascript'>alert('" + ResultMsg + "');</script>"
        response.write "<script type='text/javascript'>history.back();</script>"
        dbget.close()	:	response.End
    end if
    
'CS 삭제
elseif (mode="deletecsas") then
	
    ''Check Valid Delete - 현재는 B006 업체처리완료 , B007 완료 내역은 취소(삭제) 불가
    if (NOT ValidDeleteCS_off(csmasteridx)) then
        response.write "<script type='text/javascript'>alert(" + Chr(34) + "현재 취소 가능 상태가 아닙니다. 관리자 문의 요망." + Chr(34) + ")</script>"
        response.write "<script type='text/javascript'>history.back()</script>"
        dbget.close()	:	response.End
    end if

    If Not DeleteCSProcess_off(csmasteridx, reguserid) then
        ResultMsg = ResultMsg + "데이터 삭제시 오류"
    else
        ResultMsg = ResultMsg + "OK"
    End if
    
    ReturnUrl = "/admin/offshop/shopcscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd + "&mode=editreginfo"
   
end if

%>

<%
response.write "<script type='text/javascript'>alert('" + ResultMsg + "');</script>"
response.write "<script type='text/javascript'>location.replace('" + ReturnUrl + "');</script>"
response.End
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->