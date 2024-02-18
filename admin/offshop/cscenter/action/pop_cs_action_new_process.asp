<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_mail_Function_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim mode, modeflag2, divcd, reguserid, ipkumdiv ,title, orderno, contents_jupsu
dim finishuser, contents_finish ,requireupche, requiremakerid, ForceReturnByTen ,detailitemlist
dim opentitle, opencontents ,newasid ,isCsMailSend ,IsAllCancel ,CancelValidResultMessage
dim sqlStr, i ,ScanErr ,ResultMsg, ReturnUrl, EtcStr ,ProceedFinish ,returnmethod ,masteridxtmp
dim oordermaster ,buf_requiremakerid , masteridx , csmasteridx , cancelorderno ,GC_IsOLDOrder
	masteridx        = requestCheckVar(request.Form("masteridx"),10)
	csmasteridx   = requestCheckVar(request.Form("csmasteridx"),10)
	mode        = requestCheckVar(request.Form("mode"),32)
	modeflag2   = requestCheckVar(request.Form("modeflag2"),32)
	divcd       = requestCheckVar(request.Form("divcd"),4)
	ipkumdiv    = requestCheckVar(request.Form("ipkumdiv"),1)
	reguserid   = session("ssbctid")
	finishuser  = reguserid
	title       = requestCheckVar(html2DB(request.Form("title")),128)
	orderno = requestCheckVar(request.Form("orderno"),16)
	contents_jupsu  = requestCheckVar(html2DB(request.Form("contents_jupsu")),800)
	detailitemlist  = html2db(request.Form("detailitemlist"))
	contents_finish = requestCheckVar(html2DB(request.Form("contents_finish")),32)
	requireupche = requestCheckVar(request.Form("requireupche"),1)
	requiremakerid = requestCheckVar(request.Form("requiremakerid"),32)
	ForceReturnByTen = requestCheckVar(request.Form("ForceReturnByTen"),32)
	buf_requiremakerid  = requestCheckVar(request.Form("buf_requiremakerid"),32)
	isCsMailSend = requestCheckVar((request.Form("csmailsend")="on"),32)
	cancelorderno       = requestCheckVar(request.Form("cancelorderno"),16)

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

	'CS 접수 - 주문취소
	if (divcd="A008") then

		dbget.beginTrans
		
        '' CS Master 접수 ''html2db 사용하지 말것.
        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)  

        'CS Detail 접수(관련 상품정보)
        Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx)

		'/주문내역과 마이너스주문내역이 일치 하는지 체크
		'CancelValidResultMessage = GetPartialCancelRegValidResult_off(detailitemlist, csmasteridx, orderno,masteridx ,cancelorderno)

		if (CancelValidResultMessage <> "") then
			ScanErr = CancelValidResultMessage
		end if	        		

		'/마스터 취소처리
        'Call masterCancelProcess_off(masteridx ,cancelorderno)
            
    	''바로 완료처리로 진행 할지 여부 - AsDetail 입력후 검사
        ProceedFinish = IsDirectProceedFinish_off(divcd, csmasteridx, masteridx, EtcStr)
        contents_finish = ""

        '' 완료처리 프로세스
        If (ProceedFinish) then
			'/디테일 취소처리
            Call CancelProcess_off(detailitemlist, csmasteridx, orderno,masteridx ,cancelorderno)
			Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)

			sqlStr = ""
			sqlStr = "select top 1 masteridx , detailidx , orderno" + vbcrlf
			sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_detail" + vbcrlf
			sqlStr = sqlStr & " where cancelyn='N'"
			sqlStr = sqlStr & " and masteridx = "&masteridx&"" + vbcrlf
		
			'response.write sqlStr &"<br>"
			rsget.open sqlStr ,dbget ,1
		
			if not(rsget.eof) then
				masteridxtmp = false
			else
				masteridxtmp = true
			end if
		
			rsget.close()
		
			'//디테일이 전부 취소 라면 마스터도 취소 시킨다
			if masteridxtmp then
				sqlStr = ""
				sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_master set" + vbcrlf
				sqlStr = sqlStr & " cancelyn='Y'" + vbcrlf
				sqlStr = sqlStr & " where masteridx = "&masteridx&""
		
				'response.write sqlStr &"<br>"
				dbget.execute sqlStr
			end if
	
			'//디테일 테이블 상태가 일부출고 보다 작은 상품이 존재 하지 않으면 마스터 테이블 상태를 출고완료로 바꾼다
			'전부출고
		    sqlStr = " update db_shop.dbo.tbl_shopbeasong_order_master set										" & VbCRLF
		    sqlStr = sqlStr + " ipkumdiv='8', beadaldate=getdate() 														" & VbCRLF
			sqlStr = sqlStr + " where masteridx in ( 																" & VbCRLF
		    sqlStr = sqlStr + " 	select 																	" & VbCRLF
		    sqlStr = sqlStr + " 	m.masteridx 														" & VbCRLF
		    sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shopbeasong_order_master m 							" & VbCRLF
		    sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopbeasong_order_detail d 					" & VbCRLF
		    sqlStr = sqlStr + " 		on m.masteridx=d.masteridx 										" & VbCRLF
		    sqlStr = sqlStr + " 	where d.itemid<>0 													" & VbCRLF
		    sqlStr = sqlStr + " 	and m.masteridx in ("&masteridx&") 																	" & VbCRLF
		    sqlStr = sqlStr + " 	group by m.masteridx 														" & VbCRLF
		    sqlStr = sqlStr + " 	having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0 " & VbCRLF
		    sqlStr = sqlStr + " ) 																			" & VbCRLF
		
		    'response.write sqlStr &"<br>"
		    dbget.Execute sqlStr
        ELSE
            ResultMsg = ResultMsg + "->. 상품 준비중 상태인 상품이 존재하므로\n\n 주문 취소 접수만 진행 되었습니다.\n\n 업체 전화 확인후 완료 처리하셔야 합니다."
        End If
	
        ResultMsg = ResultMsg & "OK"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
        
        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
            'dbget.RollBackTrans
            
			response.write "<script>"
			response.write "	alert('"&ResultMsg&"');"
			response.write "	location.replace('"&ReturnUrl&"');"
			response.write "</script>"
			dbget.close()	:	response.End	            
        Else
            dbget.RollBackTrans
            
            response.write "<script>"
            response.write "	alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")"
            response.write "</script>"
            dbget.close()	:	response.End
        End If
  		
	'CS 접수 - 기타사항 / 출고시유의사항 / 업체 추가 정산비
    elseif (divcd="A009") or (divcd="A006") or (divcd="A700") then
             
        dbget.beginTrans

        '' CS Master 접수 ''html2db 사용하지 말것.
        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)    

        'CS Detail 접수(관련 상품정보)
        Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx)

        '업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
        if (requiremakerid<>"") then
            call RegCSMasterAddUpche_off(csmasteridx, requiremakerid)
        end if

        if (isCsMailSend) then
            Call SendCsActionMail_off(csmasteridx)
        End If
        
        ResultMsg = ResultMsg + "\nOK"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
            
			response.write "<script>location.replace('"&ReturnUrl&"');</script>"
			dbget.close()	:	response.End	            
        Else
            dbget.RollBackTrans
            
            response.write "<script>"
            response.write "	alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")"
            response.write "</script>"
            dbget.close()	:	response.End
        End If

	'CS 접수 - 누락재발송, 서비스발송
    elseif (divcd="A001") or (divcd="A002") then
  	              
        dbget.beginTrans

        '' CS Master 접수 ''html2db 사용하지 말것.
        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)
        
        'CS Detail 접수(관련 상품정보)
        Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx)

    
		'업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
        if (requiremakerid<>"") then
            call RegCSMasterAddUpche_off(csmasteridx, requiremakerid)
        else
        
        	'/매장배송인 경우 '/차후 텐바이텐 배송이 생길경우 구분해서 넣어야함
        	'/매장배송 requiremaejang : Y  : 텐바이텐배송 requiremaejang : N
        	call RegCSMasterAddmaejang_off(csmasteridx)
        end if

        ResultMsg = "접수완료"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans

			response.write "<script>alert('OK'); location.replace('"&ReturnUrl&"');</script>"
			dbget.close()	:	response.End
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"            
            dbget.close()	:	response.End
        End If

        if (isCsMailSend) then
            Call SendCsActionMail_off(csmasteridx)
        End If
        
    'CS 접수 - 맞교환출고
    elseif (divcd="A000") then

        dbget.beginTrans

		'' CS Master 접수 ''html2db 사용하지 말것.
        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)
   
		'CS Detail 접수(관련 상품정보)
		Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx)
 
		'업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
        if (requiremakerid<>"") then            
            call RegCSMasterAddUpche_off(csmasteridx, requiremakerid)

            ResultMsg = "맞교환 접수완료 - 업체배송"        
        
        else
        	'/매장배송인 경우 '/차후 텐바이텐 배송이 생길경우 구분해서 넣어야함
      	        	
        	'/맞교환 출고 접수 매장배송 변경 requiremaejang : Y
        	call RegCSMasterAddmaejang_off(csmasteridx)
        	
        	'매장 배송의 경우 맞교환 회수 접수
            '' CS Master 접수 ''html2db 사용하지 말것.
            newasid = RegCSMaster_off("A013", orderno, reguserid, "맞교환 회수접수", contents_jupsu, masteridx)
			
			'CS Detail 접수(관련 상품정보)			
            Call AddCSDetailByArrStr_off(detailitemlist, newasid, orderno ,masteridx)
			
			'/맞교환 회수 접수 매장배송 변경 requiremaejang : Y
        	call RegCSMasterAddmaejang_off(newasid)

             ResultMsg = "맞교환 출고 접수 및 회수접수 완료 - 매장 배송"
        end if

        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"            
            dbget.close()	:	response.End
        End If

        ''이메일 발송 맞교환 접수
        if (isCsMailSend) then
            Call SendCsActionMail_off(csmasteridx)

            ''맞교환 회수가 있을경우
            if (newasid>0) then
                Call SendCsActionMail_off(newasid)
            end if
        End If    
    else
        ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if

''접수 내역 수정
elseif (mode="editcsas") then

    dbget.beginTrans

	Call EditCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu ,csmasteridx)

    '' CS Detail 수정
    Call EditCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno)

    ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd + "&mode=editreginfo"
		
    If (Err.Number = 0) and (ScanErr="") Then
        dbget.CommitTrans
    
    	ResultMsg = ResultMsg + "OK"
    Else
        dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"        
        dbget.close()	:	response.End
    End If

'CS 접수 내역 완료처리
elseif (mode="finishcsas") then	    
	'response.write "divcd : " & divcd & "<br>"
	
	'CS 접수 내역 완료처리 - 주문취소
    if (divcd="A008") then
		
		dbget.beginTrans
		
		'/디테일 취소처리
	    Call CancelProcess_off(detailitemlist, csmasteridx, orderno,masteridx ,cancelorderno)
		Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)	

		sqlStr = ""
		sqlStr = "select top 1 masteridx , detailidx , orderno" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_detail" + vbcrlf
		sqlStr = sqlStr & " where cancelyn='N'"
		sqlStr = sqlStr & " and masteridx = "&masteridx&"" + vbcrlf
	
		'response.write sqlStr &"<br>"
		rsget.open sqlStr ,dbget ,1
	
		if not(rsget.eof) then
			masteridxtmp = false
		else
			masteridxtmp = true
		end if
	
		rsget.close()
	
		'//디테일이 전부 취소 라면 마스터도 취소 시킨다
		if masteridxtmp then
			sqlStr = ""
			sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_master set" + vbcrlf
			sqlStr = sqlStr & " cancelyn='Y'" + vbcrlf
			sqlStr = sqlStr & " where masteridx = "&masteridx&""
	
			'response.write sqlStr &"<br>"
			dbget.execute sqlStr
		end if

		'//디테일 테이블 상태가 일부출고 보다 작은 상품이 존재 하지 않으면 마스터 테이블 상태를 출고완료로 바꾼다
		'전부출고
	    sqlStr = " update db_shop.dbo.tbl_shopbeasong_order_master set										" & VbCRLF
	    sqlStr = sqlStr + " ipkumdiv='8', beadaldate=getdate() 														" & VbCRLF
		sqlStr = sqlStr + " where masteridx in ( 																" & VbCRLF
	    sqlStr = sqlStr + " 	select 																	" & VbCRLF
	    sqlStr = sqlStr + " 	m.masteridx 														" & VbCRLF
	    sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shopbeasong_order_master m 							" & VbCRLF
	    sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopbeasong_order_detail d 					" & VbCRLF
	    sqlStr = sqlStr + " 		on m.masteridx=d.masteridx 										" & VbCRLF
	    sqlStr = sqlStr + " 	where d.itemid<>0 													" & VbCRLF
	    sqlStr = sqlStr + " 	and m.masteridx in ("&masteridx&") 																	" & VbCRLF
	    sqlStr = sqlStr + " 	group by m.masteridx 														" & VbCRLF
	    sqlStr = sqlStr + " 	having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0 " & VbCRLF
	    sqlStr = sqlStr + " ) 																			" & VbCRLF
	
	    'response.write sqlStr &"<br>"
	    dbget.Execute sqlStr
	    
		ResultMsg = "처리 완료"
		ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
		
		If (Err.Number = 0) and (ScanErr="") Then
		    dbget.CommitTrans
		Else
		    dbget.RollBackTrans
		    response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"	        
		    dbget.close()	:	response.End
		End If
				
	'CS 접수 내역 완료처리 - 맞교환 출고 / 누락 / 서비스 발송 / 기타 /  출고시 유의사항
    elseif  (divcd="A000") or (divcd="A001") or (divcd="A002") or (divcd="A009") or (divcd="A006") or (divcd="A005") or (divcd="A700") then
	
		dbget.beginTrans
		
		Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)
		
		ResultMsg = "처리 완료"
		ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
		
		If (Err.Number = 0) and (ScanErr="") Then
		    dbget.CommitTrans
		Else
		    dbget.RollBackTrans
		    response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"	        
		    dbget.close()	:	response.End
		End If
		
		If (isCsMailSend) then
		    if ((divcd="A000") or (divcd="A001") or (divcd="A002")) then
		        
		        ''맞교환/누락/서비스 완료 메일
		        Call SendCsActionMail_off(csmasteridx)
		    end if
		End If
	
	'CS 접수 내역 완료처리 - 맞교환회수(매장배송)
    elseif (divcd="A013") then    	
    	        
        dbget.beginTrans
        
		Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)

        ResultMsg = "처리 완료"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            dbget.close()	:	response.End
        End If

        ''맞교환 완료 메일
        If (isCsMailSend) then
            Call SendCsActionMail_off(csmasteridx)
        End If
	        
	else
        ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if

'' 업체 처리완료 => 접수상태로변경
elseif (mode="upcheconfirm2jupsu") then
	    
    sqlStr = " select top 1 currstate from db_shop.dbo.tbl_shopbeasong_cs_master"
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)
	
	'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if (rsget("currstate")<>"B006") then
	            ResultMsg = "업체 처리 완료 상태가 아닙니다. 수정 불가"
	        end if
		else
		    ResultMsg = "코드없음. 수정 불가"
		end if
	rsget.Close

    if (ResultMsg="") then
        sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_master" + VbCrlf
        sqlStr = sqlStr + "set currstate='B001'" + VbCrlf
        sqlStr = sqlStr + ",contents_jupsu='" + (contents_jupsu) + "'" + VbCrlf
        sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)
        dbget.Execute sqlStr

        ResultMsg = "처리 완료"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if
    
'CS 삭제
elseif (mode="deletecsas") then
	
    ''Check Valid Delete - 현재는 B006 업체처리완료 , B007 완료 내역은 취소(삭제) 불가
    if (NOT ValidDeleteCS_off(csmasteridx)) then
        response.write "<script>alert(" + Chr(34) + "현재 취소 가능 상태가 아닙니다. 관리자 문의 요망." + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if

    If Not DeleteCSProcess_off(csmasteridx, reguserid) then
        ResultMsg = ResultMsg + "데이터 삭제시 오류"
    else
        ResultMsg = ResultMsg + "OK"
    End if
    
    ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd + "&mode=editreginfo"
   
end if

%>

<%
response.write "<script>alert('" + ResultMsg + "');</script>"
response.write "<script>location.replace('" + ReturnUrl + "');</script>"
response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->