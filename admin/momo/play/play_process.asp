<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성놀이
' Hieditor : 2010.12.22 허진원 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim playSn ,startdate ,enddate ,isusing, playLinkType, evt_code, linkURL, itemid
dim i , mode , plyItemSn, chkCnt
	playSn = request("playSn")
	startdate = request("startdate")
	enddate = request("enddate")
	isusing = request("isusing")
	playLinkType = request("playLinkType")
	evt_code = request("evt_code")
	linkURL = request("linkURL")
	mode = request("mode")
	plyItemSn = request("plyItemSn")
	itemid = request("itemid")

dim referer , sql
referer = request.ServerVariables("HTTP_REFERER")

'//신규 & 수정 
if mode = "add" then
			
	'//신규	
	if playSn = "" then

        '// 중복 등록 검사
        sql = "select count(playSn) " + vbcrlf	
		sql = sql & " from db_momo.dbo.tbl_momo_playInfo" + vbcrlf
		sql = sql & " where isusing = 'Y' and playStartDate = '"&startdate&"' and playEndDate = '"&enddate&"'"

        rsget.Open sql, dbget, 1
        	chkCnt = rsget(0)
        rsget.Close
		if chkCnt>0  then
			response.write "<script language='javascript'>alert('해당 날짜에 대한 내역이 이미 존재 합니다');self.close();</script>"
			dbget.close() : response.end
		end if

		'// 이벤트 유효성 검사
		if playLinkType="E" then
	        sql = "select count(evt_code) " + vbcrlf	
			sql = sql & " from db_event.dbo.tbl_event" + vbcrlf
			sql = sql & " where evt_code='"&evt_code&"' and evt_using = 'Y' and evt_startdate<='"&startdate&"' and evt_enddate>='"&enddate&"'"

	        rsget.Open sql, dbget, 1
	        	chkCnt = rsget(0)
	        rsget.Close
			if chkCnt<=0  then
				response.write "<script language='javascript'>alert('해당 기간에 진행하는 이벤트가 없습니다.');history.back();</script>"
				dbget.close() : response.end
			end if
		end if

		'// 놀이정보 저장 처리
		sql = "insert into db_momo.dbo.tbl_momo_playInfo (playLinkType, evt_code, linkURL, playStartDate, playEndDate, isusing)" + vbcrlf
		sql = sql & " values (" + vbcrlf
		sql = sql & " '"&playLinkType&"'"
		sql = sql & " ,'"&evt_code&"'"
		sql = sql & " ,'"&html2db(linkURL)&"'"
		sql = sql & " ,'"&html2db(startdate)&"'"
		sql = sql & " ,'"&html2db(enddate)&"'"
		sql = sql & " ,'"&isusing&"'"
		sql = sql & " )"
	
		dbget.execute sql

	'//수정	
	else 
	
		sql = "update db_momo.dbo.tbl_momo_playInfo set" + vbcrlf	
		sql = sql & " playLinkType='"&playLinkType&"'" + vbcrlf
		sql = sql & " ,evt_code='"&evt_code&"'" + vbcrlf
		sql = sql & " ,linkURL='"&html2db(linkURL)&"'" + vbcrlf
		sql = sql & " ,playStartDate='"&html2db(startdate)&"'" + vbcrlf
		sql = sql & " ,playEndDate='"&html2db(enddate)&"'" + vbcrlf		
		sql = sql & " ,isusing='"&isusing&"'" + vbcrlf						
		sql = sql & " where playSn = "&playSn&"" + vbcrlf	
		
		dbget.execute sql
		
	end if			

	response.write "<script language='javascript'>"
	response.write "	opener.location.reload();"
	response.write "	alert('OK');"
	response.write "	self.close();"
	response.write "</script>"

'//상품등록
elseif mode = "itemAdd" then

	if playSn = "" then
		response.write "<script>alert('공감놀이 아이디 값이 없습니다.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	'// 전송된 아이템 코드값 확인
	if Right(itemid,1)="," then
		itemid = Left(itemid,Len(itemid)-1)
	end if

	'// 추가
	sql = "insert into db_momo.dbo.tbl_momo_playItem" &_
			" (playSn, itemid)" &_
			" select '" + Cstr(playSn) + "', itemid" &_
			" from [db_item].[dbo].tbl_item" &_
			" where itemid in (" + itemid + ")" 
	dbget.execute sql
	
	response.write "<script language='javascript'>"
	response.write "	location.replace('" + referer + "');"		
	response.write "</script>"

'//선택 상품삭제
elseif mode = "itemDel" then

	if playSn = "" then
		response.write "<script>alert('공감놀이 아이디 값이 없습니다.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	'// 전송된 아이템 코드값 확인
	if Right(plyItemSn,1)="," then
		plyItemSn = Left(plyItemSn,Len(plyItemSn)-1)
	end if

	'// 삭제
	sql = "delete db_momo.dbo.tbl_momo_playItem " + vbcrlf
	sql = sql & " where  plyItemSn in (" & plyItemSn & ") "
	dbget.execute sql
	
	response.write "<script language='javascript'>"
	response.write "	location.replace('" & referer & "');"
	response.write "</script>"

'//이벤트상품등록
elseif mode = "evtItemAdd" then
			
	if playSn = "" then
		response.write "<script>alert('공감놀이 아이디 값이 없습니다.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	'// 이벤트상품 존재여부 확인
    sql = "select count(itemid) " + vbcrlf	
	sql = sql & " from db_event.dbo.tbl_eventitem" + vbcrlf
	sql = sql & " where evt_code='"&evt_code&"' "

    rsget.Open sql, dbget, 1
    	chkCnt = rsget(0)
    rsget.Close
	if chkCnt<=0  then
		response.write "<script language='javascript'>alert('해당 이벤트에 등록된 상품이 없습니다.\n이벤트에서 상품을 먼저 등록해주세요.');history.back();</script>"
		dbget.close() : response.end
	end if

	'// 기존 상품 리셋
	sql = "delete db_momo.dbo.tbl_momo_playItem " + vbcrlf
	sql = sql & " where playSn=" & playSn
	dbget.execute sql

	'// 이벤트 상품 추가
	sql = "insert into db_momo.dbo.tbl_momo_playItem" &_
			" (playSn, itemid)" &_
			" select '" + Cstr(playSn) + "', itemid" &_
			" from [db_event].[dbo].tbl_eventitem" &_
			" where evt_code='"&evt_code&"' "
	dbget.execute sql

	response.write "<script language='javascript'>"
	response.write "	location.replace('" + referer + "');"		
	response.write "</script>"
end if
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
