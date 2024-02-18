<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 우수회원샵
' Hieditor : 2009.12.28 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/specialshop/specialshop_cls.asp"-->

<%
dim id,openDate,status,i , isusing , sql , itemidarr, title, endDate, mode , idx
	id = requestCheckVar(getNumeric(request("id")),10)
	openDate = requestCheckVar(request("openDate"),32)
	status = requestCheckVar(request("status"),1)
	mode = requestCheckVar(request("mode"),32)
	isusing = requestCheckVar(request("isusing"),1)
	mode = requestCheckVar(request("mode"),32)
	itemidarr = request("itemidarr")
	idx = requestCheckVar(getNumeric(request("idx")),10)
	title = requestCheckVar(request("title"),200)
	endDate = requestCheckVar(request("endDate"),32)
	
'//상세등록
if mode = "reg" then
	'//신규등록
	if id = "" then
		if title <> "" and not(isnull(title)) then
			title = ReplaceBracket(title)
		end If

		sql = " select * from db_item.dbo.tbl_specialShop with (nolock) where" + vbcrlf 
		sql = sql & " isusing='Y' and '"&openDate&"' between openDate and endDate"
		
		'response.write sql &"<br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
		if not(rsget.bof or rsget.eof) then
			response.write "<script type='text/javascript'>alert('입력하신 오픈날짜에 진행중인 테마가 있습니다.\n확인후 등록 하세요'); history.go(-1);</script>"
			dbget.close() : response.end	
		end if
		rsget.close

		sql = "insert into db_item.dbo.tbl_specialShop (title,openDate,endDate,status,isusing) " + vbcrlf
		sql = sql & " values (" + vbcrlf
		sql = sql & " '"&title&"'" + vbcrlf
		sql = sql & " ,'"& openDate &"'" + vbcrlf
		sql = sql & " ,'"& endDate &"'" + vbcrlf
		sql = sql & " ,"&status&"" + vbcrlf
		sql = sql & " ,'"&isusing&"'" + vbcrlf
		sql = sql & " )" + vbcrlf
		
		'response.write sql &"<br>"
		dbget.execute sql
				
	'//수정
	else
		if title <> "" and not(isnull(title)) then
			title = ReplaceBracket(title)
		end If

		sql = "update db_item.dbo.tbl_specialShop set"
		sql = sql & " title = '"& title&"'" + vbcrlf
		sql = sql & " ,openDate = '"& openDate&"'" + vbcrlf
		sql = sql & " ,endDate = '"& endDate&"'" + vbcrlf
		sql = sql & " ,status = "&status&""+ vbcrlf
		sql = sql & " ,isusing = '"&isusing&"'"+ vbcrlf
		sql = sql & " where id = "&id&""

		'response.write sql &"<br>"
		dbget.execute sql
	
	end if

	response.write "<script type='text/javascript'>alert('저장되었습니다.'); opener.location.reload(); self.close();</script>"

'//상품등록
elseif mode= "itemadd" then
	itemidarr = Trim(Replace(itemidarr," ",""))
	itemidarr = Trim(Replace(itemidarr,vbCrLf,","))
	
	if right(itemidarr,1) = "," then
		itemidarr = left(itemidarr,len(itemidarr)-1)
	end if
	
	if id = "" then
		response.write "<script type='text/javascript'>alert('id값이 없습니다'); self.close();</script>"
	end if
	
	sql = "select count(*) from db_item.dbo.tbl_specialShopitem with (nolock) where"
	sql = sql & " id = '"&id&"' and itemid in ("&itemidarr&")"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
	if rsget(0) > 0 then
		response.write "<script type='text/javascript'>alert('이미 등록된 상품이 있습니다.\n확인하고 다시 입력하세요.'); history.back();</script>"
		rsget.close()
		dbget.close()
		response.end
	else
		rsget.close()
	end if
	
	sql = "insert into db_item.dbo.tbl_specialShopitem (id,itemid,isusing)" + vbcrlf
	sql = sql & " (select "&id&" , itemid,'Y' " + vbcrlf
	sql = sql & " from db_item.dbo.tbl_item " + vbcrlf	
	sql = sql & " where itemid in ("&itemidarr&"))" + vbcrlf

	'response.write sql &"<br>"
	'dbget.close()
	'response.end
	dbget.execute sql

	response.write "<script type='text/javascript'>alert('저장되었습니다.'); opener.location.reload(); location.href='specialshop_edititem.asp?id='+"&id&";</script>"			

'//상품삭제
elseif mode= "dellitem" then
	
	if id = "" or idx="" then
		response.write "<script type='text/javascript'>alert(id값이나 idx값이 없습니다'); self.close();</script>"
	end if	
	
	sql = "update db_item.dbo.tbl_specialShopitem set " + vbcrlf
	sql = sql & " isusing='N'" + vbcrlf
	sql = sql & " where idx = "&idx&""

	'response.write sql &"<br>"
	dbget.execute sql
	
	response.write "<script type='text/javascript'>alert('삭제되었습니다.'); opener.location.reload(); location.href='specialshop_edititem.asp?id='+"&id&";</script>"			

'//이벤트 상태 변경
elseif mode= "statuschange" then
		
	sql = "exec db_item.dbo.sp_Ten_specialShop" + vbcrlf

	'response.write sql &"<br>"
	dbget.execute sql
	
	response.write "<script type='text/javascript'>alert('이벤트 상태가 적용되었습니다.'); opener.location.reload(); self.close();</script>"	
	
'//상품 실서버 적용
elseif mode= "itemupdate" then
		
	sql = "exec db_item.dbo.sp_Ten_specialShop_itemupdate" + vbcrlf

	'response.write sql &"<br>"
	dbget.execute sql
	
	response.write "<script type='text/javascript'>alert('상품이 적용되었습니다.'); opener.location.reload(); self.close();</script>"		
end if			
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

