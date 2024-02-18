<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 스타일픽 관리
' Hieditor : 2011.04.05 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
dim mode , sqlstr , i , menupos ,evtidx,cd1 , title, subcopy, state, startdate, enddate
dim partMDid ,partWDid , isusing , comment ,banner_img ,lastadminid ,opendate ,closedate
dim strAdd ,itemidarr ,totalcount ,tmpitemid , tmpitem ,evtitemidxarr
	mode = request("mode")
	menupos = request("menupos")
	cd1 = request("cd1")
	lastadminid = session("ssBctId")
	evtidx = request("evtidx")
	title = request("title")
	subcopy = request("subcopy")
	state = request("state")
	startdate = left(request("startdate"),10)
	enddate = left(request("enddate"),10)
	partMDid = request("partMDid")
	partWDid = request("partWDid")
	isusing = request("isusing")
	comment = request("comment")
	banner_img = request("banner_img")
	itemidarr = request("itemidarr")
	evtitemidxarr = request("evtitemidxarr")
	
dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	
'/이벤트등록
if mode = "eventedit" then

	'/신규등록
	if evtidx = "" then

		if checkNotValidHTML(comment) then
		%>
	
		<script>
		alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');
		history.go(-1);
		</script>		
	
		<%
		dbget.close()	:	response.End
		end if

		'상태가 오픈일때 오픈일 등록
		opendate = "null"
		closedate = "null"
		
		IF state = 7 THEN
			opendate = "getdate()"
		ELSEIF state = 9 THEN
			closedate = "getdate()"
		END IF

		sqlstr = "insert into db_giftplus.dbo.tbl_stylepick_event" + vbcrlf
		sqlstr = sqlstr & " (title,subcopy ,state ,banner_img ,startdate ,enddate ,isusing,comment" + vbcrlf
		sqlstr = sqlstr & " ,lastadminid ,cd1 ,opendate ,closedate ,partMDid ,partWDid) values (" + vbcrlf
		sqlstr = sqlstr & " '"&html2db(title)&"','"&html2db(subcopy)&"' ,"&state&",'"&html2db(banner_img)&"'" + vbcrlf
		sqlstr = sqlstr & " ,'"&html2db(startdate)&" 00:00:00','"&html2db(enddate)&" 23:59:59','"&isusing&"','"&html2db(comment)&"'" + vbcrlf
		sqlstr = sqlstr & " ,'"&lastadminid&"','"&cd1&"',"&opendate&","&closedate&",'"&partMDid&"','"&partWDid&"'" + vbcrlf
		sqlstr = sqlstr & " )"
		
		'response.write sqlstr &"<Br>"
		dbget.execute sqlstr

		response.write	"<script language='javascript'>"
		response.write	"	alert('OK');"
		response.write "	opener.location.reload();"
		response.write "	self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End
	
	'//수정
	else
		if checkNotValidHTML(comment) then
		%>
	
		<script>
		alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');
		history.go(-1);
		</script>		
	
		<%
		dbget.close()	:	response.End
		end if
	
		
		strAdd=""
	 	IF (state = 7 and opendate ="" ) THEN 	'오픈처리일 설정
			strAdd = ", opendate = getdate() "
		ELSEIF (state = 9 and closedate ="" ) THEN
			strAdd = ", closedate = getdate() "	'종료처리일 설정
		END IF
	
		'종료일 이전에 종료시 종료일 현재 날짜로 변경
		IF state = 9 and  datediff("d",enddate,date()) < 0 THEN
			enddate = date()
		END IF

		sqlstr = "update db_giftplus.dbo.tbl_stylepick_event set" + vbcrlf
		sqlstr = sqlstr & " title = '"&html2db(title)&"'" + vbcrlf
		sqlstr = sqlstr & " ,subcopy = '"&html2db(subcopy)&"'" + vbcrlf
		sqlstr = sqlstr & " ,state = '"&state&"'" + vbcrlf
		sqlstr = sqlstr & " ,banner_img = '"&html2db(banner_img)&"'" + vbcrlf
		sqlstr = sqlstr & " ,startdate = '"&html2db(startdate)&" 00:00:00'" + vbcrlf
		sqlstr = sqlstr & " ,enddate = '"&html2db(enddate)&" 23:59:59'" + vbcrlf
		sqlstr = sqlstr & " ,isusing = '"&isusing&"'" + vbcrlf
		sqlstr = sqlstr & " ,comment = '"&html2db(comment)&"'" + vbcrlf
		sqlstr = sqlstr & " ,lastadminid = '"&lastadminid&"'" + vbcrlf
		sqlstr = sqlstr & " ,cd1 = '"&cd1&"'" + vbcrlf
		sqlstr = sqlstr & " ,partMDid = '"&partMDid&"'" + vbcrlf
		sqlstr = sqlstr & " ,partWDid = '"&partWDid&"' " & strAdd
		sqlstr = sqlstr & " where evtidx ="&evtidx&""

		'response.write sqlstr &"<Br>"
		dbget.execute sqlstr

		response.write	"<script language='javascript'>"
		response.write	"	alert('OK');"
		response.write "	opener.location.reload();"
		response.write "	location.replace('" + referer + "');"
		response.write	"</script>"
		dbget.close()	:	response.End
		
	end if

'/이벤트 상품 등록
elseif mode = "evtitemadd" then

	if itemidarr = "" or evtidx = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('코드에 문제가 있습니다.관리자 문의 하세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End	
	end if
	
	'/다른카테고리에 있는 중복 상품 체크
    sqlStr = sqlStr & " select"
    sqlStr = sqlStr & " ei.evtitemidx ,ei.evtidx ,ei.itemid ,ei.regdate ,ei.isusing"
    sqlStr = sqlStr & " ,e.evtidx,e.title,e.subcopy,e.state,e.banner_img,e.startdate,e.enddate"
    sqlStr = sqlStr & " ,e.isusing,e.regdate,e.comment,e.lastadminid,e.cd1,e.opendate,c1.catename"
    sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_event_item ei"
    sqlStr = sqlStr & " join db_giftplus.dbo.tbl_stylepick_event e"
    sqlStr = sqlStr & " 	on ei.evtidx = e.evtidx"
    sqlStr = sqlStr & " 	and e.state <> 9 and getdate() <= e.enddate"
    sqlStr = sqlStr & " 	and e.isusing='Y'"
    sqlStr = sqlStr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
    sqlStr = sqlStr & " 	on e.cd1 = c1.cd1"
    sqlStr = sqlStr & " 	and c1.isusing='Y'"
    sqlStr = sqlStr & " where ei.isusing='Y'"
    sqlStr = sqlStr & " and ei.evtidx <> "&evtidx&""
	sqlstr = sqlstr & " and ei.itemid in ("&itemidarr&")"
				
	'response.write sqlstr &"<Br>"
	rsget.open sqlstr ,dbget,1
	
	totalcount = rsget.recordcount
	
	if not rsget.EOF then
		do until rsget.EOF
		
		i = i + 1
		
		if tmpitem = "" then tmpitem = "\n\n타카테고리 진행중인 이벤트 중복 등록상품.. 참고하세요\n※10건 까지 노출됩니다\n\n"
		
		'/10건까지 노출
		if i+1 <= 10 then
			tmpitem = tmpitem & "["& rsget("catename") &" / 기획전코드:" & rsget("evtidx") &"] 상품코드:" & rsget("itemid") & "\n"
		end if
		
		tmpitemid = tmpitemid & rsget("itemid")
		
		if totalcount <> i then tmpitemid = tmpitemid &","
					
		rsget.movenext
		loop
	end if
	
	rsget.Close

	sqlstr = "insert into db_giftplus.dbo.tbl_stylepick_event_item (evtidx ,itemid ,isusing)" + vbcrlf
	sqlstr = sqlstr & "	select" + vbcrlf	
	sqlstr = sqlstr & "	"&evtidx&" ,i.itemid , 'Y'" + vbcrlf	
	sqlstr = sqlstr & "	from db_item.dbo.tbl_item i" + vbcrlf
	sqlstr = sqlstr & "	left join [db_giftplus].dbo.tbl_stylepick_event_item ei" + vbcrlf
	sqlstr = sqlstr & "	on i.itemid = ei.itemid" + vbcrlf
	sqlstr = sqlstr & "		and ei.isusing='Y'" + vbcrlf
	sqlStr = sqlStr & " 	and ei.evtidx = "&evtidx&""	
	sqlstr = sqlstr & "	where i.isusing = 'Y'" + vbcrlf
	sqlstr = sqlstr & "	and i.itemid in ("&itemidarr&")" + vbcrlf
	sqlstr = sqlstr & "	and ei.itemid is null" + vbcrlf		'/같은 카테고리내 중복 상품 제낌
	
	if tmpitemid <> "" then
		'sqlstr = sqlstr & "	and ei.itemid not in ("&tmpitemid&")" + vbcrlf		'/다른 카테고리내 중복 상품 제낌
	end if

	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr
	
	response.write	"<script language='javascript'>"
	response.write	"	alert('저장되었습니다"&tmpitem&"');"		
	response.write "	parent.frm.itemidarr.value = '';"
	response.write "	parent.frm.itemcount.value = '0';"
	response.write "	parent.opener.location.href='/admin/stylepick/stylepick_event_item.asp?evtidx="&evtidx&"&menupos="&menupos&"';"
	response.write "	location.href='about:blank'"
	response.write	"</script>"
	dbget.close()	:	response.End
	
'/상품 삭제
elseif mode = "evtitemdel" then

	if evtitemidxarr = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('코드에 문제가 있습니다.관리자 문의 하세요');"
		response.write "	location.replace('" + referer + "');"
		response.write "</script>"	
		dbget.close()	:	response.End	
	end if
	
	sqlstr = "update db_giftplus.dbo.tbl_stylepick_event_item set" + vbcrlf
	sqlstr = sqlstr & " isusing='N'"
	sqlstr = sqlstr & " where evtitemidx in ("&evtitemidxarr&")"
	
	'response.write sqlstr &"<Br>"
    dbget.Execute sqlStr		
	
	response.write	"<script language='javascript'>"
	response.write	"	alert('삭제되었습니다');"
	response.write "	location.href='/admin/stylepick/stylepick_event_item.asp?evtidx="&evtidx&"&menupos="&menupos&"';"
	response.write	"</script>"
	dbget.close()	:	response.End		

end if	
%>	
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	