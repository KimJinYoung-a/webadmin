<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��Ÿ���� ����
' Hieditor : 2011.04.05 �ѿ�� ����
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
	
'/�̺�Ʈ���
if mode = "eventedit" then

	'/�űԵ��
	if evtidx = "" then

		if checkNotValidHTML(comment) then
		%>
	
		<script>
		alert('���뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');
		history.go(-1);
		</script>		
	
		<%
		dbget.close()	:	response.End
		end if

		'���°� �����϶� ������ ���
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
	
	'//����
	else
		if checkNotValidHTML(comment) then
		%>
	
		<script>
		alert('���뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');
		history.go(-1);
		</script>		
	
		<%
		dbget.close()	:	response.End
		end if
	
		
		strAdd=""
	 	IF (state = 7 and opendate ="" ) THEN 	'����ó���� ����
			strAdd = ", opendate = getdate() "
		ELSEIF (state = 9 and closedate ="" ) THEN
			strAdd = ", closedate = getdate() "	'����ó���� ����
		END IF
	
		'������ ������ ����� ������ ���� ��¥�� ����
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

'/�̺�Ʈ ��ǰ ���
elseif mode = "evtitemadd" then

	if itemidarr = "" or evtidx = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('�ڵ忡 ������ �ֽ��ϴ�.������ ���� �ϼ���');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End	
	end if
	
	'/�ٸ�ī�װ��� �ִ� �ߺ� ��ǰ üũ
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
		
		if tmpitem = "" then tmpitem = "\n\nŸī�װ� �������� �̺�Ʈ �ߺ� ��ϻ�ǰ.. �����ϼ���\n��10�� ���� ����˴ϴ�\n\n"
		
		'/10�Ǳ��� ����
		if i+1 <= 10 then
			tmpitem = tmpitem & "["& rsget("catename") &" / ��ȹ���ڵ�:" & rsget("evtidx") &"] ��ǰ�ڵ�:" & rsget("itemid") & "\n"
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
	sqlstr = sqlstr & "	and ei.itemid is null" + vbcrlf		'/���� ī�װ��� �ߺ� ��ǰ ����
	
	if tmpitemid <> "" then
		'sqlstr = sqlstr & "	and ei.itemid not in ("&tmpitemid&")" + vbcrlf		'/�ٸ� ī�װ��� �ߺ� ��ǰ ����
	end if

	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr
	
	response.write	"<script language='javascript'>"
	response.write	"	alert('����Ǿ����ϴ�"&tmpitem&"');"		
	response.write "	parent.frm.itemidarr.value = '';"
	response.write "	parent.frm.itemcount.value = '0';"
	response.write "	parent.opener.location.href='/admin/stylepick/stylepick_event_item.asp?evtidx="&evtidx&"&menupos="&menupos&"';"
	response.write "	location.href='about:blank'"
	response.write	"</script>"
	dbget.close()	:	response.End
	
'/��ǰ ����
elseif mode = "evtitemdel" then

	if evtitemidxarr = "" then
		response.write "<script language='javascript'>"
		response.write	"	alert('�ڵ忡 ������ �ֽ��ϴ�.������ ���� �ϼ���');"
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
	response.write	"	alert('�����Ǿ����ϴ�');"
	response.write "	location.href='/admin/stylepick/stylepick_event_item.asp?evtidx="&evtidx&"&menupos="&menupos&"';"
	response.write	"</script>"
	dbget.close()	:	response.End		

end if	
%>	
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	