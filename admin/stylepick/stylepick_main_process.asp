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
dim mainidx,cd1,mainimage,state,startdate,enddate,isusing,lastadminid,opendate
dim closedate,partMDid,partWDid ,mode , sqlstr , i , menupos ,mainimagelink
dim strAdd ,mainctidx ,gubun ,gubunvalue ,copy ,link ,contentsyn, comment
	startdate = left(request("startdate"),10)
	enddate = left(request("enddate"),10)
	mode = request("mode")
	contentsyn = request("contentsyn")
	mainidx = request("mainidx")	
	cd1 = request("cd1")
	menupos = request("menupos")
	mainimage = request("mainimage")
	state = request("state")
	isusing = request("isusing")
	lastadminid = session("ssBctId")
	partMDid = request("partMDid")
	partWDid = request("partWDid")
	mainimagelink = request("mainimagelink")
	mainctidx = request("mainctidx")
	gubun = request("gubun")
	gubunvalue = request("gubunvalue")
	isusing = request("isusing")
	copy = request("copy")
	link = request("link")
	comment = request("comment")	

	if contentsyn = "on" then
		contentsyn = "Y"
	else
		contentsyn = "N"
	end if	

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	
'/���
if mode = "mainedit" then

	if checkNotValidHTML(comment) then
	%>

	<script>
	alert('���뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');
	history.go(-1);
	</script>		

	<%
	dbget.close()	:	response.End
	end if

	'/�űԵ��
	if mainidx = "" then

		'���°� �����϶� ������ ���
		opendate = "null"
		closedate = "null"
		
		IF state = 7 THEN
			opendate = "getdate()"
		ELSEIF state = 9 THEN
			closedate = "getdate()"
		END IF

		sqlstr = "insert into db_giftplus.dbo.tbl_stylepick_main" + vbcrlf
		sqlstr = sqlstr & " (cd1,mainimage,state,startdate,enddate,isusing,lastadminid,opendate"
		sqlstr = sqlstr & " ,closedate,partMDid,partWDid,mainimagelink,contentsyn,comment) values (" + vbcrlf
		sqlstr = sqlstr & " '"&cd1&"','"&html2db(mainimage)&"' ,"&state&",'"&html2db(startdate)&" 00:00:00'" + vbcrlf
		sqlstr = sqlstr & " ,'"&html2db(enddate)&" 23:59:59','"&isusing&"','"&lastadminid&"',"&opendate&"" + vbcrlf
		sqlstr = sqlstr & " ,"&closedate&",'"&partMDid&"','"&partWDid&"','"&html2db(mainimagelink)&"','N','"&html2db(comment)&"'" + vbcrlf
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

		dbget.beginTrans
		
		sqlstr = "update db_giftplus.dbo.tbl_stylepick_main set" + vbcrlf
		sqlstr = sqlstr & " cd1 = '"&cd1&"'" + vbcrlf
		sqlstr = sqlstr & " ,contentsyn = '"&contentsyn&"'" + vbcrlf
		sqlstr = sqlstr & " ,mainimage = '"&html2db(mainimage)&"'" + vbcrlf
		sqlstr = sqlstr & " ,mainimagelink = '"&html2db(mainimagelink)&"'" + vbcrlf	
		sqlstr = sqlstr & " ,state = '"&state&"'" + vbcrlf
		sqlstr = sqlstr & " ,startdate = '"&html2db(startdate)&" 00:00:00'" + vbcrlf
		sqlstr = sqlstr & " ,enddate = '"&html2db(enddate)&" 23:59:59'" + vbcrlf
		sqlstr = sqlstr & " ,isusing = '"&isusing&"'" + vbcrlf
		sqlstr = sqlstr & " ,lastadminid = '"&lastadminid&"'" + vbcrlf
		sqlstr = sqlstr & " ,partMDid = '"&partMDid&"'" + vbcrlf
		sqlstr = sqlstr & " ,partWDid = '"&partWDid&"' " & strAdd
		sqlstr = sqlstr & " ,comment = '"&html2db(comment)&"'" + vbcrlf
		sqlstr = sqlstr & " where mainidx ="&mainidx&""

		'response.write sqlstr &"<Br>"
		dbget.execute sqlstr

		if gubun <> "" then
			gubun = split(gubun,",")
			gubunvalue = split(gubunvalue,",")
			copy = split(copy,",")
			link = split(link,",")
			
			'/�������� ������
			sqlstr = "update db_giftplus .dbo.tbl_stylepick_main_contents set" + vbcrlf
			sqlstr = sqlstr & " isusing='N'" + vbcrlf
			sqlstr = sqlstr & " where mainidx="&mainidx&"" + vbcrlf

			'response.write sqlstr &"<Br>"
			dbget.execute sqlstr
			
			for i = 0 to ubound(gubun)
				'/���
				sqlstr = "insert into db_giftplus.dbo.tbl_stylepick_main_contents" + vbcrlf
				sqlstr = sqlstr & " (mainidx ,gubun ,gubunvalue ,isusing ,copy ,link ,lastadminid"
				sqlstr = sqlstr & " ) values (" + vbcrlf
				sqlstr = sqlstr & " "&mainidx&","&gubun(i)&" ,"&gubunvalue(i)&",'Y'" + vbcrlf
				sqlstr = sqlstr & " ,'"&html2db(copy(i))&"','"&html2db(link(i))&"','"&lastadminid&"'" + vbcrlf
				sqlstr = sqlstr & " )"

				'response.write sqlstr &"<Br>"
				dbget.execute sqlstr
			next
			
		end if

		If Err.Number = 0 Then
		    dbget.CommitTrans

			response.write	"<script language='javascript'>"
			response.write	"	alert('OK');"
			response.write "	opener.location.reload();"
			response.write "	location.replace('" + referer + "');"
			response.write	"</script>"
			dbget.close()	:	response.End
		Else
		    dbget.RollBackTrans

			response.write	"<script language='javascript'>"
			response.write	"	alert('�������� ó���� �ƴմϴ�');"
			response.write "	self.close();"
			response.write	"</script>"
			dbget.close()	:	response.End		    
		End If		
	end if
end if	
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	