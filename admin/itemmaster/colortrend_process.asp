<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �÷�Ʈ���� ����
' Hieditor : 2012.03.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/color/colortrend_cls.asp"-->

<%
dim mode ,menupos ,mainimage ,textimage ,ctcode ,colorcode ,isusing ,state ,mainimagelink ,startdate , mainimagelinknew
Dim listimage , Nmainimage , partmdid ,  partwdid , viewno , colortitle
dim adminid , sqlstr ,itemid ,idx , orderno ,i ,allusing
	mode = requestcheckvar(request("mode"),32)
	mainimage = requestcheckvar(request("mainimage"),256)
	textimage = requestcheckvar(request("textimage"),256)
	listimage = requestcheckvar(request("listimage"),256)
	Nmainimage = requestcheckvar(request("Nmainimage"),256)
	ctcode = requestcheckvar(request("ctcode"),10)
	colorcode = requestcheckvar(request("iCD"),10)
	isusing = requestcheckvar(request("isusing"),1)
	state = requestcheckvar(request("state"),1)
	startdate = requestcheckvar(request("startdate"),10)
	mainimagelink = request("mainimagelink")
	mainimagelinknew = request("mainimagelinknew") '2013��
	itemid = request("itemid")
	idx = request("idx")
	orderno = request("orderno")
	allusing = request("allusing")

	viewno = requestcheckvar(request("viewno"),10)
	partwdid = requestcheckvar(request("partwdid"),10)
	partmdid = requestcheckvar(request("partmdid"),10)
	colortitle = requestcheckvar(request("colortitle"),50)

adminid = session("ssBctId")
	
dim refer
	refer = request.ServerVariables("HTTP_REFERER")
	
'/���
if mode = "trendreg" then
	
	if checkNotValidHTML(mainimagelink)  Or checkNotValidHTML(mainimagelinknew)  then
	%>

	<script>
		alert('���� �̹��� �ʿ� ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');
		history.go(-1);
	</script>		

	<%
	dbget.close()	:	response.End
	end if
	
	'/�űԵ��
	if ctcode = "" then

		sqlstr = "insert into db_item.dbo.tbl_colortrend" + vbcrlf
		sqlstr = sqlstr & " (colorcode ,isusing ,state ,mainimage ,mainimagelink , mainimagelinknew ,textimage"
		sqlstr = sqlstr & " ,startdate ,lastupdate ,regdate ,lastadminid , listimg , Nmainimg , viewno , partmdid , partwdid , colortitle ) values (" + vbcrlf
		sqlstr = sqlstr & " "&colorcode&",'"&isusing&"' ,"&state&",'"&mainimage&"','"&html2db(mainimagelink)&"' ,'"&html2db(mainimagelinknew)&"' ,'"&textimage&"'" + vbcrlf
		sqlstr = sqlstr & " ,'" &html2db(startdate)&" 00:00:00',getdate(),getdate(),'"&adminid&"'" + vbcrlf
		sqlstr = sqlstr & " ,'"& listimage &"','"& Nmainimage &"',"& viewno &",'"& partmdid &"','"& partwdid &"','"& colortitle &"'" + vbcrlf
		sqlstr = sqlstr & " )"
		
		'response.write sqlstr &"<Br>"
		dbget.execute sqlstr

		response.write	"<script language='javascript'>"
		response.write	"	alert('OK');"
		response.write "	opener.location.reload();"
		response.write "	self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End
	
	'/����	
	else
		sqlstr = "update db_item.dbo.tbl_colortrend" + vbcrlf
		sqlstr = sqlstr & " set colorcode = "&colorcode&"" + vbcrlf
		sqlstr = sqlstr & " ,isusing = '"&isusing&"'" + vbcrlf
		sqlstr = sqlstr & " ,state = "&state&"" + vbcrlf
		sqlstr = sqlstr & " ,mainimage = '"&mainimage&"'" + vbcrlf				
		sqlstr = sqlstr & " ,mainimagelink = '"&html2db(mainimagelink)&"'" + vbcrlf		
		sqlstr = sqlstr & " ,mainimagelinknew = '"&html2db(mainimagelinknew)&"'" + vbcrlf		
		sqlstr = sqlstr & " ,textimage = '"&textimage&"'" + vbcrlf
		sqlstr = sqlstr & " ,startdate = '"&html2db(startdate)&" 00:00:00'" + vbcrlf
		sqlstr = sqlstr & " ,lastupdate = getdate()" + vbcrlf
		sqlstr = sqlstr & " ,lastadminid = '"&adminid&"'" + vbcrlf
		sqlstr = sqlstr & " ,listimg = '"&listimage&"'" + vbcrlf
		sqlstr = sqlstr & " ,Nmainimg = '"&Nmainimage&"'" + vbcrlf
		sqlstr = sqlstr & " ,viewno = "&viewno&"" + vbcrlf
		sqlstr = sqlstr & " ,partmdid = '"&partmdid&"'" + vbcrlf
		sqlstr = sqlstr & " ,partwdid = '"&partwdid&"'" + vbcrlf
		sqlstr = sqlstr & " ,colortitle = '"&colortitle&"'" + vbcrlf
		sqlstr = sqlstr & " where ctcode ="&ctcode&""

		response.write sqlstr &"<Br>"
		dbget.execute sqlstr

		response.write	"<script language='javascript'>"
		response.write	"	alert('OK');"
		response.write "	opener.location.reload();"
		response.write "	self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End		
	end if

'//�÷�Ʈ���� ��ǰ�߰�
elseif mode="itemadd" then

	if colorcode = "" then
		response.write	"<script language='javascript'>"
		response.write	"	alert('�÷��ڵ尡 �����ϴ�');"
		response.write	"	location.href='"&refer&"';"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

	if Right(itemid,1)="," then
		itemid = Left(itemid,Len(itemid)-1)
	end if

	sqlstr = "insert into db_item.dbo.tbl_colortrend_item (" + vbcrlf
	sqlstr = sqlstr & " colorCode ,itemid ,orderno ,isusing ,regdate)" + vbcrlf
	sqlstr = sqlstr & " select " + vbcrlf
	sqlstr = sqlstr & " "&colorcode&" ,i.itemid , 100 , 'Y' ,getdate()" + vbcrlf
	sqlstr = sqlstr & " from [db_item].[dbo].tbl_item i" + vbcrlf
	sqlstr = sqlstr & " left join db_item.dbo.tbl_colortrend_item ti" + vbcrlf
	sqlstr = sqlstr & " 	on i.itemid = ti.itemid" + vbcrlf
	sqlstr = sqlstr & " 	and ti.colorCode = "&colorcode&"" + vbcrlf
	sqlstr = sqlstr & " 	and ti.isusing = 'Y'" + vbcrlf
	sqlstr = sqlstr & " where i.itemid in (" + itemid + ")" + vbcrlf
	sqlstr = sqlstr & " and ti.idx is null"		'/�̵̹�ϵǾ� �ִ� ��ǰ ����
	
	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr

	response.write	"<script language='javascript'>"
	response.write	"	alert('OK');"
	response.write	"	location.href='"&refer&"';"
	response.write	"</script>"
	dbget.close()	:	response.End

'//�÷�Ʈ���� ��ǰ����
elseif mode="delitem" then

	if idx = "" then
		response.write	"<script language='javascript'>"
		response.write	"	alert('�ε��� �ڵ尡 �����ϴ�');"
		response.write	"	location.href='"&refer&"';"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

	if Right(idx,1)="," then
		idx = Left(idx,Len(idx)-1)
	end if

	sqlstr = "update db_item.dbo.tbl_colortrend_item" + vbcrlf
	sqlstr = sqlstr & " set isusing = 'N'" + vbcrlf
	sqlstr = sqlstr & " where idx in (" + idx + ")"

	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr

	response.write	"<script language='javascript'>"
	response.write	"	alert('OK');"
	response.write	"	location.href='"&refer&"';"
	response.write	"</script>"
	dbget.close()	:	response.End

'//���ļ��� ����
elseif mode="ChangeSort" then

	if idx = "" then
		response.write	"<script language='javascript'>"
		response.write	"	alert('�ε��� �ڵ尡 �����ϴ�');"
		response.write	"	location.href='"&refer&"';"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

	idx = split(idx,",")
	orderno = split(orderno,",")
	
	for i=0 to ubound(idx)-1
		sqlStr = "update db_item.dbo.tbl_colortrend_item" + vbcrlf
		sqlStr = sqlStr & " set orderno='" & orderno(i) & "'" + vbcrlf
		sqlStr = sqlStr & " where idx='" & idx(i) & "';" & vbCrLf
					
		'response.write sqlstr &"<Br>"
		dbget.execute sqlstr
	next

	response.write	"<script language='javascript'>"
	response.write	"	alert('OK');"
	response.write	"	location.href='"&refer&"';"
	response.write	"</script>"
	dbget.close()	:	response.End

'//��뿩�� ����
elseif mode="chisusing" then

	if idx = "" then
		response.write	"<script language='javascript'>"
		response.write	"	alert('�ε��� �ڵ尡 �����ϴ�');"
		response.write	"	location.href='"&refer&"';"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

	if Right(idx,1)="," then
		idx = Left(idx,Len(idx)-1)
	end if
	
	sqlStr = " update ti" + vbcrlf
	sqlStr = sqlStr & " set ti.isusing='" & allusing & "'" + vbcrlf
	sqlstr = sqlstr & " from db_item.dbo.tbl_colortrend_item ti" + vbcrlf
	sqlstr = sqlstr & " left join db_item.dbo.tbl_colortrend_item tii" + vbcrlf
	sqlstr = sqlstr & " 	on ti.itemid = tii.itemid" + vbcrlf
	sqlstr = sqlstr & " 	and ti.colorcode = tii.colorcode" + vbcrlf	
	sqlstr = sqlstr & " 	and tii.idx not in (" + idx + ")" + vbcrlf
	sqlstr = sqlstr & " 	and tii.isusing = 'Y'" + vbcrlf
	sqlstr = sqlstr & " where ti.idx in (" & idx & ")" + vbcrlf
	sqlstr = sqlstr & " and tii.idx is null"		'/��뿩�ΰ� N �λ�ǰ�� Y���� �����... ������ Y�� ��ǰ�� �ִ��� üũ�� ���� �ϸ� ������ ���

	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr
		
	response.write	"<script language='javascript'>"
	response.write	"	alert('OK');"
	response.write	"	location.href='"&refer&"';"
	response.write	"</script>"
	dbget.close()	:	response.End
end if	
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->