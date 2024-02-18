<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 컬러트랜드 관리
' Hieditor : 2012.03.29 한용민 생성
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
Dim listimage , Nmainimage , partmdid ,  partwdid , viewno , colortitle, playidx
dim adminid , sqlstr ,itemid ,idx , orderno ,i ,allusing
	mode = requestcheckvar(request("mode"),32)
	mainimage = requestcheckvar(request("mainimage"),256)
	textimage = requestcheckvar(request("textimage"),256)
	listimage = requestcheckvar(request("listimage"),256)
	Nmainimage = requestcheckvar(request("Nmainimage"),256)
	ctcode = requestcheckvar(request("ctcode"),10)
	playidx = requestcheckvar(request("playidx"),10)
	isusing = requestcheckvar(request("isusing"),1)
	state = requestcheckvar(request("state"),1)
	startdate = requestcheckvar(request("startdate"),10)
	mainimagelink = request("mainimagelink")
	mainimagelinknew = request("mainimagelinknew") '2013용
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


'/등록
if mode="itemadd" then

	if playidx = "" then
		response.write	"<script language='javascript'>"
		response.write	"	alert('PLAY idx가 없습니다');"
		response.write	"	location.href='"&refer&"';"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

	if Right(itemid,1)="," then
		itemid = Left(itemid,Len(itemid)-1)
	end if

	sqlstr = "insert into db_sitemaster.dbo.tbl_play_mo_item (" + vbcrlf
	sqlstr = sqlstr & " playidx, itemid ,orderno ,isusing ,regdate)" + vbcrlf
	sqlstr = sqlstr & " select " + vbcrlf
	sqlstr = sqlstr & " '" & playidx & "', i.itemid , 100 , 'Y' ,getdate()" + vbcrlf
	sqlstr = sqlstr & " from [db_item].[dbo].tbl_item i" + vbcrlf
	sqlstr = sqlstr & " left join db_sitemaster.dbo.tbl_play_mo_item ti" + vbcrlf
	sqlstr = sqlstr & " 	on i.itemid = ti.itemid" + vbcrlf
	sqlstr = sqlstr & " 	and ti.playidx = '" & playidx & "'" + vbcrlf
	sqlstr = sqlstr & " 	and ti.isusing = 'Y'" + vbcrlf
	sqlstr = sqlstr & " where i.itemid in (" + itemid + ")" + vbcrlf
	sqlstr = sqlstr & " and ti.idx is null"		'/이미등록되어 있는 상품 제낌
	
	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr

	response.write	"<script language='javascript'>"
	response.write	"	alert('OK');"
	response.write	"	location.href='"&refer&"';"
	response.write	"</script>"
	dbget.close()	:	response.End

'//컬러트렌드 상품삭제
elseif mode="delitem" then

	if idx = "" then
		response.write	"<script language='javascript'>"
		response.write	"	alert('인덱스 코드가 없습니다');"
		response.write	"	location.href='"&refer&"';"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

	if Right(idx,1)="," then
		idx = Left(idx,Len(idx)-1)
	end if

	sqlstr = "update db_sitemaster.dbo.tbl_play_mo_item" + vbcrlf
	sqlstr = sqlstr & " set isusing = 'N'" + vbcrlf
	sqlstr = sqlstr & " where idx in (" + idx + ")"

	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr

	response.write	"<script language='javascript'>"
	response.write	"	alert('OK');"
	response.write	"	location.href='"&refer&"';"
	response.write	"</script>"
	dbget.close()	:	response.End

'//정렬순서 저장
elseif mode="ChangeSort" then

	if idx = "" then
		response.write	"<script language='javascript'>"
		response.write	"	alert('인덱스 코드가 없습니다');"
		response.write	"	location.href='"&refer&"';"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

	idx = split(idx,",")
	orderno = split(orderno,",")
	
	for i=0 to ubound(idx)-1
		sqlStr = "update db_sitemaster.dbo.tbl_play_mo_item" + vbcrlf
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

'//사용여부 수정
elseif mode="chisusing" then

	if idx = "" then
		response.write	"<script language='javascript'>"
		response.write	"	alert('인덱스 코드가 없습니다');"
		response.write	"	location.href='"&refer&"';"
		response.write	"</script>"
		dbget.close()	:	response.End
	end if

	if Right(idx,1)="," then
		idx = Left(idx,Len(idx)-1)
	end if
	
	sqlStr = " update ti" + vbcrlf
	sqlStr = sqlStr & " set ti.isusing='" & allusing & "'" + vbcrlf
	sqlstr = sqlstr & " from db_sitemaster.dbo.tbl_play_mo_item ti" + vbcrlf
	sqlstr = sqlstr & " left join db_sitemaster.dbo.tbl_play_mo_item tii" + vbcrlf
	sqlstr = sqlstr & " 	on ti.itemid = tii.itemid" + vbcrlf
	sqlstr = sqlstr & " 	and ti.playidx = tii.playidx" + vbcrlf
	sqlstr = sqlstr & " 	and tii.idx not in (" + idx + ")" + vbcrlf
	sqlstr = sqlstr & " 	and tii.isusing = 'Y'" + vbcrlf
	sqlstr = sqlstr & " where ti.idx in (" & idx & ")" + vbcrlf
	sqlstr = sqlstr & " and tii.idx is null"		'/사용여부가 N 인상품을 Y으로 변경시... 기존에 Y인 상품이 있는지 체크후 존재 하면 제끼고 등록

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