<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Event 등록  
' History : 2008.04.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim evt_type,evt_name, evt_partner, evt_comment,isusing,startdate,enddate,listimgName,regdate,evt_code,eventdate,ticket_isusing, write_work
dim main1imgName,main2imgName,barnerimgName,barner2imgName,barner3imgName,image_main_link , comment , mode
dim main3imgName, main4imgName, main5imgName
dim m_isusing, m_img_icon, m_img_main1, m_img_main2, m_main_content, m_cmt_desc, m_sortNo, m_evtbn_code
Dim edid , emid, evt_kind

	evt_type = request("evt_type")
	evt_kind = request("evt_kind")
	evt_partner = html2db(request("evt_partner"))
	evt_name = html2db(request("evt_name"))
	evt_comment = html2db(request("evt_comment"))
	isusing = request("isusing")
	startdate =	request("startdate")
	enddate = request("enddate")
	listimgName = request("listimgName")
	main1imgName = html2db(request("main1imgName"))
	main2imgName = html2db(request("main2imgName"))
	main3imgName = html2db(request("main3imgName"))
	main4imgName = html2db(request("main4imgName"))
	main5imgName = html2db(request("main5imgName"))	
	barnerimgName = request("barnerimgName")
	barner2imgName = request("barner2imgName")
	barner3imgName = request("barner3imgName")
	regdate = request("regdate") 					
	image_main_link = html2db(request("image_main_link"))
	evt_code = request("evt_code")
	eventdate = request("eventdate")
	ticket_isusing = request("ticket_isusing")
	comment = request("comment")
	mode = request("mode")
	write_work = request("write_work")

	edid		=	request("selDId")
	emid		=	request("selMId")

	m_isusing		= request("m_isusing")
	m_img_icon		= html2db(request("m_img_icon"))
	m_img_main1		= html2db(request("m_img_main1"))
	m_img_main2		= html2db(request("m_img_main2"))
	m_main_content	= html2db(replace(request("m_main_content"),"'",""""))

	m_cmt_desc		= html2db(replace(request("m_cmt_desc"),"'",""""))
	m_sortNo		= html2db(request("m_sortNo"))
	m_evtbn_code	= request("m_evtbn_code")
	
	if m_evtbn_code = "" then m_evtbn_code = 0
	If write_work = "" Then write_work = "N"
	if comment = "" then comment="OFF"
	if m_isusing = "" then m_isusing="N"
	if m_sortNo = "" then m_sortNo="10"

dim sql

'//신규저장
if mode = "add" then

	sql = "insert into db_culture_station.dbo.tbl_culturestation_event" + vbcrlf
	sql = sql & " (comment,evt_type, evt_name, evt_partner, evt_comment, startdate ,enddate ,eventdate,isusing ,ticket_isusing"
	sql = sql & "	, image_list,image_main,image_main2,image_barner,image_barner2,image_barner3,image_main_link,image_main3,image_main4,image_main5,write_work"
	sql = sql & "	, m_isusing, m_img_icon, m_img_main1, m_img_main2, m_main_content, m_cmt_desc, m_sortNo, web_sortNo,designerid , partMDid, m_evtbn_code, evt_kind)" + vbcrlf
	sql = sql & " values ("  + vbcrlf
	sql = sql & " '"&comment&"'" + vbcrlf
	sql = sql & " ,"&evt_type&"" + vbcrlf
	sql = sql & " ,'"& html2db(request("evt_name")) &"'" + vbcrlf
	sql = sql & " ,'"& html2db(request("evt_partner")) &"'" + vbcrlf
	sql = sql & " ,'"& html2db(request("evt_comment")) &"'" + vbcrlf
	sql = sql & " ,'"&startdate&"'" + vbcrlf
	sql = sql & " ,'"&enddate&" 23:59:59'" + vbcrlf
	sql = sql & " ,'"&eventdate&"'" + vbcrlf
	sql = sql & " ,'"&isusing&"'" + vbcrlf
	sql = sql & " ,'"&ticket_isusing&"'" + vbcrlf
	sql = sql & " ,'"&listimgName&"'" + vbcrlf
	sql = sql & " ,'"&main1imgName&"'" + vbcrlf
	sql = sql & " ,'"&main2imgName&"'" + vbcrlf
	sql = sql & " ,'"&barnerimgName&"'"	 + vbcrlf
	sql = sql & " ,'"&barner2imgName&"'"	 + vbcrlf
	sql = sql & " ,'"&barner3imgName&"'"	 + vbcrlf
	sql = sql & " ,'"&image_main_link&"'" + vbcrlf
	sql = sql & " ,'"&main3imgName&"'" + vbcrlf
	sql = sql & " ,'"&main4imgName&"'"	 + vbcrlf
	sql = sql & " ,'"&main5imgName&"'"	 + vbcrlf
	sql = sql & " ,'"&write_work&"'"	 + vbcrlf
	sql = sql & " ,'"&m_isusing&"'"	 + vbcrlf
	sql = sql & " ,'"&m_img_icon&"'"	 + vbcrlf
	sql = sql & " ,'"&m_img_main1&"'"	 + vbcrlf
	sql = sql & " ,'"&m_img_main2&"'"	 + vbcrlf
	sql = sql & " ,'"&m_main_content&"'"	 + vbcrlf
	sql = sql & " ,'"&m_cmt_desc&"'"	 + vbcrlf
	sql = sql & " ,'"&m_sortNo&"',10"	 + vbcrlf
	sql = sql & " ,'"&edid&"'"	 + vbcrlf
	sql = sql & " ,'"&emid&"'"	 + vbcrlf
	sql = sql & " ,"&m_evtbn_code&""	 + vbcrlf
	sql = sql & " ,"&evt_kind&""	 + vbcrlf
	sql = sql & ")"
	'response.write sql
	dbget.execute sql

'//수정	
elseif mode = "edit" then

	sql = "update db_culture_station.dbo.tbl_culturestation_event set" + vbcrlf
	sql = sql & " comment='"&comment&"'" + vbcrlf
	sql = sql & " ,evt_type="&evt_type&"" + vbcrlf
	sql = sql & " ,evt_name='"& html2db(request("evt_name")) &"'" + vbcrlf
	sql = sql & " ,evt_partner='"& html2db(request("evt_partner")) &"'" + vbcrlf
	sql = sql & " ,evt_comment='"& html2db(request("evt_comment")) &"'" + vbcrlf
	sql = sql & " ,startdate='"&startdate&"'" + vbcrlf
	sql = sql & " ,enddate='"&enddate&" 23:59:59'" + vbcrlf
	sql = sql & " ,eventdate='"&eventdate&"'" + vbcrlf	
	sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
	sql = sql & " ,ticket_isusing='"&ticket_isusing&"'" + vbcrlf
	sql = sql & " ,image_list='"&listimgName&"'" + vbcrlf
	sql = sql & " ,image_main='"&main1imgName&"'" + vbcrlf
	sql = sql & " ,image_main2='"&main2imgName&"'" + vbcrlf
	sql = sql & " ,image_barner='"&barnerimgName&"'" + vbcrlf	
	sql = sql & " ,image_barner2='"&barner2imgName&"'" + vbcrlf	
	sql = sql & " ,image_barner3='"&barner3imgName&"'" + vbcrlf	
	sql = sql & " ,image_main_link='"&image_main_link&"'" + vbcrlf
	sql = sql & " ,image_main3='"&main3imgName&"'" + vbcrlf
	sql = sql & " ,image_main4='"&main4imgName&"'" + vbcrlf
	sql = sql & " ,image_main5='"&main5imgName&"'" + vbcrlf
	sql = sql & " ,write_work='"&write_work&"'" + vbcrlf
	sql = sql & " ,m_isusing='"&m_isusing&"'" + vbcrlf
	sql = sql & " ,m_img_icon='"&m_img_icon&"'" + vbcrlf
	sql = sql & " ,m_img_main1='"&m_img_main1&"'" + vbcrlf
	sql = sql & " ,m_img_main2='"&m_img_main2&"'" + vbcrlf
	sql = sql & " ,m_main_content='"&m_main_content&"'" + vbcrlf
	sql = sql & " ,m_cmt_desc='"&m_cmt_desc&"'" + vbcrlf
	sql = sql & " ,m_sortNo='"&m_sortNo&"'" + vbcrlf
	sql = sql & " ,designerid='"&edid&"'" + vbcrlf
	sql = sql & " ,partMDid='"&emid&"'" + vbcrlf
	sql = sql & " ,m_evtbn_code="&m_evtbn_code&"" + vbcrlf
	sql = sql & " ,evt_kind="&evt_kind&"" + vbcrlf
	sql = sql & " where evt_code = "&evt_code&"" + vbcrlf
	
'	response.write sql
	dbget.execute sql
elseif mode = "del" Then
	sql = "update db_culture_station.dbo.tbl_culturestation_event set" + vbcrlf
	sql = sql & " m_main_content=''" + vbcrlf
	sql = sql & " ,m_cmt_desc=''" + vbcrlf
	sql = sql & " where evt_code = "&evt_code&"" + vbcrlf
	dbget.execute sql
end if	
%>
<% If mode = "del" Then %>
<% Response.write "ok"%>
<% Else %>
<script>
	opener.location.reload();
	<% If Request("isimgdel") = "o" Then %>
		location.href = "/admin/culturestation/event_edit.asp?evt_code=<%=evt_code%>";
	<% Else %>
	self.close();
	<% End If %>
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->