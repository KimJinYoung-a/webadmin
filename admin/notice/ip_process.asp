<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 사내 ip 관리
' History : 2008.07.01 한용민 생성 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/10x10staffcls.asp" -->
<%
Dim company_ip ,  id, company_name ,part_sn_box , gubuncd ,ipidx , mode ,sql ,tmpcnt ,isusing
	menupos = request("menupos")
	company_ip = request("company_ip")
	id = request("id")	
	company_name = request("company_name")
	part_sn_box = request("part_sn_box")
	gubuncd = request("gubuncd")
	ipidx = request("ipidx")
	isusing = request("isusing")
	mode = request("mode")	

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

'//ip 신규등록 & 수정
if mode = "edit" then
	'//수정
	if ipidx <> "" then
		sql = "update db_partner.dbo.tbl_equipment_ip set"	+vbcrlf
		sql = sql & " id = '"& id &"'"	+vbcrlf
		sql = sql & " ,company_name = '"& company_name &"'"	+vbcrlf
		sql = sql & " ,part_sn = '"& part_sn_box &"'"	+vbcrlf
		sql = sql & " ,isusing = '"& isusing &"'"	+vbcrlf
		sql = sql & " where ipidx = "& ipidx &""	+vbcrlf		
		
		'response.write sql & "<Br>"
		dbget.execute sql	

		response.write "<script language='javascript'>"
		response.write "	alert('수정 되었습니다');"
		response.write "	location.href='/admin/notice/ip_list.asp?menupos="&menupos&"&mode="&gubuncd&"';"
		response.write "</script>"
		dbget.close()	:	response.end
	
	'/신규등록
	else	
		tmpcnt=0
		sql = "select count(*) as cnt from db_partner.dbo.tbl_equipment_ip"
		sql = sql & " where isusing='Y' and company_ip='"&company_ip&"' and gubuncd = '"&gubuncd&"'"
	
		'response.write sql &"<br>"
		rsget.open sql,dbget,1
		
		if not rsget.EOF  then
		  tmpcnt = rsget("cnt")			
		end if
	
		rsget.close
		
		if tmpcnt <> 0 then
			response.write "<script language='javascript'>"
			response.write "	alert('이미등록된 IP가 있습니다');"	
			response.write "	location.href='/admin/notice/ip_list.asp?menupos="&menupos&"&gubuncd="&gubuncd&"';"
			response.write "</script>"	
			dbget.close() : response.end
		end if
		
		sql = ""
		sql = "insert into db_partner.dbo.tbl_equipment_ip (gubuncd ,company_ip ,id ,company_name ,part_sn) values (" + vbcrlf
		sql = sql & " '"&gubuncd&"'" + vbcrlf
		sql = sql & " ,'"&company_ip&"'" + vbcrlf
		sql = sql & " ,'"&id&"'" + vbcrlf
		sql = sql & " ,'"&company_name&"'" + vbcrlf
		sql = sql & " ,'"&part_sn_box&"'" + vbcrlf
		sql = sql & " )" + vbcrlf
		
		'response.write sql &"<br>"
		dbget.execute sql
				
		response.write "<script language='javascript'>"
		response.write "	alert('신규등록 되었습니다.');"
		response.write "	location.href='/admin/notice/ip_list.asp?menupos="&menupos&"&gubuncd="&gubuncd&"';"
		response.write "</script>"
		dbget.close()	:	response.end				
	end if
end if
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
