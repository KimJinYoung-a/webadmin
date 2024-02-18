<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : 모바일 mdpick
' History : 2013.12.17 한용민
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mobile/mdpick_cls.asp" -->

<%
dim menupos, adminid, mode, acURL, itemidarr, sqlstr, i, idx, isusing, orderno, startdate, enddate
	itemidarr = Request("itemidarr")
	acURL = request("acURL")
	mode = request("mode")
	menupos = request("menupos")
	idx = request("idx")
	isusing = request("isusing")
	orderno = request("orderno")
	startdate = request("startdate")
	enddate = request("enddate")
	
adminid=session("ssBctId")

if mode = "I" then
	if itemidarr="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('상품이 없습니다.');"
		response.write "</script>"
		dbget.close()	:	response.end
	end if		

	itemidarr = split(itemidarr,",")

	IF isarray(itemidarr) THEN
		for i = 0 to ubound(itemidarr)
		
		sqlstr = "if not exists(" + vbcrlf
		sqlstr = sqlstr & " 	select top 1 *" + vbcrlf
		sqlstr = sqlstr & " 	from db_sitemaster.dbo.tbl_mobile_main_mdpick" + vbcrlf
		sqlstr = sqlstr & " 	where isusing='Y'" + vbcrlf
		sqlstr = sqlstr & " 	and itemid="&trim(itemidarr(i))&"" + vbcrlf
		sqlstr = sqlstr & " )" + vbcrlf
		sqlstr = sqlstr & " 	insert into db_sitemaster.dbo.tbl_mobile_main_mdpick(" + vbcrlf
		sqlstr = sqlstr & " 	itemid, isusing, orderno, regdate, lastdate, regadminid, lastadminid" + vbcrlf
		sqlstr = sqlstr & " 	)" + vbcrlf
		sqlstr = sqlstr & " 		select top 500" + vbcrlf
		sqlstr = sqlstr & " 		i.itemid, 'Y', 99, getdate(), getdate(), '"&adminid&"', '"&adminid&"'" + vbcrlf
		sqlstr = sqlstr & " 		from db_item.dbo.tbl_item i" + vbcrlf
		sqlstr = sqlstr & " 		where i.isusing='Y' and i.itemid="&trim(itemidarr(i))&""

		'response.write sqlstr &"<Br>"
		dbget.execute sqlstr
		next
	END IF

%>
	<script langauge="javascript">
		alert('저장되었습니다.\n\n중복된 상품이거나, 상품사용여부가 사용중이 아닌 상품은 제외되고 등록됩니다');
		location.href ="about:blank";
		//parent.close();
		//parent.history.go(0);
		//parent.location.reload();
	</script>
<%
	dbget.close()	:	response.End

'/저장
elseif mode = "mdpickedit" then
	if idx="" or isusing="" or orderno="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('필요한 조건이 없습니다.');"
		response.write "</script>"
		dbget.close()	:	response.end
	end if

	sqlstr = "update db_sitemaster.dbo.tbl_mobile_main_mdpick" + vbcrlf
	sqlstr = sqlstr & " set isusing='"&isusing&"'" + vbcrlf
	sqlstr = sqlstr & " ,startdate='"&startdate&"'" + vbcrlf
	sqlstr = sqlstr & " ,enddate='"&enddate&"'" + vbcrlf
	sqlstr = sqlstr & " ,orderno="&orderno&"" + vbcrlf
	sqlstr = sqlstr & " ,lastdate=getdate()" + vbcrlf	
	sqlstr = sqlstr & " ,lastadminid='"&adminid&"'" + vbcrlf
	sqlstr = sqlstr & " where idx = "&idx&""
	
	'response.write sqlstr
	dbget.execute sqlstr

	response.write "<script type='text/javascript'>"
	response.write "	alert('저장되었습니다.');"
	response.write "	opener.location.reload();"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.end
	
else
	response.write "<script type='text/javascript'>"
	response.write "	alert('구분자가 없습니다.');"
	response.write "</script>"
	dbget.close()	:	response.end
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
