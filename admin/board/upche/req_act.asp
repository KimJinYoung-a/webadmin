<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<%
'###########################################################
' Description : 업체 입점문의
' History : 서동석 생성
'			2008.09.01 한용민 수정
'###########################################################

dim companyrequest
dim boarditem
dim id, mode,cd1 , cd2 , cd3
dim user,comment,commmode
dim ipjumYN
dim page
dim gubun,sellgubun,onlymifinish,research,searchkey,catevalue, dispcate
Dim strSQL, menupos
	page=requestCheckVar(getNumeric(request("pg")),10)
	id = requestCheckVar(getNumeric(request("id")),10)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
	mode = request("mode")
	cd1 = request("cd1")
	cd2 = left(request("cd2"),3)
	cd3 = right(request("cd3"),3)
	user=html2db(request("user"))
	comment=html2db(request("comment"))
	sellgubun=request("sellgubun")
	ipjumYN=request("ipjumYN")
	gubun = request("gubun")
	onlymifinish = request("onlymifinish")
	research = request("research")
	searchkey = request("searchkey")
	catevalue=request("catevalue")
	commmode=request("commmode")
	dispcate	=  requestCheckVar(Request("disp"),16) 
dim opt
	opt="pg="+ page + "&id=" + id + "&mode=" + mode + "&cd1=" + cd1 + "&user=" + user + "&sellgubun=" + sellgubun + "&ipjumYN=" + ipjumYN + "&gubun=" + gubun + "&onlymifinish=" + onlymifinish + "&research=" + research +"&searchkey=" + searchkey + "&catevalue="+ catevalue + "&commmode=" + commmode + "&menupos=" + menupos 

set companyrequest = New CCompanyRequest

if (mode = "finish") then

        companyrequest.finish(id)
		response.write "<script>opener.location.reload();self.close();</script>"
		
elseif (mode="del") then
		
		companyrequest.delitem id
		response.write "<script>opener.location.reload();self.close();</script>"

elseif (mode="chcate") then

		companyrequest.catechange id,dispcate
		
		response.write "<script>location.replace('/admin/board/upche/req_view.asp?" + opt + "');</script>"
		
elseif (mode="chsell") then
		
		companyrequest.sellchange id,sellgubun
						
		response.write "<script>location.replace('/admin/board/upche/req_view.asp?" + opt + "');</script>"
		
elseif (mode="ipjum") then
		
		companyrequest.ipjumchange id,ipjumYN
						
		response.write "<script>location.replace('/admin/board/upche/req_view.asp?" + opt + "');</script>"
		
elseif (mode="comm") then
		
		companyrequest.writecomm id,user,comment
						
		response.write "<script>location.replace('/admin/board/upche/req_view.asp?" + opt + "');</script>"

elseif (mode="delworkid") Then
	strSQL = ""
	strSQL = strSQL & " UPDATE [db_cs].[dbo].tbl_company_request SET "
	strSQL = strSQL & " workid = null "
	strSQL = strSQL & " WHERE id = '"&id&"' "
	dbget.execute strSQL
	response.write "<script>location.replace('/admin/board/upche/req_view2.asp?id=" + id + "');</script>"

elseif (mode="reqdel") Then
	strSQL = "UPDATE [db_cs].[dbo].tbl_company_request" & vbcrlf
	strSQL = strSQL & " SET isusing='N' WHERE" & vbcrlf
	strSQL = strSQL & " id = '"&id&"'" & vbcrlf

	'response.write strSQL & "<br>"
	dbget.execute strSQL

	response.write "<script type='text/javascript'>"
	response.write "	alert('삭제 되었습니다.');"
	response.write "	self.close();"
	response.write "	opener.location.reload();"
	response.write "</script>"
end if

set companyrequest=nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->