<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 브랜드페이지 관리 
' History : 2009.03.26 한용민 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/brandstreet/brandstreet_upche_cls.asp"-->

<%
dim  itemid , mode
	mode = requestCheckVar(request("mode"),30)
	itemid = requestCheckVar(request("itemid"),300)
	itemid = left(itemid,len(itemid)-1)
dim sql 

if mode = "" or itemid = "" then
	response.write "<script>"
	response.write "alert('오류가 발생했습니다. 시스템팀에 문의하세요.');"
	response.write "self.close()"
	response.write "</script>"
	dbget.close()	:	response.End
	
end if

'//중단배너처리
if mode = "isusing_no" then
	
	sql = "update db_brand.dbo.tbl_upche_brandstreet set" + vbcrlf
	sql = sql & " isusing='N'" + vbcrlf				
	sql = sql & " where idx in ("&itemid&")" + vbcrlf
	
	'response.write sql&"<Br>"
	dbget.execute sql	
end if

%>

<script language="javascript">
	opener.location.reload();
	self.close();
</script>


<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

