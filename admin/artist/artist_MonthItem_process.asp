<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2012.03.29 김진영 생성
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/artistGalleryCls.asp" -->
<% 
dim mode , itemid, comment, sortNo, isusing, idx
	mode = requestcheckvar(request("mode"),25)
	idx = request("idx")
	itemid = requestcheckvar(request("itemid"),10)
	comment = request("comment")
	sortNo = request("sortNo")
	isusing = request("isusing")
	menupos 	= requestCheckVar(request("menupos"),128)
dim referer , sql
referer = request.ServerVariables("HTTP_REFERER")	

''//메인배너 상품 등록
if mode = "add" then
	sql = "insert into db_contents.dbo.tbl_artist_shop_MonthItem (itemid ,comment, sortNo ,isusing, regdate)" + vbcrlf
	sql = sql & " values (" + vbcrlf
	sql = sql & ""&itemid&"," + vbcrlf
	sql = sql & "'"&html2db(comment)&"'," + vbcrlf
	sql = sql & ""&sortNo&"," + vbcrlf
	sql = sql & "'"&isusing&"'," + vbcrlf
	sql = sql & "getdate()" + vbcrlf
	sql = sql & " )" + vbcrlf
	
	dbget.execute sql
elseif mode = "edit" then
	sql = "update db_contents.dbo.tbl_artist_shop_MonthItem set " + VbCrlf
	sql = sql + " itemid = '"&itemid&"'," + VbCrlf
	sql = sql + " comment = '"&html2db(comment)&"'," + VbCrlf
	sql = sql + " sortNo = '"&sortNo&"', " + VbCrlf
	sql = sql + " isusing = '"&isusing&"'" + VbCrlf
	sql = sql + " where (idx = '"&idx&"') "
	dbget.execute sql
end if	
%>	
<script language="javascript">
alert('저장되었습니다');
location.href='/admin/artist/artist_MonthItemList.asp?menupos=<%=menupos%>&mm=2';
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->