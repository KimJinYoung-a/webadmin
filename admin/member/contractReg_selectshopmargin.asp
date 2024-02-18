<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 계약 관리 해당 브랜드 샵별 기본마진 구하기~
' Hieditor : 2010.05.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->

<%
dim shopid , makerid , sqlStr , defaultmargin , shopname
	makerid = request("makerid")
	shopid = request("shopid")

if 	shopid = "" or makerid = "" then
	response.write "<script>alert(파라메타값이 없습니다.관리자문의);</script>"
	dbget.close()	: response.end
end if

sqlStr = "select top 1" +vbcrlf
sqlStr = sqlStr & " s.shopid ,s.makerid ,isnull(s.defaultmargin,0) as defaultmargin, u.shopname" +vbcrlf
sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_designer s" +vbcrlf
sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user u" +vbcrlf
sqlStr = sqlStr & " on s.shopid = u.userid" +vbcrlf
sqlStr = sqlStr & " where s.shopid = '"&shopid&"' and s.makerid = '"&makerid&"'"

'response.write sqlStr &"<br>"
rsget.Open sqlStr,dbget,1
if not rsget.EOF  then
    defaultmargin = rsget("defaultmargin")
    shopname = rsget("shopname")
end if
rsget.close

'response.write defaultmargin

if defaultmargin <> "" then
%>
	<script language='javascript'>
		//alert('선택하신 브랜드에 대한 <%=shopname%>샵 기본마진이 <%= defaultmargin %>% 입니다.\n마진이 틀릴경우 기본마진을 직접 입력하세요.\n ex) 35%');
		parent.frmReg.$$DEFAULT_MARGIN$$.value = '<%= defaultmargin %>%';
		parent.frmReg.$$A_STORE$$.value = '<%= shopname %>';
	</script>
	
<% 
	dbget.close()	: response.end
else
%>

<script language='javascript'>
	alert('해당 브랜드에 대한 샵 마진정보가 없습니다. 매장과 기본마진 란은 직접 입력하세요.\n ex) 35');
	parent.frmReg.$$DEFAULT_MARGIN$$.value = '0%';
	parent.frmReg.$$A_STORE$$.value = '';	
</script>

<% 
end if
	dbget.close()	: response.end
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->