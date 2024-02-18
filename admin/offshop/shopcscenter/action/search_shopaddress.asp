<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->

<%
dim sql , shopid ,shopname ,shopphone ,shopzipcode ,shopaddr1 ,shopaddr2 ,mode
	shopid = requestCheckVar(request("shopid"),32)
	mode = requestCheckVar(request("mode"),10)

if shopid = "" then response.end	:	dbget.close()

if mode = "shopselected" then
	sql = " select top 1"
	sql = sql & " userid,shopname ,shopphone ,shopzipcode ,shopaddr1 ,shopaddr2"
	sql = sql & " from [db_shop].[dbo].tbl_shop_user"
	sql = sql & " where isusing='Y' "
	sql = sql & " and userid = '"&shopid&"'"
	
	rsget.open sql,dbget,1
	if  not rsget.EOF  then
		shopid = rsget("userid")
		shopname =  rsget("shopname")
		shopphone =  rsget("shopphone")
		shopzipcode =  rsget("shopzipcode")
		shopaddr1 =  db2html(rsget("shopaddr1"))
		shopaddr2 =  db2html(rsget("shopaddr2"))
	end if
	rsget.close
%>
	<% if shopid <> "" then %>

	<script type='text/javascript'>
		parent.deliveryaction('<%= shopname %>' ,'<%= shopphone %>' ,'<%= shopphone %>' ,'<%= shopzipcode %>' ,'<%= shopaddr1 %>' ,'<%= shopaddr2 %>')
	</script>
	
	<% end if %>

<% end if %>

<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->