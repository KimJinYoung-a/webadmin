<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �귣�� ��� ���� �ش� �귣�� ���� �⺻���� ���ϱ�~
' Hieditor : 2010.05.25 �ѿ�� ����
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
	response.write "<script>alert(�Ķ��Ÿ���� �����ϴ�.�����ڹ���);</script>"
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
		//alert('�����Ͻ� �귣�忡 ���� <%=shopname%>�� �⺻������ <%= defaultmargin %>% �Դϴ�.\n������ Ʋ����� �⺻������ ���� �Է��ϼ���.\n ex) 35%');
		parent.frmReg.$$DEFAULT_MARGIN$$.value = '<%= defaultmargin %>%';
		parent.frmReg.$$A_STORE$$.value = '<%= shopname %>';
	</script>
	
<% 
	dbget.close()	: response.end
else
%>

<script language='javascript'>
	alert('�ش� �귣�忡 ���� �� ���������� �����ϴ�. ����� �⺻���� ���� ���� �Է��ϼ���.\n ex) 35');
	parent.frmReg.$$DEFAULT_MARGIN$$.value = '0%';
	parent.frmReg.$$A_STORE$$.value = '';	
</script>

<% 
end if
	dbget.close()	: response.end
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->