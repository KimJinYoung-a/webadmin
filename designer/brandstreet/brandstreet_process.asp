<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü �귣�������� ���� 
' History : 2009.03.26 �ѿ�� ����
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
	response.write "alert('������ �߻��߽��ϴ�. �ý������� �����ϼ���.');"
	response.write "self.close()"
	response.write "</script>"
	dbget.close()	:	response.End
	
end if

'//�ߴܹ��ó��
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

