<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �ش� ���̵� ������ȣ�� ��Ʈ �ѹ��� ��ȯ(���۽�)
' History : 2016.04.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim empno, userid
	empno = requestcheckvar(request("empno"),10)
	userid = requestcheckvar(request("userid"),32)

if empno="" and userid="" then
	response.write "�����ڰ� �����ϴ�"
	dbget.close() : response.end
end if

'//������ ��Ʈ��ȣ
dim part_sn
	part_sn = getpart_sn(empno, userid)

response.write part_sn
dbget.close() : response.end
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->