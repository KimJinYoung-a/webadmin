<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "euc-kr"
%>
<%
'####################################################
' Description : ī�װ�
' History : ���ʻ����ڸ�
'			2017.04.10 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<%
	if request("isWt")="W" then
		'��� ��� ��ǰ
		response.Write getDispCategoryWait(requestCheckVar(request("itemid"),10))
	else
		'�ǵ�� ��ǰ
		response.Write getDispCategory(requestCheckVar(request("itemid"),10))
	end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->