<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  �������� ����ٹ�����
' History : 2011.03.17 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/staff/staff_cls.asp"-->
<%
dim SearchText ,oMember, arrList, iTotCnt , SearchType , shopid
	SearchText = request("SearchText")
	SearchType = request("SearchType")
	shopid = request("shopid")
	
'// ���̵�� �˻�
Set oMember = new CAgitCalendar
	oMember.Frectshopid = shopid
	oMember.FrectSearchType 	= SearchType	'�˻�����
	oMember.FrectSearchText 	= SearchText
	oMember.Frectstatediv 		= "Y"			
	oMember.fnGetMemberList

IF oMember.ftotalcount > 0 THEN
%>

	<script language="javascript">
	
		parent.frm.chkCfm.value="Y";
		parent.frm.userid.value="<%= oMember.FOneItem.Fuserid %>";
		parent.frm.empno.value="<%= oMember.FOneItem.fempno %>";
		parent.frm.username.value="<%= oMember.FOneItem.fusername %>";
		parent.frm.posit_sn.value="<%= oMember.FOneItem.fposit_sn %>";
		parent.frm.part_sn.value="<%= oMember.FOneItem.fpart_sn %>";
		
	</script>

<% Else %>

	<script language="javascript">
		alert("�ش� ������ ���� ���� �ʽ��ϴ�\nȮ�� �� �ٽ� �˻����ּ���.");
	</script>

<% End If %>

<%
set oMember = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->