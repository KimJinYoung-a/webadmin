		<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��� �˻� �� ���� ����
' History : 2011.03.10 ������ ����
'						2017.04.06 ������ ������� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<%
dim userid, empno, sdt
	userid = requestCheckvar(request("uid"),32)
   empno = requestCheckvar(Request("sEn"),32)
   sdt =  requestCheckvar(Request("sdt"),10)
dim oMember, arrList, iTotCnt
dim clsap
dim totap, useap, reqap,payap, ispenalty, psdate, pedate, pkind

if empno="" then Response.End

'// ������� �˻�
Set oMember = new CTenByTenMember
oMember.FPagesize 		= 10
oMember.FCurrPage 		= 1
oMember.FSearchType 	= "3"	'�˻�����(ȸ��ID)
oMember.FSearchText 	= empno
oMember.Fstatediv 		= "Y"
oMember.Fextparttime 	= "0"	'0:�����, 1:�����̻�
	
arrList = oMember.fnGetMemberList
set oMember = nothing


'// ���̵� ã��� ���°��
IF isArray(arrList) THEN
	
set clsap = new CMyAgit
		clsap.FRectEmpno = empno
		clsap.FRectChkStart = sDt
		clsap.fnGetMyAgit
		totap = clsap.FtotPoint
		useap = clsap.FusePoint 
		pkind = clsap.Fpenaltykind
		psdate = clsap.Fpenaltysdate
		pedate = clsap.Fpenaltyedate
set clsap = nothing
%>
<script language="javascript"> 
	parent.frm.chkCfm.value="Y";	
	parent.frm.sEn.value="<%=arrList(0,0)%>";
	parent.frm.userid.value="<%=arrList(2,0)%>";
	parent.frm.username.value="<%=arrList(1,0)%>";
	parent.frm.posit_nm.value="<%=arrList(13,0)%>";
	parent.frm.avPoint.value="<%=totap-useap%>";
	parent.frm.department_nm.value="<%=arrList(27,0)%>"; 
	parent.frm.userPhone.value="<%=arrList(18,0)%>";
	parent.frm.userHP.value="<%=arrList(17,0)%>";
</script>
<%
Else
	'// �̸����� �˻�
	Set oMember = new CTenByTenMember
	oMember.FPagesize 		= 10
	oMember.FCurrPage 		= 1
	oMember.FSearchType 	= "2"	'�˻�����(ȸ����)
	oMember.FSearchText 	= empno
	oMember.Fstatediv 		= "Y"
	oMember.Fextparttime 	= "0"	'0:�����, 1:�����̻�
		
	arrList = oMember.fnGetMemberList
	iTotCnt = oMember.FTotCnt
	set oMember = nothing

	IF isArray(arrList) THEN
		IF iTotCnt>1 then
%>
<script language="javascript">
	var psu = window.open("popSelectUser.asp?unm=<%=Server.URLEncode(empno)%>","popSelUsr","width=500,height=200,scrollbars=yes");
	psu.focus();
</script>
<%
		else 
set clsap = new CMyAgit
		clsap.FRectEmpno = empno
		clsap.FRectChkStart = sDt
		clsap.fnGetMyAgit
		totap = clsap.FtotPoint
		useap = clsap.FusePoint 
		pkind = clsap.Fpenaltykind
		psdate = clsap.Fpenaltysdate
		pedate = clsap.Fpenaltyedate
set clsap = nothing

%>
<script language="javascript">
	parent.frm.chkCfm.value="Y";
	parent.frm.sEn.value="<%=arrList(0,0)%>";
	parent.frm.userid.value="<%=arrList(2,0)%>";
	parent.frm.username.value="<%=arrList(1,0)%>";
	parent.frm.posit_nm.value="<%=arrList(13,0)%>";
	parent.frm.avPoint.value="<%=totap-useap%>";
	parent.frm.department_nm.value="<%=arrList(27,0)%>"; 
	parent.frm.userPhone.value="<%=arrList(18,0)%>";
	parent.frm.userHP.value="<%=arrList(17,0)%>";
</script>
<%
		end if
	else
%>
<script language="javascript">
alert("�Է��Ͻ� [<%=empno%>]��(��) �ٹ����� ������ ���  �Ǵ� �̸��� �ƴմϴ�.\nȮ�� �� �ٽ� �˻����ּ���.");
</script>
<%
	End if
End If
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->