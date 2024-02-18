<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����Ʈ �г�Ƽ ����Ʈ
' History : 2018.05.08 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
Dim searchKey, searchString, penaltyKind
Dim penaltyStateDiv
Dim cAgit, arrList, intLoop

iCurrPage		= requestCheckvar(request("iC"),10)
searchKey		= requestCheckvar(request("SearchKey"),1)
searchString	= requestCheckvar(request("SearchKey"),17)
penaltyKind 	= requestCheckvar(request("PenaltyKind"),1)
penaltyStateDiv	= requestCheckvar(request("PenaltyStateDiv"),1)

iPageSize		= 20
if iCurrPage ="" then iCurrPage =1

set cAgit = new CAgitPoint
	cAgit.FCurrPage 		= iCurrPage
	cAgit.FPageSize 		= iPageSize
	cAgit.FRectSearchKey	= SearchKey
	cAgit.FRectSearchString = SearchString
	cAgit.FRectPenaltyKind	= penaltyKind
	cAgit.FRectStateDiv		= penaltyStateDiv

	arrList = cAgit.fnGetAgitPenaltyList
	iTotCnt = cAgit.FTotCnt 
	iTotalPage = cAgit.FTotPage 
set cAgit	= nothing
%>
<script type="text/javascript">
function jsViewBook(idx){
 	var p = window.open("/admin/member/tenbyten/agit/pop_tenbyten_Agit_Edit_admin.asp?idx="+idx,"popNAgit","width=700,height=700,scrollbars=yes,resize=yes");
		p.focus();
 }
</script>
<!-- �˻� -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="iC" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�˻�:
			<select name="SearchKey" class="select">
				<option value="">::����::</option>
				<option value="1" >���̵�</option>
				<option value="2">����ڸ�</option>
				<option value="3">���</option>
			</select>
			<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
			&nbsp;
			����:
			<select name="PenaltyKind" class="select">
				<option value="">��ü</option>
				<option value="1" >5���� ���</option>
				<option value="2">���� ���</option>
				<option value="3">No-Show</option>
				<option value="3">������ �г�Ƽ</option>
			</select>
			&nbsp;
			<label><input type="checkbox" name="penaltyStateDiv" value="1" <%=chkIIF(penaltyStateDiv="1","checked","")%> /> ��������</label>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="�˻�">
		</td>
	</tr> 
</table>
</form>

<!-- ��� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="8">
			�˻���� : <b><%=iTotCnt%></b>
			&nbsp;
			������ : <b><%= iCurrPage %> / <%=iTotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>idx</td>
		<td>���</td>
		<td>ID</td>
		<td>�̸�</td>
		<td>����Ʈ</td>
		<td>�г�Ƽ �Ⱓ</td>
		<td>�����</td>
		<td>����/����</td>
	</tr>
<%
'' 0			1			2			3		4		5		6		7		8		9			10
'' �ҳ�Ƽ��ȣ	�����ȣ	�г�Ƽ����	������	������	�����	���	���̵�	�̸�	����Ʈ��ȣ	�г�Ƽ����
	if isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
%>
	<tr bgcolor="#ffffff" height="30">
		<td align="center"><a href="javascript:jsViewBook('<%=arrList(1,intLoop)%>');" title="���� ���ຸ��"><%=arrList(0,intLoop)%></a></td>
		<td align="center"><%=arrList(6,intLoop)%></td>
		<td align="center"><%=arrList(7,intLoop)%></td>
		<td align="center"><%=arrList(8,intLoop)%></td>
		<td align="center"><%=AgitName(arrList(9,intLoop))%></td>
		<td align="center"><a href="javascript:jsViewBook('<%=arrList(1,intLoop)%>');" title="���� ���ຸ��"><%=left(arrList(3,intLoop),10) & " ~ " & left(arrList(4,intLoop),10)%></a></td>
		<td align="center"><%=left(arrList(5,intLoop),10)%></td>
		<td align="center"><%=PenaltyKindName(arrList(2,intLoop)) & chkIIF(isNull(arrList(10,intLoop)),"","<br/>" & arrList(10,intLoop)) %></td>
	</tr>
<%
		Next
	End if
%>
</table>

<!-- ����¡ó�� --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
 <!-- #include virtual="/lib/db/dbclose.asp" -->