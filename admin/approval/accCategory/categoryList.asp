<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ī�װ� ����Ʈ
' History : 2012.08.07 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/accCategoryCls.asp"-->
<%
Dim clsAcc, arrList, intLoop
Dim icidx, idepth, ipcateidx
Dim iTotCnt 
ipcateidx	= requestCheckvar(Request("selCL"),10)
IF ipcateidx = "" THEN ipcateidx = 0
IF ipcateidx = 0 THEN
	idepth = 1
ELSE
	idepth = 2
END IF	
Set clsAcc = new CAccCategory
	clsAcc.FACCDepth = idepth
	clsAcc.FACCPCateIdx	= ipcateidx
	arrList = clsAcc.fnGetAccCategoryList 	 
%>

<script language="javascript">
<!--
//���ε��
function jsNewReg(){
	var winD = window.open("popcategorydata.asp?selCL=<%=ipcateidx%>&menupos=<%=menupos%>","popD","width=400, height=300, resizable=yes, scrollbars=yes");
	winD.focus();
}
//����
function jsModReg(categoryidx){
	var winD = window.open("popcategorydata.asp?icidx="+categoryidx+"&menupos=<%=menupos%>","popD","width=400, height=300, resizable=yes, scrollbars=yes");
	winD.focus();
}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="page" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					�� ���� :
					<select name="selCL">
					<option value="0">--�ֻ���--</option>
					<% 
					clsAcc.sbGetOptACCCategory	 1,0,ipcateidx 
				  %>
					</select>
				</td> 
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<%
Set clsAcc = nothing 
%> 
<!-- #include virtual="/lib/db/dbclose.asp" --> 
<tr>
	<td><input type="button" class="button" value="�űԵ��" onClick="jsNewReg();"></td>
</tr>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
				<td>����Idx</td>
				<td>�����</td> 
				<td>ǥ�ü���</td> 
				<td>ó��</td>
			</tr>
			<%IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">				
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></a></td>	
				<td><%=arrList(1,intLoop)%></td>	
				<td><%=arrList(2,intLoop)%></td>
				<td><input type="button" class="button" value="����" onClick="jsModReg(<%=arrList(0,intLoop)%>)"  ></td>
			</tr>
		<%	Next
			ELSE%>
			<tr height=30 align="center" bgcolor="#FFFFFF">				
				<td colspan="3">��ϵ� ������ �����ϴ�.</td>	
			</tr>
			<%END IF%>
		</table>	
	</td>   
</tr> 
</table>
<!-- ������ �� -->
</body>
</html>
 



	