<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : scm �α��� ������ ���̹��� ���� 
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<!-- #include virtual="/lib/classes/hitchhiker/scmMngCls.asp" -->
<%
  dim iTotCnt,iCurrPage,iPageSize,iPerCnt
  dim CScmMng
  dim arrList, intLoop
	iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
	set CScmMng = new ClsScmMng
	CScmMng.FCPage = iCurrpage
	CScmMng.FPSize = iPageSize
	arrList = CScmMng.fnGetScmMngList
	iTotCnt = CScmMng.FTotCnt
	set CScmMng = nothing
%>
<script type="text/javascript">
	function jsReg(sValue){
		location. href = "/admin/hitchhiker/scmMng/loginImgMng_reg.asp?menupos=<%=menupos%>&idx="+sValue;
	}
</script>
<table width="100%" align="center">
	<tr>
		<td align="right"><input type="button" name="newBT" class="button" value="���ȭ�� ����" onclick="jsReg('');"></td>
	</tr>
</table>


<% '����Ʈ-------------------------------------------------------------------------------------------------------- %>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="7" >  
			�˻���� : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr  align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td><b>idx</b></td>
		<td><b>���ȭ��</b></td>
		<td><b>�ۼ���</b></td>
		<td><b>�ۼ���</b></td>
	</tr>
	<%if isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
		%>
	<tr  align="center" bgcolor="#FFFFFF">
		<td><a href="javascript:jsReg(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></a></td>
	 	<td><%IF arrList(1,intLoop) <> "" then%><img src="<%=arrList(1,intLoop)%>" width="50"><%end if%></td>
	 	<td><%=arrList(4,intLoop)%>(<%=arrList(2,intLoop)%>)</td>
	 	<td><%=arrList(3,intLoop)%></td>
	</tr> 
<%	Next
	else%>
	<tr bgcolor="#FFFFFF">
	 	<td colspan="4" align="center">��ϵ� ������ �����ϴ�.</td>
	</tr>
	<%end if%>
</table>
<!-- ����¡ó�� --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, iPerCnt,menupos %>
		</td>
	</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
