<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ڱݰ��� �μ�
' History : 2011.04.21 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim clsBS
Dim arrList, intLoop, taxKey
Dim sUSEYN,sBS_NM,iType
Dim blnView, blnSale
sBS_NM = requestCheckvar(Request("sBS_NM"),100)
sUSEYN = requestCheckvar(Request("sUSEYN"),3)
iType = requestCheckvar(Request("iType"),1)
taxKey = request("taxKey")

blnView = "Y"
blnSale = "N"

Set clsBS = new CBizSection
	clsBS.FBS_NM 	= sBS_NM
	clsBS.FUSE_YN = sUSEYN
	clsBS.FView		= blnView
	clsBS.FSale		= blnSale
	arrList = clsBS.fnGetBizSectionList
Set clsBS = nothing
%>

<script language="javascript">
<!--

   //�˻�
   function jsSearch(){
    document.frm.submit();
   }
   function chromeOpenerFuncBug(a, b){
		window.opener.document.frmAct.mode.value = "modiBizSec"
		window.opener.document.frmAct.bizSecCd.value = a;
		window.opener.document.frmAct.taxKey.value = "<%= taxKey %>";
		window.opener.document.frmAct.matchSeq.value="0"
		window.opener.document.frmAct.submit();
		self.close();
   }
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
	<tr>
	<td><strong>�μ�  ����</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
		<form name="frm" method="post" action=""><% 'popGetBiz.asp %>
		<tr align="center" bgcolor="#FFFFFF" >
			<td align="left">&nbsp;
			 �μ���: <input type="text" name="sBS_NM" size="20" value="<%=sBS_NM%>">
			</td>
			<td rowspan="2" width="50" bgcolor="#EEEEEE">
				<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<form name="frmReg" method="post">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>�μ���</td>
				<td>ó��</td>
			</tr>
			<%  Dim oldPCD
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
					IF oldPCD <> arrList(2,intLoop) THEN
				%>
				<tr bgcolor="#FFFFFF"  height=30 >
					<td><%=arrList(2,intLoop)%>&nbsp; <%=arrList(4,intLoop)%></td>
					<td></td>
				</tr>
				<%END IF%>
			<tr height=30 align="center" bgcolor="<%IF arrList(3,intLoop) ="N" THEN%>#EFEFEF<%ELSE%>#FFFFFF<%END IF%>">
				<td align="left">  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					 �� <input type="hidden" name="sNM" value="<%=arrList(1,intLoop)%>">
					 <%=arrList(7,intLoop)%>&nbsp; <%=arrList(1,intLoop)%>
					 	<% if arrList(7,intLoop)<>arrList(0,intLoop) then %>
						&nbsp;<font color="#CCCCCC">(<%=arrList(0,intLoop)%>)</font>
						<% end if %>
					 <%IF arrList(4,intLoop) <> "" THEN%><input type="hidden" name="hidPM" value="<%=arrList(4,intLoop)%>"><%END IF%>
				</td>
				<!-- <td> <%IF arrList(2,intLoop) <> "" THEN%><input type="button" class="button" value="����" onClick="opener.jsSetPart('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>');self.close();"><% END IF %></td> -->
				<td> <%IF arrList(2,intLoop) <> "" THEN%><input type="button" class="button" value="����" onClick="chromeOpenerFuncBug('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>');"><% END IF %></td>
			</tr>
		<%		oldPCD  = arrList(2,intLoop)
				Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">
				<td colspan="2">��ϵ� ������ �����ϴ�.</td>
			</tr>
		<%END IF%>
		</table>
	</form>
	</td>
</tr>
</table>
<!-- ������ �� -->
</body>
</html>
 <!-- #include virtual="/lib/db/dbclose.asp" -->



