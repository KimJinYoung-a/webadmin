<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 서류제출여부 수정
' History : 2011.05.17 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"--> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<%
dim ipayrequestidx, blnTakeDoc
ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)
blnTakeDoc	= requestCheckvar(Request("blnTD"),10)
%>
<script language="javascript">
<!--
	function jsSubmitTakeDoc(){
		document.frm.submit();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
<form name="frm" method="post" action="procpayrequest.asp"> 
<input type="hidden" name="hidM" value="T">
<input type="hidden" name="ipridx" value="<%=ipayrequestIdx%>">
<tr>
	<td>서류제출여부 수정<br><hr width=100%></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
		<tr> 
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>"> 결제요청서 idx </td>
			<td bgcolor="#FFFFFF"><%=ipayrequestIdx%></td>
		</tr>
		<tr> 
			<td  align="center"  bgcolor="<%= adminColor("tabletop") %>"> 서류제출여부 </td>
			<td bgcolor="#FFFFFF"><input type="radio" name="rdoTD" value="1" <%IF blnTakeDoc THEN%>checked<%END IF%>>Y&nbsp;<input type="radio" name="rdoTD" value="0" <%IF not blnTakeDoc THEN%>checked<%END IF%>>N</td>
		</tr> 
		</table>
	</td>
</tr>
<Tr>
<td align="center" colspan="3" height="50"><input type="button" value="확인" class="button" onClick="jsSubmitTakeDoc();"></td>
</tr>
</form>
</table>
</body>
</html> 