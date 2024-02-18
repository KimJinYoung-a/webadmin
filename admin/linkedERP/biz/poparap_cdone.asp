<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 수지항목
' History : 2011.04.21 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpArapCls.asp"-->
<%
dim menupos, i, searcharap_cdname, clsAccount, iOpExpPartIdx, arrAccount
menupos = requestCheckvar(getNumeric(Request("menupos")),10)
searcharap_cdname = requestCheckvar(Request("searcharap_cdname"),50) 
iOpExpPartIdx = requestCheckvar(Request("selP"),10)

IF iOpExpPartIdx = "" THEN iOpExpPartIdx = 0

 '수지항목 리스트
set clsAccount = new COpExpAccount
	clsAccount.FOpExpPartIdx = iOpExpPartIdx
	clsAccount.frectarap_nm = searcharap_cdname
	arrAccount = clsAccount.fnGetArapRegList
set clsAccount = nothing
%>  
 
<script type="text/javascript">

//검색
function jsSearch(){  
document.frm.submit();

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">  
	<tr>
	<td><strong>부서  선택</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<form name="frm" method="post" action="" style="margin:0px;">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
		<tr align="center" bgcolor="#FFFFFF" > 
			<td align="left">&nbsp; 
			 수지항목: <input type="text" name="searcharap_cdname" size="30" value="<%=searcharap_cdname%>">
			</td>
			<td rowspan="2" width="50" bgcolor="#EEEEEE">
				<input type="button" class="button_s" value="검색" onClick="jsSearch();">
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>  
<tr>
	<td>
		<!-- 상단 띠 시작 --> 
		<form name="frmReg" method="post" style="margin:0px;">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   
				<td>수지항목</td>  
				<td>처리</td> 
			</tr>
			<%
			If isArray(arrAccount) THEN
			For i = 0 To UBound(arrAccount,2)
			%>
			<tr height=30 align="center" bgcolor="#FFFFFF"> 
				<td align="left">
					 <input type="hidden" name="arap_cd" value="<%=arrAccount(0,i)%>">
					 <input type="hidden" name="arap_name" value="<%=chkIIF(arrAccount(2,i),"[사용]","[지급]") & arrAccount(1,i)%>">
					 <%=chkIIF(arrAccount(2,i),"[사용]","[지급]") & arrAccount(1,i)%>
				</td>	 
				<td>
					<input type="button" class="button" value="선택" onClick="opener.jsSetarap_cd('<%=arrAccount(0,i)%>','<%=chkIIF(arrAccount(2,i),"[사용]","[지급]") & arrAccount(1,i)%>');self.close();">
				</td>
			</tr>
			<%
			Next
			ELSE
			%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="2">등록된 내용이 없습니다.</td>	
			</tr>
		<%END IF%>
		</table>	
		</form>
	</td> 
</tr>  
</table>
<!-- 페이지 끝 -->
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" --> 	



	