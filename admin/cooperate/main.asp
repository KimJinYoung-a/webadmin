<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->

<%
	Dim cooperateView, vSSC1, vSSC2, vSSC3, vSSC4, vSSC5, vCSC1, vCSC2, vCSC3, vCSC4, vCSC5, vReferCnt	'### vSSC - SendStateCount, vCSC - ComeStateCount
	Set cooperateView = New CCooperate
	cooperateView.FUserID = session("ssBctId")
	
	cooperateView.FGubun = "SEND"
	cooperateView.fnGetCooperateCount_PopVer
	vSSC1 = cooperateView.FState1Cnt
	vSSC2 = cooperateView.FState2Cnt
	vSSC3 = cooperateView.FState3Cnt
	vSSC4 = cooperateView.FState4Cnt
	vSSC5 = cooperateView.FState5Cnt
	
	cooperateView.FGubun = "COME"
	cooperateView.fnGetCooperateCount_PopVer
	vCSC1		= cooperateView.FState1Cnt
	vCSC2		= cooperateView.FState2Cnt
	vCSC3		= cooperateView.FState3Cnt
	vCSC4		= cooperateView.FState4Cnt
	vCSC5		= cooperateView.FState5Cnt
	vReferCnt	= cooperateView.FReferCnt
	Set cooperateView = Nothing
%>

<script language="javascript">
function jsGoMenu(mn){
	top.location.href = "/admin/cooperate/popIndex.asp?mn="+mn+""; 
}
</script>

<table cellpadding="5" cellspacing="1" class="a" border="0">
<tr>
	<td>보낸업무협조<br></td>
</tr>
<tr>
	<td>
		<table width="100%" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor="#BABABA">
		<tr bgcolor="#EFEFEF" align="center">
			<td width="110">기안</td>
			<td width="110">작업중</td>
			<td width="110">작업완료</td>
			<td width="110">반려</td>
			<td width="110">반려 후 최종완료</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td><a href="javascript:jsGoMenu('C11');"><%=vSSC1%></a></td>
			<td><a href="javascript:jsGoMenu('C12');"><%=vSSC2%></a></td>
			<td><a href="javascript:jsGoMenu('C13');"><%=vSSC3%></a></td>
			<td><a href="javascript:jsGoMenu('C14');"><%=vSSC4%></a></td>
			<td><a href="javascript:jsGoMenu('C15');"><%=vSSC5%></a></td>
		</tr>
		</table>
	</td>
</tr>
</table>

<table cellpadding="5" cellspacing="1" class="a" border="0">
<tr>
	<td style="padding-top:30px;">받은업무협조<br></td>
</tr>
<tr>
	<td>
		<table width="100%" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor="#BABABA">
		<tr bgcolor="#EFEFEF" align="center"> 
			<td width="110">기안</td>
			<td width="110">작업중</td>
			<td width="110">작업완료</td>
			<td width="110">반려</td>
			<td width="110">반려 후 최종완료</td>
			<td width="110">참조</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td><a href="javascript:jsGoMenu('C21');"><%=vCSC1%></a></td>
			<td><a href="javascript:jsGoMenu('C22');"><%=vCSC2%></a></td>
			<td><a href="javascript:jsGoMenu('C23');"><%=vCSC3%></a></td>
			<td><a href="javascript:jsGoMenu('C24');"><%=vCSC4%></a></td>
			<td><a href="javascript:jsGoMenu('C25');"><%=vCSC5%></a></td>
			<td><a href="javascript:jsGoMenu('C26');"><%=vReferCnt%></a></td>
		</tr>
		</table>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->