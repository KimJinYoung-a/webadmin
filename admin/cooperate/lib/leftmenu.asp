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
	Dim vMenu
	vMenu = NullFillWith(requestCheckVar(Request("mn"),10),"0")

	Dim cooperateView, vSSC1, vSSC2, vSSC3, vSSC4, vSSC5, vCSC1, vCSC2, vCSC3, vCSC4, vCSC5, vReferCnt, vIsNew1, vIsNew2, vIsNew3, vIsNew4, vIsNew5, vIsNew6	'### vSSC - SendStateCount, vCSC - ComeStateCount
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
	
	cooperateView.FGubun = "ISNEW"
	cooperateView.fnGetCooperateCount_PopVer
	vIsNew1		= cooperateView.FState1Cnt
	vIsNew2		= cooperateView.FState2Cnt
	vIsNew3		= cooperateView.FState3Cnt
	vIsNew4		= cooperateView.FState4Cnt
	vIsNew5		= cooperateView.FState5Cnt
	vIsNew6		= cooperateView.FReferCnt
	Set cooperateView = Nothing
%>

<script language="javascript">
function jsGoMenu(mn){
	top.location.href = "/admin/cooperate/popIndex.asp?mn="+mn+""; 
}

function jsNewReg() {
 	 var winNewReg = window.open("/admin/cooperate/cooperate_write.asp","winNewReg","width=900, height=800, scrollbars=yes, resizeable=yes");
	 winNewReg.focus();
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" border="0">
<tr height="15">
	<td nowrap ><a href="javascript:jsGoMenu('C00');" ><img src="/images/paper2.gif" border="0"> <%=CHKIIF(vMenu="C00","<font color=""#4E9FC6""><b>업무협조 Home</b></font>","업무협조 Home")%></a></td>
</tr>
<tr height="15">
	<td nowrap ><a href="javascript:jsNewReg();"><img src="/images/paper2.gif" border="0"> 신규등록</a></td>
</tr>
<tr nowrap valign="top">
	<td><img src="/images/openfolder.png" align="absmidde" id="imgR" border="0">&nbsp;<a href="javascript:jsGoMenu('C10');"><%=CHKIIF(vMenu="C10","<font color=""#4E9FC6""><b>보낸업무협조</b></font>","보낸업무협조")%></a>
		<table width="100%"  align="center" cellpadding="1" cellspacing="1" class="a" border="0" >
		<tr>
			<td>
				<table width="100%"  align="center" cellpadding="0" cellspacing="1" class="a" border="0" >
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('C11');"><%=CHKIIF(vMenu="C11","<font color=""#4E9FC6""><b>기안 ("&vSSC1&")</b></font>","기안 ("&vSSC1&")")%></a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('C12');"><%=CHKIIF(vMenu="C12","<font color=""#4E9FC6""><b>작업중 ("&vSSC2&")</b></font>","작업중 ("&vSSC2&")")%></a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('C13');"><%=CHKIIF(vMenu="C13","<font color=""#4E9FC6""><b>작업완료 ("&vSSC3&")</b></font>","작업완료 ("&vSSC3&")")%></a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('C14');"><%=CHKIIF(vMenu="C14","<font color=""#4E9FC6""><b>반려 ("&vSSC4&")</b></font>","반려 ("&vSSC4&")")%></a></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr nowrap valign="top">
	<td><img src="/images/openfolder.png" align="absmidde" id="imgR" border="0">&nbsp;<a href="javascript:jsGoMenu('C20');"><%=CHKIIF(vMenu="C20","<font color=""#4E9FC6""><b>받은업무협조</b></font>","받은업무협조")%></a>
		<table width="100%"  align="center" cellpadding="1" cellspacing="1" class="a" border="0" >
		<tr>
			<td>
				<table width="100%"  align="center" cellpadding="0" cellspacing="1" class="a" border="0" >
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('C21');"><%=CHKIIF(vMenu="C21","<font color=""#4E9FC6""><b>기안 ("&vCSC1&")</b></font>","기안 ("&vCSC1&")")%></a> <%=CHKIIF(vIsNew1<>0,"<span style=""vertical-align：top;border:1;font-size:10px;color:blue;""> new</span>","")%></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('C22');"><%=CHKIIF(vMenu="C22","<font color=""#4E9FC6""><b>작업중 ("&vCSC2&")</b></font>","작업중 ("&vCSC2&")")%></a> <%=CHKIIF(vIsNew2<>0,"<span style=""vertical-align：top;border:1;font-size:10px;color:blue;""> new</span>","")%></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('C23');"><%=CHKIIF(vMenu="C23","<font color=""#4E9FC6""><b>작업완료 ("&vCSC3&")</b></font>","작업완료 ("&vCSC3&")")%></a> <%=CHKIIF(vIsNew3<>0,"<span style=""vertical-align：top;border:1;font-size:10px;color:blue;""> new</span>","")%></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('C24');"><%=CHKIIF(vMenu="C24","<font color=""#4E9FC6""><b>반려 ("&vCSC4&")</b></font>","반려 ("&vCSC4&")")%></a> <%=CHKIIF(vIsNew4<>0,"<span style=""vertical-align：top;border:1;font-size:10px;color:blue;""> new</span>","")%></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('C26');"><%=CHKIIF(vMenu="C26","<font color=""#4E9FC6""><b>참조 ("&vReferCnt&")</b></font>","참조 ("&vReferCnt&")")%></a> <%=CHKIIF(vIsNew6<>0,"<span style=""vertical-align：top;border:1;font-size:10px;color:blue;""> new</span>","")%></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<br><br><br>
<input type="button" class="button" value="<%=CHKIIF(g_VertiHoriz="h","읽기창:오른쪽으로","읽기창:아래쪽으로")%>" onClick="top.document.location.href='/admin/cooperate/vertihoriz.asp?mn=<%=Request("mn")%>';">

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->