 <%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전자결재 폼 선택
' History : 2011.03.14 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" -->  
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
Dim clseapp
Dim iTotSendCount,iTotReceiveCount,iTotReceiveViewCount,iTotpaySendCount 
Dim iReportstate0,iReportstate1,iReportstate3,iReportstate5,iReportstate7
Dim iReportstate100,iReportstate110,iReportstate710,iReportstate130,iReportstate150 
Dim iReportstate101,iReportstate111,iReportstate711,iReportstate131,iReportstate151 
Dim iPayRequeststate9,iPayRequeststate1,iPayRequeststate5,iPayRequeststate7
Dim irefercount, iauthcount
set clseapp = new CEApproval
	clseapp.FadminId = session("ssBctId")
	clseapp.fnGetMainCount
	
	iTotSendCount        = clseapp.FTotSendCount
	iTotReceiveCount     = clseapp.FTotReceiveCount
	iTotReceiveViewCount = clseapp.FTotReceiveViewCount
	iTotpaySendCount     = clseapp.FTotpaySendCount
	iauthCount			 = clseapp.FauthCount 
		 
	clseapp.fnGetLeftMenu
	iReportstate0  = clseapp.FReportstate0  
	iReportstate1 = clseapp.FReportstate1  
	iReportstate3 = clseapp.FReportstate3  
	iReportstate5 = clseapp.FReportstate5  
	iReportstate7 = clseapp.FReportstate7  
	iReportstate100= clseapp.FReportstate100 
	iReportstate110= clseapp.FReportstate110 
	iReportstate710= clseapp.FReportstate710 
	iReportstate130= clseapp.FReportstate130 
	iReportstate150= clseapp.FReportstate150  
	iReportstate101= clseapp.FReportstate101 
	iReportstate111= clseapp.FReportstate111 
	iReportstate711= clseapp.FReportstate711 
	iReportstate131= clseapp.FReportstate131 
	iReportstate151= clseapp.FReportstate151
	iPayRequeststate9= clseapp.FPayRequeststate9 
	iPayRequeststate1= clseapp.FPayRequeststate1 
	iPayRequeststate5= clseapp.FPayRequeststate5 
	iPayRequeststate7= clseapp.FPayRequeststate7  
	irefercount		= clseapp.FreferCount  
set clseapp = nothing

%> 
<!-- #include virtual="/lib/db/dbclose.asp" -->
<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->
<script language="javascript">
	function jsGoMenu(iRMenu){
		top.location.href = "/admin/approval/eapp/popIndex.asp?iRM="+iRMenu; 
	}
</script>
</head>
<body leftmargin ="0" topmargin="0">
<table width="100%" height="100%" cellpadding="0" cellspacing="0"  border="0"  class="a" > 
<tr>
	<td valign="top">	
		<table width="100%" cellpadding="3" cellspacing="1" class="a" border="0">  
		<tr>
			<td>
				<table width="600" cellpadding="5" cellspacing="1" class="a" border="0"> 
				<tr>
					<td>보낸결재함<br> </td>
				</tr> 
				<tr>
					<td>
						<table width="100%" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor="#BABABA">
						<tr bgcolor="#EFEFEF" align="center"> 
							<td rowspan="2" width="100">결재문서</td>
							<td width="100">작성중</td>
							<td width="100">진행문서</td>
							<td width="100">승인문서</td>
							<td width="100">보류문서</td>
							<td width="100">반려문서</td>
						</tr>  
						<tr  bgcolor="#FFFFFF"  align="center">  
							<td><a href="javascript:jsGoMenu('M010');"><%=iReportstate0%></a></td>
							<td><a href="javascript:jsGoMenu('M011');"><%=iReportstate1%></a></td>
							<td><a href="javascript:jsGoMenu('M017');"><%=iReportstate7%></a></td>
							<td><a href="javascript:jsGoMenu('M013');"><%=iReportstate3%></a></td>
							<td><a href="javascript:jsGoMenu('M015');"><%=iReportstate5%></a></td>
						</tr> 
						<tr bgcolor="#EFEFEF" align="center">
							<td  rowspan="2">결제요청서</td>
							<td>결제요청</td>
							<td>결제승인</td>
							<td>결제완료</td>
							<td>결제반려</td>
							<td></td>
						</tr>
						<tr  bgcolor="#FFFFFF"  align="center">  
							<td><a href="javascript:jsGoMenu('M021');" ><%=iPayRequeststate1%></a></td>
							<td><a href="javascript:jsGoMenu('M027');" ><%=iPayRequeststate7%></a></td>
							<td><a href="javascript:jsGoMenu('M029');" ><%=iPayRequeststate9%></a></td>
							<td><a href="javascript:jsGoMenu('M025');" ><%=iPayRequeststate5%></a></td>
							<td></td>
						</tr>
						</table>
					</td>
				</tr> 
				<tr>
					<td style="padding-top:30px;">받은결재함<br> </td>
				</tr>
				<tr>
					<td>
						<table width="100%" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor="#BABABA">
						<tr bgcolor="#EFEFEF" align="center"> 
							<td rowspan="2" width="100">결재문서</td>
							<td width="100">결재대기</td>
							<td width="100">결재완료<br>(진행중)</td>
							<td width="100">결재완료<br>(최종승인)</td>
							<td width="100">결재보류</td>
							<td width="100">결재반려</td>
							<td width="100">결재참조</td>
						</tr>	
						<tr  bgcolor="#FFFFFF"  align="center">  
							<td><font color="blue"><a href="javascript:jsGoMenu('M110');"><%=iReportstate100+iReportstate101%></a></td>
							<td><font color="blue"><a href="javascript:jsGoMenu('M111');"><%=iReportstate110+iReportstate111%></a></td>
							<td><font color="blue"><a href="javascript:jsGoMenu('M171');"><%=iReportstate710+iReportstate711%></a></td>
							<td><font color="blue"><a href="javascript:jsGoMenu('M113');"><%=iReportstate130+iReportstate131%></a></td>
							<td><font color="blue"><a href="javascript:jsGoMenu('M115');"><%=iReportstate150+iReportstate151%></a></td>
							<td><a href="javascript:jsGoMenu('M112');"><%=irefercount%></a></td> 
						</tr>      
						<tr bgcolor="#EFEFEF" align="center">
							<td  rowspan="2">결제요청서</td>
							<td>결재선</td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
						<tr  bgcolor="#FFFFFF"  align="center">  
							<td><a href="javascript:jsGoMenu('M120');"><%=iauthCount%></a></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
						</tr>
						</table>
					</td>
				</tr>  
				</table> 
			</td>
		</tr> 
		</table>
	</td>
</tr> 
</table>
</body>
</html>
