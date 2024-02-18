<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전자결재 왼쪽메뉴
' History : 2011.03.14 정윤정  생성
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->  
<!-- #include virtual="/lib/db/dbopen.asp" -->  
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
Dim clsLeapp 
Dim iReportstate0,iReportstate1,iReportstate3,iReportstate5,iReportstate7
Dim iReportstate100,iReportstate110,iReportstate710,iReportstate130,iReportstate150 
Dim iReportstate101,iReportstate111,iReportstate711,iReportstate131,iReportstate151 
Dim iPayRequeststate9,iPayRequeststate1,iPayRequeststate5,iPayRequeststate7,iPayRequeststate0
Dim iPayRequeststate000,iPayRequeststate001,iPayRequeststate110,iPayRequeststate111,iPayRequeststate710,iPayRequeststate711,iPayRequeststate970,iPayRequeststate971,iPayRequeststate550,iPayRequeststate551
Dim irefercount, iauthcount,idoccount
Dim iRectMenu 
iRectMenu = requestCheckvar(Request("iRM"),10)
IF iRectMenu ="" THEN iRectMenu = "M999"
set clsLeapp = new CEApproval
clsLeapp.FadminId = session("ssBctId")
clsLeapp.fnGetLeftMenu
iReportstate0  = clsLeapp.FReportstate0  
iReportstate1 = clsLeapp.FReportstate1  
iReportstate3 = clsLeapp.FReportstate3  
iReportstate5 = clsLeapp.FReportstate5  
iReportstate7 = clsLeapp.FReportstate7  
iReportstate100= clsLeapp.FReportstate100 
iReportstate110= clsLeapp.FReportstate110 
iReportstate710= clsLeapp.FReportstate710 
iReportstate130= clsLeapp.FReportstate130 
iReportstate150= clsLeapp.FReportstate150  
iReportstate101= clsLeapp.FReportstate101 
iReportstate111= clsLeapp.FReportstate111 
iReportstate711= clsLeapp.FReportstate711 
iReportstate131= clsLeapp.FReportstate131 
iReportstate151= clsLeapp.FReportstate151
iPayRequeststate0= clsLeapp.FPayRequeststate0 
iPayRequeststate9= clsLeapp.FPayRequeststate9 
iPayRequeststate1= clsLeapp.FPayRequeststate1 
iPayRequeststate5= clsLeapp.FPayRequeststate5 
iPayRequeststate7= clsLeapp.FPayRequeststate7  
irefercount		= clsLeapp.FreferCount 
iauthcount		= clsLeapp.FauthCount
iPayRequeststate000	= clsLeapp.FPayRequeststate000
iPayRequeststate001 = clsLeapp.FPayRequeststate001
iPayRequeststate110 = clsLeapp.FPayRequeststate110
iPayRequeststate111 = clsLeapp.FPayRequeststate111
iPayRequeststate710 = clsLeapp.FPayRequeststate710
iPayRequeststate711 = clsLeapp.FPayRequeststate711
iPayRequeststate970 = clsLeapp.FPayRequeststate970
iPayRequeststate971 = clsLeapp.FPayRequeststate971
iPayRequeststate550 = clsLeapp.FPayRequeststate550
iPayRequeststate551 = clsLeapp.FPayRequeststate551
idoccount						= clsLeapp.FDocCount
set clsLeapp = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<html>
<head>
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->
<script language="javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">
	function jsGoMenu(iRMenu){
		top.location.href = "/admin/approval/eapp/popIndex.asp?iRM="+iRMenu; 
	}
</script>
</head>
<body leftmargin ="0" topmargin="0"	>
<table width="100%" height="100%" align="center" cellpadding="3" cellspacing="0" class="a"   border="0">    
<tr height="15">
	<td nowrap ><a href="javascript:jsGoMenu('M999');" ><img src="/images/paper2.gif" border="0"> <%IF iRectMenu="M999" THEN%><font color="#4E9FC6"><b><%END IF%>전자결재홈</a></td>
</tr>
<tr height="15">
	<td nowrap ><a href="javascript:jsPopView('/admin/approval/eapp/regeappform.asp');"><img src="/images/paper2.gif" border="0"> <%IF iRectMenu="M000" THEN%><font color="#4E9FC6"><b><%END IF%>신규작성</a></td>
</tr>
<tr>
	<td nowrap valign="top" height="30"><img src="/images/openfolder.png" align="absmidde" id="imgS" border="0">&nbsp;보낸결재함 
		<table width="100%"  align="center" cellpadding="1" cellspacing="1" class="a" border="0">
		<tr>
			<td style="padding-left:10px;"><a href="javascript:jsGoMenu('T010');"><img src="/images/openfolder.png" align="absmidde" id="imgS1" border="0">&nbsp;<%IF iRectMenu="T010" THEN%><font color="#4E9FC6"><b><%END IF%>결재문서</a> 
				<table width="100%"  align="center" cellpadding="1" cellspacing="1" class="a" border="0">
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M010');"><%IF iRectMenu="M010" THEN%><font color="#4E9FC6"><b><%END IF%>작성중(임시저장) (<%=iReportstate0%>)</a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M011');"><%IF iRectMenu="M011" THEN%><font color="#4E9FC6"><b><%END IF%>진행문서</a> (<%=iReportstate1%>)</td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M017');"><%IF iRectMenu="M017" THEN%><font color="#4E9FC6"><b><%END IF%>승인문서</a> (<%=iReportstate7%>)</td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M013');"><font color="gray"><%IF iRectMenu="M013" THEN%><font color="#4E9FC6"><b><%END IF%>보류문서 (<%=iReportstate3%>)</font></a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M015');"><font color="gray"><%IF iRectMenu="M015" THEN%><font color="#4E9FC6"><b><%END IF%>반려문서 (<%=iReportstate5%>)</font></a></td>
				</tr>
				</table> 
			</td>
		</tr> 
		<tr>
			<td style="padding-left:10px;"><a href="javascript:jsGoMenu('T020');"><img src="/images/openfolder.png" align="absmidde" id="imgS2" border="0">&nbsp;결제요청서</a>
				<table width="100%"  align="center" cellpadding="1" cellspacing="1" class="a" border="0">
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M020');" ><%IF iRectMenu="M020" THEN%><font color="#4E9FC6"><b><%END IF%>작성중(임시저장) (<%=iPayRequeststate0%>)</a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M021');" ><%IF iRectMenu="M021" THEN%><font color="#4E9FC6"><b><%END IF%>결제요청 (<%=iPayRequeststate1%>)</a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M027');" ><%IF iRectMenu="M027" THEN%><font color="#4E9FC6"><b><%END IF%>결제승인 (<%=iPayRequeststate7%>)</a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M029');" ><%IF iRectMenu="M029" THEN%><font color="#4E9FC6"><b><%END IF%>결제완료 (<%=iPayRequeststate9%>)</a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M025');" ><font color="gray"><%IF iRectMenu="M025" THEN%><font color="#4E9FC6"><b><%END IF%>결제반려 (<%=iPayRequeststate5%>)</font></a></td>
				</tr>
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M028');" ><font color="gray"><%IF iRectMenu="M025" THEN%><font color="#4E9FC6"><b><%END IF%>계산서차후수취관리 (<%=idoccount%>)</font></a></td>
				</tr>
				</table> 
			</td>
		</tr> 
		</table>	 
	</td>
</tr>
<tr nowrap valign="top">
	<td><img src="/images/openfolder.png" align="absmidde" id="imgR" border="0">&nbsp;받은결재함  
		<table width="100%"  align="center" cellpadding="1" cellspacing="1" class="a" border="0" >	
		<tr>
			<td style="padding-left:10px;"><a href="javascript:jsGoMenu('T011');"><img src="/images/openfolder.png" align="absmidde" id="imgR1" border="0">&nbsp;결재문서</a>
					<table width="100%"  align="center" cellpadding="0" cellspacing="1" class="a" border="0" >	
					<tr>
						<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M110');"><%IF iRectMenu="M110" THEN%><font color="#4E9FC6"><b><%END IF%>결재대기 (<%=iReportstate100+iReportstate101%>)<%IF iReportstate100>0 THEN%><span  style="vertical-align：top;border:1;font-size:10px;color:blue;"> new</span><%END IF%></a></td>
					</tr>
					<tr>
						<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M111');"><%IF iRectMenu="M111" THEN%><font color="#4E9FC6"><b><%END IF%>결재완료(진행중) (<%=iReportstate110+iReportstate111%>)<%IF iReportstate110>0 THEN%><span  style="vertical-align：top;border:1;font-size:10px;color:blue;"> new</span><%END IF%></a></td>
					</tr>
					<tr>
						<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M171');"><%IF iRectMenu="M171" THEN%><font color="#4E9FC6"><b><%END IF%>결재완료(최종승인) (<%=iReportstate710+iReportstate711%>)<%IF iReportstate710>0 THEN%><span  style="vertical-align：top;border:1;font-size:10px;color:blue;"> new</span><%END IF%></a></td>
					</tr>
					<tr>
						<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M113');"><font color="gray"><%IF iRectMenu="M113" THEN%><font color="#4E9FC6"><b><%END IF%>결재보류 (<%=iReportstate130+iReportstate131%>)<%IF iReportstate130>0 THEN%><span  style="vertical-align：top;border:1;font-size:10px;color:blue;"> new</span><%END IF%></a></td>
					</tr>
					<tr>
						<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M115');"><font color="gray"><%IF iRectMenu="M115" THEN%><font color="#4E9FC6"><b><%END IF%>결재반려 (<%=iReportstate150+iReportstate151%>)<%IF iReportstate150>0 THEN%><span  style="vertical-align：top;border:1;font-size:10px;color:blue;"> new</span><%END IF%></a></td>
					</tr>
					<tr>
						<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M112');"><font color="gray"><%IF iRectMenu="M112" THEN%><font color="#4E9FC6"><b><%END IF%>참조 (<%=irefercount%>)</font></td>
					</tr>
					</table>	 
			</td>
		</tr> 
		<tr>
			<td style="padding-left:10px;"><a href="javascript:jsGoMenu('T021');"><img src="/images/openfolder.png" align="absmidde" id="imgR2" border="0">&nbsp;결제요청서</a>
				<table width="100%"  align="center" cellpadding="1" cellspacing="1" class="a" border="0" >
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('M120');"><%IF iRectMenu="M120" THEN%><font color="#4E9FC6"><b><%END IF%>결재선 (<%=iauthcount%>)</a></td>
				</tr>  
				</table>
			</td>
		</tr>  
		</table>   
	</td>
</tr> 
<%IF session("ssAdminPsn") =	8 OR session("ssAdminLsn") <= 2 THEN%>
<tr nowrap valign="top">
	<td><hr width=100%> <img src="/images/openfolder.png" align="absmidde" id="imgR" border="0">&nbsp;<font color="blue">재무회계</font>  
		<table width="100%"  align="center" cellpadding="1" cellspacing="1" class="a" border="0" >	
		<tr>
			<td style="padding-left:10px;">
				<table width="100%"  align="center" cellpadding="1" cellspacing="1" class="a" border="0" >  
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F100');"><%IF iRectMenu="F100" THEN%><font color="#4E9FC6"><b><%END IF%>결제요청전승인 (<font color="blue"><%=iPayRequeststate000%></font>/<%=iPayRequeststate001%>)</a></td>
				</tr>
			
				<tr>
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F110');"><%IF iRectMenu="F110" THEN%><font color="#4E9FC6"><b><%END IF%>결제요청 (<font color="blue"><%=iPayRequeststate110%></font>/<%=iPayRequeststate111%>)</a></td>
				</tr>	
				<tr>	
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F711');"><%IF iRectMenu="F711" THEN%><font color="#4E9FC6"><b><%END IF%>결제확인(결제예정) (<font color="blue"><%=iPayRequeststate710%></font>/<%=iPayRequeststate711%>)</a></td>
				</tr>	
				<tr>	
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F971');"><%IF iRectMenu="F971" THEN%><font color="#4E9FC6"><b><%END IF%>결제완료 (<font color="blue"><%=iPayRequeststate970%></font>/<%=iPayRequeststate971%>)</a></td>
				</tr>	
				<tr>	
					<td style="padding-left:15px;"><a href="javascript:jsGoMenu('F551');"><%IF iRectMenu="F551" THEN%><font color="#4E9FC6"><b><%ELSE%><font color="gray"><%END IF%>결제반려 (<font color="blue"><%=iPayRequeststate550%></font>/<%=iPayRequeststate551%>)</font></a></td>
				</tr> 
				</table> 
			</td>
		</tr> 
<%END IF%> 
	</table>
</td>
</tr>
</table>
</body>
</html> 
	
 
	
	
	
	
	
	
