<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������û�� ����Ʈ
' History : 2011.03.14 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
Dim clsPay
Dim sadminId,ireportidx,ipayrequestidx,ipayrequeststate
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim arrList,intLoop, reportname, reportprice, regdate, username
	reportname =  requestCheckvar(Request("reportname"),300)
	reportprice =  requestCheckvar(getNumeric(Request("reportprice")),10)
	regdate =  requestCheckvar(Request("regdate"),10)
	username =  requestCheckvar(Request("username"),32)
	 
	iPageSize = 30
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
	 
	sadminId =  session("ssBctId")
	ipayrequeststate	=requestCheckvar(Request("iprs"),10) 
 	ireportidx =  requestCheckvar(Request("iridx"),10) 
 	ipayrequestidx=  requestCheckvar(Request("iPRidx"),10) 
 	IF ipayrequestidx = "" THEN ipayrequestidx = 0
 		
'���� �⺻ �� ���� ��������
set clsPay = new CPayRequest
	clsPay.FadminId 	= sadminId
	clsPay.Fpayrequeststate = ipayrequeststate
	clsPay.FCurrpage 	= iCurrpage
	clsPay.FPagesize	= ipagesize
	clsPay.Freportname	= reportname
	clsPay.Freportprice	= reportprice
	clsPay.Fregdate	= regdate
	clsPay.Fusername	= username
	arrList = clsPay.fnGetPayRequestList
	iTotCnt = clsPay.FTotCnt 
set clsPay = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

 Dim iRectMenu 
 iRectMenu =  requestCheckvar(Request("iRM"),10)
%>
<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"--> 
<script type="text/javascript" src="eapp.js"></script>  
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--   
	function jsMod(reportidx,payrequestidx){  
			top.eappDetail.location.href = "regpayrequest.asp?iridx="+reportidx+"&ipridx="+payrequestidx+"&iRM=<%=iRectMenu%>"; 
	 }

	function frmsubmit(page){
		frmList.iCP.value = page;
		frmList.submit();
	}
//-->
</script>
</head>
<body leftmargin="0" topmargin="0">
	<div style="height:100%;overflow-y:auto;">
<table width="100%" height="100%" cellpadding="0" cellspacing="0"  border="0">
<tr> 
	<td valign="top">
		<table width="100%" cellpadding="0" cellspacing="1" class="a" border="0"> 
		<tr> 
			<td height="25"><font color="#4E9FC6"><b>���������� >������û�� <%IF ipayrequeststate>=0 THEN%>><%=fnGetMenu("S2",ipayrequeststate,"")%><%END IF%></b></font></td>
		</tr> 
		<tr>
			<td>
				<!-- �˻� ���� -->
				<form name="frmList" method="post" action="" style="padding:0">
				<input type="hidden" name="iCP" value="<%=iCurrPage%>"> 
				<input type="hidden" name="iprs" value="<%=ipayrequeststate%>">
				<input type="hidden" name="iridx" value="<%=ireportidx%>">
				<input type="hidden" name="ipridx" value="<%=ipayrequestidx%>">
				<input type="hidden" name="iRM" value="<%=iRectMenu%>">
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="#FFFFFF" >
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
					<td align="left">
						* ��û���� : <input type="text" name="reportname" id="reportname" value="<%= reportname %>" size="15" maxlength=150 class="text">
						<Br>
						* ��û�ݾ� : <input type="text" name="reportprice" id="reportprice" value="<%= reportprice %>" size="8" maxlength=10 class="text">
						* �ۼ��� : <input type="text" name="username" id="username" value="<%= username %>" size="8" maxlength=10 class="text">
					</td>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
						<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('1');">
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF" >
					<td align="left">
						* ��û�� : <input type="text" id="termSdt" name="regdate" size="8" maxlength=10 value="<%= regdate %>" />
						<img src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkStart_trigger" onclick="return false;" />
						<script type="text/javascript">
							var CAL_Start = new Calendar({
								inputField : "termSdt", trigger    : "ChkStart_trigger",
								onSelect: function() {
									var date = Calendar.intToDate(this.selection.get());
									//CAL_End.args.min = date;
									//CAL_End.redraw();
									this.hide();
								}, bottomBar: true, dateFormat: "%Y-%m-%d" <%'=chkIIF(regdate<>"",", max: " & replace(regdate,"-",""),"")%>
							});
						</script>
					</td>
				</tr>
				</table>
				</form>
				<!-- �˻� �� -->
				<Br>
			</td>
	 	</tr>
		<tr>
			<td>  
		 	<!----------------- ����Ʈ ---------------------------> 
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0"  bgcolor="#cccccc">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td>Idx</td> 
					<td>������û����</td> 
					<td nowrap>�����ݾ�</td>
					<td nowrap>������û��</td>
						<td>������</td> 
				</tr> 
				<%IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
				%>
				<tr  id="t<%=arrList(0,intLoop)%>" bgcolor="#FFFFFF" align="center" onclick="jsMod(<%=arrList(8,intLoop)%>,<%=arrList(0,intLoop)%>);ChangeColor('document.all.t<%=arrList(0,intLoop)%>','#CEF6EC','FFFFFF');" style="cursor:hand;"> 
					<td><%=arrList(0,intLoop)%></td>
					<td>������û��(<%=arrList(11,intLoop)%>)<br><font color="Gray"><%=arrList(9,intLoop)%></font></td> 
					<td><%=formatnumber(arrList(2,intLoop),0)%></td>
					<td><%IF arrList(1,intLoop) <> "" THEN%><%=formatdate(arrList(1,intLoop),"0000-00-00")%><%END IF%></td>  
					<td nowrap><%IF arrList(4,intLoop) <> "" THEN%><%=formatdate(arrList(4,intLoop),"0000-00-00")%><%END IF%></td>
				</tr> 
				<%	
					Next
					ELSE	
				%>
				<tr>
					<td colspan="8" align="center" bgcolor="#FFFFFF">��ϵ� ������ �����ϴ�.</td>
				</tr>
				<%END IF%>
				</table> 
			</td>
		</tr>
				 
				<!-- ������ ���� -->
						<%
						Dim iStartPage,iEndPage,iX,iPerCnt
						iPerCnt = 10
						
						iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1
						
						If (iCurrPage mod iPerCnt) = 0 Then
							iEndPage = iCurrPage
						Else
							iEndPage = iStartPage + (iPerCnt-1)
						End If
						%>
							<tr height="25" >
								<td colspan="15" align="center">
									<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
									    <tr valign="bottom" height="25">        
									        <td valign="bottom" align="center">
									         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
											<% else %>[pre]<% end if %>
									        <%
												for ix = iStartPage  to iEndPage
													if (ix > iTotalPage) then Exit for
													if Cint(ix) = Cint(iCurrPage) then
											%>
												<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
											<%		else %>
												<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
											<%
													end if
												next
											%>
									    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
											<% else %>[next]<% end if %>
									        </td>        
									    </tr>        
									</table>
							 
				<!-- ������ �� -->
			<!-----------------/ ����Ʈ ---------------------------> 
			</td>
		</tr>	 
		</table>
	</td> 	 
</tr>
</table>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->