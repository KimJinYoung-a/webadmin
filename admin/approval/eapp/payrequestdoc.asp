<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������û�� ���ó ���ļ���  ����Ʈ
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
Dim arrList,intLoop  
	 
	iPageSize = 30
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
	 
	sadminId =  session("ssBctId")
 	ireportidx =  requestCheckvar(Request("iridx"),10) 
 	ipayrequestidx=  requestCheckvar(Request("iPRidx"),10) 
 	IF ipayrequestidx = "" THEN ipayrequestidx = 0
 		
'���� �⺻ �� ���� ��������
set clsPay = new CPayRequest
	clsPay.FadminId 	= sadminId 
	clsPay.FCurrpage 	= iCurrpage
	clsPay.FPagesize	= ipagesize
	arrList = clsPay.fnGetPayRequestDocList
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
<script language="javascript">
<!--   
	function jsMod(reportidx,payrequestidx){  
			top.eappDetail.location.href = "modpayrequestdoc.asp?iridx="+reportidx+"&ipridx="+payrequestidx+"&iRM=<%=iRectMenu%>"; 
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
			<td height="25"><font color="#4E9FC6"><b>���������� >������û��> ��꼭���ļ������ </b></font></td>
		</tr> 
		<tr>
			<td>  
		 	<!----------------- ����Ʈ ---------------------------> 
				<form name="frmList" method="post" action="" style="padding:0">
				<input type="hidden" name="iCP" value="<%=iCurrPage%>"> 
				<input type="hidden" name="iprs" value="<%=ipayrequeststate%>">
				<input type="hidden" name="iridx" value="<%=ireportidx%>">
				<input type="hidden" name="ipridx" value="<%=ipayrequestidx%>">
				<input type="hidden" name="iRM" value="<%=iRectMenu%>">
				</form>  
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