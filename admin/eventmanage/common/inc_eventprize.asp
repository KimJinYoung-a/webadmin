<%
'###########################################################
' Page : /admin/eventmanage/common/eventprize_regist.asp
' Description : ��÷�� ���ó�� include
' History : 2007.02.13 ������ ����
'###########################################################

 Dim cEvtPrize
 Dim arrPrize, intLoop
 Dim iTotCnt
 Dim iPageSize, iCurrpage ,iDelCnt
 Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
 Dim arrPrizeType, arrPrizeStatus
  
 IF egKindCode = "" THEN egKindCode = 0	
	
 iCurrpage 	= requestCheckVar(Request("iC"),10)		 '���� ������ ��ȣ
 IF iCurrpage = "" THEN	iCurrpage = 1
	
	iPageSize = 30		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
	set cEvtPrize = new ClsEventPrize
	cEvtPrize.FECode	  	= eCode			'�̺�Ʈ �ڵ�
	cEvtPrize.FEGKindCode 	= egKindCode	'�׷��ڵ�(�ΰŽ�,��ȭ�̺�Ʈ ȸ��)
	cEvtPrize.FCPage 		= iCurrpage
	cEvtPrize.FPSize 		= iPageSize
	arrPrize = cEvtPrize.fnGetPrize		'��÷����
	iTotCnt = cEvtPrize.FTotCnt			'��ü ������  ��
	set cEvtPrize = nothing
	arrPrizeType = fnSetCommonCodeArr("evtprizetype",False)
	arrPrizeStatus= fnSetCommonCodeArr("evtprizestatus",False)
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��	
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language="javascript">
<!--
//��÷�� ���
  function jsSetWinner(eC,egKC,epC){
  	var winW, popURL;
  	if (epC > 0){
  		popURL ="/admin/eventmanage/event/pop_event_changewinner.asp?epC="+epC;  		
  	}else{
  		popURL="/admin/eventmanage/event/pop_event_winner.asp?eC="+eC+"&egKC="+egKC;
  	}
  	winW = window.open(popURL,'popW','width=1000, height=700, scrollbars=yes');
  	winW.focus();
  }
  
  //��÷ ���
  
  	//����¡ó��
		function jsGoPage(iP){
		document.frmPrize.iC.value = iP;
		document.frmPrize.submit();
	}

function tnCheckAll(bool, comp){
    var frm = comp.form;
    if (!comp.length){
        comp.checked = bool;
    }else{
        for (var i=0;i<comp.length;i++){
            comp[i].checked = bool;
        }
    }
}

function jsSMSSendPop(){
	if($("input:checkbox[name='cksel']:checked").length<1){
		alert("�߼۴���� �������ּ���.");
	}else{
		frm = document.frmPrize;
		window.open('', 'popSMS', 'width=500, height=700');
		frm.action = "/admin/eventmanage/common/pop_prize_sms_send.asp";
		frm.target = "popSMS";
		frm.method = "post";
		frm.submit();
	}
}
//-->
</script>
<table width="100%" border="0" align="left" class="a" cellpadding="0" cellspacing="1">
	<tr>
		<td>
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1">
			<tr>
				<td>
				<input type="button" name="btnadd"  value="SMS ����" onClick="javascript:jsSMSSendPop();" class="button">
				<input type="button" name="btnadd"  value="�� ��÷���" onClick="javascript:jsSetWinner(<%=eCode%>,<%=egKindCode%>,0);" class="button">
				</td>
			</tr>	
			</table>
		</td>	
	<tr>
		<td>
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmPrize" method="post" >
			<input type="hidden" name="menupos" value="<%=menupos%>">
			<input type="hidden" name="iC" value="<%=iCurrpage%>">
			<input type="hidden" name="eC" value="<%=eCode%>">	
			<input type="hidden" name="egKC" value="<%=egKindCode%>">			
			<tr bgcolor="#FFFFFF" height="25">
				<td colspan="10">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
			</tr>		
			<tr>
				<td align="center"  width="50" bgcolor="<%= adminColor("tabletop") %>"><input type="checkbox" onClick="tnCheckAll(this.checked,frmPrize.cksel);" /></td>
				<td align="center"  width="50" bgcolor="<%= adminColor("tabletop") %>">��÷�ڵ�</td>							
				<td align="center"  width="50" bgcolor="<%= adminColor("tabletop") %>">���</td>
				<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�����Ī</td>
				<td align="center"  width="70" bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td align="center"  bgcolor="<%= adminColor("tabletop") %>">����ǰ��(��ǰ��ȣ)</td>							
				<td align="center"  width="100" bgcolor="<%= adminColor("tabletop") %>">��÷��</td>
				<td align="center"  width="150"  bgcolor="<%= adminColor("tabletop") %>">��÷Ȯ�αⰣ</td>
				<td align="center"  width="100" bgcolor="<%= adminColor("tabletop") %>">����</td>				
				<td align="center"  width="60" bgcolor="<%= adminColor("tabletop") %>">�絵<br>��÷�ڵ�</td>
			</tr>
			<%IF isArray(arrPrize) THEN%>	
				<%For intLoop = 0 To UBound(arrPrize,2)	%>
				<tr>
					<td bgcolor="#FFFFFF" align="center"><input type='checkbox' name='cksel' id="cksel<%=intLoop%>" value='<%=arrPrize(0,intLoop)%>' /></td>
					<td bgcolor="#FFFFFF" align="center"><%=arrPrize(0,intLoop)%></td>
					<td bgcolor="#FFFFFF" align="center"><%=arrPrize(1,intLoop)%></td>
					<td bgcolor="#FFFFFF" align="center"><%=arrPrize(2,intLoop)%></td>
					<td bgcolor="#FFFFFF" align="center"><%=fnGetCommCodeArrDesc(arrPrizeType,arrPrize(14,intLoop))%></td>
					<td bgcolor="#FFFFFF"  align="left">&nbsp;<%=arrPrize(11,intLoop)%><%IF arrPrize(13,intLoop) <> 0 THEN%>(<%=arrPrize(13,intLoop)%>)<%END IF%></td>
					<td bgcolor="#FFFFFF"  align="center"><%=arrPrize(5,intLoop)%></td>
					<td bgcolor="#FFFFFF" align="left">&nbsp;<%if arrPrize(7,intLoop) <> "1900-01-01" then%><%=arrPrize(7,intLoop)%> ~ <%if arrPrize(8,intLoop) <> "1900-01-01" then%><%=arrPrize(8,intLoop)%><%end if%><%end if%></td>
					<td bgcolor="#FFFFFF" align="center">
						<%IF arrPrize(9,intLoop) = 5 THEN %>							
							<input type="button" class="button" value="�絵��û" onClick="jsSetWinner(<%=eCode%>,<%=egKindCode%>,<%=arrPrize(0,intLoop)%>);">
						<%ELSEIF datediff("d",arrPrize(8,intLoop),date()) > 0 AND  arrPrize(9,intLoop) = 0 THEN%>								
							<input type="button" class="button" value="�Ⱓ����" onClick="jsSetWinner(<%=eCode%>,<%=egKindCode%>,<%=arrPrize(0,intLoop)%>);">
						<%ELSE%>	
							<%=fnGetCommCodeArrDesc(arrPrizeStatus,arrPrize(9,intLoop))%>
						<%END IF%>		
						</td>	
					<td bgcolor="#FFFFFF" align="center"><%=arrPrize(15,intLoop)%></td>
				</tr>	
				<%Next%>				
			<%else%>	
				<tr>
					<td bgcolor="#FFFFFF" colspan="10" align="center">��÷������ �����ϴ�.</td>
				</tr>
			<%END IF%>	
			</table>	
		</td>
			
	</tr>
		
		
	<tr>
		<td>
			<!-- ����¡ó�� -->
			<%
			iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1
			
			If (iCurrpage mod iPerCnt) = 0 Then
				iEndPage = iCurrpage
			Else
				iEndPage = iStartPage + (iPerCnt-1)
			End If
			%>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			    <tr valign="bottom" height="25">        
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(iCurrpage) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>        
			    </tr>    
			    </form>
			</table>
		</td>
	</tr>
</table>	