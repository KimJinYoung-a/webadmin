<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventPrizeCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim clsEPrize
Dim arrList, intLoop
Dim iTotCnt, iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
Dim sSearchUserid, ievtprizeType, ievtprizeStatus,ievtCode,ieventkind, ievtName, vImg, vEGubun, vDate1, vDate2

iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
	sSearchUserid 	= requestCheckVar(Request("searchUserid"),32) 
	ievtprizeType 	= requestCheckVar(Request("evtprizetype"),4) 
	ievtprizeStatus = requestCheckVar(Request("evtprizestatus"),4) 
	ieventkind		= requestCheckVar(Request("eventkind"),4) 
	ievtCode		= requestCheckVar(Request("evtcode"),10)
	ievtName		= requestCheckVar(Request("evtname"),100)
	vEGubun			= NullFillWith(requestCheckVar(Request("egubun"),1),"e")
	vDate1			= request("date1")
	vDate2			= request("date2")
	
	set clsEPrize = new CEventPrize
	clsEPrize.FGubun = vEGubun
	clsEPrize.FSUserid = sSearchUserid
	clsEPrize.FRectRegDate1 = vDate1
	clsEPrize.FRectRegDate2 = vDate2
	clsEPrize.FEKind	= ieventkind
	clsEPrize.FEPType	= ievtprizeType
	clsEPrize.FEPStatus = ievtprizeStatus
	clsEPrize.FEEventCode = ievtCode
	clsEPrize.FEEventName = ievtName
	clsEPrize.FPSize = iPageSize
	clsEPrize.FCPage = iCurrpage
	If Not (sSearchUserid = "" AND ievtCode = "") Then
		arrList = clsEPrize.fnGetEventJoinList
	End IF
	iTotCnt = clsEPrize.FTotCnt
	
	
		iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
	
	Dim arrevtprizetype, arrevtprizestatus, arreventkind
	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	arrevtprizetype 	= fnSetCommonCodeArr("evtprizetype",False)	
	arrevtprizestatus 	= fnSetCommonCodeArr("evtprizestatus",False)	
	arreventkind		= fnSetCommonCodeArr("eventkind",False)	
%>
<script language="javascript">
<!--
	function jsGoPage(iP){
		document.frm.iC.value = iP;
		document.frm.submit();
	}
  
//-->
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="iC">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<table cellpadding="0" cellspacing="0" border="0" class="a">
			<tr>
				<td>
					<input type="radio" name="egubun" value="e"  onClick="frm.submit();" <%=CHKIIF(vEGubun="e","checked","")%>>�Ϲ��̺�Ʈ&nbsp;
					<input type="radio" name="egubun" value="f"  onClick="frm.submit();" <%=CHKIIF(vEGubun="f","checked","")%>>�������ΰŽ�&nbsp;
					<input type="radio" name="egubun" value="c"  onClick="frm.submit();" <%=CHKIIF(vEGubun="c","checked","")%>>Culture Station&nbsp;
				</td>
			</tr>
			<tr height="5"><td></td></tr>
			<tr>
				<td>
					&nbsp;���̵� : <input type="text" size="16" maxlength="32" name="searchUserid" value="<%=sSearchUserid%>">
					&nbsp;�̺�Ʈ �ڵ� : <input type="text" name="evtcode" value="<%=ievtCode%>" size="7" />
					&nbsp;������ : 
					<input type="text" name="date1" size="10" maxlength=10 readonly value="<%= vDate1 %>">
					<a href="javascript:calendarOpen(frm.date1);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
					&nbsp;~&nbsp;
					<input type="text" name="date2" size="10" maxlength=10 readonly value="<%= vDate2 %>">
					<a href="javascript:calendarOpen(frm.date2);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
					<!--&nbsp;�̺�Ʈ�� : <input type="text" name="evtname" value="<%=ievtName%>" size="30">//-->
				</td>
			</tr>
			</table>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
</table>	
<br>
�� <b>���̵�</b>�� <b>�̺�Ʈ�ڵ�</b> �� <b>�ϳ��� �ݵ�� �Է�</b>�ؾ� ����Ʈ�� ��Ÿ���ϴ�.(�˻��� ���� ��������)</b><Br><Br>
��<font size=4 color="red"> <b>[ON]�̺�Ʈ����>>[����]��������ƮNEW �� ������ֽñ� �ٶ��ϴ�. �� �������� ������ �Դϴ�.</b></font>
<br><br>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%=iTotCnt%></b>
			&nbsp;
			������ : <b><%=iCurrpage%>/<%=iTotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">�̺�Ʈ����</td> 
		<td width="60">�̺�Ʈ�ڵ�</td>       	
    	<td>�̺�Ʈ��</td>
    	<td width="100">����ID</td>
    	<td width="150">������</td>
      	<td width="100">��÷�ڹ�ǥ��</td>
      	<td width="100">����</td>
    </tr>   
    <%IF isArray(arrList) THEN
    	For intLoop = 0 To UBound(arrList,2)
    	
			If arrList(5,intLoop) = "Y" Then
				vImg = "yes"
			Else
				vImg = "no"
			End If
    	%> 
     <tr align="center" height="25" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
     	<td>		
     		<%
    			If vEGubun = "e" Then
    				rw fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))
    			ElseIf vEGubun = "f" Then
    				rw "�������ΰŽ�"
    			ElseIf vEGubun = "c" Then
    				rw "Culture Station(" & fnGetCommCodeArrDescCulture(arrList(1,intLoop)) & ")"
    			End If
     		%>
     	</td>
     	<td><%IF vEGubun = "f" Then%>
     			<a href="http://www.10x10.co.kr/designfingers/designfingers.asp?fingerid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a>
     		<%ELSEIF vEGubun = "c" Then%>
     			<a href="http://www.10x10.co.kr/culturestation/culturestation_event.asp?evt_code=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a>
     		<%ELSE%>
     			<a href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a>
     		<%END IF%>
     	</td>
    	<td align="left"><%=arrList(2,intLoop)%></td>
    	<td><%=arrList(7,intLoop)%></td>
    	<td><%=arrList(6,intLoop)%></td>
      	<td><%=Replace(formatdate(arrList(4,intLoop),"0000-00-00"),"1900-01-01","&nbsp;")%></td>
      	<td><img src="http://fiximage.10x10.co.kr/web2009/category/view_qna_<%=vImg%>_01.gif" style="margin-bottom:3px;"></td>
    </tr>   
	<% Next
    ELSE%>
     <tr align="center" bgcolor="#FFFFFF">
    	<td colspan="8" align="center">��ϵ� ������ �����ϴ�.</td>
    </tr>	
	<%END IF%>
</table>   
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
<%set clsEPrize = nothing%>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" --> 