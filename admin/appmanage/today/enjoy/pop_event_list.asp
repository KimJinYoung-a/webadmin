<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/pop_event_list.asp
' Description :  ���� �̺�Ʈ ���� ��������
' History : 2007.03.20 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
	'��������
	Dim menupos
	Dim cEvtList	
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
	Dim sEvt,strTxt,sKind , sDt , eDt
	dim contenType
	dim idx	

	idx = request("idx")
	sDt = request("sDt")
	eDt = request("eDt")
	menupos = request("menupos")
	sKind 	= request("eventkind")
	sEvt 	= request.form("selEvt")
	strTxt 	= request.form("sEtxt")
	contenType	= request("type")

	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrpage = Request("iC")	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
		
	'������ ��������
	set cEvtList = new ClsEvent	
		cEvtList.FCPage = iCurrpage	'����������
		cEvtList.FPSize = iPageSize '���������� ���̴� ���ڵ尹�� 
		
		cEvtList.FSKind = sKind '�˻� ����
		cEvtList.FSfEvt = sEvt '�˻� �̺�Ʈ�� or �̺�Ʈ�ڵ�
		cEvtList.FSeTxt = strTxt '�˻���	
		
 		arrList = cEvtList.fnGetEventLastList	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotCnt	'��ü ������  ��
 	set cEvtList = nothing
 	 	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
	Dim arreventkind, arreventstate
	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	arreventkind = fnSetCommonCodeArr("eventkind",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)	
%>
<script language="javascript">
<!--
	//�˻�
	function jsSearch(){	
	 if(document.frmLast.selEvt.options[document.frmLast.selEvt.selectedIndex].value == "evt_code") {
	   if(!IsDigit(document.frmLast.sEtxt.value)){
	    alert("�̺�Ʈ �ڵ�� ���ڸ� �Է°����մϴ�.");
	    document.frmLast.sEtxt.focus();
	    return false;
	   }
	 }		
	}
	
	//�θ�â�� �� �ѱ��
	function jsSetEvtCont(ieC){
	  if(typeof(opener.document) == "object"){		 
	     <% if (request("pTarget")<>"") then %>
	     opener.location.href = "<%= request("pTarget") %>&eC="+ieC;
		 <% else %>
		 	<% if contenType = "3" then %>
		 		opener.location.href = "mainTopExhibition_insert.asp?idx=<%=idx%>&menupos=<%=menupos%>&sDt=<%=sDt%>&eDt=<%=eDt%>&linktype=1&eC="+ieC;
			 <% else %>
			 	opener.location.href = "enjoy_insert.asp?menupos=<%=menupos%>&sDt=<%=sDt%>&eDt=<%=eDt%>&linktype=1&eC="+ieC;
			 <% end if %>
		 <% end if %>
		 window.close();
	  }	
	}
	
	//�������̵�
	function jsGoPage(iP){
		document.frmLast.iC.value = iP;
		document.frmLast.submit();
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ���� �̺�Ʈ ����Ʈ </div>
<table width="500" border="0" align="left" class="a" cellpadding="3" cellspacing="0" >
<form name="frmLast" method="post" action="pop_event_list.asp" onSubmit="return jsSearch();">
<input type="hidden" name="iC">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="pTarget" value="<%=request("pTarget")%>">
<input type="hidden" name="sDt" value="<%=request("sDt")%>">
<input type="hidden" name="eDt" value="<%=request("eDt")%>">
<input type="hidden" name="type" value="<%=request("type")%>">
<input type="hidden" name="idx" value="<%=request("idx")%>">
<tr>
	<td>
		����: <%sbGetOptEventCodeValue "eventkind",sKind,True,"onChange=""document.frmLast.submit();"""%>&nbsp;&nbsp;&nbsp;
		�ڵ�/��: <select name="selEvt">
        			<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
        			<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>�̺�Ʈ��</option>
        			</select>
        			<input type="text" name="sEtxt" size="15" value="<%=strTxt%>">
        			 <input type="image" src="/images/icon_search.jpg">
    </td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" >
			<td align="center" width="10%">�ڵ�</td>	
			<td align="center" width="15%">����</td>
			<td align="center">�̺�Ʈ��</td>	
			<td align="center" width="15%">����</td>	
		</tr>
		<%IF isArray(arrList) THEN 
			For intLoop = 0 To UBound(arrList,2)
			%>
		<tr bgcolor="#FFFFFF" onClick="jsSetEvtCont(<%=arrList(0,intLoop)%>);" style="cursor:hand;" onMouseOver="this.style.backgroundColor='#FFFFEC'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
			<td  align="center"><%=arrList(0,intLoop)%></td>
		
			<td  align="center"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></td>
	
			<td><%=db2html(arrList(4,intLoop))%></td>
	
			<td  align="center"><%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%></td>
		</tr>
		<% Next %>
		<%ELSE%>
		<tr><td colspan="4"  bgcolor="#FFFFFF" align="center">��ϵ� ������ �����ϴ�.</td></tr>
		<%END IF%>
		</table>	
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
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
</table>
	</td>
	
</tr>
</form>
</table>


<!-- #include virtual="/lib/db/dbclose.asp" -->