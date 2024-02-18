<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/pop_event_lastlist.asp
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
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	'��������
	Dim menupos
Dim cEvtList, cDisp, vCateCode, i, eventstate, startdate, enddate
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, vOpenerForm
	Dim sEvt,strTxt,sKind
	
	startdate = Request("startdate")
	enddate = Request("enddate")
	eventstate = Request("eventstate")
	vCateCode = Request("catecode")
	menupos = request("menupos")
	sKind 	= request("eventkind")
	sEvt 	= request("selEvt")
	strTxt 	= request("sEtxt")
	vOpenerForm = request("openerform")
	
	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrpage = Request("iC")	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF
	if sEvt = "" then sEvt = "evt_code"
	if sKind = "" then sKind = "1"
	
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
	'������ ��������
	set cEvtList = new ClsEvent	
		cEvtList.FCPage = iCurrpage	'����������
		cEvtList.FPSize = iPageSize '���������� ���̴� ���ڵ尹�� 
		
		cEvtList.FSKind = sKind '�˻� ����
		cEvtList.FSfEvt = sEvt '�˻� �̺�Ʈ�� or �̺�Ʈ�ڵ�
		cEvtList.FSeTxt = strTxt '�˻���
		cEvtList.FRectState = eventstate '���°�
		cEvtList.FRectSDate = startdate '������
		cEvtList.FRectEDate = enddate '������
		cEvtList.FRectDisp = vCateCode
		
 		arrList = cEvtList.fnGetEventLastList	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotCnt	'��ü ������  ��
 	set cEvtList = nothing
 	 	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
	Dim arreventkind, arreventstate
	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	arreventkind = fnSetCommonCodeArr("eventkind",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)	
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--
	window.onload = function(){
		window.resizeTo(1000, 820);
	}

	//�˻�
	function jsSearch(){
		if(document.frmLast.sEtxt.value != ""){
			if(document.frmLast.selEvt.options[document.frmLast.selEvt.selectedIndex].value == "evt_code") {
				if(!IsDigit(document.frmLast.sEtxt.value)){
					alert("�̺�Ʈ �ڵ�� ���ڸ� �Է°����մϴ�.");
					document.frmLast.sEtxt.focus();
					return false;
				}
			}
		}
	}
	
	//�θ�â�� �� �ѱ��
	function jsSetEvtCont(ieC){
	  if(typeof(opener.document) == "object"){		 
	     <% if (request("pTarget")<>"") then %>
	     opener.location.href = "<%= request("pTarget") %>&eC="+ieC;
	     <% else %>
	     	<% If vOpenerForm <> "" Then %>
	     		opener.<%=vOpenerForm%>.value = ieC;
	    	<% Else %>
		 		opener.location.href = "event_regist.asp?menupos=<%=menupos%>&eC="+ieC;
		 	<% End If %>
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
<form name="frmLast" method="get" action="pop_event_lastlist.asp" onSubmit="return jsSearch();" style="margin:0px;">
<input type="hidden" name="iC">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="pTarget" value="<%=request("pTarget")%>">
<input type="hidden" name="openerform" value="<%=vOpenerForm%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="5%" height="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td height="30" align="left">����ī�װ� :
		<%
		SET cDisp = New cDispCate
		cDisp.FCurrPage = 1
		cDisp.FPageSize = 2000
		cDisp.FRectDepth = 1
		cDisp.FRectUseYN = "Y"
		cDisp.GetDispCateList()
		
		If cDisp.FResultCount > 0 Then
			Response.Write "<select name=""catecode"" class=""select"" onChange=""document.frmLast.submit();"">" & vbCrLf
			Response.Write "<option value="""">����</option>" & vbCrLf
			For i=0 To cDisp.FResultCount-1
				Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
			Next
			Response.Write "</select>&nbsp;&nbsp;&nbsp;"
		End If
		Set cDisp = Nothing
		%>
	</td>
	<td height="30" rowspan="3"><input type="submit" value=" ��  �� " class="button" style="width:70px;height:50px;" onfocus="this.blur();"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="30">
		�̺�Ʈ���� : <%sbGetOptEventCodeValue "eventkind",sKind,True,"onChange=""document.frmLast.submit();"""%>&nbsp;&nbsp;&nbsp;
		�ڵ�/�� : <select name="selEvt">
        			<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
        			<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>�̺�Ʈ��</option>
        			</select>
        			<input type="text" name="sEtxt" size="15" value="<%=strTxt%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="30">
		������� : <% sbGetOptStatusCodeValue "eventstate",eventstate,true,"class=""select"" onChange=""document.frmLast.submit();""" %>&nbsp;&nbsp;&nbsp;
		�Ⱓ :&nbsp;
        <input id="startdate" type="text" name="startdate" value="<%= startdate %>" maxlength="10" size="10">
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
        <input type="text" name="dummy0" value="00:00:00" size="8" readonly class="text_ro" />
	    <script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "startdate",
			trigger    : "startdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
		&nbsp;~&nbsp;
        <input id="enddate" type="text" name="enddate" value="<%= enddate %>" maxlength="10" size="10">
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
        <input type="text" name="dummy1" value="23:59:59" size="8" readonly class="text_ro" />
	    <script type="text/javascript">
		var CAL_End = new Calendar({
			inputField : "enddate",
			trigger    : "enddate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			},
			bottomBar: true,
			dateFormat: "%Y-%m-%d"
		});
		</script>
	</td>
</tr>
</table>
<br />
<table width="100%" border="0" align="left" class="a" cellpadding="0" cellspacing="0" >
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" >
			<td align="center" width="15%">����</td>
			<td align="center" width="7%">�ڵ�</td>
			<td align="center" width="15%">�������</td>
			<td align="center">�̺�Ʈ��</td>
			<td align="center" width="15%">ī�װ�</td>
			<td align="center" width="10%">������</td>
			<td align="center" width="10%">������</td>
		</tr>
		<%IF isArray(arrList) THEN 
			For intLoop = 0 To UBound(arrList,2)
			%>
		<tr bgcolor="#FFFFFF" onClick="jsSetEvtCont(<%=arrList(0,intLoop)%>);" style="cursor:hand;" onMouseOver="this.style.backgroundColor='#DDDDDD'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
			<td height="25" align="center"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></td>
			<td  align="center"><%=arrList(0,intLoop)%></td>
			<td  align="center"><%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%></td>
			<td><%=db2html(arrList(4,intLoop))%></td>
			<td  align="center"><%=arrList(9,intLoop)%></td>
			<td  align="center"><%=arrList(5,intLoop)%></td>
			<td  align="center"><%=arrList(6,intLoop)%></td>
		</tr>
		<% Next %>
		<%ELSE%>
		<tr><td colspan="10"  bgcolor="#FFFFFF" align="center">��ϵ� ������ �����ϴ�.</td></tr>
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
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->