<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/index.asp
' Description :  �̺�Ʈ ��� - ȭ�鼳��
' History : 2007.02.07 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->
<%
	Call fnSetEventCommonCode '�����ڵ� ���ø����̼� ������ ����

	'��������
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
	Dim strparm

	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrpage = Request("iC")	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����

	'## �˻� #############################
	sDate = Request("selDate")  '�Ⱓ
	sSdate = Request("iSD")
	sEdate = Request("iED")

	sEvt = Request("selEvt")  '�̺�Ʈ �ڵ�/�� �˻�
	strTxt = Request("sEtxt")

	sCategory	= Request("selC") 'ī�װ�
	sState	 = Request("eventstate")'�̺�Ʈ ����
	sKind = Request("eventkind")	'�̺�Ʈ����

	strparm = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&eventstate="&sState&"&eventkind="&sKind
	'#######################################

	'������ ��������
	set cEvtList = new ClsEvent
		cEvtList.FCPage = iCurrpage	'����������
		cEvtList.FPSize = iPageSize '���������� ���̴� ���ڵ尹��

		cEvtList.FSfDate = sDate '�Ⱓ �˻� ����
		cEvtList.FSsDate = sSdate '�˻� ������
		cEvtList.FSeDate = sEdate '�˻� ������
		cEvtList.FSfEvt = sEvt '�˻� �̺�Ʈ�� or �̺�Ʈ�ڵ�
		cEvtList.FSeTxt = strTxt '�˻���
		cEvtList.FScategory = sCategory '�˻� ī�װ�
		cEvtList.FSstate = sState '�˻� ����
		cEvtList.FSkind = sKind

 		arrList = cEvtList.fnGetEventList	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotCnt	'��ü ������  ��
 	set cEvtList = nothing

	iTotalPage 	=  Int(iTotCnt/iPageSize)	'��ü ������ ��
	IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1
%>
<script language="javascript">
<!--
	function jsGoPage(iP){
		document.frmEvt.iC.value = iP;
		document.frmEvt.submit();
	}

	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}

	function jsSearch(frm, sType){
	if (sType == "A"){
			frm.iSD.value = "";
			frm.iED.value = "";
			frm.eventstate.value = "";
			frm.sEtxt.value = "";
			frm.selC.value = "";
		}

		frm.submit();
	}

	function jsSchedule(){
		var winS;
		winS = window.open('pop_event_schedule.asp','popwin','width=800, height=600, scrollbars=yes');
		winS.focus();
	}

	function jsChSelect(iVal){
		alert(iVal);
		alert(document.frmEvt.eventkind.value);
		alert(document.frmEvt.eventkind.options[document.frmEvt.eventkind.selectedIndex].value);
		document.frmEvt.submit();
	}
	function jsStatistics(iVal){
		window.open('pop_evt_Statistics.asp?eC='+ iVal,'sPop','');
	}
	function jsSettingEvt(iVal){
		window.open('pop_evt_Setting.asp?eC='+ iVal,'sPop','');
	}
//-->
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frmEvt" method="get"  action="index.asp" onSubmit="return jsSearch(this,'E');">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="iC">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2">
        	<table border="0"  cellpadding="1" cellspacing="3" class="a">
        	<tr>
        		<td width="30" align="right"></td>
        		<td colspan="2">
        			<!-- �̺�Ʈ ���� -->
        			<%sbGetOptEventCodeValue "eventkind", sKind, True,"onChange='javascript:document.frmEvt.submit();'"%>
        			<!-- ī�װ� -->
        			<% sbGetOptCategoryLarge "selC", sCategory ,"onChange='javascript:document.frmEvt.submit();'" %>
        			<!-- �̺�Ʈ ���� -->
        			<%sbGetOptEventCodeStateValue "eventstate", sState, True,"onChange='javascript:document.frmEvt.submit();'"%>
        		</td>
        	</tr>
        	<tr>
        		<td width="30" align="right"></td>
        		<td><select name="selEvt">
        			<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
        			<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>�̺�Ʈ��</option>
        			</select>
        			<input type="text" name="sEtxt" value="<%=strTxt%>">
        		&nbsp;&nbsp;<br>�Ⱓ:
        	 	 <select name="selDate">
        			<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
        			<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
        		 </select>
        		<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:pointer;">
        		 ~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:pointer;">&nbsp;&nbsp;
        		</td>
        		<td  colspan="2" align="right" valign="bottom">&nbsp;&nbsp;
        			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
        			<input type="button" class="button" value="��ü����" onClick="jsSearch(document.frmEvt, 'A')">
        		</td>
        	</tr>
        	</table>
        </td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="35">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
       	<img src="/images/icon_new_registration.gif" onclick="jsGoUrl('event_regist.asp?menupos=<%=menupos%>');" style="cursor:pointer;">
    	</td>
    	<td align="right">
       	<input type="button" class="button" value="������" onclick="jsSchedule();">
       <!--	<input type="button" value="���" onclick=" ">  -->
       <!--	����: <select name="selSort">
       	<option value="1">�̺�Ʈ�ڵ峻������</option>

       	</select>-->
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">�̺�Ʈ�ڵ�</td>
    	<td width="80">�������</td>
      	<td width="100">����</td>
      	<td>�̺�Ʈ��</td>
      	<td width="80">ī�װ�</td>
      	<td width="70">������</td>
      	<td width="70">������</td>
      	<td width="110">���</td>
    </tr>
    <%IF isArray(arrList) THEN
    	For intLoop = 0 To UBound(arrList,2)
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><a href="<%=wwwUrl%>/event/eventmain.asp?eventid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a></td>
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetEventCodeDesc("eventstate",arrList(8,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetEventCodeDesc("eventkind",arrList(1,intLoop))%></a></td>
      	<td><a href="event_entryList.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=db2html(arrList(4,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(12,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(5,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(6,intLoop)%></a></td>
      	<td align="left"><input type="button" value="ȭ��" class="input_b" onClick="javascript:jsGoUrl('event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')">
      		<input type="button" value="��ǰ" class="input_b" onClick="javascript:jsGoUrl('eventitem_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')">
      		<%IF arrList(13,intLoop) > "1900-01-01" THEN%><input type="button" value="��÷" class="input_b" onClick="jsStatistics('<%=arrList(0,intLoop)%>')"><input type="button" class="input_b" value="����ۼ�" onclick="jsSettingEvt('<%=arrList(0,intLoop)%>');"><%END IF%>

      	</td>
    </tr>
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="8">��ϵ� ������ �����ϴ�.</td>
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
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->