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
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
	Call fnSetEventCommonCode '�����ڵ� ���ø����̼� ������ ����
	
	'��������
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
	Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
	
	Dim iIsSale,iIsGift,iIsCoupon,fstchk
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
	
	if sSdate="" then sSdate= dateserial(year(now()),month(now()),day(now()))
	if sEdate="" then sEdate = dateserial(year(now()),month(now()),day(now()))
	sEvt = Request("selEvt")  '�̺�Ʈ �ڵ�/�� �˻�
	strTxt = Request("sEtxt")
	
	sCategory	= Request("selC") 'ī�װ�
	sState	 = Request("eventstate")'�̺�Ʈ ����	
	sKind = Request("eventkind")	'�̺�Ʈ����
	
	if sState ="" then sState="7"
	
	fstchk = request("fstchk")	
	iIsSale = request("iIsSale")
	iIsGift = request("iIsGift")
	iIsCoupon = request("iIsCoupon")
	
	if fstchk="" then 
		
		if iIsGift="" then 
			iIsGift="on"		
		end if 	
	end if
	
	

		
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
		
		cEvtList.FSisSale = iIsSale
		cEvtList.FSisGift = iIsGift
		cEvtList.FSisCoupon = iIsCoupon
		
 		arrList = cEvtList.fnGetEventList_LOG	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotCnt	'��ü ������  ��
 	set cEvtList = nothing
 	
	iTotalPage 	=  Int(iTotCnt/iPageSize)	'��ü ������ ��
	IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1		
		
	function setChkStr(st)
		if st="on" then setChkStr="checked"
	end function
	
	function getGiftItems(evtcode)
		dim sql
		sql =" select top 10 gift_itemname " &_
				" from db_event.dbo.tbl_gift " &_
				" where evt_code='"&evtcode&"'"
				
		rsget.open sql,dbget,1
		
		if not rsget.eof then
			response.write rsget("gift_itemname")
		end if
		rsget.close
	end function

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
//-->

function DrawReceiptPrintobj_TEC(elementid,printname){
        var objstring = "";
        var e;
        objstring = '<OBJECT name="' + elementid + '" ';
        objstring = objstring + ' classid="clsid:E76C9051-A8C4-458E-9F60-3C14DB9EECF9" ';
        objstring = objstring + ' codebase="http://billyman/Tec_dol.cab#version=1,5,0,0" ';
        objstring = objstring + ' width=0 ';
        objstring = objstring + ' height=0 ';
        objstring = objstring + ' align=center ';
        objstring = objstring + ' hspace=0 ';
        objstring = objstring + ' vspace=0 ';
        objstring = objstring + ' > ';
        objstring = objstring + ' <PARAM Name="PrinterName" Value="' + printname + '"> ';
        objstring = objstring + ' </OBJECT>';
        
        document.write(objstring);
}

function eventindexprint(ievt_code, ievt_name01, ievt_name02, ievt_startdate, ievt_enddate, ievt_gift01, ievt_gift02, ievt_gift03){
	var X = 1;
	var Y = 1;
	var F = 1;
	
	// TEC_DO3 : 452
	if (TEC_DO3.IsDriver == 1){
           X = 1.05;
           Y = 1.05;
           F = 1.2;
           
			TEC_DO3.SetPaper(900,600);
			TEC_DO3.OffsetX = 20;
			TEC_DO3.OffsetY = 20;
			TEC_DO3.PrinterOpen();
			
			
			
			

			TEC_DO3.PrintText(50*X, 50*Y, "HY�߰��", 30*F, 0, 0, "[������]");
			TEC_DO3.PrintText(250*X, 50*Y, "HY�߰��", 30*F, 1, 0, ievt_startdate);
			
			TEC_DO3.PrintText(50*X, 100*Y, "HY�߰��", 30*F, 0, 0, "[������]");
			TEC_DO3.PrintText(250*X, 100*Y, "HY�߰��", 30*F, 1, 0, ievt_enddate);
			
			TEC_DO3.PrintText(500*X, 30*Y, "Arial Bold", 150*F, 0, 0, ievt_code);
			
			TEC_DO3.PrintText(50*X, 175*Y, "HY�߰��", 30*F, 0, 0, "[�̺�Ʈ��]");
			TEC_DO3.PrintText(50*X, 225*Y, "HY�߰��", 30*F, 1, 0, ievt_name01);
			TEC_DO3.PrintText(50*X, 275*Y, "HY�߰��", 30*F, 1, 0, ievt_name02);
			
			TEC_DO3.PrintText(50*X, 350*Y, "HY�߰��", 30*F, 0, 0, "[����ǰ��]");
			TEC_DO3.PrintText(50*X, 400*Y, "HY�߰��", 30*F, 1, 0, ievt_gift01);	<!-- ���ٿ� �ѱ� 24�ڱ��� ������ �Ʒ��� -->
			TEC_DO3.PrintText(50*X, 450*Y, "HY�߰��", 30*F, 1, 0, ievt_gift02);
			TEC_DO3.PrintText(50*X, 500*Y, "HY�߰��", 30*F, 1, 0, ievt_gift03);
			
			TEC_DO3.PrinterClose();


    }else window.status = "TEC B-452 ����̹��� ��ġ�� �ּ���"
}

DrawReceiptPrintobj_TEC("TEC_DO3","TEC B-452");


</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmEvt" method="get"  action="" onSubmit="return jsSearch(this,'E');">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="fstchk" value="on">
	<input type="hidden" name="iC">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�̺�Ʈ����:<%sbGetOptEventCodeValue "eventkind", sKind, True,"onChange='javascript:document.frmEvt.submit();'"%>
			&nbsp;
			ī�װ�:<% sbGetOptCategoryLarge "selC", sCategory ,"onChange='javascript:document.frmEvt.submit();'" %>
			&nbsp;
			�������:<%sbGetOptEventCodeValue "eventstate", sState, True,"onChange='javascript:document.frmEvt.submit();'"%>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�̺�ƮŸ��:
			<input type="checkbox" name="iIsSale" <%= setChkStr(iIsSale) %>>����
			<input type="checkbox" name="iIsGift" <%= setChkStr(iIsGift) %>>����ǰ
			<input type="checkbox" name="iIsCoupon" <%= setChkStr(iIsCoupon) %>>����
			&nbsp;
			�ڵ�/��:
			<select class="select" name="selEvt">
    			<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
    			<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>�̺�Ʈ��</option>
			</select>
			<input type="text" class="text" name="sEtxt" value="<%=strTxt%>">
			&nbsp;
			�Ⱓ:
    	 	<!--
    	 	<select name="selDate">        	 	 	
    			<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
    			<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
    		</select>
    		-->        		
    		<input type="text" class="text" size="11" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:pointer;">
    		 ~ <input type="text" class="text" size="11" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:pointer;">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>
<input type="button" value="��ü����" onClick="jsSearch(document.frmEvt, 'A')">

<p>
<!-- ǥ �߰��� ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">       	
			�߰��˻����� : ������������ �̺�Ʈ(�������� �����ϰ� ������ ���̿� �־�� �ϸ�, �������� �ȵ��̺�Ʈ)<br>
			������¿� �ϳ� �� �߰� --> ��ǰ�Ϸ�(����ǰ)
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</table>
<!-- ǥ �߰��� �� -->

<!-- ǥ �߰��� ����
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">       	
       	<img src="/images/icon_new_registration.gif" onclick="jsGoUrl('event_regist.asp?menupos=<%=menupos%>');" style="cursor:hand;">     
    	</td>
    	<td align="right">
       	<input type="button" value="������" onclick="jsSchedule();">       	
       <!--	<input type="button" value="���" onclick=" ">  -->
       <!--	����: <select name="selSort">
       	<option value="1">�̺�Ʈ�ڵ峻������</option>
       	
       	</select>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</table>
ǥ �߰��� �� -->
 
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">�̺�Ʈ<br>�ڵ�</td>
      	<td width="90">�������</td>
      	<td width="100">����</td>
      	<!--<td width="60">Ÿ��</td>-->
      	<td>�̺�Ʈ��</td>      	
      	<td width="100">����̹���</td>   
      	<td width="120">ī�װ�</td>
      	<td width="60">������</td>
      	<td width="60">������</td>
      	<td width="80">����</td>
      	<td >����ǰ </td>
      	<td width="60">�ε������</td>
    </tr>
    <%IF isArray(arrList) THEN 
    	For intLoop = 0 To UBound(arrList,2)
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><a href="<%=vwwwUrl%>/event/eventmain.asp?eventid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetEventCodeDesc("eventstate",arrList(8,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetEventCodeDesc("eventkind",arrList(1,intLoop))%></a></td>
      	<!--<td></td>-->
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=db2html(arrList(4,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%IF arrList(10,intLoop) <> "" THEN%> <img src="<%=arrList(10,intLoop)%>" width="100" border="0"><%END IF%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(12,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(5,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(6,intLoop)%></a></td>
      	<td><input type="button" value="ȭ��" class="button" onClick="javascript:jsGoUrl('event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')">
      		<input type="button" value="��ǰ" class="button" onClick="javascript:jsGoUrl('eventitem_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')">
      		<%IF arrList(13,intLoop) > "1900-01-01" THEN%><input type="button" value="��÷" class="input_b" onClick="jsGoUrl('eventprize_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')"><%END IF%>
      	</td>
    	<td ><%= getGiftItems(arrList(0,intLoop)) %> </td>
    	<td><input type="button" value="���" class="button" onClick="eventindexprint('<%=arrList(0,intLoop)%>', '<%= left(db2html(arrList(4,intLoop)),24) %>','<%= mid(db2html(arrList(4,intLoop)),25) %>','<%=arrList(5,intLoop)%>','<%=arrList(6,intLoop)%>','<%= left(getGiftItems(arrList(0,intLoop)),24) %>','<%= mid(getGiftItems(arrList(0,intLoop)),25,24) %>','<%= mid(getGiftItems(arrList(0,intLoop)),49) %>')"></td>
    </tr>   
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="11">��ϵ� ������ �����ϴ�.</td>
   	</tr>	
   <%END IF%>
   
    <!-- ����¡ó�� -->
    <%		
	iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	
	
	If (iCurrpage mod iPerCnt) = 0 Then																
		iEndPage = iCurrpage
	Else								
		iEndPage = iStartPage + (iPerCnt-1)
	End If	
	%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
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


<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" --> -->