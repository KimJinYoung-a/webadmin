<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	PageName 	: /admin/hitchhiker/index.asp
'	Description : ��ġ����Ŀ ��ûȸ������Ʈ �ٿ�� �߼�Ȯ��
'	History		: 2006.11.30 ������ ����
'                 2012.02.13 ������ - �̴ϴ޷� ��ü
'				  2016.07.19 �ѿ�� ���� SSL ����
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhikerCls.asp"-->
<%
Dim clsHList, arrHVol, intHV, iHVol, arrAVol, intA , iAVol, blnSend, searchTxt, search
Dim arrHList, intH, iPageSize, iCurrpage, iTotCnt, iStartPage, iEndPage, iTotalPage, ix,iPerCnt
Dim chkList,chkView, startDate, endDate
	iHVol = Request("iHV")
	iAVol =	 Request("iAV")
	startDate	= Request("startDate")
	endDate	= Request("endDate")
	iCurrpage = Request("iC")	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	searchTxt = Request("searchTxt")
	search= Request("search")
	chkList = Request("chkList")
	IF chkList ="view" THEN
		chkView  = ""
	ELSE
		chkView  = "none"
	END IF

set clsHList =  new Chitchhiker
	arrHVol	= clsHList.fnGetHVol	'1.����ȸ�� ��������
	IF iHVol = "" THEN
		IF isArray(arrHVol) THEN
		iHVol	= arrHVol(0,0)
		END IF
	END IF

	clsHList.FHVol = iHVol			'Set ����ȸ��
	arrAVol = clsHList.fnGetApplyVol	'2.��ûȸ�� ��������

	IF iAVol = "" and blnSend ="" THEN
		IF isArray(arrAVol) THEN
			iAVol = arrAVol(0,0)
		END IF
	END IF

	clsHList.FAVol = iAVol			'Set ��ûȸ��
	clsHList.FisSend = blnSend		'Set �߼ۿ���

	IF chkList = "view" THEN
	clsHList.FPSize = iPageSize		'Set ������ ������
	clsHList.FCPage = iCurrpage		'Set ���� ������ ��ȣ
	clsHList.FSearch = search
	clsHList.FSearchTxt = searchTxt	'Set �˻���
	clsHList.FSDate = startDate		'�˻�������
	clsHList.FEDate = endDate		'�˻�������
	arrHList = clsHList.fnGetList	'3.��û ����Ʈ ��������
	iTotCnt = clsHList.FTotCnt 		'��û����Ʈ �� ���� ��������
	ELSE
		arrHList = NULL
	END IF
set clsHList = nothing

'��ü ������ ��
iTotalPage 	=  Int(iTotCnt/iPageSize)
IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1

%>
<script type="text/javascript">
<!--
	function jsSetPage(frm, sVal){
		eval("frm."+sVal).value = "";
		frm.submit();
	}

	function jsGetList(frm){
		frm.chkList.value = "view";
		frm.submit();
	}


	function jsDownList(frm){
	    var iHV = frm.iHV.value;
	    var iAV = frm.iAV.value;
	    var iType = frm.blnS.value;
	    var sSdt = frm.startdate.value;
	    var sEdt = frm.enddate.value;
//alert(iType);
//return;
	    var popwin = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/downHitchhiker.asp?iHV=' + iHV + '&iAV='+iAV+'&iType='+iType+'&startDate='+sSdt+'&endDate='+sEdt,'downHitchhiker','width=400, height=300');
		popwin.focus();

	}

	function jsSend(frm){
		var smsyn='';
		smsyn = 'N';
		if (frm.smsyn.value=='Y'){
			if(confirm("���Բ� �߼�ó�� ���ڹ߼��� ���� �ϼ̽��ϴ�. �߼� �Ͻðڽ��ϱ�?") == true) {
				smsyn = 'Y';
			}else{
				return false;
			}
		}

		if(frm.blnS.value == 2){
			alert("�̹� �߼۵� ��� ����Ʈ�� ���ؼ��� �߼�Ȯ�� �Ұ����մϴ�.");
			return;
		}
		if(confirm("�߼�Ȯ�� ó���Ͻðڽ��ϱ�?")){
			frm.pMode.value = "C";
			frm.action ="<%= getSCMSSLURL %>/admin/hitchhiker/processHitchhiker.asp";
			frm.submit();
		}
	}

	function jsApply(iH){
		var winApply;
		winApply = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/registHitchhiker.asp?pMode=A&iHV='+iH,'popWin','width=400, height=300');
		winApply.focus();
	}

	//�Ķ��Ÿ�� �������� �ѱ��� ����. Ű ������ idx ���� �ѱ�.		2017.07.06 �ѿ��
	function jsReApply(iH, temp, iA, idx){
		var winApply;
		winApply = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/registHitchhiker.asp?pMode=R&iHV='+iH+'&idx='+idx+'&iAV='+iA,'popWin','width=400, height=300');
		winApply.focus();
	}

	function jsGoPage(iP){
		document.frmList.iC.value = iP;
		document.frmList.submit();
	}
	function jsLogView(iHVol,iAVol){
		var winApply;
		winApply = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/LogListHitchhiker.asp?iHV='+iHVol+'&iAV='+iAVol,'popWin','width=800, height=800');
		winApply.focus();
	}

	//�Ķ��Ÿ�� �������� �ѱ��� ����. Ű ������ idx ���� �ѱ�.		2017.07.06 �ѿ��
	function jsAddrUPdate(iH,temp,idx){
		var winAddr;
		winAddr = window.open('<%= getSCMSSLURL %>/admin/hitchhiker/updateHitchhikerAddr.asp?iHV='+iH+'&idx='+idx,'popWin','width=400, height=300');
		winAddr.focus();
	}

//-->
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!-- ǥ ��ܹ� ����-->
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1">
<tr>
	<td width="95"><font color="gray">+ ����Ʈ���� :</td>
	<td> <font color="gray">��ġ����Ŀ ����ȸ�� ���� - ��û�߼�ȸ�� �Ǵ� �߼ۿ��θ� ���� - ����Ʈ�����ư - �ش����ǿ� �ش��ϴ� ��û�� ����Ʈ Ȯ��</td>
</tr>

<tr>
	<td  width="95" valign="top"><font color="gray">+ �߼�Ȯ�� ó�� :</td>
	<td> <font color="gray">��ġ����Ŀ ����ȸ�� ���� - �߼ۿ��θ� �̹߼����� ���� - �߼�Ȯ��ó����ư - �̹߼�ó���ǿ� ���� �߼�Ȯ��ó��<br>
		 ��ġ����Ŀ ����ȸ�� ���� - ��û�߼�ȸ�� ���� - �߼�Ȯ��ó����ư - �߼ۿ��ο� ������� �߼�Ȯ�� ��ó��
	</td>
</tr>
<tr>
	<td width="95"><font color="gray">+ �߼۽�û :</td>
	<td><font color="gray">��ġ����Ŀ ����ȸ�� ���� - ���߼۽�û��ư - �ش� ����ȸ���� �̹߼� ȸ���� ��ûó��</td>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#F3F3FF" >
	<td>
		<input type = "button" class="button" onclick="javascript:location.href='<%= getSCMSSLURL %>/admin/eventmanage/hitchhiker/index.asp';" value="��ġ����Ŀ VIP �ּ��Է� ����">
	</td>
</tr>
</table>

<form name="frmH" method="post" action="index.asp" style="margin:0px;">
<input type="hidden" name="chkList" value="<%=chkList%>">
<input type="hidden" name="pMode" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#F3F3FF" >
		<td rowspan="2" width="200" align="center">
			��ġ����Ŀ ����ȸ�� Vol.
        	<select name="iHV" onChange="jsSetPage(document.frmH,'blnS');">
			<% IF isArray(arrHVol) THEN %>
				<%FOR intHV = 0 TO UBound(arrHVol,2) %>
			<option value="<%=arrHVol(0,intHV)%>" <%IF Cint(iHVol) = Cint(arrHVol(0,intHV)) THEN %> selected<%END IF%>><%=arrHVol(0,intHV)%></option>
				<%NEXT%>
			<% END IF %>
			</select>
		</td>
		<td width="180" align="right">
			��û�߼� ȸ��
		<select name="iAV" onChange="jsSetPage(document.frmH,'blnS');" style="width:75px;">
		<option value="">--����--</option>
		<% IF isArray(arrAVol) THEN %>
			<%FOR intA = 0 TO UBound(arrAVol,2) %>
		<option value="<%=arrAVol(0,intA)%>" <%IF CInt(iAVol) = CInt(arrAVol(0,intA)) THEN %> selected<%END IF%>><%=arrAVol(0,intA)%></option>
			<%NEXT%>
		<% END IF %>
		</select>
		&nbsp;&nbsp;
		</td>
		<td rowspan="2" align="center">
	        ������ <input id="startdate" name="startdate" value="<%=startdate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /><br>
	        ������ <input id="enddate" name="enddate" value="<%=enddate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "startdate", trigger    : "startdate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "enddate", trigger    : "enddate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
		</td>
		<td rowspan="2">&nbsp;&nbsp;
			<input type="button" value="����Ʈ����" class="a" onClick="jsGetList(document.frmH);">
		<% If C_ADMIN_AUTH or C_SYSTEM_Part or session("ssAdminPsn") = "14" or session("ssAdminPsn") = "23" or C_logics_Part or C_CriticInfoUserLV3 Then %>
			<input type="button" value="�����ٿ�" class="a" onClick="jsDownList(document.frmH);">
			<br />
			���ڹ߼� :
			<select name="smsyn" class="select">
				<option value="">-CHOICE-</option>
				<option value="Y">Y</option>
			</select>
			<input type="button" value="�߼�Ȯ��ó��" class="a" onClick="jsSend(document.frmH);">
			<br />
		<% End If %>
			<input type="button" value="���߼۽�û" class="a" onClick="jsApply(<%=iHVol%>);">
			<% If chkList = "view" Then %>
			<input type="button" value="��߼�Log����" class="a" onClick="jsLogView(<%=iHVol%>,<%=iAVol%>);">
			<% End If %>
		</td>

	</tr>
	<tr bgcolor="#F3F3FF" >
		<td align="right">
		�߼ۿ���
		<!--
		<select name="blnS" onChange="jsSetPage(document.frmH,'iAV');" class="a" style="width:75px;">
		-->
		<select name="blnS" class="a" style="width:75px;">
			<option value="">--����--</option>
			<option value="1" <%IF blnSend ="1" THEN%>selected<%END IF%>>�̹߼�</option>
			<option value="2" <%IF blnSend ="2" THEN%>selected<%END IF%>>�߼�</option>
		</select>
		&nbsp;&nbsp;
		</td>
	</tr>
</table>
</form>
<!-- ǥ ��ܹ� ��-->
<br>
<div id="hlist" style="display:<%=chkView%>;">
<form name="frmList" method="post" action="<%= getSCMSSLURL %>/admin/hitchhiker/index.asp">
<input type="hidden" name="iHV" value="<%=iHVol%>">
<input type="hidden" name="iAV" value="<%=iAVol%>">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="chkList" value="<%=chkList%>">
<input type="hidden" name="startdate" value="<%= startdate %>">
<input type="hidden" name="enddate" value="<%= enddate %>">
<!---- �˻�------------->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="30" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<select name="search">
        		<option value="">-����-</option>
        		<option value="userid" <% If Search = "userid" Then response.write "selected" End If %>>���̵�</option>
        		<option value="username" <% If Search = "username" Then response.write "selected" End If %>>�̸�</option>
        		<option value="receviename" <% If Search = "receviename" Then response.write "selected" End If %>>������</option>
        	</select>
			<input type="text" name="searchTxt" maxlength="32" value="<%=searchTxt%>">
        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!---- /�˻�------------->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ȣ</td>
	<td>���̵�</td>
	<td>�̸�</td>

	<% if (FALSE) then %>
	<td>������</td>
	<td>�����ȣ</td>
	<td>�ּ�</td>
	<td>��ȭ��ȣ</td>
	<td>�ڵ���</td>
	<% end if %>

	<td>��û��</td>
	<td>�߼���</td>
	<td>��߼�</td>
</tr>
<%IF isArray(arrHList) THEN %>
	<%FOR intH =0 TO UBound(arrHList,2)%>
	<tr align="center" bgcolor="ffffff">
	<td><%=iTotCnt-intH-(iPageSize*(iCurrpage-1))%></td>
	<td><%= printUserId(arrHList(3,intH), 2, "*") %></td>

	<% IF isNull(arrHList(10,intH)) THEN %>
		<td onclick="jsAddrUPdate('<%=arrHList(0,intH)%>','','<%=arrHList(13,intH)%>');" style="cursor:pointer"><%=arrHList(4,intH)%></td>
	<% else %>
		<td ><%=arrHList(4,intH)%></td>
	<% end if %>

	<% if (FALSE) then %>
	<td><%=arrHList(12,intH)%></td>
	<td><%=arrHList(5,intH)%></td>

	<% IF isNull(arrHList(10,intH)) THEN %>
		<td align="left" onclick="jsAddrUPdate('<%=arrHList(0,intH)%>','','<%=arrHList(13,intH)%>');" style="cursor:pointer"><%=arrHList(6,intH)%>&nbsp;<%=db2html(arrHList(7,intH))%></td>
	<% Else %>
		<td align="left"><%=arrHList(6,intH)%>&nbsp;<%=db2html(arrHList(7,intH))%></td>
	<% End If %>

	<td><%=arrHList(8,intH)%></td>
	<td><%=arrHList(9,intH)%></td>
	<% end if %>

	<td><%=arrHList(2,intH)%></td>
	<td><%IF isNull(arrHList(10,intH)) THEN%><font color="red">�̹߼�</font><%ELSE%><%=arrHList(10,intH)%><%END IF%></td>
	<td>
		<% IF (isNull(arrHList(10,intH)) OR DateDiff("d", NOW, arrHList(10,intH)) > -7) and Not(C_ADMIN_AUTH or C_SYSTEM_Part or C_CSUser or session("ssAdminPsn")="14" or session("ssAdminPsn")="23") THEN %>
			&nbsp;
		<% ElseIf C_ADMIN_AUTH or C_SYSTEM_Part or C_CSUser or session("ssAdminPsn")="14" or session("ssAdminPsn")="23" Then %>
			<input type="button" value="��û" onClick="jsReApply(<%=iHVol%>,'','<%=iAVol%>','<%=arrHList(13,intH)%>');" class="a">
		<% Else %>
			<input type="button" value="��û" onClick="jsReApply(<%=iHVol%>,'','<%=iAVol%>','<%=arrHList(13,intH)%>');" class="a">
		<% End If %>
	</td>
	</tr>
	<%NEXT%>
	<%ELSE%>
	<tr align="center" bgcolor="ffffff">
	<td colspan="10" align="center">��ϵ� ������ �����ϴ�.</td>
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
    <td  background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
</form>
<!-- /����¡ó�� -->
</div>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
