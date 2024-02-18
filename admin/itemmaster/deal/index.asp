<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/index.asp
' Description :  �� �̺�Ʈ ����
' History : 2017.08.22 ������
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->


<%
	'��������
	Dim iCurrpage, iPageSize, iPerCnt, isResearch, sSdate, sEdate, intLoop, stext, dispCate
	Dim oDeal, arrList, iTotCnt, iTotalPage, strTxt, sdiv, datediv, viewdiv, isusing, arrCate, maxDepth

	dispCate	= requestCheckVar(Request("disp"),16) 		'���� ī�װ�
	maxDepth = 2
	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����

	isusing 		= requestCheckVar(Request("isusing"),1)
	viewdiv 		= requestCheckVar(Request("viewdiv"),1)
	datediv 		= requestCheckVar(Request("datediv"),1)
	sdiv 		= requestCheckVar(Request("sdiv"),10)
	strTxt 		= requestCheckVar(Request("stext"),32)
	
	isResearch = requestCheckVar(Request("isResearch"),1)
	if isResearch ="" then isResearch ="0"
	'## �˻� #############################
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	'������ ��������
	set oDeal = new ClsDeal
		oDeal.FCPage = iCurrpage		'����������
		oDeal.FPSize = iPageSize		'���������� ���̴� ���ڵ尹��
		oDeal.FSearchDateDiv 	= datediv	'�˻��� ����
		oDeal.FSsDate 	= sSdate	'�˻� ������
		oDeal.FSeDate 	= sEdate	'�˻� ������
		oDeal.FSearchDiv 	= sdiv	'�˻�����
		oDeal.FSeTxt 	= strTxt	'�˻���
		oDeal.FSViewDiv 	= viewdiv	'���� ����
		oDeal.FSIsUsing 	= isusing	'��� ����
		oDeal.FSdispCate 	= dispCate	'����ī�װ� �˻�
 		arrList = oDeal.fnGetDealList	'�����͸�� ��������
 		iTotCnt = oDeal.FTotCnt	'��ü ������  ��
 	set oDeal = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
<!--
	window.document.domain = "10x10.co.kr";
	function jsSearch(sType){
		var frm = document.frmEvt
		if (sType == "A"){
				frm.iSD.value = "";
				frm.iED.value = "";
				frm.eventstate.value = "";
				frm.sEtxt.value = "";
				frm.selC.value = "";
		}
		if(frm.sdiv.value=="itemid" && frm.stext.value!=""){
			if(isNaN(frm.stext.value)){
				alert("��ǰ��ȣ �˻��� ���ڸ� �Է����ּ���!");
				return false;
			}
		}

		frm.submit();
	}
	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}
	function TnEditDeal(url){
		location.href=url;
	}

	//�̸�����
	function jsOpen(sPURL,sTG){ 
	    if (sTG =="M" ){ 
	        var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes,location=yes");
	    }
	}

	function fnDealInfoUpdate(){
		$.ajax({
			type: "POST",
			url: "ajaxDealInfoUpdate.asp",
			data: "mode=all",
			cache: false,
			async: false,
			success: function(message) {
				if(message=="OK") {
					alert("������Ʈ �Ϸ�.");
				} else {
					alert("���� �� ������ �����ϴ�.");
				}
			}
		});
	}

	function fnDealItemInfoUpdate(itemid){
		$.ajax({
			type: "POST",
			url: "ajaxDealInfoUpdate.asp",
			data: "mode=one&itemid="+itemid,
			cache: false,
			async: false,
			success: function(message) {
				if(message=="OK") {
					alert("������Ʈ �Ϸ�.");
				} else {
					alert("���� �� ������ �����ϴ�.");
				}
			}
		});
	}

    function TnDevDealSaveAPICall(itemid){
		$.ajax({
			type: "POST",
			url: "<%= ItemUploadUrl %>/linkweb/items/deal_itemregisterTempWithImage_process.asp",
			data: "itemid=" + itemid,
			dataType: "JSON",
			cache: false,
			success: function(data) {
				alert(data.message);
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
    }
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmEvt" method="get"  action="index.asp" onSubmit="return jsSearch('E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">�˻� ����</td>
	<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
		<table>
		<tr>
			<td>
				�Ⱓ:
				<select name="datediv">
					<option value="S"<% If datediv="S" Then Response.write " selected" %>>������</option>
					<option value="E"<% If datediv="E" Then Response.write " selected" %>>������</option>
					<option value="R"<% If datediv="R" Then Response.write " selected" %>>�ۼ���</option>
				</select>
				<input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
				<input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "iSD", trigger    : "iSD_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "iED", trigger    : "iED_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</td>
		</tr>
		<tr>
			<td>
				���� : 
				<select name="viewdiv" class="select">
					<option value="" selected>��ü</option>
					<option value="1"<% If viewdiv="1" Then Response.write " selected" %>>��õ�</option>
					<option value="2"<% If viewdiv="2" Then Response.write " selected" %>>�Ⱓ��</option>
				</select>
				��뿩�� : 
				<select name="isusing" class="select">
					<option value="" selected>��ü</option>
					<option value="Y"<% If isusing="Y" Then Response.write " selected" %>>���</option>
					<option value="N"<% If isusing="N" Then Response.write " selected" %>>������</option>
				</select>
				���� ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
			</td>
		</tr>
		<tr>
			<td>
				�˻��� : 
				<select name="sdiv" class="select">
					<option value="itemid"<% If sdiv="itemid" Then Response.write " selected" %>>����ǰ�ڵ�</option>
					<option value="itemname"<% If sdiv="itemname" Then Response.write " selected" %>>��ǰ��</option>
					<option value="register"<% If sdiv="register" Then Response.write " selected" %>>�ۼ���</option>
					<option value="makerid"<% If sdiv="makerid" Then Response.write " selected" %>>�귣����̵�</option>
				</select>
				<input type="text" name="stext" size="50" value="<%=strTxt%>" onkeydown="if(event.keyCode==13) jsSearch('E');">
			</td>
		</tr>
		</table>
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>" align="center"><input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch('E');"></td>
</tr>
</form>
</table><br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="13">
			<table width="100%">
			<tr>
				<td>�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
				<td align="right"><% if session("ssBctId")="seojb1983" or C_ADMIN_AUTH then %><input type="button" class="button" style="width:105;" value="����������Ʈ" onclick="fnDealInfoUpdate();">&nbsp;&nbsp;&nbsp;<% end if %><input type="button" class="button" style="width:105;" value="�� �������" onclick="jsGoUrl('/admin/dataanalysis/report/weeklysimplereport.asp?menupos=4019&reporttype=dealsales');">&nbsp;&nbsp;&nbsp;<input type="button" class="button" style="width:105;" value="���" onclick="jsGoUrl('/admin/itemmaster/deal/new_deal_reg.asp?menupos=<%=menupos%>');"></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>���ڵ�</td>
		<td>����ǰ�ڵ�</td>
		<td>ī�װ�</td>
		<td>����</td>
		<td>����Ⱓ</td>
		<td>��뿩��</td>
		<td>����ǰ��</td>
		<td>�ٹ����ٰ�</td>
		<td>������</td>
		<td>�ۼ���</td>
		<td>�ۼ���</td>
		<td>�̸�����</td>
		<td>����</td>
	 </tr>
	 <% If isArray(arrList) Then %>
	 <% For intLoop = 0 To UBound(arrList,2) %>
	 <% if arrList(14,intLoop)="0" then %>
	 <tr bgcolor="#EEEEEE">
	 <% else %>
	 <tr bgcolor="#FFFFFF">
	 <% End If %>
		<td align="center"><%=arrList(0,intLoop)%></td>
		<td align="center"><%=arrList(1,intLoop)%></td>
		<td align="center">
		<%
			If arrList(12,intLoop) <> "" Then
			arrCate = Split(arrList(12,intLoop),"^^")
			If ubound(arrCate)>0 Then
			Response.write arrCate(0) & " > " & arrCate(1)
			Else
			Response.write arrCate(0)
			End If
			End If
		%>
		</td>
		<td align="center"><% If arrList(2,intLoop)="1" Then %>��õ�<% Else %>�Ⱓ��<% End If %></td>
		<td align="center"><% If arrList(2,intLoop)<>"2" Then %>��� ����<% Else %><%=FormatDateTime(arrList(3,intLoop),2)%> ~ <%=FormatDateTime(arrList(4,intLoop),2)%> <% End If %></td>
		<td align="center"><% if arrList(14,intLoop)="0" then %>��ϴ��<% Else %><% If arrList(13,intLoop) = "Y" Then %>���<% Else %>������<% End If %><% End If %></td>
		<td><%=arrList(5,intLoop)%></td>
		<td align="right"><%=FormatNumber(arrList(6,intLoop),0)%>��<% If arrList(10,intLoop) = "Y" Then %>~<% Else %>&nbsp;&nbsp;<% End If %></td>
		<td align="center"><% If arrList(11,intLoop) = "Y" Then %>~<% End If %><%=arrList(7,intLoop)%>%</td>
		<td align="center"><%=arrList(8,intLoop)%></td>
		<td align="center"><%=left(arrList(9,intLoop),10)%></td>
		<td align="center"><a href="<%=vwwwUrl%>/deal/deal.asp?itemid=<%=arrList(1,intLoop)%>" target="_blank"><img src="/images/iexplorer.gif" border="0"></a>&nbsp;<a href="javascript:jsOpen('<%=vmobileUrl%>/deal/deal.asp?itemid=<%=arrList(1,intLoop)%>','M');"><img src="/images/iexplorer.gif" border="0"></a></td>
		<td align="center">
			<% if arrList(14,intLoop)="0" and (application("Svr_Info")="Dev") then %>
			<input type="button" class="button" style="width:105;" value="�������" onclick="TnDevDealSaveAPICall(<%= arrList(1,intLoop) %>);">
			<% End If %>
			<input type="button" class="button" style="width:105;" value="����" onclick="TnEditDeal('/admin/itemmaster/deal/new_deal_edit.asp?idx=<%= arrList(0,intLoop) %>');"<% if arrList(14,intLoop)="0" then %> disabled<% End If %>>
			<input type="button" class="button" style="width:105;" value="����������Ʈ" onclick="fnDealItemInfoUpdate(<%= arrList(1,intLoop) %>);">
		</td>
	 </tr>
	 <% Next %>
	 <% Else %>
	 <tr bgcolor="#FFFFFF">
		<td colspan="13" align="center" height="25">
			��ϵ� ������ �����ϴ�.
		</td>
	 </tr>
	 <% End If %>
	 <tr bgcolor="#FFFFFF">
		<td colspan="13" bgcolor="#FFFFFF" align="center">
			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	 </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->