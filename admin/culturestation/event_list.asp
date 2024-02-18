<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Event  
' History : 2009.04.02 �ѿ�� ����
'           2012.01.12 ������; ����� �߰�, ����� ����
'           2013.06.04 ������; ������¿� ���� ���� �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<%
Dim oip,i,page,evt_type_search,isusing_search,evt_code_search, evt_name_search, evt_code_count, evt_mobile_yn, evt_partner_search
Dim edid, emid, rowColor
Dim sDate,sSdate,sEdate, sortMtd, srchStat
	evt_code_search = request("evt_code_search")
	evt_name_search = request("evt_name_search")
	evt_partner_search = request("evt_partner_search")
	evt_type_search = request("evt_type_searchbox")
	isusing_search = request("isusing_searchbox")
	evt_code_count = request("evt_code_countbox")
	evt_mobile_yn = request("evt_mobile_yn")
	menupos = request("menupos")
	page = request("page")
	sortMtd = request("sortMtd")
	srchStat = request("srchStat")

	edid  		= requestCheckVar(Request("selDId"),32)		'��� �����̳�
	emid  		= requestCheckVar(Request("selMId"),32)		'��� MD

	sDate 		= requestCheckVar(Request("selDate"),1)  	'�Ⱓ
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	if page = "" then page = 1

'// �̺�Ʈ ����Ʈ
set oip = new cevent_list
	oip.FPageSize = 50
	oip.FCurrPage = page
	oip.frectevt_type = evt_type_search
	oip.frectisusing = isusing_search
	oip.frectevt_code = evt_code_search
	oip.frectevt_partner = evt_partner_search
	oip.frectevt_name = evt_name_search
	oip.frectevt_code_count = evt_code_count
	oip.frectM_isUsing = evt_mobile_yn
	oip.frectSortMethod = sortMtd
	oip.frectStatus = srchStat

	oip.fedid	= edid
	oip.femid	= emid

	oip.fdate	= sDate
	oip.fsdate	= sSdate
	oip.fedate	= sEdate

	oip.fevent_list()
%>

<script language="javascript">

function event_edit(evt_code){
	var event_edit = window.open('/admin/culturestation/event_edit.asp?evt_code='+evt_code,'addreg','width=800,height=768,scrollbars=yes,resizable=yes');
	event_edit.focus();
}

function AnSelectAllFrame(bool){
	var frm = document.frmBuyPrc;
	if(frm.chkitem.length>1) {
		for (var i=0;i<frm.chkitem.length;i++){
			if (frm.chkitem[i].disabled!=true){
				frm.chkitem[i].checked = bool;
				AnCheckClick(frm.chkitem[i]);
			}
		}
	} else {
		frm.chkitem.checked = bool;
		AnCheckClick(frm.chkitem);
	}
}	

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}	

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;

	var frm = document.frmBuyPrc;
	if(frm.chkitem.length>1) {
		for (var i=0;i<frm.chkitem.length;i++){
			pass = ((pass)||(frm.chkitem[i].checked));
		}
	} else {
		pass = ((pass)||(frm.chkitem.checked));
	}

	if (!pass) {
		return false;
	}
	return true;
}

// MainPage Category Image����   
function maincategoryimage(evt_code,evt_type){

	var maincategoryimage;
	maincategoryimage = window.open("<%=wwwUrl%>/chtml/culturestation_maincate_imagemake.asp?evt_code=" +evt_code + '&evt_type='+evt_type, "maincategoryimage","width=400,height=300,scrollbars=yes,resizable=yes");
	maincategoryimage.focus();
}

// SubPage Category���� 
function AssignReal(upfrm,evt_type){
if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}	

	var frm = document.frmBuyPrc;
	if(frm.chkitem.length>1) {
		for (var i=0;i<frm.chkitem.length;i++){
			if (frm.chkitem[i].checked)
				upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code[i].value + "," ;
		}
	} else {
		if (frm.chkitem.checked)
			upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code.value + "," ;
	}

	var tot;
	tot = upfrm.evt_code.value;
	upfrm.evt_code.value = ""

	var AssignReal;
	AssignReal = window.open("<%=wwwUrl%>/chtml/culturestation_categorymake_new.asp?evt_code=" +tot + '&evt_type='+evt_type, "AssignReal","width=400,height=300,scrollbars=yes,resizable=yes");
	AssignReal.focus();
}

function prize(evt_code){

	 var prize = window.open('/admin/culturestation/event_prize.asp?evt_code='+evt_code,'prize','width=800,height=600,scrollbars=yes,resizable=yes');
	 prize.focus();

}

function comment_list(evt_code){

	 var comment_list = window.open('/admin/culturestation/event_comment_list.asp?evt_code='+evt_code,'comment_list','width=800,height=600,scrollbars=yes,resizable=yes');
	 comment_list.focus();

}

function save_mSortNo() {
	var frm = document.frmBuyPrc;
	frm.action="event_sortNo_process.asp";
	frm.submit();
}

function save_webSortNo() {
	var frm = document.frmBuyPrc;
	frm.action="event_websortNo_process.asp";
	frm.submit();
}

function goPage(pg) {
	var frm = document.frm;
	frm.evt_code.value="";
	frm.page.value=pg;
	frm.action="";
	frm.submit();
}
function RefreshMainCorItemRec(upfrm){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	var frm = document.frmBuyPrc;
	if(frm.chkitem.length>1) {
		for (var i=0;i<frm.chkitem.length;i++){
			if (frm.chkitem[i].checked)
				upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code[i].value + "," ;
		}
	} else {
		if (frm.chkitem.checked)
			upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code.value + "," ;
	}

	var tot;
	tot = upfrm.evt_code.value;
	upfrm.evt_code.value = ""
	var AssignReal;
//	alert(tot);
//	return false;
	AssignReal = window.open("<%=wwwUrl%>/chtml/main_curture_make12banner.asp?evt_code=" +tot, "AssignReal","width=400,height=300,scrollbars=yes,resizable=yes");
	AssignReal.focus();
}

//�����
function MORefreshMainCorItemRec(upfrm){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	var frm = document.frmBuyPrc;
	if(frm.chkitem.length>1) {
		for (var i=0;i<frm.chkitem.length;i++){
			if (frm.chkitem[i].checked)
				upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code[i].value + "," ;
		}
	} else {
		if (frm.chkitem.checked)
			upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code.value + "," ;
	}

	var tot;
	tot = upfrm.evt_code.value;
	upfrm.evt_code.value = ""
	var AssignReal;
	AssignReal = window.open("<%=mobileUrl%>/chtml/main/loader/2015loader/main_curture_make12banner.asp?evt_code=" +tot, "AssignReal","width=400,height=300,scrollbars=yes,resizable=yes");
	AssignReal.focus();
}



function TnSearchEvtSelect(objval){
	if(objval=="evt_code_search"){
		$("#evt_code_search").css("display","");
		$("#evt_name_search").css("display","none").val("");
		$("#evt_partner_search").css("display","none").val("");
	}else if(objval=="evt_name_search"){
		$("#evt_code_search").css("display","none").val("");
		$("#evt_name_search").css("display","");
		$("#evt_partner_search").css("display","none").val("");
	}else{
		$("#evt_code_search").css("display","none").val("");
		$("#evt_name_search").css("display","none").val("");
		$("#evt_partner_search").css("display","");
	}
}
</script>
<script type="text/javascript" src="/js/jquery-2.2.2.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
function event_contents_delete(evtcode){
	var str = $.ajax({
		type: "POST",
		url: "/admin/culturestation/event_edit_process.asp",
		data: "mode=del&evt_code="+evtcode,
		dataType: "text",
		async: false
	}).responseText;
	if (str  == "ok"){
		alert("�����Ͻ� �̺�Ʈ�� �������� ���� �Ǿ����ϴ�.");
	}else{
		alert("������ ������ ������ �ֽ��ϴ�. ���� �ٶ��ϴ�.");
	}
}

$(function(){
	TnSearchEvtSelect(document.frm.ses.value);
});
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="evt_code">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page">
	<input type="hidden" name="sortMtd" value="<%=sortMtd%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">�Ⱓ : 
			<select name="selDate" class="select">
		    	<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
		    	<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
		    	<option value="V" <%if Cstr(sDate) = "V" THEN %>selected<%END IF%>>��ǥ�� ����</option>
			</select>
	        <input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
	        <input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	        /
	        <input type="checkbox" name="srchStat" value="Y" <%=chkIIF(srchStat="Y","checked","")%> />�������� �̺�Ʈ�� ����
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
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm.page.value=1;frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			����:
			<select name="evt_type_searchbox" value="<%=evt_type_search%>" class="select">
				<option value="" <% if evt_type_search = "" then response.write " selected" %>>��ü</option>
				<option value="0" <% if evt_type_search = "0" then response.write " selected" %>>������</option>
				<option value="1" <% if evt_type_search = "1" then response.write " selected" %>>�о��</option>
				<option value="2" <% if evt_type_search = "2" then response.write " selected" %>>����</option>
			</select> /
			��뿩��:
			<select name="isusing_searchbox" value="<%=isusing_search%>" class="select">
				<option value="" <% if isusing_search = "" then response.write " selected" %>>��ü</option>
				<option value="Y" <% if isusing_search = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing_search = "N" then response.write " selected" %>>N</option>
			</select> /
			�ڸ�Ʈ ���:
			<select name="evt_code_countbox" value="<%=evt_code_count%>" class="select">
				<option value="" <% if evt_code_count = "" then response.write " selected" %>>��ü</option>
				<option value="Y" <% if evt_code_count = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if evt_code_count = "N" then response.write " selected" %>>N</option>
			</select> /
			����� ���:
			<select name="evt_mobile_yn" value="<%=evt_mobile_yn%>" class="select">
				<option value="" <% if evt_mobile_yn = "" then response.write " selected" %>>��ü</option>
				<option value="Y" <% if evt_mobile_yn = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if evt_mobile_yn = "N" then response.write " selected" %>>N</option>
			</select>
			WD���: <%sbGetDesignerid "selDId",edid, "onChange='javascript:document.frm.submit();'"%>
			�����ô��: <%sbGetMKTid "selMId",emid, "onChange='javascript:document.frm.submit();'"%>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >

		<td align="left">
			�˻��� :
			<select name="ses" onChange="TnSearchEvtSelect(this.value)">
				<option value="evt_code_search" <%=chkIIF(evt_code_search<>"","selected","")%>>�̺�Ʈ�ڵ�</option>
				<option value="evt_name_search" <%=chkIIF(evt_name_search<>"","selected","")%>>�̺�Ʈ��</option>
				<option value="evt_partner_search" <%=chkIIF(evt_partner_search<>"","selected","")%>>�����ü</option>
			</select>
			<input type="text" name="evt_code_search" id="evt_code_search" value="<%= evt_code_search%>" size="20" class="text">
			<input type="text" name="evt_name_search" id="evt_name_search" value="<%= evt_name_search%>" size="20" class="text" style="display:none">
			<input type="text" name="evt_partner_search" id="evt_partner_search" value="<%= evt_partner_search%>" size="20" class="text" style="display:none">
		</td>	
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<img src="/images/icon_reload.gif" onClick="javascript:RefreshMainCorItemRec(frm);" style="cursor:pointer" align="absmiddle" alt="XML�����">����Ʈ�� ����
	</td>
</tr>
<!-- <tr>
	<td align="left">
		<img src="/images/icon_reload.gif" onClick="javascript:MORefreshMainCorItemRec(frm);" style="cursor:pointer" align="absmiddle" alt="XML�����">����� ����Ʈ�� ����(2������)
	</td>
</tr> -->
</table>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if evt_type_search <> "" and evt_code_search = "" then %>				
				&nbsp;&nbsp;<a href="javascript:AssignReal(frm,<%= evt_type_search %>);"><img src="/images/refreshcpage.gif" border="0">SubPage Category����</a>		
			<% end if %>
		</td>
		<td align="right">
			<% if evt_mobile_yn="Y" then %>	<input type="button" class="button" value="Mobile���ļ��� ����" onclick="save_mSortNo();"> /<% end if %>
			<% if (evt_type_search<>"" or srchStat="Y") and evt_code_search="" and evt_mobile_yn<>"Y" then %><input type="button" class="button" value="Web���ļ��� ����" onclick="save_webSortNo();"> /<% end if %>
			<select class="select" onchange="document.frm.sortMtd.value=this.value;document.frm.submit();">
				<option value="">��ϼ�</option>
				<option value="ws" <%=chkIIF(sortMtd="ws","selected","")%>>�� ���ļ�</option>
				<option value="ms" <%=chkIIF(sortMtd="ms","selected","")%>>����� ���ļ�</option>
			</select> /
			<input type="button" class="button" value="Event���" onclick="event_edit('');">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<form action="" name="frmBuyPrc" method="POST" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="evt_code_search" value="<%= evt_code_search %>">
<input type="hidden" name="evt_name_search" value="<%= evt_name_search %>">
<input type="hidden" name="evt_partner_search" value="<%= evt_partner_search %>">
<input type="hidden" name="evt_type_searchbox" value="<%= evt_type_search %>">
<input type="hidden" name="isusing_searchbox" value="<%= isusing_search %>">
<input type="hidden" name="evt_code_countbox" value="<%= evt_code_count %>">
<input type="hidden" name="evt_mobile_yn" value="<%= evt_mobile_yn %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="sortMtd" value="<%=sortMtd%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oip.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="9">
			�˻���� : <b><%= oip.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
		<td colspan="7" align="right">
			����:
			<span style="padding:0 3px;border:1px #ccc solid;background-color:#fff;">���� ������</span>&nbsp;
			<span style="padding:0 3px;border:1px #ccc solid;background-color:#cfc;">��÷��O/����</span>&nbsp;
			<span style="padding:0 3px;border:1px #ccc solid;background-color:#fea;">��÷��X/����</span>&nbsp;
			<span style="padding:0 3px;border:1px #ccc solid;background-color:#fcc;">��������</span>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td align="center" >�̺�Ʈ �ڵ�</td>
		<td align="center">�̹���</td>
		<td align="center">New���</td>
		<td align="center">�̺�Ʈ Ÿ��</td>
		<td align="center">�̺�Ʈ��</td>
		<td align="center">�����ü</td>
		<td align="center">������</td>
		<td align="center">������</td>
		<td align="center">��ǥ��</td>
		<td align="center">���</td>
		<td align="center">�ڸ�Ʈ��</td>
		<td align="center">�����</td>
		<!-- <td align="center" title="�� ����Ʈ ��Ͽ� ǥ�õ� ���� ����">Web����</td>
		<td align="center" title="����� ����Ʈ ��Ͽ� ǥ�õ� ���� ����">Mobile����</td> -->
		<td align="center">�����ô��</td>
		<td align="center">WD���</td>
		<td align="center">���</td>
    </tr>
	<% for i=0 to oip.FresultCount-1 %>
	<%
		'// ������¿� ���� ���� ����
		if oip.FItemList(i).fisusing="N" and oip.FItemList(i).fprizeyn="Y" then
			'��������
			rowColor = "#FFCCCC"
		elseif (oip.FItemList(i).fisusing="N" or (datediff("d",oip.FItemList(i).fenddate,date)>0)) and oip.FItemList(i).fprizeyn="N" then
			'��÷��X/����
			rowColor = "#FFEEAA"
		elseif oip.FItemList(i).fisusing="Y" and oip.FItemList(i).fprizeyn="Y" then
			'��÷��O/����
			rowColor = "#CCFFCC"
		else
			'������
			rowColor = "#FFFFFF"
		end if
	%>
    <tr align="center" bgcolor="<%=rowColor%>">
			<td align="center">
				<input type="checkbox" name="chkitem" value="<%= oip.FItemList(i).fevt_code %>" onClick="AnCheckClick(this);">
				<input type="hidden" name="evt_code" value="<%= oip.FItemList(i).fevt_code %>">
			</td>
			<td align="center">
				<a href="<%=wwwUrl%>/culturestation/culturestation_event.asp?evt_code=<%= oip.FItemList(i).fevt_code %>" target="_blink" onfocus="this.blur()">
				<%= oip.FItemList(i).fevt_code %><br>[����]</a>
			</td>		
			<td align="center">
				<image src="<%=webImgUrl%>/culturestation/2009/list/<%= oip.FItemList(i).fimage_list %>" width="40" height="40" border=0>
			</td>
			<td align="center">
				<image src="<%=webImgUrl%>/culturestation/2009/list120/<%= oip.FItemList(i).fimage_list %>" width="40" border=0>
			</td>
			<td align="center">
			<% if oip.FItemList(i).fevt_type = "0" then 
					response.write "������"
				elseif oip.FItemList(i).fevt_type = "1" then
					response.write "�о��"
				else
					response.write "����"							
				end if%></td>
			<td align="center">
				<%= oip.FItemList(i).fevt_name %>
			</td>
			<td align="center">
				<%= oip.FItemList(i).fevt_partner %>
			</td>
			<td align="center"><%= left(oip.FItemList(i).fstartdate,10) %></td>		
			<td align="center"><%= left(oip.FItemList(i).fenddate,10) %></td>
			<td align="center"><%= left(oip.FItemList(i).feventdate,10) %></td>
			<td align="center"><%= "<span title='�̺�Ʈ ��뿩��'>" & oip.FItemList(i).fisusing & "</span> <font color='darkgray' title='����� ��뿩��'>(" & oip.FItemList(i).fm_isusing & ")</font>" %></td>
			<td align="center">
				<% if oip.FItemList(i).fevt_code_count = 0 then %>
				0
				<% else %>
					<a href="javascript:comment_list(<%= oip.FItemList(i).fevt_code %>);" onfocus="this.blur()">
					<%= oip.FItemList(i).fevt_code_count %><br>[����]</a>
				<% end if %>
			</td>
			<td align="center"><%= oip.FItemList(i).fsubcount %>
			<%
				If oip.FItemList(i).fsubcount > 0 Then
					Response.Write "<br><a href=""javascript:"" onClick=""window.open('pop_event_votelist_xls.asp?eC=" & oip.FItemList(i).fevt_code & "','voteXls','width=400,height=150');"">[xls�ٿ�]</a>"
				End If
			%>
			</td>

			<!-- <td align="center">
				<input name="web_sortNo" type="text" value="<%= oip.FItemList(i).fweb_sortNo %>" <%=chkIIF(oip.FItemList(i).fisusing="Y","class='text'","class='text_ro' readonly")%> style="width:24px; text-align:center;">
			</td>

			<td align="center">
				<input name="m_sortNo" type="text" value="<%= oip.FItemList(i).fm_sortNo %>" <%=chkIIF(oip.FItemList(i).fm_isUsing="Y","class='text'","class='text_ro' readonly")%> style="width:24px; text-align:center;">
			</td> -->
			<td align="center"><%= oip.FItemList(i).femName %></td>
			<td align="center"><%= oip.FItemList(i).fedName %></td>
			<td align="center">
				<input type="button" onclick="event_edit('<%= oip.FItemList(i).fevt_code %>');" value="����" class="button">&nbsp;<input type="button" onclick="event_contents_delete('<%= oip.FItemList(i).fevt_code %>');" value="�ʱ�ȭ" class="button">
				<input type="button" class="button" value="��÷�ڵ�� (<%= oip.FItemList(i).fprizeyn %>)" onclick="prize(<%= oip.FItemList(i).fevt_code %>);">
				<br><a href="javascript:maincategoryimage(<%= oip.FItemList(i).fevt_code %>,<%= oip.FItemList(i).fevt_type %>);">
				<img src="/images/refreshcpage.gif" border="0">MainCate Image����</a>
			</td>
    </tr>   
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="16" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="javascript:goPage(<%= oip.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:goPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="javascript:goPage(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

