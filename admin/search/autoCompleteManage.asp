<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/search_manageCls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)
'// ���ã��
dim IsMenuFavoriteAdded
IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)


Dim i, cAuto, vIdx, vAutoType, vTitle, vURL_PC, vURL_M, vIcon, vRegUserName, vRegdate, vLastUserName, vLastdate, vMemo, vUseYN, vSortNo
vIdx = requestCheckVar(Request("idx"),15)

If vIdx <> "" Then
	Set cAuto = New CSearchMng
	cAuto.FRectIdx = vIdx
	cAuto.sbAutoCompleteDetail

	vAutoType = cAuto.FOneItem.Fautotype
	vTitle = cAuto.FOneItem.Ftitle
	vURL_PC = cAuto.FOneItem.Furl_pc
	vURL_M = cAuto.FOneItem.Furl_m
	vIcon = cAuto.FOneItem.Ficon
	vRegUserName = cAuto.FOneItem.Fregusername
	vRegdate = cAuto.FOneItem.Fregdate
	vLastUserName = cAuto.FOneItem.Flastusername
	vLastdate = cAuto.FOneItem.Flastdate
	vMemo = cAuto.FOneItem.Fmemo
	vSortNo = cAuto.FOneItem.Fsortno
	vUseYN = cAuto.FOneItem.Fuseyn
	Set cAuto = Nothing
Else
	vUseYN = "y"
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<script language='javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;
	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "���ã�⿡�� �����Ͻðڽ��ϱ�?";
	} else {
		msg = "���ã�⿡ �߰��Ͻðڽ��ϱ�?";
	}
	ret = confirm(msg);
	if (ret) {
		frm.submit();
	}
}

function jsAutoCompleteSave(){
	if($(":radio[name=autotype]:checked").length == "0"){
		alert("�ڵ��ϼ� �Ӽ��� �����ϼ���.");
		return;
	}
	if($("#title").val() == ""){
		alert("������ �Է��ϼ���.");
		return;
	}
	if(!frm1.autotype[3].checked){
		if($("#url_pc").val() == ""){
			alert("URL PC�� �Է��ϼ���.");
			return;
		}
		if($("#url_m").val() == ""){
			alert("URL M�� �Է��ϼ���.");
			return;
		}
	}else{
		$("#url_pc").val("");
		$("#url_m").val("");
	}
	if($(":radio[name=icon]:checked").length == "0"){
		alert("������ ������ �����ϼ���.");
		return;
	}

	frm1.submit();
}

//��ũ������
function showDrop(g){
	$("#selectLink"+g+"").show();
}

function linkcopy(g){
	var val = $("#url_"+g+"").val();
	$("#selectLink"+g+"").css("display","none");
}

//�����Է�
function populateTextBox(v,g){
	var val = v;
	$("#url_"+g+"").val(val);
	$("#selectLink"+g+"").css("display","none");
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play2016Cls.asp" -->
<div class="contSectFix scrl">
	<div class="cont">
		<form name="frm1" action="autoCompleteProc.asp" method="post">
		<input type="hidden" name="idx" value="<%=vIdx%>">
		<div class="searchWrap inputWrap">
			<h3>- �ڵ��ϼ� ����</h3>
			<table class="writeTb tMar10">
				<colgroup>
					<col width="15%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>�ڵ��ϼ� �Ӽ� *</div></th>
					<td>
						<span class="rMar10"><input type="radio" id="sc" name="autotype" value="sc" <%=CHKIIF(vAutoType="sc","checked","")%> /> <label for="sc">�ٷΰ���</label></span>
						<span class="rMar10"><input type="radio" id="ca" name="autotype" value="ca" <%=CHKIIF(vAutoType="ca","checked","")%> /> <label for="ca">ī�װ�</label></span>
						<span class="rMar10"><input type="radio" id="br" name="autotype" value="br" <%=CHKIIF(vAutoType="br","checked","")%> /> <label for="br">�귣��</label></span>
						<span class="rMar10"><input type="radio" id="ky" name="autotype" value="ky" <%=CHKIIF(vAutoType="ky","checked","")%> /> <label for="ky">Ű����</label></span>
					</td>
				</tr>
				<tr>
					<th><div>���� *</div></th>
					<td><input type="text" class="formTxt" id="title" name="title" value="<%=vTitle%>" maxlength="20" placeholder="20�� �̳��� �ڵ��ϼ� ������ �Է����ּ���." style="width:50%" /></td>
				</tr>
				<tr>
					<th></th>
					<td><strong>
						<font color="blue">�� �ڵ��ϼ� �Ӽ��� "�ٷΰ���" �� ���<br />ī�װ�, �귣��, Ű���� �� �� ����Ƽ��� ������ �������� �Ұ��մϴ�. ������������ �Է��ϼ���.</font>
					</strong></td>
				</tr>
				<tr>
					<th><div>URL PC *</div></th>
					<td>
						<div class="selectLink">
							<input type="text" class="formTxt" value="<%=CHKIIF(vURL_PC="","��ũ�� �Է�(����)",vURL_PC)%>" onclick="showDrop('pc');" id="url_pc" name="url_pc" onkeyup="linkcopy('pc');" maxlength="200" />
							<ul style="display:none;" id="selectLinkpc">
								<li onclick="populateTextBox('<%=CHKIIF(vURL_PC="","",vURL_PC)%>','pc');">���þ���</li>
								<li onclick="populateTextBox('/category/category_prd.asp?itemid=��ǰ�ڵ�','pc');">/category/category_prd.asp?itemid=��ǰ�ڵ�</li>
								<li onclick="populateTextBox('/shopping/category_list.asp?disp=ī�װ�','pc');">/shopping/category_list.asp?disp=ī�װ�</li>
								<li onclick="populateTextBox('/street/street_brand.asp?makerid=�귣����̵�','pc');">/street/street_brand.asp?makerid=�귣����̵�</li>
								<li onclick="populateTextBox('/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�','pc');">/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�</li>
								<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�','pc');">/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�</li>
								<li onclick="populateTextBox('/gift/gifttalk/','pc');">����Ʈ</li>
								<li onclick="populateTextBox('/wish/index.asp','pc');">����</li>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<th><div>URL M *</div></th>
					<td>
						<div class="selectLink">
							<input type="text" class="formTxt" value="<%=CHKIIF(vURL_M="","��ũ�� �Է�(����)",vURL_M)%>" onclick="showDrop('m');" id="url_m" name="url_m" onkeyup="linkcopy('m');" maxlength="200" />
							<ul style="display:none;" id="selectLinkm">
								<li onclick="populateTextBox('<%=CHKIIF(vURL_M="","",vURL_M)%>','m');">���þ���</li>
								<li onclick="populateTextBox('/category/category_itemPrd.asp?itemid=��ǰ�ڵ�','m');">/category/category_itemPrd.asp?itemid=��ǰ�ڵ�</li>
								<li onclick="populateTextBox('/category/category_list.asp?disp=ī�װ�','m');">/category/category_list.asp?disp=ī�װ�</li>
								<li onclick="populateTextBox('/street/street_brand.asp?makerid=�귣����̵�','m');">/street/street_brand.asp?makerid=�귣����̵�</li>
								<li onclick="populateTextBox('/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�','m');">/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�</li>
								<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�','m');">/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�</li>
								<li onclick="populateTextBox('/gift/gifttalk/','m');">����Ʈ</li>
								<li onclick="populateTextBox('/wish/index.asp','m');">����</li>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<th><div>������ ���� *</div></th>
					<td>
						<span class="rMar10"><input type="radio" id="none" name="icon" value="none" <%=CHKIIF(vIcon="none","checked","")%> /> <label for="none">������</label></span>
						<span class="rMar10"><input type="radio" id="best" name="icon" value="best" <%=CHKIIF(vIcon="best","checked","")%> /> <label for="best">����Ʈ</label></span>
						<span class="rMar10"><input type="radio" id="jump" name="icon" value="jump" <%=CHKIIF(vIcon="jump","checked","")%> /> <label for="jump">�޻�� �˻���</label></span>
					</td>
				</tr>
				<tr>
					<th><div>��뿩�� *</div></th>
					<td>
						<span class="rMar10"><input type="radio" id="useyny" name="useyn" value="y" <%=CHKIIF(vUseYN="y","checked","")%> /> <label for="useyny">�����</label></span>
						<span class="rMar10"><input type="radio" id="useynn" name="useyn" value="n" <%=CHKIIF(vUseYN="n","checked","")%> /> <label for="useynn">������</label></span>
					</td>
				</tr>
				<% If vIdx <> "" Then %>
				<tr>
					<th><div>�ۼ���</div></th>
					<td>�����۾��� : <%=vRegUserName%>, �������۾��� : <%=vLastUserName%></td>
				</tr>
				<tr>
					<th><div>�ۼ���</div></th>
					<td>�����ۼ��� : <%=vRegdate%>, �������ۼ��� : <%=vLastdate%></td>
				</tr>
				<% End If %>
				<tr>
					<th><div>���</div></th>
					<td><textarea class="formTxtA" rows="6" style="width:99%;" id="memo" name="memo"><%=vMemo%></textarea></td>
				</tr>
				</tbody>
			</table>
			<div class="tMar20 ct">
				<input type="button" value="����" onclick="jsAutoCompleteSave();" class="cRd1" style="width:100px; height:30px;" />
				<input type="button" value="���" onclick="window.close();" style="width:100px; height:30px;" />
			</div>
		</div>
		</form>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->