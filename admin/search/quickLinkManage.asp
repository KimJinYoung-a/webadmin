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


Dim i, cQuick, vIdx, vQuickType, vName, vSubCopy, vURL_PC, vURL_M, vViewGubun, vRegUserName, vHtmlCont, vBtnName, vBtnPCLink, vBtnMLink
Dim vSDate, vEDate, vRegdate, vLastUserName, vLastdate, vMemo, vUseYN, vKwArr, vBgGubun, vBgColor, vBgImgPC, vBgImgM, vQuickBrID
Dim vQImgUseYN, vQImgPC, vQImgM, vBtnColor, vShhmmss, vEhhmmss
vIdx = requestCheckVar(Request("idx"),15)
vQuickType = NullFillWith(requestCheckVar(Request("type"),3),"txt")
if vQuickType = "" then vQuickType="nor"
If vIdx <> "" Then
	Set cQuick = New CSearchMng
	cQuick.FRectIdx = vIdx
	cQuick.sbQuickLinkDetail

	vQuickType = cQuick.FOneItem.Fquicktype
	vName = cQuick.FOneItem.Fquickname
	vURL_PC = cQuick.FOneItem.Furl_pc
	vURL_M = cQuick.FOneItem.Furl_m
	vViewGubun = cQuick.FOneItem.Fviewgubun
	vSDate = cQuick.FOneItem.Fsdate
	vEDate = cQuick.FOneItem.Fedate
	vShhmmss = cQuick.FOneItem.Fshhmmss
	vEhhmmss = cQuick.FOneItem.Fehhmmss
	vRegUserName = cQuick.FOneItem.Fregusername
	vRegdate = cQuick.FOneItem.Fregdate
	vLastUserName = cQuick.FOneItem.Flastusername
	vLastdate = cQuick.FOneItem.Flastdate
	vMemo = cQuick.FOneItem.Fmemo
	vUseYN = cQuick.FOneItem.Fuseyn
	vHtmlCont = cQuick.FOneItem.Fhtmlcont
	vBtnName = cQuick.FOneItem.Fbtnname
	vBtnPCLink = cQuick.FOneItem.Fbtn_pclink
	vBtnMLink = cQuick.FOneItem.Fbtn_mlink
	vBgGubun = cQuick.FOneItem.Fbggubun
	vBgColor = cQuick.FOneItem.Fbgcolor
	vBgImgPC = cQuick.FOneItem.Fbgimgpc
	vBgImgM = cQuick.FOneItem.Fbgimgm
	vQImgUseYN = cQuick.FOneItem.Fqimg_useyn
	vQImgPC = cQuick.FOneItem.Fqimgpc
	vQImgM = cQuick.FOneItem.Fqimgm
	vBtnColor = cQuick.FOneItem.Fbtn_color
	vQuickBrID = cQuick.FOneItem.Fbrandid
	vSubCopy = cQuick.FOneItem.Fsubcopy
	
	If IsArray(cQuick.FKeywordArr) Then
		For i =0 To UBound(cQuick.FKeywordArr,2)
			If i = 0 Then
				vKwArr = cQuick.FKeywordArr(0,i)
			Else
				vKwArr = vKwArr & "," & cQuick.FKeywordArr(0,i)
			End If
		Next
	End If
	Set cQuick = Nothing
Else
	vViewGubun = "period"
	vUseYN = "y"
	vBgGubun = "c"
	vQImgUseYN = "y"
	vShhmmss = "10:00:00"
	vEhhmmss = "09:59:59"
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
<style type="text/css">
.colorbtn {border-width:2px; border-style:solid; border-color:Red;}
</style>
<script language='javascript'>
document.domain = "10x10.co.kr";

function goTypeChange(a){
	location.href = "quickLinkManage.asp?idx=<%=vIdx%>&type="+a+"";
}

function jsQuickLinkSave(){
	<% If vQuickType = "" Then %>
	if($(":radio[name=quicktype]:checked").length == "0"){
		alert("����ũ �Ӽ��� �����ϼ���.");
		return;
	}
	<% End If %>
	<% If vQuickType = "brd" Then %>
	if($("#quickbrid").val() == ""){
		alert("�귣�带 �����ϼ���.");
		return;
	}
	<% End If %>
	if($("#url_pc").val() == ""){
		alert("URL PC�� �Է��ϼ���.");
		return;
	}
	if($("#url_m").val() == ""){
		alert("URL M�� �Է��ϼ���.");
		return;
	}
	if($("#keyword").val() == ""){
		alert("�˻� Ű���带 �Է��ϼ���.");
		return;
	}
	if($("#sdate").val() == "" || $("#edate").val() == ""){
		alert("������, �������� �Է����ּ���.");
		return;
	}
	<% If vQuickType = "nor" OR vQuickType = "set" OR vQuickType = "brd" Then %>
		if($("#oneClrUse").is(":checked") == true){
			if($("#bgcolor").val() == ""){
				alert("����ũ��� �ܻ������ ��� �÷��� ���� �Ǵ� ���� �Է����ּ���.");
				return;
			}
		}
		if($("#imgUse").is(":checked") == true){
			if($("#qbgimgurlpc").val() == "" && $("#qbgimgurlm").val() == ""){
				alert("����ũ��� �̹��� ����� ��� PC, Mobile �̹����� ������ּ���.");
				return;
			}
		}
	<% End If %>
	<% If vQuickType = "set" OR vQuickType = "brd" Then %>
		if($("#qimgUseY").is(":checked") == true){
			if($("#qimgurlpc").val() == "" && $("#qimgurlm").val() == ""){
				alert("����ũ �̹��� ����� ��� PC, Mobile �̹����� ������ּ���.");
				return;
			}
		}
	<% End If %>
	const regExp = /[0-9a-fA-F]{6}/;
    const bgColor = $("#bgcolor").val();
    if(!bgColor.match(regExp)){
       alert("16���� �����ڵ尡 �ƴմϴ�.");
       return;
    }
	frm1.submit();
}

//�귣�� ID �˻� �˾�â
function jsSearchBrandID1(frmName,compName,compName2){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("popBrandSearch_search.asp?isjsdomain=o&frmName=" + frmName + "&compName=" + compName + "&compName2=" + compName2 + "&socname_kr=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}

function jsUploadImg(a,b){
	document.domain ="10x10.co.kr";
	var popupl;
	popupl = window.open('/admin/search/pop_uploadimg.asp?folder=quick&span='+b+'&sname='+a+'','popupl','width=370,height=150');
	popupl.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

//��ũ������
function showDrop(g){
	$("#selectLink"+g+"").show();
}

function linkcopy(g){
	var val = $("#url_"+g+"").val();
	$("#selectLink"+g+"").css("display","none");
}
function showDropL(g){
	$("#selectbtnLink"+g+"").show();
}
function linkcopyL(g){
	var val = $("#url_"+g+"").val();
	$("#selectbtnLink"+g+"").css("display","none");
}
//�����Է�
function populateTextBox(v,g){
	var val = v;
	$("#url_"+g+"").val(val);
	$("#selectLink"+g+"").css("display","none");
}
function populateTextBoxL(v,g){
	var val = v;
	$("#btnlink"+g+"").val(val);
	$("#selectbtnLink"+g+"").css("display","none");
}

function jsViewGubunClear(){
	$("#sdate").val("");
	$("#edate").val("");
}

function jsBgGubun(g){
	if(g == "c"){
		$("#bgcolorselect").show();
		$("#qbgimg").hide();
	}else{
		$("#bgcolorselect").hide();
		$("#qbgimg").show();
	}
}

function jsBGColor(a,v,btn,bgc){
	$("#"+a+" > span > button").removeClass("colorbtn");
	$("#"+btn+"").addClass("colorbtn");
	$("#"+v+"").val(bgc);
}

function jsQimgGubun(g){
	if(g == "y"){
		$("#qimg").show();
	}else{
		$("#qimg").hide();
	}
}

function fnCheckColorCode(obj){
	var colorCode =  /^[A-Za-z0-9]*$/;
	if(!colorCode.test(obj.value)){
		alert("�Է°��� ����,���� ���ո� �Է� �����մϴ�.");
	}
}
</script>
<style>
.formTxtA::placeholder {color:#bbb;}
</style>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<div class="contSectFix scrl">
	<div class="cont">
	<form name="frm1" id="frm1" action="quickLinkproc.asp" method="post" style="margin:0px;">
	<input type=hidden name=quicktype value="nor">
	<input type="hidden" name="idx" value="<%=vIdx%>">
		<div class="searchWrap inputWrap">
			<h3>- ����ũ �⺻ ����</h3>
			<table class="writeTb tMar10">
				<colgroup>
					<col width="16%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>����ũ ��</div></th>
					<td>
						<div>
							<textarea class="formTxtA" rows="2" style="width:99%;" id="quickname" name="quickname" placeholder="����ũ���� �Է����ּ���.(����)"><%=vName%></textarea>
						</div>
						<div>
							<input type="text" class="formTxt" id="bgcolor" name="bgcolor" value="<%=vBgColor%>" onkeypress="fnCheckColorCode(this);" style="width:27%" placeholder="���� �Է�(����)" />
						</div>
					</td>
				</tr>
				<tr>
					<th><div>����ũ �̹���</div></th>
					<td>
						<p>
							<span class="rMar10"><input type="radio" id="qimgUseY" name="qimg_useyn" value="y" <%=CHKIIF(vQImgUseYN="y","checked","")%> onClick="jsQimgGubun('y');" /> <label for="qimgUseY">���</label></span>
							<span class="rMar10"><input type="radio" id="qimgUseN" name="qimg_useyn" value="n" <%=CHKIIF(vQImgUseYN="n","checked","")%> onClick="jsQimgGubun('n');" /> <label for="qimgUseN">������</label></span>
						</p>
						<p class="tPad10" id="qimg" style="display:<%=CHKIIF(vQImgUseYN="y","block","none")%>">
							<input type="button" value="PC ���ε�" onClick="jsUploadImg('qimgurlpc','qimgurlpcspan');" /><br /><br />
							<span id="qimgurlpcspan" style="padding:5px 5px 5px 0;"><%
								If vQImgPC <> "" Then
									Response.Write "<img src='" & vQImgPC & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vQImgPC & "');>"
									Response.Write "<a href=javascript:jsDelImg('qimgurlpc','qimgurlpcspan');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" id="qimgurlpc" name="qimgurlpc" value="<%=vQImgPC%>">
							<br /><br />
							<input type="button" value="Mobile ���ε�" onClick="jsUploadImg('qimgurlm','qimgurlmspan');" /><br /><br />
							<span id="qimgurlmspan" style="padding:5px 5px 5px 0;"><%
								If vQImgM <> "" Then
									Response.Write "<img src='" & vQImgM & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vQImgM & "');>"
									Response.Write "<a href=javascript:jsDelImg('qimgurlm','qimgurlmspan');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" id="qimgurlm" name="qimgurlm" value="<%=vQImgM%>">
							<br /><span class="tPad10 fs11 cBl3">* 2Mb ������(1125x240������) png, jpg, gif���� �̹��������� �������ּ���.</span>
						</p>
					</td>
				</tr>
				<tr style="display:none">
					<th><div>����ũ ����ī��</div></th>
					<td><input type="text" class="formTxt" id="subcopy" name="subcopy" value="<%=vSubCopy%>" placeholder="����ī�Ǹ� ����� ��� �ؽ�Ʈ�� �Է����ּ���." style="width:99%" maxlength="70" /></td>
				</tr>
				<tr>
					<th><div>URL PC *</div></th>
					<td>
						<div class="selectLink">
							<input type="text" class="formTxt" value="<%=CHKIIF(vURL_PC="","",vURL_PC)%>" placeholder="��ũ�� �Է�(����)" onclick="showDrop('pc');" id="url_pc" name="url_pc" onkeyup="linkcopy('pc');" maxlength="200" />
							<ul style="display:none;" id="selectLinkpc">
								<li onclick="populateTextBox('<%=CHKIIF(vURL_PC="","",vURL_PC)%>','pc');">���þ���</li>
								<li onclick="populateTextBox('/category/category_prd.asp?itemid=��ǰ�ڵ�','pc');">/category/category_prd.asp?itemid=��ǰ�ڵ�</li>
								<li onclick="populateTextBox('/category/category_main2020.asp?disp=ī�װ�','pc');">/category/category_main2020.asp?disp=ī�װ�</li>
								<li onclick="populateTextBox('/brand/brand_detail2020.asp?brandid=�귣����̵�','pc');">/brand/brand_detail2020.asp?brandid=�귣����̵�</li>
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
							<input type="text" class="formTxt" value="<%=CHKIIF(vURL_M="","",vURL_M)%>" placeholder="��ũ�� �Է�(����)" onclick="showDrop('m');" id="url_m" name="url_m" onkeyup="linkcopy('m');" maxlength="200" />
							<ul style="display:none;" id="selectLinkm">
								<li onclick="populateTextBox('<%=CHKIIF(vURL_M="","",vURL_M)%>','m');">���þ���</li>
								<li onclick="populateTextBox('/category/category_itemPrd.asp?itemid=��ǰ�ڵ�','m');">/category/category_itemPrd.asp?itemid=��ǰ�ڵ�</li>
								<li onclick="populateTextBox('/category/category_main2020.asp?disp=ī�װ�','m');">/category/category_main2020.asp?disp=ī�װ�</li>
								<li onclick="populateTextBox('/brand/brand_detail2020.asp?brandid=�귣����̵�','m');">/brand/brand_detail2020.asp?brandid=�귣����̵�</li>
								<li onclick="populateTextBox('/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�','m');">/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�</li>
								<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�','m');">/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�</li>
								<li onclick="populateTextBox('/gift/gifttalk/','m');">����Ʈ</li>
								<li onclick="populateTextBox('/wish/index.asp','m');">����</li>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<th><div>�˻� Ű���� *</div></th>
					<td>
						<input type="text" class="formTxt" id="keyword" name="keyword" value="<%=vKwArr%>" placeholder="����ũ�� ������ �˻� Ű���带 ',(��ǥ)' �������� �Է����ּ���." style="width:99%" maxlength="200" />
						<input type="hidden" id="keyword_in_db" name="keyword_in_db" value="<%=vKwArr%>">
					</td>
				</tr>
				<tr>
					<th><div>���� �Ⱓ *</div></th>
					<td>
						<span><input type="hidden" id="termSet" name="viewgubun" value="<%=vViewGubun%>" /></span>
						<span>
							<input type="text" class="formTxt" id="sdate" name="sdate" value="<%=vSDate%>" style="width:100px" placeholder="������" maxlength="10" readonly />
							<img src="/images/admin_calendar.png" id="sdate_trigger" alt="�޷����� �˻�" />
							<script language="javascript">
								var CAL_Start = new Calendar({
									inputField : "sdate", trigger    : "sdate_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_End.args.min = date;
										CAL_End.redraw();
										this.hide();
									}, bottomBar: true, dateFormat: "%Y-%m-%d"
								});
							</script>
							<input type="text" class="formTxt" id="shhmmss" name="shhmmss" value="<%=vShhmmss%>" style="width:60px" maxlength="8" readonly />
							~
							<input type="text" class="formTxt" id="edate" name="edate" value="<%=vEDate%>" style="width:100px" placeholder="������" maxlength="10" readonly />
							<img src="/images/admin_calendar.png" id="edate_trigger" alt="�޷����� �˻�" />
							<script language="javascript">
								var CAL_End = new Calendar({
									inputField : "edate", trigger    : "edate_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_Start.args.max = date;
										CAL_Start.redraw();
										this.hide();
									}, bottomBar: true, dateFormat: "%Y-%m-%d"
								});
							</script>
							<input type="text" class="formTxt" id="ehhmmss" name="ehhmmss" value="<%=vEhhmmss%>" style="width:60px" maxlength="8" readonly />
						</span>
					</td>
				</tr>
				<tr>
					<th><div>��� ���� *</div></th>
					<td>
						<span class="rMar10"><input type="radio" id="useyny" name="useyn" value="y" <%=CHKIIF(vUseYN="y","checked","")%> /> <label for="useyny">�����</label></span>
						<span class="rMar10"><input type="radio" id="useynn" name="useyn" value="n" <%=CHKIIF(vUseYN="n","checked","")%> /> <label for="useynn">������</label></span>
					</td>
				</tr>
				<tr>
					<th><div>���</div></th>
					<td><textarea class="formTxtA" rows="6" style="width:99%;" id="memo" name="memo"><%=vMemo%></textarea></td>
				</tr>
				</tbody>
			</table>
		</div>
		<% If vQuickType = "nor" OR vQuickType = "set" OR vQuickType = "brd" Then '### �⺻��, ������, �귣���� �� %>
            <div class="pad20" style="display:none">
                <h3>- ����ũ �Ӽ� ����</h3>
                <table class="tbType1 writeTb tMar10">
                    <colgroup>
                        <col width="16%" /><col width="" />
                    </colgroup>
                    <tbody>
                    <tr>
                        <th><div>��ư��</div></th>
                        <td><input type="text" class="formTxt" id="btnname" name="btnname" value="<%=vBtnName%>" placeholder="��ư���� 25�� �̳��� �Է����ּ���" style="width:99%" maxlength="25" /></td>
                    </tr>
                    <tr>
                        <th><div>��ư URL</div></th>
                        <td>
                            <p>
                                <div class="selectLink">
                                    <input type="text" class="formTxt" value="<%=CHKIIF(vBtnPCLink="","",vBtnPCLink)%>" placeholder="��ư���� �̵��� PC URL�� �Է�(����)" onclick="showDropL('pc');" id="btnlinkpc" name="btnlinkpc" onkeyup="linkcopyL('pc');" maxlength="200" />
                                    <ul style="display:none;" id="selectbtnLinkpc">
                                        <li onclick="populateTextBoxL('<%=CHKIIF(vBtnPCLink="","",vBtnPCLink)%>','pc');">���þ���</li>
                                        <li onclick="populateTextBoxL('/category/category_prd.asp?itemid=��ǰ�ڵ�','pc');">/category/category_prd.asp?itemid=��ǰ�ڵ�</li>
                                        <li onclick="populateTextBoxL('/shopping/category_list.asp?disp=ī�װ�','pc');">/shopping/category_list.asp?disp=ī�װ�</li>
                                        <li onclick="populateTextBoxL('/street/street_brand.asp?makerid=�귣����̵�','pc');">/street/street_brand.asp?makerid=�귣����̵�</li>
                                        <li onclick="populateTextBoxL('/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�','pc');">/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�</li>
                                        <li onclick="populateTextBoxL('/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�','pc');">/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�</li>
                                        <li onclick="populateTextBoxL('/gift/gifttalk/','pc');">����Ʈ</li>
                                        <li onclick="populateTextBoxL('/wish/index.asp','pc');">����</li>
                                    </ul>
                                </div>
                            </p>
                            <p class="tPad05">
                                <div class="selectLink">
                                    <input type="text" class="formTxt" value="<%=CHKIIF(vBtnMLink="","",vBtnMLink)%>" placeholder="��ư���� �̵��� Mobile URL�� �Է�(����)" onclick="showDropL('m');" id="btnlinkm" name="btnlinkm" onkeyup="linkcopyL('m');" maxlength="200" />
                                    <ul style="display:none;" id="selectbtnLinkm">
                                        <li onclick="populateTextBoxL('<%=CHKIIF(vBtnMLink="","",vBtnMLink)%>','m');">���þ���</li>
                                        <li onclick="populateTextBoxL('/category/category_itemPrd.asp?itemid=��ǰ�ڵ�','m');">/category/category_itemPrd.asp?itemid=��ǰ�ڵ�</li>
                                        <li onclick="populateTextBoxL('/category/category_list.asp?disp=ī�װ�','m');">/category/category_list.asp?disp=ī�װ�</li>
                                        <li onclick="populateTextBoxL('/street/street_brand.asp?makerid=�귣����̵�','m');">/street/street_brand.asp?makerid=�귣����̵�</li>
                                        <li onclick="populateTextBoxL('/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�','m');">/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�</li>
                                        <li onclick="populateTextBoxL('/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�','m');">/culturestation/culturestation_event.asp?evt_code=��ó�����̼��̺�Ʈ�ڵ�</li>
                                        <li onclick="populateTextBoxL('/gift/gifttalk/','m');">����Ʈ</li>
                                        <li onclick="populateTextBoxL('/wish/index.asp','m');">����</li>
                                    </ul>
                                </div>
                            </p>
                        </td>
                    </tr>
                    <% If vQuickType = "set" OR vQuickType = "brd" Then '### ������, �귣���� �� %>
                    <tr>
                        <th><div>��ư Color</div></th>
                        <td>
                            <p class="tPad10" id="btncolorselect">
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBtnColor="BAD3E0","colorbtn","")%>" id="btncolor1" onClick="jsBGColor('btncolorselect','btn_color','btncolor1','BAD3E0');" style="background-color:#BAD3E0"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBtnColor="84adc2","colorbtn","")%>" id="btncolor2" onClick="jsBGColor('btncolorselect','btn_color','btncolor2','84adc2');" style="background-color:#84adc2"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBtnColor="9c7c6b","colorbtn","")%>" id="btncolor3" onClick="jsBGColor('btncolorselect','btn_color','btncolor3','9c7c6b');" style="background-color:#9c7c6b"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBtnColor="7a88b8","colorbtn","")%>" id="btncolor4" onClick="jsBGColor('btncolorselect','btn_color','btncolor4','7a88b8');" style="background-color:#7a88b8"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBtnColor="eff7fd","colorbtn","")%>" id="btncolor5" onClick="jsBGColor('btncolorselect','btn_color','btncolor5','eff7fd');" style="background-color:#eff7fd"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBtnColor="fceef2","colorbtn","")%>" id="btncolor6" onClick="jsBGColor('btncolorselect','btn_color','btncolor6','fceef2');" style="background-color:#fceef2"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBtnColor="e9f4ed","colorbtn","")%>" id="btncolor7" onClick="jsBGColor('btncolorselect','btn_color','btncolor7','e9f4ed');" style="background-color:#e9f4ed"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBtnColor="fbf3e7","colorbtn","")%>" id="btncolor8" onClick="jsBGColor('btncolorselect','btn_color','btncolor8','fbf3e7');" style="background-color:#fbf3e7"></button></span>
                                <span>#<input type="text" class="formTxt vTop" id="btn_color" name="btn_color" value="<%=vBtnColor%>" style="width:10%" maxlength="6" /></span>
                            </p>
                        </td>
                    </tr>
                    <% End If %>
                    <tr>
                        <th><div>����ũ��� ���� *</div></th>
                        <td>
                            <p>
                                <span class="rMar10"><input type="radio" id="oneClrUse" name="bggubun" value="c" <%=CHKIIF(vBgGubun="c","checked","")%> onClick="jsBgGubun('c');" /> <label for="oneClrUse">�ܻ� ���</label></span>
                                <span class="rMar10"><input type="radio" id="imgUse" name="bggubun" value="i" <%=CHKIIF(vBgGubun="i","checked","")%> onClick="jsBgGubun('i');" /> <label for="imgUse">�̹��� ���</label></span>
                            </p>
                            <p class="tPad10" id="bgcolorselect" style="display:<%=CHKIIF(vBgGubun="c","block","none")%>">
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="BAD3E0","colorbtn","")%>" id="color1" onClick="jsBGColor('bgcolorselect','bgcolor','color1','BAD3E0');" style="background-color:#BAD3E0"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="84adc2","colorbtn","")%>" id="color2" onClick="jsBGColor('bgcolorselect','bgcolor','color2','84adc2');" style="background-color:#84adc2"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="9c7c6b","colorbtn","")%>" id="color3" onClick="jsBGColor('bgcolorselect','bgcolor','color3','9c7c6b');" style="background-color:#9c7c6b"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="7a88b8","colorbtn","")%>" id="color4" onClick="jsBGColor('bgcolorselect','bgcolor','color4','7a88b8');" style="background-color:#7a88b8"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="eff7fd","colorbtn","")%>" id="color5" onClick="jsBGColor('bgcolorselect','bgcolor','color5','eff7fd');" style="background-color:#eff7fd"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="fceef2","colorbtn","")%>" id="color6" onClick="jsBGColor('bgcolorselect','bgcolor','color6','fceef2');" style="background-color:#fceef2"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="e9f4ed","colorbtn","")%>" id="color7" onClick="jsBGColor('bgcolorselect','bgcolor','color7','e9f4ed');" style="background-color:#e9f4ed"></button></span>
                                <span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="fbf3e7","colorbtn","")%>" id="color8" onClick="jsBGColor('bgcolorselect','bgcolor','color8','fbf3e7');" style="background-color:#fbf3e7"></button></span>
                                <!--<span>#<input type="text" class="formTxt vTop" id="bgcolor" name="bgcolor" value="<%=vBgColor%>" style="width:10%" maxlength="6" /></span>-->
                            </p>
                            <p class="tPad10" id="qbgimg" style="display:<%=CHKIIF(vBgGubun="i","block","none")%>">
                                <input type="button" value="PC ���ε�" onClick="jsUploadImg('qbgimgurlpc','qbgimgurlpcspan');" /><br /><br />
                                <span id="qbgimgurlpcspan" style="padding:5px 5px 5px 0;"><%
                                    If vBgImgPC <> "" Then
                                        Response.Write "<img src='" & vBgImgPC & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vBgImgPC & "');>"
                                        Response.Write "<a href=javascript:jsDelImg('qbgimgurlpc','qbgimgurlpcspan');><img src='/images/icon_delete2.gif' border='0'></a>"
                                    End If
                                %></span>
                                <input type="hidden" id="qbgimgurlpc" name="qbgimgurlpc" value="<%=vBgImgPC%>">
                                <br /><br />
                                <input type="button" value="Mobile ���ε�" onClick="jsUploadImg('qbgimgurlm','qbgimgurlmspan');" /><br /><br />
                                <span id="qbgimgurlmspan" style="padding:5px 5px 5px 0;"><%
                                    If vBgImgM <> "" Then
                                        Response.Write "<img src='" & vBgImgM & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vBgImgM & "');>"
                                        Response.Write "<a href=javascript:jsDelImg('qbgimgurlm','qbgimgurlmspan');><img src='/images/icon_delete2.gif' border='0'></a>"
                                    End If
                                %></span>
                                <input type="hidden" id="qbgimgurlm" name="qbgimgurlm" value="<%=vBgImgM%>">
                                <br /><span class="tPad10 fs11 cBl3">* 2Mb ������(1024x200������) png, jpg, gif���� �̹��������� �������ּ���.</span>
                            </p>
                        </td>
                    </tr>
                    <% If vQuickType = "set" OR vQuickType = "brd" Then '### ������, �귣���� �� %>
                    <tr>
                        <th><div>����ũ �̹��� *</div></th>
                        <td>
                            <p>
                                <span class="rMar10"><input type="radio" id="qimgUseY" name="qimg_useyn" value="y" <%=CHKIIF(vQImgUseYN="y","checked","")%> onClick="jsQimgGubun('y');" /> <label for="qimgUseY">���</label></span>
                                <span class="rMar10"><input type="radio" id="qimgUseN" name="qimg_useyn" value="n" <%=CHKIIF(vQImgUseYN="n","checked","")%> onClick="jsQimgGubun('n');" /> <label for="qimgUseN">������</label></span>
                            </p>
                            <p class="tPad10" id="qimg" style="display:<%=CHKIIF(vQImgUseYN="y","block","none")%>">
                                <input type="button" value="PC ���ε�" onClick="jsUploadImg('qimgurlpc','qimgurlpcspan');" /><br /><br />
                                <span id="qimgurlpcspan" style="padding:5px 5px 5px 0;"><%
                                    If vQImgPC <> "" Then
                                        Response.Write "<img src='" & vQImgPC & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vQImgPC & "');>"
                                        Response.Write "<a href=javascript:jsDelImg('qimgurlpc','qimgurlpcspan');><img src='/images/icon_delete2.gif' border='0'></a>"
                                    End If
                                %></span>
                                <input type="hidden" id="qimgurlpc" name="qimgurlpc" value="<%=vQImgPC%>">
                                <br /><br />
                                <input type="button" value="Mobile ���ε�" onClick="jsUploadImg('qimgurlm','qimgurlmspan');" /><br /><br />
                                <span id="qimgurlmspan" style="padding:5px 5px 5px 0;"><%
                                    If vQImgM <> "" Then
                                        Response.Write "<img src='" & vQImgM & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vQImgM & "');>"
                                        Response.Write "<a href=javascript:jsDelImg('qimgurlm','qimgurlmspan');><img src='/images/icon_delete2.gif' border='0'></a>"
                                    End If
                                %></span>
                                <input type="hidden" id="qimgurlm" name="qimgurlm" value="<%=vQImgM%>">
                                <br /><span class="tPad10 fs11 cBl3">* 2Mb ������(1024x200������) png, jpg, gif���� �̹��������� �������ּ���.</span>
                            </p>
                        </td>
                    </tr>
                    <% End If %>
                    </tbody>
                </table>
            </div>
		<% End If %>
		<% If vQuickType = "cus" Then '### Ŀ������ �� %>
		<div class="pad20">
			<h3>- ����ũ �Ӽ� ����</h3>
			<table class="tbType1 writeTb tMar10">
				<colgroup>
					<col width="16%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>HTML</div></th>
					<td><textarea class="formTxtA" rows="25" style="width:99%;" id="htmlcont" name="htmlcont"><%=vHtmlCont%></textarea></td>
				</tr>
				</tbody>
			</table>
		</div>
		<% End If %>
		<div class="pad20">
			<div class="ct">
				<input type="button" value="����" onclick="jsQuickLinkSave();" class="cRd1" style="width:100px; height:30px;" />
				<input type="button" value="���" onclick="window.close();" style="width:100px; height:30px;" />
			</div>
		</div>
	</form>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->