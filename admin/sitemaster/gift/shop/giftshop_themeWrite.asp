<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftShop_cls.asp" -->
<%
'###############################################
' Discription : GIFT SHOP �׸� ��ǰ ����
' History : 2014.04.07 ������ : �ű� ����
'###############################################

	'// ���� ����
	Dim oGiftShop, i
	Dim themeIdx, subject, subDesc, userid, regdate, frontItemid, isOpen, isPick, isUsing, tag, sortNo
	Dim viewCount, commentCount, pickImage

	'// �Ķ���� ����
	themeIdx = getNumeric(requestCheckVar(request("themeIdx"),10))

	'// �׸� ���� ����
	if themeIdx<>"" then
		Set oGiftShop = new CGiftShop
		oGiftShop.FRectIdx = themeIdx
		oGiftShop.GetThemeInfo
		if oGiftShop.FResultCount>0 then
			subject		= oGiftShop.FOneItem.Fsubject
			subDesc		= oGiftShop.FOneItem.FsubDesc
			userid		= oGiftShop.FOneItem.Fuserid
			regdate		= oGiftShop.FOneItem.Fregdate
			frontItemid	= oGiftShop.FOneItem.FfrontItemid
			isOpen		= oGiftShop.FOneItem.FisOpen
			isPick		= oGiftShop.FOneItem.FisPick
			isUsing		= oGiftShop.FOneItem.FisUsing
			sortNo		= oGiftShop.FOneItem.FsortNo
			tag			= oGiftShop.FOneItem.Ftag
			viewCount	= oGiftShop.FOneItem.FviewCount
			commentCount	= oGiftShop.FOneItem.FcommentCount
			pickImage	= oGiftShop.FOneItem.FpickImage
		end if
		Set oGiftShop = Nothing
	end if

	if sortNo="" then sortNo=0
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function(){
	$(".chkBox").buttonset().children().next().attr("style","font-size:11px;");
	$(".btn").button();
});

function fnSelKeyword(elm){
	if($("#lyrKeyword input[type='checkbox']:checked").length>3) {
		alert("Ű����� 3������ ���� �����մϴ�.");
		$(elm).attr("checked",false);
	}
	var keyId="";
	$("#lyrKeyword input[type='checkbox']:checked").each(function(){
		if(keyId!="") keyId += ",";
		keyId += $(this).val();
	});
	document.frm.arrKeyIdx.value = keyId;
}

function SaveTheme(frm){
	if(frm.subject.value=="") {
		alert("�׸� ������ �Է����ּ���.");
		frm.subject.focus();
		return;
	}

	if(frm.sortNo.value=="") {
		alert("�׸� ���� �켱������ �Է����ּ���.");
		frm.sortNo.focus();
		return;
	}

	if(!$("input[type='checkbox']").is(":checked")) {
		alert("Ű���带 �������ּ���.");
		return;
	}

	if(frm.isOpen.value=="Y"&&$("#itemList input[name='itemid']").length<4) {
		alert("��ϵ� ��ǰ�� 4�� �̸��� ��쿡�� ���������� �� �� �����ϴ�.");
		return;
	}

	if(frm.isOpen.value=="Y"&&frm.frontItemid.value=="0") {
		alert("��ǥ ��ǰ�� �������ּ���.");
		return;
	}

	frm.submit();
}

function fnChkAll(elm) {
	$("#itemList input[name='itemid']").attr("checked",$(elm).is(":checked"));
}

function fnChkDelete() {
	var arrIID="";
	if(!$("#itemList input[name='itemid']").is(":checked")) {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}
	$("#itemList input[name='itemid']:checked").each(function(){
		if(arrIID!="") arrIID += ",";
		arrIID += $(this).val();
	});
	
	window.open("/admin/sitemaster/gift/shop/doRegItemCdArray.asp?themeIdx=<%=themeIdx%>&mode=d&subItemidArray="+arrIID, "popup_item", "width=300,height=200,scrollbars=yes,resizable=yes");
}

// ��ǰ�˻� �ϰ� ���
function popRegSearchItem() {
    var acUrl = encodeURIComponent("/admin/sitemaster/gift/shop/doRegItemCdArray.asp?themeIdx=<%=themeIdx%>&mode=i");
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL="+acUrl, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
}

function jsSetImg(sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('pop_theme_upload.asp?yr=<%=Year(regdate)%>&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   $("#"+sName).val('');
	   $("#"+sSpan).fadeOut();
	}
}

function fnChgIsPick(elm) {
	if($(elm).val()=="Y") {
		$("#rowTTImg").show();
	} else {
		$("#rowTTImg").hide();
	}
}

function fnChkFrontItem(iid) {
	document.frm.frontItemid.value=iid;
}
</script>
<!-- ���������� ���� ���� -->
<form name="frm" method="POST" action="doGiftShopTheme.asp" style="margin:0;">
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<input type="hidden" name="frontItemid" value="<%= frontItemid %>" />
<input type="hidden" name="arrKeyIdx" value="<%= tag %>" />
<input type="hidden" name="mode" value="<%=chkIIF(themeIdx="","i","u")%>" />
<p><b>�� �׸� ����</b></p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<colgroup>
	<col width="120" />
	<col width="*" />
	<col width="120" />
	<col width="*" />
</colgroup>
<% if themeIdx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�׸� ��ȣ</td>
    <td>
        <%=themeIdx %>
        <input type="hidden" name="themeIdx" value="<%=themeIdx %>" />
    </td>
    <td bgcolor="#DDDDFF">����Ͻ�</td>
    <td>
        <%=regdate %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��ȸ��</td>
    <td>
        <%=viewCount %>
    </td>
    <td bgcolor="#DDDDFF">��ۼ�</td>
    <td>
        <%=commentCount %>
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�׸� ���� <span style="color:#F03030" title="�ʼ�">��</span></td>
    <td colspan="3">
		<input type="text" name="subject" size="24" maxlength="18" value="<%=subject%>" class="text" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�ΰ� ����</td>
    <td colspan="3">
		<input type="text" name="subDesc" size="60" maxlength="40" value="<%=subDesc%>" class="text" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��������</td>
    <td>
		<select name="isOpen" class="select">
		<option value="Y" <%=chkIIF(isOpen="Y","selected","")%>>����</option>
		<option value="N" <%=chkIIF(isOpen="N" or isOpen="","selected","")%>>�����</option>
		</select>
    </td>
    <td bgcolor="#DDDDFF">��������</td>
    <td>
		<select name="isPick" class="select" onchange="fnChgIsPick(this)">
		<option value="Y" <%=chkIIF(isPick="Y" or isPick="","selected","")%>>10x10's Pick</option>
		<option value="N" <%=chkIIF(isPick="N","selected","")%>>User's Pick</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�켱���� <span style="color:#F03030" title="�ʼ�">��</span></td>
    <td>
		<input type="text" name="sortNo" size="4" value="<%=sortNo%>" class="text" />
    </td>
    <td bgcolor="#DDDDFF">��뿩��</td>
    <td>
		<select name="isUsing" class="select">
		<option value="Y" <%=chkIIF(isUsing="Y" or isUsing="","selected","")%>>���</option>
		<option value="N" <%=chkIIF(isUsing="N","selected","")%>>����</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�׸� Ű���� <span style="color:#F03030" title="�ʼ�">��</span></td>
    <td colspan="3" id="lyrKeyword">
		<%=getGiftKeyword("fnSelKeyword(this)",tag)%>
    </td>
</tr>
<% if (isPick="Y" or themeIdx="") or date<"2014-04-15" then %>
<tr bgcolor="#FFFFFF" id="rowTTImg" style="<%=chkIIF(isPick="Y" or themeIdx="","","display:none;")%>">
    <td bgcolor="#DDDDFF">Ÿ��Ʋ �̹���</td>
    <td colspan="3">
		<input type="hidden" name="pickImage" id="pickImage" value="<%=pickImage%>">
		<input type="button" value="�̹��� ���" onClick="jsSetImg('<%=pickImage%>','pickImage','lyTitleImg')" class="button">
		<span style="color:#A06060; font-size:11px;">�� 1100 �� 170px (200kb������ JPEG, GIF, PNG)</span>
		<div id="lyTitleImg" style="padding: 5 5 5 5">
		<% if Not(pickImage="" or isNull(pickImage)) then %>
			<img src="<%=pickImage%>" width="100%">
			<a href="javascript:jsDelImg('pickImage','lyTitleImg');"><img src="/images/icon_delete2.gif" border="0"></a>
		<% end if %>
		</div>
    </td>
</tr>
<% end if %>
<tr bgcolor="#F8F8F8">
    <td colspan="4" align="center">
    	<input type="button" value=" ��� " onClick="history.back();" class="btn"> &nbsp;
    	<input type="button" value=" �� �� " onClick="SaveTheme(this.form);" class="btn">
    </td>
</tr>
</table>
</form>
<%
	'// ��ϵ� �׸����
	if themeIdx<>"" then
		Set oGiftShop = new CGiftShop
		oGiftShop.FPageSize=200
		oGiftShop.FRectIdx = themeIdx
		oGiftShop.GetThemeItemList
%>
<p><b>�� ��ǰ ����</b></p>
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="7">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	�� <%=oGiftShop.FTotalCount%> �� /
		    	<input type="button" value="����" class="button" onClick="fnChkDelete()" />
		    </td>
		    <td align="right">
		    	<input type="button" value="��ǰ �߰�" class="button" onClick="popRegSearchItem()" />
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="30" />
<col width="90" />
<col width="70" />
<col width="*" />
<col width="110" />
<col width="80" />
<col width="80" />
<tr align="center" bgcolor="#DDDDFF">
    <td><input type="checkbox" name="chkALL" value="all" onclick="fnChkAll(this)"></td>
    <td>��ǰ�ڵ�(��ǥ)</td>
    <td>�̹���</td>
    <td>��ǰ��</td>
    <td>�ǸŰ�</td>
    <td>ǰ������</td>
    <td>�����</td>
</tr>
<tbody id="itemList">
<%	For i=0 to oGiftShop.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><input type="checkbox" name="itemid" value="<%=oGiftShop.FItemList(i).Fitemid%>"></td>
    <td>
    	<label>
    	<input type="radio" name="chkFront" value="<%=oGiftShop.FItemList(i).Fitemid%>" <%=chkIIF(oGiftShop.FItemList(i).Fitemid=frontItemid,"checked","")%> onclick="fnChkFrontItem(this.value)">
    	<%=oGiftShop.FItemList(i).Fitemid%>
    	</label>
    </td>
    <td><img src="<%=oGiftShop.FItemList(i).FsmallImage%>"></td>
    <td align="left">
    	<font color="#606060">[<%=oGiftShop.FItemList(i).Fbrandname%>]</font>
    	<%=oGiftShop.FItemList(i).Fitemname%>
    </td>
    <td><%=FormatNumber(oGiftShop.FItemList(i).FsellCash,0)%>��</td>
    <td><%=oGiftShop.FItemList(i).isSoldOut%></td>
    <td><%=Left(oGiftShop.FItemList(i).Fregdate,10)%></td>
</tr>
<%	Next %>
</tbody>
</table>
</form>
<%
		Set oGiftShop = Nothing
	end if
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->