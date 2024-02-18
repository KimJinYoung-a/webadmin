<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��ǰ�˻�
' History : 2009.04.07 ������ ����
'			2012.08.29 �ѿ�� ����
'####################################################
%>
<% If request.cookies("commonpop")("islogics") <> "ok" Then %>
<%'<!-- #include virtual="/admin/incSessionAdmin.asp" -->%>
<% server.Execute("/admin/incSessionAdmin.asp") %>
<% End If %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itembarcode/totalitembarcodeCls.asp" -->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->

<%

	Dim oitembar, itembarcode, siteSeq, itemgubun, itemid, itemoption
	itembarcode = requestCheckVar(request("itemcode"),32)
	If itembarcode = "" Then
		itembarcode = requestCheckVar(request("itembarcode"),32)
	End If


	If Len(itembarcode) <= 8 AND itembarcode <> "" and IsNumeric(itembarcode) Then
		'// ��ǰ�ڵ�
		siteSeq = "10"
		itemgubun = "10"
		itemid = BF_GetFormattedItemId(itembarcode)
		itemoption = "0000"
		itembarcode = itemgubun & itemid & itemoption
	Else
		if BF_IsMaybeTenBarcode(itembarcode) then
			siteSeq 	= BF_GetItemGubun(itembarcode)
			itemgubun 	= BF_GetItemGubun(itembarcode)
			itemid 		= BF_GetItemId(itembarcode)
			itemoption 	= BF_GetItemOption(itembarcode)
		ElseIf Len(itembarcode) > 7 Then
			Set oitembar = new CTotalItemBarCode
			oitembar.FRectBarcode = itembarcode
			oitembar.getTotalItemCodeSearch

			siteSeq 	= oitembar.FOneItem.FsiteSeq
			itemgubun 	= oitembar.FOneItem.FsiteItemGubun
			itemid		= oitembar.FOneItem.FsiteItemid
			itemoption	= oitembar.FOneItem.FsiteItemOption

			itembarcode = itemgubun & BF_GetFormattedItemId(itemid) & itemoption
			If itembarcode = "" Then
				itembarcode = request("itemcode")
			End If
			Set oitembar = Nothing
		End If
	End If


	'IF itembarcode <> "" AND IsNumeric(itembarcode) = False Then
	'	rw "<script>alert('�߸��� �����Դϴ�.');history.back();</script>"
	'	Response.End
	'End If


	Set oitembar = new CTotalItemBarCode

	oitembar.FRectSiteSeq = siteSeq
	oitembar.FRectItemgubun = itemgubun

	oitembar.FRectItemID = itemid
	'oitembar.FRectItemoption = itemoption

	If itemid<>"" Then
		If (CStr(siteSeq) = "10") Then
			oitembar.getTotalItemBarcodeON
		Else
			oitembar.getTotalItemBarcodeOFF
		End If
	End If

	Dim i, vRegCount
	vRegCount = 0
	For i=0 To oitembar.FResultCount-1
		If oitembar.FItemList(i).FPublicBarcode <> "" Then
			vRegCount = vRegCount + 1
		End If
	Next
%>
<script type='text/javascript'>
var pIdx = 0;

//���ڵ����
function barcodeManageRe(itemcode)
{
	var popbarcodemanageRe = window.open('/admin/stock/popBarcodeManage.asp?itemcode=' + itemcode,'popbarcodemanageRe','width=550,height=400,resizable=yes,scrollbars=yes');
	popbarcodemanageRe.focus();
}

function InputRackcode(frm){
	if (frm.itemrackcode.value.length!=4){
		alert("��ǰ ���ڵ带 ��Ȯ�� �Է��ϼ���. 4�ڸ�");
		frm.itemrackcode.focus();
		return;
	}

	if (confirm("��ǰ ���ڵ带 �����Ͻðڽ��ϱ�?")){
		frm.submit();
	}
}

function Research(frm){
	if(document.frmbar.itembarcode.value == "")
	{
		alert("���ڵ带 ��Ȯ�� �Է��ϼ���.");
		document.frmbar.itembarcode.focus();
		return;
	}
	document.location.href = "<%=CurrURL()%>?itemcode="+document.frmbar.itembarcode.value+"";
}

function InputBarcode(){
	var inputcount = 0;

	<% if oitembar.FResultCount>0 then %>
		<% if oitembar.FResultCount=1 then %>
			if(document.frmbar.publicbar.value != "")
			{
				if (document.frmbar.publicbar.value.length<8){
					alert("���ڵ带 ��Ȯ�� �Է��ϼ���.");
					document.frmbar.publicbar.focus();
					return;
				}

				inputcount = inputcount + 1;
			}
			if(inputcount < 1)
			{
				alert("���ڵ带 �Է��ؾ� �մϴ�.");
				document.frmbar.publicbar.focus();
				return;
			}
		<% else %>
			for(var i=0; i<<%=oitembar.FResultCount%>; i++)
			{
				if(document.frmbar.publicbar[i].value != "")
				{
					if (document.frmbar.publicbar[i].value.length<8){
						alert("���ڵ带 ��Ȯ�� �Է��ϼ���.");
						document.frmbar.publicbar[i].focus();
						return;
					}

					inputcount = inputcount + 1;
				}
			}

			if(inputcount < 1)
			{
				// 2016-01-26, skyer9
				alert("���ڵ带 �Է��ؾ� �մϴ�.");
				// alert("���ڵ带 �ּ� 2�� �̻��� �Է��ؾ� �մϴ�.");
				for(var i=0; i<<%=oitembar.FResultCount%>; i++)
				{
					if(document.frmbar.publicbar[i].value == "")
					{
						document.frmbar.publicbar[i].focus();
						break;
					}
				}
				return;
			}
		<% end if %>
	<% else %>
		alert("�����ڵ� �˻� �� ������ڵ带 ����ϼ���.");
		document.frmbar.itembarcode.focus();
		return;
	<% end if %>


	if (document.frmbar.itemid.value.length<1){
		alert("�����ڵ� �˻� �� ������ڵ带 ����ϼ���.");
		document.frmbar.itembarcode.focus();
		return;
	}

	//return;
	if (confirm("���� ���ڵ带 �����Ͻðڽ��ϱ�?")){
		document.frmbar.submit();
	}
}

function GetOnLoad(){
<% If Request("isok") <> "o" Then %>
	<% if oitembar.FResultCount>0 then %>
	    <% if oitembar.FResultCount=1 then %>
    		document.frmbar.publicbar.focus();
    		document.frmbar.publicbar.select();
    	<% else %>
    		eval("document.frmbar.publicbar["+pIdx+"]").focus();
    		eval("document.frmbar.publicbar["+pIdx+"]").select();
    	<% end if %>
	<% else %>
	document.frmbar.itembarcode.focus();
	<% end if %>
<% end if %>
}

function jsNextFocus(i)
{
	var isFull = "x";
	if("<%=oitembar.FResultCount%>" == i)
	{
		InputBarcode();
	}
	else
	{
		for(var i=0; i<<%=oitembar.FResultCount%>; i++)
		{
			if(document.frmbar.publicbar[i].value == "")
			{
				document.frmbar.publicbar[i].focus();
				break;
			}

			if(i == <%=oitembar.FResultCount-1%>)
			{
				isFull = "o";
			}
		}

		if(isFull == "o")
		{
			document.frmbar.publicbar[i-1].focus();
		}
	}
}

function FocusAndSelect(frm, obj){
	obj.focus();
	obj.select();
}

function jsDeleteBarcode(itemcode)
{
	document.frmbarperone.action.value = "delete";
	document.frmbarperone.itemcode.value = itemcode;
	document.frmbarperone.submit();
}

function jsMessageReset()
{
	<% For i=0 To oitembar.FResultCount-1 %>
	document.getElementById("publicbarspan<%= oitembar.FItemList(i).FsiteItemOption %>").innerHTML = "";
	<% Next %>
	document.getElementById("notregmessage").innerHTML = "";
}
window.onload=GetOnLoad;
</script>
<table width="512" height="220" border="0" align="left" cellpadding="0" cellspacing="0" class="a">
<tr valign="bottom">
	<td width="10" height="10" align="right" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif" bgcolor="#F3F3FF"></td>
	<td width="10" height="10" align="left" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td height="20" background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
	<td height="20" background="/images/tbl_blue_round_06.gif" bgcolor="#F3F3FF"><img src="/images/icon_star.gif" align="absbottom"><font color="red">&nbsp;<strong>��ǰ ������ڵ� �Է�</strong></font></td>
	<td height="20" background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
</tr>
<form name="frmbar" method="post" action="popBarcodeManage_proc.asp" target="itembarcodeframe">
<input type="hidden" name="itemcode" value="<%= itembarcode %>">
<input type="hidden" name="optioncount" value="<%=oitembar.FResultCount%>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<% if oitembar.FResultCount>0 then %>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border=0 cellspacing=0 cellpadding=2 class=a bgcolor="#F3F3FF">
		<tr>
			<td width="220">�귣�� : <%= oitembar.FItemList(0).FbrandName %> (<%= oitembar.FItemList(0).Fmakerid %>)</td>
			<td>���ڵ�:&nbsp;
				<input type="text" name="itembarcode" value="<%= itembarcode %>" size=20 AUTOCOMPLETE="off" onFocus="FocusAndSelect(frmbar, frmbar.itembarcode);" onKeyPress="if (event.keyCode == 13){ Research(); return false;}">&nbsp;
				<input type=button value="�˻�" onclick="Research(frmbar)" >
			</td>
		</tr>
		<tr><td colspan=2>��ǰ�� : <%= oitembar.FItemList(0).FsiteItemName %></td></tr>
		<tr><td colspan=2>��ǰ���� : <%= FormatNumber(oitembar.FItemList(0).Forgsellprice,0) %></td></tr>
		<tr><td height="5" colspan=2></td></tr>
		</table>

		<table border=0 cellspacing=0 cellpadding=2 class=a>
		<tr><td height="30" colspan="10">&nbsp;�� �ϳ��� �ɼǿ� �ϳ��� ���ڵ常 ��� ����. �ߺ� ���ڵ� ��� �Ұ�.</td></tr>
		<tr>
			<td width="110" align="center" valign="top"><img src="<%= oitembar.FItemList(0).FImageList %>"></td>
			<td width="370" valign="top">
				<table width="370" border="0" cellspacing="0" cellpadding="1" class="a">
				<% for i=0 to oitembar.FResultCount-1 %>
				<tr>
					<% if oitembar.FItemList(i).FsiteItemOptionName="" then %>
					<td width="210" height="25" align="center">�ɼǾ���</td>
					<% else %>
						<% if itemoption=oitembar.FItemList(i).FsiteItemOption then %>
						<td width="210" height="25" bgcolor="#F0F0F0"><script >pIdx=<%= i %>;</script><b>[<%=oitembar.FItemList(i).FsiteItemOption%>]<%= oitembar.FItemList(i).FsiteItemOptionName %></b><%=oitembar.FItemList(i).FOptaddprice%></td>
						<% else %>
						<td width="210" height="25">[<%=oitembar.FItemList(i).FsiteItemOption%>]<%= oitembar.FItemList(i).FsiteItemOptionName %><%=oitembar.FItemList(i).FOptaddprice%></td>
						<% end if %>
					<% end if %>
					<td align="right" width="160">
						<input type="hidden" name="itemoption" value="<%= oitembar.FItemList(i).FsiteItemOption %>">
						<input type="text" name="publicbar" value="<%= oitembar.FItemList(i).FPublicBarcode %>" size=20 maxlength=20 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ InputBarcode(); return false;}">
					</td>
					<td>
					<% If oitembar.FItemList(i).FPublicBarcode <> "" Then %>
						<input type="button" class="button" value="X" onClick="jsDeleteBarcode('<%= itemgubun & itemid & oitembar.FItemList(i).FsiteItemOption %>');">
					<% End If %>
					</td>
				</tr>
				<tr>
					<td colspan="3"><span id="publicbarspan<%= oitembar.FItemList(i).FsiteItemOption %>"></span></td>
				</tr>
				<% next %>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="10" align="right"><span id="notregmessage"></span></td>
		</tr>
		<tr>
			<td height="60" align="right" colspan="10">
				<table class="a">
				<tr>
					<td><br>
						�� ���ڵ彺ĳ�ʷ� �Է½� ��� �Է��ϸ� �ڵ����� ��Ϲ�ư Ŭ����.&nbsp;<br>
						�� ���ڵ�� ������ 000000000100 ���ڵ带 ��ĵ�ϸ� ��.
					</td>
					<td><input type="button" class="button" value="��   ��" style="width:100px;height:50px;" onclick="InputBarcode()">&nbsp;</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td height="10" colspan="10"></td>
		</tr>
		</table>
	</td>
	<td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
</tr>
<% else %>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border=0 cellspacing=0 cellpadding=2 class=a bgcolor="#F3F3FF">
		<tr>
			<td width="220"></td>
			<td>���ڵ�:&nbsp;
				<input type="text" name="itembarcode" value="<%= itembarcode %>" size=20 AUTOCOMPLETE="off" onFocus="FocusAndSelect(frmbar, frmbar.itembarcode);" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">&nbsp;
				<input type=button value="�˻�" onclick="Research(frmbar)" >
				<br>&nbsp;
			</td>
		</tr>
		</table>
		<table width="100%" border=0 cellspacing=0 cellpadding=2 class=a>
		<tr>
			<td align=center valign=center><br>�˻��� ���� �����ϴ�.<br>���ڵ� �Է��� �˻��ϼ���.</td>
		</tr>
		</table>
	</td>
	<td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
</tr>
<% end if %>
<tr valign="top" bgcolor="#F3F3FF">
	<td height="10" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td height="10" background="/images/tbl_blue_round_08.gif"></td>
	<td height="10"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
</form>
<form name="frmbarperone" method="post" action="popBarcodeManage_proc.asp" target="itembarcodeframe">
<input type="hidden" name="action" value="">
<input type="hidden" name="itemcode" value="">
</form>
<iframe src="about:blank" id="itembarcodeframe" name="itembarcodeframe" width="0" height="0"></iframe>
<% If Request("isok") = "o" Then %>
<script type='text/javascript'>
FocusAndSelect(frmbar, frmbar.itembarcode);
document.getElementById('notregmessage').innerHTML = '<font color=blue>* ����Ϸ�.</font>';
</script>
<%
End If

set oitembar = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
