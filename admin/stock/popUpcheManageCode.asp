<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  상품검색
' History : 2009.04.07 서동석 생성
'			2012.08.29 한용민 수정
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
	itembarcode = request("itemcode")
	If itembarcode = "" Then
		itembarcode = request("itembarcode")
	End If


	If Len(itembarcode) <= 8 AND itembarcode <> "" and IsNumeric(itembarcode) Then
		'// 상품코드
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


	IF itembarcode <> "" AND IsNumeric(itembarcode) = False Then
		rw "<script>alert('잘못된 접근입니다.');history.back();</script>"
		Response.End
	End If


	Set oitembar = new CTotalItemBarCode

	oitembar.FRectSiteSeq = siteSeq
	oitembar.FRectItemgubun = itemgubun

	oitembar.FRectItemID = itemid
	'oitembar.FRectItemoption = itemoption

	If itemid<>"" Then
		If (CStr(siteSeq) = "10") Then
			oitembar.getTotalUpcheManagecodeON
		Else
			oitembar.getTotalUpcheManagecodeOFF
		End If
	End If

	Dim i, vRegCount
	vRegCount = 0
	For i=0 To oitembar.FResultCount-1
		If oitembar.FItemList(i).Fupchemanagecode <> "" Then
			vRegCount = vRegCount + 1
		End If
	Next
%>

<script type='text/javascript'>
var pIdx = 0;

function Research(frm){
	if(document.frmbar.itembarcode.value == "")
	{
		alert("업체코드를 정확히 입력하세요.");
		document.frmbar.itembarcode.focus();
		return;
	}
	document.location.href = "<%=CurrURL()%>?itemcode="+document.frmbar.itembarcode.value+"";
}

function GetOnLoad(){
<% If Request("isok") <> "o" Then %>
	<% if oitembar.FResultCount>0 then %>
	    <% if oitembar.FResultCount=1 then %>
    		document.frmbar.upchemanagecode.focus();
    	<% else %>
    		eval("document.frmbar.upchemanagecode["+pIdx+"]").focus();
    		eval("document.frmbar.upchemanagecode["+pIdx+"]").select();
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
		InputUpcheManageCode();
	}
	else
	{
		for(var i=0; i<<%=oitembar.FResultCount%>; i++)
		{
			if(document.frmbar.upchemanagecode[i].value == "")
			{
				document.frmbar.upchemanagecode[i].focus();
				break;
			}

			if(i == <%=oitembar.FResultCount-1%>)
			{
				isFull = "o";
			}
		}

		if(isFull == "o")
		{
			document.frmbar.upchemanagecode[i-1].focus();
		}
	}
}

function InputUpcheManageCode(){
	var inputcount = 0;

	<% If oitembar.FResultCount < 2 Then %>
	//alert("옵션이 없는것은 업체코드 저장할 수 없습니다.");
	//return;
	<% End If %>

	<% if oitembar.FResultCount>0 then %>
		<% if oitembar.FResultCount=1 then %>
			if(document.frmbar.upchemanagecode.value != "")
			{
				inputcount = inputcount + 1;
			}
			if(inputcount < 1)
			{
				alert("업체코드를 입력해야 합니다.");
				document.frmbar.upchemanagecode.focus();
				return;
			}
		<% else %>
			for(var i=0; i<<%=oitembar.FResultCount%>; i++)
			{
				if(document.frmbar.upchemanagecode[i].value != "")
				{
					inputcount = inputcount + 1;
				}
			}

			if(inputcount < 2)
			{
				alert("업체코드를 최소 2개 이상은 입력해야 합니다.");
				for(var i=0; i<<%=oitembar.FResultCount%>; i++)
				{
					if(document.frmbar.upchemanagecode[i].value == "")
					{
						document.frmbar.upchemanagecode[i].focus();
						break;
					}
				}
				return;
			}
		<% end if %>
	<% else %>
		alert("물류코드 검색 후 범용바코드를 등록하세요.");
		document.frmbar.itembarcode.focus();
		return;
	<% end if %>


	if (document.frmbar.itemid.value.length<1){
		alert("물류코드 검색 후 범용바코드를 등록하세요.");
		document.frmbar.itembarcode.focus();
		return;
	}

	//return;
	if (confirm("업체코드를 저장하시겠습니까?")){
		document.frmbar.submit();
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

<form name="frmbar" method="post" action="popUpcheManageCode_proc.asp" target="itemmanamecodeframe" style="margin:0px;">
<input type="hidden" name="itemcode" value="<%= itembarcode %>">
<input type="hidden" name="optioncount" value="<%=oitembar.FResultCount%>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<table width="512" height="220" border="0" align="left" cellpadding="0" cellspacing="0" class="a">
<tr valign="bottom">
	<td width="10" height="10" align="right" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif" bgcolor="#F3F3FF"></td>
	<td width="10" height="10" align="left" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td height="20" background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
	<td height="20" background="/images/tbl_blue_round_06.gif" bgcolor="#F3F3FF"><img src="/images/icon_star.gif" align="absbottom"><font color="red">&nbsp;<strong>업체코드 입력</strong></font></td>
	<td height="20" background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
</tr>
<% if oitembar.FResultCount>0 then %>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border=0 cellspacing=0 cellpadding=2 class=a bgcolor="#F3F3FF">
		<tr>
			<td width="220">브랜드 : <%= oitembar.FItemList(0).FbrandName %> (<%= oitembar.FItemList(0).Fmakerid %>)</td>
			<td>바코드:&nbsp;
				<input type="text" name="itembarcode" value="<%= itembarcode %>" size=20 AUTOCOMPLETE="off" onFocus="FocusAndSelect(frmbar, frmbar.itembarcode);" onKeyPress="if (event.keyCode == 13){ Research(); return false;}">&nbsp;
				<input type=button value="검색" onclick="Research(frmbar)" >
			</td>
		</tr>
		<tr><td colspan=2>상품명 : <%= oitembar.FItemList(0).FsiteItemName %></td></tr>
		<tr><td colspan=2>상품가격 : <%= FormatNumber(oitembar.FItemList(0).Forgsellprice,0) %></td></tr>
		<tr><td height="5" colspan="2" align="right" style="padding-right:50px;"><font color="red">&nbsp;<strong>※ 업체코드 입력</strong></font></td></tr>
		</table>

		<table border=0 cellspacing=0 cellpadding=2 class=a>
		<tr><td height="10" colspan="10"></td></tr>
		<tr>
			<td width="110" align="center" valign="top"><img src="<%= oitembar.FItemList(0).FImageList %>"></td>
			<td width="370" valign="top">
				<table width="370" border="0" cellspacing="0" cellpadding="1" class="a">
				<% for i=0 to oitembar.FResultCount-1 %>
				<tr>
					<% if oitembar.FItemList(i).FsiteItemOptionName="" then %>
					<td width="210" height="25" align="center">옵션없음</td>
					<% else %>
						<% if itemoption=oitembar.FItemList(i).FsiteItemOption then %>
						<td width="210" height="25" bgcolor="#F0F0F0"><script >pIdx=<%= i %>;</script><b>[<%=oitembar.FItemList(i).FsiteItemOption%>]<%= oitembar.FItemList(i).FsiteItemOptionName %></b><%=oitembar.FItemList(i).FOptaddprice%></td>
						<% else %>
						<td width="210" height="25">[<%=oitembar.FItemList(i).FsiteItemOption%>]<%= oitembar.FItemList(i).FsiteItemOptionName %><%=oitembar.FItemList(i).FOptaddprice%></td>
						<% end if %>
					<% end if %>
					<td align="right" width="160">
						<input type="hidden" name="itemoption" value="<%= oitembar.FItemList(i).FsiteItemOption %>">
						<input type="text" name="upchemanagecode" value="<%= oitembar.FItemList(i).Fupchemanagecode %>" size=25 maxlength=32 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ jsNextFocus('<%=i+1%>'); return false;}">
					</td>
					<td>
					<% If oitembar.FItemList(i).Fupchemanagecode <> "" Then %>
						<input type="button" class="button" value="X" onClick="jsDeleteBarcode('<%= itemgubun & CHKIIF(ItemId>=1000000, Format00(8,ItemId), Format00(6,ItemId)) & oitembar.FItemList(i).FsiteItemOption %>');">
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
						※ 바코드스캐너로 입력시 모두 입력하면 자동으로 등록버튼 클릭됨.&nbsp;<br>
						※ 바코드로 삭제시 000000000100 바코드를 스캔하면 됨.
					</td>
					<td><input type="button" class="button" value="등   록" style="width:100px;height:50px;" onclick="InputUpcheManageCode()">&nbsp;</td>
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
			<td>바코드:&nbsp;
				<input type="text" name="itembarcode" value="<%= itembarcode %>" size=20 AUTOCOMPLETE="off" onFocus="FocusAndSelect(frmbar, frmbar.itembarcode);" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">&nbsp;
				<input type=button value="검색" onclick="Research(frmbar)" >
				<br>&nbsp;
			</td>
		</tr>
		</table>
		<table width="100%" border=0 cellspacing=0 cellpadding=2 class=a>
		<tr>
			<td align=center valign=center><br>검색된 값이 없습니다.<br>바코드 입력후 검색하세요.</td>
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
<form name="frmbarperone" method="post" action="popUpcheManageCode_proc.asp" target="itemmanamecodeframe">
<input type="hidden" name="action" value="">
<input type="hidden" name="itemcode" value="">
</form>
<iframe src="about:blank" id="itemmanamecodeframe" name="itemmanamecodeframe" width="0" height="0"></iframe>
<% If Request("isok") = "o" Then %>
<script type='text/javascript'>
FocusAndSelect(frmbar, frmbar.itembarcode);
document.getElementById('notregmessage').innerHTML = '<font color=blue>* 저장완료.</font>';
</script>
<%
End If
set oitembar = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
