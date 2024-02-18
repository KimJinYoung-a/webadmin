<%@ language=vbscript %>
<% option explicit
	Response.Expires = -1440
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
%>
<%
'####################################################
' Description :  ��ǰ ����
' History : 2010.03.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/offshop/gift/gift_Cls.asp"-->
<%
Dim clsGift, shopitemname, giftkind_code ,arrList, i
dim itemgubun , itemoption , clsedit , itemid
dim tmp_shopitemname
	itemgubun = requestCheckVar(Request("itemgubun"),40)
	shopitemname = requestCheckVar(Request("shopitemname"),40)
	'giftkind_code  = requestCheckVar(Request("giftkind_code"),10)

'Ư������ �˻�
tmp_shopitemname = Replace(shopitemname, "[", "[[]")
tmp_shopitemname = Replace(tmp_shopitemname, "%", "[%]")

set clsGift = new cgift_list
	clsGift.FRectItemGubun = itemgubun			'��ǰ�˻�
	clsGift.FrectsTxt = tmp_shopitemname

	clsGift.fnGetGiftKind

IF tmp_shopitemname <> "" and clsGift.ftotalcount = 0 and IsNumeric(tmp_shopitemname) THEN
	'��ǰ������ �˻��� �ȵǸ� ��ǰ�ڵ�� �˻�
	clsGift.FRectItemGubun = itemgubun			'��ǰ�ڵ�˻�
	clsGift.FrectsTxt = ""
	clsGift.FRectShopItemid = tmp_shopitemname

	clsGift.fnGetGiftKind
end if
%>
<script language="javascript">

	// �˻�
	function jsSearch(){
		document.frmSearch.submit();
	}

	// ��� �Ǵ� �˻� ȭ������ ����
	function jsChangeMode(giftkind_code){
		if (giftkind_code ==""){
		document.frmSearch.shopitemname.value="";
		}
		document.frmSearch.giftkind_code.value = giftkind_code;
		document.frmSearch.submit();
	}

	// ��ǰ ���
	function jsSubmitGiftKind(){
		if(!frmGift.shopitemname.value){
			 alert("��ǰ���� �Է����ּ���");
			 frmGift.shopitemname.focus();
			 return false;
		}
		frmGift.mode.value='giftitemedit';
		frmGift.submit();
	}

	//�˻��� ��ǰ ����
	function jsSetGiftKind(shopitemname, itemgubun, shopitemid, itemoption){
		opener.document.all.shopitemname.value= shopitemname;

		opener.document.all.itemgubun.value= itemgubun;
		opener.document.all.shopitemid.value= shopitemid;
		opener.document.all.itemoption.value= itemoption;

		window.close();
	}

</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ��ǰ�˻� </div>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="0" >
<form name="frmSearch" method="get" action="popTargetItemReg.asp" >
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<tr>
	<td>��ǰ��/��ǰ�ڵ� : <input type="text" name="shopitemname" size="30" maxlength="60" value="<%= shopitemname %>">
		<input type="button" class="button" value="�˻�" onClick="jsSearch();">
	</td>
</form>
</tr>
<tr>
	<td><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="<%= adminColor("tabletop") %>">
			<td align="center">����</td>
			<td align="center">��ǰ�ڵ�</td>
			<!--
			<td align="center">�ɼ��ڵ�</td>
			-->
			<td align="center">��ǰ��</td>
			<td align="center">�����</td>
			<td align="center">���</td>
		</tr>
		<%
		For i =0 To clsGift.ftotalcount - 1
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= clsGift.FItemList(i).fitemgubun %></td>
			<td align="center"><%= clsGift.FItemList(i).fshopitemid %></td>
			<!--
			<td align="center"><%= clsGift.FItemList(i).fitemoption %></td>
			-->
			<td align="center"><%= clsGift.FItemList(i).fshopitemname %></td>
			<td align="center"><%= FormatDate(clsGift.FItemList(i).fregdate,"0000.00.00") %></td>
			<td align="center">
				<input type="button" value="����" class="button" onClick="jsSetGiftKind('<%= clsGift.FItemList(i).fshopitemname %>','<%= clsGift.FItemList(i).fitemgubun %>','<%= clsGift.FItemList(i).fshopitemid %>','<%= clsGift.FItemList(i).fitemoption %>');">
			</td>
		</tr>
		<% Next	%>
		</table>
		<br>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<%IF shopitemname <> "" and clsGift.ftotalcount = 0 THEN %>
		<tr><td colspan="2"  bgcolor="#FFFFFF"><font color="#E08050"><%= shopitemname %></font>�� �ش��ϴ� ��ǰ�� �����ϴ�.</td></tr>
		<% else %>
		<tr><td colspan="2"  bgcolor="#FFFFFF">* ��ǰ�� �ִ� 30���� ǥ�õ˴ϴ�.</td></tr>
		<%END IF%>
		</table>
	</td>
</tr>
</table>

<%
set clsGift = nothing
set clsedit = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->