<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

response.write "������� �޴� &nbsp;&nbsp;&nbsp;&nbsp;"
response.write " Go �� �޴�"
dbget.close()	:	response.End

dim designerid,itemid,itemname
dim page,dispyn,sellyn,usingyn, cdl, strParam
designerid = request("designerid")
itemid = request("itemid")
itemname = request("itemname")
dispyn = request("dispyn")
sellyn = request("sellyn")
usingyn = request("usingyn")
cdl = request("cdl")
page = request("page")
if page="" then page=1

strParam = "&itemid=" & itemid & "&itemname=" & itemname & "&designerid=" & designerid & "&dispyn=" & dispyn & "&sellyn=" & sellyn & "&cdl=" & cdl

dim obuyprice
set obuyprice = new CBuyPrice
obuyprice.FCurrPage = page
obuyprice.FPageSize = 100
obuyprice.FSearchItemName = itemname
obuyprice.FSearchDesigner = designerid
obuyprice.FSearchItemid = itemid
obuyprice.FSearchDispYn = dispyn
obuyprice.FSearchSellYn = sellyn
obuyprice.FSearchusingyn = usingyn
obuyprice.FRectCD1 = cdl
obuyprice.getPrcList

dim i, disexists
%>
<!--
<h2><font color=red><center>������</center></font></h2>
-->
<script language='javascript'>
function AnBuyPriceSaveFrame2(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� ��ǰ�� �����ϴ�.');
		return;
	}

	var ret = confirm('���� ��ǰ�� �����Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;
					upfrm.sellcash.value = upfrm.sellcash.value + "|" + frm.sellcash.value;
					upfrm.sellvat.value = upfrm.sellvat.value + "|" + frm.sellvat.value;
					upfrm.buycash.value = upfrm.buycash.value + "|" + frm.buycash.value;
					upfrm.buyvat.value = upfrm.buyvat.value + "|" + frm.buyvat.value;
					upfrm.marginrate.value = upfrm.marginrate.value + "|" + frm.marginrate.value;
					upfrm.vtinclude.value = upfrm.vtinclude.value + "|" +frm.vtinclude.value;
				}
			}
		}
		upfrm.submit();
	}
}

function AnAllCalcu2(frm){
	var frmtarget;

	if (!IsDouble(frm.mgall.value)){
		alert('�������� ���ڸ� �����մϴ�.');
		frm.mgall.focus();
		return;
	}

	if ((frm.mgall.value<0)&&(frm.mgall.value>100)){
		alert('�������� 0~100�� �����մϴ�.');
		frm.mgall.focus();
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frmtarget = document.forms[i];
		if (frmtarget.name.substr(0,9)=="frmBuyPrc") {
			if (!frmtarget.cksel.checked) continue;
			frmtarget.marginrate.value = frm.mgall.value;
			AnAutoCalcu2(frmtarget,true);
		}
	}
}

function AnAutoCalcu2(frm,bool){
	var buf;
	if (!IsDigit(frm.sellcash.value)){
		alert('�ǸŰ��� ���ڸ� �����մϴ�.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDouble(frm.marginrate.value)){
		alert('�������� ���ڸ� �����մϴ�.');
		frm.marginrate.focus();
		return;
	}

	if ((frm.marginrate.value<0)&&(frm.marginrate.value>100)){
		alert('�������� 0~100�� �����մϴ�.');
		frm.marginrate.focus();
		return;
	}

	if (bool){
		frm.sellvat.value = parseInt(Math.round(frm.sellcash.value/11));
		buf = parseInt(Math.round(frm.sellcash.value*(1-frm.marginrate.value/100.0)));
		frm.buycash.value = buf;
		frm.buyvat.value = parseInt(Math.round(buf/11));
		frm.tmpbuycash.value = parseInt(Math.round(buf-frm.buyvat.value));
		//frm.buyvat.value = Math.floor(buf/11);
		//frm.tmpbuycash.value = Math.floor(buf-frm.buyvat.value);
		//frm.buycash.value = Math.floor(frm.buyvat.value * 1 + frm.tmpbuycash.value * 1);

	}else{
		frm.sellvat.value = parseInt(Math.round(frm.sellcash.value/11));
		frm.tmpbuycash.value = parseInt(Math.round(frm.sellcash.value*(1-frm.marginrate.value/100)));
		frm.buycash.value = parseInt(Math.round(frm.tmpbuycash.value*1.1));
		frm.buyvat.value = parseInt(Math.round(frm.tmpbuycash.value*0.1));
		//frm.tmpbuycash.value = Math.floor(frm.sellcash.value*(1-frm.marginrate.value/100));
		//frm.buyvat.value = Math.floor(frm.buycash.value/10);
		//frm.buycash.value = Math.floor(frm.buyvat.value * 1  + frm.tmpbuycash.value * 1 );
	}
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}
</script>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr>
		<td class="a" >
		�귣�� :
		<% drawSelectBoxDesignerWithName "designerid",designerid %>
		&nbsp;
		ī�װ� :
		<% SelectBoxBrandCategory "cdl" ,cdl %>
		<!--
		<select name="cdl">
			<option value="">��ü</option>
			<option value="10">�����ι���,���μ�ǰ</option>
			<option value="40">Ű��Ʈ,��,���,ī�޶�</option>
			<option value="15">���׸���,��������</option>
			<option value="25">�ֹ��Ȱ,���,�ٵ�</option>
			<option value="30">�м� �Ƿ�</option>
			<option value="32">�м� ��ȭ</option>
			<option value="35">���,�ð�,���</option>
			<option value="50">�ö��</option>
		</select>
		-->
		<script language="javascript">frm.cdl.value='<%=cdl%>'</script>
		<br>
		��ǰID :
		<input type="text" name="itemid" value="<%= itemid %>" size="50"> (��ǥ�� �����Է°���)
		<br>
		��ǰ�� :
		<input type="text" name="itemname" value="<%= itemname %>" size="12" maxlength="32">
		���ÿ��� :
		<select name="dispyn">
     	<option value='' selected>����</option>
     	<option value='Y' <% if dispyn="Y" then response.write "selected" %> >Y</option>
     	<option value='N' <% if dispyn="N" then response.write "selected" %> >N</option>
     	</select>
     	&nbsp;
		�Ǹſ��� :
		<select name="sellyn">
     	<option value='' selected>����</option>
     	<option value='Y' <% if sellyn="Y" then response.write "selected" %> >Y</option>
     	<option value='N' <% if sellyn="N" then response.write "selected" %> >N</option>
     	</select>
     	&nbsp;
     	��뿩�� :
		<select name="usingyn">
     	<option value='' selected>����</option>
     	<option value='Y' <% if usingyn="Y" then response.write "selected" %> >Y</option>
     	<option value='N' <% if usingyn="N" then response.write "selected" %> >N</option>
     	</select>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#000000">
<tr bgcolor="FFFFFF">
	<td colspan="15" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(obuyprice.FTotalPage,0) %> count: <%= FormatNumber(obuyprice.FTotalCount,0) %></td>
</tr>
<tr bgcolor="FFFFFF">
	<form name="frmttl" onsubmit="return false;">
	<td colspan="6" height="30"><input type="button" value="��ü����" onClick="AnSelectAllFrame(true)">&nbsp;<input type="button" value="���û�ǰ����" onClick="AnBuyPriceSaveFrame2()"></td>
	<td colspan="7" height="30" align=right>������ : <input type="text" name="mgall" size="3" maxlength="5"> &nbsp;<input type="button" value="���û�ǰ����" onClick="AnAllCalcu2(frmttl)"></td>
	</form>
</tr>
<tr bgcolor="DDDDFF" align="center">
	<td>����</td>
	<td>��ǰID</td>
	<td>��ǰ��</td>
	<td>�귣��ID</td>
	<td>��۱���</td>
	<td>�����ǸŰ�</td>
	<td>���԰�</td>
	<td>������</td>
	<td>����<br>����</td>
	<td>����<br>����</td>
	<td>����<br>����</td>
	<td>�Ǹ�<br>����</td>
	<td>���<br>����</td>
</tr>

<% for i=0 to obuyprice.FresultCount-1 %>
<form name="frmBuyPrc_<%= obuyprice.FItemList(i).FItemID %>" method="post" onSubmit="return CheckNDobuyprice(this);" action="doitembuyprice.asp">
<input type="hidden" name="itemid" value="<%= obuyprice.FItemList(i).FItemID %>">
<input type="hidden" name="sellvat" value="<%= obuyprice.FItemList(i).FSellVat %>">
<input type="hidden" name="tmpbuycash" value="<%= obuyprice.FItemList(i).FBuyPrice %>">
<input type="hidden" name="buyvat" value="<%= obuyprice.FItemList(i).FBuyVat %>">
<input type="hidden" name="vtinclude" value="<%= obuyprice.FItemList(i).Fvatinclude %>">

<tr bgcolor="#FFFFFF" align="center">

        <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if obuyprice.FItemList(i).FSailYn="Y" then response.write "disabled" : disexists= true %> ></td>
	<td><a href="javascript:PopItemSellEdit('<%= obuyprice.FItemList(i).FItemID %>');"><%= obuyprice.FItemList(i).FItemID %></a></td>
	<td align=left><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= obuyprice.FItemList(i).FItemID %>" target="_blank"><%= obuyprice.FItemList(i).FItemName %></a></td>
	<td align=left><%= obuyprice.FItemList(i).FMakerID %></td>
	<td align="center">
	<% if obuyprice.FItemList(i).FBaesongGB="1" then %>
		10x10
	<% else %>
	   	<font color=red><%= BaesongCd2Name(obuyprice.FItemList(i).FBaesongGB) %></font>
	<% end if %>
	</td>
	<td align="right"><input type="text" name="sellcash" size="6" value="<%= obuyprice.FItemList(i).FSellPrice %>"></td>
	<td align="right"><input type="text" name="buycash" size="6" value="<%= obuyprice.FItemList(i).FBuyPrice %>"></td>
	<td align="right"><input type="text" name="marginrate" size="3" maxlength="3" value="<%= obuyprice.FItemList(i).GetCalcuMarginrate %>" readonly style="border:0; text-align:right">%</td>
	<% if obuyprice.FItemList(i).FVatInclude="N" then %>
	<td align="center">�鼼</td>
	<% else %>
	<td align="center"></td>
	<% end if %>
	<% if obuyprice.FItemList(i).Fsailyn="Y" then %>
	<td align="center"><font color=red>����</font></td>
	<% else %>
	<td align="center"></td>
	<% end if %>
	<% if obuyprice.FItemList(i).FDisplayYn="Y" then %>
	<td align="center"><%= obuyprice.FItemList(i).FDisplayYn %></td>
	<% else %>
	<td align="center"><font color="red"><%= obuyprice.FItemList(i).FDisplayYn %></font></td>
	<% end if %>

	<% if obuyprice.FItemList(i).FSellYn="Y" then %>
	<td align="center"><%= obuyprice.FItemList(i).FSellYn %></td>
	<% else %>
	<td align="center"><font color="red"><%= obuyprice.FItemList(i).FSellYn %></font></td>
	<% end if %>

	<% if obuyprice.FItemList(i).FIsUsing="Y" then %>
	<td align="center"><%= obuyprice.FItemList(i).FIsUsing %></td>
	<% else %>
	<td align="center"><font color="red"><%= obuyprice.FItemList(i).FIsUsing %></font></td>
	<% end if %>
</tr>

</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center">
	<% if obuyprice.HasPreScroll then %>
		<a href="?page=<%= obuyprice.StartScrollPage-1 & strParam %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + obuyprice.StartScrollPage to obuyprice.FScrollCount + obuyprice.StartScrollPage - 1 %>
		<% if i>obuyprice.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i & strParam %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if obuyprice.HasNextScroll then %>
		<a href="?page=<%= i & strParam %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td colspan="14" height="20">
</tr>
<form name="frmArrupdate" method="post" action="doitembuyprice.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="sellcash" value="">
<input type="hidden" name="sellvat" value="">
<input type="hidden" name="marginrate" value="">
<input type="hidden" name="buycash" value="">
<input type="hidden" name="buyvat" value="">
<input type="hidden" name="vtinclude" value="">
</form>
</table>
<%
set obuyprice = Nothing
%>
<% if disexists=true and page=1 then %>
<script>alert('���� �������� ��ǰ�� �̰����� ���� �Ұ��� �մϴ�.');</script>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->