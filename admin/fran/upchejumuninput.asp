<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
' �����ϴµ�
response.end
%>

<script language='javascript'>
function MakeJumunByIdx(idxarr,designerid,etcstr){
	//alert(idxarr);
	//alert(designerid);
	document.dumifrm.idxarr.value=idxarr;
	document.dumifrm.designerid.value=designerid;
	document.dumifrm.etcstr.value=etcstr;
	document.dumifrm.submit();
}

function PopFranBalju2Upchebalju(frm){
	var designerid,baljuid,popwin;
	designerid = frm.designerid.value;
	baljuid = frm.baljuid.value;
	popwin = window.open('popfranbalju2upchebalju.asp?designerid=' + designerid + '&baljuid=' + baljuid  ,'franbalju2upchebalju','width=800,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function PopFranBalju2UpchebaljuByID(designerid){
    var baljuid,popwin;
	baljuid = "10x10";
	popwin = window.open('popfranbalju2upchebalju.asp?designerid=' + designerid + '&baljuid=' + baljuid  ,'franbalju2upchebalju','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
<%
''��ü�����ֹ����ۼ�

dim idxarr,designerid,statecd, etcstr
dim designer
dim includepreorderno, shortyn
dim research

idxarr      		= request("idxarr")
designerid  		= request("designerid")
statecd     		= request("statecd")
designer    		= request("designer")
etcstr      		= request("etcstr")
shortyn   			= request("shortyn")
includepreorderno   = request("includepreorderno")
research   			= request("research")

if (research = "") then
	'shortyn = "Y"
	'includepreorderno = "Y"
end if

if (includepreorderno = "Y") then
	shortyn = "Y"
end if



dim oupchejumun,iidx
dim DefaultItemMwDiv

iidx =0
if (idxarr<>"") and (designerid<>"") then
	if Right(idxarr,1)="," then idxarr=Left(idxarr,Len(idxarr)-1)
	if Right(etcstr,1)="," then etcstr=Left(etcstr,Len(etcstr)-1)

	DefaultItemMwDiv = GetDefaultItemMwdivByBrand(designerid)

	set oupchejumun = new COrderSheet
	oupchejumun.FRectIdxArr  = idxarr
	oupchejumun.FRectMakerid = designerid
	oupchejumun.FRectTargetid = designerid
	oupchejumun.FRectBaljuId = "10x10"
	oupchejumun.FRectBaljuname = "�ٹ�����"
	oupchejumun.FRectReguser = session("ssBctId")
	oupchejumun.FRectRegname = session("ssBctCname")
	if (DefaultItemMwDiv="M") then
	oupchejumun.FRectdivcode = "101"
	else
	oupchejumun.FRectdivcode = "111"
	end if
	oupchejumun.FRectComment = "���ֹ� : " + html2db(etcstr)

	iidx = oupchejumun.MakeUpcheJumun

	set oupchejumun = Nothing

	'�ֹ��� ���� ���ֹ� ������Ʈ
	PreOrderUpdateBySheetIdx(iidx)

	response.redirect "/admin/fran/upchejumuninputedit.asp?idx=" + CStr(iidx) + "&opage=1&ourl=upchejumunlist.asp"
	dbget.close()	:	response.End
end if

dim oordersheet1
set oordersheet1 = new COrderSheet
oordersheet1.FRectMakerid = designer
oordersheet1.FRectStatecd = statecd
oordersheet1.FRectBaljuId = "10x10"

oordersheet1.FRectShortYN = shortyn
oordersheet1.FRectIncludePreOrderNo = includepreorderno

oordersheet1.GetFranBalju2UpcheBaljuBrandlist

dim i
%>
<!--
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#FFFFFF">
	<td>
		* 2�� 1�� �ֹ� ���� <br>
		���Ծ�ü��� ������ �¶��� ����� ��������.<br>
		��Ź��ü�� ���԰�->�¶������ ������ ���� ���<br>
		��Ź��ü�� ��Ź��->�¶������ ������ ���� ���<br>
		<br>
		�̰����� ���� �ֹ��ؾ� �ϴ°��<br>
		- �����ε� ������ �ٸ����(���� ���� prixe, multiple_choice, nanishow)<br>
		- ��ü����ֹ���.(�������� ��������, �������� ������Ź)
	</td>
</tr>
</table>
-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="searchfrm" method=get">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    �귣�� : <% drawSelectBoxDesignerwithName "designer", designer %>
		    &nbsp;
			<input type=radio name="statecd" value="" <% if statecd="" then response.write "checked" %> >�ֹ����� + ��ǰ�غ�
			<input type=radio name="statecd" value="0" <% if statecd="0" then response.write "checked" %> >�ֹ�����
			<input type=radio name="statecd" value="1" <% if statecd="1" then response.write "checked" %> >��ǰ�غ�
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.searchfrm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    �������� :
			<input type=checkbox name="shortyn" value="Y" <% if shortyn = "Y" then response.write "checked" %>> ��������
			<input type=checkbox name="includepreorderno" value="Y" <% if includepreorderno = "Y" then response.write "checked" %>> ���ֹ����Ժ�����
		</td>
	</tr>
	</form>
</table>

<p>
<!--
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=post action="">
	<tr bgcolor="#FFFFFF">
	<% if designerid<>"" then %>
		<input type="hidden" name="designerid" value="<%= designerid %>">
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>����ó</td>
		<td><%= designerid %></td>
	<% else %>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>����ó</td>
		<td>
			<% drawSelectBoxDesignerwithName "designerid",designerid %>
			&nbsp;
			<input type="button" class="button" value="������ �ֹ����� �ۼ�" onclick="PopFranBalju2Upchebalju(frm);">
		</td>
	<% end if %>
	</tr>
	<tr bgcolor="#FFFFFF">
		<input type="hidden" name="baljuid" value="10x10">
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>����ó</td>
		<td>10x10</td>
	</tr>
	</form>
</table>
-->
<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<!-- <input type="button" class="button" value="�ֹ����ۼ�" onClick=""> -->
		</td>
		<td align="right">
		    * �귣�� ���̵� Ŭ���� �ۼ� ����
			/ ��ü ��ǰ �ֹ����� �̰����� �ۼ� �Ͻ� �� �����ϴ�.
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm2" method=post action="">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width=20></td>
		<td width="120">�귣��ID</td>
		<td width="100">��ǰ�ڵ�</td>
		<td width="50">�̹���</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="40">���<br>����</td>
		<td width="40">����<br>����</td>
		<td width="50">OFF<br>��ǰ�غ�</td>
		<td width="50">OFF<br>�ֹ�����</td>
		<td width="50"><b>�����<br>�հ�</b></td>
		<td width="50">�ǻ�<br>���</td>
		<td width="50"><b>����<br>����</b></td>
		<td width="100">���</td>
	</tr>
	<% for i=0 to oordersheet1.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><input type=checkbox name="cksel" onClick="AnCheckClick(this);"></td>
		<input type=hidden name="idx" value="">
		<td><a href="javascript:PopFranBalju2UpchebaljuByID('<%= oordersheet1.FItemList(i).FMakerid %>');"><%= oordersheet1.FItemList(i).FMakerid %></a></td>
		<td><%= oordersheet1.FItemList(i).FItemGubun %>-<%= CHKIIF(oordersheet1.FItemList(i).FItemId>=1000000,Format00(8,oordersheet1.FItemList(i).FItemId),Format00(6,oordersheet1.FItemList(i).FItemId)) %>-<%= oordersheet1.FItemList(i).FItemoption %></td>
		<td></td>
		<td align="left">
			<%= oordersheet1.FItemList(i).FItemName %>
				&nbsp;
			<% if oordersheet1.FItemList(i).FItemoption<>"0000" then %>
				<font color="blue"><%= oordersheet1.FItemList(i).FItemOptionname %></font>
			<% end if %>
		</td>
		<td><%= oordersheet1.FItemList(i).GetDeliverTypeString %></td>
		<td><%= oordersheet1.FItemList(i).GetMWDivString %></td>
		<td><%= oordersheet1.FItemList(i).FCount %></td>
		<td><%= oordersheet1.FItemList(i).FJupsuCount %></td>
		<td><b><%= oordersheet1.FItemList(i).FCount + oordersheet1.FItemList(i).FJupsuCount %></b></td>
		<td><%= oordersheet1.FItemList(i).Frealstock %></td>
		<td><b><%= oordersheet1.FItemList(i).Frealstock - oordersheet1.FItemList(i).FCount - oordersheet1.FItemList(i).FJupsuCount %></b></td>
		<td>
			<% if ((Not IsNull(oordersheet1.FItemList(i).FreipgoMayDate)) and (Left(oordersheet1.FItemList(i).FreipgoMayDate, 10) >= Left(DateAdd("m", -3, now()), 10) ) ) then %>
				<%= Left(oordersheet1.FItemList(i).FreipgoMayDate, 10) %><br>
			<% end if %>
			<% if oordersheet1.FItemList(i).Fpreorderno<>0 then %>
				���ֹ�:
				<% if oordersheet1.FItemList(i).Fpreorderno<>oordersheet1.FItemList(i).Fpreordernofix then response.write "</br>" + CStr(oordersheet1.FItemList(i).Fpreorderno) + "->" %>
					<%= oordersheet1.FItemList(i).Fpreordernofix %>
			<% end if %>
		</td>
	</tr>
	<% next %>
	</form>
</table>
<%
set oordersheet1 = nothing
%>
<form name="dumifrm" method=post action="">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="designerid" value="">
<input type="hidden" name="etcstr" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
