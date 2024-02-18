<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ������
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->

<%
dim idx, topidx
idx = requestCheckVar(request("idx"),10)
topidx = requestCheckVar(request("topidx"),10)

' ������
dim ofranchulgomaster
set ofranchulgomaster = new CFranjungsan
ofranchulgomaster.FRectidx = topidx
ofranchulgomaster.getOneFranJungsan

' ���긶����
dim ofranchulgodetail
set ofranchulgodetail = new CFranjungsan
ofranchulgodetail.FRectidx = idx
ofranchulgodetail.getOneFranMaeipSubmaster



dim ofranchulgojungsan

set ofranchulgojungsan = new CFranjungsan
ofranchulgojungsan.FPageSize=1000
ofranchulgojungsan.FRectIDx = idx
ofranchulgojungsan.getFranMaeipSubdetailList

dim i

dim totalsellcash,totalbuycash,totalsuplycash,totalorgsellcash
%>
<script language='javascript'>
function DellArr(frm){
	var ischecked = false;
	frm.suplycasharr.value = "";
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox")) {
			ischecked = (ischecked || e.checked);
			if (e.checked){
				if (frm.elements[i+1].type="text"){
					//frm.suplycasharr.value = frm.suplycasharr.value + frm.elements[i+1].value + ",";
				}
			}
		}
	}

	if (!ischecked) {
		alert('���� ������ �����ϴ�.');
		return;
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.mode.value = "deldetail";
		frm.submit();
	}
}


function SaveArr(frm){
	var ischecked = false;
	frm.suplycasharr.value = "";
	frm.itemnoarr.value = "";
	frm.orgsellcasharr.value = "";
	frm.buycasharr.value = "";
	frm.sellcasharr.value = "";
	
	/* ����
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox")) {
			ischecked = (ischecked || e.checked);
			if (e.checked){
				if (frm.elements[i+1].type="text"){
					frm.itemnoarr.value = frm.itemnoarr.value + frm.elements[i+1].value + ",";
					frm.orgsellcasharr.value = frm.orgsellcasharr.value + frm.elements[i+2].value + ",";
					frm.sellcasharr.value	= frm.sellcasharr.value + frm.elements[i+3].value + ",";
					frm.buycasharr.value	= frm.buycasharr.value + frm.elements[i+5].value + ",";               //��ġ�ٲܰ�� ����;;
					frm.suplycasharr.value = frm.suplycasharr.value + frm.elements[i+4].value + ",";
				}
			}
		}
	}
    */
    
    if (frm.ckidx.length){
        for (var i=0;i<frm.ckidx.length;i++){
            var e = frm.ckidx[i];
            ischecked = (ischecked || e.checked);
            if (e.checked){
    			frm.itemnoarr.value = frm.itemnoarr.value + frm.itemno[i].value + ",";
    			frm.orgsellcasharr.value = frm.orgsellcasharr.value + frm.orgsellcash[i].value + ",";
    			frm.sellcasharr.value	= frm.sellcasharr.value + frm.sellcash[i].value + ",";
    			frm.buycasharr.value	= frm.buycasharr.value + frm.buycash[i].value + ",";              
    			frm.suplycasharr.value = frm.suplycasharr.value + frm.suplycash[i].value + ",";
    		}
        }
    }else{
        var e = frm.ckidx;
        ischecked = (ischecked || e.checked);
        if (e.checked){
			frm.itemnoarr.value = frm.itemnoarr.value + frm.itemno.value + ",";
			frm.orgsellcasharr.value = frm.orgsellcasharr.value + frm.orgsellcash.value + ",";
			frm.sellcasharr.value	= frm.sellcasharr.value + frm.sellcash.value + ",";
			frm.buycasharr.value	= frm.buycasharr.value + frm.buycash.value + ",";              
			frm.suplycasharr.value = frm.suplycasharr.value + frm.suplycash.value + ",";
		}
    }
    
	if (!ischecked) {
		alert('���� ������ �����ϴ�.');
		return;
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.mode.value = "modidetail";
		frm.submit();
	}
}

function AddNewDetail(idx,topidx,makerid){
	if ("<%=ofranchulgomaster.FOneItem.FstateCd%>" >= "4")
	{
		alert("��꼭 ���� ���Ŀ��� ��Ÿ�����߰� �� �� �����ϴ�.")
		return;
	}
	if ("<%=ofranchulgodetail.FOneItem.Flinkidx%>" != "0")
	{
		alert("��Ÿ�߰��� ���������� ��Ÿ��ǰ�߰��� �� �� �ֽ��ϴ�.")
		return;
	}

	var popwin = window.open('popetcmeachul_etcjungsandetailadd.asp?idx=' + idx + '&topidx=' + topidx + '&makerid=' + makerid +'&shopid=<%= ofranchulgodetail.FOneItem.Fshopid %>','franetcjungsandetailadd','width=500, height=400, scrollbars=yes, resizable=yes');
	popwin.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>Index</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgodetail.FOneItem.Fidx %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>����ID</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgodetail.FOneItem.Fshopid %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����ڵ�</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgodetail.FOneItem.Fcode01 %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ֹ��ڵ�</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgodetail.FOneItem.Fcode02 %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�ѼҺ�</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgodetail.FOneItem.Ftotalorgsellcash,0) %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">���ǸŰ�</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgodetail.FOneItem.Ftotalsellcash,0) %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">���ް�</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgodetail.FOneItem.Ftotalsuplycash,0) %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">���԰�</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgodetail.FOneItem.Ftotalbuycash,0) %></td>
	</tr>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="��Ÿ���� �߰�" onclick="AddNewDetail('<%= idx %>','<%= topidx %>','<%= ofranchulgodetail.FOneItem.Fcode02 %>');">
		</td>
		<td align="right">
			<input type="button" class="button" value="���ó��� ����" onclick="DellArr(frmarr);">
			&nbsp;
			<input type="button" class="button" value="���ó��� ����" onclick="SaveArr(frmarr);">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align=center>
	<td width="30">����<!-- <input type=checkbox name=ck_all onclick=""> --></td>
	<td width="40">����</td>
	<td width="80">�ֹ���ȣ<br>�����ڵ�</td>
	<td width="80">�귣��ID</td>
	<td width="80">���ڵ�</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width="40">����</td>
	<td width="60">�Һ�</td>
	<td width="60">�ǸŰ�</td>
	<td width="60">���ް�</td>
	<td width="60">���԰�</td>
	<td width="60">�񱳸��԰�<br>(�������)</td>
</tr>
<form name=frmarr method=post action="etc_meachul_process.asp">
<input type=hidden name="mode" value="modidetail">
<input type=hidden name="orgsellcasharr" value="">
<input type=hidden name="sellcasharr" value="">
<input type=hidden name="buycasharr" value="">
<input type=hidden name="suplycasharr" value="">
<input type=hidden name="itemnoarr" value="">
<input type=hidden name="topidx" value="<%= topidx %>">
<% for i=0 to ofranchulgojungsan.FResultCount - 1 %>
<%
totalsuplycash = totalsuplycash + ofranchulgojungsan.FItemList(i).Fsuplycash * ofranchulgojungsan.FItemList(i).Fitemno
totalbuycash   = totalbuycash + ofranchulgojungsan.FItemList(i).Fbuycash * ofranchulgojungsan.FItemList(i).Fitemno
%>
<tr bgcolor="#FFFFFF" >
	<td align=center><input type="checkbox" name="ckidx" value="<%= ofranchulgojungsan.FItemList(i).Fidx %>" onClick="AnCheckClick(this);"></td>
	<td align=center><%= ofranchulgojungsan.FItemList(i).Flinkbaljucode %></td>
	<td align=center><%= ofranchulgojungsan.FItemList(i).Flinkmastercode %></td>
	<td align=center><%= ofranchulgojungsan.FItemList(i).Fmakerid %></td>
	<td align=center><%= ofranchulgojungsan.FItemList(i).GetBarCode %></td>
	<td><%= ofranchulgojungsan.FItemList(i).Fitemname %></td>
	<td align=center><%= ofranchulgojungsan.FItemList(i).Fitemoptionname %></td>
	<td align=center>
	<input type=text name="itemno" value="<%= ofranchulgojungsan.FItemList(i).Fitemno  %>" size=4 maxlength=8 style="border:1px #999999 solid; text-align=center">
	</td>
	<td align=right><input type=text name="orgsellcash" value="<%= ofranchulgojungsan.FItemList(i).Forgsellcash %>" size=7 maxlength=8 style="border:1px #999999 solid; text-align=right"></td>
	<td align=right><input type=text name="sellcash" value="<%= ofranchulgojungsan.FItemList(i).Fsellcash %>" size=7 maxlength=8 style="border:1px #999999 solid; text-align=right"></td>
	<td align=right><input type=text name="suplycash" value="<%= ofranchulgojungsan.FItemList(i).Fsuplycash  %>" size=7 maxlength=8 style="border:1px #999999 solid; text-align=right"></td>
	<td align=right><input type=text name="buycash" value="<%= ofranchulgojungsan.FItemList(i).Fbuycash %>" size=7 maxlength=8 style="border:1px #999999 solid; text-align=right"></td>
	<td align=right>
	    <% if (ofranchulgojungsan.FItemList(i).Fbuycash<>ofranchulgojungsan.FItemList(i).Flstbuycash) then %>
	    <font color="red"><%= FormatNumber(ofranchulgojungsan.FItemList(i).Flstbuycash,0) %></font>
	    <% else %>
	    <%= FormatNumber(ofranchulgojungsan.FItemList(i).Flstbuycash,0) %>
	    <% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" >
	<td colspan=10></td>
	<td align=right><%= FormatNumber(totalsuplycash,0) %></td>
	<td align=right><%= FormatNumber(totalbuycash,0) %></td>
	<td align=right></td>
</tr>
</form>
</table>
<%
Set ofranchulgomaster = nothing
set ofranchulgodetail = nothing
set ofranchulgojungsan = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->