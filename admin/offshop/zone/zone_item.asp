<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �𺰱�������
' Hieditor : 2010.12.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->

<%
Dim ozone,i,page,isusing , parameter , shopid ,fromDate ,toDate,yyyy1,mm1,dd1,yyyy2,mm2,dd2, zonegroup ,racktype
dim designer , itemid , itemname , searchtype ,menupos ,datefg,cdl ,cdm ,cds
	designer = RequestCheckVar(request("designer"),32)
	isusing = requestCheckVar(request("isusing"),1)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(request("itemname"),124)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	zonegroup = requestCheckVar(request("zonegroup"),10)
	racktype = requestCheckVar(request("racktype"),10)
	searchtype = requestCheckVar(request("searchtype"),1)
	menupos = requestCheckVar(request("menupos"),10)
	datefg = requestCheckVar(request("datefg"),10)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)

if datefg = "" then datefg = "maechul"
if page = "" then page = 1
if searchtype = "" then searchtype = "M"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

'����/������
if (C_IS_SHOP) then

	'/���α��� ���� �̸�
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if
end if

set ozone = new czone_list
	ozone.FPageSize = 100
	ozone.FCurrPage = page
	ozone.frectzonegroup = zonegroup
	ozone.frectracktype = racktype
	ozone.frectisusing = isusing
	ozone.frectshopid = shopid
	ozone.frectitemid = itemid
	ozone.frectitemname = itemname
	ozone.FRectStartDay = fromDate
	ozone.FRectEndDay = toDate
	ozone.FRectmakerid = designer
	ozone.frectdatefg = datefg
	ozone.FRectCDL = cdl
	ozone.FRectCDM = cdm
	ozone.FRectCDN = cds
	ozone.FRectsearchtype = searchtype

	if shopid <> "" then
		ozone.GetoffshopzoneitemMatch
	end if

if shopid = "" then response.write "<script>alert('������ �������ּ���');</script>"
%>

<script language="javascript">

	//�˻���� ��� ����
	function zone_changeall(upfrm){

		if (upfrm.shopid.value==''){
			alert('������ ������ �ּ���');
			return;
		}
		if (upfrm.zoneidx.value==''){
			alert('���� �Ͻ� ������ ������ �ּ���');
			return;
		}

		upfrm.action='zone_process.asp';
		upfrm.mode.value='zoneitemregall';
		upfrm.submit();
	}

	//���û�ǰ ����
	function zone_change(upfrm){

		if (upfrm.shopid.value==''){
			alert('������ ������ �ּ���');
			return;
		}
		if (upfrm.zoneidx.value==''){
			alert('���� �Ͻ� ������ ������ �ּ���');
			return;
		}

		upfrm.zoneidxarr.value = '';
		upfrm.itemgubunarr.value = '';
		upfrm.shopitemidarr.value = '';
		upfrm.itemoptionarr.value = '';

		if (!CheckSelected()){
				alert('���þ������� �����ϴ�.');
				return;
			}
			var frm;
				for (var i=0;i<document.forms.length;i++){
					frm = document.forms[i];
					if (frm.name.substr(0,9)=="frmBuyPrc") {
						if (frm.cksel.checked){
							upfrm.zoneidxarr.value = upfrm.zoneidxarr.value + frm.zoneidx.value + "," ;
							upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "," ;
							upfrm.shopitemidarr.value = upfrm.shopitemidarr.value + frm.shopitemid.value + "," ;
							upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "," ;
						}
					}
				}

		upfrm.action='zone_process.asp';
		upfrm.mode.value='zoneitemreg';
		upfrm.submit();
	}

	function gopage(page){
		frm.page.value=page;
		frm.submit();
	}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="mode">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="shopitemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="zoneidxarr" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		����:<% drawSelectBoxOffShop "shopid",shopid %>
		<% if searchtype = "M" then %>
			������� :
			<% drawmaechul_datefg "datefg" ,datefg ,""%>
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<% else %>
			<input type="hidden" name="datefg" value="<%=datefg%>">
			<input type="hidden" name="yyyy1" value="<%=yyyy1%>">
			<input type="hidden" name="mm1" value="<%=mm1%>">
			<input type="hidden" name="dd1" value="<%=dd1%>">
			<input type="hidden" name="yyyy2" value="<%=yyyy2%>">
			<input type="hidden" name="mm2" value="<%=mm2%>">
			<input type="hidden" name="dd2" value="<%=dd2%>">
		<% end if %>
		<Br>�׷�: <% drawSelectBoxOffShopzonegroup "zonegroup",zonegroup,"" %>
		�Ŵ�Ÿ��: <% drawSelectBoxOffShopracktype "racktype",racktype,"" %>
		��������:
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>����</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="gopage('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�귣�� : <% drawSelectBoxDesignerwithName "designer",designer %>
		��ǰ�ڵ� : <input type="text" name="itemid" value="<%=itemid %>" size=10>
		��ǰ�� : <input type="text" name="itemname" value="<%=itemname %>" size=20>
		<br>��¥�������:<input type="radio" name="searchtype" value="M" <% if searchtype = "M" then response.write " checked"%>>
		��ǰ��ü����:<input type="radio" name="searchtype" value="I" <% if searchtype = "I" then response.write " checked"%>>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
</tr>
</table>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% drawzonechange "zoneidx",shopid,"" %>
		<input type="button" value="���û�ǰ����" class="button" onclick='zone_change(frm)';>
		<% if ozone.FTotalCount > 0 then %>
			<input type="button" value="�Ѱ˻����(<%= ozone.FTotalCount %>��) �������" class="button" onclick='zone_changeall(frm)';>
		<% end if %>
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ozone.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= ozone.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">��ǰ��ȣ</td>
	<td align="center">��ǰ��(�ɼǸ�)</td>
	<td align="center">�귣��</td>
	<td align="center">��ī�װ�<br>��ī�װ�<br>��ī�װ�</td>
	<td align="center">�׷�</td>
	<td align="center">�Ŵ�Ÿ��</td>
	<td align="center">�󼼱�����</td>
</tr>
<% if ozone.FresultCount>0 then %>
<% for i=0 to ozone.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<input type="hidden" name="itemgubun" value="<%= ozone.FItemList(i).fitemgubun %>">
<input type="hidden" name="shopitemid" value="<%= ozone.FItemList(i).fshopitemid %>">
<input type="hidden" name="itemoption" value="<%= ozone.FItemList(i).fitemoption %>">
<input type="hidden" name="zoneidx" value="<%= ozone.FItemList(i).fzoneidx %>">
<% if ozone.FItemList(i).fzonename <> "" then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>
<tr align="center" bgcolor="#FFFFaa">
<% end if %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center">
		<%= ozone.FItemList(i).fitemgubun %>-<%= CHKIIF(ozone.FItemList(i).fshopitemid>=1000000,Format00(8,ozone.FItemList(i).fshopitemid),Format00(6,ozone.FItemList(i).fshopitemid)) %>-<%= ozone.FItemList(i).fitemoption %>
	</td>

	<td align="center">
		<%= ozone.FItemList(i).fshopitemname %>
		<% if ozone.FItemList(i).fshopitemoptionname <> "" then %>
			(<%= ozone.FItemList(i).fshopitemoptionname %>)
		<% end if %>
	</td>

	<td align="center">
		<%= ozone.FItemList(i).fmakerid %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).fcdl_nm <> "" then %>
			<%= ozone.FItemList(i).fcdl_nm %>
		<% else %>
			-
		<% end if %>
		<% if ozone.FItemList(i).fcdm_nm <> "" then %>
			<br><%= ozone.FItemList(i).fcdm_nm %>
		<% else %>
			<br>-
		<% end if %>
		<% if ozone.FItemList(i).fcds_nm <> "" then %>
			<br><%= ozone.FItemList(i).fcds_nm %>
		<% else %>
			<br>-
		<% end if %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).fzoneidx = "" or isnull(ozone.FItemList(i).fzoneidx)then %>
			������
		<% else %>
			<%= ozone.FItemList(i).fzonegroup_name %>
		<% end if %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).fzoneidx = "" or isnull(ozone.FItemList(i).fzoneidx)then %>
			������
		<% else %>
			<%= getOffShopracktype(ozone.FItemList(i).fracktype) %>
		<% end if %>
	</td>

	<td align="center">
		<% if ozone.FItemList(i).fzoneidx = "" or isnull(ozone.FItemList(i).fzoneidx)then %>
			������
		<% else %>
			<%= ozone.FItemList(i).fzonename %>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ozone.HasPreScroll then %>
			<span class="list_link"><a href="javascript:gopage('<%= ozone.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ozone.StartScrollPage to ozone.StartScrollPage + ozone.FScrollCount - 1 %>
			<% if (i > ozone.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ozone.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:gopage('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ozone.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gopage('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set ozone = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->