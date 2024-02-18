<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecturecls.asp"-->
<%
dim page
dim i, ix, olec
dim yyyy1,mm1,nowdate
dim yyyy2,mm2,dd2
dim itemid,itemname,lecturer,lecturdate,lecturdateyn

nowdate = now()
itemid = request("itemid")
itemname = request("itemname")
lecturer = request("lecturer")
lecturdate = request("lecturdate")
yyyy1 = request("yyyy1")
mm1   = request("mm1")
lecturdateyn   = request("lecturdateyn")

if yyyy1="" then
	yyyy1 = Left(Cstr(nowdate),4)
	mm1	  = Mid(Cstr(nowdate),6,2)
end if

yyyy2 = request("yyyy2")
mm2   = request("mm2")
dd2   = request("dd2")

if yyyy2="" then
	yyyy2 = Left(Cstr(nowdate),4)
	mm2	  = Mid(Cstr(nowdate),6,2)
	dd2	  = Mid(Cstr(nowdate),9,2)
end if
lecturdate = yyyy2 + "-" + mm2 + "-" + dd2
page = request("page")

if page="" then page=1

set olec = new CLecture
olec.FPageSize=100
olec.FCurrPage = page
olec.FRectYYYYMM = yyyy1 + "-" +mm1
olec.FRectItemID = itemid
olec.FRectItemName = itemname
olec.FRectLecturer = lecturer
if lecturdateyn = "on" then
olec.FRectDate = lecturdate
end if
olec.GetLectureList
%>
<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function AddLec(iitemid,iitemoption,iitemea){
	document.lecadd.itemid.value=iitemid;
	document.lecadd.itemoption.value=iitemoption;
	document.lecadd.itemea.value=iitemea;
	document.lecadd.submit();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function arrsave(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.idxarr.value = "";
	upfrm.viewidxarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (!IsDigit(frm.viewidx.value)){
				alert('ǥ�ü����� ���ڸ� �����մϴ�.');
				frm.viewidx.focus();
				return;
			}

			upfrm.idxarr.value = upfrm.idxarr.value + frm.idx.value + "|";
			upfrm.viewidxarr.value = upfrm.viewidxarr.value + frm.viewidx.value + "|";
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		upfrm.submit();
	}
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function clearText(y){
    if (y.defaultValue==y.value)
        y.value = ""
}
</script>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		�˻��� : <% DrawYMBox yyyy1,mm1 %>&nbsp;��ǰ�ڵ� : <input type="text" name="itemid" size="8" value="<% =itemid %>">&nbsp;���¸� : <input type="text" name="itemname" size="20"  value="<% =itemname %>"><br>
		����� : <input type="text" name="lecturer" size="10" value="<% =lecturer %>">&nbsp;<input type="checkbox" name="lecturdateyn" <% if lecturdateyn = "on" then response.write "checked" %>>������ : <% DrawOneDateBox2 yyyy2,mm2,dd2 %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="800" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td ><a href="lecreg.asp?mode=add">[������ ���]</a>&nbsp;&nbsp;&nbsp;<a href="javascript:arrsave();">[���û�������]</a></td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
   <td width="20" align="center"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<td align="center" width="30">Idx</td>
	<td align="center" width="50">����</td>
	<td align="center">��ǰ�ڵ�</td>
	<td align="center">���¸�</td>
	<td align="center" width="70">�����</td>
	<td align="center" width="100">������</td>
	<td align="center" width="80">������</td>
	<td align="center" width="80">(�����ñݾ�)</td>
	<td align="center" width="50">��û�ο�(����)</td>
	<td align="center" width="50">��������</td>
	<td align="center" width="50">��������</td>
	<td align="center" width="50">��������</td>
	<td align="center" width="50">���ÿ���</td>
	<td align="center" width="50">��ȸ</td>
	<td align="center" width="50">�����Է�</td>
	<td align="center" width="50">�����¿� ����</td>
</tr>
<% for i=0 to olec.FResultCount - 1 %>
<form name="frmBuyPrc_<%= i %>" method="post">
<input type="hidden" name="idx" value="<%= olec.FItemList(i).Fidx %>">
<tr bgcolor="#FFFFFF">
   <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><% = olec.FItemList(i).Fidx %></td>
	<td align="center"><input type="text" name="viewidx" size="3" value="<% = olec.FItemList(i).Fviewidx %>" onFocus="clearText(this)"></td>
	<td align="center"><% = olec.FItemList(i).FLinkItemId %></td>
	<td><a href="lecreg.asp?idx=<% = olec.FItemList(i).Fidx %>&mode=edit"><% = olec.FItemList(i).Flectitle %></a></td>
	<td align="center"><% = olec.FItemList(i).Flecturer %></td>
	<td align="center"><% = FormatDateTime(olec.FItemList(i).Flecdate01,2) %></td>
	<% if olec.FItemList(i).Flecsum<>olec.FItemList(i).Fsellcash then %>
	<td align="right"><b><% = FormatNumber(olec.FItemList(i).Flecsum,0) %> ��&nbsp;</b></td>
	<% else %>
	<td align="right"><% = FormatNumber(olec.FItemList(i).Flecsum,0) %> ��&nbsp;</td>
	<% end if %>
	<td align="right"><% = FormatNumber(olec.FItemList(i).Fsellcash,0) %> ��&nbsp;</td>
	<td align="right"><% = olec.FItemList(i).FOrgLimitSold %> ��&nbsp;</td>
	<td align="center"><a href="javascript:PopItemSellEdit('<% = olec.FItemList(i).Flinkitemid %>')">����</a></td>
	<td align="center"><a href="lecregdetail.asp?idx=<% = olec.FItemList(i).Fidx %>" target="_blank">����</a></td>
	<td align="center"><%= olec.FItemList(i).Fregfinish %></td>
	<td align="center"><%= olec.FItemList(i).FIsUsing %></td>
	<td align="center"><%= olec.FItemList(i).Freadcnt %></td>
	<td align="center"><a href="javascript:AddLec('<% = olec.FItemList(i).FLinkItemId %>','0000','1');">�Է�</a></td>
	<td align="center"><a href="_lecreg.asp?idx=<% = olec.FItemList(i).Fidx %>&mode=add">����</a></td>
</tr>
</form>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="17" height="30" align="center">
		<% if olec.HasPreScroll then %>
			<a href="javascript:NextPage('<%= olec.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for ix=0 + olec.StarScrollPage to olec.FScrollCount + olec.StarScrollPage - 1 %>
			<% if ix>olec.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if olec.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<%
set olec = Nothing
%>
<form name="lecadd" method="post" action="http://www.10x10.co.kr/inipay/collegelecreg.asp">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="itemea" value="">
</form>
<form name="frmArrupdate" method="post" action="doviewidxarr.asp">
<input type="hidden" name="mode" value="viewidxedit">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="viewidxarr" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->