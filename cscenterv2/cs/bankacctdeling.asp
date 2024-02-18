<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/bankacctcls.asp"-->
<%
dim ojumun, page, daydiff

daydiff = RequestCheckvar(request("daydiff"),3)
page = RequestCheckvar(request("page"),10)
if page="" then page=1
if daydiff="" then daydiff=10

set ojumun = new CBankAcct
ojumun.FCurrPage = page
ojumun.FPageSize = 50
ojumun.FRectDayDiff = daydiff
ojumun.GetOldMiipkumList

dim i
%>
<script language='javascript'>
function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="/cscenterv2/order/orderdetail_view.asp"
	frm.submit();

}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function delitems(upfrm){

    /*
    alert('���̻� �������� �ʴ� �޴� �Դϴ�. - ���� ���� �ڵ����� �߼۵�.');
    return;
    */

	if (!CheckSelected()){
		alert('���� �ֹ��� �����ϴ�.');
		return;
	}

	var ret = confirm('������ ������ �ֹ��� ����Ͻðڽ��ϱ�?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderidx.value = upfrm.orderidx.value + frm.orderidx.value + "|" ;
					upfrm.orderserial.value = upfrm.orderserial.value + frm.orderserial.value + "|" ;
				}
			}
		}
		upfrm.mode.value="del";
		upfrm.submit();
	}
}
</script>
&nbsp;
<p>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<select class="select" name="daydiff">
				<option value="7" <% if daydiff="7" then response.write "selected" %> >7�� ����</option>
				<option value="10" <% if daydiff="10" then response.write "selected" %> >10�� ����</option>
				<option value="15" <% if daydiff="15" then response.write "selected" %> >15�� ����</option>
				<option value="30" <% if daydiff="30" then response.write "selected" %> >30�� ����</option>
			</select>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<form name="frmarr" method="post" action="dobankacct.asp">
	<input type="hidden" name="orderidx" value="">
	<input type="hidden" name="orderserial" value="">
	<input type="hidden" name="mode" value="">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�����ֹ�����" onClick="delitems(frmarr)">
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= FormatNumber(ojumun.FTotalCount,0) %></b>
			&nbsp;
			������ : <b><%= ojumun.FCurrPage %> / <%=ojumun.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="100">�ֹ���ȣ</td>
		<td width="80" align="center">Site</td>
		<td width="80">UserID</td>
		<td width="65">������</td>
		<td width="65">������</td>
		<td width="72">�����ұݾ�</td>
		<td width="72">��븶�ϸ���</td>
		<td width="72">�������</td>
		<td width="140">�ֹ���</td>
		<td>������߼ۿ���</td>
	</tr>
<% if ojumun.FresultCount<1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="13" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for i=0 to ojumun.FresultCount-1 %>
	<form name="frmBuyPrc_<%=i%>" method="post" >
	<input type="hidden" name="orderidx" value="<%= ojumun.FItemList(i).FIdx %>">
	<input type="hidden" name="orderserial" value="<%= ojumun.FItemList(i).FOrderSerial %>">
	<input type="hidden" name="menupos" value="">
	<tr align="center" bgcolor="#FFFFFF">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><a href="#" onclick="ViewOrderDetail(frmBuyPrc_<%=i%>)" class="zzz"><%= ojumun.FItemList(i).FOrderSerial %></a></td>
		<td><%= ojumun.FItemList(i).FSitename %></td>
		<td><%= ojumun.FItemList(i).FUserID %></td>
		<td><%= ojumun.FItemList(i).FBuyName %></td>
		<td><%= ojumun.FItemList(i).FReqName %></td>
		<td><%= ojumun.FItemList(i).FSubTotalPrice %></td>
		<td><%= ojumun.FItemList(i).FMileTotalPrice %></td>
		<td><%= ojumun.FItemList(i).FTenCardSpend %></td>
		<td><%= ojumun.FItemList(i).FRegDate %></td>
		<td><%= ojumun.FItemList(i).FSendRegDate %></td>
	</tr>
	</form>
	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="13" height="30" align="center">
		<% if ojumun.HasPreScroll then %>
			<a href="?page=<%= ojumun.StarScrollPage-1 %>&daydiff=<%= daydiff %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ojumun.StarScrollPage to ojumun.FScrollCount + ojumun.StarScrollPage - 1 %>
			<% if i>ojumun.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&daydiff=<%= daydiff %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ojumun.HasNextScroll then %>
			<a href="?page=<%= i %>&daydiff=<%= daydiff %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
<% end if %>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->