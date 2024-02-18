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
' Discription : GIFT SHOP �׸� ����
' History : 2014.04.07 ������ : �ű� ����
'###############################################

'// ���� ����
Dim themeIdx, isUsing, isPick, isOpen
Dim oGiftShop, lp
Dim page

'// �Ķ���� ����
themeIdx = getNumeric(requestCheckVar(request("themeIdx"),10))
isusing = requestCheckVar(request("isusing"),1)
isPick = requestCheckVar(request("isPick"),1)
isOpen = requestCheckVar(request("isOpen"),1)
page = getNumeric(requestCheckVar(request("page"),10))
if isOpen="" then isOpen="Y"				'�⺻�� ������
if isPick="" then isPick="A"				'�⺻�� ��ü (A:�׸���ü, Y:�����׸�, N:���׸�)
if page="" then page="1"

'// ���������� ���
Set oGiftShop = new CGiftShop
oGiftShop.FPageSize=15
oGiftShop.FCurrPage=page
oGiftShop.FRectIsOpen = isOpen
oGiftShop.FRectIsUsing = isusing
oGiftShop.FRectIsPick = isPick
oGiftShop.GetThemeList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function() {
  	//�˻� ��ư
  	$("input[type=submit]").button();

  	// ������ư
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
  	$(".rdoOpen").buttonset().children().next().attr("style","font-size:11px;");
});

function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� �׸��� �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� �׸��� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.target="_self";
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}

function goThemeWrite(idx) {
    location.href = '/admin/sitemaster/gift/shop/giftshop_themeWrite.asp?themeidx='+idx+'&menupos=<%= request("menupos") %>';
}

</script>

<!-- ��� �˻��� ���� -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="page" value="" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
	    ��������:
		<select name="isPick" class="select">
			<option value="A" <%=chkIIF(isPick="A","selected","")%> >�׸���ü</option>
			<option value="Y" <%=chkIIF(isPick="Y","selected","")%> >10x10 Pick</option>
			<option value="N" <%=chkIIF(isPick="N","selected","")%> >User Pick</option>
		</select>
		&nbsp;/&nbsp;
	    ��������:
		<select name="isOpen" class="select">
			<option value="A" <%=chkIIF(isOpen="A","selected","")%> >��ü</option>
			<option value="Y" <%=chkIIF(isOpen="Y","selected","")%> >����</option>
			<option value="N" <%=chkIIF(isOpen="N","selected","")%> >�����</option>
		</select>
		&nbsp;/&nbsp;
	    ��뱸��:
		<select name="isusing" class="select">
			<option value="Y" <%=chkIIF(isusing="Y","selected","")%> >�����</option>
			<option value="N" <%=chkIIF(isusing="N","selected","")%> >������</option>
		</select>
	</td>
	<td width="80" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="�˻�" />
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
    <td align="left">
    	<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
    	<% if C_ADMIN_AUTH then %><input type="button" value="��������" class="button" onClick="saveList()" title="�켱���� �� ���⿩�θ� �ϰ������մϴ�."><% end if %>
    </td>
    <td align="right">
    	<input type="button" value="������ ���" class="button" onClick="goThemeWrite('');">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ��� ���� -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="chkAll" value="N">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�˻���� : <b><%=oGiftShop.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oGiftShop.FtotalPage%></b>
	</td>
</tr>
<colgroup>
    <col width="30" />
    <col width="50" />
    <col width="90" />
    <col width="*" />
    <col width="60" />
    <col width="60" />
    <col width="110" />
    <col width="90" />
    <col width="80" />
</colgroup>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>&nbsp;</td>
    <td>��ȣ</td>
    <td>��������<br>[Pick����]</td>
    <td>����</td>
    <td>��ǰ��</td>
    <td>�켱<br>����</td>
    <td>��������</td>
    <td>�����</td>
    <td>�����</td>
</tr>
<tbody id="mainList">
<%	for lp=0 to oGiftShop.FResultCount - 1 %>
<tr align="center" bgcolor="<%=chkIIF(oGiftShop.FItemList(lp).IsOpend,"#FFFFFF","#DDDDDD")%>">
    <td><input type="checkbox" name="chkIdx" value="<%=oGiftShop.FItemList(lp).FthemeIdx%>" /></td>
    <td><a href="javascript:goThemeWrite(<%=oGiftShop.FItemList(lp).FthemeIdx%>)"><%=oGiftShop.FItemList(lp).FthemeIdx%></a></td>
    <td><a href="javascript:goThemeWrite(<%=oGiftShop.FItemList(lp).FthemeIdx%>)"><%=oGiftShop.FItemList(lp).getPickType%></a>
    	<% if oGiftShop.FItemList(lp).FisPick="Y" then %><br><%=chkIIF(oGiftShop.FItemList(lp).FpickImage="" or isNull(oGiftShop.FItemList(lp).FpickImage),"<span style=""color:#608060;"">[�Ϲ���]</span>","<span style=""color:#806060;"">[�����]</span>")%><% end if %>
    </td>
    <td align="left"><a href="javascript:goThemeWrite(<%=oGiftShop.FItemList(lp).FthemeIdx%>)"><%=oGiftShop.FItemList(lp).FSubject & "<br><font color=""#606060"">" & oGiftShop.FItemList(lp).FSubDesc%></font></a></td>
    <td ><a href="javascript:goThemeWrite(<%=oGiftShop.FItemList(lp).FthemeIdx%>)"><%=oGiftShop.FItemList(lp).FitemCount%></a></td>
    <td><input type="text" name="sort<%=oGiftShop.FItemList(lp).FthemeIdx%>" size="3" class="text" value="<%=oGiftShop.FItemList(lp).FsortNo%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoOpen">
		<input type="radio" name="open<%=oGiftShop.FItemList(lp).FthemeIdx%>" id="rdoOpen<%=lp%>_1" value="Y" <%=chkIIF(oGiftShop.FItemList(lp).FisOpen="Y","checked","")%> /><label for="rdoOpen<%=lp%>_1">����</label><input type="radio" name="open<%=oGiftShop.FItemList(lp).FthemeIdx%>" id="rdoOpen<%=lp%>_2" value="N" <%=chkIIF(oGiftShop.FItemList(lp).FisOpen="N","checked","")%> /><label for="rdoOpen<%=lp%>_2">����</label>
		</span>
    </td>
    <td><%=left(oGiftShop.FItemList(lp).Fregdate,10)%></td>
    <td><%=oGiftShop.FItemList(lp).Fadminname%></td>
    </td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="9" align="center">
    <% if oGiftShop.HasPreScroll then %>
		<a href="javascript:goPage('<%= oGiftShop.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oGiftShop.StartScrollPage to oGiftShop.FScrollCount + oGiftShop.StartScrollPage - 1 %>
		<% if lp>oGiftShop.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oGiftShop.HasNextScroll then %>
		<a href="javascript:goPage('<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<!-- ��� �� -->
<%
	Set oGiftShop = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
