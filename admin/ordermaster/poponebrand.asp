<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ������Ѻ귣��. �������Ѻ귣��
' History : �̻� ����
'			2023.09.13 �ѿ�� ����(����üũ �߰�. ǥ���ڵ����� �ҽ� ����)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/tenbalju.asp"-->
<%
dim mode, brandid, reguserid, page, divcd, found, sqlStr, odanpumbalju, i, menupos
	page = requestCheckVar(getNumeric(request("page")), 10)
	menupos = requestCheckVar(getNumeric(request("menupos")), 10)
	mode = requestCheckVar(request("mode"), 32)
	brandid = requestCheckVar(trim(request("brandid")), 32)
	divcd = requestCheckVar(request("divcd"), 1)

reguserid = session("bctid")

if page="" then page=1

if (mode="del") and (brandid<>"") then
    sqlStr = "delete from [db_item].[dbo].tbl_baljureg_brand where" + VbCrlf
    sqlStr = sqlStr + " trim(brandid)='" + Cstr(brandid) + "' "

	'response.write sqlStr & "<br>"
    dbget.Execute sqlStr

    brandid = ""
	response.write "<script type='text/javascript'>"
	response.write "	location.replace('/admin/ordermaster/poponebrand.asp?mode=&page=&menupos="& menupos &"&divcd=&brandid=');"
	response.write "</script>"
end if

if (mode="add") and (brandid<>"") and (divcd<>"") then
    sqlStr = " select count(brandid) as cnt "
    sqlStr = sqlStr + " from [db_item].[dbo].tbl_baljureg_brand with (nolock)"
    sqlStr = sqlStr + " where brandid = '" & CStr(brandid) & "' "

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	found = rsget("cnt")>0
	rsget.close
	
	if (not found) then
	    sqlStr = " insert into [db_item].[dbo].tbl_baljureg_brand " + VbCrlf
	    sqlStr = sqlStr + " (brandid, divcd, reguserid) "+ VbCrlf
	    sqlStr = sqlStr + " values('" & CStr(brandid) & "', '" & CStr(divcd) & "', '" & CStr(reguserid) & "') "

		'response.write sqlStr & "<br>"
	    dbget.Execute sqlStr
	    
	    brandid = ""
        response.write "<script type='text/javascript'>"
		response.write "	location.replace('/admin/ordermaster/poponebrand.asp?mode=&page=&menupos="& menupos &"&divcd=&brandid=');"
		response.write "</script>"
	else
		response.write "<script type='text/javascript'>alert('�̹� ��ϵ� �귣���Դϴ�.');</script>"
	end if

end if

set odanpumbalju = new CTenBalju
	odanpumbalju.FPageSize=50
	odanpumbalju.FCurrpage = page
	odanpumbalju.FRectBrandDivCD = divcd
	odanpumbalju.FRectbrandid=brandid
	odanpumbalju.GetDanpumBaljuBrandList

%>
<script type='text/javascript'>

function DelItem(brandid){
   if (confirm('���� �Ͻðڽ��ϱ�?')){
        dellfrm.mode.value="del";
        dellfrm.brandid.value= brandid;
        dellfrm.submit();
    }
}

function AddItem(frm){
    if (frm.divcd.value == ""){
        alert('������ �����ϼ���.');
        return;
    }

    if (frm.brandid.value.length<3){
        alert('�귣����̵� ��Ȯ�� �Է��ϼ���.');
        frm.brandid.focus();
        return;
    }

    frm.mode.value="add";
    frm.submit();
}

function NextPage(page){
    frmbar.page.value=page;
    frmbar.submit();
}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		����/���� �귣�� ����
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- �˻� ���� -->
<form name="frmbar" method="get" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<select name="divcd" >
			<option value="" <% if divcd="" then response.write "selected" %> >����</option>
			<option value="E" <% if divcd="E" then response.write "selected" %> >���ܺ귣��</option>
			<option value="I" <% if divcd="I" then response.write "selected" %> >���Ժ귣��</option>
		</select>
		�귣����̵� : <input type="text" name="brandid" value="<%= brandid %>" Maxlength="20" size="13" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ AddItem(frmbar); return false;}">
		<input type="button" value="�귣���߰�" onclick="AddItem(frmbar)" class="button">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="NextPage('1');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�˻���� : <b><%= odanpumbalju.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= odanpumbalju.FTotalPage %></b>
	</td>
</tr>
<% if odanpumbalju.FresultCount>0 then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">����</td>
		<td width="80">�귣����̵�</td>
		<td>�귣���</td>
		<td>��ü��</td>
		<td width="50">����</td>
	</tr>
	<% for i=0 to odanpumbalju.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= odanpumbalju.FItemList(i).GetDivCDString %></td>
		<td><%= odanpumbalju.FItemList(i).Fbrandid %></td>
		<td><%= odanpumbalju.FItemList(i).Fsocname_kor %><br><%= odanpumbalju.FItemList(i).Fsocname %></td>
		<td><%= odanpumbalju.FItemList(i).Fcompany_name %></td>
		<td>
			<input type="button" value="����" onclick="DelItem('<%= odanpumbalju.FItemList(i).Fbrandid %>');" class="button">
		</td>
	</tr>   
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
			<% if odanpumbalju.HasPreScroll then %>
				<a href="javascript:NextPage('<%= odanpumbalju.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + odanpumbalju.StartScrollPage to odanpumbalju.FScrollCount + odanpumbalju.StartScrollPage - 1 %>
				<% if i>odanpumbalju.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if odanpumbalju.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="16" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>

</table>
<form name="dellfrm" method="get" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="brandid" value="">
</form>

<%
set odanpumbalju = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->