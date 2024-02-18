<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/offshop/offshop_staffcls.asp" -->
<%
dim idx,nstaff,mode
	mode = requestCheckVar(request("mode"),32)
	idx = requestCheckVar(request("idx"),10)

if idx = "" then idx=0

set nstaff = new COffshopStaffDetail
nstaff.GetOffshopStaff idx

dim yyyy1,mm1,dd1,datearr

if mode = "add" then
yyyy1 = Cstr(Year(now()))
mm1 = Cstr(Month(now()))
dd1 = Cstr(day(now()))
else
datearr = split(left(nstaff.Fipsadate,10),"-")
yyyy1 = datearr(0)
mm1 = datearr(1)
dd1 = datearr(2)
end if

%>
<script type='text/javascript'>
<!--

function GoReplyWrite(){
//    alert('������');
//    return;
    
    var frm = document.boardfrm;
	if (frm.shopid.value == ""){
		alert("�������� ���ּ���");
		frm.shopid.focus();
	}
	else if (frm.username.value == ""){
		alert("������ �Է����ּ���");
		frm.username.focus();
	}
	else if (!frm.slevel.value){
		alert("������ ������ �ּ���");
		frm.slevel.focus();
	}
	else if (frm.contents.value == ""){
		alert("������ �Է����ּ���");
		frm.contents.focus();
	}	
	else{
		frm.submit();
	}
}

//-->
</script>

<table border="0" cellpadding="0" cellspacing="1" width="700" bgcolor="#808080" class="a" align="center">
<form method="post" name="boardfrm" action="<%= uploadImgUrl %>/linkweb/offshop/OffshopStaff_process.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="backmode" value="on">
<input type="hidden" name="idx" value="<% = idx %>">
<tr>
	<td bgcolor="#FFFFFF" height="30" align="center">������</td>
	<td bgcolor="#FFFFFF">&nbsp;
		<select name="shopid">
			<option value="">����</option>
			<%Call fnOptShopName(nstaff.Fshopid)%>			
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">�����̸�</td>
	<td>&nbsp;<input type="text" name="username" size="50" class="input_b" value="<% = nstaff.Fusername %>"></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">����</td>
	<td>&nbsp;<select name="slevel">
		<option value="">--����--</option>
		<%Call fnOptCommonCode("stafflevel",nstaff.Flevel)%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">�Ի���</td>
	<td>&nbsp;<% DrawOneDateBox yyyy1,mm1,dd1 %></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">����</td>
	<td>&nbsp;<textarea name="contents" rows="20" cols="70" class="input_b"><% = nstaff.Fcontents %></textarea></td>
</tr>
<% if mode="add" then %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">÷�λ���</td>
	<td>&nbsp;<input type="file" name="file1" size="50" class="input_b"></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">÷�λ���</td>
	<td>
	&nbsp;<input type="file" name="file1" size="50" class="input_b"><br>
	&nbsp;<input type="checkbox" name="dl_file1">���ϻ��� <img src="<% = nstaff.Ficon1 %>" width="50" height="60" border="0">
	</td>
</tr>
<% end if %>
<% if mode="edit" then %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">�������</td>
	<td>
		<input type="radio" name="isusing" value="Y" <% if nstaff.Fisusing = "Y" then response.write "checked" %>>Y <input type="radio" name="isusing" value="N" <% if nstaff.Fisusing = "N" then response.write "checked" %>>N
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="right" colspan="2"><a href="javascript:GoReplyWrite();"><font color="red">�۾���</font></a>&nbsp;&nbsp;&nbsp;</td>
</tr>
</form>
</table>

<% set nstaff = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
