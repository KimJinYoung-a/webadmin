<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 �ѿ�� 2008����Ʈ�����̵� 2009������ ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/category_main_EventBannerCls.asp"-->

<%
dim mode,i,page ,cdl , cdm , idx
	mode = request("mode")
	page = request("page")
	idx = request("idx")	
	cdl = request("cdl")
	cdm = request("cdm")	
%>

<script language="javascript">

function subcheck(){
	var frm=document.inputfrm;

	if (frm.cdl.value.length<1) {
		alert('ī�װ��� �������ּ���..');
		frm.cdl.focus();
		return;
	}
	
	if (frm.evt_code.value.length< 1 ){
		 alert('�̺�Ʈ ��ȣ�� �Է����ּ���');
	frm.evt_code.focus();
	return;
	}

	if (frm.viewidx.value.length< 1 ){
		 alert('ǥ�ü����� ���ڷ� �Է����ּ���.');
	frm.viewidx.focus();
	return;
	}

	if (frm.cdl.value == '110'){
		if (frm.cdm.value==''){
			alert('����ä���� ��ī�װ��� �����ؾ߸� �մϴ�');			
			return;
		}
	}

	frm.submit();
}

function chimg(im,v){

	frm=eval("document." + v);
	frm.src=im;
}

function popEventList(){
	var frm=document.inputfrm;

	if (frm.cdl.value.length<1) {
		alert('ī�װ��� �������ּ���..');
		frm.cdl.focus();
		return;
	}
	
	window.open('ViewEventList_Main_EventBanner.asp?selC=010','popasd','width=800,height=600,scrollbars=yes');
}

function changecontent()
{
	document.inputfrm.action='?';
	document.inputfrm.submit();

}

</script>
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="20">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top"><b>ī�װ� ���� �̺�Ʈ ���� ���/����</b></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="inputfrm" method="post" action="doMainEventBanner.asp">
<input type="hidden" name="mode" value="<% =mode %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<% if mode="add" then %>
<tr>
	<td width="100" bgcolor="#F0F0FD" align="center">ī�װ�����</td>
	<td bgcolor="#FFFFFF">
<select class='select' name="cdl">
<option value='010' selected>�����ι���</option>
</select>
<select class='select' name="cdm">
<option value='' <% If cdm = "" Then Response.Write "selected" End If %>>��ü���̾</option>
<option value='10' <% If cdm = "10" Then Response.Write "selected" End If %>>���ô��̾</option>
<option value='20' <% If cdm = "20" Then Response.Write "selected" End If %>>�Ϸ���Ʈ���̾</option>
<option value='30' <% If cdm = "30" Then Response.Write "selected" End If %>>ĳ���ʹ��̾</option>
<option value='40' <% If cdm = "40" Then Response.Write "selected" End If %>>������̾</option>
</select>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">�̺�Ʈ ��ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="evt_code" size="8"><input type="button" name="evtbtn" class="button" value="�˻�" onclick="popEventList();"></td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">ǥ�ü���</td>
	<td bgcolor="#FFFFFF"><input type="text" name="viewidx" size="4"></td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" checked>Y
		<input type="radio" name="isusing" value="N">N
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td colspan="2" align="center">
			<input type="button" value=" ���� " onclick="subcheck();"> &nbsp;&nbsp;
			<input type="button" value=" ��� " onclick="history.back();">
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New CateEventBanner
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.frectidx = idx
	fmainitem.GetEventBannerList

if cdl = "" then cdl = fmainitem.FItemList(0).fcdl
if cdm = "" then cdm = fmainitem.FItemList(0).Fcdm
%>
<tr>
	<td width="100" align="center" bgcolor="#F0F0FD">ī�װ�</td>
	<td bgcolor="#FFFFFF">
<select class='select' name="cdl">
<option value='010' selected>�����ι���</option>
</select>
<select class='select' name="cdm">
<option value='' <% If cdm = "" Then Response.Write "selected" End If %>>��ü���̾</option>
<option value='10' <% If cdm = "10" Then Response.Write "selected" End If %>>���ô��̾</option>
<option value='20' <% If cdm = "20" Then Response.Write "selected" End If %>>�Ϸ���Ʈ���̾</option>
<option value='30' <% If cdm = "30" Then Response.Write "selected" End If %>>ĳ���ʹ��̾</option>
<option value='40' <% If cdm = "40" Then Response.Write "selected" End If %>>������̾</option>
</select>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">�̺�Ʈ��</td>
	<td bgcolor="#FFFFFF">
		<%="[" & fmainitem.FItemList(0).Fevt_code & "] " & fmainitem.FItemList(0).Fevt_name %>
		<input type="hidden" name="evt_code" value="<%=fmainitem.FItemList(0).Fevt_code%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">ǥ�ü���</td>
	<td bgcolor="#FFFFFF"><input type="text" name="viewidx" size="4" value="<%=fmainitem.FItemList(0).FviewIdx%>"></td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%if fmainitem.FItemList(0).FIsusing="Y" then response.write "checked" %> checked>Y
		<input type="radio" name="isusing" value="N" <%if fmainitem.FItemList(0).FIsusing="N" then response.write "checked" %>>N
		<input type="hidden" name="orgUsing" value="<%=fmainitem.FItemList(0).FIsusing%>">
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " onclick="subcheck();"> &nbsp;&nbsp;
		<input type="button" value=" ��� " onclick="history.back();">
	</td>
</tr>
<% end if %>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
