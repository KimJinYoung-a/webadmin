<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_main_EventBannerCls.asp" -->

<%
dim mode,i,page ,cdl , cdm , idx, cDisp, vCateCode
vCateCode = Request("catecode")
	mode = request("mode")
	page = request("page")
	idx = request("idx")	
	cdl = request("cdl")
	cdm = request("cdm")
	menupos = request("menupos")
%>

<script language="javascript">

function subcheck(){
	var frm=document.inputfrm;

	if (frm.catecode.value.length<1) {
		alert('ī�װ��� ������ �ּ���..');
		frm.catecode.focus();
		return;
	}
	
	if (frm.evt_code.value.length< 1 ){
		 alert('�̺�Ʈ ��ȣ�� �Է����ּ���');
	frm.evt_code.focus();
	return;
	}

	if (frm.viewidx.value.length< 1 ){
		 alert('���Ĺ�ȣ�� ���ڷ� �Է����ּ���.');
	frm.viewidx.focus();
	return;
	}

//	if (frm.cdl.value == '110'){
//		if (frm.cdm.value==''){
//			alert('����ä���� ��ī�װ��� �����ؾ߸� �մϴ�');			
//			return;
//		}
//	}

	frm.submit();
}

function chimg(im,v){

	frm=eval("document." + v);
	frm.src=im;
}

function popEventList(){
	var frm=document.inputfrm;

	if (frm.catecode.value.length<1) {
		alert('ī�װ��� �������ּ���..');
		frm.catecode.focus();
		return;
	}
	
	window.open('ViewEventList_Main_EventBanner.asp?selC=' + frm.catecode.value,'pop','width=800,height=600,scrollbars=yes');
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
	<%
	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	'cDisp.FRectUseYN = "Y"
	cDisp.GetDispCateList()
	
	If cDisp.FResultCount > 0 Then
		Response.Write "<select name=""catecode"" class=""select"">" & vbCrLf
		Response.Write "<option value="""">����</option>" & vbCrLf
		For i=0 To cDisp.FResultCount-1
			Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
		Next
		Response.Write "</select>&nbsp;&nbsp;&nbsp;"
	End If
	Set cDisp = Nothing
	%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">�̺�Ʈ ��ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="evt_code" size="8">
	<!--<input type="button" name="evtbtn" class="button" value="�˻�" onclick="popEventList();">
	//--></td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">���Ĺ�ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="viewidx" size="4" value="99"></td>
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
	<%
	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	cDisp.GetDispCateList()
	
	If cDisp.FResultCount > 0 Then
		Response.Write "<select name=""catecode"" class=""select"">" & vbCrLf
		Response.Write "<option value="""">����</option>" & vbCrLf
		For i=0 To cDisp.FResultCount-1
			Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(fmainitem.FItemList(0).Fevt_disp)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
		Next
		Response.Write "</select>&nbsp;&nbsp;&nbsp;"
	End If
	Set cDisp = Nothing
	%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">�̺�Ʈ��</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="evt_code" size="8" value="<%=fmainitem.FItemList(0).Fevt_code%>">
		<%= fmainitem.FItemList(0).Fevt_name %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">���Ĺ�ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="viewidx" size="4" value="<%=fmainitem.FItemList(0).FviewIdx%>"></td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%if fmainitem.FItemList(0).FIsusing="Y" then response.write "checked" %>>Y
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
