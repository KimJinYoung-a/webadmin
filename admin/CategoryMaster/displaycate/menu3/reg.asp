<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMenuCls.asp"-->

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="http://webadmin.10x10.co.kr/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="http://webadmin.10x10.co.kr/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="http://webadmin.10x10.co.kr/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<Script>
function checkform(f){
	if(f.disp1.value == ""){
		alert("ī�װ��� �����ϼ���.");
		return false;
	}
	if(f.linkurl.value == ""){
		alert("��ũ�� �Է��ϼ���.");
		f.linkurl.focus();
		return false;
	}
	if(f.sortno.value == ""){
		alert("���Ĺ�ȣ�� �Է��ϼ���.");
		f.sortno.focus();
		return false;
	}
	return true;
}

function jsSetImg(){
	if(frm1.disp1.value == ""){
		alert("ī�װ��� �����ϼ���.");
		return false;
	}
	
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('pop_uploadimg.asp?disp='+frm1.disp1.value,'popImg','width=370,height=150');
	winImg.focus();
}
</script>

<%
	Dim cMenu, i, vArr, vIdx, vDisp1, vType, vSubject, vItemID, vLinkURL, vSDate, vEDate, vUseYN, vSortNo, vRegdate, vImgURL, vOrderText
	vIdx = Request("idx")
	vDisp1 = Request("disp1")
	If vSortNo = "" Then vSortNo = "0" End If
	If vUseYN = "" Then vUseYN = "y" End If
	If vIdx <> "" Then
		rsget.Open "select * from db_item.dbo.tbl_display_cate_menu_top where idx = '" & vIdx & "'",dbget,1
		vDisp1		= rsget("disp1")
		vLinkURL	= rsget("linkurl")
		vImgURL		= rsget("imgurl")
		vUseYN 		= rsget("useyn")
		vSortNo		= rsget("sortno")
		vRegdate 	= rsget("regdate")
		vOrderText	= db2html(rsget("ordertext"))
		rsget.Close
	End IF
	
	Set cMenu = New cDispCateMenu
	vArr = cMenu.GetDispCate1Depth()
	Set cMenu = Nothing
%>

<form name="frm1" action="proc.asp" method="post" onSubmit="return checkform(this);">
<input type="hidden" name="idx" value="<%=vIdx%>">
<input type="hidden" name="type" value="topbanner">
<input type="hidden" name="itemid" value="0">
<input type="hidden" name="subject" value="">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% If vIdx <> "" Then %>
<tr bgcolor="#FFFFFF">
	<td>idx</td>
	<td><%=vIdx%> (�����:<%=vRegdate%>)</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
	<td>ī�װ�</td>
	<td>
		<select name="disp1" class="select">
		<option value="">-ī�װ�����-</option>
		<%
			For i=0 To UBound(vArr,2)
				Response.Write "<option value='" & vArr(0,i) & "' " & CHKIIF(CStr(vDisp1)=CStr(vArr(0,i)),"selected","") & ">" & vArr(1,i) & "</option>" & vbCrLf
			Next
		%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>�̹���</td>
	<td>		
		<span id="imgspan"><% IF vImgURL <> "" THEN %><img src="<%=vImgURL%>" height="70"><%END IF%></span>
		&nbsp;&nbsp;&nbsp;
		<input type="button" value="�̹������ε�" onClick="jsSetImg()">
		<input type="hidden" name="imgurl" value="<%=vImgURL%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��ũ</td>
	<td>
		<input type="text" name="linkurl" class="text" value="<%=vLinkURL%>" size="60">
		<br>* http://www.10x10.co.kr �� ���� / ���� �Է��ϼ���.
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>���Ĺ�ȣ</td>
	<td>
		<input type="text" name="sortno" class="text" value="<%=vSortNo%>" size="5">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��뿩��</td>
	<td>
		<select name="useyn" class="select">
			<option value="y" <%=CHKIIF(vUseYN="y","selected","")%>>���</option>
			<option value="n" <%=CHKIIF(vUseYN="n","selected","")%>>������</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>�۾�����</td>
	<td><textarea name="ordertext" cols="100" rows="10" class="textarea"><%=vOrderText%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td></td>
	<td style="padding-top:20px;">
		<input type="submit" value="��  ��" style="height:30px;">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->