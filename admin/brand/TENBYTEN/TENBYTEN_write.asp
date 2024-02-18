<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.19 ������ ����
'			2013.08.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/TENBYTENCls.asp"-->
<%
Dim mode, idx, makerid, mlookbook, didx
	mode	= request("mode")
	idx		= request("idx")
	makerid	= request("makerid")
	menupos	= request("menupos")

If idx = "" Then
	mode = "I"
Else
	mode = "U"
End If

Dim oTENBYTEN
SET oTENBYTEN = new cTENBYTEN
	oTENBYTEN.FIdx = idx
	oTENBYTEN.sbTENBYTENmodify

makerid = oTENBYTEN.FOneitem.FMakerid
%>

<script language="javascript">

function flag_select(f){
	if(f == '1'){
		document.getElementById('img').style.display="block";
		document.getElementById('play').style.display="none";
	}else if(f == '2'){
		document.getElementById('img').style.display="none";
		document.getElementById('play').style.display="block";
	}
}

function subcheck(){
	var frm=document.frm;
	if("<%=mode%>" == "U" ){
		frm.mode.value ="U"
	}
	if(frm.makerid.value==""){
		alert('�귣�带 �����ϼ���');
		frm.makerid.focus();
		return;
	}

	if(confirm('�����Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

</script>

<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>TENBYTEN ���</b>

<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/street/doTENBYTEN_reg.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="registerID" value="<%=session("ssBctId")%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�귣��</td>
			<td bgcolor="#FFFFFF">
				<% if mode = "U" then %>
					<%=makerid%>
					<input type="hidden" name="makerid" value="<%=makerid%>">
				<% else %>
					<% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
			</td>
		</tr>
		<tr >
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̹���/������</td>
			<td bgcolor="#FFFFFF">
			<% If mode = "I" Then %>
				<select name="flag" class="select" onchange="javascript:flag_select(this.value)">
					<option value="1">�̹���</option>
					<option value="2">������</option>
				</select>
			<% ElseIf mode = "U" Then %>
				<select name="flag" class="select" onchange="javascript:flag_select(this.value)">
					<option value="1" <%=Chkiif(oTENBYTEN.FOneitem.FFlag = "1","selected","disabled")%> >�̹���</option>
					<option value="2" <%=Chkiif(oTENBYTEN.FOneitem.FFlag = "2","selected","disabled")%> >������</option>
				</select>
			<% End If %>
			</td>
		</tr>
	<% If mode = "I" Then %>
		<tr id="img">
	<% ElseIf oTENBYTEN.FOneitem.FFlag = "1" Then %>
		<tr id="img">
	<% ElseIf oTENBYTEN.FOneitem.FFlag = "2" Then %>
		<tr id="img" style="display:none;">
	<% End If %>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
			<td bgcolor="#FFFFFF">
				<input type="file" name="imgurl" value="" size="55"><br>
				
				<% If oTENBYTEN.FOneItem.FImgurl <> "" Then %>
					<img src="<%=uploadUrl%>/brandstreet/TENBYTEN/<%=oTENBYTEN.FOneItem.FImgurl%>">
					<br>Filename : http://testimgstatic.10x10.co.kr/brandstreet/TENBYTEN/<%=oTENBYTEN.FOneItem.FImgurl%><br>
				<% End If %>
				
				Ŭ��URL : <input type="text" name="linkurl" value="<%=oTENBYTEN.FOneitem.FLinkurl%>" size="80" maxlength=80 class="text">
			</td>
		</tr>
	<% If oTENBYTEN.FOneitem.FFlag = "2" Then %>
		<tr id="play">
	<% Else %>
		<tr id="play" style="display:none;">
	<% End If %>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">������ũ�ּ�</td>
			<td bgcolor="#FFFFFF">
				�� ������ �ּҸ� �Է��� �ּ���.
				<Br>ex) <font color="red">http://www.youtube.com/embed/sDwatprn1mo?wmode=opaque</font>
				<Br><textarea name="playurl" rows="10" cols="69"><%=oTENBYTEN.FOneitem.FPlayurl%></textarea>
			</td>
		</tr>
		<tr height="30" align="center">
			<td bgcolor="#FFFFFF" colspan="2">
				<input type="button" value="����" class="button" onclick="javascript:subcheck();">
			</td>
		</tr>
	</td>
</tr>
</table>
</form>

<%
Set oTENBYTEN = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->