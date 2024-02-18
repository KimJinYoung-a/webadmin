<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/managerCls.asp"-->
<%
Dim mode, idx, makerid, mlookbook, didx, regdate, lastupdate, brandgubun, regadminid, lastadminid
dim subtopimage, designis
	mode	= request("mode")
	idx		= request("idx")
	makerid	= request("makerid")
	menupos	= request("menupos")

If idx = "" Then
	mode = "I"
Else
	mode = "U"
End If

Dim omanager
SET omanager = new cmanager
	omanager.FrectIdx = idx
	
	if idx <> "" then
		omanager.sbmanagermodify
		
		if omanager.ftotalcount > 0 then
        	idx = omanager.FOneItem.Fidx
        	makerid = omanager.FOneItem.Fmakerid
        	regdate = omanager.FOneItem.Fregdate
        	lastupdate = omanager.FOneItem.Flastupdate
        	brandgubun = omanager.FOneItem.Fbrandgubun
        	regadminid = omanager.FOneItem.Fregadminid
        	lastadminid = omanager.FOneItem.Flastadminid
        	subtopimage = omanager.FOneItem.Fsubtopimage
        	designis = omanager.FOneItem.fdesignis
		end if
	end if

if brandgubun="" then brandgubun="1"
%>

<script language="javascript">

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
	if(frm.brandgubun.options[frm.brandgubun.selectedIndex].value==""){
		alert('�귣�屸���� �����ϼ���');
		frm.brandgubun.focus();
		return;
	}

	if(confirm('�����Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

</script>

<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>�귣�屸������</b>

<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/street/domanager_reg.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="registerID" value="<%=session("ssBctId")%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td width="150" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
			<td bgcolor="#FFFFFF">
				<%= idx %>				
			</td>
		</tr>
		<tr >
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�귣��</td>
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
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�귣�屸��</td>
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxbrandgubun "brandgubun",brandgubun , "" %>

				<% if idx = "" or isnull(idx) then %>
					(������)
				<% end if %>				
			</td>
		</tr>
		<tr >		
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�̹���<br>(�����̾��귣������)</td>
			<td bgcolor="#FFFFFF">
				<input type="file" name="subtopimage" value="" size="55"><br>
				<% If subtopimage <> "" Then %>
				<img src="<%=uploadUrl%>/brandstreet/manager/<%=subtopimage%>">
				<br>Filename : http://testimgstatic.10x10.co.kr/brandstreet/manager/<%=subtopimage%><br>
				<% End If %>
				
				<% if designis <> "" then %>
					<Br>�������̶�?(Hello) : <%= designis %>
				<% end if %>
			</td>
		</tr>
		<tr >
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">��������</td>
			<td bgcolor="#FFFFFF">
				<% if lastadminid<>"" and not isnull(lastadminid) then %>
					<%= lastupdate %>
					<br>(<%= lastadminid %>)
				<% end if %>
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
Set omanager = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->