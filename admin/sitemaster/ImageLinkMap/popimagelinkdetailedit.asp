<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : popimagelinkdetailedit.asp
' Discription : �̹��� ���� ��ũ �� �Է�
' History : 2019.08.06 ������ : �ű��ۼ�
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/imageLinkCls.asp" -->
<%
Dim isusing, fixtype, validdate, prevDate
Dim idx, poscode, reload, gubun, edid
Dim culturecode, masterIdx, itemid, posX, posY
	idx = request("idx")
    masterIdx = request("masteridx")
	poscode = request("poscode")
    itemid = request("itemid")
    posX = request("posX")
    posY = request("posY")

    If masterIdx = "" Then
        response.write "<script>alert('�������� ��η� ������ �ּ���');history.back();</script>"
        response.end
    End If

    If idx = "" Then
        idx = 0
    End If

	dim oLinkDetailContents
    set oLinkDetailContents = new CimageLink
    oLinkDetailContents.FRectIdx = idx
    oLinkDetailContents.GetOneDetailContents

    If itemid = "" then
        itemid = oLinkDetailContents.FOneItem.FItemid
    End If


%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

	function SaveMainContents(frm){

		if (frm.Itemid.value==""){
	        alert('��ǰ�ڵ带 �Է����ּ���.');
	        frm.Itemid.focus();
	        return;
	    }

	    if (confirm('���� �Ͻðڽ��ϱ�?')){
	        frm.submit();
	    }
	}

    function popRegSearchItem() {
    <% if masterIdx <> "" then %>
        var popwinsub = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=<%=Server.URLEncode("/admin/sitemaster/ImageLinkMap/doSubRegItemCdArray.asp?listidx="&masterIdx&"&idx="&idx&"&posX="&posX&"&posY="&posY)%>", "popup_imagelinkitemsub", "width=800,height=500,scrollbars=yes,resizable=yes");
        popwinsub.focus();
    <% else %>
        alert("���� �̹����� ���� ������ּ���.");
    <% end if %>
    }

</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doLinkImageRegDetail.asp" onsubmit="return false;">
<input type="hidden" name="masteridx" value="<%=masterIdx%>">
<input type="hidden" name="posX" value="<%=posX%>">
<input type="hidden" name="posY" value="<%=posY%>">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF" align="center">Idx</td>
    <td>
        <% if oLinkDetailContents.FOneItem.Fidx<>"" then %>
        <%= oLinkDetailContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oLinkDetailContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="200" bgcolor="#DDDDFF" align="center">��ǰ�ڵ�</td>
    <td>
		<input type="text" name="Itemid" value="<%=itemid%>">&nbsp;
        <input type="button" value="��ǰ�߰�" onclick="popRegSearchItem();">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="200" bgcolor="#DDDDFF" align="center">������ ����</td>
    <td>
        <select name="IconType">
            <option value="1" <% If oLinkDetailContents.FOneItem.FIconType = "1" Then %>selected<% End If %>>�����</option>
            <option value="2" <% If oLinkDetailContents.FOneItem.FIconType = "2" Then %>selected<% End If %>>���׶��</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF" align="center">��뿩��</td>
  <td>
  	<input type="radio" name="Isusing" value="Y"<% If oLinkDetailContents.FOneItem.FIsusing="Y" Or oLinkDetailContents.FOneItem.FIsusing="" Then Response.write " checked" %>> �����
	<input type="radio" name="Isusing" value="N"<% If oLinkDetailContents.FOneItem.FIsusing="N" Then Response.write " checked" %>> ������
  </td>
</tr>
<% If oLinkDetailContents.FOneItem.Fadminid<>"" Then %>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">�۾���</td>
  <td>
  	�۾��� : <%=oLinkDetailContents.FOneItem.Fadminid %><br>
	�����۾��� : <%=oLinkDetailContents.FOneItem.Flastadminid %>
  </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" �� �� " onclick="opener.location.reload();window.close();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oLinkDetailContents = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
