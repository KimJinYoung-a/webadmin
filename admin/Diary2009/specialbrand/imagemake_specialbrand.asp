<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����� �귣�� ����Ʈ ���/����
' History : 2016.09.07 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/diary2009/classes/specialbrandCls.asp"-->

<%
dim reload , ix, tmp
dim idx, makerid, brandtext, mainbrandimg, brandmovieurl, itemid, isusing, sortnum, regdate
dim pcmainbrandtextimg, momainbrandimg, leftright
	idx = request("idx")
	reload = request("reload")
	if idx="" then idx=0

if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End
end if

dim oMainContents
	set oMainContents = new DiaryCls
	oMainContents.FRectIdx = idx
	oMainContents.fcontents_oneitem

	idx				=	oMainContents.FOneItem.fidx
	makerid		=	oMainContents.FOneItem.fbrandid
	brandtext		=	oMainContents.FOneItem.fbrandtext
	mainbrandimg	=	oMainContents.FOneItem.fmainbrandimg
	brandmovieurl	=	oMainContents.FOneItem.fbrandmovieurl
	itemid			=	oMainContents.FOneItem.fitemid
	isusing		=	oMainContents.FOneItem.fisusing
	sortnum		=	oMainContents.FOneItem.fsortnum
	regdate		=	oMainContents.FOneItem.fregdate

	pcmainbrandtextimg		=	oMainContents.FOneItem.fpcmainbrandtextimg
	momainbrandimg		=	oMainContents.FOneItem.fmomainbrandimg
	leftright		=	oMainContents.FOneItem.Fleftright


	if sortnum < 1 then
		sortnum = 99
	end if
%>
<script language='javascript'>

function SaveMainContents(frm){
    if (frm.makerid.value.length<1){
        alert('�귣�� id�� �Է��� �ּ���.');
        frm.makerid.focus();
        return;
    }

    if (frm.itemid.value.length<1){
        alert('��ǰ�ڵ带 �Է����ּ���\n(�� : 11111,22222,33333');
        frm.itemid.focus();
        return;
    }

    <% if idx < 1 then %>
	    if (frm.file1.value.length<1){
	        alert('��ǥ �̹����� ������ּ���.1');
	        frm.file1.focus();
	        return;
	    }
	<% else %>
	    if (frm.fileval.value.length<1){
	        alert('��ǥ �̹����� ������ּ���.2');
	        frm.fileval.focus();
	        return;
	    }
	<% end if %>

    if (frm.sortnum.value.length<1){
        alert('�̹��� �켱������ �Է� �ϼ���.');
        frm.sortnum.focus();
        return;
    }
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);" class="button">
		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/diary/specialbrand/specialbrand_image_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="ckUserId" value="<%=session("ssBctId")%>">
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">Idx :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	       	 <%= idx %>
	       	 <input type="hidden" name="idx" value="<%= idx %>">
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">�귣��ID :<br><font color="red">(�ʼ�)</font></td>
	    <td style="white-space:nowrap;"><%	drawSelectBoxDesignerWithName "makerid", makerid %> </td>
	  <!--  <td>
	    	 <input type="text" name="brandid" maxlength="32" value="<%'= brandid %>">
	    </td>
	   -->
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ǰ�ڵ� :<br><font color="red">(�ʼ�)</font></td>
	    <td>
			<input type="text" name="itemid" size="80" maxlength="63" value="<%= itemid %>">
			<br>
			<font color="red">
				�� �ִ� 8��<br>
				�� �޸��� ���� ex : 1111111,2222222,3333333
			</font>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">PC ��ǥ�̹��� :<br><font color="red">(�ʼ�)</font></td>
	  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
		  <br><img src="<%=uploadUrl%>/diary/specialbrand/<%= mainbrandimg %>" border="0">
		  <br><%=uploadUrl%>/diary/specialbrand/<%= mainbrandimg %>
	  <% end if %>
	  <input type="hidden" name="fileval" value="<%= mainbrandimg %>">
	  </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">PC ���� �̹��� :<br><font color="red">(�ʼ�)</font></td>
	  <td><input type="file" name="file2" value="" size="32" maxlength="32" class="file">
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
		  <br><img src="<%=uploadUrl%>/diary/specialbrand/<%= pcmainbrandtextimg %>" border="0">
		  <br><%=uploadUrl%>/diary/specialbrand/<%= pcmainbrandtextimg %>
	  <% end if %>
	  <input type="hidden" name="filevalpctext" value="<%= pcmainbrandtextimg %>">
	  </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��-�� ����(PC) :</td>
	    <td>
		<% if oMainContents.FOneItem.Fleftright="R" then %>
			<input type="radio" name="leftright" value="L">��
			<input type="radio" name="leftright" value="R" checked>��
		<% else %>
			<input type="radio" name="leftright" value="L" checked>��
			<input type="radio" name="leftright" value="R">��
		<% end if %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">����� ��ǥ �̹��� :<br><font color="red">(�ʼ�)</font></td>
	  <td><input type="file" name="file3" value="" size="32" maxlength="32" class="file">
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
		  <br><img src="<%=uploadUrl%>/diary/specialbrand/<%= momainbrandimg %>" border="0">
		  <br><%=uploadUrl%>/diary/specialbrand/<%= momainbrandimg %>
	  <% end if %>
	  <input type="hidden" name="filevalmo" value="<%= momainbrandimg %>">
	  </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">����� �귣�弳�� :</td>
	    <td>
	    	 <input type="text" name="brandtext" size="80" maxlength="150" value="<%= brandtext %>">
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">������URL :</td>
	    <td>
			<input type="text" name="brandmovieurl" size="80" value="<%= brandmovieurl %>">
			<br>
			<font color="red">
				�� ������ : �ҽ��ڵ� ���� (�� : </font><font color="blue">http://www.youtube.com/embed/qj4rn1I_dC8 </font><font color="red">)<br>
				�� ��޿� : copy embed code ���� (�� :</font><font color="blue"> //player.vimeo.com/video/102309330</font><font color="red"> ) http: ����
			</font>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">�켱���� :</td>
	    <td>
	    	 <input type="text" name="sortnum" maxlength="2" size="5" value="<%= sortnum %>">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��뿩�� :</td>
	    <td>
		<% if isusing = "N" then %>
			<input type="radio" name="isusing" value="Y">�����
			<input type="radio" name="isusing" value="N" checked >������
		<% else %>
			<input type="radio" name="isusing" value="Y" checked >�����
			<input type="radio" name="isusing" value="N">������
		<% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">����� :</td>
	    <td>
	        <%= regdate %>
	    </td>
	</tr>
	
</form>
</table>
<%
set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
