<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ڳʰ���
' History : 2009.09.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/corner/corner_cls.asp"-->

<%
dim idx, poscode,reload , ix
	idx = RequestCheckvar(request("idx"),10)
	poscode = RequestCheckvar(request("poscode"),10)
	reload = RequestCheckvar(request("reload"),10)
	if idx="" then idx=0

if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
end if

dim oMainContents
	set oMainContents = new cposcode_list
	oMainContents.FRectIdx = idx
	oMainContents.fcontents_oneitem

dim oposcode, defaultMapStr
	set oposcode = new cposcode_list
	oposcode.FRectPosCode = poscode	
	if poscode<>"" then
	    oposcode.fposcode_oneitem
	end if

%>
<script language='javascript'>

function SaveMainContents(frm){
    if (frm.poscode.value.length<1){
        alert('������ ���� ���� �ϼ���.');
        frm.poscode.focus();
        return;
    }
    
    if (frm.image_order.value.length<1){
        alert('�̹��� �켱������ �Է� �ϼ���.');
        frm.image_order.focus();
        return;
    }
            
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function ChangeLinktype(comp){
    if (comp.value=="M"){
       document.all.link_M.style.display = "";
       document.all.link_L.style.display = "none";
    }else{
       document.all.link_M.style.display = "none";
       document.all.link_L.style.display = "";
    }
}

function ChangeGubun(comp){
    location.href = "?poscode=" + comp.value;
    // nothing;
}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="center">
			
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="<%=imgFingers%>/linkweb/corner/image_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="ckUserId" value="<%=session("ssBctId")%>">
<!--<input type="hidden" name="ckUserId" value="<%=request.Cookies("partner")("userid")%>">-->
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">Idx :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	        <%= oMainContents.FOneItem.Fidx %>
	        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
	        <% else %>
	
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">���и� :</td>
	    <td>
	    	
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
				<%= oMainContents.FOneItem.Fposname %> (<%= oMainContents.FOneItem.Fposcode %>)
				<input type="hidden" name="poscode" value="<%= oMainContents.FOneItem.Fposcode %>">
	        <% else %>
        <% call DrawMainPosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'") %>
		        	
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">�̹������Ŀ켱���� :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
					<select name="image_order">
						<option>����</option>
						<% for ix = 1 to 50 %>
							<option value="<%=ix%>" <% if cint(oMainContents.FOneItem.fimage_order) = cint(ix) then response.write " selected"%>><%= ix %></option>				
						<% next %>						
					</select>
	        <% else %>
	            <% if poscode<>"" then %>
					<select name="image_order">
						<option>����</option>
						<% for ix = 1 to 50 %>
							<option value="<%=ix%>"><%= ix %></option>				
						<% next %>						
					</select>
					�Ǽ��� ����� ���ڰ� ������� �켱����
	            <% else %>
	            <font color="red">������ ���� �����ϼ���</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>	
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ũ���� :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	        <%= oMainContents.FOneItem.fimagetype %>
	        <% else %>
	            <% if poscode<>"" then %>
	            <%= oposcode.FOneItem.fimagetype %>
	            <% else %>
	            <font color="red">������ ���� �����ϼ���</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">�̹��� :</td>
	  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <br><img src="<%=imgFingers%>/corner/main/<%= oMainContents.FOneItem.fimagepath %>" border="0"> 
	  <br><%=imgFingers%>/culturestation/main/<%= oMainContents.FOneItem.fimagepath %>
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">����� �̹����� :</td>
	  <td>
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <input type="text" name="imagehcount" value="<%= oMainContents.FOneItem.fimagecount %>" size="2" maxlength="2"> 
	  <% else %>
	        <% if poscode<>"" then %>
	        <%= oposcode.FOneItem.fimagecount %>
	        <% else %>
	        <font color="red">������ ���� �����ϼ���</font>
	        <% end if %>
	  <% end if %>
	  </td>
	</tr>	
	<tr bgcolor="#FFFFFF">
	  <td width="150"  align="center">�̹���Width :</td>
	  <td>
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <input type="text" name="imagewidth" value="<%= oMainContents.FOneItem.Fimagewidth %>" size="8" maxlength="16"> 
	  <% else %>
	        <% if poscode<>"" then %>
	        <%= oposcode.FOneItem.Fimagewidth %>
	        <% else %>
	        <font color="red">������ ���� �����ϼ���</font>
	        <% end if %>
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">�̹���Height :</td>
	  <td>
	  <% if oMainContents.FOneItem.Fidx<>""  then %>
	  <input type="text" name="imageheight" value="<%= oMainContents.FOneItem.Fimageheight %>" size="8" maxlength="16"> 
	  <% else %>
	        <% if poscode<>"" then %>
	        <%= oposcode.FOneItem.Fimageheight %>
	        <% else %>
	        <font color="red">������ ���� �����ϼ���</font>
	        <% end if %>
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ũ�� :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	            <% if oMainContents.FOneItem.fimagetype="map" then %>
	            <textarea name="linkpath" cols="60" rows="6"><%= oMainContents.FOneItem.flinkpath %></textarea>
	            <% else %>
	            <input type="text" name="linkpath" value="<%= oMainContents.FOneItem.flinkpath %>" maxlength="128" size="60">
	            <% end if %>
	        <% else %>
	            <% if poscode<>"" then %>
	                <% if oposcode.FOneItem.fimagetype="map" then %>
	                    <textarea name="linkpath" cols="60" rows="6"><map name='Map<%=poscode%>'></map></textarea>
	                    <br>(�̹����� ������ ���� ����)
	                <% else %>
	                    <input type="text" name="linkpath" value="" maxlength="128" size="60">
	                    <br>(����η� ǥ���� �ּ���  ex: /culturestation/culturestation_event.asp?evt_code=7)
	                <% end if %>
	            <% else %>
	            <font color="red">������ ���� �����ϼ���</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">����� :</td>
	    <td>
	        <%= oMainContents.FOneItem.Fregdate %>
	    </td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��뿩�� :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fisusing="N" then %>
	        <input type="radio" name="isusing" value="Y">�����
	        <input type="radio" name="isusing" value="N" checked >������
	        <% else %>
	        <input type="radio" name="isusing" value="Y" checked >�����
	        <input type="radio" name="isusing" value="N">������
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td  align="center" colspan=2>
	    	<input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);" class="button">
	    </td>
	</tr>	
</form>
</table>
<%
set oposcode = Nothing
set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->