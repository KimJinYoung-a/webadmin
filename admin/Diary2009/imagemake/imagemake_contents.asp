<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���̾���丮
' History : 2008.10.12 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->

<%
dim idx, poscode,reload , ix, tmp
	idx = request("idx")
	poscode = request("poscode")
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

dim oposcode, defaultMapStr
	set oposcode = new DiaryCls
	oposcode.FRectPosCode = poscode
	if poscode<>"" then
	    oposcode.fposcode_oneitem
	end if

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript'>

function SaveMainContents(frm){
    if (frm.poscode.value.length<1){
        alert('������ ���� ���� �ϼ���.');
        frm.poscode.focus();
        return;
    }

    if (frm.linkpath.value.length<1){
        alert('��ũ ���� �Է� �ϼ���.');
        frm.linkpath.focus();
        return;
    }

    if (frm.image_order.value.length<1){
        alert('�̹��� �켱������ �Է� �ϼ���.');
        frm.image_order.focus();
        return;
    }

	<% if poscode="18" or CStr(oMainContents.FOneItem.Fposcode)="18" then %>
	    if (frm.itemid.value.length<1){
	        alert('��ǰ �ڵ带 �Է��� �ּ���.');
	        frm.itemid.focus();
	        return;
	    }
	<% end if %>
    
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


$(function(){
	//�޷´�ȭâ ����
	var arrDayMin = ["��","��","ȭ","��","��","��","��"];
	var arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
	$("#event_start").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
		numberOfMonths: 2,
		showCurrentAtPos: 1,
		showOn: "button",
		<% if Idx<>"" then %>maxDate: "<%= left(oMainContents.FOneItem.fevent_end,10) %>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#event_end" ).datepicker( "option", "minDate", selectedDate );
		}
	});
	$("#event_end").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
		numberOfMonths: 2,
		showOn: "button",
		<% if Idx<>"" then %>minDate: "<%= left(oMainContents.FOneItem.fevent_start,10) %>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#event_start" ).datepicker( "option", "maxDate", selectedDate );
		}
	});
});

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
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/diary/image_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
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
	    	<%
	    		Dim vFlashTo
	    		If CStr(poscode) = "2" Then
	    			vFlashTo = 999
	    		Else
	    			If CStr(poscode) = "" AND CStr(oMainContents.FOneItem.Fposcode) = "2" Then
	    				vFlashTo = 999
	    			Else
	    				vFlashTo = 50
	    			End If
	    		End If
	    	%>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
					<select name="image_order">
						<option value="0">����</option>
						<% for ix = 1 to vFlashTo %>
							<option value="<%=ix%>" <% if cint(oMainContents.FOneItem.fimage_order) = cint(ix) then response.write " selected"%>><%= ix %></option>
						<% next %>
					</select>
	        <% else %>
	            <% if CStr(poscode) <> "" then %>
					<select name="image_order">
						<option value="0">����</option>
						<% for ix = 1 to vFlashTo %>
							<option value="<%=ix%>"><%= ix %></option>
						<% next %>
					</select>
					<% if CStr(poscode) = 3 then %>
					<script>frmcontents.image_order.options[1].selected = true;</script>
					<% end if %>
					�Ǽ��� ����� ���ڰ� Ŭ��� �켱����
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
	            <% if CStr(poscode) <> "" then %>
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
	  <br><img src="<%=uploadUrl%>/diary/main/<%= oMainContents.FOneItem.fimagepath %>" border="0" width="750px;">
	  <br><%=uploadUrl%>/diary/main/<%= oMainContents.FOneItem.fimagepath %>
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="150" align="center">�������� �ؽ�Ʈ</td>
		<td><input type="text" name="swipertext" value="<%= oMainContents.FOneItem.Fswipertext %>" size="30" maxlength="30"></td>		
	</tr>
	<!--2016 ���ι�� �¿� �÷��ڵ� �߰� ���¿� 150921-->
	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">��� ��,�� �÷��ڵ�</td>
	  <td>
	  ��:<input type="text" name="colorcodeleft" value="<%= oMainContents.FOneItem.fcolorcodeleft %>" size="5" maxlength="10">
	  ��:<input type="text" name="colorcoderight" value="<%= oMainContents.FOneItem.fcolorcoderight %>" size="5" maxlength="10">
	  ex) FF0000
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">����� �̹����� :</td>
	  <td>
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <input type="text" name="imagehcount" value="<%= oMainContents.FOneItem.fimagecount %>" size="2" maxlength="2">
	  <% else %>
	        <% if CStr(poscode) <> "" then %>
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
	        <% if CStr(poscode) <> "" then %>
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
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <input type="text" name="imageheight" value="<%= oMainContents.FOneItem.Fimageheight %>" size="8" maxlength="16">
	  <% else %>
	        <% if CStr(poscode) <> "" then %>
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
	<% If CStr(poscode) = "27" OR CStr(oMainContents.FOneItem.Fposcode) = "27" Then %>
	    ��ǰ�ڵ� :
	    ����� �����ϰ� ��ǰ�ڵ尪��.<br>
	<% Else %>
		<% If CStr(poscode) = "3" Then %>
			<b>��� �ڵ�ȿ��� ' �� ����ϸ� �ȵǰ� �ݵ�� " �� ����ؾ���.</b><br>
		<% End If %>
	<% End If %>
			<%
			'//�������
			if oMainContents.FOneItem.Fidx<>"" then
			%>
			<% if oMainContents.FOneItem.fimagetype="map" then %>
				<textarea name="linkpath" cols="60" rows="6"><%= oMainContents.FOneItem.flinkpath %></textarea>
			<% else %>
				<input type="text" name="linkpath" value="<%= oMainContents.FOneItem.flinkpath %>" maxlength="128" size="60">
			<% end if %>
	        <%
			'// �űԵ��
	        else
	            if CStr(poscode) <> "" then
			%>
	                <% if oposcode.FOneItem.fimagetype="map" then %>
						<textarea name="linkpath" cols="60" rows="6"></textarea>
	                    <br>
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
<%
'//�������
if oMainContents.FOneItem.Fidx<>"" then
%>
	<%' if oMainContents.FOneItem.Fposcode = "600"  then %>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ǰ�ڵ�</td>
	    <td>
	        <input type="text" name="itemid" value="<%= oMainContents.FOneItem.fevt_code %>" maxlength="128" size="60">
	        <% if CStr(oMainContents.FOneItem.Fposcode)="18" then %>
	        	<br><font color="red">�� ��ǥ ��ǰ�ڵ带 �����ž� ��ʰ� ���ɴϴ�.</font>
	        <% end if %>
	    </td>
	</tr>
	<%' end if %>
<%
'// �űԵ��
else
'	if CStr(poscode) <> "" and CStr(poscode) = "600" then
%>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ǰ�ڵ�</td>
	    <td>
	        <input type="text" name="itemid" value="" maxlength="128" size="60">
	    </td>
	</tr>
<%
'	end if
end if %>

<%
'//�������
if oMainContents.FOneItem.Fidx<>"" then

	'if CStr(oMainContents.FOneItem.Fposcode) = "402" or CStr(oMainContents.FOneItem.Fposcode) = "1100" then
	if CStr(oMainContents.FOneItem.Fposcode) = "16" or CStr(oMainContents.FOneItem.Fposcode) = "17" or CStr(oMainContents.FOneItem.Fposcode) = "19" or CStr(oMainContents.FOneItem.Fposcode) = "18" or CStr(oMainContents.FOneItem.Fposcode) = "20" then
%>
	<tr align="center" bgcolor="#FFFFFF">
		<td nowrap width="152">������</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="text" name="event_start"  id="event_start" size=10 value="<%= left(oMainContents.FOneItem.fevent_start,10) %>">

<!--			<a href="" onclick="calendarOpen3(frmcontents.event_start,'������',frmcontents.event_start.value); return false;">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
-->


			~<input type="text" name="event_end" id="event_end" size=10  value="<%= left(oMainContents.FOneItem.fevent_end,10) %>">
<!--
			<a href="" onclick="calendarOpen3(frmcontents.event_end,'��������',frmcontents.event_end.value); return false;">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>	
-->
		</td>
	</tr>
<%
	end if

'// �űԵ��
else

	'if CStr(poscode) <> "" and (CStr(poscode) = "402" or CStr(poscode) = "1100") then
	if CStr(poscode) ="16" or CStr(poscode) = "17" or CStr(poscode) = "19" or CStr(poscode) = "18" or CStr(poscode) = "20" then
%>
	<tr align="center" bgcolor="#FFFFFF">
		<td nowrap width="152">������</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="text" name="event_start" id="event_start" size=10 value="">

<!--
			<a href="" onclick="calendarOpen3(frmcontents.event_start,'������',frmcontents.event_start.value); return false;">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
-->
			~<input type="text" name="event_end" id="event_end" size=10  value="">

<!--
			<a href="" onclick="calendarOpen3(frmcontents.event_end,'��������',frmcontents.event_end.value); return false;">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>	
-->
		</td>
	</tr>
<%
	end if
end if %>

<%
'//�������
if idx<>0 then
%>
	<% if CStr(oMainContents.FOneItem.fposcode) = "200"  then %>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ǰ�ڵ�</td>
	    <td>
	        <input type="text" name="itemid" value="<%= oMainContents.FOneItem.fevt_code %>" maxlength="128" size="30">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ǰŸ��</td>
	    <td>
	    	<select name="itemtype">
	    		<option <% if oMainContents.FOneItem.fitemtype = "" then response.write " selected" %>>����</option>
	    		<option value="story" <% if oMainContents.FOneItem.fitemtype = "story" then response.write " selected" %>>story</option>
	    		<option value="event" <% if oMainContents.FOneItem.fitemtype = "event" then response.write " selected" %>>event</option>
	    	</select>
	    </td>
	</tr>
	<% end if %>
<%
'// �űԵ��
else
	if CStr(poscode) = "200" then
%>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ǰ�ڵ�</td>
	    <td>
	        <input type="text" name="itemid"  maxlength="128" size="30">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">��ǰŸ��</td>
	    <td>
	    	<select name="itemtype">
	    		<option>����</option>
	    		<option value="story">story</option>
	    		<option value="event">event</option>
	    	</select>
	    </td>
	</tr>
<%
	end if
end if %>

<input type="hidden" name="groupcode" value="0">
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
</form>
</table>
<%
set oposcode = Nothing
set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
