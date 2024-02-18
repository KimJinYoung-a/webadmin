<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.23 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<%
dim idx, poscode,reload , ix
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	if idx="" then idx=0

if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
end if

dim oMainContents
	set oMainContents = new organizerCls
	oMainContents.FRectIdx = idx
	oMainContents.fcontents_oneitem

dim oposcode, defaultMapStr
	set oposcode = new organizerCls
	oposcode.FRectPosCode = poscode	
	if poscode<>"" then
	    oposcode.fposcode_oneitem
	end if

%>
<script language='javascript'>

function SaveMainContents(frm){
    if (frm.poscode.value.length<1){
        alert('구분을 먼저 선택 하세요.');
        frm.poscode.focus();
        return;
    }
    
    if (frm.linkpath.value.length<1){
        alert('링크 값을 입력 하세요.');
        frm.linkpath.focus();
        return;
    }

    if (frm.image_order.value.length<1){
        alert('이미지 우선순위를 입력 하세요.');
        frm.image_order.focus();
        return;
    }
            
    if (confirm('저장 하시겠습니까?')){
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

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);" class="button">
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="<%=staticimgurl%>/linkweb/organizer/organizer_image_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
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
	    <td width="150" align="center">구분명 :</td>
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
	    <td width="150" align="center">이미지정렬우선순위 :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
					<select name="image_order">
						<option>선택</option>
						<% for ix = 1 to 50 %>
							<option value="<%=ix%>" <% if cint(oMainContents.FOneItem.fimage_order) = cint(ix) then response.write " selected"%>><%= ix %></option>				
						<% next %>						
					</select>
	        <% else %>
	            <% if poscode<>"" then %>
					<select name="image_order">
						<option>선택</option>
						<% for ix = 1 to 50 %>
							<option value="<%=ix%>"><%= ix %></option>				
						<% next %>						
					</select>
					실서버 적용시 숫자가 클경우 우선노출
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>	
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">링크구분 :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	        <%= oMainContents.FOneItem.fimagetype %>
	        <% else %>
	            <% if poscode<>"" then %>
	            <%= oposcode.FOneItem.fimagetype %>
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">이미지 :</td>
	  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <br><img src="<%=uploadUrl%>/organizer/main/<%= oMainContents.FOneItem.fimagepath %>" border="0"> 
	  <br><%=uploadUrl%>/organizer/main/<%= oMainContents.FOneItem.fimagepath %>
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">사용할 이미지수 :</td>
	  <td>
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <input type="text" name="imagehcount" value="<%= oMainContents.FOneItem.fimagecount %>" size="2" maxlength="2"> 
	  <% else %>
	        <% if poscode<>"" then %>
	        <%= oposcode.FOneItem.fimagecount %>
	        <% else %>
	        <font color="red">구분을 먼저 선택하세요</font>
	        <% end if %>
	  <% end if %>
	  </td>
	</tr>	
	<tr bgcolor="#FFFFFF">
	  <td width="150"  align="center">이미지Width :</td>
	  <td>
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <input type="text" name="imagewidth" value="<%= oMainContents.FOneItem.Fimagewidth %>" size="8" maxlength="16"> 
	  <% else %>
	        <% if poscode<>"" then %>
	        <%= oposcode.FOneItem.Fimagewidth %>
	        <% else %>
	        <font color="red">구분을 먼저 선택하세요</font>
	        <% end if %>
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	  <td width="150" align="center">이미지Height :</td>
	  <td>
	  <% if oMainContents.FOneItem.Fidx<>"" then %>
	  <input type="text" name="imageheight" value="<%= oMainContents.FOneItem.Fimageheight %>" size="8" maxlength="16"> 
	  <% else %>
	        <% if poscode<>"" then %>
	        <%= oposcode.FOneItem.Fimageheight %>
	        <% else %>
	        <font color="red">구분을 먼저 선택하세요</font>
	        <% end if %>
	  <% end if %>
	  </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<% If poscode = 99 OR oMainContents.FOneItem.Fposcode = 99 Then %>
	    <td width="150" align="center">상품코드 :</td>
	    <td>상대경로 무시하고 상품코드값만.<br>
	<% Else %>
	    <td width="150" align="center">링크값 :</td>
	    <td> ※이미지맵 경로 절대 바꾸시면 안됩니다.<br>
	<% End If %>
	        <% 
	        '//수정모드
	        if oMainContents.FOneItem.Fidx<>"" then 
	        %>
	            <% if oMainContents.FOneItem.fimagetype="map" then %>
	            <textarea name="linkpath" cols="60" rows="6"><%= oMainContents.FOneItem.flinkpath %></textarea>
	            <% else %>
	            <input type="text" name="linkpath" value="<%= oMainContents.FOneItem.flinkpath %>" maxlength="128" size="60">
	            <% end if %>
	        <% 
			'// 신규등록	        	
	        else 
	            if poscode<>"" then 
			%>
	                <% if oposcode.FOneItem.fimagetype="map" then %>
						<textarea name="linkpath" cols="60" rows="6"><map name='Map<%=poscode%>'></map></textarea>
	                    <br>(이미지맵 변수값 변경 금지)
	                <% else %>
	                    <input type="text" name="linkpath" value="" maxlength="128" size="60">
	                    <br>(상대경로로 표시해 주세요  ex: /culturestation/culturestation_event.asp?evt_code=7)
	                <% end if %>
	            <% else %>
	            <font color="red">구분을 먼저 선택하세요</font>
	            <% end if %>
	        <% end if %>
	    </td>
	</tr>
<% 
'//수정모드
if idx<>0 then 
%>	
	<% if oMainContents.FOneItem.fposcode="500"  then %>  
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">상품코드</td>
	    <td>
	        <input type="text" name="itemid" value="<%= oMainContents.FOneItem.fevt_code %>" maxlength="128" size="30">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">상품타입</td>
	    <td>
	    	<select name="itemtype">
	    		<option <% if oMainContents.FOneItem.fitemtype = "" then response.write " selected" %>>선택</option>
	    		<option value="story" <% if oMainContents.FOneItem.fitemtype = "story" then response.write " selected" %>>story</option>
	    		<option value="event" <% if oMainContents.FOneItem.fitemtype = "event" then response.write " selected" %>>event</option>	    			
	    	</select>
	    </td>
	</tr>
	<% end if %>
<% 
'// 신규등록	        	
else
	if poscode = "500" then
%>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">상품코드</td>
	    <td>
	        <input type="text" name="itemid"  maxlength="128" size="30">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">상품타입</td>
	    <td>
	    	<select name="itemtype">
	    		<option>선택</option>
	    		<option value="story">story</option>
	    		<option value="event">event</option>	    			
	    	</select>
	    </td>
	</tr>
<% 
	end if 
end if %>	

	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">등록일 :</td>
	    <td>
	        <%= oMainContents.FOneItem.Fregdate %>
	    </td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">사용여부 :</td>
	    <td>
	        <% if oMainContents.FOneItem.Fisusing="N" then %>
	        <input type="radio" name="isusing" value="Y">사용함
	        <input type="radio" name="isusing" value="N" checked >사용안함
	        <% else %>
	        <input type="radio" name="isusing" value="Y" checked >사용함
	        <input type="radio" name="isusing" value="N">사용안함
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
