<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2010.01.19 한용민 수정
'	Description : 핑거스 맵등록
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%

dim map_title,map_tel,map_addr,map_etc,map_titleimg,map_img,map_html 
dim map_menuimg , map_menu_overimg ,map_subway ,map_bus
dim map_idx : map_idx = requestCheckVar(request("map_idx"),9)
dim mode    : mode = requestCheckVar(request("mode"),9)
dim sql
dim i,FResultCount
dim Fidx,FMap_title,FMpap

	sql = " select top 100 * from [db_academy].dbo.tbl_map_Info"
		if (map_idx<>"") and (mode="") then
		    sql = sql + " where map_idx="&map_idx
		end if
	sql = sql + " order by map_idx  "
	
	rsAcademyget.open sql,dbAcademyget,1

	if not rsAcademyget.eof then
		i=0
		FResultCount= rsAcademyget.recordcount
		redim Fidx(FResultCount)
		redim FMap_title(FResultCount)
		redim FMap_titleimg(FResultCount)
		redim FMap(FResultCount)
	    redim FMap_addr(FResultCount)
	    redim FMap_tel(FResultCount)
	    redim FMap_etc(FResultCount)
	    redim FMap_html(FResultCount)
	    redim fmap_menuimg(FResultCount)
	    redim fmap_menu_overimg(FResultCount)
	    redim fmap_subway(FResultCount)
	    redim fmap_bus(FResultCount)
	                
		do until rsAcademyget.eof
			Fidx(i)	= rsAcademyget("map_idx")
			FMap_title(i)	=rsAcademyget("map_title")
			FMap_titleimg(i) =rsAcademyget("map_titleimg")
			FMap(i)	=	rsAcademyget("map_img")
	        FMap_addr(i)	= db2html(rsAcademyget("map_addr"))
	        FMap_tel(i)	= rsAcademyget("map_tel")
	        FMap_etc(i)	= rsAcademyget("map_etc")
	        FMap_html(i) = db2html(rsAcademyget("map_html"))
	        fmap_menuimg(i)	= rsAcademyget("map_menuimg")
	        fmap_menu_overimg(i)	= rsAcademyget("map_menu_overimg")
	        fmap_subway(i)	= rsAcademyget("map_subway")
	        fmap_bus(i)	= rsAcademyget("map_bus")
	                                
			i=i+1
			rsAcademyget.movenext
		loop
	end if
	
	rsAcademyget.close
%>

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">

	function selmap(vv,idx,sv){
		opener.lecfrm.lec_mapimg.value=vv;
		opener.lecfrm.map_idx.value=idx;
		opener.lecfrm.lec_space.value=sv;
		self.close();
	}
	
	function confirmSubmit(frm){
	    if (frm.mode.value=="add"){
	        if (confirm('등록 하시겠습니까?')){
	            frm.submit();
	        }
	    }else{
	        if (confirm('수정 하시겠습니까?')){
	            frm.submit();
	        }
	    }
	}
	
	function delmap(map_idx){
	    if (confirm('삭제 하시겠습니까?')){
	        frm_map.mode.value="del";
	        frm_map.map_idx.value=map_idx;
	        frm_map.action='/academy/lecture/lib/pop_lec_mapimg_process.asp';
	        frm_map.submit();
	    }
	}
	
	function Showimg(img){
		var win = window.open('','ww','width=400,height=400,scrollbars=auto,resizable=yes');
		
		win.document.write ('<html><body>');
		win.document.write ('<table border="0" cellpadding="0" cellspacing="0"><tr><td>');
		win.document.write ('<img src="' + img + '" onclick="self.close();" style="cursor:hand;">');
		win.document.write ('</td></tr></table>');
		win.document.write ('</body></html>');
	}
	
</script>

<table width="100%" class="a" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<div style="overflow-y:scroll; width:100%; height:460; padding:0px">
				<table width="100%" class="a" border="0" cellpadding="1" cellspacing="1" bgcolor="#BABABA" class="a">
					<tr>
						<td align="center" bgcolor="#DDDDFF">선택</td>
						<td align="center" bgcolor="#DDDDFF">번호</td>
						<td align="center" bgcolor="#DDDDFF">설명</td>
						<td align="center" bgcolor="#DDDDFF">전화</td>
						<td align="center" bgcolor="#DDDDFF">주소</td>
						<td align="center" bgcolor="#DDDDFF">이미지</td>
						<td align="center" bgcolor="#DDDDFF">선택</td>
					</tr>
					<% For i=0 to FResultCount -1 %>
					<tr bgcolor="#FFFFFF">
						<td bgcolor="#FFFFFF"><input type="button" value="선택" class="button" onclick="javascript:selmap('<%= FMap(i) %>','<%= Fidx(i) %>','<%= FMap_title(i) %>');"></td>
						<td align=center><%= Fidx(i) %></td>
						<td ><%= FMap_title(i) %></td>
						<td ><%= FMap_addr(i) %></td>
						<td ><%= FMap_tel(i) %></td>
						<td><img src="<%= FMap(i) %>" height=50 border="0" width=50 onclick="javascript:Showimg('<%= FMap(i) %>')" style="cursor:hand;"></td>
					    <td>
					    	<input type="button" value="수정" onclick="location.href='?map_idx=<%=Fidx(i) %>'" class="button">
					    	<input type="button" value="삭제" onclick="delmap(<%= Fidx(i) %>);" class="button">
					    </td>
					</tr>
					<% next %>
				</table>
			</div>
		</td>
	</tr>
</table>
<table width="100%" class="a" border="0" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
	<form name="mapfrm" method="post" action="<%=UploadImgFingers%>/linkweb/cscenter/mapimage_proc.asp" onsubmit="confirmSubmit(this);return false;" enctype="multipart/form-data">	
	<% if (map_idx<>"") and (FResultCount>0) then %>
	<input type="hidden" name="mode" value="edit">
	<%
	map_title   = FMap_title(0)
	map_tel     = FMap_tel(0)
	map_addr    = FMap_addr(0)
	map_etc     = FMap_etc(0)
	map_titleimg= FMap_titleimg(0)
	map_img     = FMap(0)
	map_html    = FMap_html(0)
	map_menuimg     = fmap_menuimg(0)
	map_menu_overimg    = fmap_menu_overimg(0)
	map_subway     = fmap_subway(0)
	map_bus    = fmap_bus(0)	
	%>
	<tr>
		<td bgcolor="#DDDDFF">번호</td>
		<td bgcolor="#FFFFFF"><input class="text_ro" type="text" size="3" name="map_idx" value="<%= map_idx %>" readOnly >
		<input type="button" value="목록으로" onClick="location.href='pop_lec_mapimg.asp'" class="button">
		</td>
	</tr>
	<% else %>
	<input type="hidden" name="mode" value="add">
	<tr>
		<td bgcolor="#DDDDFF">번호</td>
		<td bgcolor="#FFFFFF"><input type="text" size="3" name="map_idx" value=""  ></td>
	</tr>
	<% end if %>
	<tr>
		<td bgcolor="#DDDDFF">제목</td>
		<td bgcolor="#FFFFFF"><input type="text" size="50" name="map_title" value="<%= map_title %>"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">주소</td>
		<td bgcolor="#FFFFFF"><input type="text" size="50" name="map_addr" value="<%= map_addr %>" maxlength="100"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">주소(기타)</td>
		<td bgcolor="#FFFFFF"><input type="text" size="50" name="map_etc" value="<%= map_etc %>" maxlength="100"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">전화</td>
		<td bgcolor="#FFFFFF"><input type="text" size="30" name="map_tel" value="<%= map_tel %>" maxlength="30"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">지하철</td>
		<td bgcolor="#FFFFFF"><textarea name="map_subway" cols="70" rows="2"><%= map_subway %></textarea></td>
	</tr>		
	<tr>
		<td bgcolor="#DDDDFF">버스</td>
		<td bgcolor="#FFFFFF"><textarea name="map_bus" cols="70" rows="2"><%= map_bus %></textarea></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">제목이미지</td>
		<td bgcolor="#FFFFFF">
			<% if map_titleimg <> "" then %>
			<img src="<%=map_titleimg%>"><br>
			<% end if %>
			<input type="file" name="map_titleimg" size="32" maxlength="32" class="file">		
		</td>
	</tr>	
	<tr>
		<td bgcolor="#DDDDFF">메인이미지</td>
		<td bgcolor="#FFFFFF">
			<% if map_img <> "" then %>
			<img src="<%=map_img%>"><br>
			<% end if %>
			<input type="file" name="map_img" size="32" maxlength="32" class="file">	
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">메뉴이미지</td>
		<td bgcolor="#FFFFFF">
			<% if map_menuimg <> "" then %>
			<img src="<%=map_menuimg%>"><br>
			<% end if %>
			<input type="file" name="map_menuimg" size="32" maxlength="32" class="file">				
		</td>
	</tr>		
	<tr>
		<td bgcolor="#DDDDFF">메뉴오버이미지</td>
		<td bgcolor="#FFFFFF">
			<% if map_menu_overimg <> "" then %>
			<img src="<%=map_menu_overimg%>"><br>
			<% end if %>
			<input type="file" name="map_menu_overimg" size="32" maxlength="32" class="file">			
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">기타(태그 사용가능, 최하단에 표시 됩니다)</td>
		<td bgcolor="#FFFFFF">
			<textarea name="map_html" cols="70" rows="5"><%= map_html %></textarea>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" colspan=2 align="center">
		<% if (map_idx<>"") then %>
		<input type="button" value="수정" onclick="confirmSubmit(mapfrm);" class="button">
		<% else %>
		<input type="button" value="신규저장" onclick="confirmSubmit(mapfrm);" class="button">
		<% end if %>
		</td>
	</tr>
	</form>
</table>

<form name="frm_map" method="post">
	<input type="hidden" name="mode">
	<input type="hidden" name="map_idx">
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->