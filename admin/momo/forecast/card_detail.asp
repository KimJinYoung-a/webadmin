<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성예보
' Hieditor : 2010.11.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oforecast,i , oforecast_detail ,page , idx , image_url , forecastgubun, link_url, couponidx
dim cardidx ,startdate ,enddate ,isusing ,regdate , contents
	cardidx = requestcheckvar(request("cardidx"),10)
	idx = requestcheckvar(request("idx"),10)
	page = request("page")
	if page = "" then page = 1
			
if cardidx = "" then
	response.write "<script language='javascript'>alert('해당 번호가 없습니다');self.close();</script>"
end if
	
'//상세
set oforecast_detail = new cforecast_list
	oforecast_detail.frectidx = idx
	
	'//수정일경우에만 쿼리
	if idx <> "" then
		oforecast_detail.fcarddetail_oneitem()
	end if
	
	if oforecast_detail.ftotalcount > 0 then
		idx = oforecast_detail.FOneItem.fidx
		cardidx = oforecast_detail.FOneItem.fcardidx
		forecastgubun = oforecast_detail.FOneItem.fforecastgubun
		image_url = oforecast_detail.FOneItem.fimage_url
		contents = oforecast_detail.FOneItem.fcontents
		isusing = oforecast_detail.FOneItem.fisusing
		link_url = oforecast_detail.FOneItem.flink_url
		couponidx = oforecast_detail.FOneItem.fcouponidx
	end if

'// 리스트
set oforecast = new cforecast_list
	oforecast.FPageSize = 20
	oforecast.FCurrPage = page
	oforecast.frectcardidx = cardidx
	oforecast.fcard_detaillist()	
%>

<script language="javascript">

document.domain = "10x10.co.kr";

function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

	window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

//저장
function reg(){
	if (frm.forecastgubun.value==''){
	alert('카드구분을 입력해주세요');
	frm.forecastgubun.focus();
	return;
	}
	if (frm.contents.value==''){
	alert('상세설명을 입력해주세요');
	frm.contents.focus();
	return;
	}				
	if (frm.isusing.value==''){
	alert('사용여부를 선택해주세요');
	return;
	}
	
	frm.action='/admin/momo/forecast/card_process.asp';
	frm.mode.value='detailadd';
	frm.submit();
}

function cardedit(idx){
	frm.idx.value=idx;
	frm.submit();
}
	
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= cardidx %><input type="hidden" name="cardidx" value="<%= cardidx %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>카드번호</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= idx %><input type="hidden" name="idx" value="<%= idx %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>카드구분</td>
	<td bgcolor="#FFFFFF" align="left">
		<% if idx = "" then %>
			<% drawforecastgubun "forecastgubun", forecastgubun , ""%>
		<% else %>
			<%= getforecastgubun(forecastgubun) %><input type="hidden" name="forecastgubun" value="<%=forecastgubun%>">
		<% end if %>			
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상세설명</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="contents" style="width:450px; height:100px;"><%=contents%></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>이미지<br>235x309</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('image_urldiv','image_url','image','2000','235','true');"/>		
		<input type="hidden" name="image_url" value="<%= image_url %>">
		<div align="right" id="image_urldiv"><% IF image_url<>"" THEN %><img src="<%=webImgUrl%>/momo/forecast/image/<%= image_url %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>URL</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="link_url" value="<%=link_url%>" size="60">	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>보너스쿠폰 ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="couponidx" value="<%=couponidx%>" size="5">	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2><input type="button" onclick="reg();" value="저장" class="button"></td>
</tr>
</form>
</table>

<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="makeThumbYn" value="">
</form>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">			
		</td>
		<td align="right">		
			<input type="button" onclick="cardedit('');" value="신규등록" class="button">					
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oforecast.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oforecast.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oforecast.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">	
	<td align="center">번호</td>
	<td align="center">카드번호</td>
	<td align="center">구분</td>	
	<td align="center">사용여부</td>
	<td align="center">비고</td>
</tr>
<% for i=0 to oforecast.FresultCount-1 %>

<% if oforecast.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<%= oforecast.FItemList(i).fcardidx %>
	</td>
	<td align="center">
		<%= oforecast.FItemList(i).fidx %>
	</td>				
	<td align="center">
		<%= getforecastgubun(oforecast.FItemList(i).fforecastgubun) %>
	</td>		
	<td align="center">
		<%= oforecast.FItemList(i).fisusing %>
	</td>
		
	<td align="center">
		<input type="button" onclick="cardedit(<%= oforecast.FItemList(i).fidx %>);" class="button" value="수정">			
	</td>			
</tr>   

<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if oforecast.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= oforecast.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oforecast.StartScrollPage to oforecast.StartScrollPage + oforecast.FScrollCount - 1 %>
			<% if (i > oforecast.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oforecast.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&isusing=<%=isusing%>>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oforecast.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oforecast_detail = nothing
	set oforecast = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->