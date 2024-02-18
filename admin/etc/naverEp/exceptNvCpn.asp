<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<%
Dim sqlStr
Dim research : research = requestCheckvar(request("research"), 10)
Dim makerid	 : makerid	= requestCheckvar(request("makerid"), 32)
Dim page	 : page	= requestCheckvar(request("page"), 10)

Dim itemidarr : itemidarr	= request("itemidarr")
Dim exceptTp : exceptTp	= requestCheckvar(request("exceptTp"), 10)
''dim NOTmakerid, orderby, isusing

Dim nMaker

if (exceptTp="") then exceptTp="B"
If page = "" Then page = 1

Dim tmpArr
if itemidarr<>"" then
	itemidarr = Replace(itemidarr," ",",")
	itemidarr = Replace(itemidarr,vbCrLf,",")
	tmpArr = split(itemidarr,",")
	itemidarr = ""
	for i=0 to uBound(tmpArr)
		if isNumeric(tmpArr(i)) then
			itemidarr = itemidarr & chkIIF(itemidarr<>"",",","") & trim(tmpArr(i))
		end if
	next
end if

SET nMaker = new epShop
	nMaker.FCurrPage				= page
	nMaker.FPageSize				= 20
	nMaker.FMakerId					= makerid
	nMaker.FRectItemid				= itemidarr
	if (exceptTp="I") then
		nMaker.getNaverCpnExceptItemList
	else
    	nMaker.getNaverCpnExceptBrandList
	end if
%>
<script language='javascript'>
function goPage(pg){
    var frm = document.frmsearch;
    frm.page.value=pg;
	frm.submit();
}

function addDisableCpn(){
	var iURI = "pop_exceptNvCpnAdd.asp?exceptTp=<%=exceptTp%>";
	var popwin = window.open(iURI,'pop_exceptNvCpnAdd','scrollbars=yes,resizable=yes,width=600,height=400');
	popwin.focus();
}


function pop_disableCpn(iexcepttp,imakeridorItemid){
	if (iexcepttp=="B") {
		var iURI = "pop_exceptNvCpnAdd.asp?exceptTp="+iexcepttp+"&mode=R&makerid="+imakeridorItemid;
	}else{
		var iURI = "pop_exceptNvCpnAdd.asp?exceptTp="+iexcepttp+"&mode=R&itemid="+imakeridorItemid;
	}

	var popwin = window.open(iURI,'pop_exceptNvCpnAdd','scrollbars=yes,resizable=yes,width=600,height=400');
	popwin.focus();
}

/*
function disableCpn(iexcepttp,imakeridorItemid){
	var frm = document.frmAct;
	var configmMsg = imakeridorItemid + ' 브랜드를 네이버쿠폰 적용 불가능으로 변경하시겠습니까?';
	if (iexcepttp=="I") configmMsg = imakeridorItemid + ' 상품을 네이버쿠폰 적용 불가능으로 변경하시겠습니까?';
	if (confirm(configmMsg)){
		if (iexcepttp=="I"){
			frm.itemid.value = imakeridorItemid;
		}else{
			frm.makerid.value = imakeridorItemid;
		}
		
		frm.excepttp.value =iexcepttp;
		frm.mode.value = "R"
		frm.submit();
	}
}
*/

function enableCpn(iexcepttp,imakeridorItemid){
	var frm = document.frmAct;
	var configmMsg = imakeridorItemid + ' 브랜드를 네이버쿠폰 적용 가능으로 변경하시겠습니까?';
	if (iexcepttp=="I") configmMsg = imakeridorItemid + ' 상품을 네이버쿠폰 적용 가능으로 변경하시겠습니까?';
	if (confirm(configmMsg)){
		var popwin = window.open("",'iblankpop','scrollbars=yes,resizable=yes,width=400,height=400');

		if (iexcepttp=="I"){
			frm.itemid.value = imakeridorItemid;
		}else{
			frm.makerid.value = imakeridorItemid;
		}
		frm.excepttp.value =iexcepttp;
		frm.mode.value = "D"
		frm.target="iblankpop";
		frm.submit();
	}
}


function showNvCpnList(){
	var iURI = "/admin/shopmaster/itemcouponitemlisteidtMulti.asp?onlynv=Y&exceptnvcpn=Y"
	var popwin = window.open(iURI,'itemcouponitemlisteidtMulti2','scrollbars=yes,resizable=yes,width=1200,height=800');
	popwin.focus();
}

</script>
<!-- 검색 시작 -->
<form name="frmsearch" method="get"  >
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a" >
		<tr>
		    <td width="90%">Naver 전용 쿠폰 
			<input type="radio" name="exceptTp" value="B" <%= CHKIIF(exceptTp="B","checked","")%> >제외 브랜드
			<input type="radio" name="exceptTp" value="I" <%= CHKIIF(exceptTp="I","checked","")%> >제외 상품
			</td>
			<td rowspan="2" valign="middle">
				상품코드:<br><textarea name="itemidarr" style="width:300px; height:50px;"><%= itemidarr %></textarea>
			</td>
		    <td rowspan="2" width="10%" align="center"><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
				브랜드ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >&nbsp;
				
				<% if (FALSE) then %>
                판매여부 : 
				<select name="isusing" class="select">
					<option value="">-Choice-</option>
					<option value="Y" <%= Chkiif(isusing = "Y", "selected", "") %> >판매</option>
					<option value="N" <%= Chkiif(isusing = "N", "selected", "") %> >판매안함</option>
				</select>
				&nbsp;
				정렬기준 : 
				<select name="orderby" class="select">
					<option value="">-Choice-</option>
					<option value="lastupdate" <%= Chkiif(orderby = "lastupdate", "selected", "") %> >최종수정일</option>
					<option value="best" <%= Chkiif(orderby = "best", "selected", "") %> >베스트브랜드</option>
				</select>
                <% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td>
		<input type="button" value="쿠폰삭제할 상품목록보기" onClick="showNvCpnList()">
	</td>
	<td align="right">
		<input type="button" value="쿠폰 적용 불가 신규등록" onClick="addDisableCpn()">
	</td>
</tr>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<% if (exceptTp="I") then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>브랜드ID</td>
		<td>상품코드</td>
		<td width="50">이미지</td>
		<td>상품명</td>
		<td>등록일</td>
		<td>등록자</td>
		<td>제외만료일</td>
		<td>Action</td>
	</tr>
	<% If nMaker.FResultCount > 0 Then %>
	<% For i = 0 To nMaker.FResultCount - 1 %>
	<tr bgcolor="<%=CHKIIF(nMaker.FItemList(i).FisExpired,"#DDDDDD","#FFFFFF")%>" height="30" align="center" height="25">
		<td><%=nMaker.FItemList(i).FMakerid%></td>
		<td><%=nMaker.FItemList(i).Fitemid%></td>
		<td><img src="<%=nMaker.FItemList(i).FsmallImage%>" width="50"></td>
		<td><%=nMaker.FItemList(i).Fitemname%></td>
		<td><%=nMaker.FItemList(i).FRegdate%></td>
		<td><%=nMaker.FItemList(i).FRegid%></td>
		<td>
			<%=CHKIIF(isNULL(nMaker.FItemList(i).FAsignMaxDt),"",nMaker.FItemList(i).FAsignMaxDt)%></td>
		<td>
			<% if (nMaker.FItemList(i).FisExpired) then %>
				<input type="button" value="쿠폰 적용 불가로 변경" onClick="pop_disableCpn('<%=exceptTp%>','<%=nMaker.FItemList(i).Fitemid%>')">
			<% else %>
				<input type="button" value="쿠폰 적용 가능 으로 변경" onClick="enableCpn('<%=exceptTp%>','<%=nMaker.FItemList(i).Fitemid%>')">
			<% end if %>
		</td>
	</tr>
	<% Next %>
	<tr height="30">
		<td colspan="8" align="center" bgcolor="#FFFFFF">
		<% If nMaker.HasPreScroll Then %>
			<a href="javascript:goPage('<%= nMaker.StartScrollPage-1 %>');">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i=0 + nMaker.StartScrollPage To nMaker.FScrollCount + nMaker.StartScrollPage - 1 %>
			<% If i>nMaker.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If nMaker.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>');">[next]</a>
		<% Else %>
		[next]
		<% End If %>
		</td>
	</tr>
	<% Else %>
	<tr height="50">
		<td colspan="16" align="center" bgcolor="#FFFFFF">
			등록된 상품이 없습니다
		</td>
	</tr>
	<% End If %>
<% else %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>브랜드ID</td>
		<td>브랜드명</td>
		<td>등록일</td>
		<td>등록자</td>
		<td>제외만료일</td>
		<td>Action</td>
	</tr>
	<% If nMaker.FResultCount > 0 Then %>
	<% For i = 0 To nMaker.FResultCount - 1 %>
	<tr bgcolor="<%=CHKIIF(nMaker.FItemList(i).FisExpired,"#DDDDDD","#FFFFFF")%>" height="30" align="center" height="25">
		<td><%=nMaker.FItemList(i).FMakerid%></td>
		<td><%=nMaker.FItemList(i).FSocName_kor%></td>
		<td><%=nMaker.FItemList(i).FRegdate%></td>
		<td><%=nMaker.FItemList(i).FRegid%></td>
		<td>
			<%=CHKIIF(isNULL(nMaker.FItemList(i).FAsignMaxDt),"",nMaker.FItemList(i).FAsignMaxDt)%></td>
		<td>
			<% if (nMaker.FItemList(i).FisExpired) then %>
				<input type="button" value="쿠폰 적용 불가로 변경" onClick="pop_disableCpn('<%=exceptTp%>','<%=nMaker.FItemList(i).FMakerid%>')">
			<% else %>
				<input type="button" value="쿠폰 적용 가능 으로 변경" onClick="enableCpn('<%=exceptTp%>','<%=nMaker.FItemList(i).FMakerid%>')">
			<% end if %>
		</td>
	</tr>
	<% Next %>
	<tr height="30">
		<td colspan="6" align="center" bgcolor="#FFFFFF">
		<% If nMaker.HasPreScroll Then %>
			<a href="javascript:goPage('<%= nMaker.StartScrollPage-1 %>');">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i=0 + nMaker.StartScrollPage To nMaker.FScrollCount + nMaker.StartScrollPage - 1 %>
			<% If i>nMaker.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If nMaker.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>');">[next]</a>
		<% Else %>
		[next]
		<% End If %>
		</td>
	</tr>
	<% Else %>
	<tr height="50">
		<td colspan="16" align="center" bgcolor="#FFFFFF">
			등록된 브랜드가 없습니다
		</td>
	</tr>
	<% End If %>
<% end if %>
</form>
</table>
<% SET nMaker = nothing %>
<form name="frmAct" method="post" action="exceptNvCpn_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="excepttp" value="">
<input type="hidden" name="makerid" value="">
<input type="hidden" name="itemid" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->