<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  해외출고관리
' History : 2009.04.07 서동석 생성
'			2022.07.22 한용민 수정(홀쎄일 카톤박스 결제 추가, 보안강화, 소스표준화)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<%
dim page, shopid, showmichulgo, tplgubun, research, i, pgsz
	menupos = requestCheckvar(getNumeric(request("menupos")),10)
	page = requestCheckvar(getNumeric(request("page")),10)
	pgsz = requestCheckvar(getNumeric(request("pgsz")),10)
	shopid = requestCheckvar(request("shopid"),32)
	showmichulgo = requestCheckvar(request("showmichulgo"),2)
	research = requestCheckvar(request("research"),2)
	tplgubun = requestCheckvar(request("tplgubun"),32)

if (page = "") then
	page = 1
end if

if (pgsz = "") then
	pgsz = 20
end if

page = CLng(page)
pgsz = CLng(pgsz)

dim occartoonbox
set occartoonbox = new CCartoonBox
occartoonbox.FRectShopid = shopid
occartoonbox.FRectShowMichulgo = showmichulgo
occartoonbox.FCurrPage = page
occartoonbox.Fpagesize = pgsz
occartoonbox.FtplGubun = tplgubun
occartoonbox.GetMasterList

dim oinnerboxlist
set oinnerboxlist = new CCartoonBox
oinnerboxlist.FRectMasterIdx = -1
oinnerboxlist.FRectShopid = "ALL"
oinnerboxlist.FtplGubun = tplgubun
oinnerboxlist.GetInnerBoxInserted   ''이쿼리가 느림. 주석처리 2016/09/06 eastone

dim shopidlist, tmpshopid
shopidlist = ""
tmpshopid = ""
for i = 0 to oinnerboxlist.FResultCount - 1
	if (tmpshopid <> oinnerboxlist.FItemList(i).Fshopid) then
		if (shopidlist = "") then
			shopidlist = oinnerboxlist.FItemList(i).Fshopid
		else
			shopidlist = shopidlist + ", " + oinnerboxlist.FItemList(i).Fshopid
		end if

		tmpshopid = oinnerboxlist.FItemList(i).Fshopid
	end if
next

%>

<script type='text/javascript'>

function popJungsanMaster(iid){
	var popwin = window.open('/admin/offshop/franmeaippopsubmaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function smssendreg(masteridx){
	var popwin = window.open('/admin/fran/jumun_smssendreg.asp?masteridx='+masteridx+'&paymentgroup=CARTOONBOX','regsmssend','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			ShopID : 
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			&nbsp;
			3PL 구분 : <% Call drawSelectBoxTPLGubun("tplgubun", tplgubun) %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			미출고만 : <input type="checkbox" class="checkbox" name="showmichulgo" value="Y" <% if (showmichulgo = "Y") then %>checked<% end if %>>
			&nbsp;
			표시갯수 :
			<select class="select" name="pgsz">
				<option value="20" <%= CHKIIF((pgsz = 20), "selected", "") %> >20</option>
				<option value="100" <%= CHKIIF((pgsz = 100), "selected", "") %> >100</option>
				<option value="500" <%= CHKIIF((pgsz = 500), "selected", "") %> >500</option>
			</select>
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<form name="cartoonboxaction" action="cartoonbox_modify.asp" style="margin:0px;">
<input type="hidden" name="mode" value="new">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="right">
		<% if (oinnerboxlist.FResultCount > 0) then %>
			<font color=red>총 <%= oinnerboxlist.FResultCount %> 개의 미지정 박스가 있습니다.(<%= shopidlist %>)</font>
		<% end if %>
		<input type="button" value="새작업등록" onclick="javascript:document.cartoonboxaction.submit();" class="button">
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="11">
			검색결과 : <b><%= occartoonbox.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= occartoonbox.FTotalpage %></b>
		</td>
	</tr>
	<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">IDX</td>
		<td>작업명</td>
		<td>샵아이디</td>
		<td>샵이름</td>
		<td width="60">상태</td>
		<td>wholesale<br>결제상태</td>
		<!--
		<td width="80">출고요청일</td>
		-->
		<td width="80">출고일</td>
		<td width="60">운송방법</td>
		<td width="80">정산IDX</td>
		<td width="60">작성자</td>
		<td width="80">등록일</td>
	</tr>
	<% if occartoonbox.FResultCount >0 then %>
	<% for i=0 to occartoonbox.FResultcount-1 %>
	<tr height="25" bgcolor="#FFFFFF">
		<td align="center"><%= occartoonbox.FItemList(i).Fidx %></td>
		<td align="center"><a href="cartoonbox_modify.asp?menupos=<%= menupos %>&idx=<%= occartoonbox.FItemList(i).Fidx %>"><%= occartoonbox.FItemList(i).Ftitle %></a></td>
		<td align="center"><a href="cartoonbox_modify.asp?menupos=<%= menupos %>&idx=<%= occartoonbox.FItemList(i).Fidx %>"><%= occartoonbox.FItemList(i).Fshopid %></a></td>
		<td align="center"><a href="cartoonbox_modify.asp?menupos=<%= menupos %>&idx=<%= occartoonbox.FItemList(i).Fidx %>"><%= occartoonbox.FItemList(i).Fshopname %></a></td>
		<td align="center">
			<font color="<%= occartoonbox.FItemList(i).GetStateColor %>"><%= occartoonbox.FItemList(i).GetStateName %></font>
		</td>
		<td align="center">
			<%= occartoonbox.FItemList(i).getcartoonboxpaymentstatus %>
			<% if occartoonbox.FItemList(i).fsmssenddate<>"" and not(isnull(occartoonbox.FItemList(i).fsmssenddate)) then %>
				<br>문자발송:
				<br><%= left(occartoonbox.FItemList(i).fsmssenddate,10) %>
				<br><%= mid(occartoonbox.FItemList(i).fsmssenddate,12,22) %>
			<% else %>
				<br><input type="button" onclick="smssendreg('<%= occartoonbox.FItemList(i).Fidx %>')" value="문자발송" class="button">
			<% end if %>
		</td>
		<!--
		<td align="center"><%= occartoonbox.FItemList(i).Frequestdt %></td>
		-->
		<td align="center"><%= occartoonbox.FItemList(i).Fdeliverdt %></td>
		<td align="center"><%= occartoonbox.FItemList(i).GetDeliverMethodName %></td>
		<td align="center"><a href="javascript:popJungsanMaster(<%= occartoonbox.FItemList(i).Fjungsanidx %>)"><%= occartoonbox.FItemList(i).Fjungsanidx %></a></td>
		<td align="center"><%= occartoonbox.FItemList(i).Freguserid %></td>
		<td align="center"><%= Left(occartoonbox.FItemList(i).Fregdate, 10) %></td>
	</tr>
	<% next %>
	<% else %>
<tr bgcolor="#FFFFFF">
		<td colspan="10" align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="10" align="center">
			<%
			dim strparam
			strparam = "&shopid=" + CStr(shopid)

			strparam = strparam + "&menupos=" + CStr(menupos)
			strparam = strparam + "&showmichulgo=" + CStr(showmichulgo)

			%>
			<% if occartoonbox.HasPreScroll then %>
				<a href="?page=<%= occartoonbox.StartScrollPage-1 %>&research=on<%= strparam %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + occartoonbox.StartScrollPage to occartoonbox.FScrollCount + occartoonbox.StartScrollPage - 1 %>
				<% if i>occartoonbox.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if occartoonbox.HasNextScroll then %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set occartoonbox = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
