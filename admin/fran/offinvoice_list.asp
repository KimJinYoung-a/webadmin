<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 오프라인주문서관리
' History : 2010.06.03 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<%

menupos = request("menupos")



dim page, shopid
dim research, i

page = request("page")
shopid = request("shopid")
research = request("research")

if (page = "") then
	page = 1
end if



'================================================================================
dim ocoffinvoice

set ocoffinvoice = new COffInvoice

ocoffinvoice.FRectShopid = shopid

ocoffinvoice.FCurrPage = page
ocoffinvoice.Fpagesize = 25

ocoffinvoice.GetMasterList

%>

<script language='javascript'>

function PopDownloadExportDeclareFile(masteridx,ino) {
	var popwin;

	popwin = window.open('<%= uploadImgUrl %>/linkweb/offinvoice/offinvoice_download.asp?idx=' + masteridx+'&ino='+ino,'PopDownloadExportDeclareFile','width=100,height=100,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popJungsanMaster(iid){
	var popwin = window.open('/admin/offshop/franmeaippopsubmaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			ShopID : 
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="offinvoiceaction" action="offinvoice_modify.asp">
<form name="mode" value="new">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr>
	<td align="right">
		<!--
		<input type="button" value="새인보이스등록" onclick="javascript:document.offinvoiceaction.submit();" class="button">
		-->
		* 인보이스 작성은 가맹점정산관리(매출) 에서 할 수 있습니다.
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18">
			검색결과 : <b><%= ocoffinvoice.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ocoffinvoice.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">IDX</td>
		<td>샵아이디</td>
		<td>정산IDX</td>
		<td>작업IDX</td>
		<td>인보이스<br>NO</td>
		<td>운송<br>방법</td>
		<td>운임<br>부담</td>
		<td>정산<br>시기</td>
		<td>박스<br>수량</td>
		<td>총상품금액<br>(원)</td>
		<td>총운임<br>(원)</td>
		<td>작성화폐</td>
		<td>수출환율</td>
		<td>총상품금액<br>(외환)</td>
		<td>총운임<br>(외환)</td>
		<td width="80">등록일</td>
		<td>상태</td>
		<td>비고</td>
	</tr>
	<% if ocoffinvoice.FResultCount >0 then %>
	<% for i=0 to ocoffinvoice.FResultcount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= ocoffinvoice.FItemList(i).Fidx %></td>
		<td align="center"><a href="offinvoice_modify.asp?menupos=<%= menupos %>&idx=<%= ocoffinvoice.FItemList(i).Fidx %>"><%= ocoffinvoice.FItemList(i).Fshopid %><br><%= ocoffinvoice.FItemList(i).Fshopname %></a></td>
		<td align="center">
			<a href="javascript:popJungsanMaster(<%= ocoffinvoice.FItemList(i).Fjungsanidx %>)"><%= ocoffinvoice.FItemList(i).Fjungsanidx %></a>
		</td>
		<td align="center">
			<a href="javascript:PopExportSheet(<%= ocoffinvoice.FItemList(i).Fworkidx %>)"><%= ocoffinvoice.FItemList(i).Fworkidx %></a>
		</td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Finvoiceno %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetDeliverMethodName %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetExportMethodName %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetJungsanTypeName %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Ftotalboxno %></td>
		<td align="right">
			<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalgoodsprice, 0) %>&nbsp;
		</td>
		<td align="right">
			<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalboxprice, 0) %>&nbsp;
		</td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Fpriceunit %></td>
		<td align="center"><%= FormatNumber(ocoffinvoice.FItemList(i).Fexchangerate, 0) %> 원</td>
		<td align="right">
			<% if (ocoffinvoice.FItemList(i).Fexchangerate <> "") and (Not IsNull(ocoffinvoice.FItemList(i).Fexchangerate)) and (ocoffinvoice.FItemList(i).Fexchangerate <> "0") then %>
				<% if (ocoffinvoice.FItemList(i).Fpriceunit = "JPY") then %>
					<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalgoodsprice/(ocoffinvoice.FItemList(i).Fexchangerate/100)), 0), 0) %>&nbsp;
				<% else %>
					<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalgoodsprice/ocoffinvoice.FItemList(i).Fexchangerate), 2), 2) %>&nbsp;
				<% end if %>
			<% end if %>
		</td>
		<td align="right">
			<% if (ocoffinvoice.FItemList(i).Fexchangerate <> "") and (Not IsNull(ocoffinvoice.FItemList(i).Fexchangerate)) and (ocoffinvoice.FItemList(i).Fexchangerate <> "0") then %>
				<% if (ocoffinvoice.FItemList(i).Fpriceunit = "JPY") then %>
					<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalboxprice/(ocoffinvoice.FItemList(i).Fexchangerate/100)), 0), 0) %>&nbsp;
				<% else %>
					<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalboxprice/ocoffinvoice.FItemList(i).Fexchangerate), 2), 2) %>&nbsp;
				<% end if %>
			<% end if %>
		</td>
		<td align="center"><%= Left(ocoffinvoice.FItemList(i).Fregdate, 10) %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetStateCDName %></td>
		<td align="center">
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename <> "") then %>
			<input type="button" class="button" value="수출신고필증1" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,1)">
			<% end if %>
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename2 <> "") then %>
			<input type="button" class="button" value="수출신고필증2" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,2)">
			<% end if %>
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename3 <> "") then %>
			<input type="button" class="button" value="수출신고필증3" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,3)">
			<% end if %>
		</td>
	</tr>
	<% next %>
	<% else %>
<tr bgcolor="#FFFFFF">
		<td colspan=18 align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18" align="center">
			<%
			dim strparam
			strparam = "&shopid=" + CStr(shopid)

			strparam = strparam + "&menupos=" + CStr(menupos)

			%>
			<% if ocoffinvoice.HasPreScroll then %>
				<a href="?page=<%= ocoffinvoice.StartScrollPage-1 %>&research=on<%= strparam %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ocoffinvoice.StartScrollPage to ocoffinvoice.FScrollCount + ocoffinvoice.StartScrollPage - 1 %>
				<% if i>ocoffinvoice.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ocoffinvoice.HasNextScroll then %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>


<%
set ocoffinvoice = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
