<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/brandStaticCls.asp"-->
<%

dim yyyy1, mm1
Dim makerID, page, ordby
dim i

yyyy1	= req("yyyy1", Left(DateAdd("m", -1, Now()),4))
mm1		= req("mm1", Mid(DateAdd("m", -1, Now()),6,2))
makerID = req("makerID", "")
page = req("page", "1")
ordby = req("ordby", "itemreview")

dim rs
dim oCBrandService

set oCBrandService = new CBrandService
oCBrandService.FRectYYYYMM = yyyy1 & "-"& mm1
oCBrandService.FRectMakerid = makerID
oCBrandService.FCurrPage = page
oCBrandService.FPageSize = 100
oCBrandService.FRectOrderBy = ordby
rs = oCBrandService.GetBrandServiceByActionList()

class CBrandServiceItem
	public Fyyyymm
	public Fmakerid
	public FeventRegCnt
	public FnewItemRegCnt
	public FitemReviewCnt
	public FitemReviewPointSUM
	public FitemWishCnt
	public FbrandZzimCnt
	public FitemQnaRegCnt
	public FitemQnaAnsCnt
	public FitemQnaAnsDaySUM
	public Fregdate
	public Flastupdate

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

function toClass(rs, i)
	dim result
	'// yyyymm, makerid, eventRegCnt, newItemRegCnt, itemReviewCnt, itemReviewPointSUM, itemWishCnt, brandZzimCnt, itemQnaRegCnt, itemQnaAnsCnt, itemQnaAnsDaySUM, regdate, lastupdate
	set result = new CBrandServiceItem
	result.Fyyyymm 			= rs(1,i)
	result.Fmakerid 		= rs(2,i)
	result.FeventRegCnt 	= rs(3,i)
	result.FnewItemRegCnt 	= rs(4,i)
	result.FitemReviewCnt 	= rs(5,i)
	result.FitemReviewPointSUM 	= rs(6,i)
	result.FitemWishCnt 		= rs(7,i)
	result.FbrandZzimCnt 		= rs(8,i)
	result.FitemQnaRegCnt 		= rs(9,i)
	result.FitemQnaAnsCnt 		= rs(10,i)
	result.FitemQnaAnsDaySUM 	= rs(11,i)
	result.Fregdate 			= rs(12,i)
	result.Flastupdate 			= rs(13,i)

	set toClass = result
end function

dim rowCnt, item, val

%>

<script language='javascript'>
function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

function jsPopDashBoard(makerid) {
    var popwin = window.open("/admin/brandStatic/brandServicePointDashBoard.asp?menupos=4024&makerID=" + makerid,"jsPopDashBoard","width=1400 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
		<td align="left">
	       	년월 :
			<% DrawYMBox yyyy1,mm1 %>
			&nbsp;
			브랜드ID :
			<input type="text" class="text" name="makerID" value="<%=makerID%>">
		</td>

		<td rowspan="2" width="80" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	       	정렬순서 :
			<select class="select" name="ordby">
				<option value="makerid" <%= CHKIIF(ordby="makerid", "selected", "") %>>브랜드</option>
				<option value="evtcnt" <%= CHKIIF(ordby="evtcnt", "selected", "") %>>이벤트 등록건수</option>
				<option value="newitem" <%= CHKIIF(ordby="newitem", "selected", "") %>>신상품 등록건수</option>
				<option value="itemreview" <%= CHKIIF(ordby="itemreview", "selected", "") %>>상품후기 등록건수</option>
				<option value="itemwish" <%= CHKIIF(ordby="itemwish", "selected", "") %>>상품위시 등록건수</option>
				<option value="brndzzim" <%= CHKIIF(ordby="brndzzim", "selected", "") %>>브랜드찜 등록건수</option>
			</select>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

* 이벤트 건수는 이벤트시작일 기준입니다.<br />
* 신상품 건수는 판매개시일 기준입니다.<br />
* 상품후기 건수는 최근 3개월 등록건수입니다.<br />
* 상품위시 건수는 최근 3개월 등록건수입니다.<br />
* 브랜드찜 건수는 최근 3개월 등록건수입니다.<br />
* 상품문의 건수는 최근 3개월 등록건수입니다.<br />
* 7일 이상 답변이 달리지 않거나, 7일을 넘어 답변이 달리는 경우 답변등록일은 7일로 산정합니다.<br />
* 답변 등록일 산정에 휴일을 제외하지 않았습니다.

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">년월</td>
		<td width="250">브랜드</td>
		<td width="80">이벤트<br />등록건수</td>
		<td width="80">신상품<br />등록건수</td>
		<td width="80">상품후기<br />등록건수</td>
		<td width="80">상품후기<br />평점</td>
		<td width="80">상품위시<br />등록건수</td>
		<td width="80">브랜드찜<br />등록건수</td>
		<td width="80">상품문의<br />등록건수</td>
		<td width="80">평균답변<br />등록일수</td>
		<td>비고</td>
	</tr>
	<%
	If IsArray(rs) Then
		rowCnt = UBound(rs,2) + 1
		For i = 0 To UBound(rs,2)
			set item = toClass(rs, i)
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= item.Fyyyymm %></td>
		<td><a href="javascript:jsPopDashBoard('<%= item.Fmakerid %>')"><%= item.Fmakerid %></a></td>
		<td><%= FormatNumber(item.FeventRegCnt,0) %></td>
		<td><%= FormatNumber(item.FnewItemRegCnt,0) %></td>
		<td><%= FormatNumber(item.FitemReviewCnt,0) %></td>
		<td>
			<%
			if (item.FitemReviewCnt > 0) then
				response.write FormatNumber(item.FitemReviewPointSUM/item.FitemReviewCnt,2)
			else
				response.write "-"
			end if
			%>
		</td>
		<td><%= FormatNumber(item.FitemWishCnt,0) %></td>
		<td><%= FormatNumber(item.FbrandZzimCnt,0) %></td>
		<td><%= FormatNumber(item.FitemQnaRegCnt,0) %></td>
		<td>
			<%
			if (item.FitemQnaRegCnt > 0) then
				response.write FormatNumber(item.FitemQnaAnsDaySUM/item.FitemQnaRegCnt,2)
			else
				response.write "-"
			end if
			%>
		</td>
		<td></td>
	</tr>
	<%
		next
	end if
	%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19" align="center">
		<% if oCBrandService.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCBrandService.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCBrandService.StartScrollPage to oCBrandService.FScrollCount + oCBrandService.StartScrollPage - 1 %>
			<% if i>oCBrandService.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCBrandService.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
