<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  채널 베스트셀러
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/category_reportcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdate,searchnextdate
dim orderserial,itemid,oreport
dim topn,cdl,cdm,page
dim ckpointsearch,ckipkumdiv4
dim ix,iy,cknodate
dim order_desum
dim rectdispy, rectselly, ordertype, rdsite
dim oldlist, sitename
dim sellchnl, userlevel
dim vPurchasetype, inc3pl
Dim dispCate, DlvType
Dim optExists, research

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
cdl = request("cdl")
cdm = request("cdm")
orderserial = request("orderserial")
itemid = request("itemid")
topn = request("topn")
ckpointsearch = request("ckpointsearch")
cknodate = request("cknodate")
order_desum = request("order_desum")
rectdispy = request("rectdispy")
rectselly = request("rectselly")
ordertype = request("ordertype")
if ordertype="" then ordertype="ea"
oldlist = request("oldlist")
rdsite = request("rdsite")
sitename = request("sitename")
sellchnl  = NullFillWith(request("sellchnl"),"")
dispCate = requestCheckvar(request("disp"),16)

vPurchasetype = request("purchasetype")
inc3pl = request("inc3pl")
optExists = request("optExists")
research = request("research")
userlevel = request("userlevel")
DlvType = request("dlvtype")

''기본조건 옵션별로 보기
If (research = "") Then
	optExists = ""
End If

if sitename<>"" then
	if rdsite<>"on" then
		rdsite = sitename
	else
		sitename = ""
	end if
end if

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

topn = request("topn")
if (topn="") then topn=100

set oreport = new CCategoryReport

if cknodate="" then
	oreport.FRectFromDate = yyyy1 + "-" + mm1 + "-" + dd1
	oreport.FRectToDate = searchnextdate
end if

oreport.FRectCD1 = cdl
oreport.FRectCD2 = cdm
oreport.FPageSize = topn
oreport.FCurrPage = page
oreport.FRectDispY = rectdispy
oreport.FRectSellY = rectselly
oreport.FRectRdsite = rdsite
oreport.FRectOrdertype = ordertype
oreport.FRectOldJumun = oldlist
oreport.FRectSellChannelDiv = sellchnl
oreport.FRectPurchasetype = vPurchasetype ''2014/01/27
oreport.FRectInc3pl = inc3pl  ''2014/01/27
oreport.FRectDispCate = dispCate
oreport.FRectOptExists = optExists
oreport.FRectUserLevel = userlevel
oreport.FRectDlvType = DlvType
oreport.ONSearchCategoryBestseller

'// 사이트구분로검색시작
Sub Drawsitename(selectboxname, sitename)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str

	response.write "<select name='" & selectboxname & "' class='select'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if sitename ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">전체</option>"								'선택이란 단어가 나오도록.
	response.write "<option value='10x10' "
		if sitename ="10x10" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">10x10</option>"

	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = " select id from [db_partner].[dbo].tbl_partner"
	userquery = userquery + " where 1=1"
	userquery = userquery + " and id <> '' and id is not null"
	userquery = userquery + " and userdiv= '999' and isusing='Y' "
	userquery = userquery + " group by id"

	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(sitename) = Lcase(rsget("id")) then 	'검색될 이름과 db에 저장된 이름을 비교해서 맞다면, //
				tem_str = " selected"								'// 검색어로 선택
			end if

			response.write "<option value='" & rsget("id") & "' " & tem_str & ">" & rsget("id") & "</option>"
			tem_str = ""				'rsget에 id를 선택하고 검색할 값으로 선택
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function ViewOrderDetail(itemid){
    var popwin = window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"category_prd");
    popwin.focus();
}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function MonthDiff(d1, d2) {
	d1 = d1.split("-");
	d2 = d2.split("-");

	d1 = new Date(d1[0], d1[1] - 1, d1[2]);
	d2 = new Date(d2[0], d2[1] - 1, d2[2]);

	var d1Y = d1.getFullYear();
	var d2Y = d2.getFullYear();
	var d1M = d1.getMonth();
	var d2M = d2.getMonth();

	return (d2M+12*d2Y)-(d1M+12*d1Y);
}

function ReSearch(ifrm){
	var v = ifrm.topn.value;
	if (!IsDigit(v)){
		alert('숫자만 가능합니다.');
		ifrm.topn.focus();
		return;
	}

	if (v>1000){
		alert('천건 이하만 검색가능합니다.');
		ifrm.topn.focus();
		return;
	}

	if ((CheckDateValid(ifrm.yyyy1.value, ifrm.mm1.value, ifrm.dd1.value) == true) && (CheckDateValid(ifrm.yyyy2.value, ifrm.mm2.value, ifrm.dd2.value) == true)) {
		if (MonthDiff(ifrm.yyyy1.value + "-" + ifrm.mm1.value + "-" + ifrm.dd1.value, ifrm.yyyy2.value + "-" + ifrm.mm2.value + "-" + ifrm.dd2.value) >= 3) {
			alert("최대 3개월까지만 검색이 가능합니다.");
			return;
		}

		ifrm.submit();
	}

	//ifrm.submit();
}
</script>

	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="70" bgcolor="<%= adminColor("gray") %>">검색조건</td>
		<td align="left">
			<table class="a" border="0" cellpadding="3">
			<tr>
				<td class="a" >
				기간:
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
			</td>
		</tr>
		<tr>
			<td>
				관리카테고리:
				<% SelectBoxCategoryLarge cdl %>&nbsp;
				<% if cdl="110" then DrawSelectBoxCategoryMid "cdm",cdl,cdm %>
				&nbsp;&nbsp;전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->


			</td>
		</tr>
		<tr>
			<td>
			  사이트: <% Drawsitename "sitename",sitename %>
	      &nbsp;&nbsp;채널구분:
	        <% drawSellChannelComboBoxGroup "sellchnl",sellchnl %>  
			&nbsp;&nbsp;구매유형: 
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
			&nbsp;&nbsp;<b>매출처:</b>
		    <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
		    &nbsp;&nbsp;회원등급:
		    <% Call DrawselectboxUserLevel("userlevel", userlevel, "") %>
			&nbsp;&nbsp;배송구분:
			<select name="dlvtype" class="select">
				<option value="">전체</option>
				<option value="N" <%=chkIIF(DlvType="N","selected","")%>>텐바이텐 배송</option>
				<option value="Y" <%=chkIIF(DlvType="Y","selected","")%>>업체 배송</option>
			</select>
			</td>
		</tr>
		<tr>
			<td>
				<input type="checkbox" name="rectselly" <% if rectselly="on" then response.write "checked" %> >판매하는아이템만
				<input type="checkbox" name="rectdispy" <% if rectdispy="on" then response.write "checked" %> >전시하는아이템만
				<input type="checkbox" name="rdsite" <% if rdsite="on" then response.write "checked" %> >모바일판매만
				&nbsp;&nbsp;정렬:
				<input type="radio" name="ordertype" value="ea" <% if ordertype="ea" then response.write "checked" %>>수량순
				<input type="radio" name="ordertype" value="totalprice" <% if ordertype="totalprice" then response.write "checked" %>>매출순
				<input type="radio" name="ordertype" value="gain" <% if ordertype="gain" then response.write "checked" %>>수익순
				<input type="radio" name="ordertype" value="unitCost" <% if ordertype="unitCost" then response.write "checked" %>>객단가순
				   &nbsp;&nbsp; 검색갯수 :
				<input type="text" name="topn" value="<%= topn %>" size="7" maxlength="6" >
				&nbsp;<input type="checkbox" name="optExists" <%= ChkIIF(optExists="on","checked","") %> >옵션별로 보기
			</td>
		</tr>

	</table>
	</td>
	<td class="a" align="center"  bgcolor="<%= adminColor("gray") %>">
			<a href="javascript:ReSearch(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
	</form>
	<br>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#E6E6E6">
	<td colspan="12" height="25" align="right">검색결과 : 총 <font color="red"><% = oreport.FResultCount %></font>개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor="#E6E6E6">
	<td width="30" align="center">순위</td>
	<td width="50" align="center">이미지</td>
	<td width="50" align="center">상품번호</td>
	<td  align="center">상품</td>
<% If optExists = "" Then %>
	<td width="100" align="center">브랜드ID</td>
<% Else %>
	<td width="50">단가</td>
	<td width="100" align="center">브랜드ID</td>
	<td width="80" align="center">옵션</td>
<% End If %>
	<td width="65" align="center">판매갯수</td>
	<td width="65" align="center">(현재)판매가</td>
	<td width="100" align="center">판매가합</td>
	<td width="100" align="center">매입가합</td>
	<td width="100" align="center">수익</td>
	<td width="70" align="center">마진율</td>
</tr>
<% if oreport.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="12" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to oreport.FResultCount -1 %>
<%
Dim totalsumprice, totalbuyprice, totalitemno
totalitemno   =  totalitemno + oreport.FItemList(ix).FItemNo
totalsumprice =  totalsumprice + oreport.FItemList(ix).Fselltotal
totalbuyprice =  totalbuyprice + oreport.FItemList(ix).Fbuytotal

%>
	<tr class="a" bgcolor="#FFFFFF" height="50">
		<td align="center"><%=ix+1%></td>
		<td><img src="<%= oreport.FItemList(ix).FImageSmall %>" width=50></td>
		<td align="center" height="25"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oreport.FItemList(ix).FItemID %>" class="zzz" target="_blank"><%= oreport.FItemList(ix).FItemID  %></a></td>
		<td align="center"><%= oreport.FItemList(ix).FItemName %></td>
	<% If optExists = "" Then %>
		<td align="center"><%= oreport.FItemList(ix).FMakerid %></td>
	<% Else %>
		<td align="center"><%= Chkiif(optExists="on", FormatNumber(oreport.FItemList(ix).FItemCost,0), "") %></td>
		<td align="center"><%= oreport.FItemList(ix).FMakerid %></td>
		<% if (oreport.FItemList(ix).FItemOptionStr="") then %>
			<td align="center">&nbsp;</td>
		<% else %>
			<td align="center"><%= oreport.FItemList(ix).FItemOptionStr %></td>
		<% end if %>
	<% End If %>
		<td align="center"><%= oreport.FItemList(ix).FItemNo %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Forgprice,0) %>
		    <%
			'할인가
			if oreport.FItemList(ix).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((oreport.FItemList(ix).Forgprice-oreport.FItemList(ix).Fsailprice)/oreport.FItemList(ix).Forgprice*100) & "%할)" & FormatNumber(oreport.FItemList(ix).Fsailprice,0) & "</font>"
			end if
			'쿠폰가
			if oreport.FItemList(ix).FitemCouponYn="Y" then
				Select Case oreport.FItemList(ix).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oreport.FItemList(ix).GetCouponAssignPrice(),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oreport.FItemList(ix).GetCouponAssignPrice(),0) & "</font>"
				end Select

			end if%></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Fselltotal,0) %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Fbuytotal,0) %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Fselltotal-oreport.FItemList(ix).Fbuytotal,0) %></td>
	    <td align="center">
	        <% if oreport.FItemList(ix).Fselltotal<>0 then %>
	        <%= 100-CLng(oreport.FItemList(ix).Fbuytotal/oreport.FItemList(ix).Fselltotal*100*100)/100 %> %
	        <% end if %>
	    </td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" align="center">Total</td>
	    <td colspan="<%=CHKIIF(optExists ="", "3", "5") %>"></td>
	    <td align="center"><%= FormatNumber(totalitemno,0) %></td>
	    <td>&nbsp;</td>
	    <td align="right"><%= FormatNumber(totalsumprice,0) %></td>
	    <td align="right"><%= FormatNumber(totalbuyprice,0) %></td>
	    <td align="right"><%= FormatNumber(totalsumprice-totalbuyprice,0) %></td>
	    <td align="center">
	        <% if totalsumprice<>0 then %>
	        <%= 100-CLng(totalbuyprice/totalsumprice*100*100)/100 %> %
	        <% end if %>
	    </td>
	</tr>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
