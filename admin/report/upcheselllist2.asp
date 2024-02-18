<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  매출
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->

<%


dim searchtype
dim searchrect

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2,tmpDate
dim nowdate,searchNowDate, searchnextdate
dim orderserial,itemid,designer
dim research, oldlist
dim ckminusorder, vPurchaseType, vMDPick, rdsite, vMDPickMo, vMDPickMo1, vMDPickMo2, vMDPickMo3, vMDPickMo4, vMDPickMoArr
dim cdl,cdm,cds
dim channelDiv
Dim dispCate

nowdate = Left(CStr(now()),10)

designer = requestCheckvar(request("designer"),32)
orderserial = getNumeric(request("orderserial"))
itemid = getNumeric(request("itemid"))
''searchtype = requestCheckvar(request("searchtype"),10)  ''불필요
''searchrect = requestCheckvar(request("searchrect"),32)  ''불필요
yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)
dd1 = requestCheckvar(request("dd1"),2)
yyyy2 = requestCheckvar(request("yyyy2"),4)
mm2 = requestCheckvar(request("mm2"),2)
dd2 = requestCheckvar(request("dd2"),2)
research = requestCheckvar(request("research"),10)
oldlist = requestCheckvar(request("oldlist"),10)
ckminusorder = requestCheckvar(request("ckminusorder"),10)
vPurchaseType = requestCheckVar(request("purchasetype"),2)
vMDPick = requestCheckvar(request("mdpick"),10)
'rdsite = request.queryString("rdsite") '' 쿠키와 혼동을 피하기 위해 //2014/12/04

cdl = requestCheckvar(request("cdl"),3)
cdm = requestCheckvar(request("cdm"),3)
cds = requestCheckvar(request("cds"),3)
dispCate = requestCheckvar(request("disp"),16)

vMDPickMo = requestCheckvar(request("mdpickmo"),10)

vMDPickMo1  = requestCheckvar(request("mdpickmo1"),10)
vMDPickMo2  = requestCheckvar(request("mdpickmo2"),10)
vMDPickMo3  = requestCheckvar(request("mdpickmo3"),10)
vMDPickMo4  = requestCheckvar(request("mdpickmo4"),10)

vMDPickMoArr = ""
if(vMDPickMo1<>"") then vMDPickMoArr=vMDPickMoArr&"1,"
if(vMDPickMo2<>"") then vMDPickMoArr=vMDPickMoArr&"2,"
if(vMDPickMo3<>"") then vMDPickMoArr=vMDPickMoArr&"3,"
if(vMDPickMo4<>"") then vMDPickMoArr=vMDPickMoArr&"4,"    

if (vMDPickMoArr<>"") then vMDPickMoArr=LEFT(vMDPickMoArr,LEN(vMDPickMoArr)-1)

channelDiv  = NullFillWith(request("channelDiv"),"")

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	searchNowDate = nowdate
else
	'날짜 재배치
	tmpDate = DateSerial(yyyy1,mm1,dd1)
	yyyy1 = year(tmpDate)
	mm1 = Format00(2,month(tmpDate))
	dd1 = Format00(2,day(tmpDate))
	searchNowDate = tmpDate
end if

if yyyy2="" then
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
else
	'날짜 재배치
	tmpDate = DateSerial(yyyy2,mm2,dd2)
	yyyy2 = year(tmpDate)
	mm2 = Format00(2,month(tmpDate))
	dd2 = Format00(2,day(tmpDate))
end if
searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 & "-" & mm2 & "-" & dd2),1)),10)


dim cknodate,ckdelsearch,ckipkumdiv4
dim datetype
cknodate = requestCheckvar(request("cknodate"),10)
ckdelsearch = requestCheckvar(request("ckdelsearch"),10)
ckipkumdiv4 = requestCheckvar(request("ckipkumdiv4"),10)
datetype = requestCheckvar(request("datetype"),10)
if (datetype="") then datetype="jumunil"

dim page
dim ojumun

page = requestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new CJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = searchNowDate
	ojumun.FRectRegEnd = searchnextdate
end if

'response.write "rdsite:"&rdsite

if research="" then
	ckipkumdiv4 = "on"
end if

if searchtype="01" then
	ojumun.FRectBuyname = searchrect
elseif searchtype="02" then
	ojumun.FRectReqName = searchrect
elseif searchtype="03" then
	ojumun.FRectUserID = searchrect
elseif searchtype="04" then
	ojumun.FRectIpkumName = searchrect
elseif searchtype="06" then
	ojumun.FRectSubTotalPrice = searchrect
end if

ojumun.FRectDelNoSearch="on"
ojumun.FRectItemid = itemid
ojumun.FRectDesignerID = designer
ojumun.FPageSize = 500
ojumun.FCurrPage = page
ojumun.FRectIpkumDiv4 = ckipkumdiv4
ojumun.FRectOrderSerial = orderserial
ojumun.FRectDateType = datetype
ojumun.FRectOldJumun = oldlist
ojumun.FRectMinusOrderInclude = ckminusorder
ojumun.FRectBrandPurchaseType = vPurchaseType
ojumun.FRectCDL = cdl
ojumun.FRectCDM = cdm
ojumun.FRectCDS = cds
ojumun.FIsMDPick = vMDPick
ojumun.FIsMDPickMo = vMDPickMo
ojumun.FIsMDPickMoArr = vMDPickMoArr
ojumun.FIsRdSite = rdsite
ojumun.FRectChannelDiv = channelDiv
ojumun.FRectDispCate = dispCate
ojumun.SearchJumunListByupcheSelllist2


dim ix,iy

'response.write ojumun.FRectOrderSerial
'dbget.close()	:	response.End

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function ViewOrderDetail(itemid){


window.open("http://www.10x10.co.kr/street/designershop.asp?itemid=" + itemid,"sample");


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
</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<table class="a" border="0" cellpadding="3">
			<Tr>
				<td > 
				 기간: <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %> 
				  &nbsp;&nbsp; <input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역 
				 <br>
					<input type="radio" name="datetype" value="jumunil" <% if (datetype="jumunil") then response.write "checked" %> >주문일 
					<input type="radio" name="datetype" value="ipkumil" <% if (datetype="ipkumil") then response.write "checked" %> >결제일 
					<input type="radio" name="datetype" value="beadal" <% if (datetype="beadal") then response.write "checked" %> >배송일 

					(<input type="checkbox" name="ckipkumdiv4" <% if ckipkumdiv4="on" then response.write "checked" %> >결제완료이상검색) 
					(<input type="checkbox" name="ckminusorder" <% if ckminusorder="on" then response.write "checked" %> >반품포함) 
				</td> 
			</tr>
			<tr>
				<td  bgcolor="#FFFFFF" >
				 브랜드: <% drawSelectBoxDesigner "designer",designer %> 
				 &nbsp;&nbsp; item번호: <input type="text" name="itemid" value="<%= itemid %>" size="11" maxlength="16"> 
				</td>
			</tr>
			<tr>
					<td  bgcolor="#FFFFFF" >
				 <!-- #include virtual="/common/module/categoryselectbox.asp"-->  
					&nbsp;&nbsp;전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->   
				</td>
			</tr> 
			<tr>
				<td> 
				 채널구분 : <% drawSellChannelComboBox "channelDiv",channelDiv %>

				 &nbsp;&nbsp;구매유형: 
				 <% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
				&nbsp;&nbsp;<input type="checkbox" name="mdpick" <% if vMDPick="on" then response.write "checked" %> />PC웹 MD'S Pick 상품 
				&nbsp; 
				<!--
				<input type="checkbox" name="mdpickmo" <% if vMDPickMo="on" then response.write "checked" %> />모바일MD'S Pick 상품
				-->
				
				<input type="checkbox" name="mdpickmo1" <% if vMDPickMo1="on" then response.write "checked" %> />모바일MD'S Pick 상품
				<input type="checkbox" name="mdpickmo2" <% if vMDPickMo2="on" then response.write "checked" %> />모바일 New 상품
				<input type="checkbox" name="mdpickmo4" <% if vMDPickMo4="on" then response.write "checked" %> />모바일 On Sales 상품
				
				<!--
				<input type="checkbox" name="mdpickmo3" <% if vMDPickMo3="on" then response.write "checked" %> />모바일 Best 상품
				 -->
				<% if (FALSE) then %>
				<!--&nbsp; <input type="checkbox" name="rdsite" <% if rdsite="on" then response.write "checked" %> />모바일판매만  //-->
			    <% end if %>
		</td>
	</tr> 
</table>
</td>
<td class="a" align="center"  bgcolor="<%= adminColor("gray") %>">
					<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
				</td>
			</tr>
</table>
</form>
<Br>
(최대 <%= ojumun.FPageSize %> 건검색) 
<Br>
<table width="100%" border="1" cellpadding="0" cellspacing="0" class="a">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">상품번호</td>
		<td >상품</td>
		<td width="80">옵션</td>
		<td width="70">가격</td>
		<td width="60">총갯수</td>
		<td width="70">합계매출</td>
		<td width="70">합계수익</td>
	</tr>
<% if ojumun.FResultCount<1 then %>
	<tr align="center">
		<td colspan="12" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FResultCount -1 %>
<%
Dim sumprice,totalsumprice,totalitemNo, sumprofit, totalProfit
sumprice = ojumun.FMasterItemList(ix).FItemCost * ojumun.FMasterItemList(ix).FItemNo
sumprofit = (ojumun.FMasterItemList(ix).FItemCost * ojumun.FMasterItemList(ix).FItemNo) - (ojumun.FMasterItemList(ix).Fbuycash * ojumun.FMasterItemList(ix).FItemNo)

totalitemNo=totalitemNo + ojumun.FMasterItemList(ix).FItemNo
totalsumprice =  totalsumprice + sumprice
totalProfit = totalProfit + sumprofit
%>
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr align="center" class="a">
	<% else %>
	<tr align="center" class="gray">
	<% end if %>
		<td align="center" height="25"><a href="<%=wwwURL%>/<%= ojumun.FMasterItemList(ix).FItemID  %>" class="zzz" target="_blank"><%= ojumun.FMasterItemList(ix).FItemID  %></a></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FItemName %></td>
		<% if (ojumun.FMasterItemList(ix).FItemOptionStr="") then %>
			<td align="center">&nbsp;</td>
		<% else %>
			<td align="center"><%= ojumun.FMasterItemList(ix).FItemOptionStr %></td>
		<% end if %>
		<td align="center"><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0)  %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FItemNo %></td>
		<td align="center"><%= FormatNumber(sumprice,0) %></td>
		<td align="center"><%= FormatNumber(sumprofit,0) %></td>
	</tr>
	<% next %>
	<tr align="center">
		<td colspan="7" height="25" align="right">현재 페이지의 상품 합계 개수 :<font color="red"><%= totalitemNo %></font>&nbsp;/&nbsp;합계 매출 : <font color="red"><% =FormatNumber(totalsumprice,0) %></font>원&nbsp;/&nbsp;합계 수익 : <font color="red"><% =FormatNumber(totalprofit,0) %></font>원&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	</tr>
<% end if %>

	<tr align="center">
<!--
		<td colspan="7" height="30" align="center">
		<% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
			<% if ix>ojumun.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ojumun.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
-->
		</td>
	</tr>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->