<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/deliverypolicycls.asp"-->
<%

dim designer, mduserid , catecode, defaultdeliveryType, isusingbrand, isusingitem, mwdiv
dim i

dim currpage, pagesize

currpage 	= requestCheckvar(request("currpage"),32)
designer 	= requestCheckvar(request("designer"),32)
mduserid 	= requestCheckvar(request("mduserid"),32)
catecode 	= requestCheckvar(request("catecode"),3)
defaultdeliveryType 	= requestCheckvar(request("defaultdeliveryType"),1)
isusingbrand 	= requestCheckvar(request("isusingbrand"),1)
isusingitem 	= requestCheckvar(request("isusingitem"),1)
mwdiv 	= requestCheckvar(request("mwdiv"),1)

pagesize = 5000
if (currpage = "") then
	currpage = 1
end if

'==============================================================================
dim ODeliveryPolicy

set ODeliveryPolicy = new CDeliveryPolicy

ODeliveryPolicy.FPageSize = pagesize
ODeliveryPolicy.FCurrPage = currpage
ODeliveryPolicy.FRectUserID = designer
ODeliveryPolicy.FRectMDUserID = mduserid
ODeliveryPolicy.FRectCategoryCode = catecode
ODeliveryPolicy.FRectDefaultDeliveryType = defaultdeliveryType
ODeliveryPolicy.FRectIsUsingBrand = isusingbrand
ODeliveryPolicy.FRectIsUsingItem = isusingitem
ODeliveryPolicy.FRectMWDiv = mwdiv

ODeliveryPolicy.GetList

'Excel Header
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=brandDeliveryPolicy_page" & CStr(currpage) & ".xls"
Response.CacheControl = "public"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="application/vnd.ms-excel;charset=euc-kr">
<style type='text/css'>
	td {border:0.3px solid #666;}
	.txt {mso-number-format:'\@'}
	.num {mso-number-format:"#,##0"}
</style>
</head>
<body>
<table cellpadding="3" cellspacing="1">
	<tr>
		<td colspan="16">
			검색결과 : <b><%= ODeliveryPolicy.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= currpage %> / <%= ODeliveryPolicy.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center">
    	<td rowspan="2">브랜드ID</td>
		<td rowspan="2">스트리트명</td>
      	<td rowspan="2">회사명</td>
      	<td rowspan="2" width="80">브랜드 배송정책</td>

      	<td colspan="6">상품정보(상품수)</td>
      	<td colspan="3">상품정보(거래구분)</td>
      	<td rowspan="2">전체상품수</td>

      	<td colspan="2">개별배송비기준(원)</td>
    </tr>
	<tr align="center">
      	<td>1만원미만</td>
      	<td>1만원대</td>
      	<td>2만원대</td>
      	<td>3만원대</td>
      	<td>5만원대</td>
      	<td>5만원이상</td>

      	<td>업체</td>
      	<td>위탁</td>
      	<td>매입</td>

      	<td>무료배송 최소금액</td>
      	<td>개별배송비</td>
    </tr>
<% if ODeliveryPolicy.FresultCount < 1 then %>
	<tr align="center">
		<td colspan="16" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for i = 0 to ODeliveryPolicy.FresultCount - 1 %>

    <tr align="center">
    	<td class="txt"><%= ODeliveryPolicy.FItemList(i).Fuserid %></td>
    	<td class="txt"><%= ODeliveryPolicy.FItemList(i).Fsocname_kor %></td>
      	<td class="txt"><%= ODeliveryPolicy.FItemList(i).Fconame %></td>
      	<td class="txt"><% if (ODeliveryPolicy.FItemList(i).FdefaultdeliveryType <> "ETC") then %><%= ODeliveryPolicy.FItemList(i).FdefaultdeliveryType %><% end if %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).Fprice0 %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).Fprice10000 %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).Fprice20000 %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).Fprice30000 %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).Fprice40000 %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).Fprice50000 %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).Fupchecount %></td>
      	<td class="num"><% if (ODeliveryPolicy.FItemList(i).FdefaultdeliveryType <> "ETC") and ((ODeliveryPolicy.FItemList(i).Fwitakcount <> 0) or (ODeliveryPolicy.FItemList(i).Fmaeipcount <> 0)) then %><font color=red><b><% end if %><%= ODeliveryPolicy.FItemList(i).Fwitakcount %></td>
      	<td class="num"><% if (ODeliveryPolicy.FItemList(i).FdefaultdeliveryType <> "ETC") and ((ODeliveryPolicy.FItemList(i).Fwitakcount <> 0) or (ODeliveryPolicy.FItemList(i).Fmaeipcount <> 0)) then %><font color=red><b><% end if %><%= ODeliveryPolicy.FItemList(i).Fmaeipcount %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).Fitemcount %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).FdefaultFreeBeasongLimit %></td>
      	<td class="num"><%= ODeliveryPolicy.FItemList(i).FdefaultDeliverPay %></td>
    </tr>
	<% next %>
<% end if %>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->