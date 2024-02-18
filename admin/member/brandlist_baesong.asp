<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/deliverypolicycls.asp"-->

<%

dim designer, mduserid , catecode, defaultdeliveryType, isusingbrand, isusingitem, mwdiv
dim i

dim currpage

currpage 	= requestCheckvar(request("currpage"),32)
designer 	= requestCheckvar(request("designer"),32)
mduserid 	= requestCheckvar(request("mduserid"),32)
catecode 	= requestCheckvar(request("catecode"),3)
defaultdeliveryType 	= requestCheckvar(request("defaultdeliveryType"),4)
isusingbrand 	= requestCheckvar(request("isusingbrand"),1)
isusingitem 	= requestCheckvar(request("isusingitem"),1)
mwdiv 	= requestCheckvar(request("mwdiv"),1)

if (currpage = "") then
	currpage = 1
end if



'==============================================================================
dim ODeliveryPolicy

set ODeliveryPolicy = new CDeliveryPolicy

ODeliveryPolicy.FPageSize = 50
ODeliveryPolicy.FCurrPage = currpage
ODeliveryPolicy.FRectUserID = designer
ODeliveryPolicy.FRectMDUserID = mduserid
ODeliveryPolicy.FRectCategoryCode = catecode
ODeliveryPolicy.FRectDefaultDeliveryType = defaultdeliveryType
ODeliveryPolicy.FRectIsUsingBrand = isusingbrand
ODeliveryPolicy.FRectIsUsingItem = isusingitem
ODeliveryPolicy.FRectMWDiv = mwdiv



ODeliveryPolicy.GetList

%>
<script language='javascript'>
function popItemSellEdit(designerid,mwdiv,usingyn){
	var popwin = window.open('/admin/shopmaster/itemviewset.asp?menupos=24&makerid=' + designerid + '&mwdiv=' + mwdiv + '&usingyn=' + usingyn  ,'popItemSellEdit','width=1000,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}
</script>
<!-- 검색 시작 -->



<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
    		브랜드 : <% drawSelectBoxDesignerwithName "designer", designer %>
			&nbsp;
			담당자 : <% drawSelectBoxCoWorker "mduserid", mduserid %>
			&nbsp;
			카테고리 : <% SelectBoxBrandCategory "catecode", catecode %>

		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			브랜드 배송정책 : <% drawPartnerCommCodeBox True,"deliveryType","defaultdeliveryType", defaultdeliveryType,"" %>
	     	&nbsp;
			브랜드 사용여부 :
			<select class="select" name="isusingbrand">
		     	<option value='' >전체</option>
		     	<option value='Y' <% if (isusingbrand = "Y") then %>selected<% end if %>>사용</option>
		     	<option value='N' <% if (isusingbrand = "N") then %>selected<% end if %>>사용안함</option>
	     	</select>
	     	&nbsp;
			상품 사용여부 :
			<select class="select" name="isusingitem">
		     	<option value='' selected>전체</option>
		     	<option value='Y' <% if (isusingitem = "Y") then %>selected<% end if %>>사용</option>
		     	<option value='N' <% if (isusingitem = "N") then %>selected<% end if %>>사용안함</option>
	     	</select>
			<!-- 무지 느리다. 2015-04-08, skyer9
	     	&nbsp;
			거래구분 :
			<select class="select" name="mwdiv">
		     	<option value='' selected>전체</option>
		     	<option value='M' <% if (mwdiv = "M") then %>selected<% end if %>>매입</option>
				<option value='W' <% if (mwdiv = "W") then %>selected<% end if %>>위탁</option>
				<option value='U' <% if (mwdiv = "U") then %>selected<% end if %>>업체</option>
	     	</select>
			-->
            &nbsp;
            <input type="checkbox" name="exctpl" value="Y" checked disabled> 3PL 제외

		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<!-- 엑셀받기 -->
<%
	Dim exlPsz, exlPg
	exlPsz = 5000
	exlPg = ceil(ODeliveryPolicy.FTotalCount/exlPsz)
%>
<script>
	function fnGetExcel(pg) {
		window.open("brandlist_baesong_excel.asp?currpage="+pg+"&designer=<%=designer%>&mduserid=<%=mduserid%>&defaultdeliveryType=<%=defaultdeliveryType%>&catecode=<%=catecode%>&mwdiv=<%=mwdiv%>&isusingbrand=<%=isusingbrand%>&isusingitem=<%=isusingitem%>");
	}
</script>

<div style="text-align:right; margin:10px 5px;">
	<select id="exlPage" class="select" style="vertical-align: middle;">
	<% for i=1 to exlPg %>
	<option value="<%=i%>"><%=((i-1)*exlPsz)+1%>~<%=chkIIF(i*exlPsz<ODeliveryPolicy.FTotalCount,i*exlPsz,ODeliveryPolicy.FTotalCount)%></option>
	<% next %>
	</select>
	<img src="/images/btn_excel.gif" onClick="fnGetExcel(document.getElementById('exlPage').value)" style="cursor:pointer;vertical-align: middle;" />
</div>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">
			검색결과 : <b><%= ODeliveryPolicy.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= currpage %> / <%= ODeliveryPolicy.FTotalPage %></b>
		</td>
		<td colspan="16" align=right>
			*상품이 모두 업체배송일 경우만 설정변경이 가능합니다.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2">브랜드ID</td>
		<td rowspan="2">스트리트명</td>
      	<td rowspan="2">회사명</td>
      	<td rowspan="2" width="80">브랜드<br>배송정책</td>

      	<td colspan="6">상품정보(상품수)</td>
      	<td colspan="3">상품정보(거래구분)</td>
      	<td rowspan="2" width="60">전체상품수</td>

      	<td colspan="2">개별배송비기준(원)</td>
      	<td rowspan="2" width="80">비고</td>

    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      	<td width="60">1만원미만</td>
      	<td width="50">1만원대</td>
      	<td width="50">2만원대</td>
      	<td width="50">3만원대</td>
      	<td width="50">5만원대</td>
      	<td width="60">5만원이상</td>

      	<td width="40">업체</td>
      	<td width="40">위탁</td>
      	<td width="40">매입</td>

      	<td width="60">무료배송<br>최소금액</td>
      	<td width="70">개별배송비</td>
    </tr>
<% if ODeliveryPolicy.FresultCount < 1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="20" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for i = 0 to ODeliveryPolicy.FresultCount - 1 %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= ODeliveryPolicy.FItemList(i).Fuserid %></td>
    	<td><%= ODeliveryPolicy.FItemList(i).Fsocname_kor %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fconame %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).FdefaultdeliveryType %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice0 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice10000 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice20000 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice30000 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice40000 %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fprice50000 %></td>
      	<td><a href="javascript:popItemSellEdit('<%= ODeliveryPolicy.FItemList(i).Fuserid %>','U','<%= isusingitem %>');"><%= ODeliveryPolicy.FItemList(i).Fupchecount %></a></td>
      	<td><a href="javascript:popItemSellEdit('<%= ODeliveryPolicy.FItemList(i).Fuserid %>','W','<%= isusingitem %>');"><% if (ODeliveryPolicy.FItemList(i).FdefaultdeliveryType <> "ETC") and ((ODeliveryPolicy.FItemList(i).Fwitakcount <> 0) or (ODeliveryPolicy.FItemList(i).Fmaeipcount <> 0)) then %><font color=red><b><% end if %><%= ODeliveryPolicy.FItemList(i).Fwitakcount %></a></td>
      	<td><a href="javascript:popItemSellEdit('<%= ODeliveryPolicy.FItemList(i).Fuserid %>','M','<%= isusingitem %>');"><% if (ODeliveryPolicy.FItemList(i).FdefaultdeliveryType <> "ETC") and ((ODeliveryPolicy.FItemList(i).Fwitakcount <> 0) or (ODeliveryPolicy.FItemList(i).Fmaeipcount <> 0)) then %><font color=red><b><% end if %><%= ODeliveryPolicy.FItemList(i).Fmaeipcount %></a></td>
      	<td><%= ODeliveryPolicy.FItemList(i).Fitemcount %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).FdefaultFreeBeasongLimit %></td>
      	<td><%= ODeliveryPolicy.FItemList(i).FdefaultDeliverPay %></td>
      	<td>
		<!-- 팀장이상 + 텐바이텐사업팀 - MD파트 정직이상 설정가능(권한제한 낮춤:2011.09.01) -->
		<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or ((session("ssAdminLsn") <= "4") and (session("ssAdminPsn") = "11"))) then %>
			<% if (ODeliveryPolicy.FItemList(i).FdefaultdeliveryType = "ETC") then %>
				<% if (ODeliveryPolicy.FItemList(i).Fwitakcount = 0) and (ODeliveryPolicy.FItemList(i).Fmaeipcount = 0) then %>
			<input type="button" class="button" value="신규설정" onClick="PopBrandAdminUsingChange('<%= ODeliveryPolicy.FItemList(i).Fuserid %>')">
				<% end if %>
      		<% else %>
      		<input type="button" class="button" value="설정변경" onClick="PopBrandAdminUsingChange('<%= ODeliveryPolicy.FItemList(i).Fuserid %>')">
      		<% end if %>
      	<% end if %>
      	</td>
    </tr>
	<% next %>
<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
    		<% if ODeliveryPolicy.HasPreScroll then %>
    			<a href="?currpage=<%= ODeliveryPolicy.StartScrollPage-1 %>&menupos=<%= menupos %>&designer=<%= designer %>&mduserid=<%= mduserid %>&catecode=<%= catecode %>&defaultdeliveryType=<%= defaultdeliveryType %>&isusingbrand=<%= isusingbrand %>&isusingitem=<%= isusingitem %>">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i = (0 + ODeliveryPolicy.StartScrollPage) to (ODeliveryPolicy.FScrollCount + ODeliveryPolicy.StartScrollPage - 1) %>
    			<% if i>ODeliveryPolicy.FTotalpage then Exit for %>
    			<% if CStr(currpage)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="?currpage=<%= i %>&menupos=<%= menupos %>&designer=<%= designer %>&mduserid=<%= mduserid %>&catecode=<%= catecode %>&defaultdeliveryType=<%= defaultdeliveryType %>&isusingbrand=<%= isusingbrand %>&isusingitem=<%= isusingitem %>">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if ODeliveryPolicy.HasNextScroll then %>
    			<a href="?currpage=<%= i %>&menupos=<%= menupos %>&designer=<%= designer %>&mduserid=<%= mduserid %>&catecode=<%= catecode %>&defaultdeliveryType=<%= defaultdeliveryType %>&isusingbrand=<%= isusingbrand %>&isusingitem=<%= isusingitem %>">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>


















<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
