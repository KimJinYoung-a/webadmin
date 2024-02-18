<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품수정
' History : 2009.04.17 최초생성자모름
'			2016.07.06 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

dim itemid, oitem, oitemvideo
dim makerid
dim chkMWAuth 'mw 변경가능한 권한인지 체크 
dim rentalItemFlag

itemid = request("itemid")
makerid = request("makerid")
menupos = request("menupos")
if (itemid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

Set oitemvideo = New CItem
oitemvideo.FRectItemId = itemid
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetItemContentsVideo

dim oitemAddImage
set oitemAddImage = new CItemAddImage
oitemAddImage.FRectItemID = itemid

if oitem.FResultCount>0 then
    ''상품 추가이미지 접수.
    oitemAddImage.GetOneItemAddImageList
end if

''연관상품 목록 접수
dim strItemRelation
strItemRelation = GetItemRelationStr(itemid)

'==============================================================================
''업체 기본계약 구분 
dim defaultmargin, defaultmaeipdiv, defaultFreeBeasongLimit, defaultDeliverPay, defaultDeliveryType
dim jungsangubun, companyno
dim sqlStr
sqlStr = "select c.defaultmargine, c.maeipdiv as defaultmaeipdiv, "
sqlStr = sqlStr + " IsNULL(c.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit,"
sqlStr = sqlStr + " IsNULL(c.defaultDeliverPay,0) as defaultDeliverPay,"
sqlStr = sqlStr + " IsNULL(c.defaultDeliveryType,'') as defaultDeliveryType"
sqlStr = sqlStr + "  , p.jungsan_gubun, p.company_no "
sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c as c "
sqlStr = sqlStr + "  inner join db_partner.dbo.tbl_partner as p on c.userid = p.id " 
sqlStr = sqlStr + " where c.userid='" & oitem.FOneItem.Fmakerid & "'" 
rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        defaultmargin           = rsget("defaultmargine")
        defaultmaeipdiv         = rsget("defaultmaeipdiv")
        defaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
        defaultDeliverPay       = rsget("defaultDeliverPay")
        defaultDeliveryType     = rsget("defaultDeliveryType")
        jungsangubun						= rsget("jungsan_gubun")
        companyno							= rsget("company_no")
    end if
rsget.close

'==============================================================================
'세일마진
dim sailmargine, orgmargine, margine

''수정
if oitem.FOneItem.Fsailprice<>0 then
	sailmargine = 100-CCur(oitem.FOneItem.Fsailsuplycash/oitem.FOneItem.Fsailprice*100*100*100*100)/100/100/100
else
	sailmargine = 0
end if

if oitem.FOneItem.Forgprice<>0 then
	orgmargine = 100-CCur(oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice*100*100*100*100)/100/100/100
else
	orgmargine = 0
end if

if oitem.FOneItem.Fsellcash<>0 then
	margine = 100-CCur(oitem.FOneItem.Fbuycash/oitem.FOneItem.Fsellcash*100*100*100*100)/100/100/100     ''''*100*100 / 100/100 추가
else
	margine = 0
end if


'mw 변경가능 권한인지 체크
chkMWAuth = False
IF (Not oitem.FOneItem.FisCurrStockExists)  or C_ADMIN_AUTH  THEN chkMWAuth = True 

'// 렌탈 상품은 일단 테스트로 특정 유저만 노출함
If C_ADMIN_AUTH Then
	rentalItemFlag = true
Else
	rentalItemFlag = true
End If
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
<!-- #include file="./itemmodify_javascript.asp"-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>상품수정</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>등록된 상품을 수정합니다.</b>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/itemmodifyWithImage_process.asp" onsubmit="return false;" enctype="multipart/form-data" style="margin:0;">
<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
<input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">
<input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
<input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
<!-- 업체 기본 계약 구분 -->
<input type="hidden" name="defaultmargin" value="<%= defaultmargin %>">
<input type="hidden" name="defaultmaeipdiv" value="<%= defaultmaeipdiv %>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%= defaultFreeBeasongLimit %>">
<input type="hidden" name="defaultDeliverPay" value="<%= defaultDeliverPay %>">
<input type="hidden" name="defaultDeliveryType" value="<%= defaultDeliveryType %>">
<input type="hidden" name="jungsangubun" value="<%=jungsangubun%>">
<input type="hidden" name="companyno" value="<%=companyno%>">
<input type="hidden" name="sellreservedate" value="<%=oitem.FOneItem.Fsellreservedate%>"><!--오픈예약일-->
<input type="hidden" name="chkModSR" value="N"><!--오픈예약 취소여부-->
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>

</table>
<!-- 표 상단바 끝-->

<!-- 1.일반정보 --> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>1.일반정보</strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">브랜드ID :</td>
	<!--td bgcolor="#FFFFFF" colspan="3"><% 'SelectBoxDesignerItem oitem.FOneItem.Fmakerid %> (사용업체만 표시됩니다)</td-->
	<td bgcolor="#FFFFFF" colspan="3"><% NewDrawSelectBoxDesignerChangeMargin "designer", oitem.FOneItem.Fmakerid, "marginData", "TnDesignerNMargineAppl" %></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemname" maxlength="64" size="50" class="text" id="[on,off,off,off][상품명]" value="<%= Replace(oitem.FOneItem.Fitemname,"""","&quot;") %>">&nbsp;
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">영문상품명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemnameEng" maxlength="64" size="60" class="text_ro" readonly id="[off,off,off,off][영문상품명]" value="<%= Replace(oitem.FOneItem.FitemnameEng,"""","&quot;") %>">&nbsp;
		<input type="button" value="다국어 정보 <%=chkIIF(oitem.FOneItem.FitemnameEng="" or isnull(oitem.FOneItem.FitemnameEng),"등록","수정")%>" class="button" onclick="popMultiLangEdit(<%= oitem.FOneItem.Fitemid %>)" />
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품카피 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="designercomment" size="60" maxlength="128" class="text" id="[off,off,off,off][상품카피]" value="<%= Replace(oitem.FOneItem.Fdesignercomment,"""","&quot;") %>">
	</td>
</tr>
</table>

<!-- 2.구분 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left" >
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>2.구분</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="재고/매출 등의 관리 카테고리" style="cursor:help;">관리 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td><%=getCategoryInfo(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"><input type="button" value="추가" class="button" onClick="popCateSelect('<%=oitem.FOneItem.Fitemid%>')"></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="프론트에 진열될 카테고리" style="cursor:help;">전시 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getDispCategory(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품구분 :</td>
	<td bgcolor="#FFFFFF" >
		<label><input type="radio" name="itemdiv" value="01" <%=chkIIF(oitem.FOneItem.Fitemdiv="01","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">일반상품</label>
		<br>
		<label><input type="radio" name="itemdiv" value="<%= oitem.FOneItem.Fitemdiv %>" <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","checked","")%> onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">주문 제작상품</label>
		<input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitem.FOneItem.Fitemdiv="06","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">주문제작 문구 필요<font color=red>(주문시 이니셜등 제작문구가 필요한경우 체크)</font>
		<br>
		<label><input type="radio" name="itemdiv" value="08" <%=chkIIF(oitem.FOneItem.Fitemdiv="08","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">티켓상품</label>
		<label><input type="radio" name="itemdiv" value="09" <%=chkIIF(oitem.FOneItem.Fitemdiv="09","checked","")%> >Present상품</label>
		<label><input type="radio" name="itemdiv" value="11" <%=chkIIF(oitem.FOneItem.Fitemdiv="11","checked","")%> >상품권상품</label>
		<label><input type="radio" name="itemdiv" value="18" <%=chkIIF(oitem.FOneItem.Fitemdiv="18","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">여행상품</label>

		<% if oitem.FOneItem.Fitemdiv ="07" then %> <!-- 2014년이전 단독구매 상품 > reserveItemTp=1 / 현재는 구매제한(회원당 구매 제한) -->
			<label><input type="radio" name="itemdiv" value="07" <%=chkIIF(oitem.FOneItem.Fitemdiv="07","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">구매제한상품</label>
		<% end if %>
		<% if oitem.FOneItem.Fitemdiv ="82" then %>
			<label><input type="radio" name="itemdiv" value="82" <%=chkIIF(oitem.FOneItem.Fitemdiv="82","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">마일리지샵 상품</label>
		<% end if %>

		<label><input type="radio" name="itemdiv" value="75" <%=chkIIF(oitem.FOneItem.Fitemdiv="75","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">정기구독상품</label>

		<% If rentalItemFlag Then %>
			<label><input type="radio" name="itemdiv" value="30" <%=chkIIF(oitem.FOneItem.Fitemdiv="30","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">렌탈상품</label>
		<% End If %>
		<label><input type="radio" name="itemdiv" value="23" <%=chkIIF(oitem.FOneItem.Fitemdiv="23","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">B2B상품</label>
		<label><input type="radio" name="itemdiv" value="17" <%=chkIIF(oitem.FOneItem.Fitemdiv="17","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">마케팅전용상품</label>
	</td>
	<td bgcolor="#FFFFFF">
	    <div id="lyRequre" style="<%=chkIIF((oitem.FOneItem.Fitemdiv ="06") or (oitem.FOneItem.Fitemdiv ="16"),"","display:none;")%>padding-left:22px;">
		예상제작소요일 <input type="text" name="requireMakeDay" value="<%=oitem.FOneItem.FrequireMakeDay%>" size="2" class="text" id="[off,on,off,off][예상제작소요일]">일
		<font color="red">(상품발송전 상품제작 기간)</font>
		</div>
	</td>
</tr>
<% if (oitem.FOneItem.IsReserveOnlyItem) then %>
<!-- 설정은 시스템팀 only 2012/03/26 추가-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">단독(예약)구매 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
	    <label><input type="radio" name="reserveItemTp" value="0" <%=chkIIF(oitem.FOneItem.FreserveItemTp="0" And oitem.FOneItem.Fitemdiv <>"30","checked","")%>>일반</label>
		<label><input type="radio" name="reserveItemTp" value="1" <%=chkIIF(oitem.FOneItem.FreserveItemTp="1" or oitem.FOneItem.Fitemdiv="30","checked","")%>>단독(예약)구매상품</label>
	</td>
</tr>
<% end if %>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">텐바이텐 독점구분 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="tenOnlyYn" value="Y" <%=chkIIF(oitem.FOneItem.FtenOnlyYn="Y","checked","")%>>독점상품</label>
		<label><input type="radio" name="tenOnlyYn" value="N" <%=chkIIF(oitem.FOneItem.FtenOnlyYn="N","checked","")%>>일반상품</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">선착순 결제 상품 :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="2">
	<input type="hidden" name="availPayType" value="<%= oitem.FOneItem.FavailPayType %>">
	<% if (oitem.FOneItem.FavailPayType = "9") then %>
		선착순
	<% elseif (oitem.FOneItem.FavailPayType = "8") then %>
		저스트원데이
	<% elseif (oitem.FOneItem.FavailPayType = "0") then %>
		일반
	<% else %>
		<%= oitem.FOneItem.FavailPayType %>
	<% end if %>
	</td>
</tr>
</table>

<!-- 3.가격정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>3.가격정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">가격설정 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
		<tr align="center">
			<td height="25" width="90" bgcolor="#DDDDFF">선택</td>
			<td width="100" bgcolor="#DDDDFF">소비자가</td>
			<td width="100" bgcolor="#DDDDFF">공급가</td>
			<td width="100" bgcolor="#DDDDFF">마진</td>
			<td bgcolor="#DDDDFF">&nbsp;</td>
		</tr>
		<tr>
			<td height="25" bgcolor="#FFFFFF"><label><input type="radio" name="sailyn" onClick="TnCheckSailYN(itemreg)" value="N" <% if oitem.FOneItem.Fsailyn = "N" then response.write "checked" %>> 정상가격</label></td>
			<td bgcolor="#FFFFFF" align="center">
			<% if oitem.FOneItem.Fsailyn = "N" then %>
				<input type="text" name="sellcash" maxlength="16" size="8" class="text" id="[on,on,off,off][소비자가]" value="<%= oitem.FOneItem.Fsellcash %>" onkeyup="CalcuAuto(itemreg);">원
			<% else %>
				<input type="text" name="sellcash" maxlength="16" size="8" class="text" id="[on,on,off,off][소비자가]" value="<%= oitem.FOneItem.Forgprice %>" onkeyup="CalcuAuto(itemreg);">원
			<% end if %>
			</td>
			<td bgcolor="#FFFFFF" align="center">
			<% if oitem.FOneItem.Fsailyn = "N" then %>
				<input type="text" name="buycash" maxlength="16" size="8" class="text" id="[on,on,off,off][공급가]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fbuycash %>">원
			<% else %>
				<input type="text" name="buycash" maxlength="16" size="8" class="text" id="[on,on,off,off][공급가]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Forgsuplycash %>">원
			<% end if %>
			</td>
			<% if oitem.FOneItem.Fsailyn = "N" then %>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][마진]" value="<%= margine %>">%
			</td>
			<% else %>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="margin" maxlength="32" size="5" class="text" id="[on,off,off,off][마진]" value="<%= orgmargine %>">%
			</td>
			<% end if %>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" value="공급가 자동계산" class="button" onclick="CalcuAuto(itemreg);">
			</td>
		</tr>
		<tr>
			<td height="25" bgcolor="#FFFFFF"><label><input type="radio" name="sailyn" onClick="TnCheckSailYN(itemreg)" value="Y" <% if oitem.FOneItem.Fsailyn = "Y" then response.write "checked" %>> 할인가격</label></td>
			<input type="hidden" name="sailpricevat">
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailprice" maxlength="16" size="8" class="text" id="[on,on,off,off][할인소비자가]" value="<%= oitem.FOneItem.Fsailprice %>"  onkeyup="CalcuAuto(itemreg);">원
			</td>
			<input type="hidden" name="sailsuplycashvat">
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailsuplycash" maxlength="16" size="8" class="text" id="[on,on,off,off][할인공급가]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fsailsuplycash %>">원
			</td>
			<td bgcolor="#FFFFFF" align="center">
				<input type="text" name="sailmargin" maxlength="32" size="5" class="text" id="[on,off,off,off][할인마진]" value="<%= sailmargine %>">%
			</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" value="공급가 자동계산" class="button" onclick="CalcuAuto(itemreg);">
				<%
					dim itemSalePer : itemSalePer=0
					if oitem.FOneItem.Fsailyn="Y" then
						itemSalePer = oitem.FOneItem.Forgprice - oitem.FOneItem.Fsailprice
						itemSalePer = itemSalePer/oitem.FOneItem.Forgprice*100
					end if
				%>
				<span id="lyrPct"><% if itemSalePer>0 then %>할인율: <font color="#EE0000"><strong><%=formatNumber(itemSalePer,1)%>%</strong></font><% end if %></span>
			</td>
		</tr>
		</table>
		<br>
		- 공급가는 <b>부가세 포함가</b>입니다.<br>
		- 소비자가(할인가)와 마진(할인마진)을 입력하고 [공급가자동계산] 버튼을 누르면 공급가와 마일리지가 자동계산됩니다.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">마일리지 :</td>
	<td width="35%" bgcolor="#FFFFFF"><input type="text" name="mileage" maxlength="32" size="10" class="text" id="[on,on,off,off][마일리지]" value="<%= oitem.FOneItem.Fmileage %>">point</td>
	<td width="15%" bgcolor="#DDDDFF">과세, 면세 여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="vatinclude" value="Y" <% if oitem.FOneItem.Fvatinclude = "Y" then response.write "checked" %>>과세</label>
		<label><input type="radio" name="vatinclude" value="N" <% if oitem.FOneItem.Fvatinclude = "N" then response.write "checked" %>>면세</label>
	</td>
</tr>
</table>

<!-- 4.관리정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>4.관리정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%= oitem.FOneItem.Fitemid %>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="미리보기" class="button" onclick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>');">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="상품 상세 속성" style="cursor:help;">상품속성 :</td>
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3"></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">업체상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="upchemanagecode" class="text" id="[off,off,off,off][업체상품코드]" value="<%= oitem.FOneItem.Fupchemanagecode %>" size="20" maxlength="32">
		(업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">ISBN :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		ISBN 13 <input type="text" name="isbn13" class="text" value="<%= oitem.FOneItem.Fisbn13 %>" size="13" maxlength="13">
		/ 부가기호 <input type="text" name="isbn_sub" class="text" value="<%= oitem.FOneItem.FisbnSub %>" size="5" maxlength="5"><br />
		ISBN 10 <input type="text" name="isbn10" class="text" value="<%= oitem.FOneItem.Fisbn10 %>" size="10" maxlength="10"> (Optional)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">연관상품등록 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="relateItems" value="<%=strItemRelation%>" size="40" class="text" id="[off,off,off,off][연관상품]">
	    (연관상품은 최대 6개까지 등록가능, 상품번호를 콤마(,)로 구분하여 입력)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">판매여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="sellyn" value="Y" <% if oitem.FOneItem.Fsellyn = "Y" then response.write "checked" %>>판매함</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="S" <% if oitem.FOneItem.Fsellyn = "S" then response.write "checked" %>>일시품절</label>&nbsp;&nbsp;
		<label><input type="radio" name="sellyn" value="N" <% if oitem.FOneItem.Fsellyn = "N" then response.write "checked" %>>판매안함</label> 
	<%IF (oitem.FOneItem.Fsellreservedate)<> "" THEN %><font color="blue">[오픈예약: <%=oitem.FOneItem.Fsellreservedate%>]</font><%END IF%>
	</td>
	<td width="15%" bgcolor="#DDDDFF">사용여부 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<label><input type="radio" name="isusing" value="Y" onclick="TnChkIsUsing(this.form)" <%=chkIIF(oitem.FOneItem.Fisusing="Y","checked","")%>>사용함</label>&nbsp;&nbsp;
		<label><input type="radio" name="isusing" value="N" onclick="TnChkIsUsing(this.form)" <%=chkIIF(oitem.FOneItem.Fisusing="N","checked","")%>>사용안함</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">제품등록일 :</td>
	<td bgcolor="#FFFFFF" colspan="3"><%= oitem.FOneItem.FRegDate %></td>
</tr>
</table>

<!-- 5.기본정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>5.기본정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">제조사 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="makername" maxlength="32" size="25" class="text" id="[on,off,off,off][제조사]" value="<%= oitem.FOneItem.Fmakername %>">&nbsp;(제조업체명)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">원산지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		 <p> 
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="0" <%if isNull(oitem.FOneItem.Fsourcekind) or oitem.FOneItem.Fsourcekind="0" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 식품 외</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="1" <%if oitem.FOneItem.Fsourcekind="1" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 농축산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="2" <%if oitem.FOneItem.Fsourcekind="2" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 수산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="3" <%if oitem.FOneItem.Fsourcekind="3" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 축산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="4" <%if oitem.FOneItem.Fsourcekind="4" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 농수산가공품</span>
	  </p>
	  <p><input type="text" name="sourcearea" maxlength="64" size="64" class="text" id="[on,off,off,off][원산지]"  value="<%= oitem.FOneItem.Fsourcearea %>"/></p>
	  <div id="dvArea0" style="display:<%if isNull(oitem.FOneItem.Fsourcekind) or oitem.FOneItem.Fsourcekind="0" then%>block<%else%>none<%end if%>;">
	  <p><strong>ex: 한국, 중국, 중국OEM, 일본 등 </strong></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea1" style="display:<%if oitem.FOneItem.Fsourcekind ="1" then%>block<%else%>none<%end if%>;">
	  <p><strong>국내산 :</strong> 국산, 국내산 또는 시·도명, 시·군명(대한민국, 한국X)  <span style="margin-right:10px;">ex. 쌀(국산)</span></BR>
	   <strong>수입산 :</strong> 통관시의 수입국가명 <span style="margin-right:10px;">ex. 곶감(중국산)</span></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea2" style="display:<%if oitem.FOneItem.Fsourcekind ="2" then%>block<%else%>none<%end if%>;">
	  <p><strong>국내산 :</strong> 국산,국내산 또는 연근해산(양식 수산물은 시·군명 가능)   <span style="margin-right:10px;">ex. 갈치(국산), 오징어(연근해산)</span> </BR>
	  	<strong>원양산 :</strong> 원양산 또는 원양산(해역명)   <span style="margin-right:10px;">ex. 참치[원양산(대서양)]</span> </BR>
	    <strong>수입산 :</strong> 통관시의 수입국가명 <span style="margin-right:10px;">ex. 농어(중국산)</span></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea3" style="display:<%if oitem.FOneItem.Fsourcekind ="3" then%>block<%else%>none<%end if%>;">
	  <p>소고기의 경우 식육의 종류(한우/육우/젖소구분) 및 원산지   <span style="margin-right:10px;">ex. 쇠고기(횡성산 한우), 쇠고기(호주산)</span></BR>
	  - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea4" style="display:<%if oitem.FOneItem.Fsourcekind ="4" then%>block<%else%>none<%end if%>;">
	  <p><strong>98%이상 원료가 있는 경우:</strong>  한가지 원료만 표시 가능    <span style="margin-right:10px;">ex. 쇠고기(미국산)</span> </BR>
	  	<strong>복합 원료를 사용한 경우:</strong> 혼합비율이 높은 순으로 2개 국가   <span style="margin-right:10px;">ex. 고추장[밀가루(미국산),고춧가루(국내산)]</span></BR>
	  - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div> 
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품무게 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" class="text" id="[on,off,off,off][상품무게]" style="text-align:right" value="<%= oitem.FOneItem.FitemWeight %>">g &nbsp;(그램단위로 입력, ex:1.5kg→ 1500) / 해외배송시 배송비 산출을 위한 것이므로 정확히 입력.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">검색키워드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="keywords" maxlength="250" size="50" class="text" id="[on,off,off,off][검색키워드]" value="<%= oitem.FOneItem.Fkeywords %>">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)
	</td>
</tr>
</table>

<!-- 5-1.품목상세정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- 품목상세정보 </strong> &nbsp;<font color=gray>상품정보제공고시 관련 법안 추진에 따라 아래 내용을 정확히 입력해주시기 바랍니다.</font></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목선택 :</td>
	<td bgcolor="#FFFFFF">
		<% DrawInfoDiv "infoDiv", oitem.FOneItem.FinfoDiv, " onchange='chgInfoDiv(this.value);'" %>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목내용 :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="") then
			Server.Execute("act_itemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">상품상세페이지에 내용이 포함 되어있더라도 정확히 입력바랍니다. 부정확하거나 잘못된 정보 입력시, 그에 대한 책임을 물을 수도 있습니다.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:플라스틱,비즈,금,...)
	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
	</td>
</tr>
</table>
<!-- 5-2.안전인증정보 -->
<%
dim arrAuth, r, real_safetydiv, real_safetynum, safetyDivList
arrAuth = oitem.FAuthInfo
if isArray(arrAuth) THEN
	For r =0 To UBound(arrAuth,2)
		real_safetydiv = real_safetydiv & arrAuth(0,r)
		if r <> UBound(arrAuth,2) then real_safetydiv = real_safetydiv & "," end if
		
		real_safetynum = real_safetynum & arrAuth(1,r)
		if r <> UBound(arrAuth,2) then real_safetynum = real_safetynum & "," end if
		
		safetyDivList = safetyDivList & "<p class='tPad05' id='l"&arrAuth(0,r)&"'>"
		safetyDivList = safetyDivList & "- "&fnSafetyDivCodeName(arrAuth(0,r))&"("&CHKIIF(arrAuth(1,r)="x","인증번호 없음",arrAuth(1,r))&")"
		safetyDivList = safetyDivList & " <input type='button' value='삭제' class='btn3 btnIntb' onClick='jsSafetyDivListDel("&arrAuth(0,r)&");'>"
		safetyDivList = safetyDivList & "</p>"
	Next
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-left:5px;"><strong>- 안전인증정보</strong></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">
		안전인증대상 :
		<input type="button" value="안전인증 필수 품목 확인" onclick="jsSafetyPopup();" class="button" />
	</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
		<tr align="left" height="30">
			<td bgcolor="#FFFFFF">
				<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitem.FOneItem.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 대상</label>
				<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitem.FOneItem.FsafetyYn="N","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 대상아님</label>
				<label><input type="radio" name="safetyYn" value="I" <%=chkIIF(oitem.FOneItem.FsafetyYn="I","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 상품설명에 표기</label>
				<label><input type="radio" name="safetyYn" value="S" <%=chkIIF(oitem.FOneItem.FsafetyYn="S","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 안전기준준수</label>
				<input type="hidden" name="auth_go_catecode" id="auth_go_catecode" value="">
				<input type="hidden" name="real_safetydiv" id="real_safetydiv" value="<%=real_safetydiv%>">
				<input type="hidden" name="real_safetynum" id="real_safetynum" value="<%=real_safetynum%>">
				<input type="hidden" name="real_safetyidx" id="real_safetyidx" value="">
				<input type="hidden" name="real_safetynum_delete" id="real_safetynum_delete" value="">
				<input type="hidden" name="real_safetydiv_delete" id="real_safetydiv_delete" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxSafetyDivCode "safetyDiv", "", oitem.FOneItem.FsafetyYn, "" %>
				인증번호 <input type="text" name="safetyNum" id="[off,off,off,off][안전인증 인증번호]" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" value="" /><%'=oitem.FOneItem.FsafetyNum%>
				<input type="button" id="safetybtn" value="추   가" onclick="jsSafetyAuth();" class="button">
				<input type="hidden" name="issafetyauth" id="issafetyauth" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<div id="safetyDivList">
					<%=safetyDivList%>
				</div>
				<div id="safetyYnI" style="display:none;">
					<font color="blue">상품 설명에 표기(표기대상 상품인경우 상품 상세 페이지에 인증번호와 모델명, KC 마크를 꼭 표기해주세요.)</font>
				</div>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td bgcolor="#FFFFFF" colspan=2>
		* 인증정보를 입력 안 하거나, 잘못된 인증정보를 입력한 경우 발견 <strong><font color='red'>즉시 판매정지 또는 삭제</font></strong> 됩니다.<br>
		* <strong><font color='red'>안전기준준수</font></strong> 대상일경우 인증번호가 없으며, KC마크를 표시하지 않아야 됩니다.<br>
		* 입력한 인증정보는 제품안전정보센터에서 제공된 정보를 기준으로 조회되며, <strong><font color='red'>검증되지 않은 정보는 등록이 불가</font></strong>능합니다.<br>
		* 정상적인 인증정보를 입력했음에도 불구하고 등록이 안될경우에 "상품설명에 표기"로 설정이 가능하며, 상품 상세 페이지에 모델명과 표기대상 상품인경우 인증번호,KC마크를 표기해야 합니다.<br>
		* 안전인증정보 관련 문의는 홈페이지(<u><a href="http://safetykorea.kr" target="_blank">http://safetykorea.kr</a></u>)로 확인해 주시기 바랍니다.
	</td>
</tr>
</table>

<!-- 6.배송정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>6.배송정보</strong>
    </td>
    <td align="right">
    	<input type="button" class="button" value="계약조건으로 세팅" onclick="TnAutoChkDeliver()">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">매입특정구분 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <% IF chkMWAuth THEN %>
		<label><input type="radio" name="mwdiv" value="M" onclick="TnCheckUpcheYN(this.form);" <% if oitem.FOneItem.Fmwdiv = "M" then response.write "checked" %>>매입</label>
		<label><input type="radio" name="mwdiv" value="W" onclick="TnCheckUpcheYN(this.form);" <% if oitem.FOneItem.Fmwdiv = "W" then response.write "checked" %>>특정</label>
		<label><input type="radio" name="mwdiv" value="U" onclick="TnCheckUpcheYN(this.form);" <% if oitem.FOneItem.Fmwdiv = "U" then response.write "checked" %>>업체배송</label>
		&nbsp;&nbsp; - 매입특정구분에 따라 배송구분이 달라집니다. 배송구분을 확인해주세요.
		<%ELSE%> 
		<%= fnColor(oitem.FOneItem.Fmwdiv,"mw") %>
		<input type="hidden" name="mwdiv" value="<%=oitem.FOneItem.Fmwdiv%>">
		<%END IF%>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송구분 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverytype" value="1" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "1" then response.write "checked" %>>텐바이텐배송</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="2" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "2" then response.write "checked" %>>업체(무료)배송</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="4" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "4" then response.write "checked" %>>텐바이텐무료배송</label>&nbsp;
		<label><input type="radio" name="deliverytype" value="9" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "9" then response.write "checked" %>>업체조건배송(개별 배송비부과)</label>
		<label><input type="radio" name="deliverytype" value="7" onclick="TnCheckUpcheDeliverYN(this.form);" <% if oitem.FOneItem.Fdeliverytype = "7" then response.write "checked" %>>업체착불배송</label>
		<% if oitem.FOneItem.Fdeliverytype = "6" then %>
		<label><input type="radio" name="deliverytype" value="6" onclick="TnCheckUpcheDeliverYN(this.form);" checked><font color="darkred">현장수령</font></label>
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송방법 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverfixday" value="" <%=chkIIF(Trim(oitem.FOneItem.Fdeliverfixday)="" or IsNull(oitem.FOneItem.Fdeliverfixday),"checked","")%> onclick="TnCheckFixday(this.form)">택배(일반)</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="X" <%=chkIIF(oitem.FOneItem.Fdeliverfixday="X","checked","")%> onclick="TnCheckFixday(this.form)">화물</label>&nbsp;
		<label><input type="radio" name="deliverfixday" value="C" <%=chkIIF(oitem.FOneItem.Fdeliverfixday="C","checked","")%> onclick="TnCheckFixday(this.form)">플라워지정일</label>
		<label><input type="radio" name="deliverfixday" value="G" <%=chkIIF(oitem.FOneItem.Fdeliverfixday="G","checked","")%> onclick="TnCheckFixday(this.form)">해외직구</label>
		<span id="lyrFreightRng" style="display:<%=chkIIF(oitem.FOneItem.Fdeliverfixday="X","","none")%>;">
			<br />&nbsp;
			반품/교환 시 화물배송 비용(편도) :
			최소 <input type="text" name="freight_min" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_min%>" style="text-align:right;">원 ~
			최대 <input type="text" name="freight_max" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_max%>" style="text-align:right;">원
		</span>
		<br>&nbsp;<font color="red">(플라워 상품인 경우만 수도권배송, 서울배송, 플라워지정일 옵션이 사용가능합니다.)</font>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">배송지역 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="deliverarea" value="" <%=chkIIF(Trim(oitem.FOneItem.Fdeliverarea)="" or IsNull(oitem.FOneItem.Fdeliverarea),"checked","")%>>전국배송</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="C" <%=chkIIF(oitem.FOneItem.Fdeliverarea="C","checked","")%> >수도권배송</label>&nbsp;
		<label><input type="radio" name="deliverarea" value="S" <%=chkIIF(oitem.FOneItem.Fdeliverarea="S","checked","")%> >서울배송</label>
		<label><input type="checkbox" name="deliverOverseas" value="Y" <% if oitem.FOneItem.FdeliverOverseas="Y" then response.write "checked" %> title="해외배송은 상품무게가 입력이 돼야 완료됩니다.">해외배송</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">포장가능여부 :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<%= oitem.FOneItem.Fpojangok %> <!-- 읽기전용 포장 여부 수정은 다른곳에서 popup 으로. -->
	</td>
</tr>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">재입고예정일 :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<input type="text" name="reipgodate" class="text" id="[off,off,off,off][재입고예정일]" size="10" value="<%= oitem.FOneItem.FreipgoDate %>" maxlength="10">
		<a href="javascript:calendarOpen(itemreg.reipgodate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		<a href="javascript:ClearVal(itemreg.reipgodate);"><img src="/images/icon_delete2.gif" width="20" border="0" align="absmiddle"></a>
	</td>
</tr>
</table>

<!-- 7.옵션정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>7.옵션정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">옵션구분 :</td>
	<input type="hidden" name="optioncnt" value="<%= oitem.FOneItem.Foptioncnt %>">
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
	<% if oitem.FOneItem.Foptioncnt < 1 then %>
		옵션사용안함
	<% else %>
		옵션사용중
	<% end if %>
	</td>
</tr>
<tr align="left">
	<td width="15%" bgcolor="#DDDDFF">옵션설정 :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3">
		- 옵션정보는 옵션창에서 수정가능합니다.<br>
		- 옵션은 추가는 가능하지만 삭제는 불가능합니다. 정확히 입력하세요.<br>
		- 한정수량은 옵션이 있을 경우, 옵션창에서 수정이 가능합니다.<br>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" rowspan="2">색상선택 :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
	  - 상품 색상선택은 [상품 컬러 관리]에서 하실 수 있습니다.
	</td>
</tr>
</table>

<!-- 8.한정정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>8.한정정보</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td width="15%" bgcolor="#DDDDFF">한정판매구분 :</td>
	<td width="85%" colspan="3" bgcolor="#FFFFFF">
		<label><input type="radio" name="limityn" value="N" onClick="TnCheckLimitYN(itemreg)" <% if oitem.FOneItem.Flimityn = "N" then response.write "checked" %>>비한정판매</label>&nbsp;&nbsp;
		<label><input type="radio" name="limityn" value="Y" onClick="TnCheckLimitYN(itemreg)" <% if oitem.FOneItem.Flimityn = "Y" then response.write "checked" %>>한정판매</label>
	  <div id="dvDisp" style="display:none;" >
			&nbsp;-> 한정노출여부: 
			<input type="radio" name="limitdispyn" value="Y" <%IF oitem.FOneItem.Flimitdispyn="Y"  THEN%>checked<%END IF%>>노출 
			<input type="radio" name="limitdispyn" value="N" <%IF oitem.FOneItem.Flimitdispyn="N" or oitem.FOneItem.Flimitdispyn ="" THEN%>checked<%END IF%>>비노출
		</div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">한정수량 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="limitno" maxlength="32" size="8" readonly style="background-color:#E6E6E6;" class="text" id="[off,on,off,off][한정수량]" value="<%= oitem.FOneItem.Flimitno %>">
		-
		<input type="text" name="limitsold" maxlength="32" size="8" readonly style="background-color:#E6E6E6;" class="text" id="[off,on,off,off][한정판매]" value="<%= oitem.FOneItem.Flimitsold %>">
		=
		<input type="text" name="limitstock" maxlength="32" size="8" readonly style="background-color:#E6E6E6;" class="text" id="[off,on,off,off][한정재고]" value="<%= (oitem.FOneItem.Flimitno - oitem.FOneItem.Flimitsold) %>">(개)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">최소/최대 판매수 :</td>
	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		최소
		<input type="text" name="orderMinNum" maxlength="5" size="5" class="text" id="[off,on,off,off][최소판매수]" value="<%= oitem.FOneItem.ForderMinNum %>">
		/ 최대
		<input type="text" name="orderMaxNum" maxlength="5" size="5" class="text" id="[off,on,off,off][최대판매수]" value="<%= oitem.FOneItem.ForderMaxNum %>">
		(한 주문에 판매 제한 수)
	</td>
</tr>
</table>

<!-- 9.상품설명 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">
      <img src="/images/icon_arrow_down.gif" border="0" align="absbottom"> <strong>9.상품설명</strong>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 설명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="usinghtml" value="N" <%=chkIIF(oitem.FOneItem.Fusinghtml="N","checked","") %>>일반TEXT</label>
		<label><input type="radio" name="usinghtml" value="H" <%=chkIIF(oitem.FOneItem.Fusinghtml="H","checked","") %>>TEXT+HTML</label>
		<label><input type="radio" name="usinghtml" value="Y" <%=chkIIF(oitem.FOneItem.Fusinghtml="Y","checked","") %>>HTML사용</label>
		<br>
		<textarea name="itemcontent" rows="18" class="textarea" style="width:100%" id="[on,off,off,off][상품설명]"><%= oitem.FOneItem.Fitemcontent %></textarea>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">아이템 동영상 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="80" id="[off,off,off,off][아이템동영상]"><%=oitemvideo.FOneItem.FvideoFullUrl%></textarea>
	    <br>※ Youtube, Vimeo 동영상만 가능(Youtube : 소스코드값 입력, Vimeo : 임베딩값 입력)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문시 유의사항 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="ordercomment" rows="5" cols="90" class="textarea" id="[off,off,off,off][유의사항]"><%= oitem.FOneItem.Fordercomment %></textarea><br>
		<font color="red">특별한 배송기간이나 주문시 확인해야만 하는 사항</font>을 입력하시면 고객불만이나 환불을 줄일수 있습니다.
	</td>
</tr>
</table>

<!-- 10.이미지정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left" style="padding-bottom:5px;">
      <img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> <strong>10.이미지정보</strong>
		<br>- 텐바이텐에서 이미지를 등록할 경우에는 필수항목인 기본이미지만 입력하시기 바랍니다.
		<br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
		<br>&nbsp;&nbsp;(이미지사이즈나 <font color=red>가로세로폭의 사이즈</font>를 규격에 넘지 않게 등록해주세요. 규격초과시 등록이 되지 않습니다.)
		<br>- <font color=red>포토샾에서 Save For Web으로, Optimize체크, 압축율 80%이하</font>로 만드신 후 올려주시기 바랍니다.
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">기본이미지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fbasicimage <> "") then %>
		<div id="divimgbasic" style="display:block;">
		<img src="<%= oitem.FOneItem.Fbasicimage %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgbasic" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgbasic" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgbasic,40, 1000, 1000)"> (<font color=red>필수</font>,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="basic">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">아이콘이미지(자동생성) :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	<% if (oitem.FOneItem.Ficon1image <> "") then %>
		<img src="<%= oitem.FOneItem.Ficon1image %>" width="200" height="200">
	<% end if %>
	<% if (oitem.FOneItem.Ficon2image <> "") then %>
		<img src="<%= oitem.FOneItem.Ficon2image %>" >
	<% end if %>
	<% if (oitem.FOneItem.Flistimage120 <> "") then %>
		<img src="<%= oitem.FOneItem.Flistimage120 %>" width="120" height="120">
	<% end if %>
	<% if (oitem.FOneItem.Flistimage <> "") then %>
		<img src="<%= oitem.FOneItem.Flistimage %>" width="100" height="100">
	<% end if %>
	<% if (oitem.FOneItem.Fsmallimage <> "") then %>
		<img src="<%= oitem.FOneItem.Fsmallimage %>" width="50" height="50">
	<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">흰배경(누끼)이미지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmaskimage <> "") then %>
		<div id="divimgmask" style="display:block;">
		<img src="<%= oitem.FOneItem.Fmaskimage %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgmask" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgmask" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgmask,40, 1000, 1000)"> (<font color=red>필수</font>,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="mask">
	</td>
</tr>

<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">텐바이텐기본이미지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Ftentenimage <> "") then %>
		<div id="divimgtenten" style="display:block;">
		<img src="<%= oitem.FOneItem.Ftentenimage %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgtenten" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgtenten" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgtenten,40, 1000, 1000)"> (선택,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="tenten">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">텐바이텐기본썸네일이미지(자동생성) :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	<% if (oitem.FOneItem.Ftentenimage1000 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage1000 %>" width="400" height="400" title="1000*1000이미지">
	<% end if %>
	<% if (oitem.FOneItem.Ftentenimage600 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage600 %>" width="300" height="300" title="600*600이미지">
	<% end if %>
	<% if (oitem.FOneItem.Ftentenimage400 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage400 %>" width="200" height="200" title="400*400이미지">
	<% end if %>
	<% if (oitem.FOneItem.Ftentenimage200 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage200 %>" width="150" height="150" title="200*200이미지">
	<% end if %>
	<% if (oitem.FOneItem.Ftentenimage50 <> "") then %>
		<img src="<%= oitem.FOneItem.Ftentenimage50 %>" width="50" height="50" title="50*50이미지">
	<% end if %>
	</td>
</tr>

<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,1) <> "") then %>
		<div id="divimgadd1" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,1) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd1" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add1">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지2 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,2) <> "") then %>
		<div id="divimgadd2" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,2) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd2" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd2,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지3 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,3) <> "") then %>
		<div id="divimgadd3" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,3) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd3" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd3,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add3">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지4 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,4) <> "") then %>
		<div id="divimgadd4" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,4) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd4" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd4,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add4">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지5 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,5) <> "") then %>
		<div id="divimgadd5" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,5) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd5" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd5,40, 1000, 1000)"> (선택,1000X1000,jpg,gif)
		<input type="hidden" name="add5">
	</td>
</tr>
</table>
<%
	Dim cImg, k, vArr, j
	set cImg = new CItemAddImage
	cImg.FRectItemID = itemid
	vArr = cImg.GetAddImageListIMGTYPE1
%>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
	<% If isArray(vArr) Then
			If vArr(3,UBound(vArr,2)) > 0 Then
			For k = 1 To vArr(3,UBound(vArr,2))
	%>
			  <tr align="left">
			  	<td height="30" width="15%" bgcolor="#DDDDFF">상품설명이미지 #<%= (k) %> :</td>
			  	<td bgcolor="#FFFFFF">
		  		<%
		  		If cImg.IsImgExist(vArr,k) Then
		    		For j = 0 To UBound(vArr,2)
		    			If CStr(vArr(3,j)) = CStr(k) AND (vArr(4,j) <> "" and isNull(vArr(4,j)) = False) Then
							Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:block;""><img src=""" & webImgUrl & "/item/contentsimage/" & GetImageSubFolderByItemid(vArr(1,j)) & "/" & vArr(4,j) & """ height=""250""></div>"
							Exit For
		    			End If
		    		Next
				Else
					Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:none;""></div>"
				End If
				%>
			      <input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40, <%= (k-1) %>);" class="text" size="40">
			      <input type="button" value="#<%= (k) %> 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname<%=CHKIIF(vArr(3,UBound(vArr,2))=1,"","["&(k-1)&"]")%>,40, 1000, 1000, <%= (k-1) %>)"> (선택,800X1600, Max 800KB,jpg,gif)
			      <input type="hidden" name="addimggubun" value="<%= (k) %>">
			      <input type="hidden" name="addimgdel" value="">
			  	</td>
			  </tr>
	<%
			Next
			End IF
		Else
	%>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">PC상품설명이미지 #1 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname1" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40,0);" class="text" size="40">
				<input type="button" value="#1 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 800, 1600, 0)"> (선택,800X1600, Max 800KB,jpg,gif)
				<input type="hidden" name="addimggubun" value="1">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">PC상품설명이미지 #2 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname2" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40,1);" class="text" size="40">
				<input type="button" value="#2 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 800, 1600, 1)"> (선택,800X1600, Max 800KB,jpg,gif)
				<input type="hidden" name="addimggubun" value="2">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">PC상품설명이미지 #3 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname3" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40,2);" class="text" size="40">
				<input type="button" value="#3 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 800, 1600, 2)"> (선택,800X1600, Max 800KB,jpg,gif)
				<input type="hidden" name="addimggubun" value="3">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
	<%
	   End IF %>
</table>
<%	set cImg = nothing %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td bgcolor="#FFFFFF" height="30">
      <input type="button" value="PC상품설명이미지추가" class="button" onClick="InsertImageUp()">
      <font color="red">* 업로드가 된 이미지가 제대로 안나오면 새로고침(CTRL + F5(콘트롤 F5 버튼))을 해주세요.</font>
  	</td>
  </tr>
</table>

<%
	Dim cmImg, mk, vmArr, mj
	set cmImg = new CItemAddImage
	cmImg.FRectItemID = itemid
	vmArr = cmImg.GetAddImageListIMGTYPE2
%>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="MobileimgIn">
	<% If isArray(vmArr) Then
			If vmArr(3,UBound(vmArr,2)) > 0 Then
			For mk = 1 To vmArr(3,UBound(vmArr,2))
	%>

			  <tr align="left">
				<td height="30" width="15%" bgcolor="#DDDDFF">모바일상품상세이미지 #<%= (mk) %> :</td>
				<td bgcolor="#FFFFFF">
				<%
				If cmImg.IsImgExist(vmArr,mk) Then
					For mj = 0 To UBound(vmArr,2)
						If CStr(vmArr(3,mj)) = CStr(mk) AND (vmArr(4,mj) <> "" and isNull(vmArr(4,mj)) = False) Then
							Response.Write "<div id=""divaddmobileimgname"&(mk)&""" style=""display:block;""><img src=""" & webImgUrl & "/item/contentsimage/" & GetImageSubFolderByItemid(vmArr(1,mj)) & "/" & vmArr(4,mj) & """ height=""250""></div>"
							Exit For
						End If
					Next
				Else
					Response.Write "<div id=""divaddmobileimgname"&(mk)&""" style=""display:none;""></div>"
				End If
				%>
				  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40, <%= (mk-1) %>);" class="text" size="40">
				  <input type="button" value="#<%= (mk) %> 이미지지우기" class="button" onClick="ClearImage3(this.form.addmoblieimgname<%=CHKIIF(vmArr(3,UBound(vmArr,2))=1,"","["&(mk-1)&"]")%>,40, 640, 1200, <%= (mk-1) %>)"> (선택,400X800, Max 400KB,jpg,gif)
				  <input type="hidden" name="addmobileimggubun" value="<%= (mk) %>">
				  <input type="hidden" name="addmobileimgdel" value="">
				</td>
			  </tr>
	<%
			Next
			End IF
		Else
	%>
		  <tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">모바일상품상세이미지 #1 :</td>
			<td bgcolor="#FFFFFF">
			  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
			  <input type="button" value="#1 이미지지우기" class="button" onClick="ClearImage2(this.form.addmobileimgname[0],40, 640, 1200)"> (선택,640X1200, Max 400KB,jpg,gif)
				<input type="hidden" name="addmobileimggubun" value="1">
				<input type="hidden" name="addmobileimgdel" value="">
			</td>
		  </tr>
		  <tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">모바일상품상세이미지 #2 :</td>
			<td bgcolor="#FFFFFF">
			  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
			  <input type="button" value="#2 이미지지우기" class="button" onClick="ClearImage2(this.form.addmobileimgname[1],40, 640, 1200)"> (선택,640X1200, Max 400KB,jpg,gif)
				<input type="hidden" name="addmobileimggubun" value="2">
				<input type="hidden" name="addmobileimgdel" value="">
			</td>
		  </tr>
		  <tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">모바일상품상세이미지 #3 :</td>
			<td bgcolor="#FFFFFF">
			  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
			  <input type="button" value="#3 이미지지우기" class="button" onClick="ClearImage2(this.form.addmobileimgname[2],40, 640, 1200)"> (선택,640X1200, Max 400KB,jpg,gif)
				<input type="hidden" name="addmobileimggubun" value="3">
				<input type="hidden" name="addmobileimgdel" value="">
			</td>
		  </tr>
	<%
	   End IF %>
</table>
<%	set cmImg = nothing %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>※ 모바일 상품상세 이미지는 앞으로 이 영역으로 대체 됩니다. html은 사용하지 않을 예정이오니 이쪽으로 업로드 해주시기 바랍니다.<br>※ 모바일 상품상세에는 이미지를 잘라서 올려주시기 바랍니다.</strong></font>
 	</td>
 </tr>
  <tr align="left">
  	<td bgcolor="#FFFFFF">
      <input type="button" value="모바일상품상세이미지추가" class="button" onClick="InsertMobileImageUp()">
      <font color="red">* 업로드가 된 이미지가 제대로 안나오면 새로고침(CTRL + F5(콘트롤 F5 버튼))을 해주세요.</font>
  	</td>
  </tr>
</table>


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>※ 기존의 제품설명이미지는 사용하지 않고 상품설명이미지를 사용합니다. 기존에 등록된 제품설명이미지는 사용은 하되 추가 수정은 되지않고 삭제만 됩니다.</strong></font>
 	</td>
 </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">제품설명이미지 #1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmainimage <> "") then %>
		<div id="divimgmain" style="display:block;">
		<img src="<%=oitem.FOneItem.Fmainimage %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain" style="display:none;"></div>
	  <% end if %>
		<input type="button" value="이미지지우기" class="button" onClick="oldClearImage('main', 40, 800, 1600)"> (선택,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">제품설명이미지 #2:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmainimage2 <> "") then %>
		<div id="divimgmain2" style="display:block;">
		<img src="<%=oitem.FOneItem.Fmainimage2 %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain2" style="display:none;"></div>
	  <% end if %>
		<input type="button" value="이미지지우기" class="button" onClick="oldClearImage('main2', 40, 800, 1600)"> (선택,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">제품설명이미지 #3:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmainimage3 <> "") then %>
		<div id="divimgmain3" style="display:block;">
		<img src="<%=oitem.FOneItem.Fmainimage3 %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain3" style="display:none;"></div>
	  <% end if %>
		<input type="button" value="이미지지우기" class="button" onClick="oldClearImage('main3', 40, 800, 1600)"> (선택,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main3">
	</td>
</tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <!--<input type="button" value="저장하기" class="button" onClick="SubmitSave()">//-->
          <input type="button" value="취소하기" class="button" onClick="self.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
</form>

<% if application("Svr_Info")	= "Dev" then %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="600" height="600"></iframe>
<% else %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<p>
<script type="text/javascript">

// 매입특정구분 및 배송구분세팅
TnCheckUpcheYN(itemreg);
for (var i = 0; i < itemreg.elements.length; i++) {
    if (itemreg.elements[i].name == "deliverytype") {
        if (itemreg.elements[i].value == "<%= oitem.FOneItem.Fdeliverytype %>") {
            itemreg.elements[i].checked = true;
        }
    }
}

// 한정
TnSilentCheckLimitYN(itemreg);
// 세일
CheckSailEnDisabled(itemreg);

itemreg.designer.readOnly = true;

	// 안전인증체크. 전안법
	jsSafetyCheck('<%= oitem.FOneItem.FsafetyYn %>','');
</script>

<%
set oitem = Nothing
set oitemAddImage = Nothing
Set oitemvideo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->