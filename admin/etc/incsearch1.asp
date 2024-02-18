<%
Dim infoLoop, infoDivValue
Dim foreignMall : foreignMall = "N"
Dim hiddenMall : hiddenMall = "N"
Dim mLoop, vPurchasetype
vPurchasetype = request("purchasetype")
Select Case Request.ServerVariables("SCRIPT_NAME")
	Case "/admin/etc/my11st/my11stItem.asp"				foreignMall = "Y"
	Case "/admin/etc/zilingo/zilingoItem.asp"			foreignMall = "Y"
	Case "/admin/etc/shopify/shopifyItem.asp"			foreignMall = "Y"
	Case "/admin/etc/shopify/shopifyNewItem.asp"		foreignMall = "Y"
	Case "/admin/etc/nvstorefarmClass/nvClassItem.asp"	hiddenMall = "Y"
End Select
%>
<script language='javascript'>
function checkComp(comp){
	if ((comp.name=="bestOrd")||(comp.name=="bestOrdMall")){
		if ((comp.name=="bestOrd")&&(comp.checked)){
			comp.form.bestOrdMall.checked=false;
		}
		if ((comp.name=="bestOrdMall")&&(comp.checked)){
			comp.form.bestOrd.checked=false;
		}
	}
}
</script>
<label><input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>베스트순(10x10)</b></label>&nbsp;
<label><input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>베스트순(제휴몰)</b></label>&nbsp;
<br />
판매(10x10)
<select name="sellyn" class="select">
	<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >전체
	<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >판매
	<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >품절
</select>&nbsp;
한정
<select name="limityn" class="select">
	<option value="">전체
	<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >한정
	<option value="N" <%= CHkIIF(limityn="N","selected","") %> >일반
</select>&nbsp;
세일
<select name="sailyn" class="select">
	<option value="">전체
	<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >세일Y
	<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >세일N
</select>&nbsp;
<% If CMALLNAME = "11st1010" OR CMALLNAME = "lfmall" OR CMALLNAME = "cjmall" OR CMALLNAME = "interpark" OR CMALLNAME = "shintvshopping" OR CMALLNAME = "skstoa" OR CMALLNAME = "wetoo1300k" OR CMALLNAME = "lotteimall" OR CMALLNAME = "kakaostore" OR CMALLNAME = "boribori1010" OR CMALLNAME = "wconcept1010" OR CMALLNAME = "benepia1010" OR CMALLNAME = "qooi1010" OR CMALLNAME = "gmarket1010" OR CMALLNAME = "auction1010" Then %>
마진(<%= getOutmallstandardMargin %>%)
<% Else %>
마진(<%= CMAXMARGIN %>%)
<% End If %>
<select name="startMargin" class="select">
	<option value="">-선택-</option>
	<% For mLoop = 0 to 100 %>
	<option value="<%= mLoop %>" <%= CHkIIF(CStr(startMargin) = CStr(mLoop),"selected","") %>><%= mLoop %></option>
	<% Next %>
</select>
~
<select name="endMargin" class="select">
	<option value="">-선택-</option>
	<% For mLoop = 0 to 100 %>
	<option value="<%= mLoop %>" <%= CHkIIF(CStr(endMargin) = CStr(mLoop),"selected","") %>><%= mLoop %></option>
	<% Next %>
</select>&nbsp;
<% If hiddenMall = "N" Then %>
제작
<select name="isMadeHand" class="select">
	<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >전체</option>
	<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y</option>
	<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N</option>
	<option value="T" <%= CHkIIF(isMadeHand="T","selected","") %> >주문제작문구</option>
</select>&nbsp;
<% End If %>
옵션
<select name="isOption" class="select">
	<option value="" <%= CHkIIF(isOption="","selected","") %> >전체
	<option value="optAll" <%= CHkIIF(isOption="optAll","selected","") %> >옵션전체
	<option value="optaddpricey" <%= CHkIIF(isOption="optaddpricey","selected","") %> >추가금액Y
	<option value="optaddpricen" <%= CHkIIF(isOption="optaddpricen","selected","") %> >추가금액N
	<option value="optN" <%= CHkIIF(isOption="optN","selected","") %> >단품
</select>&nbsp;
품목
<select name="infodiv" class="select">
	<option value="" <%= CHkIIF(infoDiv="","selected","") %> >전체
	<option value="Y" <%= CHkIIF(infoDiv="Y","selected","") %> >입력
	<option value="N" <%= CHkIIF(infoDiv="N","selected","") %> >미입력
<%
	For infoLoop = 1 To 35
		If infoLoop < 10 Then
			infoDivValue = "0"&infoLoop
		Else
			infoDivValue = infoLoop
		End If
%>
	<option value="<%=infoDivValue%>" <%= CHkIIF(CStr(infodiv) = CStr(infoDivValue),"selected","") %> ><%= infoDivValue %>
<% Next %>
	<option value="47" <%= CHkIIF(CStr(infodiv) = "47","selected","") %> >47
	<option value="48" <%= CHkIIF(CStr(infodiv) = "48","selected","") %> >48
</select>&nbsp;
<% If foreignMall = "N" Then %>
	<% If hiddenMall = "N" Then %>
제외브랜드
<select name="notinmakerid" class="select">
	<option value="" <%= CHkIIF(notinmakerid="","selected","") %> >전체
	<option value="Y" <%= CHkIIF(notinmakerid="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(notinmakerid="N","selected","") %> >N
</select>&nbsp;
	<% End If %>
제외상품
<select name="notinitemid" class="select">
	<option value="" <%= CHkIIF(notinitemid="","selected","") %> >전체
	<option value="Y" <%= CHkIIF(notinitemid="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(notinitemid="N","selected","") %> >N
</select>&nbsp;
옵션추가금액
<select name="priceOption" class="select">
	<option value="" <%= CHkIIF(priceOption="","selected","") %> >전체
	<option value="Y" <%= CHkIIF(priceOption="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(priceOption="N","selected","") %> >N
</select>&nbsp;
	<% If hiddenMall = "N" Then %>
특가
<select name="isSpecialPrice" class="select">
    <option value="" <%= CHkIIF(isSpecialPrice="","selected","") %> >전체
    <option value="Y" <%= CHkIIF(isSpecialPrice="Y","selected","") %> >Y
</select>&nbsp;
	<% End If %>
<% End If %>
<br />
판매(제휴몰)
<select name="extsellyn" class="select">
	<option value="" <%= CHkIIF(extsellyn="","selected","") %> >전체
	<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >판매
	<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >품절
<% If cmallname ="gsshop" Then %>
	<option value="E" <%= CHkIIF(extsellyn="E","selected","") %> >대기
<% Else %>
	<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >종료
	<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >종료제외
	<% If cmallname ="interpark" Then %>
	<option value="SP" <%= CHkIIF(extsellyn="SP","selected","") %> >미입력
	<% End If %>
<% End If %>
</select>&nbsp;
전송제외상품
<select name="exctrans" class="select">
	<option value="" <%= CHkIIF(exctrans="","selected","") %> >전체</option>
	<option value="Y" <%= CHkIIF(exctrans="Y","selected","") %> >Y</option>
	<option value="N" <%= CHkIIF(exctrans="N","selected","") %> >N</option>
	<option value="F" <%= CHkIIF(exctrans="F","selected","") %> >N(FAIL)</option>
</select>&nbsp;
오류
<select name="failCntExists" class="select">
	<option value="" <%= CHkIIF(failCntExists="","selected","") %> >전체</option>
	<option value="Y" <%= CHkIIF(failCntExists="Y","selected","") %> >등록수정오류1회이상</option>
	<option value="N" <%= CHkIIF(failCntExists="N","selected","") %> >등록수정오류0회</option>
	<option value="5U" <%= CHkIIF(failCntExists="5U","selected","") %> >오류6회 이상</option>
	<option value="5D" <%= CHkIIF(failCntExists="5D","selected","") %> >오류5회 이하</option>
</select>&nbsp;
배송구분
<% drawBeadalDiv "deliverytype", deliverytype %>&nbsp;
거래구분
<% drawSelectBoxMWU "mwdiv", mwdiv %>
<% If CMALLNAME = "coupang" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	고시일치
	<select name="GosiEqual" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(GosiEqual="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(GosiEqual="N","selected","") %> >미매칭
	</select>&nbsp;
	출고지
	<select name="MatchShipping" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchShipping="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchShipping="N","selected","") %> >미매칭
	</select>&nbsp;
	옵션수차이 :
	<select name="regedOptOver" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(regedOptOver="Y","selected","") %> >초과
		<option value="N" <%= CHkIIF(regedOptOver="N","selected","") %> >미만
	</select>&nbsp;
	스케줄제외상품
	<select name="scheduleNotInItemid" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "ssg" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	적용마진
	<input type="text" name="setMargin" value="<%= setMargin%>" class="text" size="2" maxlength="2">
	스케줄제외상품
	<select name="scheduleNotInItemid" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "wetoo1300k" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	브랜드
	<select name="MatchBrand" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchBrand="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchBrand="N","selected","") %> >미매칭
	</select>&nbsp;
<% ElseIf CMALLNAME = "hmall1010" Then %>
	이미지
	<select name="MatchIMG" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchIMG="Y","selected","") %> >등록
		<option value="N" <%= CHkIIF(MatchIMG="N","selected","") %> >미등록
	</select>&nbsp;
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	스케줄제외상품
	<select name="scheduleNotInItemid" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
	적용마진
	<input type="text" name="setMargin" value="<%= setMargin%>" class="text" size="5" maxlength="5">
<% ElseIf CMALLNAME = "auction1010" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	스케줄제외상품
	<select name="scheduleNotInItemid" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "ezwel" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	상품분류
	<select name="MatchPrddiv" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >미매칭
	</select>&nbsp;
<% ElseIf CMALLNAME = "gmarket1010" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	스케줄제외상품
	<select name="scheduleNotInItemid" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
	G9등록여부
	<select name="MatchG9" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchG9="Y","selected","") %> >등록
		<option value="N" <%= CHkIIF(MatchG9="N","selected","") %> >미등록
	</select>&nbsp;
	금액
	<select name="sellpriceChk" class="select">
		<option value="">전체
		<option value="samman" <%= CHkIIF(sellpriceChk="samman","selected","") %> >3만원이상
	</select>&nbsp;
<% ElseIf CMALLNAME = "gsshop" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	상품분류
	<select name="MatchPrddiv" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >미매칭
	</select>&nbsp;
	스케줄제외상품
	<select name="scheduleNotInItemid" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "nvstorefarm" or CMALLNAME = "nvstoregift" or CMALLNAME = "Mylittlewhoopee" or CMALLNAME = "WMP" or CMALLNAME = "interpark" or CMALLNAME = "lfmall" or CMALLNAME = "11st1010" or CMALLNAME = "shintvshopping" or CMALLNAME = "skstoa" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	<% If CMALLNAME = "lfmall" Then %>
	품목분류
	<select name="MatchDiv" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchDiv="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchDiv="N","selected","") %> >미매칭
	</select>&nbsp;
	<% End If %>
	<% If CMALLNAME = "skstoa" Then %>
	적용마진
	<input type="text" name="setMargin" value="<%= setMargin%>" class="text" size="2" maxlength="2">
	<% End If %>
	스케줄제외상품
	<select name="scheduleNotInItemid" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "lotteon" or CMALLNAME = "lotteimall" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	스케줄제외상품
	<select name="scheduleNotInItemid" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% ElseIf CMALLNAME = "cjmall" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	상품분류
	<select name="MatchPrddiv" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >미매칭
	</select>&nbsp;
<% ElseIf CMALLNAME = "boribori1010" OR CMALLNAME = "wconcept1010" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
	브랜드
	<select name="MatchBrand" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchBrand="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchBrand="N","selected","") %> >미매칭
	</select>&nbsp;
<% ElseIf CMALLNAME = "qooi1010" OR CMALLNAME = "benepia1010" Then %>
	카테고리
	<select name="MatchCate" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
	</select>&nbsp;
<% ElseIf CMALLNAME = "sabangnet" Then %>
	스케줄제외상품
	<select name="scheduleNotInItemid" class="select">
		<option value="">전체
		<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
		<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
	</select>&nbsp;
<% End If %>
<br />
제휴사용(상품)
<select name="isextusing" class="select">
	<option value="">전체</option>
	<option value="Y" <%= CHkIIF(isextusing="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(isextusing="N","selected","") %> >N
</select>&nbsp;
제휴사용(브랜드)
<select name="cisextusing" class="select">
	<option value="">전체</option>
	<option value="Y" <%= CHkIIF(cisextusing="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(cisextusing="N","selected","") %> >N
</select>&nbsp;
3개월판매량
<select name="rctsellcnt" class="select">
	<option value="">전체</option>
	<option value="0" <%= CHkIIF(rctsellcnt="0","selected","") %> >0
	<option value="1" <%= CHkIIF(rctsellcnt="1","selected","") %> >1개이상
</select>&nbsp;
구매유형: 
<select name="purchasetype" class="select">
	<option value="">전체</option>
	<option value="1"	<%= CHkIIF(vPurchasetype="1","selected","") %> >일반유통
	<option value="3"	<%= CHkIIF(vPurchasetype="3","selected","") %> >PB
	<option value="4"	<%= CHkIIF(vPurchasetype="4","selected","") %> >사입
	<option value="5"	<%= CHkIIF(vPurchasetype="5","selected","") %> >ODM
	<option value="7"	<%= CHkIIF(vPurchasetype="7","selected","") %> >브랜드수입
	<option value="6"	<%= CHkIIF(vPurchasetype="6","selected","") %> >수입
	<option value="8"	<%= CHkIIF(vPurchasetype="8","selected","") %> >제작
	<option value="9"	<%= CHkIIF(vPurchasetype="9","selected","") %> >해외직구
	<option value="10"	<%= CHkIIF(vPurchasetype="10","selected","") %> >B2B
	<option value="356"	<%= CHkIIF(vPurchasetype="356","selected","") %> >PB/ODM/수입만
	<option value="101"	<%= CHkIIF(vPurchasetype="101","selected","") %> >일반유통 제외
	<option value="102"	<%= CHkIIF(vPurchasetype="102","selected","") %> >전략상품만
</select>&nbsp;