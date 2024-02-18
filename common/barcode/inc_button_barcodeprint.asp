<%
session.codePage = 65001
%>
<%
'###########################################################
' Description : 바코드 출력 프린트 설정 매뉴
' Hieditor : 2016.12.15 한용민 생성
'###########################################################
%>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<select name="papername">
			<option value="BQ" <% if (papername = "BQ") then %>selected<% end if %>>장바구니 쇼카드(QR코드)</option>
			<option value="T" <% if (papername = "T") then %>selected<% end if %>>쇼카드(물류코드)</option>
			<option value="G" <% if (papername = "G") then %>selected<% end if %>>쇼카드(범용바코드)</option>
			<option value="Q" <% if (papername = "Q") then %>selected<% end if %>>쇼카드(QR코드)</option>
			<option value="I" <% if (papername = "I") then %>selected<% end if %>>쇼카드(이미지)</option>
			<option value="M" <% if (papername = "M") then %>selected<% end if %>>쇼카드(물류코드,이미지)</option>
			<option value="N" <% if (papername = "N") then %>selected<% end if %>>쇼카드(범용바코드,이미지)</option>
			<option value="" <% if (papername = "") then %>selected<% end if %>>쇼카드(글씨만)</option>
		</select>
		<select name="itemcopydispyn">
			<option value="Y" <% if (itemcopydispyn = "Y") then %>selected<% end if %>>상품설명표시</option>
			<option value="N" <% if (itemcopydispyn = "N") then %>selected<% end if %>>상품설명표시안함</option>
		</select>
		<select name="itemoptionyn">
			<option value="Y" <% if (itemoptionyn = "Y") then %>selected<% end if %>>옵션표시</option>
			<option value="N" <% if (itemoptionyn = "N") then %>selected<% end if %>>옵션표시안함</option>
		</select>
		<input type="button" class="button" value="쇼카드출력" onClick="paperBarcodePrint('<%= onoffgubun %>');">
	</td>
	<td align="right">
		<% if printername = "TTP-243_45x22" then %>
			<input type="button" class="button" value="상품바코드출력(구)" onClick="BarcodePrint('T')">
			<input type="button" class="button" value="상품바코드출력" onClick="CssBarcodeprint('T')">
			<input type="button" class="button" value="상품범용바코드출력(구)" onClick="BarcodePrint('G')">
			<input type="button" class="button" value="상품범용바코드출력" onClick="CssBarcodeprint('G')">
		<% elseif printername = "TTP-243_35x15" then %>
			<input type="button" class="button" value="쥬얼리바코드출력(구)" onClick="jewellery_BarcodePrint('T');">
			<input type="button" class="button" value="쥬얼리바코드출력" onClick="jewelleryCssBarcodePrint('T');">
			<input type="button" class="button" value="쥬얼리범용바코드출력(구)" onClick="jewellery_BarcodePrint('G');">
			<input type="button" class="button" value="쥬얼리범용바코드출력" onClick="jewelleryCssBarcodePrint('G');">
		<% elseif printername = "TTP-243_45x45" then %>
			<input type="button" class="button" value="해외바코드출력" onClick="foreign_BarcodePrint('A');">
		<% elseif printername = "TTP-243_80x50" then %>
			<input type="button" class="button" value="인덱스출력" onClick="IndexCssBarcodePrint();">
			<input type="button" class="button" value="인덱스출력(구)" onClick="IndexBarcodePrint();">
			<input type="button" class="button" value="수기인덱스출력" onClick="IndexSudongBarcodePrint();">
		<% elseif printername = "TEC_B-FV4_45x22" then %>
			<input type="button" class="button" value="상품바코드출력(구)" onClick="BarcodePrint('T')">
			<input type="button" class="button" value="상품바코드출력" onClick="CssBarcodeprint('T')">
			<input type="button" class="button" value="상품범용바코드출력(구)" onClick="BarcodePrint('G')">
			<input type="button" class="button" value="상품범용바코드출력" onClick="CssBarcodeprint('G')">
		<% elseif printername = "TEC_B-FV4_35x15" then %>
			<input type="button" class="button" value="쥬얼리바코드출력(구)" onClick="jewellery_BarcodePrint('T');">
			<input type="button" class="button" value="쥬얼리바코드출력" onClick="jewelleryCssBarcodePrint('T');">
			<input type="button" class="button" value="쥬얼리범용바코드출력(구)" onClick="jewellery_BarcodePrint('G');">
			<input type="button" class="button" value="쥬얼리범용바코드출력" onClick="jewelleryCssBarcodePrint('G');">
		<% elseif printername = "TEC_B-FV4_45x45" then %>
			<input type="button" class="button" value="해외바코드출력" onClick="foreign_BarcodePrint('A');">
		<% elseif printername = "TEC_B-FV4_80x50" then %>
			<input type="button" class="button" value="인덱스출력" onClick="IndexCssBarcodePrint();">
			<input type="button" class="button" value="인덱스출력(구)" onClick="IndexBarcodePrint();">
			<input type="button" class="button" value="수기인덱스출력" onClick="IndexSudongBarcodePrint();">
		<% else %>
			<input type="button" class="button" value="상품바코드출력" onClick="CssFORMTECBarcodeprint('T')">
			<input type="button" class="button" value="상품범용바코드출력" onClick="CssFORMTECBarcodeprint('G')">
		<% end if %>
	</td>
</tr>
</table>