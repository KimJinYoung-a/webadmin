

<table width="100%" class="a">
<tr>
	<td valign="top">
		실제상품 등록 :
		<input class="button" type="button" id="btnRegSel" value="상품등록" onClick="ebaySelectRegProcess('<%= gubun %>');">&nbsp;&nbsp;
		<br><br>
		실제상품 수정 :

	</td>
	<td align="right" valign="top">
		<br><br>
		선택상품을
		<Select name="chgSellYn" class="select">
			<option value="N">판매중지</option>
			<option value="Y">판매</option>
		</Select>(으)로
		<input class="button" type="button" id="btnSellYn" value="변경" onClick="AuctionSellYnProcess(frmReg.chgSellYn.value);">
	</td>
</tr>
</table>