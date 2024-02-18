<%	'####### 전년실적 전용 inc 파일

	Dim vArr, vCateExist
	vCateExist = "x"	'### 전년실적 없을때 나오는  html 때문에 구분값 줌.
	vArr = olist.FLastYearArray
	
	'목표 Total
	Dim vTotalTarget1, vTotalTarget2, vTotalTarget3, vTotalTarget4, vTotalTarget5, vTotalTarget6, vTotalTarget7, vTotalTarget8, vTotalTarget9, vTotalTarget10, vTotalTarget11, vTotalTarget12
	Dim vProfit1, vProfit2, vProfit3, vProfit4, vProfit5, vProfit6, vProfit7, vProfit8, vProfit9, vProfit10, vProfit11, vProfit12
	Dim vAllTotTarget, vAllTotProfit, vAllTotItemCost, vAllTotMaechul
	
	'실적 Total
	Dim vItemcost1, vItemcost2, vItemcost3, vItemcost4, vItemcost5, vItemcost6, vItemcost7, vItemcost8, vItemcost9, vItemcost10, vItemcost11, vItemcost12
	Dim vMaechulProf1, vMaechulProf2, vMaechulProf3, vMaechulProf4, vMaechulProf5, vMaechulProf6, vMaechulProf7, vMaechulProf8, vMaechulProf9, vMaechulProf10, vMaechulProf11, vMaechulProf12
	
	'전년목표 Total
	Dim vLastTotalTarget1, vLastTotalTarget2, vLastTotalTarget3, vLastTotalTarget4, vLastTotalTarget5, vLastTotalTarget6, vLastTotalTarget7, vLastTotalTarget8, vLastTotalTarget9, vLastTotalTarget10, vLastTotalTarget11, vLastTotalTarget12
	Dim vLastProfit1, vLastProfit2, vLastProfit3, vLastProfit4, vLastProfit5, vLastProfit6, vLastProfit7, vLastProfit8, vLastProfit9, vLastProfit10, vLastProfit11, vLastProfit12
	Dim vLastAllTotTarget, vLastAllTotProfit, vLastAllTotItemCost, vLastAllTotMaechul
	
	'전년실적 Total
	Dim vLastItemcost1, vLastItemcost2, vLastItemcost3, vLastItemcost4, vLastItemcost5, vLastItemcost6, vLastItemcost7, vLastItemcost8, vLastItemcost9, vLastItemcost10, vLastItemcost11, vLastItemcost12
	Dim vLastMaechulProf1, vLastMaechulProf2, vLastMaechulProf3, vLastMaechulProf4, vLastMaechulProf5, vLastMaechulProf6, vLastMaechulProf7, vLastMaechulProf8, vLastMaechulProf9, vLastMaechulProf10, vLastMaechulProf11, vLastMaechulProf12
	
	
	If olist.FResultCount > 0 Then
		For i = 0 to olist.FResultCount -1
			vTotalTarget1 	= vTotalTarget1 + olist.FItemList(i).FTarget1
			vProfit1 		= vProfit1 + olist.FItemList(i).FProfit1
			vItemcost1 		= vItemcost1 + olist.FItemList(i).FItemcost1
			vMaechulProf1 	= vMaechulProf1 + olist.FItemList(i).FMaechulProfit1
			
			vTotalTarget2 	= vTotalTarget2 + olist.FItemList(i).FTarget2
			vProfit2 		= vProfit2 + olist.FItemList(i).FProfit2
			vItemcost2 		= vItemcost2 + olist.FItemList(i).FItemcost2
			vMaechulProf2 	= vMaechulProf2 + olist.FItemList(i).FMaechulProfit2
			
			vTotalTarget3 	= vTotalTarget3 + olist.FItemList(i).FTarget3
			vProfit3 		= vProfit3 + olist.FItemList(i).FProfit3
			vItemcost3 		= vItemcost3 + olist.FItemList(i).FItemcost3
			vMaechulProf3 	= vMaechulProf3 + olist.FItemList(i).FMaechulProfit3
			
			vTotalTarget4 	= vTotalTarget4 + olist.FItemList(i).FTarget4
			vProfit4 		= vProfit4 + olist.FItemList(i).FProfit4
			vItemcost4 		= vItemcost4 + olist.FItemList(i).FItemcost4
			vMaechulProf4 	= vMaechulProf4 + olist.FItemList(i).FMaechulProfit4
			
			vTotalTarget5 	= vTotalTarget5 + olist.FItemList(i).FTarget5
			vProfit5 		= vProfit5 + olist.FItemList(i).FProfit5
			vItemcost5 		= vItemcost5 + olist.FItemList(i).FItemcost5
			vMaechulProf5 	= vMaechulProf5 + olist.FItemList(i).FMaechulProfit5
			
			vTotalTarget6 	= vTotalTarget6 + olist.FItemList(i).FTarget6
			vProfit6 		= vProfit6 + olist.FItemList(i).FProfit6
			vItemcost6 		= vItemcost6 + olist.FItemList(i).FItemcost6
			vMaechulProf6 	= vMaechulProf6 + olist.FItemList(i).FMaechulProfit6
			
			vTotalTarget7 	= vTotalTarget7 + olist.FItemList(i).FTarget7
			vProfit7 		= vProfit7 + olist.FItemList(i).FProfit7
			vItemcost7 		= vItemcost7 + olist.FItemList(i).FItemcost7
			vMaechulProf7 	= vMaechulProf7 + olist.FItemList(i).FMaechulProfit7
			
			vTotalTarget8 	= vTotalTarget8 + olist.FItemList(i).FTarget8
			vProfit8 		= vProfit8 + olist.FItemList(i).FProfit8
			vItemcost8 		= vItemcost8 + olist.FItemList(i).FItemcost8
			vMaechulProf8 	= vMaechulProf8 + olist.FItemList(i).FMaechulProfit8
			
			vTotalTarget9 	= vTotalTarget9 + olist.FItemList(i).FTarget9
			vProfit9 		= vProfit9 + olist.FItemList(i).FProfit9
			vItemcost9 		= vItemcost9 + olist.FItemList(i).FItemcost9
			vMaechulProf9 	= vMaechulProf9 + olist.FItemList(i).FMaechulProfit9
			
			vTotalTarget10 	= vTotalTarget10 + olist.FItemList(i).FTarget10
			vProfit10 		= vProfit10 + olist.FItemList(i).FProfit10
			vItemcost10 	= vItemcost10 + olist.FItemList(i).FItemcost10
			vMaechulProf10 	= vMaechulProf10 + olist.FItemList(i).FMaechulProfit10
			
			vTotalTarget11 	= vTotalTarget11 + olist.FItemList(i).FTarget11
			vProfit11 		= vProfit11 + olist.FItemList(i).FProfit11
			vItemcost11 	= vItemcost11 + olist.FItemList(i).FItemcost11
			vMaechulProf11 	= vMaechulProf11 + olist.FItemList(i).FMaechulProfit11
			
			vTotalTarget12 	= vTotalTarget12 + olist.FItemList(i).FTarget12
			vProfit12 		= vProfit12 + olist.FItemList(i).FProfit12
			vItemcost12 	= vItemcost12 + olist.FItemList(i).FItemcost12
			vMaechulProf12 	= vMaechulProf12 + olist.FItemList(i).FMaechulProfit12
			
			If isArray(vArr) Then
				For k = 0 To UBound(vArr,2)
		
				If CStr(olist.FItemList(i).FCateCode) = CStr(vArr(0,k)) Then
					vLastTotalTarget1 	= vLastTotalTarget1 + vArr(4,k)
					vLastProfit1 		= vLastProfit1 + vArr(5,k)
					vLastItemcost1 		= vLastItemcost1 + vArr(6,k)
					vLastMaechulProf1 	= vLastMaechulProf1 + vArr(7,k)
					
					vLastTotalTarget2 	= vLastTotalTarget2 + vArr(8,k)
					vLastProfit2 		= vLastProfit2 + vArr(9,k)
					vLastItemcost2 		= vLastItemcost2 + vArr(10,k)
					vLastMaechulProf2 	= vLastMaechulProf2 + vArr(11,k)
					
					vLastTotalTarget3 	= vLastTotalTarget3 + vArr(12,k)
					vLastProfit3 		= vLastProfit3 + vArr(13,k)
					vLastItemcost3 		= vLastItemcost3 + vArr(14,k)
					vLastMaechulProf3 	= vLastMaechulProf3 + vArr(15,k)
					
					vLastTotalTarget4 	= vLastTotalTarget4 + vArr(16,k)
					vLastProfit4 		= vLastProfit4 + vArr(17,k)
					vLastItemcost4 		= vLastItemcost4 + vArr(18,k)
					vLastMaechulProf4 	= vLastMaechulProf4 + vArr(19,k)
					
					vLastTotalTarget5 	= vLastTotalTarget5 + vArr(20,k)
					vLastProfit5 		= vLastProfit5 + vArr(21,k)
					vLastItemcost5 		= vLastItemcost5 + vArr(22,k)
					vLastMaechulProf5 	= vLastMaechulProf5 + vArr(23,k)
					
					vLastTotalTarget6 	= vLastTotalTarget6 + vArr(24,k)
					vLastProfit6 		= vLastProfit6 + vArr(25,k)
					vLastItemcost6 		= vLastItemcost6 + vArr(26,k)
					vLastMaechulProf6 	= vLastMaechulProf6 + vArr(27,k)
					
					vLastTotalTarget7 	= vLastTotalTarget7 + vArr(28,k)
					vLastProfit7 		= vLastProfit7 + vArr(29,k)
					vLastItemcost7 		= vLastItemcost7 + vArr(30,k)
					vLastMaechulProf7 	= vLastMaechulProf7 + vArr(31,k)
					
					vLastTotalTarget8 	= vLastTotalTarget8 + vArr(32,k)
					vLastProfit8 		= vLastProfit8 + vArr(33,k)
					vLastItemcost8 		= vLastItemcost8 + vArr(34,k)
					vLastMaechulProf8 	= vLastMaechulProf8 + vArr(35,k)
					
					vLastTotalTarget9 	= vLastTotalTarget9 + vArr(36,k)
					vLastProfit9 		= vLastProfit9 + vArr(37,k)
					vLastItemcost9 		= vLastItemcost9 + vArr(38,k)
					vLastMaechulProf9 	= vLastMaechulProf9 + vArr(39,k)
					
					vLastTotalTarget10 	= vLastTotalTarget10 + vArr(40,k)
					vLastProfit10 		= vLastProfit10 + vArr(41,k)
					vLastItemcost10 	= vLastItemcost10 + vArr(42,k)
					vLastMaechulProf10 	= vLastMaechulProf10 + vArr(43,k)
					
					vLastTotalTarget11 	= vLastTotalTarget11 + vArr(44,k)
					vLastProfit11 		= vLastProfit11 + vArr(45,k)
					vLastItemcost11 	= vLastItemcost11 + vArr(46,k)
					vLastMaechulProf11 	= vLastMaechulProf11 + vArr(47,k)
					
					vLastTotalTarget12 	= vLastTotalTarget12 + vArr(48,k)
					vLastProfit12 		= vLastProfit12 + vArr(49,k)
					vLastItemcost12 	= vLastItemcost12 + vArr(50,k)
					vLastMaechulProf12 	= vLastMaechulProf12 + vArr(51,k)
					Exit For
				End If
				Next
			End If
		Next
		
		vAllTotTarget = vTotalTarget1 + vTotalTarget2 + vTotalTarget3 + vTotalTarget4 + vTotalTarget5 + vTotalTarget6 + vTotalTarget7 + vTotalTarget8 + vTotalTarget9 + vTotalTarget10 + vTotalTarget11 + vTotalTarget12
		vAllTotProfit = vProfit1 + vProfit2 + vProfit3 + vProfit4 + vProfit5 + vProfit6 + vProfit7 + vProfit8 + vProfit9 + vProfit10 + vProfit11 + vProfit12
		vAllTotItemCost = vItemcost1 + vItemcost2 + vItemcost3 + vItemcost4 + vItemcost5 + vItemcost6 + vItemcost7 + vItemcost8 + vItemcost9 + vItemcost10 + vItemcost11 + vItemcost12
		vAllTotMaechul = vMaechulProf1 + vMaechulProf2 + vMaechulProf3 + vMaechulProf4 + vMaechulProf5 + vMaechulProf6 + vMaechulProf7 + vMaechulProf8 + vMaechulProf9 + vMaechulProf10 + vMaechulProf11 + vMaechulProf12
		
		vLastAllTotTarget = vLastTotalTarget1 + vLastTotalTarget2 + vLastTotalTarget3 + vLastTotalTarget4 + vLastTotalTarget5 + vLastTotalTarget6 + vLastTotalTarget7 + vLastTotalTarget8 + vLastTotalTarget9 + vLastTotalTarget10 + vLastTotalTarget11 + vLastTotalTarget12
		vLastAllTotProfit = vLastProfit1 + vLastProfit2 + vLastProfit3 + vLastProfit4 + vLastProfit5 + vLastProfit6 + vLastProfit7 + vLastProfit8 + vLastProfit9 + vLastProfit10 + vLastProfit11 + vLastProfit12
		vLastAllTotItemCost = vLastItemcost1 + vLastItemcost2 + vLastItemcost3 + vLastItemcost4 + vLastItemcost5 + vLastItemcost6 + vLastItemcost7 + vLastItemcost8 + vLastItemcost9 + vLastItemcost10 + vLastItemcost11 + vLastItemcost12
		vLastAllTotMaechul = vLastMaechulProf1 + vLastMaechulProf2 + vLastMaechulProf3 + vLastMaechulProf4 + vLastMaechulProf5 + vLastMaechulProf6 + vLastMaechulProf7 + vLastMaechulProf8 + vLastMaechulProf9 + vLastMaechulProf10 + vLastMaechulProf11 + vLastMaechulProf12
	End If
	
	
	
sub sbCateNotExistHTML()
%>
<tr align="center"  bgcolor="#F0F0F0" height="25">
	<td rowspan="5">전년실적</td><td align="right">구매총액</td>
	<td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td>
	<td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right"><strong>0</strong></td><td align="right"><strong>0%</strong></td>
</tr>
<tr align="center"  bgcolor="#F0F0F0" height="25">
	<td align="right">달성율</td>
	<td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td>
	<td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td></td>
</tr>
<tr align="center"  bgcolor="#F0F0F0" height="25">
	<td align="right">수익</td>
	<td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td>
	<td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right">0</td><td align="right"><strong>0</strong></td><td align="right"><strong>0%</strong></td>
</tr>
<tr align="center"  bgcolor="#F0F0F0" height="25">
	<td align="right">달성율</td>
	<td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td>
	<td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td></td>
</tr>
<tr align="center"  bgcolor="#F0F0F0" height="25">
	<td align="right">수익율</td>
	<td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td>
	<td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td align="right">0%</td><td></td>
</tr>
<%
end sub
%>