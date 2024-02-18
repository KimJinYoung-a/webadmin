<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ezwel/ezwelcls.asp"-->
<%
Dim oEzwel, i, ezwelStatcd, ezwelGoodNoArray, itemids, mode
Dim page, confirmGoodnoArr
Dim sqlStr
ezwelStatcd = request("ezwelStatcd")
itemids = request("itemids")
confirmGoodnoArr = trim(request("confirmGoodnoArr"))
mode = request("mode")
page = request("page")

If page = "" Then page = 1
If ezwelStatcd = "" Then ezwelStatcd = 3

If mode = "I" Then
	If confirmGoodnoArr <> "" then
		Dim iA2, arrTemp2, arrconfirmGoodnoArr
		confirmGoodnoArr = replace(confirmGoodnoArr,",",chr(10))
		confirmGoodnoArr = replace(confirmGoodnoArr,chr(13),"")
		arrTemp2 = Split(confirmGoodnoArr,chr(10))
		iA2 = 0
		Do While iA2 <= ubound(arrTemp2)
			If Trim(arrTemp2(iA2))<>"" then
				arrconfirmGoodnoArr = arrconfirmGoodnoArr& "'"& trim(arrTemp2(iA2)) & "',"
			End If
			iA2 = iA2 + 1
		Loop
		confirmGoodnoArr = left(arrconfirmGoodnoArr,len(arrconfirmGoodnoArr)-1)

		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_etcmall.dbo.tbl_ezwel_regitem"
		sqlStr = sqlStr & " SET ezwelstatcd= '7' "
		sqlStr = sqlStr & " , ezwelSellYn = 'Y' "
		sqlStr = sqlStr & " , accFailCnt = 0 "
		sqlStr = sqlStr & " WHERE ezwelGoodno in ("& confirmGoodnoArr &") "
		dbget.execute sqlStr
		Response.Write "<script>alert('처리되었습니다.');self.close();</script>"
		Response.End
	Else
		Response.Write "<script>alert('상품코드 없음');self.close();</script>"
		Response.End
	End If
Else
	Set oEzwel = new CEzwel
		oEzwel.FCurrPage	= page
		oEzwel.FPageSize	= 1000
		oEzwel.FRectExtNotReg = ezwelStatcd
		oEzwel.getEzwelStatcdList
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function statCdConfirm(){
	if ($("#confirmGoodnoArr").val() == "") {
		alert('승인할 상품을 입력하세요');
		$("#confirmGoodnoArr").focus();
		return;
	}


	if(confirm("이지웰페어 MD에게 승인처리 확인 후 진행해야 합니다.\n\n확인 시 리스트의 판매상태는 Y로, 승인상태는 전시중으로 변경됩니다.\n\n승인처리 계속 하시겠습니까?")){
		document.frms.submit();
	}
}
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

</script>
<!--
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="<%= page %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		등록여부 :
		<select name="ezwelStatcd" class="select" >
			<option value="" >전체
			<option value="3" <%= CHkIIF(ezwelStatcd="3","selected","") %> >Ezwel 승인예정
			<option value="4" <%= CHkIIF(ezwelStatcd="4","selected","") %> >Ezwel 재판매예정
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
-->
<p />
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="<%= page %>">
</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oEzwel.FTotalCount,0) %></b>
		<% If oEzwel.FResultCount > 0 Then %>
			<br />
			<form name="frms" method="post">
				<input type="hidden" name="mode" value="I">
				승인할 ezwel상품코드 : <textarea id="confirmGoodnoArr" name="confirmGoodnoArr"></textarea>
			</form>
			<input type="button" class="button" value="승인처리" onclick="statCdConfirm();">
		<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">Ezwel등록일<br>Ezwel최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">Ezwel<br>가격및판매</td>
	<td width="70">Ezwel<br>실가격</td>
	<td width="70">Ezwel<br>상품번호</td>
	<td width="70">Ezwel<br>비고</td>
</tr>
<% For i=0 to oEzwel.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center"><%= oEzwel.FItemList(i).FItemID %>
		<% If oEzwel.FItemList(i).FLimitYn= "Y" Then %><br><%= oEzwel.FItemList(i).getLimitHtmlStr %></font><% End If %>
	</td>
	<td align="left"><%= oEzwel.FItemList(i).FMakerid %> <%= oEzwel.FItemList(i).getDeliverytypeName %><br><%= oEzwel.FItemList(i).FItemName %></td>
	<td align="center"><%= oEzwel.FItemList(i).FRegdate %><br><%= oEzwel.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oEzwel.FItemList(i).FEzwelRegdate %><br><%= oEzwel.FItemList(i).FEzwelLastUpdate %></td>
	<td align="right">
		<% If oEzwel.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oEzwel.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oEzwel.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oEzwel.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oEzwel.FItemList(i).Fsellcash = 0 Then
		elseif (oEzwel.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oEzwel.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oEzwel.FItemList(i).FOrgSuplycash/oEzwel.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oEzwel.FItemList(i).Fbuycash/oEzwel.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oEzwel.FItemList(i).Fbuycash/oEzwel.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oEzwel.FItemList(i).IsSoldOut Then
			If oEzwel.FItemList(i).FSellyn = "N" Then
	%>
		<font color="red">품절</font>
	<%
			Else
	%>
		<font color="red">일시<br>품절</font>
	<%
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oEzwel.FItemList(i).FEzwelStatCd > 0) Then
			If Not IsNULL(oEzwel.FItemList(i).FEzwelPrice) Then
				If (oEzwel.FItemList(i).Fsellcash<>oEzwel.FItemList(i).FEzwelPrice) Then
	%>
					<strong><%= formatNumber(oEzwel.FItemList(i).FEzwelPrice,0) %></strong>
	<%
				Else
					response.write formatNumber(oEzwel.FItemList(i).FEzwelPrice,0)
				End If
	%>
				<br>
	<%
				If Not IsNULL(oEzwel.FItemList(i).FSpecialPrice) Then
					If (now() >= oEzwel.FItemList(i).FStartDate) And (now() <= oEzwel.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oEzwel.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oEzwel.FItemList(i).FSellyn="Y" and oEzwel.FItemList(i).FEzwelSellYn<>"Y") or (oEzwel.FItemList(i).FSellyn<>"Y" and oEzwel.FItemList(i).FEzwelSellYn="Y") Then
	%>
					<strong><%= oEzwel.FItemList(i).FEzwelSellYn %></strong>
	<%
				Else
					response.write oEzwel.FItemList(i).FEzwelSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<%
			If Not IsNULL(oEzwel.FItemList(i).FEzwelPrice) Then
				response.write FormatNumber(Fix(oEzwel.FItemList(i).FEzwelPrice/100)*100,0)
		 	End If
		%>
	</td>
	<td align="center">
	<%
    	Response.Write "<span style='cursor:pointer;' onclick=window.open('http://shop.ezwel.com/shopNew/goods/preview/goodsDetailView.ez?preview=yes&goodsBean.goodsCd="&oEzwel.FItemList(i).FEzwelGoodNo&"')>"&oEzwel.FItemList(i).FEzwelGoodNo&"</span><br>"
	%>
	</td>
	<td align="center">
	<%
		Select Case oEzwel.FItemList(i).FEzwelStatcd
			Case "3"	response.write "승인예정"
			Case "4"	response.write "재판매예정"
		End Select
	%>
	</td>
</tr>
<%
	ezwelGoodNoArray = ezwelGoodNoArray & oEzwel.FItemList(i).FezwelGoodNo & VBCRLF
Next
%>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oEzwel.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEzwel.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oEzwel.StartScrollPage to oEzwel.FScrollCount + oEzwel.StartScrollPage - 1 %>
    		<% if i>oEzwel.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oEzwel.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
Ezwel상품번호Array : <textarea id="ezwelGoodnoArr" name="ezwelGoodnoArr"><%= ezwelGoodNoArray %></textarea>&nbsp;
<button onclick="copyId();">Copy</button>
<script>
function copyId() {
	var ttt = document.getElementById("ezwelGoodnoArr");
	ttt.select();
	document.execCommand("copy");
}
</script>
<% Set oEzwel = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->