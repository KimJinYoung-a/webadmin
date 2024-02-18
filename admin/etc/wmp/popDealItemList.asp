<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wmp/wmpCls.asp"-->
<%
Dim makerid, itemid, isGetDate
Dim page, i
Dim oDealItem
page                = request("page")
makerid				= requestCheckVar(request("makerid"), 32)
itemid  			= request("itemid")
isGetDate           = requestCheckVar(request("isGetDate"), 1)

If page = "" Then page = 1
'텐바이텐 상품코드 엔터키로 검색되게
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

SET oDealItem = new CWmp
	oDealItem.FCurrPage					= page
	oDealItem.FPageSize					= 50
    oDealItem.FRectMakerid				= makerid
	oDealItem.FRectItemID				= itemid
    oDealItem.FRectIsGetDate		   	= isGetDate
    oDealItem.getDealItemList
%>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
function popDealItem(){
	var popDealItem = window.open("/admin/etc/wmp/popDealItem.asp","popDealItem","width=700,height=400,scrollbars=yes,resizable=yes");
	popDealItem.focus();
}
function fnModifyMustPrice(iidx){
	var popMustPrice = window.open("/admin/etc/wmp/popDealItem.asp?idx="+iidx+"&isModify=Y","popMustPrice","width=700,height=400,scrollbars=yes,resizable=yes");
	popMustPrice.focus();
}
function popOption(iitemid){
	var popOption = window.open("/admin/etc/wmp/popDealOption.asp?itemid="+iitemid,"popOption","width=700,height=400,scrollbars=yes,resizable=yes");
	popOption.focus();
}
function fnDelItems(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}
	if (confirm('선택하신 ' + chkSel + '개 삭제 하시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.mode.value = "D";
		document.frmSvArr.action = "/admin/etc/wmp/procDealItem.asp"
		document.frmSvArr.submit();
    }
}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<br /><br />
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
        &nbsp;
        딜진행여부(현재날짜기준) :
        <select name="isGetDate" class="select">
            <option value="" >-Choice-</option>
            <option value="Y" <%= CHKiif(isGetDate="Y","selected","") %> >진행중</option>
        </select>
    </td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br />
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		검색결과 : <b><%= FormatNumber(oDealItem.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDealItem.FTotalPage,0) %></b>
	</td>
	<td align="right">
        <input type="button" class="button" value="관리" onclick="popDealItem();" />
        &nbsp;
        <input type="button" class="button" value="삭제" onclick="fnDelItems();" />
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
    <td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명<br><font color="blue">변경상품명</font></td>
    <td width="300">딜기간</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">옵션관리</td>
	<td width="80">수정자ID</td>
</tr>
<% For i = 0 To oDealItem.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oDealItem.FItemList(i).FItemId %>"></td>
	<td><img src="<%= oDealItem.FItemList(i).Fsmallimage %>" width="50"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oDealItem.FItemList(i).FItemID%>" target="_blank"><%= oDealItem.FItemList(i).FItemID %></a>
	</td>
	<td align="left" style="cursor:pointer;" onclick="fnModifyMustPrice('<%= oDealItem.FItemList(i).FIdx %>');">
        <%= oDealItem.FItemList(i).FMakerid %><%= oDealItem.FItemList(i).getDeliverytypeName %><br><%= oDealItem.FItemList(i).FItemName %>
		<br/>
		<font color="blue"><%= oDealItem.FItemList(i).FNewItemName %></font>
    </td>
	<td>
		<%= FormatDate(oDealItem.FItemList(i).FStartDate,"0000-00-00 00:00:00") %> <br />~ <%= FormatDate(oDealItem.FItemList(i).FEndDate,"0000-00-00 00:00:00") %>
	</td>
	<td align="right">
	<% If oDealItem.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oDealItem.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oDealItem.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oDealItem.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
	<%
		If oDealItem.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oDealItem.FItemList(i).Fbuycash/oDealItem.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oDealItem.FItemList(i).IsSoldOut Then
			If oDealItem.FItemList(i).FSellyn = "N" Then
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
		If oDealItem.FItemList(i).FItemdiv = "06" OR oDealItem.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
    <td align="center">
		<input type="button" class="button" value="옵션" onclick="popOption('<%= oDealItem.FItemList(i).FItemId %>');">
	</td>
	<td align="center"><%= Chkiif(oDealItem.FItemList(i).Freguserid <> "", oDealItem.FItemList(i).Freguserid, oDealItem.FItemList(i).FLastUpdateUserId ) %></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18" align="center">
	<% If oDealItem.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oDealItem.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oDealItem.StartScrollPage To oDealItem.FScrollCount + oDealItem.StartScrollPage - 1 %>
		<% If i>oDealItem.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oDealItem.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oDealItem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->