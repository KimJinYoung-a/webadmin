<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/mustPriceCls.asp"-->
<%
Dim makerid, itemid, mallgubun, isGetDate
Dim page, i, mwdiv
Dim oMustPrice
page                = request("page")
makerid				= requestCheckVar(request("makerid"), 32)
itemid  			= request("itemid")
mallgubun           = requestCheckVar(request("mallgubun"), 32)
isGetDate           = requestCheckVar(request("isGetDate"), 1)
mwdiv				= request("mwdiv")

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

SET oMustPrice = new CMustPrice
	oMustPrice.FCurrPage					= page
	oMustPrice.FPageSize					= 50
    oMustPrice.FRectMakerid					= makerid
	oMustPrice.FRectItemID					= itemid
    oMustPrice.FRectMallgubun				= mallgubun
    oMustPrice.FRectIsGetDate		    	= isGetDate
	oMustPrice.FRectMwdiv		    		= mwdiv
    oMustPrice.getMustPirceItemList
%>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
function popMustPrice(){
	var popMustPrice = window.open("/admin/etc/popMustPrice.asp","popMustPrice","width=700,height=400,scrollbars=yes,resizable=yes");
	popMustPrice.focus();
}
function fnModifyMustPrice(iidx, mallid){
	var popMustPrice = window.open("/admin/etc/popMustPrice.asp?idx="+iidx+"&isModify=Y&mallid="+mallid,"popMustPrice","width=700,height=400,scrollbars=yes,resizable=yes");
	popMustPrice.focus();
}
function popUploadExcel(){
	var popwin;
	popwin = window.open("/admin/etc/mustprice/popUploadMustPrice.asp", "popup_item", "width=500,height=230,scrollbars=yes,resizable=yes");
	popwin.focus();
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
		document.frmSvArr.action = "/admin/etc/mustPrice_process.asp"
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
        몰 구분 :
        <select name="mallgubun" class="select">
            <option value="">-Choice-</option>
            <option value="ssg" <%= CHKiif(mallgubun="ssg","selected","") %> >SSG</option>
            <option value="coupang" <%= CHKiif(mallgubun="coupang","selected","") %> >쿠팡</option>
            <option value="halfclub" <%= CHKiif(mallgubun="halfclub","selected","") %> >하프클럽</option>
			<option value="hmall1010" <%= CHKiif(mallgubun="hmall1010","selected","") %> >HMall</option>
            <option value="auction1010" <%= CHKiif(mallgubun="auction1010","selected","") %> >옥션</option>
            <option value="ezwel" <%= CHKiif(mallgubun="ezwel","selected","") %> >이지웰페어</option>
            <option value="gmarket1010" <%= CHKiif(mallgubun="gmarket1010","selected","") %> >G마켓</option>
            <option value="gsshop" <%= CHKiif(mallgubun="gsshop","selected","") %> >GSShop</option>
            <option value="interpark" <%= CHKiif(mallgubun="interpark","selected","") %> >인터파크</option>
            <option value="nvstorefarm" <%= CHKiif(mallgubun="nvstorefarm","selected","") %> >스토어팜</option>
			<option value="Mylittlewhoopee" <%= Chkiif(mallgubun = "Mylittlewhoopee", "selected", "") %>>스토어팜 캣앤독</option>
			<option value="nvstoregift" <%= CHKiif(mallgubun="nvstoregift","selected","") %> >스토어팜 선물하기</option>
            <option value="WMP" <%= CHKiif(mallgubun="WMP","selected","") %> >위메프</option>
			<option value="11st1010" <%= CHKiif(mallgubun="11st1010","selected","") %> >11번가</option>
            <option value="lotteCom" <%= CHKiif(mallgubun="lotteCom","selected","") %> >롯데닷컴</option>
            <option value="lotteimall" <%= CHKiif(mallgubun="lotteimall","selected","") %> >롯데아이몰</option>
			<option value="lotteon" <%= CHKiif(mallgubun="lotteon","selected","") %> >롯데On</option>
			<option value="skstoa" <%= CHKiif(mallgubun="skstoa","selected","") %> >SKSTOA</option>
			<option value="shintvshopping" <%= CHKiif(mallgubun="shintvshopping","selected","") %> >신세계TV쇼핑</option>
			<option value="wetoo1300k" <%= CHKiif(mallgubun="wetoo1300k","selected","") %> >1300k</option>
            <option value="cjmall" <%= CHKiif(mallgubun="cjmall","selected","") %> >CJMall</option>
			<option value="lfmall" <%= Chkiif(mallgubun = "lfmall", "selected", "") %>>LFmall</option>
			<option value="sabangnet" <%= Chkiif(mallgubun = "sabangnet", "selected", "") %>>사방넷</option>
			<option value="kakaogift" <%= Chkiif(mallgubun = "kakaogift", "selected", "") %>>카카오기프트</option>
			<option value="kakaostore" <%= Chkiif(mallgubun = "kakaostore", "selected", "") %>>카카오톡스토어</option>
			<option value="boribori1010" <%= Chkiif(mallgubun = "boribori1010", "selected", "") %>>보리보리</option>
			<option value="wconcept1010" <%= Chkiif(mallgubun = "wconcept1010", "selected", "") %>>W컨셉</option>
			<option value="benepia1010" <%= Chkiif(mallgubun = "benepia1010", "selected", "") %>>베네피아</option>
        </select>
        &nbsp;
        특가진행여부(현재날짜기준) :
        <select name="isGetDate" class="select">
            <option value="" >-Choice-</option>
            <option value="Y" <%= CHKiif(isGetDate="Y","selected","") %> >진행중</option>
        </select>
		거래구분 <% drawSelectBoxMWU "mwdiv", mwdiv %>
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
<input type="hidden" name="mallgubun" value="<%= mallgubun %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		검색결과 : <b><%= FormatNumber(oMustPrice.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oMustPrice.FTotalPage,0) %></b>
	</td>
	<td align="right" colspan="2">
        <input type="button" class="button" value="관리" onclick="popMustPrice();" />
		&nbsp;
        <input type="button" class="button" value="엑셀등록" onclick="popUploadExcel();" />
    <% If mallgubun <> "" Then %>
        &nbsp;
        <input type="button" class="button" value="삭제" onclick="fnDelItems();" />
    <% End If %>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="70">몰구분</td>
    <td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
    <td width="200">특가기간</td>
    <td width="70">특가</td>
	<td width="70">특가시<br>마진</td>
	<td width="70">특가시<br>매입가</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">거래구분</td>
	<td width="70">주문제작<br>여부</td>
	<td width="80">수정자ID</td>
</tr>
<% For i = 0 To oMustPrice.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oMustPrice.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oMustPrice.FItemList(i).Fsmallimage %>" width="50"></td>
    <td><%= oMustPrice.FItemList(i).FMallgubun %></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oMustPrice.FItemList(i).FItemID%>" target="_blank"><%= oMustPrice.FItemList(i).FItemID %></a>
	</td>
	<td align="left" style="cursor:pointer;" onclick="fnModifyMustPrice('<%= oMustPrice.FItemList(i).FIdx %>', '<%= oMustPrice.FItemList(i).FMallgubun %>');">
        <%= oMustPrice.FItemList(i).FMakerid %><%= oMustPrice.FItemList(i).getDeliverytypeName %><br><%= oMustPrice.FItemList(i).FItemName %>
    </td>
	<td>
		<%= FormatDate(oMustPrice.FItemList(i).FStartDate,"0000-00-00 00:00:00") %> <br />~ <%= FormatDate(oMustPrice.FItemList(i).FEndDate,"0000-00-00 00:00:00") %>
	</td>
	<td align="right">
		<%= FormatNumber(oMustPrice.FItemList(i).FMustPrice,0) %>
	</td>
	<td align="right">
	<%
		If oMustPrice.FItemList(i).FMustMargin = 0 Then
			response.write "설정안됨"
		Else
			response.write oMustPrice.FItemList(i).FMustMargin & "%"
		End If
	%>
	</td>
	<td align="right">
	<%
		If oMustPrice.FItemList(i).FMustMargin = 0 Then
			response.write "설정안됨"
		Else
			response.write FormatNumber(oMustPrice.FItemList(i).FMustBuyPrice,0)
		End If
	%>
	</td>
	<td align="right">
	<% If oMustPrice.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oMustPrice.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oMustPrice.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oMustPrice.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
	<%
		If oMustPrice.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oMustPrice.FItemList(i).Fbuycash/oMustPrice.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oMustPrice.FItemList(i).IsSoldOut Then
			If oMustPrice.FItemList(i).FSellyn = "N" Then
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
		Select Case oMustPrice.FItemList(i).FMWDiv
			Case "M"	response.write "매입"
			Case "W"	response.write "위탁"
			Case "U"	response.write "업체"
		End Select
	%>
	</td>
	<td align="center">
	<%
		If oMustPrice.FItemList(i).FItemdiv = "06" OR oMustPrice.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center"><%= Chkiif(oMustPrice.FItemList(i).Freguserid <> "", oMustPrice.FItemList(i).Freguserid, oMustPrice.FItemList(i).FLastUpdateUserId ) %></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18" align="center">
	<% If oMustPrice.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oMustPrice.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oMustPrice.StartScrollPage To oMustPrice.FScrollCount + oMustPrice.StartScrollPage - 1 %>
		<% If i>oMustPrice.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oMustPrice.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oMustPrice = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->