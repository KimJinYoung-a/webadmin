<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->

<%
	Dim iCurrentpage, Giftlist, i, iTotCnt, vEvtCode, vSDate, page, vGubun, vItemID, vItemName, vUseYN, vSoldOUT, vParam
	Dim diffPrc, diffCost
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page 			= NullFillWith(requestCheckVar(request("page"),5),1)
	vGubun			= NullFillWith(requestCheckVar(request("gubun"),10),"")
	vItemID			= NullFillWith(requestCheckVar(request("itemid"),10),"")
	vItemName		= NullFillWith(requestCheckVar(request("itemname"),100),"")
	vUseYN			= NullFillWith(requestCheckVar(request("useyn"),1),"")
	vSoldOUT		= NullFillWith(requestCheckVar(request("soldout"),1),"")
	diffPrc			= NullFillWith(requestCheckVar(request("diffPrc"),2),"")
	diffCost		= NullFillWith(requestCheckVar(request("diffCost"),2),"")
	vParam = "&menupos="&Request("menupos")&"&gubun="&vGubun&"&itemid="&vItemID&"&useyn="&vUseYN&"&soldout="&vSoldOUT&"&itemname="&vItemName&"&diffPrc="&diffprc&"&diffCost="&diffCost

	Set Giftlist = new ClsGift
	Giftlist.FCurrPage = page
	Giftlist.FGubun = vGubun
	Giftlist.FItemID = vItemID
	Giftlist.FItemName = vItemName
	Giftlist.FUseYN = vUseYN
	Giftlist.FSoldOUT = vSoldOUT
	Giftlist.FRectdiffPrc = diffPrc
	Giftlist.FRectdiffCost = diffCost
	Giftlist.FGiftList

	iTotCnt = Giftlist.ftotalcount
%>

<script language="javascript">
function GiftWrite(id)
{
	var gift = window.open('gift_write.asp?idx='+id+'','gift','width=300,height=250');
	gift.focus();
}
</script>

<!-- 리스트 시작 -->
<form name="frm" method="get" action="index.asp">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="40" bgcolor="FFFFFF">
	<td colspan="10">
		구분 :
		<select name="gubun">
			<option value="">-선택-</option>
			<option value="giftting" <%=CHKIIF(vGubun="giftting","selected","")%>>카카오기프트</option>
			<option value="gifticon" <%=CHKIIF(vGubun="gifticon","selected","")%>>기프티콘</option>
			<option value="celectory" <%=CHKIIF(vGubun="celectory","selected","")%>>셀렉토리</option>
			<option value="gsisuper" <%=CHKIIF(vGubun="gsisuper","selected","")%>>GS아이슈퍼</option>
		</select>
		&nbsp;
		상품코드 :
		<input type="text" name="itemid" value="<%=vItemID%>" maxlength="9" size="10">
		&nbsp;
		상품명 :
		<input type="text" name="itemname" value="<%=vItemName%>" size="30">
		&nbsp;
		기프티콘 사용여부 :
		<select name="useyn">
			<option value="">-선택-</option>
			<option value="Y" <%=CHKIIF(vUseYN="Y","selected","")%>>Y</option>
			<option value="N" <%=CHKIIF(vUseYN="N","selected","")%>>N</option>
		</select>
		&nbsp;
		품절여부 :
		<select name="soldout">
			<option value="">-선택-</option>
			<option value="Y" <%=CHKIIF(vSoldOUT="Y","selected","")%>>품절</option>
			<option value="N" <%=CHKIIF(vSoldOUT="N","selected","")%>>판매중</option>
		</select>
		<br>
		<input type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> >10x10 판매가<>기프티콘 상품가
		&nbsp;
		<input type="checkbox" name="diffCost" <%= ChkIIF(diffCost="on","checked","") %> >10x10(예상)배송비<>기프티콘 배송비
	</td>
	<td colspan="10">
		<input type="submit" class="button" value="검 색">
		&nbsp;
	</td>
</tr>
</table>
</form>
<%
	IF application("Svr_Info") = "Dev" THEN
		Response.Write "<font color='blue'>※ 기프트카드 374487, 374488, 374489, 374490, 374491 는 제외되서 보여집니다.</font>"
	Else
		Response.Write "<font color='blue'>※ 기프트카드 588084, 588085, 588088, 588089, 588095 는 제외되서 보여집니다.</font>"
	End If
%>
<table cellpadding="0" cellspacing="0" border="0" class="a">
<tr height="30">
	<td width="120">
		Total Count : <b><%= iTotCnt %></b>
	</td>
	<td>
		<input type="button" value=" 상품등록 " class="button" onClick="GiftWrite('')">
	</td>
</tr>
</table>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#E6E6E6">
	<td align="center">구분</td>
	<td align="center">상품</td>
	<td align="center">상품코드</td>
	<td align="center">상품명</td>
	<td align="center">총판매가</td>
	<td align="center">상품가</td>
	<td align="center">배송비</td>
	<td align="center">10x10<br>배송구분</td>
	<td align="center">10x10<br>판매가</td>
	<td align="center">10x10<br>배송비</td>
	<td align="center">10x10<br>품절여부</td>
	<td align="center">기프티콘<br>사용여부</td>
	<td align="center"></td>
</tr>
<%
	If Giftlist.FResultCount <> 0 Then
		For i = 0 To Giftlist.FResultCount -1
%>
		<tr bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
			<td width="70" align="center">
			<%
				If Giftlist.FItemList(i).fgubun = "giftting" Then
					Response.Write "kakaogift"
				ElseIf Giftlist.FItemList(i).fgubun = "gifticon" Then
					Response.Write "기프티콘"
				ElseIf Giftlist.FItemList(i).fgubun = "celectory" Then
					Response.Write "셀렉토리"
				ElseIf Giftlist.FItemList(i).fgubun = "gsisuper" Then
					Response.Write "GS아이슈퍼"
				End IF
			%>
			</td>
			<td width="60" align="center"><a href="http://www.10x10.co.kr/<%=Giftlist.FItemList(i).fitemid%>" target="_blank"><img src="<%=Giftlist.FItemList(i).fsmallimage%>" border="0"></a></td>
			<td width="60" align="center"><%=Giftlist.FItemList(i).fitemid%></td>
			<td><%=Giftlist.FItemList(i).fitemname%></td>
			<td width="60" align="center"><%=FormatNumber(Giftlist.FItemList(i).ftot_sellcash,0) %></td>
			<td width="60" align="center"><%=FormatNumber(Giftlist.FItemList(i).fsellcash,0) %></td>
			<td width="60" align="center"><%=FormatNumber(Giftlist.FItemList(i).fdili_itemcost,0) %></td>
			<td width="90" align="center"><%= Giftlist.FItemList(i).getDeliverytypeName %></td>
			<td width="90" align="center"><%=FormatNumber(Giftlist.FItemList(i).FTenSellcash,0) %></td>
			<td width="60" align="center"><%=FormatNumber(Giftlist.FItemList(i).FItemcost,0) %></td>
			<td width="60" align="center"><%=CHKIIF(Giftlist.FItemList(i).fsoldout="True","<b><font color=red>품절</font></b>","판매중") %></td>
			<td width="60" align="center"><%=Giftlist.FItemList(i).fuseyn %></td>
			<td width="60" align="center"><input type="button" class="button" value="수정" onClick="GiftWrite(<%=Giftlist.FItemList(i).fidx%>)"></td>
		</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
		</tr>
<%
	End If
%>
<tr bgcolor="#FFFFFF">
	<td align="center" style="padding:10 0 10 0" colspan="13">
		<a href="?page=1<%=vParam%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev02.gif" width="9" height="9" border="0" /></a>
		<% if Giftlist.HasPreScroll then %>
			&nbsp;&nbsp;<a href="?page=<%= Giftlist.StartScrollPage-1 %><%=vParam%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" /></a>
		<% else %>
			&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" />
		<% end if %>
		<%
		for i = 0 + Giftlist.StartScrollPage to Giftlist.StartScrollPage + Giftlist.FScrollCount - 1
		if (i > Giftlist.FTotalpage) then Exit for
		if CStr(i) = CStr(Giftlist.FCurrPage) then
		%>
			&nbsp;&nbsp;&nbsp;&nbsp;<span class="eng11pxblack"><b><%= i %></b></span>
		<% else %>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="?page=<%= i %><%=vParam%>" style="cursor:pointer"><%= i %></a>
		<%
		end if
		next
		%>
		<% if Giftlist.HasNextScroll then %>
			&nbsp;&nbsp;<span class="list_link"><a href="?page=<%= i %><%=vParam%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" /></a>
		<% else %>
			&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" />
		<% end if %>
		&nbsp;&nbsp;&nbsp;<a href="?page=<%= Giftlist.FTotalpage %><%=vParam%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext02.gif" width="9" height="9" border="0" /></a>
	</td>
</tr>
</table>

<%
	set Giftlist = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->