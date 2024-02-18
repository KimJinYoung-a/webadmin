<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 회원 구매 히스토리
' Hieditor : 2011.02.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #Include virtual = "/lib/classes/totalpoint/totalpointCls.asp" -->

<%
dim vCardNo, posuid, pssnkey, dummikey, shopid, vTelNo, vHpNo, vSearchCardNo
Dim arrList, intLoop, arrFileList, i, arrCardList ,UserName ,CardNo
Dim vUserID, vUserSeq, vUserName, vJumin1, vJumin2_Enc, vPoint, vGrade, vSexFlag
Dim vZipCode, vAddress, vAddressDetail, vEmail, vEmailYN, vSMSYN, vUserStatus, vLastUpdate
dim vRegdate, vShopName, vUseYN, vTotalPoint	
	vUserSeq		= requestCheckVar(Request("userseq"),10)
	vUserID			= requestCheckVar(Request("userid"),32)
	UserName		= requestCheckVar(Request("username"),20)
	CardNo			= requestCheckVar(Request("cardno"),20)
	vUseYN			= requestCheckVar(Request("useyn"),20)
	vSearchCardNo	= requestCheckVar(Request("searchcardno"),20)
	posuid			= Request("posuid")
	pssnkey			= Request("pssnkey")
	dummikey		= Request("dummikey")
	shopid = request("shopid")
	menupos = request("menupos")			
		
	If vUserSeq = "" Then
		Response.Write "<script>alert('잘못된 경로입니다.');window.close();</script>"
		Response.End
	Else

		Dim totalpointView
		Set totalpointView = New TotalPoint
		totalpointView.FUserSeq = vUserSeq
		totalpointView.GetTotalPointDetail
	
		If totalpointView.FTotCnt = "0" Then
			Response.Write "<script>alert('잘못된 경로입니다.');window.close();</script>"
			dbget.close()
			Response.End
		ElseIf totalpointView.FTotCnt > "1" Then
			'Response.Write "<script>alert('중복되어 들어간 회원입니다.');</script>"
		End If
		
		vUserName		= totalpointView.FUserName
		vJumin1			= totalpointView.FJumin1
		vJumin2_Enc		= totalpointView.FJumin2_Enc
		vCardNo			= totalpointView.FCardNo
		vPoint			= totalpointView.FPoint
		vGrade			= totalpointView.FGrade
		vSexFlag		= totalpointView.FSexFlag
		vTelNo			= totalpointView.FTelNo
		vHpNo			= totalpointView.FHpNo
		vZipCode		= totalpointView.FZipCode
		vAddress		= totalpointView.FAddress
		vAddressDetail	= totalpointView.FAddressDetail
		vEmail			= totalpointView.FEmail
		vEmailYN		= totalpointView.FEmailYN
		vSMSYN			= totalpointView.FSMSYN
		vUserStatus		= totalpointView.FUserStatus
		vLastUpdate		= totalpointView.FLastUpdate
		vRegdate		= totalpointView.FRegdate
		vShopName		= totalpointView.FShopName

		totalpointView.FCardNo = vSearchCardNo
		arrList = totalpointView.GetTotalPointLogList()
		vTotalPoint		= totalpointView.FTotalPoint
		
		arrCardList = totalpointView.GetMemberCardList()
		
		set totalpointView = Nothing
	End If	
%>

<script language="javascript">

function refer(){
	frm.action='/admin/totalpoint/customer_sell_history.asp';
	frm.submit();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="50%"><img src="/images/icon_arrow_link.gif" valign="absbottom">&nbsp;<b>기본정보</b></td>
			<td width="50%" align="right">
				<input type="button" class="button" value="목록으로" onClick="refer();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">회원번호</td>
	<td width="300" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vUserSeq%></td>
	<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">고객명</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
	<%
		If vGrade <> "0" Then
			Response.Write "[아이띵소특별회원] "
		End If
		Response.Write vUserName & "&nbsp;"
		If vSexFlag = "1" Then
			Response.Write "(남)"
		ElseIf vSexFlag = "2" Then
			Response.Write "(여)"
		Else
			Response.Write "(" & vSexFlag & ")"
		End If
	%>
	</td>
</tr>

<tr>
	<td align="center"  bgcolor="<%= adminColor("tabletop") %>">전화번호</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= printUserId(vTelNo,4,"*") %></td>
	<td align="center"  bgcolor="<%= adminColor("tabletop") %>">MAIL수신여부</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vEmailYN%></td>
</tr>
<tr>
	<td align="center"  bgcolor="<%= adminColor("tabletop") %>">핸드폰번호</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= printUserId(vHpNo,4,"*") %></td>
	<td align="center"  bgcolor="<%= adminColor("tabletop") %>">SMS수신여부</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vSMSYN%></td>
</tr>
<tr>
	<td align="center"  bgcolor="<%= adminColor("tabletop") %>">사용카드번호</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vCardNo%></td>
	<td align="center"  bgcolor="<%= adminColor("tabletop") %>">포인트</td>
	<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vTotalPoint%> Point</td>
</tr>
</tr>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="userseq" value="<%=vUserSeq%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="posuid" value="<%=posuid%>">
<input type="hidden" name="pssnkey" value="<%=pssnkey%>">
<input type="hidden" name="dummikey" value="<%=dummikey%>">
<input type="hidden" name="cardno" value="<%=CardNo%>">
<input type="hidden" name="username" value="<%=UserName%>">
<input type="hidden" name="userid" value="<%=vUserID%>">
<input type="hidden" name="shopid" value="<%=shopid%>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<img src="/images/icon_arrow_link.gif" valign="absbottom">&nbsp;<b>상세정보</b>
		&nbsp;
		<select name="searchcardno" class="select" onChange="frm.submit();">
			<option value="">카드전체</option>
		<%
			IF isArray(arrCardList) THEN
				For intLoop =0 To UBound(arrCardList,2)

					Response.Write "<option value='" & arrCardList(0,intLoop) & "'"
					If vSearchCardNo = arrCardList(0,intLoop) Then
						Response.Write " selected"
					End If
					Response.Write ">" & arrCardList(0,intLoop) & "</option>"

				Next
			End If
		%>
		</select>
	</td>
</tr>
</form>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="70">카드구분</td>
	<td width="140">등록일</td>
	<td width="100">카드번호</td>
	<td width="120">거래구분</td>
	<td width="50">포인트</td>
	<td>관련주문번호</td>
</tr>
<%
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
%>

<tr align="center" bgcolor="#FFFFFF">
	<td>
		<% If Left(arrList(0,intLoop),4) = "1010" Then %>
			POINT1010
		<% ElseIf Left(arrList(0,intLoop),4) = "3253" Then %>
			아이띵소
		<% Else %>
			오프라인
		<% End If %>
	</td>
	<td><%=arrList(7,intLoop)%></td>
	<td><%=arrList(0,intLoop)%></td>
	<td>
		<%
			'### 포인트 0이고 code가 3(포인트이관)일때 카드등록으로 나타냄.
			If arrList(1,intLoop) = "0" AND arrList(8,intLoop) = "3" Then
				Response.Write arrList(4,intLoop)
			Else
				Response.Write arrList(2,intLoop)
			End IF
		%>
	</td>
	<td><%=arrList(1,intLoop)%></td>
	<td><%=arrList(5,intLoop)%></td>
</tr>
					
<%
		Next
	End If
%>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->
