<%@  language="VBScript" %>
<% option explicit %> 
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/tenmember/lib/header.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<% 
Dim oipchul,iscmlinkno   
iscmlinkno		=  requestCheckvar(Request("iSL"),10)
set oipchul = new CIpChulStorage
oipchul.FRectId = iscmlinkno
oipchul.GetIpChulMaster  

	  function sGetDivCodeName(Fdivcode)
		if Fdivcode="002" then
			sGetDivCodeName = "위탁"
		elseif Fdivcode="001" then
			sGetDivCodeName = "매입"
		elseif Fdivcode="003" then
			sGetDivCodeName = "판촉"
		elseif Fdivcode="004" then
			sGetDivCodeName = "외부"
		elseif Fdivcode="005" then
			sGetDivCodeName = "협찬"
		elseif Fdivcode="006" then
			sGetDivCodeName = "B2B"
		elseif Fdivcode="007" then
			sGetDivCodeName = "기타"
		elseif Fdivcode="101" then
			sGetDivCodeName = "위탁출고"
		elseif Fdivcode="801" then
			sGetDivCodeName = "Off매입"
		elseif Fdivcode="802" then
			sGetDivCodeName = "Off위탁"
		elseif Fdivcode="999" then
			sGetDivCodeName = "기타(정산안함)"
		end if
	end function
 %> 
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<!--전자결재-->
	<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
	<input type="hidden" name="tC" value="">
	<input type="hidden" name="ieidx" value="58"> <!-- 문서번호 지정!!-->
	<input type="hidden" name="iSL" value="<%=iscmlinkno%>">
	<input type="hidden" name="mRP" value="<%=formatnumber(oipchul.FOneItem.Ftotalbuycash*-1,0)%>">
	</form>
	<div id="divEapp" style="display:none;">
	<p>&nbsp;다음과 같이 기타출고를 진행하고자 하오니 검토 후 재가 바랍니다. </p>
	<p>&nbsp;</p>
	<p align="center">-  다  음  - </p>
	<p>&nbsp;</p>
	<p>1. 기타출고사유: </p>
	<p>&nbsp;</p>
	<p>2. 기타출고내역: </p>
	<p>&nbsp;</p>
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>출고코드</td>
			<td>출고처ID</td>
			<td>출고처명</td>
			<td>등록자</td>
			<td>요청일</td>
			<td>판매가</td>
			<td>출고가</td>
			<td>매입가</td>
			<td>구분</td>
			<td>할인율</td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td><%=oipchul.FOneItem.Fcode %></td>
			<td><%=oipchul.FOneItem.Fsocid%></td>
			<td><%= oipchul.FOneItem.Fsocname%></td>
			<td><%= oipchul.FOneItem.Fchargeid %>&nbsp;(<%= oipchul.FOneItem.Fchargename %>)</td>
			<td><%= oipchul.FOneItem.Fscheduledt %></td>
			<td><%= FormatNumber(oipchul.FOneItem.Ftotalsellcash*-1,0) %></td>
			<td><%= FormatNumber(oipchul.FOneItem.Ftotalsuplycash*-1,0) %></td>
			<td><%= FormatNumber(oipchul.FOneItem.Ftotalbuycash*-1,0) %></td>
			<td><%=sGetDivCodeName(oipchul.FOneItem.Fdivcode) %></td>
			<td><% if oipchul.FOneItem.Ftotalsellcash<>0 then %>
				  <%= 100-CLng(oipchul.FOneItem.Ftotalsuplycash/oipchul.FOneItem.Ftotalsellcash*100*100)/100 %>%
				<% end if %>
			</td>
		</tr>
	</table>
	<%if oipchul.FOneItem.Fprizecnt > 0 then%>
	<p>&nbsp;</p>
	<p>3. 당첨자정보: </p>
	<p style="color:blue">- 이벤트소득세파일 재무팀 별도제출</p>
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>No</td>
			<td>당첨자ID</td>
			<td>당첨자명</td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td>1</td>
			<td></td>
			<td></td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td>2</td>
			<td></td>
			<td></td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td>3</td>
			<td></td>
			<td></td>
		</tr>
	</table>
	<%end if%>
	<br /><br />
	<p>
		<b>* 출고처 ID</b><br />
		<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>출고처ID</td>
			<td>내용</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>etcout</td>
			<td>재고이동</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemgift</td>
			<td>당첨사은품</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemgift_all</td>
			<td>구매사은품</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemsample</td>
			<td>샘플사용 (ex.촬영샘플)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemAD</td>
			<td>광고선전비 (ex.협찬)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemgift_Biz</td>
			<td>접대비 (ex.거래처선물)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemstaff</td>
			<td>복리후생비 (ex.웰컴키트)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itempay</td>
			<td>급여귀속</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemdisuse</td>
			<td>폐기손실</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemloss</td>
			<td>감모손실</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>shopitemsample</td>
			<td>샘플사용(매장)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>shopitemloss</td>
			<td>감모손실(매장)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>parcelloss</td>
			<td>택배분실</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>donation</td>
			<td>기부</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>csservice</td>
			<td>CS비용</td>
		</tr>
		</table>
	</p>
	</div>
	 
	 <%set oipchul = nothing
	 
%>
	<script type="text/javascript">  
		document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");
	 	document.frmEapp.submit();
		</script>
	<!--/전자결재-->

