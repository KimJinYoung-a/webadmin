<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
'####################################################
' Description :  온라인 환율 관리
' History : 2013.05.02 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyheadUTF8.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->
<%
dim idx, sitename, currencyUnit, currencyChar, exchangeRate, basedate, regdate, lastupdate, makerid
dim reguserid, lastuserid, page, i, countrylangcd, multipleRate, linkPriceType
	currencyUnit = requestCheckVar(request("currencyUnit"),16)
	sitename = requestCheckVar(request("sitename"),32)
	idx = requestCheckVar(getNumeric(request("idx")),10)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
	page = requestCheckVar(getNumeric(request("page")),10)
	makerid = requestCheckVar(request("makerid"),32)
if page="" then page=1

dim oexchangerate
set oexchangerate = new cexchangerate
	oexchangerate.frectidx = idx
	oexchangerate.frectcurrencyUnit = currencyUnit
	oexchangerate.frectsitename = sitename

	if (currencyUnit <> "" and sitename <> "") or idx <> "" then
		oexchangerate.fexchangerate_oneitem

		if oexchangerate.ftotalcount > 0 then
			idx = oexchangerate.FOneItem.fidx
			sitename = oexchangerate.FOneItem.fsitename
			currencyUnit = oexchangerate.FOneItem.fcurrencyUnit
			currencyChar = oexchangerate.FOneItem.fcurrencyChar
			exchangeRate = oexchangerate.FOneItem.fexchangeRate
			basedate = oexchangerate.FOneItem.fbasedate
			regdate = oexchangerate.FOneItem.fregdate
			lastupdate = oexchangerate.FOneItem.flastupdate
			reguserid = oexchangerate.FOneItem.freguserid
			lastuserid = oexchangerate.FOneItem.flastuserid
			countrylangcd = oexchangerate.FOneItem.fcountryLangCD
			multipleRate = oexchangerate.FOneItem.fmultipleRate
			linkPriceType = oexchangerate.FOneItem.flinkPriceType
			makerid = oexchangerate.FOneItem.FMakerid
		end if
	end if

dim oexchangerateList
set oexchangerateList = new cexchangerate
	oexchangerateList.FPageSize=50
	oexchangerateList.FCurrPage= page
	oexchangerateList.fexchangerate_list

if exchangeRate = "" then exchangeRate = 0
if multipleRate = "" then multipleRate = 1
%>

<script type='text/javascript'>

function delcurrencyUnit(frm){
    if (frm.sitename.value==''){
        alert('사이트구분을 선택하세요.');
        frm.sitename.focus();
        return;
    }

    if (frm.currencyUnit.value==''){
        alert('화폐단위를 입력하세요.');
        frm.currencyUnit.focus();
        return;
    }

    if (confirm('삭제 하시겠습니까?')){
    	frm.mode.value='exchangeRatedel';
        frm.submit();
    }
}

function SavecurrencyUnit(frm){
	var tmpitemautoyn='';

    if (frm.sitename.value==''){
        alert('사이트구분을 선택하세요.');
        frm.sitename.focus();
        return;
    }

    if (frm.countrylangcd.value==''){
        alert('대표언어를 선택하세요.');
        frm.countrylangcd.focus();
        return;
    }

    if (frm.linkPriceType.value==''){
        alert('기준 가격을 선택하세요.');
        frm.linkPriceType.focus();
        return;
    }

    if (frm.multipleRate.value==''){
        alert('대표배수를 입력하세요.');
        frm.multipleRate.focus();
        return;
    }

    if (frm.currencyUnit.value==''){
        alert('화폐단위를 입력하세요.');
        frm.currencyUnit.focus();
        return;
    }

    if (frm.currencyChar.value==''){
        alert('화폐기호를 입력하세요.');
        frm.currencyChar.focus();
        return;
    }

    if (frm.exchangeRate.value==''){
        alert('환율을 입력하세요');
        frm.exchangeRate.focus();
        return;
    }

    if (frm.basedate.value==''){
        alert('기준일을 입력하세요.');
        frm.basedate.focus();
        return;
    }

	//수정일 경우에만
	if (frm.idx.value!=''){
		//사이트구분이 홀쎄일 일경우
		if (frm.sitename.value=='WSLWEB'){
			//사이트구분과 화폐단위가 수정전과 수정후가 같은거
			if (frm.orgsitename.value==frm.sitename.value && frm.orgcurrencyUnit.value==frm.currencyUnit.value){
				//환율이나 대표배수가 수정 될경우
				if (frm.orgmultipleRate.value!=frm.multipleRate.value || frm.orgexchangeRate.value!=frm.exchangeRate.value || frm.orglinkPriceType.value!=frm.linkPriceType.value){
					if (frm.itemautoyn.value=='Y'){
						tmpitemautoyn='Y';
					}
				}
			}

			if (tmpitemautoyn=='Y'){
			    if (confirm('상품단에 환율,배수가 일괄 적용 됩니다.\n원하지 않으시면 상품단 환율,배수 일괄적용을 N 으로 선택 하세요.\n저장 하시겠습니까?')){
			    	frm.mode.value='exchangeRateedit';
			        frm.submit();
			    }
			}else{
			    if (confirm('저장 하시겠습니까?\n변동부분이 없어서 상품단에 환율,배수는 일괄 적용되지 않습니다.')){
			    	frm.mode.value='exchangeRateedit';
			        frm.submit();
			    }
			}
		}else{
		    if (confirm('저장 하시겠습니까?')){
		    	frm.mode.value='exchangeRateedit';
		        frm.submit();
		    }
		}
	}else{
	    if (confirm('저장 하시겠습니까?')){
	    	frm.mode.value='exchangeRateedit';
	        frm.submit();
	    }
	}
}

//신규등록
function newcurrencyUnit(){
	location.href='/common/overseas/exchangerate/exchangerate.asp?menupos=<%=menupos%>'
}

</script>

<form name="frmcurrencyUnit" method="post" action="/common/overseas/exchangerate/exchangerate_process.asp">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode">
<input type="hidden" name="orgsitename" value="<%= sitename %>">
<input type="hidden" name="orgcountrylangcd" value="<%= countrylangcd %>">
<input type="hidden" name="orgcurrencyUnit" value="<%= currencyUnit %>">
<input type="hidden" name="orgmultipleRate" value="<%= multipleRate %>">
<input type="hidden" name="orgexchangeRate" value="<%= exchangeRate %>">
<input type="hidden" name="orglinkPriceType" value="<%= linkPriceType %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" align="center">
    <td width="250">번호</td>
    <td align="left">
    	<%= IDX %>
		<input type="hidden" name="idx" value="<%=idx%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="250">사이트구분</td>
    <td align="left">
    	<%
    	'//수정모드
    	if IDX <> "" then
    	%>
    		<%= sitename %>
    		<input type="hidden" name="sitename" value="<%= sitename %>">
		<% else %>
			<% drawSelectboxMultiSiteSitename "sitename", sitename, "" %>
		<% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="250">제휴몰ID</td>
    <td align="left">
		<input type="text" name="makerid" value="<%= makerid %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="250">대표언어</td>
    <td align="left">
		<% drawSelectboxMultiLangCountrycd "countrylangcd", countrylangcd, "" %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="250">화폐단위</td>
    <td align="left">
    	<%
    	'//수정모드
    	if IDX <> "" then
    	%>
    		<%= currencyUnit %>
    		<input type="hidden" name="currencyUnit" value="<%= currencyUnit %>">
		<% else %>
    		<input type="text" name="currencyUnit" value="<%= currencyUnit %>">		EX) USD
		<% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td>화폐기호</td>
    <td align="left">
        <input type="text" name="currencyChar" value="<%= currencyChar %>" maxlength="10" size="10">
        EX) $
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td>환율</td>
    <td align="left">
        <input type="text" name="exchangeRate" value="<%= exchangeRate %>" maxlength="10" size="10">
        EX) 1200
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="250">대표배수</td>
    <td align="left">
        <select name="linkPriceType" class="select">
            <option value="">선택
            <option value="1" <%=CHKIIF(linkPriceType=1,"selected","") %> >실판매가
            <option value="2" <%=CHKIIF(linkPriceType=2,"selected","") %> >소비자가
        </select>
        대비
		<input type="text" name="multipleRate" value="<%=multipleRate%>" size="5" maxlength="5"> 배
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td>기준일</td>
    <td align="left">
		<input type="text" class="text" name="basedate" value="<%= basedate %>" size=10 maxlength=10 readonly ><a href="#" onclick="calendarOpen(frmcurrencyUnit.basedate); return false;">
		<img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>
    	<%
    	'//수정모드
    	if IDX <> "" then
    	%>
    		<% if sitename="WSLWEB" then %>
				상품단 환율,배수 일괄적용 :
				<select name="itemautoyn" class="select">
					<option value="Y">Y</option>
					<option value="N">N</option>
				</select>
				<br><br>
			<% end if %>

			<input type="button" value="삭제" onClick="delcurrencyUnit(frmcurrencyUnit);" class="button">
		<% end if %>
	</td>
    <td>
    	<input type="button" value="저장" onClick="SavecurrencyUnit(frmcurrencyUnit);" class="button">
    </td>
</tr>
</table>
</form>

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    	<input type="button" onclick="newcurrencyUnit();" value="신규등록" class="button">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= oexchangerateList.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oexchangerateList.FTotalPage %></b>
	</td>
</tr>
<% if oexchangerateList.FResultCount > 0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>사이트구분</td>
	<td>제휴몰ID</td>
	<td>대표언어</td>
	<td>화폐단위</td>
    <td>화폐기호</td>
    <td>환율</td>
	<td>대표배수</td>
    <td>기준일</td>
    <td>비고</td>
</tr>
<% for i=0 to oexchangerateList.FResultCount-1 %>

<% if oexchangerateList.FItemList(i).fidx = idx then %>
	<tr bgcolor="orange" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='orange'; align="center">
<% else %>
	<tr bgcolor="#ffffff" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff'; align="center">
<% end if %>
	<td><%= oexchangerateList.FItemList(i).fidx %></td>
	<td><%= oexchangerateList.FItemList(i).fsitename %></td>
	<td><%= oexchangerateList.FItemList(i).FMakerid %></td>
	<td><%= oexchangerateList.FItemList(i).fcountryLangCD %></td>
    <td><%= oexchangerateList.FItemList(i).fcurrencyUnit %></td>
    <td><%= oexchangerateList.FItemList(i).fcurrencychar %></td>
    <td align="right"><%= oexchangerateList.FItemList(i).fexchangeRate %></td>
	<td>
		<%= oexchangerateList.FItemList(i).getlinkPriceTypeName %> 대비 <%= oexchangerateList.FItemList(i).fmultipleRate %> 배
	</td>
    <td><%= oexchangerateList.FItemList(i).fbasedate %></td>
    <td width=60>
    	<input type="button" onclick="location.href='?idx=<%= oexchangerateList.FItemList(i).fidx %>&page=<%= page %>'" value="수정" class="button">
    </td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	    <% if oexchangerateList.HasPreScroll then %>
			<a href="?page=<%= oexchangerateList.StartScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oexchangerateList.StartScrollPage to oexchangerateList.FScrollCount + oexchangerateList.StartScrollPage - 1 %>
			<% if i>oexchangerateList.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oexchangerateList.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
<tr bgcolor="#FFFFFF">
	<td align="center">내용이 없습니다.</td>
</tr>
<% end if %>
</table>

</body>
</html>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<%
	set oexchangerate = Nothing
	set oexchangerateList = Nothing
	session.codePage = 949
%>
<!-- 표 하단바 끝-->
<!-- #include virtual="/lib/db/dbclose.asp" -->