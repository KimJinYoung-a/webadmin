<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  온라인 해외판매대기상품
' History : 2013.05.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->

<%
response.write "사용중지 매뉴 입니다."
response.end

dim itemid, itemname, makerid, sellyn, usingyn, mwdiv, limityn, overSeaYn, weightYn, danjongyn
dim cdl, cdm, cds, sortDiv, sortDiv2, sellcash1, sellcash2, vDate1, vDate2, page, i
dim itemrackcode, vRegUserID, vIsReg, reload
dim sitename
	itemid		= request("itemid")
	itemname	= request("itemname")
	makerid		= request("makerid")
	sellyn		= request("sellyn")
	usingyn		= request("usingyn")
	mwdiv		= request("mwdiv")
	limityn		= request("limityn")
	overSeaYn	= request("overSeaYn")
	weightYn	= request("weightYn")
	itemrackcode= request("itemrackcode")
	sortDiv		= request("sortDiv")
	sortDiv2	= request("sortDiv2")
	vRegUserID	= request("reguserid")
	vIsReg		= request("isreg")
	sellcash1	= request("sellcash1")
	sellcash2	= request("sellcash2")
	vDate1		= request("date1")
	vDate2		= request("date2")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	page = request("page")
	reload = request("reload")
    sitename = request("sitename")
	danjongyn   = requestCheckvar(request("danjongyn"),10)

'기본값
if (page="") then page=1
if sitename="" then sitename="WSLWEB"
'if reload<>"ON" and mwdiv="" then mwdiv="MW"
if reload<>"ON" and overSeaYn="" then overSeaYn="Y"
'if reload<>"ON" and weightYn="" then weightYn="Y"
if sortDiv="" then sortDiv="new"
if sortDiv2="" then sortDiv2="weightup"
if reload<>"ON" and sellyn="" then sellyn="YS"
if reload<>"ON" and usingyn="" then usingyn="Y"
if vIsReg="" then vIsReg="x"

if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim oitem
set oitem = new COverSeasItem
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.FRectMakerid      = makerid
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectIsOversea	= overSeaYn
	oitem.FRectIsWeight		= weightYn
	oitem.FRectRackcode		= itemrackcode
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectSortDiv		= sortDiv
	oitem.FRectSortDiv2		= sortDiv2
	oitem.FRectRegUserID	= vRegUserID
	oitem.FRectRegDate1		= vDate1
	oitem.FRectRegDate2		= vDate2
	oitem.FRectIsReg		= vIsReg
	oitem.FRectSellcash1	= sellcash1
	oitem.FRectSellcash2	= sellcash2
	oitem.FRectSitename     = sitename
	oitem.FRectDanjongyn    = danjongyn

    if (sitename<>"") then
        oitem.GetOverSeasTargetItemListCommon
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('사이트를 선택하세요');"
		response.write "</script>"
    end if
%>

<script type='text/javascript'>

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function chgSort(srt,gb){
	if(gb == "2"){
		document.frm.sortDiv2.value= srt;
	}else{
		document.frm.sortDiv.value= srt;
	}
	document.frm.submit();
}

function chgReg(reg){
	document.frm.isreg.value= reg;
	document.frm.submit();
}

function PopItemContent(iitemid){
	var popwin = window.open('/admin/itemmaster/overseas/popItemContent.asp?itemid=' + iitemid +'&sitename=<%=sitename%>','itemWeightEdit','width=1280,height=960,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function num_check(gb){
	if(gb == "1"){
		if(isNaN(document.frm.sellcash1.value) == true)
		{
			alert("숫자만 입력해주세요.");
			document.frm.sellcash1.value = "";
			document.frm.sellcash1.focus();
		}
	}else{
		if(isNaN(document.frm.sellcash2.value) == true)
		{
			alert("숫자만 입력해주세요.");
			document.frm.sellcash2.value = "";
			document.frm.sellcash2.focus();
		}
	}
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<input type="hidden" name="reload" value="ON">
<input type="hidden" name="sortDiv" value="<%=sortDiv%>">
<input type="hidden" name="sortDiv2" value="<%=sortDiv2%>">
<input type="hidden" name="isreg" value="<%=vIsReg%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : <%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<p>
		* 렉코드 :
		<input type="text" class="text" name="itemrackcode" value="<%= itemrackcode %>" size="12" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;&nbsp;		
		* 상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
		&nbsp;&nbsp;
		* 상품명 :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick='NextPage("");'>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		* 판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;&nbsp;
     	* 사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;&nbsp;
     	* 한정:<% drawSelectBoxLimitYN "limityn", limityn %>
		&nbsp;&nbsp;
		* 단종: <% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
		&nbsp;&nbsp;
     	* 거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;&nbsp;
     	* 해외배송
		<select class="select" name="overSeaYn">
		<option value="">전체</option>
		<option value="Y" <% if overSeaYn="Y" then Response.write "selected"%>>사용</option>
		<option value="N" <% if overSeaYn="N" then Response.write "selected"%>>안함</option>
		</select>
		&nbsp;&nbsp;
     	* 무게여부
		<select class="select" name="weightYn">
		<option value="">전체</option>
		<option value="Y" <% if weightYn="Y" then Response.write "selected"%>>사용</option>
		<option value="N" <% if weightYn="N" then Response.write "selected"%>>안함</option>
		</select>
     	<br>
     	* 판매가 :
     	<input type="text" class="text" name="sellcash1" value="<%=sellcash1%>" size="10" onkeyUp="num_check('1')">
     	~<input type="text" class="text" name="sellcash2" value="<%=sellcash2%>" size="10" onkeyUp="num_check('2')">
		&nbsp;&nbsp;
		* 최종업데이트일:
		<input type="text" name="date1" size="10" maxlength=10 readonly value="<%= vDate1 %>">
		<a href="javascript:calendarOpen(frm.date1);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;~&nbsp;
		<input type="text" name="date2" size="10" maxlength=10 readonly value="<%= vDate2 %>">
		<a href="javascript:calendarOpen(frm.date2);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;&nbsp;
     	* 등록자
     	<select class="select" name="reguserid">
     	<option value="">전체</option>
     	<option value="gkclzh" <% if vRegUserID="gkclzh" then Response.write "selected"%>>박선영(gkclzh)</option>
     	<option value="grim0307" <% if vRegUserID="grim0307" then Response.write "selected"%>>강그림(grim0307)</option>
     	<option value="alsdud001919" <% if vRegUserID="alsdud001919" then Response.write "selected"%>>연민영(alsdud001919)</option>
     	<option value="">------------------</option>
     	<%
     		Dim vQuery

			vQuery = "SELECT" & vbcrlf
			vQuery = vQuery & " userid, part_sn, username, case part_sn when '11' then 'MD' when '14' then 'MKT' end as part" & vbcrlf
			vQuery = vQuery & " FROM [db_partner].[dbo].[tbl_user_tenbyten]" & vbcrlf
			vQuery = vQuery & " WHERE part_sn IN(11,14) AND isusing = 1" & vbcrlf

			' 퇴사예정자 처리	' 2018.10.16 한용민
			vQuery = vQuery & " and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
			vQuery = vQuery & " AND posit_sn = '12'" & vbcrlf
			vQuery = vQuery & " ORDER BY part_sn ASC, username ASC" & vbcrlf

			'response.write vQuery & "<br>"
     		rsget.Open vQuery,dbget,1
     		Do Until rsget.Eof
				Response.Write "<option value=""" & rsget("userid") & """ "
				If vRegUserID = rsget("userid") Then
					Response.Write " selected"
				End If
				Response.Write ">" & rsget("part") & " - " & rsget("username") & "</option>"
			rsget.MoveNext
			Loop
			rsget.close()
     	%>
     	</select>
		<br>
		<b><font color="blue">
	    * 사이트 : <% drawSelectboxMultiSiteSitename "sitename", sitename, " onchange='NextPage("""");'" %>
	    </font></b>	
	</td>
</tr>
</form>
</table>

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	※ 이매뉴는 신규등록만 가능합니다. 수정은 [ON]해외상품관리>>해외판매상품 에서 하세요.
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				검색결과 : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
			</td>
			<td align="right">
				상품등록여부 :
				<select name="reg" class="select" onchange="chgReg(this.value)">
					<option value="all" <%= CHKIIF(vIsReg="all","selected","") %>>전체보기</option>
					<option value="x" <%= CHKIIF(vIsReg="x","selected","") %>>미등록만</option>
					<option value="o" <%= CHKIIF(vIsReg="o","selected","") %>>등록만</option>
				</select>
				&nbsp;&nbsp;&nbsp;
				정렬방법 :
				1순위 <select name="sort" class="select" onchange="chgSort(this.value,'1')">
					<option value="" <% if sortDiv="" then Response.Write "selected" %>>-선택-</option>
					<option value="new" <% if sortDiv="new" then Response.Write "selected" %>>신상품순</option>
					<option value="best" <% if sortDiv="best" then Response.Write "selected" %>>인기상품순</option>
					<option value="min" <% if sortDiv="min" then Response.Write "selected" %>>낮은가격순</option>
					<option value="hi" <% if sortDiv="hi" then Response.Write "selected" %>>높은가격순</option>
					<option value="hs" <% if sortDiv="hs" then Response.Write "selected" %>>높은할인율순</option>
					<!--<option value="weight" <% if sortDiv="weight" then Response.Write "selected" %>>상품무게순</option>//-->
				</select>
				2순위 <select name="sort2" class="select" onchange="chgSort(this.value,'2')">
					<option value="weightup" <% if sortDiv2="weightup" then Response.Write "selected" %>>상품무게높은순</option>
					<option value="weightdown" <% if sortDiv2="weightdown" then Response.Write "selected" %>>상품무게낮은순</option>
				</select>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">No.</td>
	<td width=50> 이미지</td>
	<td width="100">브랜드ID</td>
	<td> 상품명</td>
	<td width="60">판매가</td>
	<td width="60">매입가</td>
	<td width="30">계약<br>구분</td>
	<td width="30">판매<br>여부</td>
	<td width="30">사용<br>여부</td>
	<td width="30">한정<br>여부</td>
	<td width="50">단종<br>여부</td>
	<td width="40">해외<br>여부</td>
	<td width="60">상품<br>무게</td>
	<td width="100">등록자</td>	
	<td width="100">비고</td>
</tr>

<% if oitem.FresultCount > 0 then %>
	<% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF" align="center">
		<td>
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left"><% =oitem.FItemList(i).Fitemname %></td>
		<td align="right">
		<%
			Response.Write "" & FormatNumber(oitem.FItemList(i).Forgprice,0) & ""
			'할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>(할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'쿠폰가
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						'Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
					Case "2"
						'Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
				end Select
			end if
		%>
		</td>
		<td align="center"><%= FormatNumber(oitem.FItemList(i).Fbuycash,0) %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
		<td align="center">
			<%= fnColor(oitem.FItemList(i).Fdanjongyn,"dj") %>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).FdeliverOverseas,"yn") %></td>
		<td align="center"><%= FormatNumber(oitem.FItemList(i).FitemWeight,0) %>g</td>
	    <td align="center"><%= oitem.FItemList(i).FRegUserID %></td>
	    <td>
	    	<% If oitem.FItemList(i).fsitename<>"" Then %>
	    		<input type="button" onClick="PopItemContent( '<%= oitem.FItemList(i).Fitemid %>');" value="수정" class="button">
	    		<br>
	    		<b>상품등록완료</b>
	    	<% Else %>
	    		<input type="button" onClick="PopItemContent('<%= oitem.FItemList(i).Fitemid %>');" value="상품등록" class="button">
	    		<br>
	    		<font color="red">상품미등록</font>
	    	<% End If %>
	    </td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="25" align="center">
			<% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
				<% if i>oitem.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oitem.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>


<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->