<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 이벤트 상품추가
' History : 2010.03.10 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<%
'변수선언
Dim evt_code , i, page ,evt_kind,evt_kinddesc,evt_name,	evt_startdate,evt_enddate
dim evt_prizedate, evt_state, regdate , selDate , strTxt , sCategory ,strparm
dim cEvtCont , cEvtItem, evt_statedesc , PriceDiffExists
	evt_code = requestCheckVar(Request("evt_code"),10)
	selDate 		= requestCheckVar(Request("selDate"),1)  	'기간
	evt_startdate 	= requestCheckVar(Request("evt_startdate"),10)
	evt_enddate 	= requestCheckVar(Request("evt_enddate"),10)
	evt_code 		= requestCheckVar(Request("evt_code"),10)  	'이벤트 코드/명 검색
	strTxt 		= requestCheckVar(Request("sEtxt"),60)
	sCategory	= requestCheckVar(Request("selC"),10) '카테고리
	evt_state	= requestCheckVar(Request("evt_state"),4)	'이벤트 상태
	evt_kind 	= requestCheckVar(Request("evt_kind"),32)	'이벤트종류
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)

if page = "" then page = 1

IF evt_code = "" THEN	'이벤트 코드값이 없을 경우 back
%>
	<script language="javascript">
		alert("전달값에 문제가 발생하였습니다. 관리자에게 문의해주십시오");
		history.back();
	</script>
<%	dbget.close()	:	response.End
END IF

strparm = "menupos="&menupos&"&selDate="&selDate&"&evt_startdate="&evt_startdate&"&evt_enddate="&evt_enddate
strparm = strparm & "&sEtxt="&strTxt&"&selC="&sCategory&"&evt_state="&evt_state&"&evt_kind="&evt_kind&""

'--이벤트 개요
set cEvtCont = new cevent_list
	cEvtCont.Frectevt_code = evt_code

	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont_off
	evt_kind 		=	cEvtCont.FOneItem.fevt_kind
	evt_kinddesc	=	cEvtCont.FOneItem.fevt_kinddesc
	evt_name = cEvtCont.FOneItem.fevt_name
	evt_startdate = cEvtCont.FOneItem.Fevt_startdate
	evt_enddate = cEvtCont.FOneItem.Fevt_enddate
	evt_prizedate =	cEvtCont.FOneItem.Fevt_prizedate
	evt_state =	cEvtCont.FOneItem.Fevt_state
	evt_statedesc =	cEvtCont.FOneItem.fevt_statedesc
	regdate	= cEvtCont.FOneItem.fevt_regdate
set cEvtCont = nothing

set cEvtItem = new cevent_list
	cEvtItem.FPageSize = 100
	cEvtItem.FCurrPage = page
	cEvtItem.Frectevt_code = evt_code
	cEvtItem.fnGetEventItem
%>

<script language="javascript">

function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

// 새상품 추가 팝업
function addnewItem(){
	var popwin;
	popwin = window.open("/common/offshop/pop_event_additemlist_off.asp?evt_code=<%=evt_code%>", "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}

//선택상품 삭제
function itemdel(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "," ;
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "," ;
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "," ;

				}
			}
		}

	upfrm.mode.value='itemdel';
	upfrm.action = "eventitem_off_process.asp";
	upfrm.submit();
	upfrm.itemidarr.value = ""
	upfrm.itemoptionarr.value = ""
	upfrm.itemgubunarr.value = ""
}

</script>

<!-- 표 상단바 시작-->
<font color="red">※ 이벤트 상품 추가</font>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<input type="hidden" name="sType" >
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemoptionarr">
<input type="hidden" name="itemgubunarr">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr>
	<td style="padding-bottom:10">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
				<td width="30%" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_code%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트명</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_name%></td>
			</tr>
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">종류</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_kinddesc%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">상태</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_statedesc%></td>
			</tr>
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">기간</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_startdate%> ~ <%=evt_enddate%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨 발표일</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%
					if evt_prizedate > "1900-01-01" then
					response.write evt_prizedate
					end if
					%>
				</td>
			</tr>
		</table>
	</td>
</tr>
</form>

<tr>
	<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
		    <tr height="35">
		        <td align="left">
		       		<input type="button" value="선택삭제" onClick="itemdel(frm);" class="button">
		       		<input type="button" value="이벤트페이지로돌아가기" onClick="location.href='index.asp?evt_code=<%=evt_code%>&menupos=<%=menupos%>';" class="button">
		    	<td align="right">
		       		<input type="button" value="새상품 추가" onclick="addnewItem();" class="button">
		        </td>
		    </tr>
		</table>
	</td>
</tr>

<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan="20" align="left">검색결과 : <b><%=cEvtItem.FTotalCount%></b>&nbsp;&nbsp;페이지 : <b><%= page %>/ <%= cEvtItem.FTotalPage %></b></td>
		</tr>
		<% if cEvtItem.FTotalCount > 0 then %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
			<td width="70">브랜드ID</td>
			<td width="90">상품코드</td>
			<td>상품명<font color="blue">[옵션명]</font></td>
			<td width="50">소비자가</td>
			<td width="50">판매가</td>
			<td width="40">할인율<br>(%)</td>
			<td width="50">매입가</td>
			<td width="50">샾공급가</td>
			<td width="30">매입<br>마진</td>
			<td width="30">공급<br>마진</td>
			<td width="30">센터<br>매입<br>구분</td>
			<td width="30">ON<br>판매</td>
			<td width="30">ON<br>단종</td>
			<td width="60">범용바코드</td>
			<td width="60">비고</td>
		</tr>
		<%
			For i = 0 to cEvtItem.fresultcount - 1
		%>
		<form action="" name="frmBuyPrc<%=i%>" method="get">
		<% if cEvtItem.FItemlist(i).Fisusing="N" then %>
		<tr bgcolor="#EEEEEE">
		<% else %>
		<tr bgcolor="#FFFFFF">
		<% end if %>
			<td  align="center">
				<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
				<input type="hidden" name="itemid" value="<%=cEvtItem.FItemlist(i).Fshopitemid%>">
				<input type="hidden" name="itemoption" value="<%=cEvtItem.FItemlist(i).Fitemoption%>">
				<input type="hidden" name="itemgubun" value="<%=cEvtItem.FItemlist(i).fitemgubun%>">
			</td>
	 		<td ><%= cEvtItem.FItemlist(i).FMakerID %></td>
	  		<td><%= cEvtItem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(cEvtItem.FItemlist(i).Fshopitemid>=1000000,Format00(8,cEvtItem.FItemlist(i).Fshopitemid),Format00(6,cEvtItem.FItemlist(i).Fshopitemid)) %>-<%= cEvtItem.FItemlist(i).Fitemoption %></td>
	  		<td>
	  			<%= cEvtItem.FItemlist(i).FShopItemName %>
	  			<% if cEvtItem.FItemlist(i).Fitemoption<>"0000" then %>
	  				<font color="blue">[<%= cEvtItem.FItemlist(i).FShopitemOptionname %>]</font>
	  			<% end if %>

	  			<% if cEvtItem.FItemlist(i).FOnlineOptaddprice<>0 then %>
	  			    <br>옵션추가금액: <%= FormatNumber(cEvtItem.FItemlist(i).FOnlineOptaddprice,0) %>
	  			<% end if %>
	  		</td>
	        <% PriceDiffExists = false %>
	        <td align="right" >
	            <%= FormatNumber(cEvtItem.FItemlist(i).FShopItemOrgprice,0) %>
	            <% if (cEvtItem.FItemlist(i).FItemGubun="10") then %>
	            <% if (cEvtItem.FItemlist(i).FOnlineitemorgprice + cEvtItem.FItemlist(i).FOnlineOptaddprice<>cEvtItem.FItemlist(i).FShopItemOrgprice)  then %>
	                <font color="red"><strong><%= cEvtItem.FItemlist(i).FOnlineitemorgprice + cEvtItem.FItemlist(i).FOnlineOptaddprice %></strong></font>
	            <% else %>
	                <% if (PriceDiffExists) then %>
	                <%= cEvtItem.FItemlist(i).FOnlineitemorgprice + cEvtItem.FItemlist(i).FOnlineOptaddprice %>
	                <% end if %>
	            <% end if %>
	            <% end if %>
	        </td>
	  		<td align="right" >
	  		    <%= FormatNumber(cEvtItem.FItemlist(i).FShopItemprice,0) %>
	  		    <% if (cEvtItem.FItemlist(i).FItemGubun="10") then %>
	            <% if (cEvtItem.FItemlist(i).FOnLineItemprice+ cEvtItem.FItemlist(i).FOnlineOptaddprice<>cEvtItem.FItemlist(i).FShopItemprice)  then %>
	  		        <font color="red"><strong><%= cEvtItem.FItemlist(i).FOnLineItemprice + cEvtItem.FItemlist(i).FOnlineOptaddprice %></strong></font>
	  		    <% else %>
	  		        <% if (PriceDiffExists) then %>
	  		        <%= cEvtItem.FItemlist(i).FOnLineItemprice + cEvtItem.FItemlist(i).FOnlineOptaddprice %>
	  		        <% end if %>
	            <% end if %>
	            <% end if %>
	  		</td>
	  		<td align="center" >
	            <% if (cEvtItem.FItemlist(i).FShopItemOrgprice<>0) then %>
	                <% if cEvtItem.FItemlist(i).FShopItemOrgprice<>cEvtItem.FItemlist(i).FShopItemprice then %>
	                OFF:<font color="#FF3333"><%= CLng((cEvtItem.FItemlist(i).FShopItemOrgprice-cEvtItem.FItemlist(i).FShopItemprice)/cEvtItem.FItemlist(i).FShopItemOrgprice*100*100)/100 %></font>
	                <% end if %>
	  		    <% end if %>

	  		    <% if (cEvtItem.FItemlist(i).FOnlineitemorgprice<>0) then %>
	  		        <% if cEvtItem.FItemlist(i).FOnlineitemorgprice<>cEvtItem.FItemlist(i).FOnLineItemprice then %>
	                ON:<font color="#FF3333"><%= CLng((cEvtItem.FItemlist(i).FOnlineitemorgprice-cEvtItem.FItemlist(i).FOnLineItemprice)/cEvtItem.FItemlist(i).FOnlineitemorgprice*100*100)/100 %></font>
	                <% end if %>
	  		    <% end if %>
	  		</td>

	  		<td align="right" ><%= FormatNumber(cEvtItem.FItemlist(i).Fshopsuplycash,0) %></td>
	  		<td align="right" ><%= FormatNumber(cEvtItem.FItemlist(i).Fshopbuyprice,0) %></td>

	  		<td align="right" >
	  		<% if (cEvtItem.FItemlist(i).FShopItemprice<>0) and (cEvtItem.FItemlist(i).Fshopsuplycash<>0) then %>
	  			<font color="blue"><%= CLng((cEvtItem.FItemlist(i).FShopItemprice-cEvtItem.FItemlist(i).Fshopsuplycash)/cEvtItem.FItemlist(i).FShopItemprice*100) %>%</font>
	  		<% end if %>
	  		</td>
	  		<td align="right" >
	  		<% if (cEvtItem.FItemlist(i).FShopItemprice<>0) and (cEvtItem.FItemlist(i).Fshopbuyprice<>0) then %>
	  			<font color="blue"><%= CLng((cEvtItem.FItemlist(i).FShopItemprice-cEvtItem.FItemlist(i).Fshopbuyprice)/cEvtItem.FItemlist(i).FShopItemprice*100) %>%</font>
	  		<% end if %>
	  	    </td>
	  	    <td align="center" ><%= cEvtItem.FItemlist(i).FCenterMwDiv %></td>
	  	    <td align="center" ><%= fnColor(cEvtItem.FItemlist(i).Fsellyn,"sellyn") %></td>
	  	    <td align="center" ><%= fnColor(cEvtItem.FItemlist(i).FonLineDanjongyn,"dj") %></td>
	  		<td align="right" ><%= cEvtItem.FItemlist(i).FextBarcode %></td>
	  		<td align="right" ></td>
		</tr>
		</form>
		<% Next %>

		<tr height="25" bgcolor="FFFFFF">
			<td colspan="18" align="center">
		       	<% if cEvtItem.HasPreScroll then %>
					<span class="list_link"><a href="?<%=strparm%>&page=<%=cEvtItem.StartScrollPage-1%>&evt_code=<%=evt_code%>">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
				<% for i = 0 + cEvtItem.StartScrollPage to cEvtItem.StartScrollPage + cEvtItem.FScrollCount - 1 %>
					<% if (i > cEvtItem.FTotalpage) then Exit for %>
					<% if CStr(i) = CStr(cEvtItem.FCurrPage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
					<% else %>
					<a href="?<%=strparm%>&page=<%=i%>&evt_code=<%=evt_code%>" class="list_link"><font color="#000000"><%= i %></font></a>
					<% end if %>
				<% next %>
				<% if cEvtItem.HasNextScroll then %>
					<span class="list_link"><a href="?<%=strparm%>&page=<%=i%>&evt_code=<%=evt_code%>">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>

		<% ELSE %>
		<tr align="center" bgcolor="#FFFFFF">
			<td colspan="20">등록된 내용이 없습니다.</td>
		</tr>
		<%END IF%>
		</table>
	</td>
</tR>
</table>
<%
	set cEvtItem = nothing
%>

<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
