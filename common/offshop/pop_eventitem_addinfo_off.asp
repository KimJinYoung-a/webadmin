<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2010.12.06 한용민 생성
' Description : 이벤트 관련 상품 추가
'				input - actionURL(db 처리에 필요한 파라미터까지 포함) ex.acURL = "/admin/eventmanage/event/eventitem_process.asp?eC=1234"
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<%
dim designer,page,usingyn ,research,imageview, defaultmargin
dim itemgubun, itemid, itemname ,cdl, cdm, cds ,onexpire ,shopid , strparm
dim i, PriceDiffExists , actionURL ,cEvtItem , eCode , scode ,egcode
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),10)
	usingyn     = RequestCheckVar(request("usingyn"),1)
	research    = RequestCheckVar(request("research"),9)
	imageview   = RequestCheckVar(request("imageview"),9)
	onexpire    = RequestCheckVar(request("onexpire"),9)
	itemgubun   = RequestCheckVar(request("itemgubun"),2)
	itemid      = RequestCheckVar(request("itemid"),9)
	itemname    = RequestCheckVar(request("itemname"),32)
	eCode    = RequestCheckVar(request("ec"),10)
	scode    = RequestCheckVar(request("sc"),10)
	egcode    = RequestCheckVar(request("egC"),10)
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
	shopid    = RequestCheckVar(request("shopid"),10)
	actionURL	= request("acURL")
	defaultmargin = RequestCheckVar(request("defaultmargin"),20)

	if egcode = "" then egcode = 0

	'if sellyn = "" then sellyn ="Y"
	if itemid<>"" then
		dim iA ,arrTemp,arrItemid

		arrTemp = Split(itemid,",")

		iA = 0
		do while iA <= ubound(arrTemp)

			if trim(arrTemp(iA))<>"" then
				'상품코드 유효성 검사(2008.08.04;허진원)
				if Not(isNumeric(trim(arrTemp(iA)))) then
					Response.Write "<script type='text/javascript'>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
					dbget.close()	:	response.End
				else
					arrItemid = arrItemid & trim(arrTemp(iA)) & ","
				end if
			end if
			iA = iA + 1
		loop

		itemid = left(arrItemid,len(arrItemid)-1)
	end if

	if page="" then page=1
	if research<>"on" then usingyn="Y"
	strparm = "designer="&designer&"&usingyn="&usingyn&""
	strparm = strparm & "&research="&research&"&imageview="&imageview&"&onexpire="&onexpire&""
	strparm = strparm & "&itemgubun="&itemgubun&"&itemid="&itemid&"&itemname="&itemname&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds

set cEvtItem = new cevent_list
	cEvtItem.FPageSize = 20
	cEvtItem.FCurrPage = page
	cEvtItem.FRectDesigner = designer
	cEvtItem.FRectOnlyUsing = usingyn
	cEvtItem.FRectItemgubun = itemgubun
	cEvtItem.FRectItemID = itemid
	cEvtItem.FRectItemName = html2db(itemname)
	cEvtItem.FRectCDL = cdl
	cEvtItem.FRectCDM = cdm
	cEvtItem.FRectCDS = cds
	cEvtItem.FRectOnlineExpiredItem = onexpire
	cEvtItem.frectevt_code = ecode
	cEvtItem.fnGetEventItem
%>

<script type='text/javascript'>

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="pop_eventitem_addinfo_off.asp";
	frm.submit();
}

function SelectItems(){
var frm;
var itemcount = 0;
frm = document.frm;

	if(typeof(frm.chkitem) !="undefined"){
		if(!frm.chkitem.length){
			if(!frm.chkitem.checked){
				alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
				return;
			}
			 frm.itemidarr.value = frm.chkitem.value;
			 itemcount = 1;
		}else{
			for(i=0;i<frm.chkitem.length;i++){
				if(frm.chkitem[i].checked) {

					 frm.itemidarr.value = frm.itemidarr.value + frm.chkitem[i].value + ",";
					 frm.itemgubunarr.value = frm.itemgubunarr.value + frm.chkitemgubun[i].value + ",";
					 frm.itemoptionarr.value = frm.itemoptionarr.value + frm.chkitemoption[i].value + ",";

				}
				itemcount = frm.chkitem.length;
			}

			if (frm.itemidarr.value == ""){
				alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
				return;
			}
		}
	}else{
		alert("추가할 상품이 없습니다.");
		return;
	}

	//frm.target = opener.name;
	frm.target = "FrameCKP";
	frm.action = "<%=actionURL%>";
	frm.itemcount.value = itemcount;
	frm.submit();
	frm.itemidarr.value = "";
	frm.itemgubunarr.value = "";
	frm.itemoptionarr.value = "";
	frm.itemcount.value = 0;
	opener.history.go(0);
	//window.close();
}

//전체 선택
function jsChkAll(){
var frm;
frm = document.frm;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
		   	 	frm.chkitem.checked = true;
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					frm.chkitem[i].checked = true;
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}
		}
	  }

	}

}

function reg(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="page" >
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="egC" value="<%=egCode%>">
<input type="hidden" name="itemidarr" >
<input type="hidden" name="itemoptionarr" >
<input type="hidden" name="itemgubunarr" >
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="acURL" value="<%=actionURL%>">
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드 : <% drawSelectBoxDesignerwithName "designer",designer  %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;
		상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
		<div style="font-size:11px; color:gray;padding-left:60px;">(쉼표로 복수입력가능)</div>
	</td>

	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		상품구분:<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
     	&nbsp;
     	오프사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
     	&nbsp;
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
		&nbsp;
		<input type="checkbox" name="onexpire" value="on" <% if onexpire="on" then response.write "checked" %> >ON품절+단종+사용안함(신상품제외)
	</td>
</tr>
</table>

<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
	<tr>
		<td  valign="bottom">
				<input type="button" value="선택상품 추가" onClick="SelectItems()" class="button">
		</td>
	</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr  bgcolor="#FFFFFF">
	<td colspan="20">
	검색결과 : <b><%= cEvtItem.FTotalCount%></b>
	&nbsp;
	페이지 : <b><%= page %> /<%=  cEvtItem.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<% if (imageview<>"") then %>
	<td width="50">이미지</td>
	<% end if %>
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
</tr>
<% if cEvtItem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF" >
    	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if cEvtItem.FresultCount > 0 then %>
    <% for i=0 to cEvtItem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
	<td  align="center">
		<input type="checkbox" name="chkitem" value="<%= cEvtItem.FItemList(i).Fshopitemid %>">
		<input type="hidden" name="chkitemoption" value="<%= cEvtItem.FItemList(i).Fitemoption %>">
		<input type="hidden" name="chkitemgubun" value="<%= cEvtItem.FItemList(i).Fitemgubun %>">
	</td>
	<% if (imageview<>"") then %>
	<td width="50"><img src="<%= cEvtItem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></td>
	<% end if %>
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
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
       	<% if cEvtItem.HasPreScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=cEvtItem.StartScrollPage-1%>);">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + cEvtItem.StartScrollPage to cEvtItem.StartScrollPage + cEvtItem.FScrollCount - 1 %>
			<% if (i > cEvtItem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(cEvtItem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if cEvtItem.HasNextScroll then %>
			<span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</form>
<% end if %>

<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
</table>
<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
	<tr>
		<td  valign="bottom">
				<input type="button" value="선택상품 추가" onClick="SelectItems()" class="button">
		</td>
	</tr>
</table>

<%
	set cEvtItem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->