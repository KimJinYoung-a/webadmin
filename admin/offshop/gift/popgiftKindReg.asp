<%@ language=vbscript %>
<% option explicit
	Response.Expires = -1440
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
%>
<%
'####################################################
' Description :  사은품 종류
' History : 2010.03.11 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/offshop/gift/gift_Cls.asp"-->
<%
Dim clsGift, giftkind_name, giftkind_code ,arrList, i
dim itemgubun , itemoption , clsedit , itemid
dim tmp_giftkind_name
	giftkind_name = requestCheckVar(Request("giftkind_name"),40)
	'giftkind_code  = requestCheckVar(Request("giftkind_code"),10)

'특수문자 검색
tmp_giftkind_name = Replace(giftkind_name, "[", "[[]")
tmp_giftkind_name = Replace(tmp_giftkind_name, "%", "[%]")

set clsGift = new cgift_list
	clsGift.FRectItemGubun = "80"			'사은품명검색
	clsGift.FrectsTxt = tmp_giftkind_name

	clsGift.fnGetGiftKind


IF giftkind_name <> "" and clsGift.ftotalcount = 0 and IsNumeric(giftkind_name) THEN
	'사은품명으로 검색이 안되면 사은품코드로 검색
	clsGift.FRectItemGubun = "80"			'사은품코드검색
	clsGift.FrectsTxt = ""
	clsGift.FRectShopItemid = giftkind_name

	clsGift.fnGetGiftKind
end if



%>
<script language="javascript">

	// 검색
	function jsSearch(){
		document.frmSearch.submit();
	}

	// 등록 또는 검색 화면으로 변경
	function jsChangeMode(giftkind_code){
		if (giftkind_code ==""){
		document.frmSearch.giftkind_name.value="";
		}
		document.frmSearch.giftkind_code.value = giftkind_code;
		document.frmSearch.submit();
	}

	// 사은품 종류등록
	function jsSubmitGiftKind(){
		if(!frmGift.giftkind_name.value){
			 alert("사은품종류명을 입력해주세요");
			 frmGift.giftkind_name.focus();
			 return false;
		}
		frmGift.mode.value='giftitemedit';
		frmGift.submit();
	}

	//검색된 사은품종류 적용
	function jsSetGiftKind(giftkind_code, giftkind_name, gift_itemgubun, gift_shopitemid, gift_itemoption){
		opener.document.all.giftkind_code.value = giftkind_code;
		opener.document.all.giftkind_name.value= giftkind_name;

		opener.document.all.gift_itemgubun.value= gift_itemgubun;
		opener.document.all.gift_shopitemid.value= gift_shopitemid;
		opener.document.all.gift_itemoption.value= gift_itemoption;

		window.close();
	}

</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 사은품 등록</div>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="0" >
<form name="frmSearch" method="get" action="popgiftKindReg.asp" >
<input type="hidden" name="giftkind_code" >
<tr>
	<td>사은품명/사은품코드 : <input type="text" name="giftkind_name" size="30" maxlength="60" value="<%=giftkind_name%>">
		<input type="button" class="button" value="검색" onClick="jsSearch();">
	</td>
</form>
</tr>
<tr>
	<td><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="<%= adminColor("tabletop") %>">
			<td align="center">구분</td>
			<td align="center">사은품코드</td>
			<!--
			<td align="center">옵션코드</td>
			-->
			<td align="center">사은품명</td>
			<td align="center">등록일</td>
			<td align="center">비고</td>
		</tr>
		<%
		For i =0 To clsGift.ftotalcount - 1
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= clsGift.FItemList(i).fitemgubun %></td>
			<td align="center"><%= clsGift.FItemList(i).fshopitemid %></td>
			<!--
			<td align="center"><%= clsGift.FItemList(i).fitemoption %></td>
			-->
			<td align="center"><%= clsGift.FItemList(i).fgiftkind_name %></td>
			<td align="center"><%= FormatDate(clsGift.FItemList(i).fregdate,"0000.00.00") %></td>
			<td align="center">
				<input type="button" value="선택" class="button" onClick="jsSetGiftKind('<%= clsGift.FItemList(i).fgiftkind_code %>' ,'<%= clsGift.FItemList(i).fgiftkind_name %>','<%= clsGift.FItemList(i).fitemgubun %>','<%= clsGift.FItemList(i).fshopitemid %>','<%= clsGift.FItemList(i).fitemoption %>');">
			</td>
		</tr>
		<% Next	%>
		</table>
		<br>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<%IF giftkind_name <> "" and clsGift.ftotalcount = 0 THEN %>
		<tr><td colspan="2"  bgcolor="#FFFFFF"><font color="#E08050"><%=giftkind_name%></font>에 해당하는 사은품이 없습니다.</td></tr>
		<% else %>
		<tr><td colspan="2"  bgcolor="#FFFFFF">* 사은품은 최대 30개만 표시됩니다.</td></tr>
		<%END IF%>
		</table>
	</td>
</tr>
</table>

<%
set clsGift = nothing
set clsedit = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->