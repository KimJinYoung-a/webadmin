<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  메이크글로비 판매대기상품
' History : 2015.10.28 원승현 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/makeglob/makeglobCls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->

<%

	Dim currpage '// 현재 페이지
	Dim pagesize '// 페이지사이즈
	Dim brandname '// 브랜드명
	Dim itemname '// 상품명
	Dim itemid '// 아이템코드
	Dim sellyn '// 상품판매여부
	Dim limityn '// 한정판매여부
	Dim isusing '// 사용여부
	Dim MakeGlobChkEN '// 영문입력여부
	Dim MakeGlobChkZH '// 중문입력여부
	Dim ghidden '// 글로비 숨김여부
	Dim gsoldout '// 글로비 품절여부
	Dim gproductkey '// 글로비 상품코드
	Dim gcheck '// 글로비 등록여부
	Dim marginSt '// 마진율 시작값
	Dim marginEd '// 마진율 종료값
	Dim sOrgpriceSt '// 판매가 시작값
	Dim sOrgpriceEd '// 판매가 종료값
	Dim baesonggubun '// 배송구분(업배, 텐배)
	Dim i, dispCate, paramvalue



	currpage = request("page")
	pagesize = 30
	brandname = request("brandname")
	itemname = request("itemname")
	itemid = request("itemid")
	sellyn = request("sellyn")
	limityn = request("limityn")
	isusing = request("isusing")
	ghidden = request("globHiddenYN")
	gsoldout = request("globSoldoutYN")
	gproductkey = request("gproductkey")
	gcheck = request("globCheckYN")
	marginSt = request("marginSt")
	marginEd = request("marginEd")
	sOrgpriceSt = request("sOrgpriceSt")
	sOrgpriceEd = request("sOrgpriceEd")
	MakeGlobChkEN = request("MakeGlobChkEN")
	MakeGlobChkZH = request("MakeGlobChkZH")
	baesonggubun = request("baesonggubun")
	dispCate = request("disp")



	'// 기본값
	If currpage = "" Then currpage = 1
	If sellyn = "" Then sellyn = "Y"
	If isusing = "" Then isusing = "Y"

	if itemid<>"" then
		dim iA ,arrTemp,arrItemid

		itemid = replace(itemid,chr(10),",")

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

	if gproductkey<>"" then
		dim iB ,arrTemp2,arrgproductkey

		arrTemp2 = Split(gproductkey,",")

		iB = 0
		do while iB <= ubound(arrTemp2)
			if Trim(arrTemp2(iB))<>"" and isNumeric(Trim(arrTemp2(iB))) then
				arrgproductkey = arrgproductkey & Trim(arrTemp2(iB)) & ","
			end if
			iB = iB + 1
		loop

		if len(arrgproductkey)>0 then
			gproductkey = left(arrgproductkey,len(arrgproductkey)-1)
		else
			if Not(isNumeric(gproductkey)) then
				gproductkey = ""
			end if
		end if
	end if


	Dim oitem
	set oitem = new CMakeGlobItem
	oitem.Fpagesize = pagesize
	oitem.Fcurrpage = currpage
	oitem.FRectBrandName = brandname
	oitem.FRectCateCode = dispCate
	oitem.FRectItemName = itemname
	oitem.FRectItemId = itemid
	oitem.FRectSellyn = sellyn
	oitem.FRectLimityn = limityn
	oitem.FRectIsUsing = isusing
	oitem.FRectGIsHidden = ghidden
	oitem.FRectGIssoldout = gsoldout
	oitem.FRectGProductKey = gproductkey
	oitem.FRectGIscheck = gcheck
	oitem.FRectMarginSt = marginSt
	oitem.FRectMarginEd = marginEd
	oitem.FRectSorgpriceSt = sOrgpriceSt
	oitem.FRectSorgpriceEd = sOrgpriceEd
	oitem.FRectBaesongGubun = baesonggubun
	oitem.GetMakeGlobItemWaitingList()

	paramvalue = "menupos=3751&page="&currpage&"&reload=ON&disp="&dispcate&"&itemname="&itemname&"&itemid="&itemid&"&sellyn="&sellyn&"&isusing="&isusing&"&limityn="&limityn&"&gproductkey="&gproductkey&"&globHiddenYN="&ghidden&"&globSoldoutYN="&gsoldout&"&globCheckYN="&gcheck&"&brandname="&brandname&"&baesonggubun="&baesonggubun

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

$(document).ready(function(){
    $("#checkall").click(function(){
        if($("#checkall").prop("checked")){
            $("input[name=productcode]").prop("checked",true);
        }else{
            $("input[name=productcode]").prop("checked",false);
        }
    })
})


function fnHiddenProc(val)
{
	var hiddenarrlist='';
	var hiddenalertText='';

	if (val=="Y")
	{
		hiddenalertText = "선택된 상품을 숨김처리 하시겠습니까?";
	}
	else
	{
		hiddenalertText = "선택된 상품을 노출 하시겠습니까?";
	}

	if (!$('input:checkbox[name=productcode]').is(':checked'))
	{
		alert("상품을 선택해주세요.");
		return false;
	}
	else
	{
		if (confirm(hiddenalertText))
		{
			document.globFrm.mode.value="hidden";
			document.globFrm.hiddenvalue.value=val;
			$("input:checkbox[name=productcode]:checked").each(function(){
				if (hiddenarrlist=="")
				{
					hiddenarrlist=$(this).val();
				}
				else
				{
					hiddenarrlist+=','+$(this).val();
				}
			});
			document.globFrm.arrproductcode.value=hiddenarrlist;
			document.globFrm.submit();
		}
		else
		{
			return false;
		}
	}


}

function fnSoldoutProc(val)
{
	var soldarrlist='';
	var soldalertText='';

	if (val=="Y")
	{
		soldalertText = "선택된 상품을 품절처리 하시겠습니까?";
	}
	else
	{
		soldalertText = "선택된 상품을 판매가능 상태로 변경하시겠습니까?";
	}

	if (!$('input:checkbox[name=productcode]').is(':checked'))
	{
		alert("상품을 선택해주세요.");
		return false;
	}
	else
	{
		if (confirm(soldalertText))
		{
			document.globFrm.mode.value="soldout";
			document.globFrm.soldoutvalue.value=val;
			$("input:checkbox[name=productcode]:checked").each(function(){
				if (soldarrlist=="")
				{
					soldarrlist=$(this).val();
				}
				else
				{
					soldarrlist+=','+$(this).val();
				}
			});
			document.globFrm.arrproductcode.value=soldarrlist;
			document.globFrm.submit();
		}
		else
		{
			return false;
		}
	}


}

function fnProductInsert()
{
	var productarrlist='';
	if (!$('input:checkbox[name=productcode]').is(':checked'))
	{
		alert("상품을 선택해주세요.");
		return false;
	}
	else
	{
		if (confirm('선택하신 상품을 등록/수정 하시겠습니까?'))
		{
			document.globFrm.mode.value="product";
			$("input:checkbox[name=productcode]:checked").each(function(){
				if (productarrlist=="")
				{
					productarrlist=$(this).val();
				}
				else
				{
					productarrlist+=','+$(this).val();
				}
			});
			document.globFrm.arrproductcode.value=productarrlist;
			document.globFrm.submit();
		}
		else
		{
			return false;
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
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : 	<input type="text" class="text" name="brandname" value="<%= brandname %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;&nbsp;
		전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		&nbsp;&nbsp;
		* 상품명 :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick='NextPage("");'>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		* 텐바이텐 판매여부:<% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;&nbsp;
     	* 텐바이텐 사용여부:<% drawSelectBoxUsingYN "isusing", isusing %>
		&nbsp;&nbsp;
     	* 텐바이텐 한정여부:<% drawSelectBoxLimitYN "limityn", limityn %>
		&nbsp;&nbsp;
     	* 배송구분: 
		<select name="baesonggubun" class="select" >
			<option value="">전체</option>
			<option value="tenbae" <% If baesonggubun="tenbae" Then %> selected <% End If %>>텐바이텐배송</option>
			<option value="upbae" <% If baesonggubun="upbae" Then %> selected <% End If %>>업체배송</option>
		</select>
		&nbsp;&nbsp;
		<p/>
     	* 마진율 : <input type="text" class="text" name="marginSt" value="<%= marginSt %>" size="10" maxlength="4"> ~ <input type="text" class="text" name="marginEd" value="<%= marginEd %>" size="10" maxlength="4">
		&nbsp;&nbsp;
     	* 판매가 : <input type="text" class="text" name="sOrgPriceSt" value="<%= sOrgPriceSt %>" size="10" maxlength="10"> ~ <input type="text" class="text" name="sOrgPriceEd" value="<%= sOrgPriceEd %>" size="10" maxlength="10">
		&nbsp;&nbsp;
		* 텐바이텐 상품코드 :
		<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		* 글로비 상품코드 :
		<input type="text" class="text" name="gproductkey" value="<%= gproductkey %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
		&nbsp;&nbsp;
		* 글로비 숨김여부:<% drawSelectBoxGHiddenYN "globHiddenYN", ghidden %>
		&nbsp;&nbsp;
     	* 글로비 품절여부:<% drawSelectBoxGsoldoutYN "globSoldoutYN", gsoldout %>
		&nbsp;&nbsp;
     	* 글로비 등록여부:<% drawSelectBoxGcheckYN "globCheckYN", gcheck %>
		&nbsp;&nbsp;
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="3" width="50" align="center"><strong>관리</strong></td>
		<td><input type="button" value="상품숨김" onclick="fnHiddenProc('Y');return false;">&nbsp;&nbsp;<input type="button" value="상품노출" onclick="fnHiddenProc('N');return false;"></td>
	</tr>
	<tr align="left" bgcolor="<%= adminColor("topbar") %>" >
		<td><input type="button" value="품절처리" onclick="fnSoldoutProc('Y');return false;">&nbsp;&nbsp;<input type="button" value="판매가능" onclick="fnSoldoutProc('N');return false;"></td>
	</tr>
	<tr align="left" bgcolor="<%= adminColor("topbar") %>" >
		<td><input type="button" value="상품등록/수정" onclick="fnProductInsert();return false;"> (기존에 이미 등록되어 있던 상품은 최신정보로 수정, 없던 상품은 신규로 추가 됩니다.)</td>
	</tr>
</table>
<br>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				검색결과 : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= currpage %> /<%=  oitem.FTotalpage %></b>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50" rowspan="2"><input type="checkbox" id="checkall"></td>
	<td width="50" rowspan="2">이미지</td>
	<td width="100" rowspan="2">브랜드명</td>
	<td rowspan="2">상품명</td>
	<td width="60" rowspan="2">상품<br>무게</td>
	<td width="60" rowspan="2">배송<br>구분</td>
	<td colspan="7" width="300"><strong>텐바이텐</strong></td>
	<td colspan="7" width="120"><strong>메이크글로비</strong></td>

</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

	<td width="60">상품코드</td>
	<td width="60">판매가</td>
	<td width="60">매입가</td>
	<td width="60">마진율</td>
	<td width="30">판매<br>여부</td>
	<td width="30">품절<br>여부</td>
	<td width="30">사용<br>여부</td>
	<td width="30">한정<br>여부</td>
	<td width="60">상품코드</td>
	<td width="30">숨김<br>여부</td>
	<td width="30">품절<br>여부</td>
	<td width="60">업데이트<br>여부</td>
	<td width="60">업데이트<br>일자</td>
</tr>


<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="19" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" <% If oitem.FItemList(i).FMakeGlobProductKey="" Or isnull(oitem.FItemList(i).FMakeGlobProductKey) Then %> bgcolor="#FFFFA5" <% Else %> bgcolor="#FFFFFF" <% End If %>align="center">

	<td align="center"><input type="checkbox" name="productcode" value="<%= oitem.FItemList(i).Fitemid %>"></td>
	<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
	<td align="left"><%= oitem.FItemList(i).FBrandName %></td>
	<td align="left"><% =oitem.FItemList(i).FitemName %></td>
	<td align="center"><%= FormatNumber((oitem.FItemList(i).FitemWeight/1000),2) %>kg</td>
	<td align="center">
		<%
			If oitem.FItemList(i).FBaesongGubun="M" Or oitem.FItemList(i).FBaesongGubun="W" Then
				Response.write "텐배"
			Else
				Response.write "업배"
			End If
		%>
	</td>
	<td>
		<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">
		<%= oitem.FItemList(i).Fitemid %></a>
	</td>
	<td align="right">
	<%
		Response.Write "" & FormatNumber(oitem.FItemList(i).Forgprice,0) & ""
		'할인가
'		if oitem.FItemList(i).Fsailyn="Y" then
'			Response.Write "<br><font color=#F08050>(현판매가)" & FormatNumber(oitem.FItemList(i).FsellCash,0) & "</font>"
'		end if

	%>
	</td>
	<td align="right">
	<%
		Response.Write "" & FormatNumber(oitem.FItemList(i).Forgsuplycash,0) & ""
	%>
	</td>
	<td align="right">
	<%
		Response.Write "" & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice,1) & ""
	%>
	</td>
	<!--td align="center"><%= FormatNumber(oitem.FItemList(i).FbuyCash,0) %></td-->
	<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
	<td align="center">
		<%
			If oitem.FItemList(i).isSoldout Then
				Response.write fnColor("Y", "yn")
			Else
				Response.write fnColor("N", "yn")
			End If
		%>
	</td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td><%= oitem.FItemList(i).FMakeGlobProductKey %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).FMakeGlobHidden,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).FMakeGlobSoldout,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).FMakeGlobupdate,"yn") %></td>
	<td align="center">
		<%
			If oitem.FItemList(i).FMakeGlobupdateTime ="1900-01-01" Then
				Response.write ""
			Else
				Response.write oitem.FItemList(i).FMakeGlobupdateTime
			End If
		%>
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="19" align="center">
		<% if oitem.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(currpage)=CStr(i) then %>
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

</table>
<form method="post" action="/admin/makeglob/proc.asp" name="globFrm">
	<input type="hidden" name="mode">
	<input type="hidden" name="hiddenvalue">
	<input type="hidden" name="soldoutvalue">
	<input type="hidden" name="arrproductcode">
	<input type="hidden" name="paramvalue" value="<%=tenEnc(paramvalue)%>">
</form>
<% end if %>

<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->