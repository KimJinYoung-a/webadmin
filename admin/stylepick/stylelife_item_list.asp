<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->
<%
	Dim cdl, cdm, cds, vIsNowCateSearch, vOnly1, vOnly2, vSort1, vSort2, vItemID, vItemName, vMakerID, cStyleLifeItemList, ocate, i, vCount, arrList, intLoop, iCurrentpage, vParam
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	vOnly1 = request("only1")
	vOnly2 = request("only2")
	vSort1 = request("sort1")
	vSort2 = NullFillWith(request("sort2"),"ne")
	vIsNowCateSearch = NullFillWith(request("nowcatesearch"),"o")
	vItemID = request("itemid")
	vItemName = request("itemname")
	vMakerID = request("makerid")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	
	vParam = "&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&only1="&vOnly1&"&only2="&vOnly2&"&sort1="&vSort1&"&sort2="&vSort2&"&nowcatesearch="&vIsNowCateSearch&"&itemid="&vItemID&"&itemname="&vItemName&"&makerid="&vMakerID&""
	
	set cStyleLifeItemList = new ClsStyleLife
 	cStyleLifeItemList.FCurrPage = iCurrentpage
 	cStyleLifeItemList.FGubun = vIsNowCateSearch
 	cStyleLifeItemList.FCate1 = cdl
 	cStyleLifeItemList.FCate2 = cdm
 	cStyleLifeItemList.FCate3 = cds
 	cStyleLifeItemList.FOnly1 = vOnly1
 	cStyleLifeItemList.FOnly2 = vOnly2
 	cStyleLifeItemList.FSort1 = vSort1
 	cStyleLifeItemList.FSort2 = vSort2
 	cStyleLifeItemList.FItemID = vItemID
 	cStyleLifeItemList.FItemName = vItemName
 	cStyleLifeItemList.FMakerID = vMakerID
	arrList = cStyleLifeItemList.FStyleLifeItemList
	vCount = cStyleLifeItemList.ftotalcount
	set cStyleLifeItemList = Nothing
	
	'### 등록된 스타일별 select box
	set ocate = new cstylepickMenu
	ocate.frectisusing = "Y"
	ocate.getstylepick_cate_cd1()
%>

<script language="javascript">
function Check_All()
{
	var chk = document.frmitem.itemid; 
	var cnt = 0;
	var ischecked = ""
	if(document.getElementsByName("chkall")(0).checked)
	{
		ischecked = "checked"
		document.getElementById("asdf").innerHTML = "해제";
	}
	else
	{
		ischecked = ""
		document.getElementById("asdf").innerHTML = "";
	}
	
	if(cnt == 0 && chk.length != 0)
	{
		for(i = 0; i < chk.length; i++)
		{
			chk.item(i).checked = ischecked;
		}
		cnt++;
	}
}
function chkfrm()
{
	frm.submit();
}
function only1Change(a)
{
	document.frm.only1.value = a;
	chkfrm();
}
function only2Change(a)
{
	document.frm.only2.value = a;
	chkfrm();
}
function sort1Change(a)
{
	document.frm.sort1.value = a;
	chkfrm();
}
function sort2Change(a)
{
	document.frm.sort2.value = a;
	chkfrm();
}
function nowChangeCate(a)
{
	document.frmitemproc.gubun.value = "oneitemchange";
	document.frmitemproc.itemid.value = a;
	document.frmitemproc.submit();
}
function checkboxCheck(gubun)
{
	var j = document.frmitem.itemid.length;
	var k = new Array();
	var m = 0;

	for(var i=0; i < <%=CHKIIF(vCount=1,"1","j")%> ; i++){
	    if (document.frmitem.itemid<%=CHKIIF(vCount=1,"","[i]")%>.checked == true)
	    {
	    	if(gubun == "i")
	    	{
	        	k[m] = document.frmitem.itemid<%=CHKIIF(vCount=1,"","[i]")%>.value;
	        }
	        else if(gubun == "c")
	        {
	        	k[m] = document.frmitem.itemcate<%=CHKIIF(vCount=1,"","[i]")%>.value;
	        }
	        m = m+1;
	    }
	}
	return k;
}
function goDefault()
{
	var i = checkboxCheck("i");
	if(i != "")
	{
		document.frmitemproc.itemid.value = i;
		document.frmitemproc.gubun.value = "default";
		document.frmitemproc.submit();
	}
	else
	{
		alert("상품을 선택해 주세요.");
		return;
	}
}
function goSetStyle()
{
	jj = document.frmitemproc.stylecate.length;
	var kk = new Array();
	mm = 0;
	for(var ii=0; ii < jj ; ii++){
	    if (document.frmitemproc.stylecate[ii].checked == true)
	    {
	        kk[mm] = document.frmitemproc.stylecate[ii].value;
	        mm = mm+1;
	    }
	}
	if(kk == "")
	{
		alert("지정할 스타일을 선택해 주세요.");
		return;
	}
	

	var i = checkboxCheck("i");
	var c = checkboxCheck("c");

	if(i != "")
	{
		document.frmitemproc.itemid.value = i;
		document.frmitemproc.itemcate.value = c;
		document.frmitemproc.gubun.value = "setstyle";
		document.frmitemproc.submit();
	}
	else
	{
		alert("상품을 선택해 주세요.");
		return;
	}
}
function goMidCate()
{
	var itemmidcategory = window.open('/admin/stylepick/stylelife_item_midcate.asp','itemmidcategory','width=800,height=700,scrollbars=yes,resizable=yes');
	itemmidcategory.focus();
}
function goNotUseItem()
{
	var i = checkboxCheck("i");
	if(i != "")
	{
		document.frmitemproc.itemid.value = i;
		document.frmitemproc.gubun.value = "notuseitem";
		document.frmitemproc.submit();
	}
	else
	{
		alert("상품을 선택해 주세요.");
		return;
	}
}
</script>

<hr>
<input type="text" id="nowcatename" name="nowcatename" value="" size="100">
<form name="frm" method="get" action="<%=CurrURL()%>">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="only1" value="<%=vOnly1%>">
<input type="hidden" name="only2" value="<%=vOnly2%>">
<input type="hidden" name="sort1" value="<%=vSort1%>">
<input type="hidden" name="sort2" value="<%=vSort2%>">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="#F4F4F4"></td>
	<td bgcolor="FFFFFF">
		<label id="so" style="cursor:pointer;"><input type="radio" name="nowcatesearch" value="o" <%=CHKIIF(vIsNowCateSearch="o","checked","")%>> 현재 카테고리에서 검색</label>&nbsp;&nbsp;&nbsp;
		<label id="sx" style="cursor:pointer;"><input type="radio" name="nowcatesearch" value="x" <%=CHKIIF(vIsNowCateSearch="x","checked","")%>> 전체 카테고리에서 검색</label>
	</td>
</tr>
<tr>
	<td bgcolor="#F4F4F4">상품코드</td>
	<td bgcolor="FFFFFF"><input type="text" name="itemid" value="<%=vItemID%>" size="70">(쉼표로 복수입력가능)</td>
</tr>
<tr>
	<td bgcolor="#F4F4F4">상품명</td>
	<td bgcolor="FFFFFF"><input type="text" name="itemname" value="<%=vItemName%>" size="70"></td>
</tr>
<tr>
	<td bgcolor="#F4F4F4">브랜드ID검색</td>
	<td bgcolor="FFFFFF"><input type="text" class="text" name="makerid" value="<%=vMakerID%>"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" ></td>
</tr>
<tr>
	<td bgcolor="FFFFFF" colspan="2"><input type="submit" class="button" value=" 검    색 "></td>
</tr>
</table>
</form>

<table width="810" class="a">
<tr>
	<td style="padding:5px 0 3px 0;">Total <b><%=vCount%></b></td>
	<td style="padding:5px 0 3px 0;" align="right">
		<label id="o1c" style="cursor:pointer;"><input type="checkbox" name="only1chk" value="o" <%=CHKIIF(vOnly1="o","checked","")%> onClick="only1Change(<%=CHKIIF(vOnly1="o","''","this.value")%>);"> 판매중만 보기</label>&nbsp;&nbsp;&nbsp;
		<label id="o2c" style="cursor:pointer;"><input type="checkbox" name="only2chk" value="o" <%=CHKIIF(vOnly2="o","checked","")%> onClick="only2Change(<%=CHKIIF(vOnly2="o","''","this.value")%>);"> 스타일 미지정 상품만 보기</label>&nbsp;&nbsp;&nbsp;
		<select name="sort1chk" onChange="sort1Change(this.value)">
			<option value="">- 스타일별 -</option>
			<% for i=0 to ocate.FresultCount-1 %>
			<option value="<%= ocate.FItemList(i).fcd1 %>" <%=CHKIIF(CStr(vSort1)=CStr(ocate.FItemList(i).fcd1),"selected","")%>><%= ocate.FItemList(i).fcatename %></option>
		<% next %>
		</select>&nbsp;&nbsp;&nbsp;
		<select name="sort2chk" onChange="sort2Change(this.value)">
			<option value="ne" <%=CHKIIF(vSort2="ne","selected","")%>>최신순</option>
			<option value="hp" <%=CHKIIF(vSort2="hp","selected","")%>>고가순</option>
			<option value="lp" <%=CHKIIF(vSort2="lp","selected","")%>>저가순</option>
			<option value="be" <%=CHKIIF(vSort2="be","selected","")%>>인기상품순</option>
			<option value="mk" <%=CHKIIF(vSort2="mk","selected","")%>>브랜드순</option>
		</select>
	</td>
</tr>
</table>

<% IF isArray(arrList) THEN %>
<form name="frmitem" style="margin:0px;">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="#FFFFFF" colspan="5" style="padding-left:8px;"><input type="checkbox" name="chkall" value="" onClick="Check_All()"> 전체선택<span id="asdf"></span></td>
</tr>
	<% For intLoop =0 To UBound(arrList,2) %>
	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
		<td width="30"><input type="checkbox" name="itemid" value="<%=arrList(0,intLoop)%>"><input type="hidden" name="itemcate" value="<%=arrList(7,intLoop)%>"></td>
		<td width="60"><img src="<%=webImgUrl%>/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(2,intLoop)%>"></td>
		<td width="520" align="left">[<%=arrList(0,intLoop)%>]<%=db2html(arrList(1,intLoop))%></td>
		<td width="170">
			<%
				If arrList(3,intLoop) <> "" Then
					Response.Write StyleLifeItemComma(arrList(3,intLoop))
				Else
					Response.Write StyleNameSelectBox(arrList(0,intLoop),arrList(7,intLoop))
				End If
			%>
		</td>
	</tr>
	<% Next %>
	<%
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	iTotalPage 	=  int((vCount-1)/20) +1
	iStartPage = (Int((iCurrentpage-1)/20)*20) + 1
	
	If (iCurrentpage mod 20) = 0 Then
		iEndPage = iCurrentpage
	Else
		iEndPage = iStartPage + (20-1)
	End If
	%>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if (iStartPage-1 )> 0 then %><a href="?iC=<%= iStartPage-1 %><%=vParam%>" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(iCurrentpage) then
			%>
				<a href="?iC=<%= ix %><%=vParam%>" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
			<%		else %>
				<a href="?iC=<%= ix %><%=vParam%>" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
			<%
					end if
				next
			%>
	    	<% if CLng(iTotalPage) > CLng(iEndPage)  then %><a href="?iC=<%= ix %><%=vParam%>" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
		</td>
	</tr>
</table>
</form>

<form name="frmitemproc" action="stylelife_item_proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="gubun" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemcate" value="">
<input type="hidden" name="returnUrl" value="<%=CurrURLQ()%>">
<table height="70" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td width="290" bgcolor="#FFFFFF" style="padding-left:10px;">
		<label id="ca010" style="cursor:pointer;"><input type="checkbox" name="stylecate" value="010">클래식</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<label id="ca020" style="cursor:pointer;"><input type="checkbox" name="stylecate" value="020">큐트</label>&nbsp;&nbsp;&nbsp;
		<label id="ca040" style="cursor:pointer;"><input type="checkbox" name="stylecate" value="040">모던</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<label id="ca050" style="cursor:pointer;"><input type="checkbox" name="stylecate" value="050">네추럴</label>
		<br>
		<label id="ca060" style="cursor:pointer;"><input type="checkbox" name="stylecate" value="060">오리엔탈</label>&nbsp;&nbsp;&nbsp;
		<label id="ca070" style="cursor:pointer;"><input type="checkbox" name="stylecate" value="070">팝</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<label id="ca080" style="cursor:pointer;"><input type="checkbox" name="stylecate" value="080">로맨틱</label>&nbsp;&nbsp;&nbsp;
		<label id="ca090" style="cursor:pointer;"><input type="checkbox" name="stylecate" value="090">빈티지</label>
	</td>
	<td width="180" bgcolor="#FFFFFF" align="center">
		<input type="button" class="button" value="선택상품 스타일 지정" onClick="goSetStyle()" style="width:150px;height:40px;">
	</td>
	<td width="200" bgcolor="#FFFFFF" align="center">
		<input type="button" class="button" value="선택상품 스타일 초기화" onClick="goDefault()" style="width:170px;height:40px;">
	</td>
	<td width="103" bgcolor="#FFFFFF" align="center">
		<input type="button" class="button" value="중분류 관리" onClick="goMidCate()" style="height:40px;">
	</td>
</tr>
</table>
<br><input type="button" class="button" value="선택상품 스타일라이프에서 완전 제외" onClick="goNotUseItem()" style="width:300px;height:40px;">
</form>
<%
Else
	Response.WRite "데이터가 없습니다."
End If %>

<script language='javascript'>
document.getElementById("nowcatename").value = <% If vIsNowCateSearch = "x" Then %>"전체 카테고리"<% Else %>parent.document.getElementById("nowcatename").value<% End If %>;
</script>

<%
set ocate = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->