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
	Dim cdl, cdm, cds, vIsNowCateSearch, vTempCate, vSort2, vItemID, vItemName, vMakerID, cStyleLifeItemList, ocate, i, vCount, arrList, intLoop, iCurrentpage, vParam
	Dim vCate1, vCate2, vCate3
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	vCate1 = request("cd1")
	vCate2 = request("cd2")
	vCate3 = request("cd3")
	vTempCate = request("temp_cate2")
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	
	vParam = "&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&cd1="&vCate1&"&cd2="&vCate2&"&cd3="&vCate3&"&sort2="&vSort2&"&nowcatesearch="&vIsNowCateSearch&"&itemid="&vItemID&"&itemname="&vItemName&"&makerid="&vMakerID&"&temp_cate2="&vTempCate&""
	
	set cStyleLifeItemList = new ClsStyleLife
 	cStyleLifeItemList.FCurrPage = iCurrentpage
 	cStyleLifeItemList.FLCate = cdl
 	cStyleLifeItemList.FMCate = cdm
 	cStyleLifeItemList.FSCate = cds
 	cStyleLifeItemList.FCate1 = vCate1
 	cStyleLifeItemList.FCate2 = vCate2
 	cStyleLifeItemList.FCate3 = vCate3
	arrList = cStyleLifeItemList.FStyleLifeItemMidCateList
	vCount = cStyleLifeItemList.ftotalcount
	set cStyleLifeItemList = Nothing
	
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
function chkfrm()
{
	frm.submit();
}
function goTempCate2(a)
{
	document.frm.temp_cate2.value = a;
	document.frm.submit();
}
function goSaveMidCate(g)
{
	if(g != "all")
	{
		var i = checkboxCheck("i");
		if(i != "")
		{
			document.frmitemproc.itemid.value = i;
		}
		else
		{
			alert("상품을 선택해 주세요.");
			return;
		}
	}
	
	if(frmitem.cate3.value == "")
	{
		alert("지정할 중분류를 선택해 주세요.");
		return;
	}
	else
	{
		document.frmitemproc.cate3.value = document.frmitem.cate3.value;
	}
	
	if(g == "all")
	{
		document.frmitemproc.gubun.value = g;
		document.frmitemproc.submit();
	}
	else
	{
		document.frmitemproc.submit();
	}
}
</script>

<form name="frm" method="get" action="<%=CurrURL()%>">
<input type="hidden" name="temp_cate2" value="">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="FFFFFF">
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
</tr>
<tr>
	<td bgcolor="FFFFFF">
		<select name="cd1">
			<option value="">-스타일-</option>
			<option value="010" <%=CHKIIF(vCate1="010","selected","")%>>클래식</option>
			<option value="020" <%=CHKIIF(vCate1="020","selected","")%>>큐트</option>
			<option value="040" <%=CHKIIF(vCate1="040","selected","")%>>모던</option>
			<option value="050" <%=CHKIIF(vCate1="050","selected","")%>>네추럴</option>
			<option value="060" <%=CHKIIF(vCate1="060","selected","")%>>오리엔탈</option>
			<option value="070" <%=CHKIIF(vCate1="070","selected","")%>>팝</option>
			<option value="080" <%=CHKIIF(vCate1="080","selected","")%>>로맨틱</option>
			<option value="090" <%=CHKIIF(vCate1="090","selected","")%>>빈티지</option>
		</select>
		<select name="cd2">
			<option value="">-분류-</option>
			<%
				rsget.Open "select cd2, catename from db_giftplus.dbo.tbl_stylepick_cate_cd2 where isusing = 'Y' order by orderno", dbget, 1
				Do Until rsget.Eof
					Response.Write "<option value=""" & rsget("cd2") & """ " & CHKIIF(CStr(vCate2)=CStr(rsget("cd2")),"selected","") & ">" & rsget("catename") & "</option>"
				rsget.MoveNext
				Loop
				rsget.Close()
			%>
		</select>
		<% If vCate2 <> "" Then %>
		<select name="cd3">
			<option value="">-중분류-</option>
			<%
				rsget.Open "select cd3, catename from db_giftplus.dbo.tbl_stylepick_cate_cd3 where isusing = 'Y' and Left(cd3,1) = '" & Mid(vCate2,2,1) & "' order by orderno", dbget, 1
				Do Until rsget.Eof
					Response.Write "<option value=""" & rsget("cd3") & """ " & CHKIIF(CStr(vCate3)=CStr(rsget("cd3")),"selected","") & ">" & rsget("catename") & "</option>"
				rsget.MoveNext
				Loop
				rsget.Close()
			%>
		</select>
		<% End If %>
	</td>
</tr>
<tr>
	<td bgcolor="FFFFFF" colspan="2"><input type="submit" class="button" value=" 검    색 "></td>
</tr>
</table>
</form>

<% IF isArray(arrList) THEN %>
<form name="frmitem" style="margin:0px;">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="#FFFFFF" colspan="5" style="padding-left:8px;"><input type="checkbox" name="chkall" value="" onClick="Check_All()"> 전체선택<span id="asdf"></span>&nbsp;&nbsp;&nbsp;Total <b><%=vCount%></b></td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="5" style="padding-left:8px;">
		<select name="cate2" onChange="goTempCate2(this.value);">
			<option value="">-분류-</option>
			<%
				rsget.Open "select cd2, catename from db_giftplus.dbo.tbl_stylepick_cate_cd2 where isusing = 'Y' order by orderno", dbget, 1
				Do Until rsget.Eof
					Response.Write "<option value=""" & rsget("cd2") & """ " & CHKIIF(CStr(vTempCate)=CStr(rsget("cd2")),"selected","") & ">" & rsget("catename") & "</option>"
				rsget.MoveNext
				Loop
				rsget.Close()
			%>
		</select>
		<% If vTempCate <> "" Then %>
		지정할 중분류
		<select name="cate3">
			<option value="">-중분류-</option>
			<%
				rsget.Open "select cd3, catename from db_giftplus.dbo.tbl_stylepick_cate_cd3 where isusing = 'Y' and Left(cd3,1) = '" & Mid(vTempCate,2,1) & "' order by orderno", dbget, 1
				Do Until rsget.Eof
					Response.Write "<option value=""" & rsget("cd3") & """>" & rsget("catename") & "</option>"
				rsget.MoveNext
				Loop
				rsget.Close()
			%>
		</select>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" class="button" value="중분류지정" style="width:80px;" onClick="goSaveMidCate('')">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" class="button" value="검색된상품<%=vCount%>개모두지정" style="width:160px;" onClick="goSaveMidCate('all')">
		<% End If %>
	</td>
</tr>
	<% For intLoop =0 To UBound(arrList,2) %>
	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
		<td width="30"><input type="checkbox" name="itemid" value="<%=arrList(0,intLoop)%>"></td>
		<td width="60"><img src="<%=webImgUrl%>/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(2,intLoop)%>"></td>
		<td width="70"><%=arrList(6,intLoop)%></td>
		<td width="650" align="left">[<%=arrList(0,intLoop)%>]<%=db2html(arrList(1,intLoop))%></td>
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

<form name="frmitemproc" action="stylelife_item_midcate_proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="gubun" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="cd1" value="<%=vCate1%>">
<input type="hidden" name="cd2" value="<%=vCate2%>">
<input type="hidden" name="cd3" value="<%=vCate3%>">
<input type="hidden" name="cate3" value="">
<input type="hidden" name="returnUrl" value="<%=CurrURLQ()%>">
</form>
<%
Else
	Response.WRite "데이터가 없습니다."
End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->