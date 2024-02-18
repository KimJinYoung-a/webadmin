<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  다이샾 플러스 상품
' History : 2010.11.09 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/PlusDIYItemCls.asp"-->
<%
dim itemid, page , oitem ,oitemoption ,i ,oPsItemList
	itemid = requestCheckvar(request("itemid"),9)
	page   = requestCheckvar(request("page"),9)

	if page="" then page=1

set oitem = new CItem
	oitem.FRectItemID = itemid

	if itemid<>"" then
		oitem.GetOneItem
	end if

set oitemoption = new CItemOption
	oitemoption.FRectItemID = itemid
	if itemid<>"" then
		oitemoption.GetItemOptionInfo
	end if

set oPsItemList = new CPlusSaleItem
	oPsItemList.FCurrPage = page
	oPsItemList.FPageSize = 20
	oPsItemList.FRectItemID = itemid
    oPsItemList.GetPlusSaleSubItemListByMainLinkItemID
%>

<script language='javascript'>

function PlusDIYItem_Main_Reg(iitemid){

	var acURL = "<%= server.UrlEncode("/academy/itemmaster/PlusDIYItem/PlusDIYItem_Process.asp?mode=PlusMainAddArr&PlusSaleItemid=") %>" + iitemid;
	var popwin = window.open("/academy/itemmaster/plusdiyitem/pop_plusdiyitem_list.asp?sellyn=Y&usingyn=Y&saleYn=&defaultmargin=0&plusSaleLinkItemID=<%=ItemID%>&acURL=" + acURL, "PlusDIYItem_Main_Reg", "width=1024,height=768,scrollbars=yes,resizable=yes");

    popwin.focus();
}

function delPlusSaleMainItem(){
    var frm;
    frm = document.frmList;
    var chkExists = false;

    if (!frm.chkitem){
        return;
    }else{
        if(!frm.chkitem.length){
            chkExists = frm.chkitem.checked;
        }else{
            for (var i=0;i<frm.chkitem.length;i++){
                if (frm.chkitem[i].checked){
                    chkExists = true;
                }
            }
        }
    }

    if (!chkExists){
        alert('선택된 내역이 없습니다.');
        return;
    }

    if (confirm('선택된 추가구성 상품을 삭제 하시겠습니까?')){
        frm.submit();
    }
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

//전체 선택
function jsChkAll(){
    var frm;
    frm = document.frmList;
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

</script>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="7">
		<!--<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">-->
	</td>
</tr>
</form>
</table>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oitem.FResultCount>0 then %>
<tr bgcolor="#FFFFFF">
	<td rowspan="<%= 6 + oitemoption.FResultCount -1 %>" width="100" valign="top" align="center">
		<img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100" border="0">
	</td>
  	<td width="60" bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
  	<td width="300"><%= oitem.FOneItem.FItemID %></td>
  	<td width="80" bgcolor="<%= adminColor("tabletop") %>"><!-- PlusSale구분 --></td>
  	<td>
  	    플러스 D.I.Y 상품
  	</td>
</tr>
<tr bgcolor="#FFFFFF">
  	<td bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
  	<td><%= oitem.FOneItem.FMakerid %></td>
  	<td bgcolor="<%= adminColor("tabletop") %>">판매여부</td>
  	<td><%= oitem.FOneItem.FSellyn %></td>
</tr>
<tr bgcolor="#FFFFFF">
  	<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
  	<td><%= oitem.FOneItem.FItemName %></td>
  	<td bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
  	<td><%= oitem.FOneItem.FIsUsing %></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("tabletop") %>">가격정보</td>
	<td colspan=4>
		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuyCash,0) %>
		&nbsp;
		<% if oitem.FOneItem.FSellcash<>0 then %>
		<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
		<% end if %>
		&nbsp;&nbsp;
		<!-- 할인여부/쿠폰적용여부 -->
		<% if (oitem.FOneItem.FsaleYn="Y") then %>
		    <font color=red>
		    <% if (oitem.FOneItem.Forgprice<>0) then %>
		        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
		    <% end if %>
		     할인
		    </font>
		<% end if %>

		<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

		    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
		    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
		<% end if %>

		&nbsp;&nbsp;
		<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>
	</td>
</tr>
    <% if oitemoption.FResultCount>1 then %>
    <!-- 옵션이 있는경우 -->
    <% for i=0 to oitemoption.FResultCount -1 %>
	    <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
	    <tr bgcolor="#FFFFFF">
	      	<td bgcolor="<%= adminColor("tabletop") %>"><font color="#AAAAAA">옵션명 :</font></td>
	      	<td><font color="#AAAAAA"><%= oitemoption.FITemList(i).FOptionName %></font></td>
	      	<td bgcolor="<%= adminColor("tabletop") %>"><font color="#AAAAAA">한정여부 : </font></td>
	      	<td><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
	    </tr>
	    <% else %>


	    <tr bgcolor="#FFFFFF">
	      	<td bgcolor="<%= adminColor("tabletop") %>">옵션명</td>
	      	<td><%= oitemoption.FITemList(i).FOptionName %></td>
	      	<td bgcolor="<%= adminColor("tabletop") %>">한정여부</td>
	      	<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
	    </tr>
	    <% end if %>
    <% next %>
    <% else %>
	<tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">옵션명</td>
      	<td>-</td>
      	<td bgcolor="<%= adminColor("tabletop") %>">한정여부</td>
      	<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
    </tr>
    <% end if %>

<% else %>
<tr bgcolor="#FFFFFF"><td align="center"> 상품 검색 결과가 없습니다. </td></tr>
<% end if %>
</table>
<br>

<% if (oitem.FResultCount>0) then %>
<!-- <b>메인 링크상품인 경우.....</b> -->
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				<img src="/images/icon_star.gif" border="0" align="absbottom">
				<b>추가 구성 상품 리스트</b>
				&nbsp;
				검색결과 : <b><%= oPsItemList.FTotalCount %></b>
			</td>
			<td align="right">
				<input type="button" class="button" value="선택상품 삭제" onClick="delPlusSaleMainItem('<%= itemid %>');">
				<input type="button" class="button" value="추가구성상품 추가" onClick="PlusDIYItem_Main_Reg('<%= itemid %>');">
			</td>
		</tr>
	</table>
	</td>
</tr>
<form name="frmList" method="post" action="PlusDIYItem_Process.asp">
<input type="hidden" name="mode" value="PlusSaleDellArr">
<input type="hidden" name="PlusSaleLinkItemid" value="<%= itemid %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td width="60">상품코드</td>
	<td width="50">이미지</td>
  	<td>브랜드ID</td>
  	<td>상품명</td>
  	<td width="50">판매가</td>
	<td width="50">매입가</td>
	<td width="40">마진</td>
	<td width="80">계약구분</td>
	<td width="40">플러스<br>할인율</td>
	<td width="35">판매<br>여부</td>
	<!-- td width="35">PS<br>사용</td -->
</tr>
<% if oPsItemList.FResultCount>0 then %>
<% for i=0 to oPsItemList.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="2"><input type="checkbox" name="chkitem" value="<%= oPsItemList.FItemList(i).FPlusSaleItemID %>"></td>
	<td rowspan="2"><%= oPsItemList.FItemList(i).FPlusSaleItemID %></td>
  	<td rowspan="2"><img src="<%= oPsItemList.FItemList(i).FImageSmall %>" width="50" height="50" border="0"></td>
  	<td rowspan="2"><%= oPsItemList.FItemList(i).FMakerID %></td>
  	<td align="left"><%= oPsItemList.FItemList(i).FItemName %></td>
  	<td align="right"><%= FormatNumber(oPsItemList.FItemList(i).FSellCash,0) %></td>
  	<td align="right"><%= FormatNumber(oPsItemList.FItemList(i).FBuyCash,0) %></td>
  	<td><%= fnPercent(oPsItemList.FItemList(i).FBuyCash,oPsItemList.FItemList(i).FSellCash,1) %></td>
  	<td><%= oPsItemList.FItemList(i).FMwdiv %></td>
  	<td rowspan="2"><%= oPsItemList.FItemList(i).FPlusSalePro %>%</td>
  	<td rowspan="2"><%= fnColor(oPsItemList.FItemList(i).FSellyn,"sellyn") %></td>
  	<!-- td rowspan="2">Y</td -->
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left"><%= Left(oPsItemList.FItemList(i).FPlusSaleStartDate,10) %> ~ <%= Left(oPsItemList.FItemList(i).FPlusSaleEndDate,10) %> <font color="<%= oPsItemList.FItemList(i).getCurrstateColor %>">(<%= oPsItemList.FItemList(i).getCurrstateName %>)</font></td>
	<td align="right"><font color="#CC33FF"><%= FormatNumber(oPsItemList.FItemList(i).getPlusSalePrice,0) %></font></td>
	<td align="right"><font color="#CC33FF"><%= FormatNumber(oPsItemList.FItemList(i).getPlusSaleBuycash,0) %></font></td>
	<td><%= fnPercent(oPsItemList.FItemList(i).getPlusSaleBuycash,oPsItemList.FItemList(i).getPlusSalePrice,1) %></td>
	<td><%= oPsItemList.FItemList(i).getMaginFlagName %></td>
</tr>
<% next %>
<tr>
    <td colspan="20" align="center" bgcolor="#FFFFFF">
        <% if oPsItemList.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oPsItemList.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oPsItemList.StartScrollPage to oPsItemList.FScrollCount + oPsItemList.StartScrollPage - 1 %>
			<% if i>oPsItemList.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oPsItemList.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
<% else %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		현재 진행중인 플러스 D.I.Y 상품이 없습니다.
	</td>
</tr>
<% end if %>
</form>
</table>
<% end if %>

<%
	set oitem = Nothing
	set oitemoption = Nothing
	set oPsItemList = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
