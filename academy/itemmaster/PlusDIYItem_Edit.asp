<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/PlusDIYItemCls.asp"-->

<%

dim itemid, page
itemid = requestCheckvar(request("itemid"),9)
page   = requestCheckvar(request("page"),9)

if page="" then page=1

dim oitem
set oitem = new CItem
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItem
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

dim oPlusSaleItem
set oPlusSaleItem = new CPlusSaleItem
oPlusSaleItem.FRectItemID = itemid

if itemid<>"" then
	oPlusSaleItem.GetOnePlusSaleSubItem
end if

dim i
dim IsPlusSaleItem        '' 플러스세일 상품인지
IsPlusSaleItem = (oPlusSaleItem.FResultCount>0)


'' 기존 IsLinkedItem 인경우
dim IsLinkedItem
if itemid<>"" then
    IsLinkedItem = oPlusSaleItem.IsPlusSaleLinkItem
end if

dim oPsItemList
set oPsItemList = new CPlusSaleItem
oPsItemList.FCurrPage = page
oPsItemList.FPageSize = 20
oPsItemList.FRectItemID = itemid
if (IsPlusSaleItem) then
    oPsItemList.FRectPlusSaleItemID = itemid
    oPsItemList.GetPlusSaleMainItemList
elseif (IsLinkedItem) then
    oPsItemList.GetPlusSaleSubItemListByMainLinkItemID
end if
%>

<script language='javascript'>

function PlusDIYItem_Main_Reg(iitemid){
	//var popwin = window.open('PlusDIYItem_Main_Reg.asp','PlusDIYItem_Main_Reg','width=900,height=600,scrollbars=yes,resizable=yes')
	//popwin.focus();
	var acURL = "<%= server.UrlEncode("/academy/itemmaster/PlusDIYItem_Process.asp?mode=PlusMainAddArr&PlusSaleItemid=") %>" + iitemid;
	var popwin = window.open("/academy/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&saleYn=&defaultmargin=0&acURL=" + acURL, "PlusDIYItem_Main_Reg", "width=800,height=500,scrollbars=yes,resizable=yes");

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

    if (confirm('선택된 메인링크 상품을 삭제 하시겠습니까?')){
        frm.submit();
    }
}

function delPlusSaleItem(){
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

    if (confirm('선택된 플러스 세일 상품을 삭제 하시겠습니까?')){
        frm.submit();
    }
}

function PlusDIYItem_Sub_Reg(iitemid){
	var popwin = window.open('PlusDIYItem_Sub_Reg.asp?PlusSaleLinkItemid=' + iitemid,'PlusDIYItem_Sub_Reg','width=900,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
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
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<p>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oitem.FResultCount>0 then %>
    <tr bgcolor="#FFFFFF">
        <% if  (IsPlusSaleItem) then %>
        <td rowspan="<%= 7 + oitemoption.FResultCount -1 %>" width="100" valign="top" align="center">
        <% else %>
    	<td rowspan="<%= 6 + oitemoption.FResultCount -1 %>" width="100" valign="top" align="center">
    	<% end if %>
    		<img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100" border="0">
		</td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
      	<td width="300"><%= oitem.FOneItem.FItemID %></td>
      	<td width="80" bgcolor="<%= adminColor("tabletop") %>"><!-- PlusSale구분 --></td>
      	<td>
      	    <% if (IsPlusSaleItem) then %>
      	    플러스세일 추가구성상품
      		<% end if %>

      	    <% if (IsLinkedItem) then %>
      	    메인 링크 상품
      	    <% end if %>
      	</td>
    </tr>
    <% if  (IsPlusSaleItem) then %>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>"><font color="red">플러스세일 정보</font></td>
        <td colspan=4>
            <%= oPlusSaleItem.FOneItem.FPlusSalePro %>% 할인
            <%= FormatNumber(oPlusSaleItem.FOneItem.getPlusSalePrice,0) %>
            /
            <%= FormatNumber(oPlusSaleItem.FOneItem.getPlusSaleBuycash,0) %>
            (<%= oPlusSaleItem.FOneItem.FPlusSaleMargin %>%)

            <%= oPlusSaleItem.FOneItem.getMaginFlagName %>
        </td>
    </tr>
    <% end if %>
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
<p>



<!-- 플러스세일상품(추가구성상품)에 등록되었을 경우, 리스트를 표시한다 -->
<% if  (IsPlusSaleItem) then %>
<br>
<!-- <b>플러스 세일 상품인 경우.....</b> -->
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_star.gif" border="0" align="absbottom">
					<b>메인 링크 상품 리스트</b>
					&nbsp;
					검색결과 : <b><%= oPsItemList.FTotalCount %></b>
				</td>
				<td align="right">
					<input type="button" class="button" value="선택상품 삭제" onClick="delPlusSaleMainItem('<%= itemid %>');">
					<input type="button" class="button" value="메인링크상품 추가" onClick="PlusDIYItem_Main_Reg('<%= itemid %>');">
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<form name="frmList" method="post" action="PlusDIYItem_Process.asp">
	<input type="hidden" name="mode" value="PlusMainDellArr">
	<input type="hidden" name="PlusSaleItemID" value="<%= itemid %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	   	<td width="60">상품코드</td>
    	<td width="50">이미지</td>
      	<td width="100">브랜드ID</td>
      	<td>상품명</td>
      	<td width="60">판매가</td>
		<td width="60">매입가</td>
		<td width="40">마진</td>
		<td width="40">거래<br>구분</td>
		<td width="40">판매<br>여부</td>
    </tr>
    <% if oPsItemList.FResultCount<1 then %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			현재 진행중인 메인 상품이 없습니다.
		</td>
	</tr>
	<% else %>
    <% for i=0 to oPsItemList.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><input type="checkbox" name="chkitem" value="<%= oPsItemList.FItemList(i).FPlusSaleLinkItemID %>"></td>
    	<td><%= oPsItemList.FItemList(i).FPlusSaleLinkItemID %></td>
    	<td><img src="<%= oPsItemList.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
      	<td><%= oPsItemList.FItemList(i).FMakerID %></td>
      	<td><%= oPsItemList.FItemList(i).FItemName %></td>
		<td align="right">
		    <% if oPsItemList.FItemList(i).FsaleYn="Y" then %>
      		<%= FormatNumber(oPsItemList.FItemList(i).FOrgPrice,0) %>
      		<br><font color=#F08050>(할)<%= FormatNumber(oPsItemList.FItemList(i).FSellcash,0) %></font>
		    <% else %>
		    <%= FormatNumber(oPsItemList.FItemList(i).FSellcash,0) %>
      		<% end if %>
      	</td>
      	<td align="right">
      	    <% if oPsItemList.FItemList(i).FsaleYn="Y" then %>
      		<%= FormatNumber(oPsItemList.FItemList(i).Forgsuplycash,0) %>
      		<br><font color=#F08050>(할)<%= FormatNumber(oPsItemList.FItemList(i).FBuycash,0) %></font>
      	    <% else %>
      	    <%= FormatNumber(oPsItemList.FItemList(i).FBuycash,0) %>
      		<% end if %>
      	</td>
      	<td>
      	    <% if oPsItemList.FItemList(i).FsaleYn="Y" then %>
      		<%= fnPercent(oPsItemList.FItemList(i).Forgsuplycash,oPsItemList.FItemList(i).FOrgPrice,1) %>
      		<br><font color=#F08050><%= fnPercent(oPsItemList.FItemList(i).Forgsuplycash,oPsItemList.FItemList(i).FOrgPrice,1) %></font>
      	    <% else %>
      	    <%= fnPercent(oPsItemList.FItemList(i).FBuycash,oPsItemList.FItemList(i).FSellcash,1) %>
      		<% end if %>
      	</td>
      	<td><%= fnColor(oPsItemList.FItemList(i).FMwDiv,"mw") %></td>
      	<td><%= fnColor(oPsItemList.FItemList(i).FSellyn,"sellyn") %></td>
    </tr>
    <% next %>
    </form>
    <tr>
        <td colspan="20" align="center" bgcolor="#FFFFFF">
            <% if oPsItemList.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= oPsItemList.StarScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oPsItemList.StarScrollPage to oPsItemList.FScrollCount + oPsItemList.StarScrollPage - 1 %>
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
    <% end if %>
</table>
<% end if %>

<% if (oitem.FResultCount>0) then %>
    <% if (IsLinkedItem) or (Not IsPlusSaleItem) then %>
    <br>
    <!-- <b>메인 링크상품인 경우.....</b> -->
    <!-- 리스트 시작 -->
    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    	<tr height="25" bgcolor="FFFFFF">
    		<td colspan="20">
    		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
    			<tr>
    				<td>
    					<img src="/images/icon_star.gif" border="0" align="absbottom">
    					<b>플러스 세일 상품 리스트</b>
    					&nbsp;
    					검색결과 : <b><%= oPsItemList.FTotalCount %></b>
    				</td>
    				<td align="right">
    					<input type="button" class="button" value="선택상품 삭제" onClick="delPlusSaleItem('<%= itemid %>');">
    					<input type="button" class="button" value="추가구성상품 추가" onClick="PlusDIYItem_Sub_Reg('<%= itemid %>');">
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
        			<a href="javascript:NextPage('<%= oPsItemList.StarScrollPage-1 %>')">[pre]</a>
        		<% else %>
        			[pre]
        		<% end if %>

        		<% for i=0 + oPsItemList.StarScrollPage to oPsItemList.FScrollCount + oPsItemList.StarScrollPage - 1 %>
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
    			현재 진행중인 플러스할인상품이 없습니다.
    		</td>
    	</tr>
    	<% end if %>
    	</form>
    </table>
    <% end if %>
<% end if %>
<% if (Not IsPlusSaleItem) and (Not IsLinkedItem) then %>
<!--
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
        <td align="center"> 플러스 세일 상품이 아닙니다. </td>
    </tr>
</table>
-->
<% end if %>
<%
set oitem = Nothing
set oitemoption = Nothing
set oPlusSaleItem = Nothing
set oPsItemList = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
