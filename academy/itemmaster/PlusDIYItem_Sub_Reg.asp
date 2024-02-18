<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/PlusDIYItemCls.asp"-->

<%
dim page
dim makerid, itemidArr, itemname
dim cdl, cdm, cds
dim openstate, research, sellyn, mwdiv
dim PlusSaleLinkItemid

page        = RequestCheckVar(request("page"),9)
makerid     = RequestCheckVar(request("makerid"),32)
itemidArr   = RequestCheckVar(request("itemidArr"),1024)
itemname    = RequestCheckVar(request("itemname"),64)
cdl         = RequestCheckVar(request("cdl"),3)
cdm         = RequestCheckVar(request("cdm"),3)
cds         = RequestCheckVar(request("cds"),3)
openstate   = RequestCheckVar(request("openstate"),32)
research    = RequestCheckVar(request("research"),32)
sellyn      = RequestCheckVar(request("sellyn"),9)
mwdiv      = RequestCheckVar(request("mwdiv"),9)
PlusSaleLinkItemid = RequestCheckVar(request("PlusSaleLinkItemid"),9)

if (research="") and (openstate="") then openstate="openscheduled"

if (page="") then page=1
itemidArr = Trim(itemidArr)
itemname  = Trim(itemname)
if (Right(itemidArr,1)=",") then itemidArr = Left(itemidArr,Len(itemidArr)-1)

dim oPlusSaleItem
set oPlusSaleItem = new CPlusSaleItem
oPlusSaleItem.FRectMakerid  = makerid
oPlusSaleItem.FRectCDL      = cdl
oPlusSaleItem.FRectCDM      = cdm
oPlusSaleItem.FRectCDS      = cds
oPlusSaleItem.FRectItemIDArr= itemidArr
oPlusSaleItem.FRectItemName = itemname
oPlusSaleItem.FRectOpenState= openstate
oPlusSaleItem.FRectMwDiv    = mwdiv
oPlusSaleItem.FRectSellYn   = sellyn
oPlusSaleItem.FRectPlusSaleLinkItemid = PlusSaleLinkItemid

oPlusSaleItem.GetPlusSaleSubItemList

dim i
%>

<script language='javascript'>
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

function AddSubItemArr(frm){
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
    
    if (confirm('선택된 상품을 추가 하시겠습니까?')){
        frm.submit();
    }
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="1">
	
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			<img src="/images/icon_star.gif" border="0" align="absbottom">
			<b>PlusSale 추가구성상품 추가</b>
		</td>
	</tr>
	<tr>
	    <td bgcolor="<%= adminColor("gray") %>">메인상품</td>
	    <td colspan="3" bgcolor="#FFFFFF" >
	    <input type="text" Class="text_ro" name="PlusSaleLinkItemid" value="<%= PlusSaleLinkItemid %>" ReadOnly size="6">
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/academy/comm/CategorySelectBox.asp"-->
			<br>
			상품코드 :
			<input type="text" class="text" name="itemidArr" value="<%= itemidArr %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
			&nbsp;
			상품명 :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
			&nbsp;
			거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
			&nbsp;
			진행상태 : 
			<select class="select" name="openstate">
              <option value="">전체</option>
              <option value="open" <%= ChkIIF(openstate="open","selected","") %> >진행중</option>
              <option value="scheduled" <%= ChkIIF(openstate="scheduled","selected","") %> >진행예정</option>
              <option value="openscheduled" <%= ChkIIF(openstate="openscheduled","selected","") %> >진행중+진행예정</option>
              <option value="expired" <%= ChkIIF(openstate="expired","selected","") %> >기간종료</option>
            </select>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					검색결과 : <b><%= oPlusSaleItem.FTotalCount %></b>
					&nbsp;
					페이지 : <b><%= page %> / <%= oPlusSaleItem.FTotalPage %></b>
				</td>
				<td align="right">
					<input type="button" class="button" value="선택상품 추가" onClick="AddSubItemArr(frmList)">
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<form name="frmList" method="post" action="PlusDIYItem_Process.asp">
	<input type="hidden" name="PlusSaleLinkItemid" value="<%= PlusSaleLinkItemid %>">
	<input type="hidden" name="mode" value="PlusSaleAddArr">
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
		<!-- td width="35"></td -->
    </tr>
    <% for i=0 to oPlusSaleItem.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td rowspan="2"><input type="checkbox" name="chkitem" value="<%= oPlusSaleItem.FItemList(i).FPlusSaleItemID %>"></td>
      	<td rowspan="2"><%= oPlusSaleItem.FItemList(i).FPlusSaleItemID %></td>
    	<td rowspan="2"><img src="<%= oPlusSaleItem.FItemList(i).FImageSmall %>" width="50" height="50" border="0"></td>
      	<td rowspan="2"><%= oPlusSaleItem.FItemList(i).FMakerID %></td>
      	<td align="left"><%= oPlusSaleItem.FItemList(i).FItemName %></td>
      	<td align="right">
      	    
      	    <%= FormatNumber(oPlusSaleItem.FItemList(i).FOrgPrice,0) %>
      	    <% if oPlusSaleItem.FItemList(i).IsCurrentSaleItem then %>
      		<br><font color=#F08050>(할)<%= FormatNumber(oPlusSaleItem.FItemList(i).FSellcash,0) %></font>
      		<% end if %>
      		
      		<% if oPlusSaleItem.FItemList(i).IsCouponItem then %>
      	        <br><font color=#5080F0>(쿠)<%= FormatNumber(oPlusSaleItem.FItemList(i).GetCouponAssignPrice,0) %></font>
      	    <% end if %>
      	</td>
      	<td align="right">
      	    
      		<%= FormatNumber(oPlusSaleItem.FItemList(i).FOrgSuplycash,0) %>
      		<% if oPlusSaleItem.FItemList(i).IsCurrentSaleItem then %>
      		<br><font color=#F08050>(할)<%= FormatNumber(oPlusSaleItem.FItemList(i).FBuycash,0) %></font>
      		<% end if %>
      	</td>
      	<td>
      		<%= fnPercent(oPlusSaleItem.FItemList(i).FOrgSuplycash,oPlusSaleItem.FItemList(i).FOrgPrice,1) %>
      		<% if oPlusSaleItem.FItemList(i).IsCurrentSaleItem then %>
      		<%= FormatNumber(oPlusSaleItem.FItemList(i).FOrgSuplycash,0) %>
      		<br><font color=#F08050><%= fnPercent(oPlusSaleItem.FItemList(i).FBuycash,oPlusSaleItem.FItemList(i).FSellcash,1) %></font>
      		<% end if %>
      	</td>
      	<td><%= fnColor(oPlusSaleItem.FItemList(i).FMwdiv,"mw") %></td>
      	<td rowspan="2">
      	    <%= oPlusSaleItem.FItemList(i).FPlusSalePro %>%
      	</td>
      	<td rowspan="2"><%= fnColor(oPlusSaleItem.FItemList(i).FSellyn,"sellyn") %></td>
      	<!-- td rowspan="2">Y</td -->
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td align="left"><%= Left(oPlusSaleItem.FItemList(i).FPlusSaleStartDate,10) %> ~ <%= Left(oPlusSaleItem.FItemList(i).FPlusSaleEndDate,10) %> (<%= oPlusSaleItem.FItemList(i).getCurrstateName %>)</td>
    	<td align="right"><font color="#CC33FF"><%= FormatNumber(oPlusSaleItem.FItemList(i).getPlusSalePrice,0) %></font></td>
    	<td align="right"><font color="#CC33FF"><%= FormatNumber(oPlusSaleItem.FItemList(i).getPlusSaleBuycash,0) %></font></td>
    	<td><font color="#CC33FF"><%= fnPercent(oPlusSaleItem.FItemList(i).getPlusSaleBuycash,oPlusSaleItem.FItemList(i).getPlusSalePrice,1) %></font></td>
    	<td><font color="#CC33FF"><%= oPlusSaleItem.FItemList(i).getMaginFlagName %></font></td>
    </tr>
    <% next %>
    
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if oPlusSaleItem.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= oPlusSaleItem.StarScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>
    
    		<% for i=0 + oPlusSaleItem.StarScrollPage to oPlusSaleItem.FScrollCount + oPlusSaleItem.StarScrollPage - 1 %>
    			<% if i>oPlusSaleItem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>
    
    		<% if oPlusSaleItem.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
	</form>
</table>

<%
set oPlusSaleItem = Nothing
%>
<!--
<p>
* 플러스 할인가 지정색(일반할인가-빨간색 / 쿠폰할인가-초록색 / 플러스할인가-보라색)<br>
* 매입상품인 경우, 무조건 텐바이텐 부담? (동일마진 / 업체부담 / 반반부담 /텐바이텐부담)
-->


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
