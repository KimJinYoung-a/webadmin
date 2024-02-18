<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/upcheitemeditcls.asp"-->

<%
dim designerid, itemid
dim research,notfinish
dim page
page = request("page")
designerid = request("designerid")
itemid = request("itemid")
research = request("research")
notfinish = request("notfinish")

if page="" then page=1

if research<>"on" then
	notfinish = "on"
end if

dim isfinishStr
if notfinish="on" then
	isfinishStr="N"
end if

dim oupcheitemedit
set oupcheitemedit = New CUpCheItemEdit
oupcheitemedit.FPageSize = 20
oupcheitemedit.FCurrPage = page
oupcheitemedit.FRectDesignerID =  designerid
oupcheitemedit.FRectItemId = itemid
oupcheitemedit.FRectNotFinish = isfinishStr

oupcheitemedit.GetReqList

dim i
%>
<script language='javascript'>

function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function DelThis2(){

}

function DelThis(frm){
	if (frm.rejectstr.value.length<1){
		alert('거부 사유를 입력해 주세요.');
		frm.rejectstr.focus();
		return;
	}

	var ret = confirm('승인 거부 하시겠습니까?');

	if (ret){
		frm.mode.value="del";
		frm.submit();
	}
}

function AccThis(frm){
	frm.mode.value="acct";

	if ((frm.limitSetno)&&(!IsDigit(frm.limitSetno.value))){
		alert('숫자만 가능합니다.');
		frm.limitno.focus();
		return;
	}

//	if (!IsDigit(frm.limitsold.value)){
//		alert('숫자만 가능합니다.');
//		frm.limitno.focus();
//		return;
//	}

	var ret = confirm('승인 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}


// ============================================================================
function ChangeDispYN(frm, divdispyn) {
    if (frm.dispyn.value == "Y") {
        frm.dispyn.value = "N";
        divdispyn.innerHTML = "<font color=red>숨김</font>";
    } else {
        frm.dispyn.value = "Y";
        divdispyn.innerHTML = "전시";
    }
}

function ChangeSellYN(frm, divsellyn) {
    if (frm.sellyn.value == "Y") {
        frm.sellyn.value = "N";
        divsellyn.innerHTML = "<font color=red>품절</font>";
    } else {
        frm.sellyn.value = "Y";
        divsellyn.innerHTML = "판매";
    }
}

function ChangeLimitYN(frm, divlimityn) {
    if (frm.limityn.value == "Y") {
        frm.limityn.value = "N";
        divlimityn.innerHTML = "일반";
        frm.limitSetno.disabled = true;
        frm.limitSetno.style.background = "#CCCCCC";
    } else {
        frm.limityn.value = "Y";
        divlimityn.innerHTML = "<font color=red>한정</font>";
        frm.limitSetno.disabled = false;
        frm.limitSetno.style.background = "#FFFFFF";
    }
}

</script>
 
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 :
			<% drawSelectBoxDesigner "designerid",designerid %>
			&nbsp;
			상품코드 :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="9">
			&nbsp;
			<input type="checkbox" name="notfinish" <% if notfinish="on" then response.write "checked" %> >미처리목록만
			<br>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<!--<input type="button" class="button" value="선택아이템저장" onClick="alert('일괄처리는 준비중입니다.');">-->
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
		<td width="50">이미지</td>
		<td width="50">상품코드</td>
		<td width="80">브랜드ID</td>
		<td>아이템명</td>
		<td width="30">거래<br>구분</td>
		<td width="70">등록일</td>
		<td width="200">요청사항<br>(요청사항을 수정하려면 마우스로 클릭하세요)</td>
		<td width="100">요청한한정수</td>
		<td width="30">거부</td>
		<td width="30">승인</td>
	</tr>
	<% for i=0 to oupcheitemedit.FResultCount -1 %>
	<form name="frmBuyPrc_<%= i %>" method="post" action="do_upche_req_itemmodify.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="idx" value="<%= oupcheitemedit.FItemList(i).Fidx %>">
	<input type="hidden" name="itemid" value="<%= oupcheitemedit.FItemList(i).FItemId %>">
	<input type="hidden" name="itemoption" value="<%= oupcheitemedit.FItemList(i).FItemOption %>">
	<input type="hidden" name="sellyn" value="<%= oupcheitemedit.FItemList(i).FSellYn %>">
	<input type="hidden" name="limityn" value="<%= oupcheitemedit.FItemList(i).FLimitYn %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td rowspan="2"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td rowspan="2"><img src="<%= oupcheitemedit.FItemList(i).FImageSmall %>" width="50" height="50" ></td>
		<td rowspan="2">
		    <a href="javascript:PopItemSellEdit('<%= oupcheitemedit.FItemList(i).FItemId %>');"><%= oupcheitemedit.FItemList(i).FItemId %></a>
		    <br>
		    (<%= oupcheitemedit.FItemList(i).FItemOption %>)
		</td>
		<td rowspan="2"><%= oupcheitemedit.FItemList(i).FMakerId %></td>
		<td rowspan="2" align="left">
			<%= oupcheitemedit.FItemList(i).FItemName %>
		<% if oupcheitemedit.FItemList(i).FItemOptionName<>"" then %>
			<br><%= oupcheitemedit.FItemList(i).FItemOptionName %>
		<% end if %>
		</td>
		<td rowspan="2"><%= fnColor(oupcheitemedit.FItemList(i).Fmwdiv,"mw") %></td>
		<td><acronym title="<%= oupcheitemedit.FItemList(i).FRegDate %>"><%= left(oupcheitemedit.FItemList(i).FRegDate,10) %></acronym></td>
		<td>
        <% if (oupcheitemedit.FItemList(i).FItemOption = "0000") or (oupcheitemedit.FItemList(i).FItemOption = "XXXX") then %>
        <!-- 옵션이 없을경우 -->
        	<table width="100%" border="0" class="a">
        		<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						    	<td width="30">판매:</td>
							    <td width="40"><%= oupcheitemedit.FItemList(i).GetOldSellYnName %>-&gt;</td>
							    <td width="30"><a href="javascript:ChangeSellYN(frmBuyPrc_<%= i %>, divsellyn<%= i %>)"><div id="divsellyn<%= i %>"><% if (oupcheitemedit.FItemlist(i).FSellYn = "Y") then %>판매<% else %><font color=red>품절</font><% end if %></div></a></td>
							    <td>(현재:<%= oupcheitemedit.FItemList(i).GetCurrSellYnName %>)</td>
						    </tr>
					  	</table>
					</td>
				</tr>

				<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						    	<td width="30">한정:</td>
						        <td width="40"><%= oupcheitemedit.FItemList(i).GetOldLimitYnName %>-&gt;</td>
						        <td width="30"><a href="javascript:ChangeLimitYN(frmBuyPrc_<%= i %>, divlimityn<%= i %>)"><div id="divlimityn<%= i %>"><% if (oupcheitemedit.FItemlist(i).FLimitYn = "N") then %>일반<% else %><font color=red>한정</font><% end if %></div></a></td>
						        <td>(현재:<%= oupcheitemedit.FItemList(i).GetCurrLimitYnName %>)</td>
						    </tr>
					    </table>
					</td>
				</tr>

				<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						        <td>
						        	한정수량설정:<input type="text" class="text" name="limitSetno" value="<%= oupcheitemedit.FItemList(i).FLimitNo %>" size="2" <%= chkIIF(oupcheitemedit.FItemlist(i).FLimitYn= "N","disabled style='background-color:#CCCCCC'","") %> >
						     	</td>
						    </tr>
					    </table>
					</td>
				</tr>
		  </table>
		<% else %>
		<!-- 옵션이 있을경우  -->
        	<table width="100%" border="0" class="a">
        		<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						    	<td width="30"><strong>옵션사용</strong>:</td>
							    <td width="40"><%= oupcheitemedit.FItemList(i).GetOldOptUsingYnName %>-&gt;</td>
							    <td width="30"><a href="javascript:ChangeSellYN(frmBuyPrc_<%= i %>, divsellyn<%= i %>)"><div id="divsellyn<%= i %>"><% if (oupcheitemedit.FItemlist(i).FSellYn = "Y") then %>판매<% else %><font color=red>품절</font><% end if %></div></a></td>
							    <td>(현재:<%= oupcheitemedit.FItemList(i).GetCurrSellYnName %>)</td>
						    </tr>
					  	</table>
					</td>
				</tr>
				<!--
				<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						    	<td width="30">한정:</td>
						        <td width="40"><%= oupcheitemedit.FItemList(i).GetOldLimitYnName %>-&gt;</td>
						        <td width="30"><a href="javascript:ChangeLimitYN(frmBuyPrc_<%= i %>, divlimityn<%= i %>)"><div id="divlimityn<%= i %>"><% if (oupcheitemedit.FItemlist(i).FLimitYn = "N") then %>일반<% else %><font color=red>한정</font><% end if %></div></a></td>
						        <td>(현재:<%= oupcheitemedit.FItemList(i).GetCurrLimitYnName %>)</td>
						    </tr>
					    </table>
					</td>
				</tr>
				-->
				<% if oupcheitemedit.FItemList(i).FLimitYn="Y" then %>
				<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						        <td>
						        	한정수량설정:<input type="text" class="text" name="limitSetno" value="<%= oupcheitemedit.FItemList(i).FLimitNo %>" size="2">
						     	</td>
						    </tr>
					    </table>
					</td>
				</tr>
				<% end if %>
		  </table>

        <% end if %>
		</td>

		<td>
			<%= oupcheitemedit.FItemList(i).FLimitNo %>
			-
			<%= oupcheitemedit.FItemList(i).FLimitSold %>
			=
			<%= oupcheitemedit.FItemList(i).GetRemainEa %>
			<br>
			(현재:<%= oupcheitemedit.FItemList(i).GetCurrRemainEa %>)
     		<br>
     		(실재고:<%= oupcheitemedit.FItemList(i).GetLimitStockNo %>)
			<!-- (<%= oupcheitemedit.FItemList(i).FCurrLimitNo %>-<%= oupcheitemedit.FItemList(i).FCurrLimitSold %>) -->
		</td>
		<td rowspan="2"><a href="javascript:DelThis(frmBuyPrc_<%= i %>)">거부</a></td>
		<td rowspan="2"><a href="javascript:AccThis(frmBuyPrc_<%= i %>)">승인</a></td>
	</tr>
	<tr>
		<td colspan="3" bgcolor="#FFFFFF">
			요청사유 : <%= oupcheitemedit.FItemList(i).FEtcStr %><br>
			거부사유 : <input type="text" class="text" name="rejectstr" value="" size="36" maxlength="64">
		</td>
	</tr>
	</form>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oupcheitemedit.HasPreScroll then %>
        		<a href="javascript:NextPage('<%= oupcheitemedit.StartScrollPage-1 %>')">[pre]</a>
        	<% else %>
        		[pre]
        	<% end if %>

        	<% for i=0 + oupcheitemedit.StartScrollPage to oupcheitemedit.FScrollCount + oupcheitemedit.StartScrollPage - 1 %>
        		<% if i>oupcheitemedit.FTotalpage then Exit for %>
        		<% if CStr(page)=CStr(i) then %>
        		<font color="red">[<%= i %>]</font>
        		<% else %>
        		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
        		<% end if %>
        	<% next %>

        	<% if oupcheitemedit.HasNextScroll then %>
        		<a href="javascript:NextPage('<%= i %>')">[next]</a>
        	<% else %>
        		[next]
        	<% end if %>
		</td>
	</tr>
</table>


<%
set oupcheitemedit = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->