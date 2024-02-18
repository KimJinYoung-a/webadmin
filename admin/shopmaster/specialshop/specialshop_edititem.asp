<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 우수회원샵
' Hieditor : 2009.12.28 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/specialshop/specialshop_cls.asp"-->

<%
dim id , i
	id = request("id")
	
if id = "" then
	response.write "<script>alert(id값이 없습니다'); self.close();</script>"
end if	
	
dim ospecialshop_item
set ospecialshop_item = new cspecialshop_list
	ospecialshop_item.frectid = id
	ospecialshop_item.fspecialshop_itemlist()	
%>

<script language="javascript">

	function SaveArr(){
		if (frm.itemidarr.value==''){
			alert('상품코드를 입력 하세요');
			frm.itemidarr.focus();
			return;
		}
		
		frm.mode.value='itemadd';	
		frm.action='/admin/shopmaster/specialshop/specialshop_process.asp';
		frm.submit();		
	}

	function dellitem(idx){	
		frm.mode.value='dellitem';	
		frm.idx.value=idx;	
		frm.action='/admin/shopmaster/specialshop/specialshop_process.asp';
		frm.submit();		
	}

</script>

<!-- 액션 시작 -->
※상품등록
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="post" action="">
<input type="hidden" name="mode">
<input type="hidden" name="id" value="<%=id %>">
<input type="hidden" name="idx" >	
	<tr>
		<td>
			<!-- input type=text name="itemidarr" size="30" maxlength=64 -->
			<textarea name="itemidarr" rows="5"></textarea>
			<input type=button value="상품추가" onClick="SaveArr(frm)" class="button">
			<br>(마우스로 긁어서 복사해서 붙여넣으면 눈에 보이진 않으나 특수문자가 복사될 수 있습니다. 그러면 에러납니다.)
		</td>
		<td class="a" align="right">
		</td>
	</tr>
</form>	
</table>
<!-- 액션 끝 -->
<br><font color="#3366FF">+ 할인율: BLUE 15%, VIP silver 20%, VIP gold 25%, STAFF 25%, FAMILY 20%</font>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ospecialshop_item.ftotalcount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ospecialshop_item.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   		
		<td align="center">상품코드</td>
		<td align="center">이미지</td>				
		<td align="center">상품명</td>
		<td align="center">품절여부</td>
		<td align="center">판매가<br/><font color="#3366FF">최대할인가(25%)</font></td>
		<td align="center">공급가</td>
		<td align="center">마진</td>
		<td align="center">비고</td>	
    </tr>
	<% for i=0 to ospecialshop_item.ftotalcount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			
	
    <% if ospecialshop_item.FItemList(i).fstatus = "3" then %>    
    <tr align="center" bgcolor="#FFFFaa">
    <% else %>    
    <tr align="center" bgcolor="#FFFFFF">
	<% end if %>	
		<td align="center">
			<a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%=ospecialshop_item.FItemList(i).fitemid%>" onfocus="this.blur()" target="_blink"><%=ospecialshop_item.FItemList(i).fitemid%></a>
		</td>	
		<td align="center">
			<img src="<%=ospecialshop_item.FItemList(i).FImageSmall%>" width=50 height=50>
		</td>		
		<td align="center">
			<%=ospecialshop_item.FItemList(i).fitemname%>
		</td>		
		<td align="center">
			<%
			if ospecialshop_item.FItemList(i).fsellyn <> "Y" then 
				response.write "품절"
			else
				response.write "판매중"
			end if				 
				
			%>
		</td>
		
		<td class="verdana-small">		 
			<% if ospecialshop_item.FItemList(i).IsSail then %><font color="#F08050"><% end if %><%= FormatNumber(ospecialshop_item.FItemList(i).FSellCash,0) %>원</font>
			 <br>  <font color="#3366FF"><%= FormatNumber(ospecialshop_item.FItemList(i).getRealPrice ,0) %>원</font>
			 
		</td>
		<td  class="verdana-small">	
			<%=FormatNumber(ospecialshop_item.FItemList(i).FBuyCash,0)%>원
		</td>
		<td class="verdana-small"><%IF ospecialshop_item.FItemList(i).getMargin < 0 then%><font color="red"><%end if%><%=FormatNumber(ospecialshop_item.FItemList(i).getMargin,1)%>%</td>
		<td align="center">
			<input type="button" class="button" value="삭제" onclick="dellitem(<%=ospecialshop_item.FItemList(i).fidx%>);">
		</td>
    </tr>   
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[등록되어 있는 상품이 없습니다.]</td>
		</tr>
	<% end if %>

</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->