<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트플러스
' History : 2010.04.05 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/giftplus/giftplus_cls.asp"-->
<%
dim cdL,cdM,cdS,arrItemID ,ecdL,ecdM , ecdS,mode ,objMnL,objMnM,objMnS ,i ,objView
	cdL= request("cdL")
	cdM= request("cdM")
	cdS= request("cdS")
	ecdL = request("ecdL")
	ecdM = request("ecdM")
	ecdS = request("ecdS")
	mode = request("mode")
	arrItemID = chkarray(request("arrItemID"))

set objView = new giftManagerView
	objView.getMenuView cdL,cdM,cdS
%>

<script language="javascript">

function selCate(va){

	document.UpdateFRM.action="?";
	document.UpdateFRM.submit();
}

function subchk(){
	
	var frm = document.UpdateFRM;
	var conf;
	if(frm.cdL.value!=""&&frm.cdM.value!=""&&frm.cdS.value!=""){
		
		if(frm.EcdL.value!=""&&frm.EcdM.value!=""&&frm.EcdS.value!=""){
			
			for (var i = 0 ;frm.mode.length;i++){
				
				if (frm.mode[i].checked){
					
					if(frm.mode[i].value=="copy"){
					
						if(confirm("상품을 복사합니다.")){
							frm.action="Item_Process.asp";
							frm.submit();
						}
					
					} else {
					
						if(confirm("상품을 이동합니다.")){
							frm.action="Item_Process.asp";
							frm.submit();
						}
						
					}
				}
			}
			
			
		}else{
			alert("대상 카테고리를 선택해주세요");
			return false;
		}
	}else {
		alert("원본 카테고리를 다시 선택해 주세요");
		self.close();
		return false;
	}

}

</script>

<table width="600" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="UpdateFRM" action="" target="" onsubmit="return subchk();">
<input type="hidden" name="arrItemID" value="<%= arrItemID %>">

<tr>
	<td width="130" bgcolor="<%= adminColor("tabletop") %>" align="center">작업 선택</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="mode" value="copy" <% IF mode ="copy" or mode="" then response.write "checked" %>> 복사 
		<input type="radio" name="mode" value="move" <% IF mode ="move" then response.write "checked" %>>이동</td>
</tr>
<% IF objView.LCode <>"" then %>
<input type="hidden" name="cdL" size="4" value="<%= objView.LCode %>" />
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">대 카테고리</td>
	<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.LCode %></font>] <%= objView.LCodeNm %>
</tr>
<% END IF %>

<% IF objView.MCode <>"" then %>
<input type="hidden" name="cdM" size="4" value="<%= objView.MCode %>" />
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">중 카테고리</td>
	<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.MCode %></font>] <%= objView.MCodeNm %>
</tr>
<% END IF %>

<% IF objView.SCode <>"" then %>
<input type="hidden" name="cdS" size="4" value="<%= objView.SCode %>" />
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">소 카테고리</td>
	<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.SCode %></font>] <%= objView.SCodeNm %>
</tr>
<% END IF %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">대상 카테고리</td>
	<td bgcolor="#FFFFFF">
		<%
			
			SET objMnL = NEW giftManagerMenu
			objMnL.getMenuListLarge 
		%>
		<select name="EcdL" onchange="selCate(this.value);">
			<option value="">대 카테고리 선택</option>
			<% FOR i =0 TO objMnL.FResultCount -1 %>
			<option value="<%= objMnL.FItemList(i).LCode %>" <% if EcdL =objMnL.FItemList(i).LCode then response.write "selected" %>><%= objMnL.FItemList(i).LCodeNm %></option>
			<% NEXT %>
		</select>
		<% SET objMnL = NOTHING %>
		
		
		<%
			
		SET objMnM = NEW giftManagerMenu
		objMnM.FRectCDL = EcdL
		objMnM.getMenuListMid
		
		
		%>
		<select name="EcdM" onchange="selCate(this.value);">
			<option value="">중 카테고리 선택</option>
			<% 
			IF objMnM.FResultcount >0 THEN 
				FOR i =0 TO objMnM.FResultCount -1 %>
			<option value="<%= objMnM.FItemList(i).MCode %>" <% if EcdM =objMnM.FItemList(i).MCode then response.write "selected" %>><%= objMnM.FItemList(i).MCodeNm %></option>
			<% 
				NEXT 
			END IF %>
		</select>
		<% SET objMnM = NOTHING %>
		
		<%
			
		SET objMnS = NEW giftManagerMenu
		objMnS.FRectCDL = EcdL
		objMnS.FRectCDM = EcdM
		objMnS.getMenuListSmall
		
		%>
		<select name="EcdS" onchange="selCate(this.value);">
			<option value="">소 카테고리 선택</option>
			<% 
			IF objMnS.FResultCount >0 THEN
				FOR i =0 TO objMnS.FResultCount -1 %>
			<option value="<%= objMnS.FItemList(i).SCode %>" <% if EcdS =objMnS.FItemList(i).SCode then response.write "selected" %>><%= objMnS.FItemList(i).SCodeNm %></option>
			<% 
				NEXT 
			END IF %>
		</select>
		<% SET objMnS = NOTHING %>
		
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="submit" class="button" value="적용"></td>
</tr>
</form>
</table> 

<% 
set objView = nothing 
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->