<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 오프라인 메일진
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<%
dim masteridx, gubun
masteridx = requestCheckVar(request("masteridx"),10)
gubun = requestCheckVar(request("gubun"),10)

%>
<script language='javascript'>
function popItemWindow(iid,frm){
	if (frmarr.masteridx.value == "")	{
		alert("메일진 구분을 선택해주세요...");
		frmarr.masteridx.focus();
		return;
	}
	if (frmarr.gubun.value == "")	{
		alert("On-Off구분을  선택해주세요...");
		frmarr.gubun.focus();
		return;
	}
	else{
	var v;
	v=frmarr.gubun.value;
	window.open("/admin/pop/viewitemlist.asp?designerid=" + iid + "&target=" + frm, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	}
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function AddIttems(){
	var ret = confirm(frmarr.itemid.value + '아이템을 추가하시겠습니까?');
	if (ret){
		frmarr.itemid.value = frmarr.itemid.value;
		frmarr.gubun.value = frmarr.gubun.value;
		frmarr.mode.value="add";
		frmarr.submit();
	}
}

function AddIttems2(){

	if (confirm(frmarr.itemidarr.value + '아이템을 추가하시겠습니까?')){
		frmarr.itemid.value = frmarr.itemidarr.value;
		frmarr.gubun.value = frmarr.gubun.value;
		frmarr.mode.value="add";
		frmarr.submit();
	}
}

</script>
<table width="650" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frmarr" method="post" action="/admin/offshop/lib/domailzinebestitem.asp">
	<input type="hidden" name="mode">
	<input type="hidden" name="itemid">
	<tr>
		<td class="a" >
		메일진구분 : <% DrawSelectBoxMailzine masteridx %>
		On-Off 구분 : 
			<select name=gubun>
				<option value="" <% if gubun="" then response.write "selected" %>>선택</option>
				<option value="01" <% if gubun="01" then response.write "selected" %>>On-line Best</option>
				<option value="02" <% if gubun="02" then response.write "selected" %>>Off-line Best</option>
			</select>
			<input type="button" value="아이템 추가" onclick="popItemWindow('','frmarr.itemid')" class="button">
		</td>
	</tr>
	<tr>
		<td class="a">
			<table width=100% border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td><input type="text" name="itemidarr" value="" size="76" class="input"></td>
				<td width="100" align="right"><input type="button" value="아이템 직접 추가" onclick="AddIttems2()" class="button"></td>
			</tr>
			</table><br>(마지막에 콤마(,)를 넣어주세요 ex:41080,40780,40759,)
		</td>
	</tr>
	</form>
</table>

<%
'메일진 선택
Sub DrawSelectBoxMailzine(byval selectedId)
   dim tmp_str,query1
   %><select name="masteridx" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select idx,regdate from [db_shop].[dbo].tbl_shopmaster_mail"
   query1 = query1 + " where isusing = 'Y'"
   query1 = query1 + " order by regdate desc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("idx")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("idx")&"' "&tmp_str&">"&FormatDate(rsget("regdate"),"0000.00.00")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->