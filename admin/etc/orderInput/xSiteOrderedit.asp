<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 주문내역 수정
' Hieditor : 2015.06.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incsessionadmin.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->

<%
dim voutmallorderseq, voutmallorderserial, vxSiteOrdercount, i
	voutmallorderseq = requestcheckvar(request("outmallorderseq"),10)

if voutmallorderseq="" then
	response.write "<script type='text/javascript'>alert('매칭키가 없습니다'); self.close();</script>"
	dbget.close()	:	response.end
end if

dim oorderedit
set oorderedit = new CxSiteTempOrder
    oorderedit.FPageSize=1
    oorderedit.FCurrPage=1
	oorderedit.frectoutmallorderseq = voutmallorderseq
	oorderedit.frectoutmallorderserial = voutmallorderserial
	oorderedit.fxsiteorderedit()

if oorderedit.ftotalcount=0 then
	response.write "<script type='text/javascript'>alert('주문건이 없습니다.'); self.close();</script>"
	dbget.close()	:	response.end
end if

if oorderedit.foneitem.fsellsite<>"cjmall" and oorderedit.foneitem.fsellsite<>"lotteimall" then
	response.write "<script type='text/javascript'>alert('CJMALL/Lotteimall 주문건만 수정이 가능합니다.'); self.close();</script>"
	dbget.close()	:	response.end
end if

if oorderedit.foneitem.fOrderSerial<>"" then
	response.write "<script type='text/javascript'>alert('기 입력 주문건.'); self.close();</script>"
	dbget.close()	:	response.end
end if

voutmallorderserial = oorderedit.foneitem.fOutMallOrderSerial

vxSiteOrdercount = getxSiteDuppReceiverCheck(voutmallorderserial)
if vxSiteOrdercount=0 then
	response.write "<script type='text/javascript'>alert('해당되는 주문건이 없습니다.'); self.close();</script>"
	dbget.close()	:	response.end
end if

if vxSiteOrdercount>1 then
	If session("ssBctID")<>"kjy8517" Then
		response.write "<script type='text/javascript'>alert('다수령지 체크 요망.'); self.close();</script>"
		dbget.close()	:	response.end		
	End If
end if
%>
<script type='text/javascript'>

	function chorderedit(){
	    var frm=document.frm;
/*
	    alert(frm.Org_OrderName.value.length);
	    alert(frm.OrderName.value.length);
*/
		if(frm.Org_OrderName.value.length <= 2 ){
		    if(frm.Org_OrderName.value.substring(0,1)!=frm.OrderName.value.substring(0,1)){
		        alert('기존 주문인과 일치하지 않습니다.');
		        frm.OrderName.focus();
		        return;
		    }
		}else{
		    if(frm.Org_OrderName.value.substring(0,2)!=frm.OrderName.value.substring(0,2)){
		        alert('기존 주문인과 일치하지 않습니다.');
		        frm.OrderName.focus();
		        return;
		    }
		}

	    if(frm.Org_ReceiveName.value.length <= 2 ){
		    if(frm.Org_ReceiveName.value.substring(0,1)!=frm.ReceiveName.value.substring(0,1)){
		        alert('기존 수령인과 일치하지 않습니다.');
		        frm.ReceiveName.focus();
		        return;
		    }
	    }else{
		    if(frm.Org_ReceiveName.value.substring(0,2)!=frm.ReceiveName.value.substring(0,2)){
		        alert('기존 수령인과 일치하지 않습니다.');
		        frm.ReceiveName.focus();
		        return;
		    }
		}
/*
		if(frm.Org_OrderName.value.length !=  frm.OrderName.value.length){
			alert('기존 주문인과 변경한 이름의 길이가 다릅니다');
			frm.OrderName.focus();
			return;
		}

		if(frm.Org_ReceiveName.value.length !=  frm.ReceiveName.value.length){
			alert('기존 수령인과 변경한 이름의 길이가 다릅니다');
			frm.ReceiveName.focus();
			return;
		}
*/

		if(confirm("[제휴주문번호 : <%= voutmallorderserial %>]\n같은 주문건이 <%= vxSiteOrdercount %>건 존재 합니다. 수정하시겠습니까?")){
			frm.mode.value="orderedit"
			frm.submit();
		}
	}
	
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			※ 같은 주문건이 총 <%= vxSiteOrdercount %>건 있습니다.
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post" action="/admin/etc/orderinput/xSiteOrderprocess.asp">
<input type="hidden" name="mode">
<input type="hidden" name="outmallorderseq" value="<%=voutmallorderseq%>">
<input type="hidden" name="outmallorderserial" value="<%=voutmallorderserial%>">
<tr bgcolor="#FFFFFF">
	<td align="center">판매쇼핑몰</td>
	<td colspan="2" >
		<%= oorderedit.foneitem.fsellsite %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">제휴주문번호</td>
	<td colspan="2"  >
		<%= oorderedit.foneitem.fOutMallOrderSerial %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"></td>
	<td>주문자</td>
	<td>수령인</td>	
	
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">현재</td>
	<td><%= oorderedit.foneitem.fOrderName %></td>
	<td><%= oorderedit.foneitem.fReceiveName %></td>	
	
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">변경</td>
	<td>
	    <input type="hidden" name="Org_OrderName" value="<%= oorderedit.foneitem.fOrderName %>">
	    <input type="text" name="OrderName" value="<%= oorderedit.foneitem.fOrderName %>" size="10">
	</td>
	<td>
	    <input type="hidden" name="Org_ReceiveName" value="<%= oorderedit.foneitem.fReceiveName %>">
	    <input type="text" name="ReceiveName" value="<%= oorderedit.foneitem.fReceiveName %>" size="10">
	</td>	
	
</tr>


<tr bgcolor="#FFFFFF">
	<td align="center" colspan="3">
		<input type="button" value="수정" onclick="chorderedit();" class="button" >
	</td>
</tr>
</form>
</table>

<%
set oorderedit = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->