<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->

<%
Dim mode	: mode = requestCheckVar(request("mode",""),32)
Dim CsAsID	: CsAsID = requestCheckVar(request("masteridx",""),10)

Dim songjangDiv	: songjangDiv = requestCheckVar(request("songjangDiv",""),3)
Dim songjangNo	: songjangNo  = requestCheckVar(request("songjangNo",""),32)

If mode = "SONGJANG" Then 

	dim sqlStr
	sqlStr = "UPDATE db_shop.dbo.tbl_shopbeasong_cs_master SET" & VbCrlf	
	sqlStr = sqlStr + " songjangDiv ="&songjangDiv&"" & VbCrlf
	sqlStr = sqlStr + " , songjangNo ='"&songjangNo&"'" & VbCrlf
	sqlStr = sqlStr + " , currState	  = (case when currState < 'B004' then 'B004' else currState end)" & VbCrlf
    sqlStr = sqlStr + " WHERE masteridx =" & CsAsID
    
    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
    
	response.write "<script>" & vbCrLf
	response.write "alert('등록되었습니다.');" & vbCrLf
	response.write "opener.location.reload();" & vbCrLf
	response.write "window.close();" & vbCrLf
	response.write "</script>" & vbCrLf
	dbget.close()	:	response.End 
End If 

Sub drawSelectBoxDeliverCompany(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Trim(Lcase(selectedId)) = Trim(Lcase(rsget("divcd"))) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub


dim ocsaslist
set ocsaslist = New corder
	ocsaslist.FRectCsAsID = CsAsID
	
	if (CsAsID<>"") then
	    ocsaslist.fGetOneCSASMaster
	end if

if (ocsaslist.ftotalcount < 1) then
    response.write "등록된 as 내역이 없습니다"
    response.end
end if
%>

<script language="javascript">

function jsSubmit()
{
	var f = document.frmWrite;

	if (!f.songjangDiv.value)
	{
		alert("택배회사를 선택해 주세요.");
		f.songjangDiv.focus();
		return;
	}
	if (!f.songjangNo.value || f.songjangNo.value.length < 8)
	{
		alert("송장번호를 입력해 주세요.");
		f.songjangNo.focus();
		return;
	}

	f.submit();
}

</script>

<!---- 팝업크기 400x195 ---->
<form name="frmWrite" action="popChangeSongjang.asp">
<input type="hidden" name="mode" value="SONGJANG">
<input type="hidden" name="masteridx" value="<%=CsAsID%>">
<table width="400" border="0" cellspacing="0" cellpadding="0">
<tr>
	<!---- 팝업제목 시작 ---->
	<td valign="top" bgcolor="#af1414">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td align="right"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/popup_logo.gif" width="254" height="15"></td>
		</tr>
		<tr>
			<td height="42" valign="bottom" style="padding:0 0 5px 20px"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/title_invoice.gif" width="107" height="23"></td>
		</tr>
		</table>
	</td>
	<!---- 팝업제목 끝 ---->
</tr>
<tr>
	<td><br><Br></td>
</tr>
<tr>
	<td align="center" class="gray11px02" style="padding:0 0 20px 0px;">
		<table border="0" cellspacing="0" cellpadding="0" style="border-top:3px solid #be0808;" class="a">
		<tr>
			<td width="100" height="31" bgcolor="#fcf6f6" class="bbstxt01" style="border-bottom:1px solid #eaeaea;">택배 회사</td>
			<td width="200" style="border-bottom:1px solid #eaeaea;padding:0 1px 0 20px;">
				<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td>
						<%Call drawSelectBoxDeliverCompany("songjangDiv",ocsaslist.FOneItem.FsongjangDiv)%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="100" height="31" bgcolor="#fcf6f6" class="bbstxt01" style="border-bottom:1px solid #eaeaea;">택배 번호</td>
			<td style="border-bottom:1px solid #eaeaea;padding:0 1px 0 20px;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><input name="songjangNo" type="text" class="input_02" style="width:140px;height:20px;" value="<%=ocsaslist.FOneItem.FsongjangNo%>" /></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" style="padding-bottom:10px;">
		<table border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td style="padding-right:7px;"><a href="javascript:jsSubmit();" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/btn_confirm.gif" width="58" height="24" border="0"/></a></td>
			<td><a href="javascript:window.close();" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2009/order/btn_cancel02.gif" width="58" height="24" border="0"/></a></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<%
set ocsaslist = Nothing
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->