<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���α� ��� ���
' History : 2008.01.15 ������ ����
'			2016.07.21 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/BonusCouponSummaryClass.asp"-->

<%
dim issuedcount, usingcount, spendcoupon, subtotalprice   , spendmileage, i, NotExpiredExists

'//������,�ý������� ��������� (2011.06.21; ������)
if Not((session("ssAdminPsn")=7) or (session("ssAdminPsn")=14) or (session("ssAdminPsn")=22) or (session("ssAdminPsn")=30) or (session("ssAdminPsn")=11)) then
	response.write "������"
	response.end
end if

dim page, mode, yyyymm, yyyy1,mm1, stdate, couponidx, userlevel, chkMonth
	mode        = request("mode")
	page        = request("page")
	chkMonth    = request("chkMonth")
	yyyy1       = request("yyyy1")
	mm1	        = request("mm1")
	couponidx   = request("couponidx")
	userlevel   = request("userlevel")

if (page="") then page=1

if yyyy1="" then
	stdate = CStr(Now)
	stdate = DateSerial(Left(stdate,4), CLng(Mid(stdate,6,2)),1)
	yyyy1 = Left(stdate,4)
	mm1 = Mid(stdate,6,2)
end if

dim sqlStr
if (mode="refreshSummary") then
    if (couponidx="246") or (couponidx="335") or (couponidx="380") or (couponidx="404") or (couponidx="418") or (couponidx="419") or (couponidx="568") or (couponidx="577") then
        rw "���� ������Ʈ - ������ ���� ���"  '''
        response.end
        ''sqlStr = "exec db_datamart.dbo.sp_ten_dataMart_mkt_bonuscoupon_result " & couponidx & ",'" & yyyy1 & "-" & mm1 & "'"
        ''db3_dbget.Execute sqlStr   '''���� ����..//���Ѿ���..
    else
       ' sqlStr = "exec db_summary.dbo.sp_ten_dataMart_mkt_bonuscoupon_result " & couponidx & ",'" & yyyy1 & "-" & mm1 & "'"
       ' dbget.Execute sqlStr
       
       ''2015/11/18���� //_V3 2016/12/21 
       sqlStr = "exec db_datamart.dbo.sp_ten_dataMart_mkt_bonuscoupon_result_V3 " & couponidx & ",''"
       ''rw "������..."
       db3_dbget.Execute sqlStr
       
       
    end if
    
    
''   ''����Ÿ ��Ʈ�������� �������� ���� 20120727-eastone :: ������ ����ȵ�..
''   sqlStr = "exec db_datamart.dbo.sp_ten_dataMart_mkt_bonuscoupon_result " & couponidx & ",'" & yyyy1 & "-" & mm1 & "'"
''   if (couponidx<>"335") then
''       rw sqlStr
''       rw "������"
''       response.end
''    end if
''    db3_dbget.Execute sqlStr 

end if

dim oCouponSummary
set oCouponSummary = new CBonusCouponSummary
oCouponSummary.FPageSize = 100
oCouponSummary.FCurrpage = page

if (chkMonth<>"") then
    oCouponSummary.FRectYYYYMM = yyyy1 + "-" + mm1
end if

oCouponSummary.FRectCouponidx  = couponidx
oCouponSummary.FRectUserLevel  = userlevel
oCouponSummary.getCouponResultSummary

NotExpiredExists = false
%>
<script type="text/javascript">

function goPage(ipage){
    frm.page.value=ipage;
    frm.submit();
}

function refreshSummary(){
    var frm = document.frmSubmit;
    
    if (confirm('��踦 ���ۼ� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
    	<input type="checkbox" name="chkMonth" <% if chkMonth<>"" then response.write "checked" %> > 
    	����� : <% DrawYMBox yyyy1,mm1 %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ���� ��ȣ : <input type="text" name="couponidx" value="<%= couponidx %>" size="5" maxlength="9">
		&nbsp;
		* ����ڷ��� :
		<% DrawselectboxUserLevel "userlevel",  userlevel, "" %>
	</td>
</tr>
</table>
<!-- �˻� �� -->

</form>

<Br>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oCouponSummary.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oCouponSummary.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="90">�����</td>
	<td width="80">������ȣ</td>
  	<td width="160">������</td>
  	<td width="90">����</td>
  	<td width="60">�����</td>
  	<td width="60">����</td>
  	<td width="60">�����</td>
  	<td width="90">����</td>
  	<td width="90">���Ǹ���<br>(���ϸ�������)</td>
  	<td width="90">���ϸ�������</td>
  	<td>���</td>
</tr>

<% if oCouponSummary.FresultCount>0 then %>
	<%
	for i=0 to oCouponSummary.FResultCount -1

	issuedcount     = issuedcount + oCouponSummary.FItemList(i).Fissuedcount
	usingcount      = usingcount + oCouponSummary.FItemList(i).Fusingcount
	spendcoupon     = spendcoupon + oCouponSummary.FItemList(i).Fspendcoupon
	subtotalprice   = subtotalprice + oCouponSummary.FItemList(i).Fsubtotalprice
	spendmileage    = spendmileage + oCouponSummary.FItemList(i).Fspendmileage
	%>
	<tr bgcolor="#FFFFFF">
    <td align="center"><%= oCouponSummary.FItemList(i).FregYYYYMM %></td>
    <td align="center"><%= oCouponSummary.FItemList(i).Fbonuscouponidx %></td>
    <td><%= oCouponSummary.FItemList(i).Fbonuscouponname %></td>
    <td align="center"><font color="<%= getUserLevelColor(oCouponSummary.FItemList(i).Fuserlevel) %>"><%= getUserLevelStr(oCouponSummary.FItemList(i).Fuserlevel) %></font></td>
    <td align="center"><%= FormatNumber(oCouponSummary.FItemList(i).Fissuedcount,0) %></td>
    <td align="center"><%= FormatNumber(oCouponSummary.FItemList(i).Fusingcount,0) %></td>
    <td align="center"><%= oCouponSummary.FItemList(i).getUsingPro() %>%</td>
    <td align="right"><%= FormatNumber(oCouponSummary.FItemList(i).Fspendcoupon,0) %></td>
    <td align="right"><%= FormatNumber(oCouponSummary.FItemList(i).Fsubtotalprice,0) %></td>
    <td align="right"><%= FormatNumber(oCouponSummary.FItemList(i).Fspendmileage,0) %></td>
    <td>
	    <% if oCouponSummary.FItemList(i).FNotExpiredCount<1 then %>
	    
	    <% else %>
	        <% NotExpiredExists = true %>
	        ��������
	    <% end if %>
	    </td>
	</tr>
	<% next %>

	<tr bgcolor="#FFFFFF">
	    <td align="center">�հ�</td>
	    <td align="center"></td>
	    <td align="center"></td>
	    <td align="center"></td>
	    <td align="center"><%= FormatNumber(issuedcount,0) %></td>
	    <td align="center"><%= FormatNumber(usingcount,0) %></td>
	    <td align="center">
		    <% if issuedcount<>0 then %>
		        <%= CLng(usingcount/issuedcount*100*100)/100 %>%
		    <% end if %>
	    </td>
	    <td align="right"><%= FormatNumber(spendcoupon,0) %></td>
	    <td align="right"><%= FormatNumber(subtotalprice,0) %></td>
	    <td align="right"><%= FormatNumber(spendmileage,0) %></td>
	    <td align="center">
		    <% if ( (NotExpiredExists) and (couponidx<>"") ) then %>
		        <a href="javascript:refreshSummary();"><img src="/images/button_reload.gif" width="60" border="0"></a>
		    <% end if %>
	    </td>
	</tr>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<%
			if oCouponSummary.HasPreScroll then
				Response.Write "<a href='javascript:goPage(" & oCouponSummary.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
			else
				Response.Write "[pre] &nbsp;"
			end if
	
			for i=0 + oCouponSummary.StartScrollPage to oCouponSummary.FScrollCount + oCouponSummary.StartScrollPage - 1
	
				if i>oCouponSummary.FTotalpage then Exit for
	
				if CStr(page)=CStr(i) then
					Response.Write " <font color='red'>[" & i & "]</font> "
				else
					Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
				end if
	
			next
	
			if oCouponSummary.HasNextScroll then
				Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
			else
				Response.Write "&nbsp; [next]"
			end if
			%>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
	    <td colspan="15">
	    	<%
	    		'// 6���� �����͸� ���÷��� ����
	    		if datediff("m",yyyy1 & "-" & mm1 & "-01",date)<6 then
	    	%>
	        <a href="javascript:refreshSummary();"><img src="/images/button_reload.gif" width="60" border="0"></a>
	        <% end if %>
	    </td>
	</tr>
<% end if %>

</table>

<%
set oCouponSummary = Nothing
%>
<form name="frmSubmit" method="post" >
<input type="hidden" name="mode" value="refreshSummary">
<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
<input type="hidden" name="mm1" value="<%= mm1 %>">
<input type="hidden" name="couponidx" value="<%= couponidx %>">
<input type="hidden" name="chkMonth" value="<%= chkMonth %>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
