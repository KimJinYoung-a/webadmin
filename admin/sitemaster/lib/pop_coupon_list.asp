<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : pop_coupon_list
' Description :  ���� ����Ʈ
' History : 2019-09-09 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
function getCouponList()

	dim SqlStr, intLoop 

	sqlStr = sqlStr & " SELECT TOP 150  "
	sqlStr = sqlStr & " 	IDX "
	sqlStr = sqlStr & " 	, couponvalue "
	sqlStr = sqlStr & " 	, couponname "
	sqlStr = sqlStr & " 	, minbuyprice "
	sqlStr = sqlStr & " 	, startdate "
	sqlStr = sqlStr & " 	, expiredate "
	sqlStr = sqlStr & " FROM DB_USER.DBO.tbl_user_coupon_master "
	sqlStr = sqlStr & " ORDER BY IDX DESC "

'       response.write sqlStr &"<br>"
'       response.end
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF then
		getCouponList = rsget.getRows()    
	end if
	rsget.close         
End function	
%>
<%
dim arrList, intLoop

arrList = getCouponList()
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language="javascript">
	function setCouponIdx(idx){
		if(confirm("������ ����Ͻðڽ��ϱ�?")){
			$(opener.document).find("#couponidx").val(idx)
			self.close();
		}		
	}
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle">���� ����Ʈ</div>
<table width="500" border="0" align="left" class="a" cellpadding="3" cellspacing="0" >
<tr>
	<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>" >
		<td align="center" width="10%">�ڵ�</td>	
		<td align="center" width="15%">��������</td>
		<td align="center">�����̸�</td>	
		<td align="center" width="15%">�ּұ��ž�</td>	
		<td align="center" width="15%">������</td>	
		<td align="center" width="15%">������</td>	
	</tr>

	<%IF isArray(arrList) THEN 
		For intLoop = 0 To UBound(arrList,2)
		%>
		<tr bgcolor="#FFFFFF" onClick="setCouponIdx(<%=arrList(0,intLoop)%>);" style="cursor:hand;" onMouseOver="this.style.backgroundColor='#FFFFEC'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
			<td  align="center"><%=arrList(0,intLoop)%></td>		
			<td  align="center"><%=arrList(1,intLoop)%></td>
			<td><%=arrList(2,intLoop)%></td>
			<td  align="center"><%=arrList(3,intLoop)%></td>
			<td  align="center"><%=arrList(4,intLoop)%></td>
			<td  align="center"><%=arrList(5,intLoop)%></td>
		</tr>
		<% Next %>	
	<% end if %>
	</table>	
</tr>

</form>
</table>


<!-- #include virtual="/lib/db/dbclose.asp" -->