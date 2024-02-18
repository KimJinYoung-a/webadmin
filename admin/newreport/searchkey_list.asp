<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim strSql
Dim arrRows,intLoop
Dim bitOrder : bitOrder = Trim(Request.Form("hidOrder"))
Dim rCount

 strSql = "select top 100 SearchKey, SearchKeyCount, RecentKeyCount "
 strSql = strSql + " from [db_search].[dbo].tbl_searchkey_statistics "
 If ( bitOrder = "1" ) Then 
 	strSql = strSql + " order by RecentKeyCount desc"
 Else	
 	strSql = strSql + " order by SearchkeyCount desc"
 End If
 	
 db2_rsget.Open strSql, db2_dbget
 
 If not db2_rsget.Eof Then
 	arrRows = db2_rsget.getRows()
 End If 
%>
<script language="javascript">
<!--
  function fnReOrder(varType){   
   document.frmOrder.hidOrder.value = varType;       
   document.frmOrder.submit();
  }
//-->
</script>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmOrder" method="post">
<input type="hidden" name="hidOrder" value="<%=bitOrder%>">
	<tr bgcolor="#EEEEEE">
	    <td align="center" width="10%">No</td>
	    <td align="center">검색어</td>
	    <td align="center"  width="20%" onClick="javascript:fnReOrder(0);" style="cursor:hand;">
	      총 검색횟수 <%If bitOrder = "" OR bitOrder = "0" THEN%>▼<%End If%>
	    </td>
	    <td align="center"  width="20%" onClick="javascript:fnReOrder(1);" style="cursor:hand;">
	     최근일주일 검색횟수<%If bitOrder = "1" THEN%>▼<%End If%>
	    </td>
	</tr>
<%For intLoop = 0 To UBound(arrRows,2)%>
	<tr bgcolor="#FFFFFF">
	    <td align="center"><%=intLoop+1%></td>
	    <td align="center"><%=arrRows(0,intLoop)%></td>
	    <td align="center"><%=arrRows(1,intLoop)%></td>
	    <td align="center"><%=arrRows(2,intLoop)%></td>
	</tr>
<%Next%>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db2close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->