<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim sqlStr, resultRow, vBrandID
	vBrandID = Request("brandid")
	If vBrandID = "" Then
		Response.write "<script >alert('브랜드를 지정하세요.')</script>"
		Response.End
	End If
	sqlStr = "UPDATE s " & _
			 "		SET s.interparklastupdate = dateadd(hh,-1,i.lastupdate) " & _
			 "	FROM [db_item].[dbo].tbl_interpark_reg_item s " & _
			 "		Left JOIN [db_item].[dbo].tbl_item i on i.itemid = s.itemid " & _
			 "		Left JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping p on i.cate_large=p.tencdl and i.cate_mid=p.tencdm and i.cate_small=p.tencdn " & _
			 "	WHERE " & _
			 "		i.makerid = '" & vBrandID & "' " & _
			 "		and s.interparkPrdNo is Not NULL " & _
			 "		and ((i.lastupdate<>'2008-10-21 00:06:16.140')) " & _
			 "		and i.basicimage is not null " & _
			 "		and i.itemdiv<50 " & _
			 "		and i.cate_large<>'' " & _
			 "		and i.cate_large<>'999' " & _
			 "		and i.sellcash>0 " & _
			 "		and i.makerid NOT IN (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun = 'interpark') " & _
			 "		and ((i.sellcash-i.buycash)/i.sellcash)*100>=11 " & _
			 "		and i.itemid<>171124 " & _
			 "		and i.itemid<>171659 " & _
			 "		and i.itemid<>171658 " & _
			 "		and i.itemid<>172515 " & _
			 "		and i.itemid<>172794 " & _
			 "		and p.SupplyCtrtSeq is Not NULL "
	dbget.Execute sqlStr, resultRow
	response.write "<script >alert('" + CStr(resultRow) + "건 수정되었습니다.')</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->