<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%
Dim refer
refer = request.ServerVariables("HTTP_REFERER")
Dim mode, tecdl, tecdm, tecdn
Dim interparkdispcategory, SupplyCtrtSeq, interparkstorecategory
Dim sqlStr
Dim oldDispCate
mode					= request("mode")
tecdl					= request("tecdl")
tecdm					= request("tecdm")
tecdn					= request("tecdn")
interparkdispcategory	= request("interparkdispcategory")
SupplyCtrtSeq			= request("SupplyCtrtSeq")
interparkstorecategory	= request("interparkstorecategory")

If (mode="cateedit") Then
	''카테고리가 변경된 경우 수정해야함 -> 수정일 변경
	oldDispCate = ""
	sqlStr = ""
	sqlStr = sqlStr & " SELECT interparkdispcategory from [db_item].[dbo].tbl_interpark_dspcategory_mapping"
	sqlStr = sqlStr & " WHERE tencdl='" & tecdl & "'"
	sqlStr = sqlStr & " and tencdm='" & tecdm & "'"
	sqlStr = sqlStr & " and tencdn='" & tecdn & "'"
	rsget.Open sqlStr,dbget,1
	If Not rsget.Eof Then
		oldDispCate = rsget("interparkdispcategory")
	End If    
	rsget.Close
	
	sqlStr = ""
	sqlStr = sqlStr & " IF Exists(SELECT * FROM [db_item].[dbo].tbl_interpark_dspcategory_mapping WHERE tencdl='"&tecdl&"' and tencdm='"&tecdm&"' and tencdn='"&tecdn&"')"
	sqlStr = sqlStr & " BEGIN"
	sqlStr = sqlStr & "     UPDATE [db_item].[dbo].tbl_interpark_dspcategory_mapping "
	sqlStr = sqlStr & "     SET interparkdispcategory='" & interparkdispcategory & "'"
	sqlStr = sqlStr & "     ,SupplyCtrtSeq=" & SupplyCtrtSeq & ""
	sqlStr = sqlStr & "     ,interparkstorecategory='" & interparkstorecategory & "'"
	sqlStr = sqlStr & "     where tencdl='" & tecdl & "'"
	sqlStr = sqlStr & "     and tencdm='" & tecdm & "'"
	sqlStr = sqlStr & "     and tencdn='" & tecdn & "'"
	sqlStr = sqlStr & " END"
	sqlStr = sqlStr & " ELSE"
	sqlStr = sqlStr & " BEGIN"
	sqlStr = sqlStr & "     INSERT INTO [db_item].[dbo].tbl_interpark_dspcategory_mapping "
	sqlStr = sqlStr & "     (tencdl, tencdm, tencdn, interparkdispcategory, SupplyCtrtSeq, interparkstorecategory) "
	sqlStr = sqlStr & "     VALUES("
	sqlStr = sqlStr & "     '" & tecdl & "'"
	sqlStr = sqlStr & "     ,'" & tecdm & "'"
	sqlStr = sqlStr & "     ,'" & tecdn & "'"
	sqlStr = sqlStr & "     ,'" & interparkdispcategory & "'"
	sqlStr = sqlStr & "     ," & SupplyCtrtSeq & ""
	sqlStr = sqlStr & "     ,'" & interparkstorecategory & "'"
	sqlStr = sqlStr & "     )"
	sqlStr = sqlStr & " END"
	dbget.Execute sqlStr
    ''전시 카테고리가 변경된 경우
	If (oldDispCate<>"") and (oldDispCate<>interparkdispcategory) Then
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE [db_item].[dbo].tbl_interpark_reg_item"
		sqlStr = sqlStr & " SET interparklastupdate='2008-01-01'"
		sqlStr = sqlStr & " WHERE itemid in ("
		sqlStr = sqlStr & "		SELECT TOP 500 r.itemid from [db_item].[dbo].tbl_interpark_reg_item r,"
		sqlStr = sqlStr & "		[db_item].[dbo].tbl_item i, [db_item].[dbo].tbl_interpark_dspcategory_mapping p"
		sqlStr = sqlStr & "		WHERE r.itemid=i.itemid"
		sqlStr = sqlStr & "		and p.interparkdispcategory='" & interparkdispcategory & "'"
		sqlStr = sqlStr & "		and p.tencdl=i.cate_large"
		sqlStr = sqlStr & "		and p.tencdm=i.cate_mid"
		sqlStr = sqlStr & "		and p.tencdn=i.cate_small"
		sqlStr = sqlStr & " )"
		dbget.Execute sqlStr
		
		'''카테고리가 변경되어도 .. 연동은 되도록..
		''''상품에 연결 - ipark상품쪽에 전시카테고리 연결. interparkSupplyCtrtSeq 는 동일해야함.. 바뀌면 상품업데이트 안됨.
		''' 2011-04-21  상품 등록시에도 필요할듯.. => 등록 성공 프로세스쪽에..
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE R"
		sqlStr = sqlStr & " SET interparkSupplyCtrtSeq=D.SupplyCtrtSeq"
		sqlStr = sqlStr & " ,interparkStoreCategory=D.interparkStoreCategory"
		sqlStr = sqlStr & " ,Pinterparkdispcategory=D.interparkdispcategory"
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_interpark_reg_item R"
		sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item i on R.itemid=i.itemid"
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_interpark_dspcategory_mapping D on D.tencdl=i.cate_large and D.tencdm=i.cate_mid and D.tencdn=i.cate_small"
		sqlStr = sqlStr & " WHERE IsNULL(R.interparkSupplyCtrtSeq,D.SupplyCtrtSeq)=D.SupplyCtrtSeq"
		sqlStr = sqlStr & " and D.SupplyCtrtSeq is Not NULL"
		sqlStr = sqlStr & " and i.cate_large='" & tecdl & "'"
		sqlStr = sqlStr & " and i.cate_mid='" & tecdm & "'"
		sqlStr = sqlStr & " and i.cate_small='" & tecdn & "'"
		sqlStr = sqlStr & " and R.interParkPrdNo is Not NULL"
		dbget.Execute sqlStr
	End If
End If
%>
<script language='javascript'>
alert('저장되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->