<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
Dim makerid, sqlStr
Dim deliveryCd, deliveryAddrCd, brandcd
makerid			= request("makerid")
deliveryCd		= request("deliveryCd")
deliveryAddrCd	= request("deliveryAddrCd")
brandcd			= request("brandcd")

sqlStr = ""
sqlStr = sqlStr & " IF Exists(SELECT * FROM db_item.dbo.tbl_gsshop_brandDelivery_mapping where makerid='"&makerid&"')"
sqlStr = sqlStr & " BEGIN"& VbCRLF
sqlStr = sqlStr & " UPDATE R" & VbCRLF
sqlStr = sqlStr & "	SET deliveryCd = '"& deliveryCd &"' "  & VbCRLF
sqlStr = sqlStr & "	,deliveryAddrCd = '"& deliveryAddrCd &"' "& VbCRLF
sqlStr = sqlStr & "	,brandcd = '"& brandcd &"' "& VbCRLF
sqlStr = sqlStr & "	,lastupdate = getdate() "& VbCRLF
sqlStr = sqlStr & "	,updateid = '"& session("ssBctID") &"' "& VbCRLF
sqlStr = sqlStr & "	FROM db_item.dbo.tbl_gsshop_brandDelivery_mapping R"& VbCRLF
sqlStr = sqlStr & " WHERE R.makerid='" & makerid & "'"
sqlStr = sqlStr & " END ELSE "
sqlStr = sqlStr & " BEGIN"& VbCRLF
sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_gsshop_brandDelivery_mapping "
sqlStr = sqlStr & " (makerid, deliveryCd, deliveryAddrCd, brandcd, regdate, regid)"
sqlStr = sqlStr & " VALUES ('"& makerid &"', '"& deliveryCd &"', '"& deliveryAddrCd &"', '"& brandcd &"', getdate(), '"& session("ssBctID") &"')"
sqlStr = sqlStr & " END "
dbget.Execute sqlStr
%>
<script language="javascript">
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->