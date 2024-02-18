<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, cnt
mode    = Request("mode")

'// 모드별 분기
Select Case mode
    CASE "deal"
        sqlStr = ""
        sqlStr = sqlStr & " SELECT itemid, regdate "
        sqlStr = sqlStr & " INTO #TMPTBL "
        sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_wemake_regitem "
        sqlStr = sqlStr & " WHERE itemid in (1830261,1830260,1830259,1875952,1876125,1635535,2124860,2124859,1989316,1871564,1989314,2126261,2127965,2127962,1834492,1834491,2240073,2237489,1829901,2044057,1829923,1829921,1875938,1989318,1895549,1989317,1871566,1871565,2126212,2126264,2126263,2126262,1871563,1871562,1830716,1830714,1830701,1830697,1835602,1835590,1835589,2087575,2087574,1965535,2237504,2237503,2237502,2237501,2237500,2237499,1945617,1945622,1945621,1945618,2032234,2032233,2127811,2237481,2124858,2127942) "
        sqlStr = sqlStr & " and regdate <= '2019-03-10 23:59:59' "
        dbget.execute sqlStr

        sqlStr = ""
        sqlStr = sqlStr & " SELECT COUNT(*) as CNT"
        sqlStr = sqlStr & " FROM #TMPTBL "
		rsget.Open sqlStr,dbget,1
			cnt = rsget("cnt")
		rsget.Close

        If cnt > 0 Then
            sqlStr = ""
            sqlStr = sqlStr & " DELETE FROM db_etcmall.dbo.tbl_wemake_regitem WHERE itemid in (SELECT itemid FROM #TMPTBL)"
            dbget.execute sqlStr

            sqlStr = ""
            sqlStr = sqlStr & " DELETE FROM db_item.dbo.tbl_OutMall_regedoption WHERE itemid in (SELECT itemid FROM #TMPTBL) and mallid = 'WMP'"
            dbget.execute sqlStr

            sqlStr = ""
            sqlStr = sqlStr & " DELETE FROM db_etcmall.dbo.tbl_outmall_API_Que WHERE itemid in (SELECT itemid FROM #TMPTBL) and mallid = 'WMP'"
            dbget.execute sqlStr
            Response.Write "<script language=javascript>alert('반영 하였습니다.');parent.location.reload();</script>"
            dbget.close()	:	response.End
        Else
            Response.Write "<script language=javascript>alert('삭제할 DATA가 없습니다.');</script>"
            dbget.close()	:	response.End
        End If
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

