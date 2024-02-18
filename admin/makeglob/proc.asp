<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  메이크글로비 proc페이지
' History : 2015.11.11 원승현 생성
'			2016.07.04 김진영 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/makeglob/makeglobCls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->

<%
	Dim mode, hiddenvalue, soldoutvalue, arrproductcode, strsql, paramvalue, tmparrproductcode, vquery, i, UICheck
	Dim tenCateCode, tengldispChkdep2, tengldispChkdep3, maxMgCateCd, sIdx, maySoldout

	mode = request("mode")
	hiddenvalue = request("hiddenvalue")
	soldoutvalue = request("soldoutvalue")
	arrproductcode = request("arrproductcode")
	paramvalue = tenDec(request("paramvalue"))
	paramvalue = Server.URLencode(paramvalue)

	Select Case Trim(mode)

		Case "hidden"

			If Trim(hiddenvalue)="" Then
				Response.write "<script>alert('정상적인 경로로 접근해주세요.');history.back();</script>"
				Response.End
			End If

			If Trim(arrproductcode)="" Then
				Response.write "<script>alert('정상적인 경로로 접근해주세요.');history.back();</script>"
				Response.End
			End If

			strsql = " update db_item.dbo.tbl_makeglob_product set hidden='"&hiddenvalue&"', makeglobYN='N', makeupdate='1900-01-01' Where product_code in ("&arrproductcode&") "
			dbget.execute strsql

			If Trim(hiddenvalue)="Y" Then
				Response.write "<script>alert('선택한 상품이 숨김처리 되었습니다.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"';</script>"
				Response.End
			Else
				Response.write "<script>alert('선택한 상품이 노출처리 되었습니다.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"';</script>"
				Response.End
			End If



		Case "soldout"
			If Trim(soldoutvalue)="" Then
				Response.write "<script>alert('정상적인 경로로 접근해주세요.');history.back();</script>"
				Response.End
			End If

			If Trim(arrproductcode)="" Then
				Response.write "<script>alert('정상적인 경로로 접근해주세요.');history.back();</script>"
				Response.End
			End If

			strsql = " update db_item.dbo.tbl_makeglob_product set soldout='"&soldoutvalue&"', makeglobYN='N', makeupdate='1900-01-01' Where product_code in ("&arrproductcode&") "
			dbget.execute strsql

			If Trim(soldoutvalue)="Y" Then
				Response.write "<script>alert('선택한 상품이 품절처리 되었습니다.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"';</script>"
				Response.End
			Else
				Response.write "<script>alert('선택한 상품이 판매가능 상태로 변경 되었습니다.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"'</script>"
				Response.End
			End If



		Case "product"

			If Trim(arrproductcode)="" Then
				Response.write "<script>alert('정상적인 경로로 접근해주세요.');history.back();</script>"
				Response.End
			End If
		
			tmparrproductcode = Split(arrproductcode, ",")

			For i=0 To UBound(tmparrproductcode)

				strsql = " Select top 1 mp.product_key, mo.idx From db_item.dbo.tbl_makeglob_product mp "
				strsql = strsql & " left join db_item.dbo.tbl_makeglob_product_option mo on mp.product_key = mo.product_key And mp.product_code = mo.product_code "
				strsql = strsql & " where mp.product_code='"&tmparrproductcode(i)&"' "
		        rsget.Open strsql,dbget, 1
				If Not(rsget.bof Or rsget.eof) Then
					UICheck = "U"
					sIdx = Trim(rsget("idx"))
				Else
					UICheck = "I"
				End If
				rsget.close
'******************************************************************************************************************************************************************
'********************************************************* '2016-07-04 김진영..maySoldOut 추가  *******************************************************************
				strsql = ""
				strsql = strsql & " SELECT TOP 1 "
				strsql = strsql & " (CASE WHEN i.isusing = 'N' "					'사용여부 N
				strsql = strsql & " 	OR i.deliverOverseas <> 'Y' "				'해외판매여부 N
				strsql = strsql & " 	OR i.itemweight < 0 "						'무게 0미만
				strsql = strsql & " 	OR i.sellyn <> 'Y' "						'품절 or 일시품절
				strsql = strsql & " 	OR i.itemid in (Select TOP 1 itemid FROM db_item.[dbo].[tbl_const_OptAddPrice_Exists] WHERE itemid = '"&tmparrproductcode(i)&"') "				'옵션추가금액 있음
				strsql = strsql & " 	OR r.jungsan_gubun = '간이과세' "				'브랜드 정산구분이 간이과세면 N
				strsql = strsql & " 	THEN 'Y' ELSE 'N' END) as maySoldOut "		'위 조건에 해당하면 maySoldOut에 Y, 아니면 N
				strsql = strsql & " FROM db_item.dbo.tbl_item i "
				strsql = strsql & " join db_partner.dbo.tbl_partner as r on i.makerid = r.id "
				strsql = strsql & " WHERE i.itemid = '"&tmparrproductcode(i)&"' "
				rsget.Open strsql,dbget, 1
				If Not(rsget.bof Or rsget.eof) Then
					maySoldOut = rsget("maySoldOut")
				End If
				rsget.close
'******************************************************************************************************************************************************************
				If UICheck="U" Then
					If maySoldOut = "Y" Then
						strsql = " update db_item.dbo.tbl_makeglob_product set soldout = 'Y', makeglobYN='N', makeupdate='1900-01-01' Where product_code in ("&tmparrproductcode(i)&") "
						dbget.execute strsql
						Response.write "<script>alert('상품코드 : "&tmparrproductcode(i)&" 판매불가..\n품절처리 합니다. ');history.back();</script>"
						Response.End
					Else
						'// 전시 카테고리가 현재 메이크글로비 DB에도 있는지 확인하여 없으면 넣어준다.
						'// 해당상품 전시 카테고리를 가져온다.
						strsql = " Select top 1 catecode From db_item.dbo.tbl_display_cate_item Where itemid='"&tmparrproductcode(i)&"' and isDefault='y' "
						rsget.Open strsql,dbget, 1
						If Not(rsget.bof Or rsget.eof) Then
							tenCateCode = Trim(rsget("catecode"))
						End If
						rsget.close

	
						'// 9자리로 잘라 현재 tbl_makeglob_cate_matching 테이블에 값이 있는지 확인한다.
						strsql = " Select dispCate From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 9)='"&tenCateCode&"' "
						rsget.Open strsql,dbget, 1
						If Not(rsget.bof Or rsget.eof) Then
							tengldispChkdep3 = "Y"
						Else
							tengldispChkdep3 = "N"
						End If
						rsget.close
	
						'// 9자리로 되어 있는 값이 없으면 6자리를 체크하여 2뎁스 내역이 있는지 확인한다.
						If tengldispChkdep3="N" Then
							strsql = " Select top 1 * From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 6)='"&Left(tenCateCode, 6)&"' "
							rsget.Open strsql,dbget, 1
							If Not(rsget.bof Or rsget.eof) Then
								tengldispChkdep2 = "Y"
							Else
								tengldispChkdep2 = "N"
							End If
							rsget.close
	
							'// 2뎁스가 있으면 현재 들어온 카테고리값을 넣어준다.
							If tengldispChkdep2 = "Y" Then
								'// 6자리 카테고리 코드 master값을 가져온다.
								strsql = " Select top 1 max(Mg_cateCd) as Mg_cateCd From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 6)='"&Left(tenCateCode, 6)&"' "
								rsget.Open strsql,dbget, 1
								If Not(rsget.bof Or rsget.eof) Then
									'// max값+1로 Mg_cateCd를 생성하여 insert한다.
									vquery = " insert into db_item.[dbo].[tbl_makeglob_Cate_matching] "
									vquery = vquery & " values ('"&CInt(rsget("Mg_cateCd"))+1&"', '"&Left(tenCateCode, 9)&"','') "
									dbget.execute vquery
								Else
									'// 2뎁스가 없으면 개발자가 직접 넣어줘야 하므로 alert 띄운다.
									Response.write "<script>alert('등록된 카테고리값이 없습니다.\n개발팀에 문의해주세요.');history.back();</script>"
									Response.End					
								End If
								rsget.close
							Else
								'// 2뎁스가 없으면 개발자가 직접 넣어줘야 하므로 alert 띄운다.
								Response.write "<script>alert('등록된 카테고리값이 없습니다.\n개발팀에 문의해주세요.');history.back();</script>"
								Response.End					
							End If
						End If
	
						'// 있는 상품이므로 update
						'// 상품 테이블 update
						vquery = " update db_item.dbo.tbl_makeglob_product "
						vquery = vquery & " set product_name=T1.itemname, list_img_url = T1.listimage120, detail_img_url = T1.icon1image, zoom_img_url = T1.basicimage, "
						vquery = vquery & " 	basic600_img_url = T1.basicimage600, basic1000_img_url = T1.basicimage1000, weight = T1.itemweight, maker_name = T1.makername, "
						vquery = vquery & " 	madein = T1.sourcearea, brand_name = T1.brandname, manufacture_date = T1.manufacture_date, launching_date = T1.sellSTDate, "
						vquery = vquery & " 	keyword = T1.keywords, [desc] = T1.itemcontent, itemsource = T1.itemsource, itemsize = T1.itemsize,cateindex = T1.cateindex, "
						vquery = vquery & " 	makeglobYN='N', lastupdate=getdate(), makeupdate='1900-01-01' "
						vquery = vquery & " From "
						vquery = vquery & " ( "
						vquery = vquery & " 	Select top 1 "
						vquery = vquery & " 		i.itemid, 'KO' as product_language, 'KRW' as currency, i.itemname,i.orgprice as product_price, "
						vquery = vquery & " 		i.orgprice as original_price, 0 as supply_price, i.listimage120, i.icon1image,i.basicimage,i.basicimage600, i.basicimage1000, 0 as mileage, "
						vquery = vquery & " 		convert(float, i.itemweight)/1000 as itemweight, ic.makername,ic.sourcearea,i.brandname,i.sellStdate as manufacture_date, i.sellStdate,  "
						vquery = vquery & " 		ic.keywords, "
'						vquery = vquery & "			ic.itemcontent, "
						vquery = vquery & "			Case WHEN i.mwdiv not in ('M', 'W') Then CONVERT(VARCHAR(100),'<p><img src=""http://webimage.10x10.co.kr/common/uploadimg/2017/china/1716265/1705301355410340028.jpg""></p>') + '<br/>' + CONVERT(VARCHAR (max),ic.itemcontent) Else ic.itemcontent end as itemcontent, "
						vquery = vquery & "			ic.itemsource, ic.itemsize, "
						vquery = vquery & " 			'N' as hidden, 'N' as soldout, '' as product_url, '' as pdt_stock, "
						vquery = vquery & " 		db_item.dbo.getMakeglobCateCd(ci.catecode,Case When isNull(i.frontMakerid,'') <>'' then i.frontMakerid else i.makerid end) as cateindex, "
						vquery = vquery & " 		'N' as makeglobYN, getdate() as regdate, getdate() as lastupdate, '' as makeupdate "
						vquery = vquery & " 	From db_item.dbo.tbl_item i  "
						vquery = vquery & " 	inner join db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid  "
						vquery = vquery & " 	inner join db_item.dbo.tbl_display_cate_item ci on i.itemid = ci.itemid And ci.isDefault='y'  "
						vquery = vquery & " 	inner join db_user.dbo.tbl_user_c c on i.makerid = c.userid  "
						vquery = vquery & " 	Where  i.deliverOverseas='Y' And i.itemweight>0  "
						vquery = vquery & " 		And i.isusing='Y' And i.sellyn='Y' And i.itemid='"&tmparrproductcode(i)&"' "
						vquery = vquery & " ) as T1	Where product_code='"&tmparrproductcode(i)&"' "
						dbget.execute vquery
	
	
						'// 옵션이 있다면 옵션도 업데이트 해준다.
						If isnull(sIdx) Or sIdx="" Then
	
						Else
							vquery = " update C "
							vquery = vquery & " 	set C.option_index_name = case when A.optionTypeName='' then '옵션명' when A.optionTypeName is null then '옵션명' else A.optionTypename end, "
							vquery = vquery & " 	C.option_index_value = A.optionname, C.option_index_price = A.optaddprice, "
							vquery = vquery & " 	C.stock = case when A.optlimityn='Y' then optlimitno-optlimitsold else 0 end, "
							vquery = vquery & " 	C.soldout = case when A.optlimityn = 'Y' then case when optlimitno-optlimitsold=0 then 'Y' else 'N' end else 'N' end, "
							vquery = vquery & " 	C.hidden = case when A.optsellyn='Y' then 'N' else 'Y' end, "
							vquery = vquery & " 	C.lastupdate = getdate() "
							vquery = vquery & " From db_item.dbo.tbl_item_option A "
							vquery = vquery & " inner join db_item.dbo.tbl_makeglob_product B on A.itemid = B.product_code "
							vquery = vquery & " inner join db_item.dbo.tbl_makeglob_product_option C on a.itemid = C.product_code And B.product_key = C.product_key And A.itemoption = C.tenoptioncode "
							vquery = vquery & " Where A.isusing='Y' And A.itemid='"&tmparrproductcode(i)&"' "
							dbget.execute vquery
						End If
					End If
				ElseIf UICheck="I" Then
					If maySoldOut = "N" Then
						'// 전시 카테고리가 현재 메이크글로비 DB에도 있는지 확인하여 없으면 넣어준다.
						'// 해당상품 전시 카테고리를 가져온다.
						strsql = " Select top 1 catecode From db_item.dbo.tbl_display_cate_item Where itemid='"&tmparrproductcode(i)&"' and isDefault='y' "
						rsget.Open strsql,dbget, 1
						If Not(rsget.bof Or rsget.eof) Then
							tenCateCode = Trim(Left(rsget("catecode"), 9))
						End If
						rsget.close
	
						'// 9자리로 잘라 현재 tbl_makeglob_cate_matching 테이블에 값이 있는지 확인한다.
						strsql = " Select dispCate From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 9)='"&tenCateCode&"' "
						rsget.Open strsql,dbget, 1
						If Not(rsget.bof Or rsget.eof) Then
							tengldispChkdep3 = "Y"
						Else
							tengldispChkdep3 = "N"
						End If
						rsget.close


						'// 9자리로 되어 있는 값이 없으면 6자리를 체크하여 2뎁스 내역이 있는지 확인한다.
						If tengldispChkdep3="N" Then
							strsql = " Select top 1 * From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 6)='"&Left(tenCateCode, 6)&"' "
							rsget.Open strsql,dbget, 1
							If Not(rsget.bof Or rsget.eof) Then
								tengldispChkdep2 = "Y"
							Else
								tengldispChkdep2 = "N"
							End If
							rsget.close
						End If

						'// 2뎁스가 있으면 현재 들어온 카테고리값을 넣어준다.
						If tengldispChkdep2 = "Y" Then
							'// 6자리 카테고리 코드 master값을 가져온다.
							strsql = " Select top 1 Mg_cateCd From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 6)='"&Left(tenCateCode, 6)&"' "
							rsget.Open strsql,dbget, 1
							If Not(rsget.bof Or rsget.eof) Then
								'// max값+1로 Mg_cateCd를 생성하여 insert한다.
								vquery = " insert into db_item.[dbo].[tbl_makeglob_Cate_matching] "
								vquery = vquery & " values ('"&rsget("Mg_cateCd")&"', '"&tenCateCode&"','') "
								dbget.execute vquery
							Else
								'// 2뎁스가 없으면 개발자가 직접 넣어줘야 하므로 alert 띄운다.
								Response.write "<script>alert('등록된 카테고리값이 없습니다.\n개발팀에 문의해주세요.');history.back();</script>"
								Response.End					
							End If
							rsget.close
						ElseIf tengldispChkdep2="N" Then
							'// 2뎁스가 없으면 개발자가 직접 넣어줘야 하므로 alert 띄운다.
							Response.write "<script>alert('등록된 카테고리값이 없습니다.\n개발팀에 문의해주세요.');history.back();</script>"
							Response.End					
	
						End If
	
	
						'// 없는 상품이므로 insert
						'// 상품 테이블에 insert
						vquery = "insert into db_item.[dbo].[tbl_makeglob_product] "
						vquery = vquery & " Select top 1 "
						vquery = vquery & "		i.itemid, 'KO' as product_language, 'KRW' as currency, i.itemname,i.orgprice as product_price, "
						vquery = vquery & "		i.orgprice as original_price, 0 as supply_price, i.listimage120, i.icon1image,i.basicimage,i.basicimage600, i.basicimage1000, 0 as mileage, "
						vquery = vquery & "		convert(float, i.itemweight)/1000 as itemweight, ic.makername,ic.sourcearea,i.brandname,i.sellStdate as manufacture_date, i.sellStdate,  "
						vquery = vquery & "		ic.keywords, "
'						vquery = vquery & "		ic.itemcontent, "
						vquery = vquery & "		Case WHEN i.mwdiv not in ('M', 'W') Then CONVERT(VARCHAR(100),'<p><img src=""http://webimage.10x10.co.kr/common/uploadimg/2017/china/1716265/1705301355410340028.jpg""></p>') + '<br/>' + CONVERT(VARCHAR (max),ic.itemcontent) Else ic.itemcontent end as itemcontent, "
						vquery = vquery & "		ic.itemsource, ic.itemsize, "
						vquery = vquery & "			'N' as hidden, 'N' as soldout, '' as product_url, '' as pdt_stock, "
						vquery = vquery & "		db_item.dbo.getMakeglobCateCd(ci.catecode,Case When isNull(i.frontMakerid,'') <>'' then i.frontMakerid else i.makerid end) as cateindex, "
						vquery = vquery & "		'N', getdate(), getdate(), '' as makeupdate "
						vquery = vquery & "	From db_item.dbo.tbl_item i  "
						vquery = vquery & "	inner join db_item.dbo.tbl_item_contents ic on i.itemid = ic.itemid  "
						vquery = vquery & "	inner join db_item.dbo.tbl_display_cate_item ci on i.itemid = ci.itemid And ci.isDefault='y'  "
						vquery = vquery & "	inner join db_user.dbo.tbl_user_c c on i.makerid = c.userid  "
						vquery = vquery & "	Where  i.deliverOverseas='Y' And i.itemweight>0  "
						vquery = vquery & "		And i.isusing='Y' And i.sellyn='Y' And i.itemid='"&tmparrproductcode(i)&"' "
						dbget.execute vquery
	
						'// 상품테이블에 넣은 후 옵션도 넣어준다.
						vquery = " insert into db_item.dbo.tbl_makeglob_product_option "
						vquery = vquery & "	Select  b.product_key, b.product_code, a.itemoption, "
						vquery = vquery & "		case when A.optionTypeName='' then '옵션명' when A.optionTypeName is null then '옵션명'	else A.optionTypeName end as option_index_name, "
						vquery = vquery & "		A.optionname, A.optaddprice, case when A.optlimityn='Y' then optlimitno-optlimitsold else 0 end as stock, "
						vquery = vquery & "		case when A.optlimityn='Y' then  "
						vquery = vquery & "			case when optlimitno-optlimitsold=0 then 'Y' else 'N' end "
						vquery = vquery & "		else 'N' end as soldout,  "
						vquery = vquery & "		case when A.optsellyn='Y' then 'N' else 'Y' end as hidden, getdate(), getdate() "
						vquery = vquery & "	From db_item.dbo.tbl_item_option A "
						vquery = vquery & "	inner join db_item.dbo.tbl_makeglob_product B on A.itemid = B.product_code "
						vquery = vquery & "	Where A.isusing='Y' And A.itemid='"&tmparrproductcode(i)&"' "
						dbget.execute vquery
					Else
						Response.write "<script>alert('상품코드 : "&tmparrproductcode(i)&" | 판매불가 상품입니다.');history.back();</script>"
						Response.End
					End If
				Else
					Response.write "<script>alert('오류가 발생하였습니다.개발팀으로 문의주세요.');history.back();</script>"
					Response.End
				End If
			Next

			Response.write "<script>alert('선택하신 상품이 등록/수정 되었습니다.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"';</script>"
			Response.End					


		Case Else
			Response.write "<script>alert('정상적인 경로로 접근해주세요.');history.back();</script>"
			Response.End

	End Select


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->