<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ũ�۷κ� proc������
' History : 2015.11.11 ������ ����
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
	Dim tenCateCode, tengldispChkdep2, tengldispChkdep3, maxMgCateCd, sIdx

	mode = request("mode")
	hiddenvalue = request("hiddenvalue")
	soldoutvalue = request("soldoutvalue")
	arrproductcode = request("arrproductcode")
	paramvalue = tenDec(request("paramvalue"))
	paramvalue = Server.URLencode(paramvalue)

	Select Case Trim(mode)

		Case "hidden"

			If Trim(hiddenvalue)="" Then
				Response.write "<script>alert('�������� ��η� �������ּ���.');history.back();</script>"
				Response.End
			End If

			If Trim(arrproductcode)="" Then
				Response.write "<script>alert('�������� ��η� �������ּ���.');history.back();</script>"
				Response.End
			End If

			strsql = " update db_item.dbo.tbl_makeglob_product set hidden='"&hiddenvalue&"', makeglobYN='N', makeupdate='1900-01-01' Where product_code in ("&arrproductcode&") "
			dbget.execute strsql

			If Trim(hiddenvalue)="Y" Then
				Response.write "<script>alert('������ ��ǰ�� ����ó�� �Ǿ����ϴ�.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"';</script>"
				Response.End
			Else
				Response.write "<script>alert('������ ��ǰ�� ����ó�� �Ǿ����ϴ�.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"';</script>"
				Response.End
			End If



		Case "soldout"
			If Trim(soldoutvalue)="" Then
				Response.write "<script>alert('�������� ��η� �������ּ���.');history.back();</script>"
				Response.End
			End If

			If Trim(arrproductcode)="" Then
				Response.write "<script>alert('�������� ��η� �������ּ���.');history.back();</script>"
				Response.End
			End If

			strsql = " update db_item.dbo.tbl_makeglob_product set soldout='"&soldoutvalue&"', makeglobYN='N', makeupdate='1900-01-01' Where product_code in ("&arrproductcode&") "
			dbget.execute strsql

			If Trim(soldoutvalue)="Y" Then
				Response.write "<script>alert('������ ��ǰ�� ǰ��ó�� �Ǿ����ϴ�.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"';</script>"
				Response.End
			Else
				Response.write "<script>alert('������ ��ǰ�� �ǸŰ��� ���·� ���� �Ǿ����ϴ�.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"'</script>"
				Response.End
			End If



		Case "product"

			If Trim(arrproductcode)="" Then
				Response.write "<script>alert('�������� ��η� �������ּ���.');history.back();</script>"
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

				If UICheck="U" Then

					'// ���� ī�װ��� ���� ����ũ�۷κ� DB���� �ִ��� Ȯ���Ͽ� ������ �־��ش�.
					'// �ش��ǰ ���� ī�װ��� �����´�.
					strsql = " Select top 1 catecode From db_item.dbo.tbl_display_cate_item Where itemid='"&tmparrproductcode(i)&"' and isDefault='y' "
					rsget.Open strsql,dbget, 1
					If Not(rsget.bof Or rsget.eof) Then
						tenCateCode = Trim(rsget("catecode"))
					End If
					rsget.close

					'// 9�ڸ��� �߶� ���� tbl_makeglob_cate_matching ���̺� ���� �ִ��� Ȯ���Ѵ�.
					strsql = " Select dispCate From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 9)='"&tenCateCode&"' "
					rsget.Open strsql,dbget, 1
					If Not(rsget.bof Or rsget.eof) Then
						tengldispChkdep3 = "Y"
					Else
						tengldispChkdep3 = "N"
					End If
					rsget.close

					'// 9�ڸ��� �Ǿ� �ִ� ���� ������ 6�ڸ��� üũ�Ͽ� 2���� ������ �ִ��� Ȯ���Ѵ�.
					If tengldispChkdep3="N" Then
						strsql = " Select top 1 * From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 6)='"&Left(tenCateCode, 6)&"' "
						rsget.Open strsql,dbget, 1
						If Not(rsget.bof Or rsget.eof) Then
							tengldispChkdep2 = "Y"
						Else
							tengldispChkdep2 = "N"
						End If
						rsget.close

						'// 2������ ������ ���� ���� ī�װ����� �־��ش�.
						If tengldispChkdep2 = "Y" Then
							'// 6�ڸ� ī�װ� �ڵ� master���� �����´�.
							strsql = " Select top 1 Mg_cateCd From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 6)='"&Left(tenCateCode, 6)&"' "
							rsget.Open strsql,dbget, 1
							If Not(rsget.bof Or rsget.eof) Then
								'// max��+1�� Mg_cateCd�� �����Ͽ� insert�Ѵ�.
								vquery = " insert into db_item.[dbo].[tbl_makeglob_Cate_matching] "
								vquery = vquery & " values ('"&rsget("Mg_cateCd")&"', '"&tenCateCode&"','') "
								dbget.execute vquery
							Else
								'// 2������ ������ �����ڰ� ���� �־���� �ϹǷ� alert ����.
								Response.write "<script>alert('��ϵ� ī�װ����� �����ϴ�.\n�������� �������ּ���.');history.back();</script>"
								Response.End					
							End If
							rsget.close
						Else
							'// 2������ ������ �����ڰ� ���� �־���� �ϹǷ� alert ����.
							Response.write "<script>alert('��ϵ� ī�װ����� �����ϴ�.\n�������� �������ּ���.');history.back();</script>"
							Response.End					
						End If
					End If

					'// �ִ� ��ǰ�̹Ƿ� update
					'// ��ǰ ���̺� update
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
					vquery = vquery & " 		ic.keywords, ic.itemcontent, ic.itemsource, ic.itemsize, "
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


					'// �ɼ��� �ִٸ� �ɼǵ� ������Ʈ ���ش�.
					If isnull(sIdx) Or sIdx="" Then

					Else
						vquery = " update C "
						vquery = vquery & " 	set C.option_index_name = case when A.optionTypeName='' then '�ɼǸ�' when A.optionTypeName is null then '�ɼǸ�' else A.optionTypename end, "
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
				ElseIf UICheck="I" Then

					'// ���� ī�װ��� ���� ����ũ�۷κ� DB���� �ִ��� Ȯ���Ͽ� ������ �־��ش�.
					'// �ش��ǰ ���� ī�װ��� �����´�.
					strsql = " Select top 1 catecode From db_item.dbo.tbl_display_cate_item Where itemid='"&tmparrproductcode(i)&"' and isDefault='y' "
					rsget.Open strsql,dbget, 1
					If Not(rsget.bof Or rsget.eof) Then
						tenCateCode = Trim(Left(rsget("catecode"), 9))
					End If
					rsget.close

					'// 9�ڸ��� �߶� ���� tbl_makeglob_cate_matching ���̺� ���� �ִ��� Ȯ���Ѵ�.
					strsql = " Select dispCate From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 9)='"&tenCateCode&"' "
					rsget.Open strsql,dbget, 1
					If Not(rsget.bof Or rsget.eof) Then
						tengldispChkdep3 = "Y"
					Else
						tengldispChkdep3 = "N"
					End If
					rsget.close

					'// 9�ڸ��� �Ǿ� �ִ� ���� ������ 6�ڸ��� üũ�Ͽ� 2���� ������ �ִ��� Ȯ���Ѵ�.
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

					'// 2������ ������ ���� ���� ī�װ����� �־��ش�.
					If tengldispChkdep2 = "Y" Then
						'// 6�ڸ� ī�װ� �ڵ� master���� �����´�.
						strsql = " Select top 1 Mg_cateCd From db_item.[dbo].[tbl_makeglob_Cate_matching] Where left(dispcate, 6)='"&Left(tenCateCode, 6)&"' "
						rsget.Open strsql,dbget, 1
						If Not(rsget.bof Or rsget.eof) Then
							'// max��+1�� Mg_cateCd�� �����Ͽ� insert�Ѵ�.
							vquery = " insert into db_item.[dbo].[tbl_makeglob_Cate_matching] "
							vquery = vquery & " values ('"&rsget("Mg_cateCd")&"', '"&tenCateCode&"','') "
							dbget.execute vquery
						Else
							'// 2������ ������ �����ڰ� ���� �־���� �ϹǷ� alert ����.
							Response.write "<script>alert('��ϵ� ī�װ����� �����ϴ�.\n�������� �������ּ���.');history.back();</script>"
							Response.End					
						End If
						rsget.close
					ElseIf tengldispChkdep2="N" Then
						'// 2������ ������ �����ڰ� ���� �־���� �ϹǷ� alert ����.
						Response.write "<script>alert('��ϵ� ī�װ����� �����ϴ�.\n�������� �������ּ���.');history.back();</script>"
						Response.End					

					End If


					'// ���� ��ǰ�̹Ƿ� insert
					'// ��ǰ ���̺� insert
					vquery = "insert into db_item.[dbo].[tbl_makeglob_product] "
					vquery = vquery & " Select top 1 "
					vquery = vquery & "		i.itemid, 'KO' as product_language, 'KRW' as currency, i.itemname,i.orgprice as product_price, "
					vquery = vquery & "		i.orgprice as original_price, 0 as supply_price, i.listimage120, i.icon1image,i.basicimage,i.basicimage600, i.basicimage1000, 0 as mileage, "
					vquery = vquery & "		convert(float, i.itemweight)/1000 as itemweight, ic.makername,ic.sourcearea,i.brandname,i.sellStdate as manufacture_date, i.sellStdate,  "
					vquery = vquery & "		ic.keywords, ic.itemcontent, ic.itemsource, ic.itemsize, "
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

					'// ��ǰ���̺� ���� �� �ɼǵ� �־��ش�.
					vquery = " insert into db_item.dbo.tbl_makeglob_product_option "
					vquery = vquery & "	Select  b.product_key, b.product_code, a.itemoption, "
					vquery = vquery & "		case when A.optionTypeName='' then '�ɼǸ�' when A.optionTypeName is null then '�ɼǸ�'	else A.optionTypeName end as option_index_name, "
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
					Response.write "<script>alert('������ �߻��Ͽ����ϴ�.���������� �����ּ���.');history.back();</script>"
					Response.End
				End If
			Next

			Response.write "<script>alert('�����Ͻ� ��ǰ�� ���/���� �Ǿ����ϴ�.');location.href='/admin/makeglob/itemwaitlist.asp?"&paramvalue&"';</script>"
			Response.End					


		Case Else
			Response.write "<script>alert('�������� ��η� �������ּ���.');history.back();</script>"
			Response.End

	End Select


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->