<%




Class CDeliveryPolicyItem
        public Fuserid
        public Fsocname_kor
        public Fconame
        public FdefaultdeliveryType
        public Fprice0
        public Fprice10000
        public Fprice20000
        public Fprice30000
        public Fprice40000
        public Fprice50000
        public Fupchecount
        public Fwitakcount
        public Fmaeipcount
        public Fitemcount
        public FdefaultFreeBeasongLimit
        public FdefaultDeliverPay

        Private Sub Class_Initialize()

        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class

Class CDeliveryPolicy
        public FItemList()
        public FOneItem

        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

        public FRectUserID
        public FRectMDUserID
        public FRectCategoryCode
        public FRectDefaultDeliveryType
        public FRectIsUsingBrand
        public FRectIsUsingItem

		public FRectMWDiv


        public Sub GetList()
            dim i,sqlStr, itemSql, brandSql
            dim startRowNum, endRowNum

            brandSql = ""

            if (FRectUserID <> "") then
                brandSql = brandSql + " and b.userid = '" + CStr(FRectUserID) + "' "
            end if

            if (FRectMDUserID <> "") then
                brandSql = brandSql + " and b.mduserid = '" + CStr(FRectMDUserID) + "' "
            end if

            if (FRectCategoryCode <> "") then
                brandSql = brandSql + " and b.catecode = '" + CStr(FRectCategoryCode) + "' "
            end if

            if (FRectDefaultDeliveryType <> "") then
                if (FRectDefaultDeliveryType = "NULL") then
                    brandSql = brandSql + " and b.defaultdeliveryType is NULL "
                else
                    brandSql = brandSql + " and b.defaultdeliveryType = '" + CStr(FRectDefaultDeliveryType) + "' "
                end if
            end if

            if (FRectIsUsingBrand <> "") then
                brandSql = brandSql + " and b.isusing = '" + CStr(FRectIsUsingBrand) + "' "
            end if

            sqlStr = " select count(*) as cnt "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " 	db_user.dbo.tbl_user_c b "
            sqlStr = sqlStr + " 	join [db_partner].[dbo].[tbl_partner] p on b.userid = p.id and p.tplcompanyid is NULL "
            sqlStr = sqlStr + " where "
            sqlStr = sqlStr + " 	1 = 1 "

            sqlStr = sqlStr + brandSql

			rsget.Open sqlStr, dbget, 1
            ''Response.write sqlStr

			if  not rsget.EOF  then
				FTotalCount = rsget("cnt")
			else
				FTotalCount = 0
			end if
			rsget.close


            startRowNum = FPageSize * (FCurrPage - 1)
            sqlStr = " select b.userid, b.socname_kor, b.coname, b.defaultFreeBeasongLimit, b.defaultDeliverPay "
            sqlStr = sqlStr + " 		,( "
            sqlStr = sqlStr + " 			case "
            sqlStr = sqlStr + " 				when b.defaultdeliveryType = '9' then '업체조건배송' "
            sqlStr = sqlStr + " 				when b.defaultdeliveryType = '7' then '업체착불배송' "
            sqlStr = sqlStr + " 				when b.defaultdeliveryType is NULL then '업체무료배송' "
            sqlStr = sqlStr + " 				else b.defaultdeliveryType "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as defaultdeliveryType "
            sqlStr = sqlStr + " 		,sum( "
            sqlStr = sqlStr + " 			case "
            sqlStr = sqlStr + " 				when i.orgprice < 10000 then 1 "
            sqlStr = sqlStr + " 				else 0 "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as price0 "
            sqlStr = sqlStr + " 		,sum( "
            sqlStr = sqlStr + " 			case  "
            sqlStr = sqlStr + " 				when i.orgprice >= 10000 and i.orgprice < 20000 then 1 "
            sqlStr = sqlStr + " 				else 0 "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as price10000 "
            sqlStr = sqlStr + " 		,sum( "
            sqlStr = sqlStr + " 			case "
            sqlStr = sqlStr + " 				when i.orgprice >= 20000 and i.orgprice < 30000 then 1 "
            sqlStr = sqlStr + " 				else 0 "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as price20000 "
            sqlStr = sqlStr + " 		,sum( "
            sqlStr = sqlStr + " 			case "
            sqlStr = sqlStr + " 				when i.orgprice >= 30000 and i.orgprice < 40000 then 1 "
            sqlStr = sqlStr + " 				else 0 "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as price30000 "
            sqlStr = sqlStr + " 		,sum( "
            sqlStr = sqlStr + " 			case "
            sqlStr = sqlStr + " 				when i.orgprice >= 40000 and i.orgprice < 50000 then 1 "
            sqlStr = sqlStr + " 				else 0 "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as price40000 "
            sqlStr = sqlStr + " 		,sum( "
            sqlStr = sqlStr + " 			case "
            sqlStr = sqlStr + " 				when i.orgprice >= 50000 then 1 "
            sqlStr = sqlStr + " 				else 0 "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as price50000 "
            sqlStr = sqlStr + " 		,sum( "
            sqlStr = sqlStr + " 			case "
            sqlStr = sqlStr + " 				when i.mwdiv = 'U' then 1 "
            sqlStr = sqlStr + " 				else 0 "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as upchecount "
            sqlStr = sqlStr + " 		,sum( "
            sqlStr = sqlStr + " 			case "
            sqlStr = sqlStr + " 				when i.mwdiv = 'W' then 1 "
            sqlStr = sqlStr + " 				else 0 "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as witakcount "
            sqlStr = sqlStr + " 		,sum( "
            sqlStr = sqlStr + " 			case "
            sqlStr = sqlStr + " 				when i.mwdiv = 'M' then 1 "
            sqlStr = sqlStr + " 				else 0 "
            sqlStr = sqlStr + " 			end "
            sqlStr = sqlStr + " 		) as maeipcount "
            sqlStr = sqlStr + " 		,sum(1) as itemcount "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " 	( "
            sqlStr = sqlStr + " 		SELECT b.userid, b.socname_kor, b.coname, b.defaultdeliveryType, b.defaultFreeBeasongLimit, b.defaultDeliverPay "
            sqlStr = sqlStr + " 		FROM db_user.dbo.tbl_user_c b "
            sqlStr = sqlStr + " 		join [db_partner].[dbo].[tbl_partner] p on b.userid = p.id and p.tplcompanyid is NULL "
            sqlStr = sqlStr + " 		where "
            sqlStr = sqlStr + " 			1 = 1 "
            sqlStr = sqlStr + brandSql
            sqlStr = sqlStr + " 		ORDER BY userid  "
        	sqlStr = sqlStr + " 		OFFSET " & startRowNum & " ROWS FETCH NEXT " & FPageSize & " ROWS ONLY "
            sqlStr = sqlStr + " 	) B "
            sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item i on B.userid = i.makerid "
            sqlStr = sqlStr + " WHERE "
            sqlStr = sqlStr + " 	1 = 1 "

            if (FRectIsUsingItem <> "") then
                sqlStr = sqlStr + " and i.isusing = '" + CStr(FRectIsUsingItem) + "' "
            end if

            if (FRectMWDiv <> "") then
                sqlStr = sqlStr + " and i.mwdiv = '" + CStr(FRectMWDiv) + "' "
            end if

            sqlStr = sqlStr + " group by b.userid, b.socname_kor, b.coname, defaultdeliveryType, b.defaultFreeBeasongLimit, b.defaultDeliverPay "
            sqlStr = sqlStr + " order by b.userid "

            rsget.Open sqlStr, dbget, 1
            ''Response.write sqlStr

			FTotalPage = (FTotalCount\FPageSize)
			if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage + 1

            FResultCount = rsget.RecordCount

            redim preserve FItemList(FResultCount)
            if  not rsget.EOF  then
                    i = 0
                    do until rsget.eof

                            set FItemList(i) = new CDeliveryPolicyItem

                            FItemList(i).Fuserid = rsget("userid")
                            FItemList(i).Fsocname_kor = rsget("socname_kor")
                            FItemList(i).Fconame = rsget("coname")
                            FItemList(i).FdefaultdeliveryType = rsget("defaultdeliveryType")
                            FItemList(i).Fprice0 = rsget("price0")
                            FItemList(i).Fprice10000 = rsget("price10000")
                            FItemList(i).Fprice20000 = rsget("price20000")
                            FItemList(i).Fprice30000 = rsget("price30000")
                            FItemList(i).Fprice40000 = rsget("price40000")
                            FItemList(i).Fprice50000 = rsget("price50000")
                            FItemList(i).Fupchecount = rsget("upchecount")
                            FItemList(i).Fwitakcount = rsget("witakcount")
                            FItemList(i).Fmaeipcount = rsget("maeipcount")
                            FItemList(i).Fitemcount = rsget("itemcount")
                            FItemList(i).FdefaultFreeBeasongLimit = rsget("defaultFreeBeasongLimit")
                            FItemList(i).FdefaultDeliverPay = rsget("defaultDeliverPay")

                            rsget.MoveNext

                            i = i + 1
                    loop
            end if
            rsget.close
        end sub

        Private Sub Class_Initialize()
                FCurrPage       = 1
                FPageSize       = 20
                FResultCount    = 0
                FScrollCount    = 10
                FTotalCount     = 0
        End Sub

        Private Sub Class_Terminate()

        End Sub

        public Function HasPreScroll()
                HasPreScroll = StartScrollPage > 1
        end Function

        public Function HasNextScroll()
                HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
        end Function

        public Function StartScrollPage()
                StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
        end Function

end Class

%>
