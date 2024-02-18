<%

Class CCSCenterIppbxUserItem
        public Flocalcallno
        public Fuserid
        public Fuseyn
        public Flastupdate

        Private Sub Class_Initialize()

        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class

Class CCSCenterIppbxUser
        public FItemList()
        public FOneItem

        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

        public Sub GetCSCenterIppbxUserList()
                dim i,sqlStr

                sqlStr = " select top 300 localcallno, userid, useyn, lastupdate "
                sqlStr = sqlStr + " from db_cs.dbo.tbl_cs_ippbx_user "
                sqlStr = sqlStr + " order by localcallno "
                'response.write sqlStr
                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount
                FTotalCount = FResultCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new CCSCenterIppbxUserItem

                                FItemList(i).Flocalcallno        = rsget("localcallno")
                                FItemList(i).Fuserid             = rsget("userid")
                                FItemList(i).Fuseyn              = rsget("useyn")
                                FItemList(i).Flastupdate         = rsget("lastupdate")

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