<%@ language=vbscript %>
    <% option explicit %>
        <% '#############################################################
' Description : ���� ����ǰ ���� '	History		: 2022.01.05 ������ ����
' ############################################################# %>
            <!-- #include virtual="/lib/function.asp"-->
            <!-- #include virtual="/lib/db/dbopen.asp" -->
            <!-- #include virtual="/lib/util/htmllib.asp" -->
            <!-- #include virtual="/admin/incSessionAdmin.asp" -->
            <!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
            </p>
            <link rel="stylesheet" type="text/css" href="/css/commonV20.css" />
            <div>
                <!-- �������� ����� -->
                <div class="content paidGift add">
                    <div class="paidGift_top">
                        <div class="search_wrap">
                            <select name="" id="select01">
                                <option value="">����</option>
                                <option value="">�÷�������</option>
                                <option value="">�������ǰ</option>
                            </select>
                            <select name="" id="select02">
                                <option value="">�������</option>
                                <option value="">������</option>
                                <option value="">���࿹��</option>
                                <option value="">����</option>
                                <option value="">������</option>
                            </select>
                            <div class="input_wrap">
                                <select name="" id="select03">
                                    <option value="">���μ�</option>
                                    <option value="">�����</option>
                                </select>
                                <span></span>
                                <button class="btn_select">�����ϱ�</button>
                                <li class="selected" style="display:none;">������<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                            </div>
                             <div class="input_wrap">
                                <select name="" id="select04">
                                    <option value="">��������</option>
                                    <option value="">���� ù ����</option>
                                    <option value="">�ֱ� �����ݾ�</option>
                                    <option value="">ȸ�����</option>
                                    <option value="">���űݾ�</option>
                                    <option value="">����Ƚ��</option>
                                    <option value="">��ǰ</option>
                                    <option value="">ī�װ�</option>
                                    <option value="">�귣��</option>
                                    <option value="">��ȹ��/�̺�Ʈ</option>
                                </select>
                                <li class="selected" style="display:none;">��ȹ��/�̺�Ʈ<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                                <li class="selected" style="display:none;">��ǰ<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                            </div>
                            <div class="input_wrap">
                                <select name="" id="select05">
                                    <option value="">��������/��ǰ/����ǰ��</option>
                                    <option value="">�������� ��ȣ</option>
                                    <option value="">����ǰ�ڵ�</option>
                                    <option value="">��ǰ�ڵ�</option>
                                </select>
                                <span></span>
                                <input type="text" placeholder="�˻�� �Է����ּ���">
                            </div>
                            <button class="btn_search"><img src="https://webadmin.10x10.co.kr/images/icon/search.png">�˻��ϱ�</button>
                        </div>
                        <div class="tgl_wrap">
                            <span>���� ����� �������ø� ����</span>
                            <div class="tgl_btn">
                                <input type="checkbox" id="tgl_btn_my">
                                <label for="tgl_btn_my"></label>
                            </div>
                        </div>
                    </div>
                    <div class="paidGift_aside">
                        <div class="list_wrap">
                            <div class="list_top">
                                <li>�� <span>103</span>��</li>
                            </div>
                            <div class="list_cont">
                                <a href=""><div class="cont cont_new">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[����] �� ��������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>���� ù ����</span></li>
                                        <li><span></span></li>
                                        <li class="state st01" style="display:none;"></li>
                                        <p>12:31:20 �ڵ�����</p>
                                    </ul>
                                    <button class="delete"><img src="https://webadmin.10x10.co.kr/images/icon/trash_red.png"></button>
                                </div></a>
                                <!-- �ӽ����� �������� -->
                                <a href=""><div class="cont" style="display:none;">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st03">�ӽ�����</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st01">���¿���</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st02">������</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st02">������</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st02">������</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st01">���¿���</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st03">����</li>
                                    </ul>
                                </div></a>
                            </div>
                            <div class="list_bottom">
                                <ul class="pagination">
                                    <li class="on"><a>1</a></li>
                                    <li class=""><a>2</a></li>
                                    <li class=""><a>3</a></li>
                                    <li class=""><a>4</a></li>
                                    <li class=""><a>5</a></li> 
                                    <li class=""><a>></a></li>
                                    <li class=""><a>>></a></li>
                                </ul>
                            </div>
                        </div>
                    </div>
                    <div class="paidGift_section">
                    <!-- ���ǳ��� : on -->
                        <div class="steps">
                            <li class="step on">���� �������� ����</li>
                            <li class="step"><span></span>�Ⱓ����</li>
                            <li class="step"><span></span>���ü���</li>
                            <li class="step"><span></span>��������</li>
                        </div>
                        <div class="step_wrap step01">
                            <div class="step_noti on"><span>���� ���� �ֹ����� ���������� ���������� ������ ����ڸ� �����մϴ�. � �����ڿ��� ������ �����ϰ� �ͳ���?</span></div>
                            <div class="btn_group">
                                <button class="type01 step01_01 on"><span>��� ������</span><p class="img"></p></button>
                                <button class="type01 step01_02"><span>Ư�� ������</span><p class="img"></p></button>
                            </div>
                            <!-- ��� ������ step_cont step01_01 on -->
                            <div class="step_cont step01_01">
                                <li>��� �����ڿ��� ���������� �����մϴ�. �Ʒ� '����' ��ư�� �����ּ���.</li>
                            </div>
                            <!-- ��� ������ step_cont step01_02 on -->
                            <div class="step_cont step01_02 on">
                                <li>���õ� ���ǿ� ��� �ش��ϴ� �����ڿ��Ը� ������ �����մϴ�. ������ ���� �� '����' ��ư�� �����ּ���.</li>
                                <div class="step_detail">
                                    <div class="step_detail_top">
                                        <h3>���� �����ϱ�</h3>
                                        <span>*������ ���� ���� �����ϸ�, ������̺�Ʈ�� �Բ� ������ �� �����ϴ�.</span>
                                    </div>
                                    <div class="step_detail_list">
                                        <button class="type02 on">
                                            <h3>���� ù ����</h3>
                                            <li>���� �� ù ������ ���</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>�ֱ� �����ݾ� ����</h3>
                                            <li>�ֱ� 5���� �� ��ۿϷ�� �ֹ����� �� �����ݾ�</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>ȸ����� �ش�</h3>
                                            <li>�ſ� ���ŵǴ� ȸ����� ����</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>���űݾ� ����</h3>
                                            <li>���� �ֹ� ���� ���űݾ� ����</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>����Ƚ�� ����</h3>
                                            <li>�ֱ� 5���� �� ��ۿϷ�� �ֹ�Ƚ�� ����</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>������ ��ǰ ����</h3>
                                            <li>Ư�� ��ǰ ���� ��</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>������ ī�װ� ��ǰ ����</h3>
                                            <li>���� �Ǵ� ���� �ֹ� �ǿ��� Ư�� ī�װ� ��ǰ ���� ��</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>������ �귣�� ��ǰ ����</h3>
                                            <li>���� �Ǵ� ���� �ֹ� �ǿ��� Ư�� �귣�� ��ǰ ���Ž�</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>������ ��ȹ��/�̺�Ʈ ��ǰ ����</h3>
                                            <li>Ư�� ��ȹ��/�̺�Ʈ ��ǰ ���� ��</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                    </div>
                                </div>
                                <div class="step_detail">
                                    <div class="step_detail_top">
                                        <h3>����� �̺�Ʈ</h3>
                                        <span>*1���� ���� �����ϸ�, ���� �����ϱ�� �Բ� ������ �� �����ϴ�.</span>
                                         <li>���� ������ �� ���� Ư���� �����̰ų�, ���󵵰� ���� ���������� Ŀ���� �׸����� �����մϴ�.</li>
                                    </div>
                                    <div class="step_detail_list">
                                        <button class="type02 on">
                                            <h3>���̾ ���丮</h3>
                                            <li>���̾ ���丮 ����ǰ ���� ���</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                         <button class="type02">
                                            <h3>�ֳ��̺�Ʈ</h3>
                                            <li>N �ֳ��̺�Ʈ ����ǰ ���� ���</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="steps_bottom">
                            <ul>
                                <button class="delete">���</button>
                                <!-- ��ư ��Ȱ��ȭ : next on -->
                                <button class="next on">����</button>
                            </ul>
                        </div>
                    </div>

                    <!-- ���̾��˾� ���� : paidGift_lyr on -->
                    <div class="lyr paidGift_lyr">
                        <div class="lyr_overlay"></div>
                     <!-- ���ú� ���̾��˾� ���� : lyr_wrap lyr01 on -->
                        <!-- �ֱ� �����ݾ� ���� -->
                        <div class="lyr_wrap lyr01 on">
                            <div class="lyr_top">
                                <li>�ֱ� �����ݾ� ����</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>�Ⱓ ����</h4>
                                    <input type="checkbox" id="lyr01_01" class="type01"><label for="lyr01_01"><span class="circle"></span>�ֱ� 5���� ����</label>
                                    <input type="checkbox" id="lyr01_02" class="type01"><label for="lyr01_02"><span class="circle"></span>�Ⱓ ��������</label>
                                    <!-- �Ⱓ �������� ���ý� ���� -->
                                        <div class="option">
                                            <div class="date_wrap">
                                                <ul class="date">
                                                    <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- ��¥ ���� -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                        <thead>
                                                                            <tr>
                                                                                <th colspan="7">
                                                                                    <ul class="btn_group">
                                                                                        <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                        <li class="btn active">2�� 2022</li>
                                                                                        <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    </ul>
                                                                                </th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div></button>
                                                    <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"><button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- ��¥ ���� -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                    <thead>
                                                                        <tr>
                                                                            <th colspan="7">
                                                                                <ul class="btn_group">
                                                                                    <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    <li class="btn active">2�� 2022</li>
                                                                                    <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                </ul>
                                                                            </th>
                                                                        </tr>
                                                                    </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div></button>
                                                </ul>
                                            </div>
                                        <li class="noti">�ֱ� 5���� �̳��� ��¥�� ������ �� �ֽ��ϴ�.</li>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�ݾ� ����</h4>
                                    <li>������ �Ⱓ���� ������ �ݾ��� ����</li>
                                    <input type="text" id="lyr01_03" placeholder="0"><label for="lyr01_03"></label>
                                    <span>�̻� ������ ��� ���� ����</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ȸ����� -->
                        <div class="lyr_wrap lyr02">
                            <div class="lyr_top">
                                <li>ȸ�����</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <input type="checkbox" id="lyr02_01" class="type01"><label for="lyr02_01"><span class="circle"></span>WHITE</label>
                                    <input type="checkbox" id="lyr02_02" class="type01"><label for="lyr02_02"><span class="circle"></span>RED</label>
                                    <input type="checkbox" id="lyr02_03" class="type01"><label for="lyr02_03"><span class="circle"></span>VIP</label>
                                    <input type="checkbox" id="lyr02_04" class="type01"><label for="lyr02_04"><span class="circle"></span>VIP GOLD</label>
                                    <input type="checkbox" id="lyr02_05" class="type01"><label for="lyr02_05"><span class="circle"></span>VVIP</label>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ���űݾ� ���� -->
                        <div class="lyr_wrap lyr03">
                            <div class="lyr_top">
                                <li>���űݾ� ����</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>��� ����</h4>
                                    <input type="checkbox" id="lyr03_01" class="type01"><label for="lyr03_01"><span class="circle"></span>��ü��ǰ</label>
                                    <input type="checkbox" id="lyr03_02" class="type01"><label for="lyr03_02"><span class="circle"></span>���ٹ�ۻ�ǰ ����</label>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�ݾ� ����</h4>
                                    <li>���� �ֹ����� ���űݾ���</li>
                                    <input type="text" id="lyr01_03" placeholder="0"><label for="lyr01_03"></label>
                                    <span>�̻� �̻��� ��� ����</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ���� Ƚ�� -->
                        <div class="lyr_wrap lyr04">
                            <div class="lyr_top">
                                <li>���� Ƚ��</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>�Ⱓ ����</h4>
                                    <input type="checkbox" id="lyr04_01" class="type01"><label for="lyr04_01"><span class="circle"></span>�ֱ� 5���� ����</label>
                                    <input type="checkbox" id="lyr04_02" class="type01"><label for="lyr04_02"><span class="circle"></span>�Ⱓ ��������</label>
                                    <!-- �Ⱓ �������� ���ý� ���� -->
                                        <div class="option">
                                            <div class="date_wrap">
                                                <ul class="date">
                                                    <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- ��¥ ���� -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                         <thead>
                                                                            <tr>
                                                                                <th colspan="7">
                                                                                    <ul class="btn_group">
                                                                                        <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                        <li class="btn active">2�� 2022</li>
                                                                                        <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    </ul>
                                                                                </th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div>
                                                    </button>
                                                    <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png">
                                                    <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- ��¥ ���� -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                         <thead>
                                                                            <tr>
                                                                                <th colspan="7">
                                                                                    <ul class="btn_group">
                                                                                        <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                        <li class="btn active">2�� 2022</li>
                                                                                        <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    </ul>
                                                                                </th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div>
                                                    </button>
                                                </ul>
                                            </div>
                                        <li class="noti">�ֱ� 5���� �̳��� ��¥�� ������ �� �ֽ��ϴ�.</li>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>���� Ƚ�� ����</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr04_03" ><label for="lyr04_03"><span class="circle"></span><input type="text" placeholder="0" > ȸ �ֹ��� �� ���</label></li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr04_04" ><label for="lyr04_04"><span class="circle"></span><input type="text" placeholder="0" > ȸ �̻� �ֹ��� �� ���</label></li>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ������ ��ǰ ���� -->
                        <div class="lyr_wrap lyr05">
                            <div class="lyr_top">
                                <li>��ǰ �����ϱ�</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>��ǰ ����<span>*�������� �����մϴ�.</span></h4>
                                    <div class="cont_wrap search">
                                        <input type="text" placeholder="��ǰ�ڵ带 �Է����ּ���">
                                        <button class="add btn_blue">�߰��ϱ�</button>
                                        <button class="search btn_white">��ǰ ã��</button>
                                    </div>
                                </div>
                                <!-- ��ǰ�߰� �� ���� -->
                                <div class="lyr_cont">
                                    <div class="ttop">
                                        <span>�� 2��</span>
                                        <button>�����׸� ����</button>
                                    </div>
                                    <div class="table">
                                        <ul class="thead">
                                            <li><input type="checkbox"></li>
                                            <li>��ǰ�ڵ�</li>
                                            <li>��ǥ�̹���/�귣��/��ǰ��</li>
                                            <li></li>
                                            <li>�ǸŰ�</li>
                                            <li></li>
                                        </ul>
                                    <div class="tbody_wrap">
                                            <ul class="tbody">
                                                <li><input type="checkbox"></li>
                                                <li>12345678</li>
                                                <li>
                                                    <img class="prd_img">
                                                    <ul><p class="prd_brand">dailylike</p>
                                                    <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                </li>
                                                <li>7% ����</li>
                                                    <li>
                                                    <p class="price01">38,000</p>
                                                    <p class="price02">32,300</p>
                                                </li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                            </ul>
                                        </div>
                                        <div class="tbody_wrap">
                                            <ul class="tbody">
                                                <li><input type="checkbox"></li>
                                                <li>12345678</li>
                                                <li>
                                                    <img class="prd_img" >
                                                    <ul><p class="prd_brand">dailylike</p>
                                                    <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                </li>
                                                <li>7% ����</li>
                                                    <li>
                                                    <p class="price01">38,000</p>
                                                    <p class="price02">32,300</p>
                                                </li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </ul>
                                        </div>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>��ǰ ���� ����</h4>
                                    <input type="checkbox" id="lyr05_01" class="type01"><label for="lyr05_01"><span class="circle"></span>1���� ���� ��</label>
                                    <input type="checkbox" id="lyr05_02" class="type01"><label for="lyr05_02"><span class="circle"></span>��� ��ǰ ���� ��</label>
                                    <li class="type01 noti">������ ��ǰ �� 1���� �����ϸ� ������ �����մϴ�.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>���� ����</h4>
                                     <!-- ���ż��� ���� ���ý� ���� -->
                                    <div class="option t02">
                                        <li>������ ��ǰ��</li>
                                        <input type="text" id="lyr05_05" value="1"><label for="lyr05_05"></label>
                                        <span>�̻� ������ ��� ���� ����</span>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ������ ī�װ� ��ǰ ���� -->
                        <div class="lyr_wrap lyr06">
                            <div class="lyr_top">
                                <li>������ ī�װ� ��ǰ ���� ��</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <div class="cont_wrap type03">
                                        <input class="type03" type="checkbox" id="lyr06_01"><label for="lyr06_01">���� �ֹ��� ����</label>
                                        <input class="type03" type="checkbox" id="lyr06_02"><label for="lyr06_02">���� �ֹ��� ����<li>�ֱ� 5���� �̳�</li></label>
                                    </div>
                                </div>
                                <!-- ���� �ֹ��� ���� ���ý� ���� -->
                                 <div class="lyr_cont">
                                        <h4>�Ⱓ ����</h4>
                                        <input type="checkbox" id="lyr06_03" class="type01"><label for="lyr06_03"><span class="circle"></span>�ֱ� 5���� ����</label>
                                        <input type="checkbox" id="lyr06_04" class="type01"><label for="lyr06_04"><span class="circle"></span>�Ⱓ ��������</label>
                                        <!-- �Ⱓ �������� ���ý� ���� -->
                                        <div class="option">
                                             <div class="date_wrap">
                                                <ul class="date">
                                                    <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- ��¥ ���� -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                         <thead>
                                                                            <tr>
                                                                                <th colspan="7">
                                                                                    <ul class="btn_group">
                                                                                        <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                        <li class="btn active">2�� 2022</li>
                                                                                        <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    </ul>
                                                                                </th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div>
                                                    </button>
                                                    <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png">
                                                    <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- ��¥ ���� -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                         <thead>
                                                                            <tr>
                                                                                <th colspan="7">
                                                                                    <ul class="btn_group">
                                                                                        <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                        <li class="btn active">2�� 2022</li>
                                                                                        <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    </ul>
                                                                                </th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div>
                                                    </button>
                                                </ul>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>ī�װ� ����<span>*�������� �����մϴ�.</span></h4>
                                        <select name="" id="">
                                            <option value="">1depth</option>
                                        </select>
                                        <select name="" id="">
                                            <option value="">2depth</option>
                                        </select>
                                        <button class="add btn_blue">ī�װ� �߰�</button>
                                         <!-- ī�װ� �߰��� ���� -->
                                        <div class="option">
                                            <div class="option_added">
                                                <li>�����ι���</li>
                                                <li>></li>
                                                <li>���̾/�÷���</li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                            <div class="option_added">
                                                <li>�����ι���</li>
                                                <li>></li>
                                                <li>���ڷ��̼�</li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�ݾ�/���� ����</h4>
                                    <div class="cont_wrap lyr06">
                                        <input type="checkbox" id="lyr06_06" class="type01"><label for="lyr06_06"><span class="circle"></span>�ּұݾ� ����</label>
                                        <input type="checkbox" id="lyr06_07" class="type01"><label for="lyr06_07"><span class="circle"></span>���ż��� ����</label>
                                        <!-- ���� �ֹ��� ���� ���ý� ���� -->
                                        <input type="checkbox" id="lyr06_08" class="type01"><label for="lyr06_08"><span class="circle"></span>�ֹ�Ƚ�� ����</label>
                                    </div>
                                    <!-- ���� ���� ���ý� ���� -->
                                    <div class="option">
                                        <li class="noti">������ ī�װ� ��ǰ �� 1���� �����ϸ� ������ �����մϴ�.</li>
                                    </div>
                                    <!-- �ּұݾ� ���� ���ý� ���� -->
                                    <div class="option t02 lyr06_06">
                                        <li>������ ī�װ� ��ǰ��</li>
                                        <input type="text" id="lyr06_09" placeholder="0" value="0"><label for="lyr06_09"></label>
                                        <span>�̻� ������ ��� ���� ����</span>
                                    </div>
                                    <!-- ���ż��� ���� ���ý� ���� -->
                                    <div class="option t02 lyr06_07">
                                        <li>������ ī�װ� ��ǰ��</li>
                                        <input type="text" id="lyr06_10" value="1"><label for="lyr06_10"></label>
                                        <span>�̻� ������ ��� ���� ����</span>
                                    </div>
                                    <!-- �ֹ�Ƚ�� ���� ���ý� ���� -->
                                    <div class="option t02 lyr06_08">
                                        <li>������ �Ⱓ���� ������ ī�װ� ��ǰ�� �����Ͽ�</li>
                                        <input type="text" id="lyr06_11" value="0"><label for="lyr06_11"></label>
                                        <span>�̻� �ֹ��� ��� ���� ����</span>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <!-- ��ư ��Ȱ��ȭ : submit on -->
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ������ �귣�� ��ǰ ���� -->
                        <div class="lyr_wrap lyr07">
                            <div class="lyr_top">
                                <li>������ �귣�� ��ǰ ���� ��</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <div class="cont_wrap type03">
                                        <input class="type03" type="checkbox" id="lyr07_01"><label for="lyr07_01">���� �ֹ��� ����</label>
                                        <input class="type03" type="checkbox" id="lyr07_02"><label for="lyr07_02">���� �ֹ��� ����<li>�ֱ� 5���� �̳�</li></label>
                                    </div>
                                </div>
                                <!-- ���� �ֹ��� ���� ���ý� ���� -->
                                 <div class="lyr_cont">
                                        <h4>�Ⱓ ����</h4>
                                        <input type="checkbox" id="lyr07_03" class="type01"><label for="lyr07_03"><span class="circle"></span>�ֱ� 5���� ����</label>
                                        <input type="checkbox" id="lyr07_04" class="type01"><label for="lyr07_04"><span class="circle"></span>�Ⱓ ��������</label>
                                        <!-- �Ⱓ �������� ���ý� ���� -->
                                        <div class="option">
                                             <div class="date_wrap">
                                               <ul class="date">
                                                    <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- ��¥ ���� -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                         <thead>
                                                                            <tr>
                                                                                <th colspan="7">
                                                                                    <ul class="btn_group">
                                                                                        <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                        <li class="btn active">2�� 2022</li>
                                                                                        <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    </ul>
                                                                                </th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div>
                                                    </button>
                                                    <img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png">
                                                    <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- ��¥ ���� -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                         <thead>
                                                                            <tr>
                                                                                <th colspan="7">
                                                                                    <ul class="btn_group">
                                                                                        <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                        <li class="btn active">2�� 2022</li>
                                                                                        <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    </ul>
                                                                                </th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div>
                                                    </button>
                                                </ul>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�귣�� ����<span>*�������� �����մϴ�.</span></h4>
                                        <div class="cont_wrap search">
                                            <input type="text" placeholder="�귣��ID�� �Է����ּ���">
                                            <button class="add btn_blue">�߰��ϱ�</button>
                                            <button class="search btn_white">�귣�� ã��</button>
                                        </div>
                                         <!-- �귣�� �߰��� ���� -->
                                        <div class="option">
                                            <div class="option_added">
                                                <li>PEANUTS10X10</li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                            <div class="option_added">
                                                <li>SANRIO</li>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�ݾ�/���� ����</h4>
                                    <div class="cont_wrap lyr07">
                                        <input type="checkbox" id="lyr07_06" class="type01"><label for="lyr07_06"><span class="circle"></span>�ּұݾ� ����</label>
                                        <input type="checkbox" id="lyr07_07" class="type01"><label for="lyr07_07"><span class="circle"></span>���ż��� ����</label>
                                        <!-- ���� �ֹ��� ���� ���ý� ���� -->
                                        <input type="checkbox" id="lyr07_08" class="type01"><label for="lyr07_08"><span class="circle"></span>�ֹ�Ƚ�� ����</label>
                                    </div>
                                    <!-- ���� ���� ���ý� ���� -->
                                    <div class="option">
                                        <li class="noti">������ �귣�� ��ǰ �� 1���� �����ϸ� ������ �����մϴ�.</li>
                                    </div>
                                    <!-- �ּұݾ� ���� ���ý� ���� -->
                                    <div class="option t02 lyr07_06">
                                        <li>������ �귣�� ��ǰ��</li>
                                        <input type="text" id="lyr07_09" value="0"><label for="lyr07_09"></label>
                                        <span>�̻� ������ ��� ���� ����</span>
                                    </div>
                                    <!-- ���ż��� ���� ���ý� ���� -->
                                    <div class="option t02 lyr07_07">
                                        <li>������ �귣�� ��ǰ��</li>
                                        <input type="text" id="lyr07_10" value="1"><label for="lyr07_10"></label>
                                        <span>�̻� ������ ��� ���� ����</span>
                                    </div>
                                    <!-- �ֹ�Ƚ�� ���� ���ý� ���� -->
                                    <div class="option t02 lyr07_08">
                                        <li>������ �Ⱓ���� ������ �귣�� ��ǰ�� �����Ͽ�</li>
                                        <input type="text" id="lyr07_11" value="1"><label for="lyr07_11"></label>
                                        <span>�̻� �ֹ��� ��� ���� ����</span>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <!-- ��ư ��Ȱ��ȭ : submit on -->
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ������ ��ȹ��/�̺�Ʈ ��ǰ ���� -->
                        <div class="lyr_wrap lyr08">
                            <div class="lyr_top">
                                <li>������ ��ȹ��/�̺�Ʈ ��ǰ ���� ��</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>��ȹ��/�̺�Ʈ ����<span>*�������� �����մϴ�.</span></h4>
                                        <div class="cont_wrap search">
                                            <input type="text" placeholder="�̺�Ʈ �ڵ带 �Է����ּ���">
                                            <button class="add btn_blue">�߰��ϱ�</button>
                                        </div>
                                         <!-- �̺�Ʈ�ڵ� �߰��� ���� -->
                                        <div class="option">
                                            <div class="option_added">
                                                <li class="e_img"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></li>
                                                <ul>
                                                    <li class="e_code">12345678</li>
                                                    <li class="e_name">�ų����� ����ī�� ��Ű��</li>
                                                </ul>
                                                <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                             <div class="option_added">
                                                <li class="e_img"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></li>
                                                <ul>
                                                    <li class="e_code">12345678</li>
                                                    <li class="e_name">�ų����� ����ī�� ��Ű��</li>
                                                </ul>
                                                 <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"  alt=""></button></li>
                                            </div>
                                        </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>��ǰ ���� ����</h4>
                                        <input type="checkbox" id="lyr08_01" class="type01"><label for="lyr08_01"><span class="circle"></span>1���� ���� ��</label>
                                        <input type="checkbox" id="lyr08_02" class="type01"><label for="lyr08_02"><span class="circle"></span>��� ��ǰ ���� ��</label>
                                    <li class="noti">������ �귣�� ��ǰ �� 1���� �����ϸ� ������ �����մϴ�.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�ݾ�/���� ����</h4>
                                        <input type="checkbox" id="lyr08_04" class="type01"><label for="lyr08_04"><span class="circle"></span>�ּұݾ� ����</label>
                                        <input type="checkbox" id="lyr08_05" class="type01"><label for="lyr08_05"><span class="circle"></span>���ż��� ����</label>
                                    <!-- �ּұݾ� ���� ���ý� ���� -->
                                    <div class="option t02 lyr08_04">
                                        <li>������ ��ȹ�� ��ǰ��</li>
                                        <input type="text" id="lyr08_06" value="0"><label for="lyr08_06"></label>
                                        <span>�̻� ������ ��� ���� ����</span>
                                    </div>
                                    <!-- ���ż��� ���� ���ý� ���� -->
                                    <div class="option t02 lyr08_05">
                                        <li>������ ��ȹ�� ��ǰ��</li>
                                        <input type="text" id="lyr08_07" value="0"><label for="lyr08_07"></label>
                                        <span>�̻� ������ ��� ���� ����</span>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <!-- ��ư ��Ȱ��ȭ : submit on -->
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ���űݾ� ���� -->
                        <div class="lyr_wrap lyr09">
                            <div class="lyr_top">
                                <li>���űݾ� ����</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>��� ����</h4>
                                        <input type="checkbox" id="lyr09_01" class="type01"><label for="lyr09_01"><span class="circle"></span>��ü��ǰ</label>
                                        <input type="checkbox" id="lyr09_02" class="type01"><label for="lyr09_02"><span class="circle"></span>���ٹ�ۻ�ǰ ����</label>
                                </div>
                                <div class="lyr_cont lyr09_03">
                                    <h4>�ݾ� ����</h4>
                                    <input type="text" id="lyr09_03"><label for="lyr09_03"></label>
                                    <span>�̻��� ��� ����</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <button class="initial">���� �ʱ�ȭ</button>
                                <ul>
                                    <button class="cancel">���</button>
                                    <!-- ��ư ��Ȱ��ȭ : submit on -->
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->