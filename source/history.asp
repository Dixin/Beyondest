<!-- #include file="include/config_other.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com v3.6.1
' 
'           http://beyondest.com
' 
'*******************************************************************

tit="历史"
tit_fir=""

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------

response.write format_img("llogin.jpg")
call format_login()
call main_update_view()
call main_stat("","jt1",1,1,1)
call user_data_top("bbs_counter","jt12",1,10)
call vote_type(1,1,"","f7f7f7")

'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%response.write format_img("rhistory.jpg")%></td></tr>
</table>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td width=1 bgcolor="<%response.write web_var(web_color,3)%>"></td><td align=center><%call history()%></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
call web_end(0)




sub history()
dim temph,gang
gang="<table height=1 bgcolor="&web_var(web_color,3)&" width='100%' cellspacing=0 cellpadding=0 border=0><tr><td></td></tr></table>"

temph="<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1983年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center></td><td>黄家驹及叶世荣于琴行老板介绍下认识，发觉彼此音乐兴趣相近，遂联同两位朋友一起组成乐队作音乐交流</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">5月</td><td>因参加结他杂志举行之「山叶结他比赛」需要为乐队命名，结他手邓炜谦把乐队命名为Beyond</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>于结他比赛取得冠军，得奖歌曲获收录在合辑《香港》</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1984年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center></td><td>黄家驹弟弟黄家强加入Beyond，负责弹奏低音结他</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1985年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>4月</td><td>开始筹备首个乐队演唱会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>结他手陈时安要到外地升学离队，因此邀请替音乐会作平面设计工作的黄贯中加入Beyond</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>自资举办「永远等待演唱会」，于坚道明爱中心举行</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1986年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>4月</td><td>推出自资盒带《再见理想》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>于艺穗会举行两场「剖析聚会演唱会」</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>签约Kinn's经理人公司，正式踏入香港乐坛</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>结他手刘志远加入Beyond</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>到台北参加「亚太流行音乐节」，与新加坡、菲律宾、日本等各地乐手同台切磋，交流表演</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1987年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>推出首张EP《永远等待》，由宝丽金出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">5月</td><td>EP《新天地》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>推出首张专辑《亚拉伯跳舞女郎》</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1988年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>3月</td><td>第二张专辑《现代舞台》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">4月</td><td>刘志远离队，Beyond成为四人乐队</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">9月</td><td>签约新艺宝唱片公司，推出第三张专辑《秘密警察》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">10月</td><td>于北京首都体育馆举行两场演唱会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">年底</td><td>出版自传《心内心外》，为电影〈午夜迷墙〉作配乐工作</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1989年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>首次获得电子传媒奖项，包括商业电台「叱咤乐坛组合银奖」；歌曲〈大地〉荣获无线电视「十大劲歌金曲」、香港电台「十大中文金曲」奖项</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">4月</td><td>推出EP《4拍4》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>无线电视剧〈淘气双子星〉首播，黄贯中及黄家强参与演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>出版第四张专辑《Beyond IV》，销量达双白金</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">10月</td><td>联同草蜢拍摄无线电视音乐特辑《够Hit斗玩Beyond+草蜢》</td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">11月</td><td>出版书刊《真的见证》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>于伊利沙伯体育馆举行七场「真的见证演唱会」，反应理想</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>推出专辑《真的见证》，Beyond于此碟重新演译多首曾写给其它歌手歌曲</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1990年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>歌曲《真的爱你》荣获多个乐坛奖项，包括无线电视「十大劲歌金曲」、香港电台「十大中文金曲」等</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">2月</td><td>四子有份参与演出之贺岁电影《吉星拱照》上映</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">4月</td><td>歌曲〈午夜迷墙〉获香港电影金像奖「最佳电影歌曲」提名</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>首次参与宣明会活动「饥馑三十」，并于闭幕礼演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">6月</td><td>电影原声EP《天若有情》出版，碟内收录四首电影歌曲，其中三首由Beyond主唱</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>首次担担纲主演电影《开心鬼救开心鬼》上映，Beyond并负责此片配乐工作，推出EP《战胜心魔》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">8月</td><td>黄家驹参与宣明会与香港电台合办「爱心第一旅」活动，前往巴布亚新畿内亚作亲善探访</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>电影《忍者龟》上映，Beyond参与该片配音工作。专辑《命运派对》出版，此碟更获三白金销量</td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">10月</td><td>拍摄无线电视音乐特辑《劲Band四斗士》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>推出首张国语专辑《大地》，正式发展台湾市场</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>于舞台「Rock N’Roll Band Stand音乐会」表演，此节目由日本NHK电视台举办，并以卫星直播</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1991年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>歌曲《光辉岁月》荣获无线电视「十大劲歌金曲」奖项，《俾面派对》则获得香港电台「十大中文金曲」奖项</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">2月</td><td>四子前往东非国家肯尼亚作亲善探访</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">4月</td><td>第二张国语专辑《光辉岁月》于台湾出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7～9月</td><td>为无线电视拍摄十二集综合节目《Beyond放暑假》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">8月</td><td>电影《Beyond日记之莫欺少年穷》上映</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">9月</td><td>专辑《犹豫》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>《忍者龟II》上映，Beyond再次为该片配音</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>Kinn's重新发行《再见理想》盒带，并推出CD版</td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>首次踏足红馆，举行五场「Beyond生命接触演唱会」</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>黄贯中参与演出之电影《老豆唔怕多》正式上映</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1992年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>《Amani》获得香港电台「十大中文金曲」奖项</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>转投华纳唱片公司，并与日本经理人公司Amuse及唱片公司Fun House Inc签约，正式开拓日本市场</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">3月</td><td>签约滚石唱片，滚石成为台湾发行之国语专辑代理</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>推出专辑《继续革命》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>黄家驹参与演出之电影《笼民》正式上映</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">9月</td><td>首张日语专辑《超越》于日本推出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>于台湾推出国语专辑《信念》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>推出粤语EP《无尽空虚》，此碟同时亦收录日语版《长城》</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1993年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>5月</td><td>推出专辑《乐与怒》，亦是四人时期最后一张专辑</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>于香港电台举办之「Beyond我地呀！Unplugged音乐会」演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>于马来西亚举办「Beyond Unplugged Live演唱会」</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">6月</td><td>于日本推出第二张日语专辑《This is Love I》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>24日凌晨1时（日本时间）于日本东京富士电视台录像游戏节目《Ucchan-nanchan no yarunara yaraneba》环节「对抗Corner」时发生意外，黄家驹不幸从三米高舞台堕下重伤，于东京女子医科大学附属医院留医六日后于30日下午4时15分（日本时间）逝世，终年31岁</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>黄家驹于香港举行丧礼，遗体于将军澳华人永远坟场下葬，其后亦于日本举行追悼会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>黄家驹凭《长城》获得新城电台颁发「劲爆作曲人」奖项，奖项由经理人公司代领</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>华纳唱片推出纪念专辑《遥望黄家驹不死音乐精神特别纪念集92-93》，并收录《为了你为了我》黄家驹重唱版本</td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">8月</td><td>华纳唱片推出另一只纪念专辑《Beyond Word & Music: Final Live with家驹》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">9月</td><td>于台湾推出专辑《海阔天空》，该碟收录多首国语及粤语作品</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">11月</td><td>召开记者招待会宣布正式复出，将以三人姿态出版唱片</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>于商业电台举办之「创作人音乐会」演出，是次音乐会于红馆举行，亦是黄家驹逝世后三子首次演出</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1994年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>歌曲《海阔天空》获得商业电台「我最喜爱的本地创作歌曲」、香港电台「十大中文金曲」；黄家驹更获得香港电台「无休止符纪念奖」、无线电视台「荣誉大奖」奖项</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">2月</td><td>转投香港滚石唱片，滚石成为亚洲地区粤语及国语专辑代理</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">6月</td><td>首次以三人阵容出版专辑《二楼后座》，专辑发行后即成为全城焦点，并创出同时以两首主打歌（《醒你》及《遥远的Paradise》）派台的先河</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>国语专辑《Paradise》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>滚石唱片邀请多位歌手重新演译黄家驹作品，推出专辑《祝您愉快——致家驹》作致敬</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1995年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>5月</td><td>出版粤语专辑《Sound》，此碟录音工作更远赴美国洛杉矶进行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">6月</td><td>于文化中心广场举行「Sound音乐会」，由商业电台协办</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">10月</td><td>出版国语专辑《爱与生活》，此碟突破一向由粤语歌曲改编国语版做法，收录四首全新国语歌曲</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">11月</td><td>与新宝岛康乐队合作，于台湾举行三场「土洋大决战音乐会」</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1996年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>2月</td><td>粤语EP《Beyond得精彩》出版，随EP附送小相集《13周年Beyond的精彩2．3事典》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">3月</td><td>于红馆举行四场「Beyond的精彩演唱会」，观众反应热烈</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">6月</td><td>于马来西亚吉隆坡默迪卡球场举行演唱会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">10月</td><td>商业电台于红馆举行「创作人音乐会声音再生万岁」，Beyond作压轴嘉宾演出</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1997年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>首度夺得商业电台之「叱咤乐坛组合金奖」奖项，此后多年Beyond囊括各大电子传媒组合奖项</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">4月</td><td>推出粤语专辑《请将手放开》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>推出粤语专辑《惊喜》，打破Beyond多年在香港「一年一碟」惯例</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1998年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>自传《拥抱Beyond岁月》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>Beyond获商业电台颁发「叱咤殿堂十大歌手／组合」奖项，黄家驹更获颁发「叱咤殿堂十大作曲人」奖项</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">2月</td><td>出版国语专辑《Here and There》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">3月</td><td>于台湾举行两场「Here and There 演唱会」</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>黄贯中参与演出之电影《生死恋》正式上映</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>推出EP《Action》，收录电影〈轰天炮4〉粤语及国语主题曲</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">11月</td><td>商业电台叱咤903节目「组Band时间」为向Beyond致敬，推出《Beyond精彩十五年：A Tribute to Beyond》专辑，邀请多队乐队重新演译多首Beyond作品</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>粤语专辑《不见不散》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>叶世荣主演电影《爱情传真》正式上映</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>1999年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>3月</td><td>于启德机场飞机库举行一场「Beyond 2000演唱会」，由新城电台协办</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">8月</td><td>参与台湾「MTV夏日高峰演唱会」，联同台湾、日本、韩国等歌手乐队演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>粤语专辑《Goodtime》出版，Beyond首次为自己唱片作监制</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>黄贯中主演电影《生命楂Fit人》正式上映</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>一连三场「Good Time演唱会」于会展新翼Hall 3举行，之后进入活动休止期</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>2000年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>年初</td><td>黄家强拍摄电影《敌对》</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">11月</td><td>黄贯中参与「剧场空间」剧社之音乐剧〈梦断维港〉演出，该话剧于文化中心广场举行六场</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">年底</td><td>黄家强担任电影《阴阳爱》配乐工作</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>2001年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>黄贯中推出首张个人专辑《Paul Wong》，此碟邀请多位音乐人包括Mick Karn、Funky、Sugizo、张亚东等客串</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>黄贯中于坚道明爱中心举行「组Band时间五周年特Gig黄贯中音乐会」</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>黄贯中参与演出之电影《枕边凶灵》上映</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">8月</td><td>叶世荣首张个人EP《美丽的时光机器》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>黄贯中专辑《黑白》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>黄家强于高山剧场Band Show「Power Punch」作表演嘉宾</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>2002年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>黄贯中首度夺得商业电台之「叱咤乐坛唱作人金奖」</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">2～4月</td><td>黄贯中为该年宣明会活动「饥馑三十」饥馑之星，并前往缅甸作亲善探访</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">4月</td><td>黄贯中EP《同根》出版，纪念缅甸之行歌曲《同根》邀得黄家强、叶世荣负责乐器弹奏部分</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>叶世荣于高山剧场「Rock On 2002」Band Show演出，同年并参与电影《现代灰姑娘》演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">10月</td><td>黄家强首张个人专辑《Be Right Back》出版，多位音乐人包括Michael Brook、Seasons Lee参与此碟制作</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">11月</td><td>黄贯中于红馆举行两场「Play It Loud演唱会」，黄家强、叶世荣为是次演唱会嘉宾，成为三子作个人发展后首次同台演出，并宣布将于翌年举行Beyond二十周年纪念演唱会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>黄贯中个人专辑《Play It Loud》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>黄家强于坚道明爱中心举行「组Band时间家强第一Gig音乐会」</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>黄家强个人专辑《Be Right Back》夺得IFPI销量奖项</td></tr></table>"&gang

temph=temph&"<table width='100%'  border='0' cellspacing='3' cellpadding='0'><tr><td width=60 align=center>2003年</td><td bgcolor="&web_var(web_color,6)&" width=40 align=center>1月</td><td>Beyond获香港电台颁发「金曲银禧荣誉大奖」</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">3月</td><td>黄家强专辑《叱咤903．组Band时间家强第一Gig－新曲 + Live》出版，碟内收录三首全新作品</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">4月</td><td>Beyond三子再度合作，出版EP《Together》，该碟也是自从《再见理想》后再度自资出版之专辑</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">4～5月</td><td>(30/4-4/5) 一连五场「Beyond超越Beyond Live」于红馆举行，由于反应口碑俱佳，主办机构宣布加场</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">5月</td><td>(18/5) 世荣出席于维园的Farm Band Show</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">6月</td><td>出版《抗战二十年相集》，友力出版发行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(21-23/6) 一连三场「Beyond超越Beyond Live」（Part II）于红馆举行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">7月</td><td>叶世荣于九展「Rock On 2003」Band Show演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(19/7) Beyond于尖沙咀HMV举行签名会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(31/7) 叶世荣于「生力清啤Wild Day Out BAR GIG」演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">8月</td><td>叶世荣个人EP《Remember You》出版</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(6/8) 叶世荣于「香港漫画节」演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(8/8) 叶世荣于「Sound & Vision Festival」演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(9/8) 叶世荣于「17 Attitude? Music No Limit Showtime」演出</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(10/8) 于北京举行《Beyond 超越 Beyond 北京巡回演唱会》记者招待会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(13/8) 出席通利琴行50周年「Life Rocks HK2003」音乐会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(19/8) 叶世荣于尖沙咀帝苑酒店 FALCON举行生日会，同Fans一齐过生日</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(23/8) 开始举行海外巡回演唱会，于北京工人体育场举行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">9月</td><td>(5-6/9) 叶世荣于新加坡举行两场Live Show</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(23-24/9) 于马来西亚举行歌迷见面会及签名会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">10月</td><td>(11/10) 第三站海外巡回演唱会于马来西亚吉隆坡默迪卡球场 (Merdeka Stadium) 举行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">11月</td><td>(19/7) (8/11) 黄家强于尖沙咀Hard Rock Cafe举行生日会</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(15/11) 第四站海外巡回演唱会于广州举行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(22/11) 第五站海外巡回演唱会于上海举行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(27/11) 第六站海外巡回演唱会于美国大西洋城举行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&">12月</td><td>(4/12) 第七站海外巡回演唱会于多伦多 Niagara Arena 举行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(8/12) 第八站海外巡回演唱会于温哥华 Queen Elizabeth Theatre 举行</td></tr><tr height=1><td></td><td colspan=2 background=images/bg_dian.gif><td></tr><tr><td></td><td align=center bgcolor="&web_var(web_color,6)&"></td><td>(8/12) 第九站海外巡回演唱会于深圳举行</td></tr></table>"

response.write format_barc("<font class=end><b>Beyond抗战事迹</font></b>",temph,3,0,5)
end sub%>