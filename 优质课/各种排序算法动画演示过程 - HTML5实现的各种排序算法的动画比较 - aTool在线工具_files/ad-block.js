// append ad block into doms
/*
  How to use it ?
  var ads = [
    [text, url, vip, date],
    [text, url, vip, date]
  ];
  var dom = document.getElementById('ads');
  appendAdBlock(dom, ads)
*/
function appendAdBlock(dom, ads) {
  var html = '';
  html += '<div class="ad-block"><ul>';
  for (var i = 0; i < ads.length; i++) {
    // 超时
    if (new Date(ads[i][3]) >= new Date()) {
      html += '<li><a href="' + ads[i][1] + '" target="_blank" class="' + (ads[i][2] ? 'red': '') + '">' + ads[i][0] + '</a></li>';  
    }
  }
  html += '</ul></div>';
  dom.innerHTML = html;
};

var ads = [
  ['2017年最牛逼的伪原创工具', 'http://www.zhipaiwu.com/', 0, '2017-05-31'],
  ['设计神器', 'http://www.tubangzhu.com/?atool', 1, '2017-07-15'],
  ['【b3qude】阿里云 vps 推荐码', 'http://www.atool.org/', 0, '2020-02-01'],
  ['短信验证码，低至0.035元每条', 'https://www.mysubmail.com/sms?s=atool', 1, '2017-06-28'],
  ['banner 免费 logo 制作', 'http://www.zhaoxi.net/', 1, '2017-11-20'],
  ['阿里云全产品优惠券', 'http://blog.5ibc.net/go/aliyun', 1, '2017-12-06'],
  ['梦达网贷工作室代办贷款信用卡', 'http://www.mengda918.com/', 0, '2017-04-07'],
  ['拥有一套自己的网页聊天系统', 'http://layim.layui.com/?from=atool', 0, '2017-07-31'],
  ['加群:广告位(150/月)', 'http://www.atool.org/', 1, '2020-02-01'],
  ['短信验证码每条3分6', 'http://www.ucpaas.com/product/sms.html?utm_source=%E7%BD%91%E7%AB%99%E5%90%88%E4%BD%9C&utm_medium=cpt&amp;utm_term=atool%E6%96%87%E5%AD%97%E9%93%BE&utm_content=&utm_campaign=%E7%9F%AD%E4%BF%A1%E9%AA%8C%E8%AF%81%E7%A0%81&cl_sr=%E7%BD%91%E7%AB%99%E5%90%88%E4%BD%9C&cl_ctnm=atool%E6%96%87%E5%AD%97%E9%93%BE&amp;id=105930', 1, '2017-08-08'],
  ['网上贷款找梦达网贷', 'http://1139581.b2b.xbaixing.com/ ', 1, '2017-08-09'],
  ['免费抢红包可提现详情看公告', 'http://d.hd9166.com/Reg/f62fa9fda86caa4b  ', 0, '2017-08-15'],
  ['百度快速排名包收录 ', 'http://www.dayidayou.com/tui#atool   ', 1, '2017-09-10'],
  ['百度快速排名包收录 ', 'http://www.dayidayou.com/xinpinshangjia/z/   ', 1, '2017-11-25'],
  ['VPS服务器免费招代理', 'http://www.93lh.com', 1, '2017-11-23'],
  ['全自动博客群发软件', 'http://soft.dayidayou.com', 1, '2017-11-25'],
  ['【推荐】专业原创文章代写', 'http://434.com.cn/fuwu/244.html', 1, '2017-12-05'],
];
var dom = document.getElementById('ad-blocks');
appendAdBlock(dom, ads);
