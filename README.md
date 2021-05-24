# Generate-an-incomplete-list-of-youth-college-studies
> 此项目利用Excel宏生成青年大学习未完成名单--适用于江西省青年大学习后台导出系统（云瓣科技）

------

### 具体的使用步骤

1. 下载并打开[`生成器/default.xlsm`](https://download.fastgit.org/yjn140/Generate-an-incomplete-list-of-youth-college-studies/releases/download/1.1/default.xlsm)
2. 点击一下“打开/保存”团员名单，把需要检查完成情况的人员名单替换上去。<u>班上团员名单或者全班的名单，取决于每个学校的要求----团员必须做&每个人都要做</u>
3. 把这个excel文件保存，以后每次都可以使用这个文件执行代码后，把未完成名单拷贝到其他excel表格
4. 从[江西青年大学习管理后台](https://jxtw.h5yunban.cn/jxtw-qndxx/admin/login.php)下载某一团支部，某一期的青年大学习完成情况。<u>一个.csv文件</u>
5. 把这个`.csv文件`更名为`导出文件.csv`,并和`生成器.xlsm`放在同一个目录（例如桌面）
6. 在`生成器/default.xlsm`界面点击**生成**
7. 然后就会生成好名单

------

### 视频介绍


<div style="position: relative; width: 100%; height: 0; padding-bottom: 75%;">
<iframe src="//player.bilibili.com/player.html?aid=796797075&bvid=BV1eC4y1t7Bg&cid=226136910&page=1" scrolling="no" border="0" frameborder="no" framespacing="0" allowfullscreen="true" style="position: absolute; width: 100%; height: 100%; Left: 0; top: 0;" ></iframe> </div>




### 更新日志

- 2021.4.15  上传生成器
- 2021.5.18  增加对.csv文件是否更名的判i断
- 2021.5.23 彻底解决数据表刷新不顺序执行的问题 [网上找到的解决方案](http://club.excelhome.net/thread-1131212-1-1.html)     （大概原因是`ActiveWorkbook.RefreshAll`这个函数在vba里面并不是顺序执行，而是在所有函数执行完之后执行。当遇到上次所使用到的.csv文件的数据被保存在生成器中，下一次制作就会因为代码无法顺序执行而导致生成出来的未完成名单出错。解决方法就是在数据--查询链接 里面关闭后台刷新）