# Generate-an-incomplete-list-of-youth-college-studies
> 此项目利用Excel宏生成青年大学习未完成名单--适用于江西省青年大学习后台导出系统（云瓣科技）

------

### 具体的使用步骤

1. 下载并打开[`生成器/default.xlsm`](https://github.com/yjn140/Generate-an-incomplete-list-of-youth-college-studies/raw/main/%E7%94%9F%E6%88%90%E5%99%A8/%E7%94%9F%E6%88%90%E5%99%A8.xlsm)
2. 点击一下“打开/保存”团员名单，把需要检查完成情况的人员名单替换上去。<u>班上团员名单或者全班的名单，取决于每个学校的要求----团员必须做&每个人都要做</u>
3. 把这个excel文件保存，以后每次都可以使用这个文件执行代码后，把未完成名单拷贝到其他excel表格
4. 从[江西青年大学习管理后台](https://jxtw.h5yunban.cn/jxtw-qndxx/admin/login.php)下载某一团支部，某一期的青年大学习完成情况。<u>一个.csv文件</u>
5. 把这个`.csv文件`更名为`导出文件.csv`,并和`生成器.xlsm`放在同一个目录（例如桌面）
6. 在`生成器/default.xlsm`界面点击**生成**
7. 然后就会生成好名单

------

### 视频介绍

### 更新日志

2021.4.15  上传生成器

2021.5.18  修复因处理时间过快而数据表未更新完全造成的名单错误；增加对.csv文件是否更名的判断