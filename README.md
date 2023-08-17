# Auto-Privacy-Leakage-Detection

处理数据[附件链接](https://cpipc.acge.org.cn/sysFile/downFile.do?fileId=3f03175a2ddb46328f35a1edd8924236)

|解题思路|需实现功能|可能的技术手段|
|:-:|:-:|:-:|
|1.数据预处理| 1.解压数据，将所有文件转换成文本格式，便于后续文本处理<br>2.对文本清洗和预处理，去除无关信息、特殊字符，保持格式一致性 | 根据不同文件类型采取相应解析方法 |
|2.敏感信息提取| 1.从预处理的文本中提取敏感信息<br>2.需要提取包含但不仅限于目标地址、端口、用户名、密码、密码hash、Token、AK/SK等 |自然语言处理（NLP）技术，正则表达式 |
|3.上下文关联| 1.能关联上下文，输出关联的敏感信息对 | |
|4.评估与优化| 1.提升敏感信息提取数量和准确性<br>2.提高算法效率（时间） | 1.多线程或并行处理技术<br>2.优化算法逻辑、数据结构|


|py文件|实现功能|
|:-:|:-:|
|main.py| 主函数 |
|classify_files.py| 1.遍历所有文件，划分文件类型<br>2.输出data_info.xlsx |
|DataPreprocessing.py| 1.解析文件，输出文本格式数据<br>2.输出preprocessed_data.txt<br>和操作日志preprocess_log.txt |


## **“华为杯”第二届中国研究生网络安全创新大赛揭榜挑战赛赛题：富文本敏感信息泄露检测**

题目1：富文本敏感信息泄露检测

**竞赛题目详细描述：**

给定一个包含大量文件的压缩包（见附件），选手需要编写程序，自动化提取其中敏感的认证信息。压缩包中的文件格式包含但不仅限于Word、Excel、PowerPoint、txt、各类系统配置文件、图片、聊天纪录等。

**挑战内容，即程序需实现功能如下：**

1、提取敏感认证信息，包含但不仅限于目标地址、端口、用户名、密码、密码hash、Token、AK/SK等。

2、程序应能关联上下文，输出关联的敏感信息对，如{"ip":"127.0.0.1","port":3306,"username":"root","password":"root"}

**评价指标：** **（先主要集中精力完成第一个，第二个需要在完成的基础上改进）**

1、提取的敏感内容条目不少于预设总条目的80%

2、提取条目相同时，时间越少越好。

答疑邮箱： xuxiaoqiang6@huawei.com

