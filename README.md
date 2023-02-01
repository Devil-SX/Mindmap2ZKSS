# Mindmap AutoCase for ZKSS
  <h1 align="center">
    <picture>
      <img alt="ZKSS" src="./pic/zkss.png" width=367.5>
    </picture>
  </h1>

# ✨功能
- 使用Mindmap编写测试用例
- 将Mindmap文件转化为中科山水测试用例格式Excel
- 自动生成测试ID

# 💡过程
中科山水（北京）科技有限公司采用Excel编写记录测试用例，表格包含以下表头：
1.  模块名称
2. \[子模块\]
3. 角色
4. 功能
5. 功能说明
6. 测试用例ID
7. 测试用例
8. 测试步骤
	1. 步骤 or 前提
	2. 业务操作
	3. 预期结果
9. 测试日期
10. 所用数据/正常测试
11. 所用数据/异常测试
12. 执行结果

其中`1~8`属于测试用例，`9~12`属于测试过程与结果

Mindmap适用编辑树状结构。除去`3.角色`和`6.测试ID`，`1~8`剩余表头用Mindmap编辑

Mindmap编辑完成后，使用Excel VBA脚本自动生成`3 6 9 10 11 12`项表头和测试id，并调整格式。最后手动设置角色

*暂不生成`2.子模块`表头，`8.1 步骤 or 前提`根据表格内容猜测自动生成步骤或前提，默认为前提*

# 📘使用

查看[使用文档](./doc/README.md)

# 🪄待开发功能
- 自动生成角色项
- 集成自动化Mindmap导出Excel，Excel执行宏
- 和Testlink等测试管理工具管理测试用例，[可参考该仓库实现](https://github.com/zhuifengshen/xmind2testcase)