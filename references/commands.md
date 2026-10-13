# 使用命令

直接用自然语言，无需技能名称前缀。新对话建议说“查看PPT风格库”而不是孤立的“查看”，便于确定对象。以下是聊天用法，不是系统注册斜杠命令。

```text
查看风格
浏览库
查看组件 特征
根据附件Word制作，风格=academic-blue，用途=defense，先给大纲
根据附件论文制作，风格=academic-minimal，用途=paper-reading，完成全部阶段
继续项目 <目录>，生成全部页面设计图
继续项目 <目录>，通过API将选定图片忠实重建为可编辑PPT，不用简单页面或组件替代原图
添加风格：来源=<PPT文件>，ID=my-style，名称=我的汇报风格，选取封面、架构、结果、总结页
添加组件：来源=<PPT文件>，第8页，选取特征堆叠，ID=feature-stack-v2
```

添加时没有对象名，代理先查看页面和选择窗格，不把来源所有对象全部收录。ID使用小写英文、数字和连字符。

代理可执行：`python <skill目录>/scripts/library.py list-styles`、`list-components --query 特征`、`browse --out <工作区>/素材浏览.html`、`new-project --out <项目> --style academic-blue --purpose defense`、`check`。用户不必手动运行终端。

查看风格输出ID、中文名、参考图片数量和资源类型。只有文档时仍可用库内图片风格，不必重传其模板。字体、字号、学校、Logo由项目决定。
