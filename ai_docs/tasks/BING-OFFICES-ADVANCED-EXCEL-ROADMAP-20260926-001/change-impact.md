# Change Impact Analysis

| 项目 | 影响 |
|---|---|
| 生产项目 | Abstractions 新增 Sheet 图片/原生校验；SpreadCheetah 映射和内容写入；NPOI/ClosedXML 原生内容、报表预检与等价规则；MiniExcel 新内容预检拒绝。 |
| 测试项目 | SheetContent 独立预期及原生/OOXML 读回；流式映射与资源合同；NPOI/ClosedXML 报表直接测试；双 TFM API 与真实包消费者。 |
| 公共 API | 仅追加图片、校验类型及 Sheet/Builder 成员；保持 netstandard2.0，不删除原有成员，不改变既有枚举值。本轮未增加 Provider。 |
| 运行时路径 | 映射绑定创建和报表静态检查前移；图片请求磁盘暂存适配 JPEG/锚点；NPOI 显式列表流式修正；默认色阶/数据条与打印边界统一。 |
| 资源与所有权 | 串行背压、取消检查及公共原子文件提交；Stream 不可回滚且不关闭。无图片容量样本不能代表图片修正路径，DOM 仍有完整工作簿内存。 |
| 依赖与发布 | 不新增/升级引擎或版本；八包仅重新打包做隔离消费者验证，不发布。 |
| 容器与容量 | 固定摘要、Runtime、字体源、非 root、只读源码和 4 GiB；真实 fontless 路径及 9 组容量样本。远程 CI 未执行，正式阈值不在本轮批准。 |

风险等级：高（新增公共契约、ZIP 图片/校验适配、跨 Provider 报表行为）。先职责测试，再双 TFM/全量门禁、容器及独立审查。Aspose、公式重新计算、高级 Custom 校验、透视表均为 Deferred，不修改实现。
