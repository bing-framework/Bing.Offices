# 自定义文件提交服务

`IFileExportCommitter` 位于 Abstractions，`DefaultFileExportCommitter` 位于 Core。官方 Provider 和外部组件均通过这些公开契约复用同目录临时文件提交，无需 `InternalsVisibleTo`。生产程序集的友元仅限仓库明确列出的测试程序集。

NPOI、ClosedXML 的正常文件导出和导入失败工作簿的 `DestinationPath` 输出使用该服务。无导入错误、失败工作簿未启用或输出到调用方 `Stream` 时，不调用文件提交器。同步入口调用 `Commit`，异步入口调用 `CommitAsync`。

## 注册与直接构造

在 Provider 注册前注册自定义实现，默认注册不会覆盖它：

```csharp
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

services.AddSingleton<IFileExportCommitter, HostFileCommitter>();
services.AddBingOfficesNpoi();
```

ClosedXML 使用相同注册方式并调用 `AddBingOfficesClosedXml`；其 DOM 准入配置继续生效。

直接构造时保留旧五参数重载，新增六参数重载的所有参数均必传：

```csharp
var importer = new Bing.Offices.Imports.NpoiExcelImporter(
    null, null, null, null, null, new HostFileCommitter());
var closedXmlImporter = new Bing.Offices.ClosedXml.Imports.ClosedXmlExcelImporter(
    null, null, null, null, null, new HostFileCommitter());
```

最后一项为 `null` 时使用默认提交器。以下装饰器保留默认文件安全行为，宿主可在服务层增加自己的逻辑：

```csharp
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.IO;

public sealed class HostFileCommitter : IFileExportCommitter
{
    private readonly DefaultFileExportCommitter _inner = new();

    public void Commit(string path, Action<Stream> write,
        CancellationToken cancellationToken, string format)
        => _inner.Commit(path, write, cancellationToken, format);

    public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
        CancellationToken cancellationToken, string format)
        => _inner.CommitAsync(path, writeAsync, cancellationToken, format);
}
```

## 实现契约

- 写委托仅调用一次。提交器拥有其创建的临时目标流；写委托不得关闭该流。调用方的输入流和直接输出流仍由调用方管理。
- 写入并刷新成功、检查取消后才能替换旧文件；异常或取消必须保留旧文件并清理临时文件。推荐装饰默认实现，避免重复实现提交算法。
- `CommitAsync` 将原令牌交给写委托并等待完成；外围 IO 使用真实异步，不用 `Task.Run`、`.Result` 或 `.Wait()`。底层文件替换和 DOM 序列化仍有同步边界。
- 内容生成失败与文件提交失败须区分；`OperationCanceledException` 保持取消语义，文件提交失败使用 `BingOfficesFileCommitException`。
- 提交器不重复通知异常观察器；公共导入/导出入口负责观察。默认提交器会保留主异常，并附带临时文件清理失败信息。
- 若按 singleton 注册，自定义实现须能安全处理并发调用，不得把每次调用的流、路径或令牌存入共享可变字段。

## 兼容性

两个导入器原五参数构造签名和默认值保持不变，新增重载不会改变现有构造调用。
移除 Core 的生产友元后，曾直接调用其内部方法的旧 NPOI / ClosedXML Provider 二进制需要与 Core 配套升级。外部组件始终使用公开接口；不要引用内部 AtomicFileCommitter。

## 后续优化建议

可独立推进第三方 Provider 契约测试套件、最小 Provider 项目模板、Provider SPI 文档索引和提交耗时诊断。本次不新增这些机制。
