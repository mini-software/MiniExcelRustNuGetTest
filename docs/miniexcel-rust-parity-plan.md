# MiniExcel 全量 Rust 后端比对与迁移计划

## 1. 目标

在不修改 `D:\git\MiniExcel` 的前提下，以它的公开 API、单元测试和实际行为作为只读基准，补齐 `D:\git\MiniExcel-Rust` 的能力，并在当前 `D:\git\MiniExcelRust` .NET 包装仓库中提供兼容入口，使 MiniExcel 的 XLSX/CSV 解析、写入、模板和工作簿操作最终全部由 Rust 执行。

完成后的生产包不得依赖或回退到 C# MiniExcel 实现。反射、`DataTable`、`IDataReader`、`IAsyncEnumerable<T>` 和 .NET 异常转换可以留在薄托管适配层，但文件格式处理、工作簿变更和模板执行必须进入 Rust。

## 2. 仓库边界

| 路径 | 角色 | 是否允许修改 |
| --- | --- | --- |
| `D:\git\MiniExcel` | C# API、行为和测试的基准实现 | 否，只读、可编译、可运行测试 |
| `D:\git\MiniExcel-Rust` | Rust 核心引擎 | 是，补齐底层能力和 Rust 测试 |
| `D:\git\MiniExcelRust` | 当前 .NET NuGet、FFI 和跨平台验证仓库 | 是，先在这里建立兼容层、比对测试和发布验证 |

实施期间不得改写、格式化或提交 `D:\git\MiniExcel` 中的任何文件。需要的新 fixture、契约文件和差异测试应放入两个 Rust 相关仓库。

## 3. 完成定义

以下条件必须同时满足，才能称为“全部 MiniExcel 方法底层换成 Rust”：

1. 公开 API 清单中的每个方法和重载都有对应实现，包含同步、异步、路径、流和 `byte[]` 变体。
2. 方法名称、泛型约束、参数名称、默认值、返回类型和可观察异常与基准版本兼容。
3. 生产依赖图中不存在 C# MiniExcel 包或程序集；差异测试项目可以仅把它作为基准 oracle 使用。
4. XLSX/CSV 的读取、写入、模板、图片和工作簿变更均由 Rust 引擎执行。
5. `D:\git\MiniExcel` 的适用单元测试已在只读目录原样通过；迁移后的兼容测试与差异测试也全部通过。
6. Windows、Linux、macOS 的 x64/Arm64 包测试通过；musl 平台继续通过正确性和资源生命周期测试。
7. 提前停止枚举、取消、异常、重复调用和流所有权测试证明没有句柄、文件描述符或非托管内存持续增长。

基准版本应在执行开始时记录提交 SHA、包版本和公开 API 快照。基准升级必须单独评审，不能在迁移过程中无提示漂移。

## 4. API 比对清单

先生成机器可读矩阵，建议字段为：`API ID`、基准签名、Rust 能力、FFI 能力、托管入口、同步测试、异步测试、路径测试、流测试、异常测试、状态和备注。

### 4.1 OpenXML 读取

- [ ] `Query` / `QueryAsync`：动态、泛型，路径和流。
- [ ] `QueryRange` / `QueryRangeAsync`：A1 地址与行列索引重载。
- [ ] `QueryTable` / `QueryTableAsync`：动态、泛型，路径和流。
- [ ] `QueryAsDataTable` / `QueryAsDataTableAsync`。
- [ ] `GetReader` / `GetDataReader` / `GetAsyncDataReader`。
- [ ] `GetSheetNames`、`GetSheetInformations`、`GetSheetDimensions`、`GetColumns` / `GetColumnNames`。
- [ ] `RetrieveComments`，包括批注、作者和回复。
- [ ] 延迟枚举、取消、空行、稀疏单元格、合并单元格填充、表头裁剪和共享字符串缓存。

### 4.2 OpenXML 写入与工作簿操作

- [ ] `SaveAs` / `Export`：路径和流，同步和异步。
- [ ] POCO、匿名对象、字典、`DataTable`、`DataSet`、`IDataReader`、`IAsyncEnumerable<T>` 和多工作表输入。
- [ ] `Insert` / `InsertSheet`：新增或替换工作表。
- [ ] `CopyAndAddSheet`：文件到文件、流到流。
- [ ] `AlterSheet`：重命名、排序和可见状态。
- [ ] 自动筛选、冻结窗格、RTL、列宽、隐藏列、换行、对齐、表头样式、日期及数字格式。
- [ ] 覆盖策略、进度回报、原子输出和失败后的目标文件状态。

### 4.3 模板与富内容

- [ ] `SaveAsByTemplate` / `FillTemplate`：所有路径、流和 `byte[]` 组合。
- [ ] 标量替换、集合展开、分组、条件、公式、行移动和缺失变量策略。
- [ ] `MergeSameCells`：合并标记、边界和既有合并区域。
- [ ] `AddPicture`：图片类型、尺寸、锚点、关系和内容类型。
- [ ] 模板处理后保持公式、合并区域、表格、批注、定义名称和绘图关系有效。

### 4.4 CSV 与格式转换

- [ ] CSV `Query`、`QueryAsDataTable`、`GetColumnNames` 和 Reader API。
- [ ] CSV `Export` / `SaveAs` 与 `Append`。
- [ ] 分隔符、换行符、引号、嵌入换行、BOM、空字符串、编码和自定义 reader/writer。
- [ ] `ConvertCsvToXlsx` 和 `ConvertXlsxToCsv` 的路径与流、同步与异步重载。

### 4.5 映射、配置和兼容入口

- [ ] `MiniExcelLibs.MiniExcel` 旧版 facade 的全部方法与重载。
- [ ] V2 `Importers`、`Exporters`、`Templaters` provider API。
- [ ] 列名称、索引、宽度、格式、隐藏、忽略、sheet 等 attributes。
- [ ] nullable、enum、GUID、URI、日期、`DateOnly`、`DateTimeOffset`、`TimeSpan`、culture 和自定义格式转换。
- [ ] Fluent Mapping：`Property`、`Collection`、`ToWorksheet`、`ToCell`、`WithFormat`、`WithFormula`、`StartAt`、`WithSpacing` 和嵌套集合。
- [ ] `ExcelType`、configuration、model、enum、exception 和 compatibility alias。
- [ ] stream `leaveOpen`、seek 要求、overwrite 默认值和参数验证行为。

## 5. 当前差距摘要

| 能力 | `MiniExcel-Rust` 核心 | 当前 .NET FFI/包装 | 主要工作 |
| --- | --- | --- | --- |
| 动态 XLSX 查询 | 已有较完整能力 | 仅同步路径查询 | 补齐 options、范围、流、异步和错误语义 |
| 泛型映射 | Rust Serde 已有基础 | 未暴露 | 建立 schema/列映射协议和 .NET 转换层 |
| metadata/table/comments | Rust 已有 | 未暴露 | 增加 ABI 和托管模型 |
| CSV | Rust 已有读写与 append | 未暴露 | 增加流式 ABI、配置和转换入口 |
| XLSX 写入 | Rust 已有基础与多 sheet 能力 | 未暴露 | 增加 schema、输入回调、进度和原子输出 |
| insert/copy/rename/reorder/visibility | Rust 已有部分能力 | 未暴露 | 增加 package 保真与回滚测试 |
| template/merge | Rust 已有基础 | 未暴露且与 C# 仍有差距 | 补齐集合、关系和公式更新语义 |
| picture | 未完整支持 | 未暴露 | 实现 OOXML drawing、media 和 relationship 写入 |
| `DataTable`/`IDataReader` | 不属于 Rust 类型系统 | 未实现 | 在托管层适配到统一 Rust row/schema 协议 |
| Fluent Mapping | 有 cell map，但不等价 | 未实现 | 托管层生成 mapping plan，Rust 执行读取/写入 |
| async/cancellation | 原生 async 有部分能力 | 未暴露 | 增加取消句柄、异步流和线程规则 |

当前 ABI v1 还会把 Rust `i64` 转成 `double`、把 Excel error 转成普通字符串，并把 duration 截断到毫秒；这些都必须在兼容工作开始前修正。

## 6. 实施阶段

### 阶段 0：冻结基准与建立矩阵

1. 记录三个仓库的 commit SHA、工具链版本和目标框架。
2. 用反射生成 `D:\git\MiniExcel` 公开 API 快照，包括生成出的同步方法。
3. 从 OpenXML、CSV、Fluent Mapping、legacy facade 测试中建立测试对应表。
4. 将必要 fixture 复制到 Rust 相关仓库，记录来源与预期 hash；不修改原 fixture。
5. 建立“缺失、签名不符、行为不符、已通过”四种状态的 API 矩阵。

退出条件：所有公开方法都有唯一 API ID 和至少一个计划中的验收测试。

### 阶段 1：稳定 ABI 与资源模型

1. 发布版本化 C header/protocol，定义长度、所有权、线程、取消和错误规则。
2. 把返回值升级为可扩展 tagged value，保留 `Int64`、日期时间精度、Excel error 和 null/empty 差异。
3. 为 path、borrowed stream callback、owned buffer、row iterator 和 writer 建立独立句柄。
4. 提供结构化 error code、错误类别、参数名和内部消息，不允许 panic 跨越 FFI。
5. 增加 ABI test vectors、畸形 frame 测试、重复 close 和提前 dispose 测试。
6. 消除当前仓库与上游重复 FFI 源码漂移：改为单一来源或加入自动同步校验。

退出条件：ABI 契约测试、内存/句柄生命周期测试和八个 RID 的加载测试通过。

### 阶段 2：完成 XLSX 读取

1. 补齐 `ReadOptions`、end cell、空行、merged fill、trim header、缓存与 sheet 选择。
2. 暴露 range、table、metadata、column、dimension 和 comments API。
3. 实现 path、stream、`byte[]` 统一读取源；严格落实 `leaveOpen`。
4. 完成 dynamic row 的列名、顺序、值类型和异常一致性。
5. 增加 `IEnumerable` / `IAsyncEnumerable`、取消和提前结束枚举。

退出条件：读取类差异测试逐行、逐列、逐类型一致，相关 OpenXML 基准测试全部通过。

### 阶段 3：类型映射、配置与 Reader

1. 托管层把 attributes、reflection 和 fluent mapping 编译为稳定 schema/mapping plan。
2. Rust 按 plan 完成列定位、值转换和错误定位；托管层只构造对象或适配 .NET 类型。
3. 在同一 native row iterator 上实现 `IDataReader`、async reader 和 `DataTable` adapter。
4. 补齐 culture、日期、enum、nullable、字段、动态列和映射异常。

退出条件：typed mapping、DataReader、DataTable、Fluent Mapping 的读取测试全部通过。

### 阶段 4：CSV 与转换

1. 暴露 CSV query、reader、export、append 和完整配置。
2. 覆盖 UTF-8/UTF-16/GBK/Windows-1252、BOM、quote 和跨行字段。
3. 让 CSV/XLSX 转换直接串接 Rust reader/writer，避免整表进入托管内存。

退出条件：CSV 单元测试、CsvHelper 互操作测试和转换 round-trip 测试通过。

### 阶段 5：XLSX 写入

1. 先实现 dynamic/schema 单 sheet，再支持 typed、异步输入和多 sheet。
2. 托管输入统一转换为 row/schema callback，覆盖所有支持的数据源类型。
3. 补齐 style、format、width、hidden、freeze、filter、RTL 和 progress。
4. 对 path 使用临时文件加原子替换；stream 失败时定义并测试可观察状态。
5. 使用 Excel、EPPlus、ClosedXML、NPOI、ExcelDataReader 或 Packaging 验证输出。

退出条件：写入结果可被基准读取且关键 OOXML 结构等价，所有 export 测试通过。

### 阶段 6：工作簿变更、模板与图片

1. 完成 insert、copy/add、rename、reorder 和 visibility，并保持无关 package parts 不变。
2. 补齐模板集合、分组、条件、公式、行偏移、merge 和 missing-value 行为。
3. 实现 picture 写入及 drawing relationship/content type 管理。
4. 对公式引用、calc chain、table range、defined names、merge、comments 和 drawings 建立变更后校验。
5. `.xlsm` 按基准行为拒绝可能丢失宏的写入操作，不静默降级。

退出条件：模板、图片、工作簿变更和第三方互操作测试全部通过。

### 阶段 7：完整 facade 与发布门禁

1. 将 provider API 和 legacy `MiniExcelLibs.MiniExcel` 的全部重载接到 Rust-backed 实现。
2. 用 API snapshot/approval test 阻止遗漏重载、默认值或 public type。
3. 删除任何生产环境 C# MiniExcel fallback 和临时双实现开关。
4. 执行全量单元、差异、互操作、压力、泄漏和跨平台 package 测试。
5. 对每个 API ID 关闭矩阵条目；不得以“Rust 暂不支持”跳过完成门槛。

退出条件：API 矩阵 100% 完成，生产依赖检查为纯 Rust 后端，发布包全平台验证通过。

## 7. 测试策略

### 7.1 三层测试

1. Rust 单元/集成测试：验证 parser、writer、template、package mutation 和错误路径。
2. .NET 兼容测试：验证签名、默认值、attributes、reflection、Reader、DataTable、异步和异常。
3. 黑盒差异测试：同一 fixture、参数和 culture 分别运行 C# 基准与 Rust-backed 包，对结果、异常和输出文件进行标准化比对。

不要只比较成功结果。每个 API 至少覆盖正常、边界、错误、取消或提前结束中的适用场景。

### 7.2 比对规则

- 查询：比较 sheet、row、column、key 顺序、CLR 类型和值；浮点、日期和 duration 使用明确规则。
- 异常：比较异常类别、触发时机、参数名和关键消息，不依赖平台路径文本。
- XLSX：先比较语义，再检查关键 OOXML parts、relationships、content types 和未变更 part 的 hash。
- CSV：比较编码后的 bytes、BOM、换行、引用和尾部换行。
- 流：覆盖 seekable/non-seekable、只读/只写、`leaveOpen` 和中途异常。
- 性能：正确性优先；通过后要求流式操作保持有界内存，且不得比 C# 基准出现未解释的数量级退化。

### 7.3 必须纳入的回归类别

- Header/headerless、Unicode、空白与重复表头、稀疏 XML、缺失 `r` attribute。
- 数字精度、bool、null/empty、日期、时间、时长、公式 cached value 和 Excel error。
- 多 sheet、隐藏 sheet、table、comments/replies、merged cells 和 shared strings。
- 泛型映射、attributes、culture、nullable、enum、GUID、URI 和 mapping failure。
- 大文件、提前停止、取消、并发、重复 5,000 次生命周期和目标文件独占重开。
- 模板分组、公式、图片、insert/copy/rename/reorder 及第三方软件打开验证。
- `D:\git\MiniExcel\tests` 下 `Issues` 目录中的历史回归案例。

## 8. 建议验证命令

只读基准：

```powershell
dotnet test D:\git\MiniExcel\tests\MiniExcel.OpenXml.Tests\MiniExcel.OpenXml.Tests.csproj --framework net10.0
dotnet test D:\git\MiniExcel\tests\MiniExcel.Csv.Tests\MiniExcel.Csv.Tests.csproj --framework net10.0
```

Rust 核心：

```powershell
Set-Location D:\git\MiniExcel-Rust
cargo +1.85.0 fmt --all -- --check
cargo +1.85.0 clippy --workspace --all-targets --all-features --locked -- -D warnings
cargo +1.85.0 test --workspace --all-targets --all-features --locked
cargo +1.85.0 doc --workspace --no-deps --all-features --locked
```

当前 .NET 包装与 NuGet：

```powershell
Set-Location D:\git\MiniExcelRust
cargo test --workspace --all-targets --locked
dotnet build .\src\MiniExcelRust\MiniExcelRust.csproj -c Release
.\build\Test-Package.ps1 -Rid win-x64
```

CI 中再扩展到所有目标 RID，并把 API snapshot、差异测试、依赖检查和 package 内容校验设为必过门禁。

## 9. 执行顺序与交付物

每个阶段使用同一节奏：先补失败的契约测试，再实现 Rust 核心，再扩展 ABI 和托管入口，最后跑差异与跨平台 package 测试。避免先批量声明全部 .NET 方法再长期保留 `NotSupportedException`。

阶段性交付物如下：

- `api-baseline.json`：C# 基准公开 API 快照。
- `parity-matrix.md` 或结构化等价文件：逐 API 状态和测试证据。
- 版本化 ABI 文档与 test vectors。
- 可复用的差异测试 runner 和标准化比较器。
- 各阶段 Rust、.NET、互操作和资源测试报告。
- 最终生产依赖报告，证明没有 C# MiniExcel runtime fallback。

首个实现批次应从阶段 0、阶段 1 和 XLSX 动态读取差距开始，不应直接进入模板或图片功能；先稳定 ABI，后续方法才能共用同一套流、错误、值和生命周期协议。