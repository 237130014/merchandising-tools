# 商品部工具

商品部日常**数据核对**与**信息抽查**工具的集合仓库（公开）。

每个工具是一个独立的 WorkBuddy 技能包，放在自己同名目录下，可直接复制到 `~/.workbuddy/skills/` 安装使用。

---

## 收录工具

### 1. `jackyun-goods-field-audit` · 吉客云货品档案字段抽查 · v1.0

对吉客云「货品档案导出」表做**字段级信息抽查**，按 8 条业务规则筛出**疑点清单交人工复核**（基调：只筛疑点、不自动判错），并把「核查结果 / 问题摘要」两列写回被核查表，方便逐行核对。

**跑法**

```bash
PY=python   # 任意已装 openpyxl 的 Python 3 即可

# 1) 抽查 → 输出疑点清单（6 个 sheet）
$PY jackyun-goods-field-audit/scripts/audit.py --src "10-15-2货品导出.xlsx"

# 2) 写回核查列（自动复用已有的「核查结果 / 问题摘要」列）
$PY jackyun-goods-field-audit/scripts/write_back.py \
    --src "10-15-2货品导出.xlsx" \
    --issues "10-15-2_抽查疑点清单_20261016.xlsx"
```

**8 条规则**

| # | 规则 | 一句话口径 |
|---|---|---|
| 1 | 属性/品牌四列一致 | 条码属性 = 编号前 2 位 = 分类；条码品牌标记 ↔ 编号末位 ↔ 规格【品牌】。数字段一律不校；★唯一 ≡ TY（末位 A/T 互通） |
| 2 | 成分 ↔ 银重/工费 | 计价 → 银重、工费必须纯数字；一口价 → 豁免 |
| 3 | 货品名称证书措辞 | 品牌组（中央/老凤祥/珍尚银/中国珠宝/萃华）须含「配国检证书」；非品牌组（唯一/IU/TY/周大金）按材质判「配证书」 |
| 4 | 银含量三列统一 | 规格 / 货品名称 / 字印要求 等级一致；三列均未写 = 不含银 → 免检 |
| 5 | 产品名一致 | 规格【品牌】后首个括号前的产品名须出现在货品名称中 |
| 6 | 字印要求前缀 | 老凤祥→老鳳祥｜中国珠宝→中珠｜珍尚银→中国黄金｜周大金→周大金｜唯一/IU/TY→FY/MY/ZZ/足银｜萃华→CH｜中央→只有银含量 |
| 7 | 抽检比例非空 | `抽检比例(规)` 为空即疑点 |
| 8 | 手镯银重取中间值 | 规格 `11-13g` → 银重写 `12`；规格 `12g` → 写 `12` |

完整口径、品牌/属性对照表、人工豁免口径 → `jackyun-goods-field-audit/references/rules.md`

---

### 2. `dewu-hot-track-report` · 得物商机推荐趋势报告 · v1.0

抓取得物后台「商机推荐 › 高搜词 › 查看更多类目热点」的**全量热点明细**，清洗后出一份八章《得物商机推荐市场趋势分析报告》（Markdown + 网页两个版本），并带上相似机会品竞品库。

**怎么抓**：用 bsk 接管**已登录的真实浏览器**（后台必须登录态，无头浏览器进不去），打开「查看更多类目热点」抽屉 → 多选类目 → 翻页时拦截 `getHotTrackPage` 接口。**接口本身返回全部明细**，所以不用逐条点开卡片，161 条数据一轮翻页就拿完。

**跑法**

```bash
# 1) 抓取（先接管标签页，见 SKILL.md 第 1 步）
$PY dewu-hot-track-report/scripts/scrape.py --session dewu --out ./dewu-runs/2026-10-20 --groups A,B --max-pages 10

# 2) 出报告（CSV + 清洗 + 下载热点图 + MD/HTML）
$PY dewu-hot-track-report/scripts/build_report.py --data-dir ./dewu-runs/2026-10-20/data --base ./dewu-runs
```

**默认两组类目**

| 组 | 范围 |
|---|---|
| A | 流行饰品/眼镜/服配 下的 首饰、玛瑙/玉髓、水晶、黑曜石、朱砂、其他玉石、石英岩玉 |
| B | 黄金/珠宝/翡翠/玉石 整个一级类目 |

换个范围就写一份 config 传 `--config`（格式见 `references/groups.example.json`），`mode` 取 `top`（勾一级）或 `children`（勾指定二级）。

**清洗口径**（业务方确认过，别擅自调松）

| 步骤 | 规则 |
|---|---|
| 1 类目白名单 | 50 条类目，命中即留 |
| 2 国际大牌 | 30 个品牌中英文别名命中即剔（避免"和大牌同款"类伪机会） |
| 3 热度阈值 | 搜索热度 > 全量均值保留；再按成交热度均值分高/低成交两组 |

第 3 步会砍掉 80% 以上——热度指数极度右偏（中位数远低于均值），这是预期结果。

**前置条件**：bsk（`bsk doctor` 自检）+ 浏览器已登录得物后台 + 一个停在商机推荐›高搜词 的标签页 + `pandas`、`jieba`。

**两个量纲别搞错**：`searchHeat` 是裸数字指数（6650811 = 665 万，不要再乘 10000）；`searchHeatQoq` 单位是 0.01%（2187 = +21.87%）。

页面操作细节与 7 个坑位 → `dewu-hot-track-report/references/page-flow.md`；字段映射与缺字段覆盖率 → `references/field-schema.md`

---

## 安装（同事看这里）

### 方式一 · 一句话自动安装（推荐）

**不需要手动下载、解压、拷文件。** 打开 WorkBuddy 新建对话，把下面这段整句发过去，它会自己装完：

```
帮我装 WorkBuddy 技能：下载 https://github.com/237130014/merchandising-tools/archive/refs/heads/main.zip 到临时目录并解压，把里面的 jackyun-goods-field-audit 和 dewu-hot-track-report 两个文件夹都装到 C:\Users\%USERNAME%\.workbuddy\skills\ 下（如果已有同名文件夹，先改名备份成 .bak-时间戳），装好后逐个确认 skills\<技能名>\SKILL.md 存在、并确认 python 能 import openpyxl 以及 pandas、jieba，最后告诉我哪些装成功了、哪些失败、要不要重启 WorkBuddy
```

装完**完全退出 WorkBuddy 再重开**（托盘图标也要退），新建对话拖入货品表，发：

```
用吉客云货品档案抽查技能检查这个表
```

> 为什么不用 `irm ... | iex` 那种一行安装：WorkBuddy 的安全策略会拦截 `iex`（执行任意远程代码）。上面这句走的是「下载 + 解压 + 复制文件」，不执行任何下载来的代码，能力上等价且能通过校验。

### 方式二 · 自己在终端跑脚本

仓库根目录有 `install.ps1`，**一次装完仓库里的所有技能**（下载、落位、校验层级、补装依赖都由它完成）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/237130014/merchandising-tools/main/install.ps1 | iex"
```

只想装其中一个，把脚本下载到本地后加 `-Skill`：

```powershell
powershell -File install.ps1 -Skill dewu-hot-track-report
```

可用环境变量 `MT_SKILLS_DIR` 指定安装目录（默认 `~/.workbuddy/skills`）。装过的旧版本会先备份成 `<技能名>.bak-时间戳`，不会直接覆盖。

### 方式三 · 手动下载

**1. 下载**：本页右上角绿色 `Code` → `Download ZIP`。

**2. 解压**，把需要的技能文件夹（`jackyun-goods-field-audit`、`dewu-hot-track-report`）**整个**拷到：

```
C:\Users\<用户名>\.workbuddy\skills\
```

放好后层级必须是 `skills\<技能名>\SKILL.md`（点进去能**直接**看到 `SKILL.md`，不能多套一层同名文件夹）。`.workbuddy` 是隐藏目录，需在资源管理器「查看」里勾选「隐藏的项目」；`skills` 不存在就手动新建。

**3. 完全退出 WorkBuddy 再重开**（托盘图标也要退），新建对话后发对应的一句：

```
吉客云：用吉客云货品档案抽查技能检查这个表
得物  ：跑一次得物商机推荐报告
```

详细图文步骤 + 排错见 **[安装说明_同事版.md](安装说明_同事版.md)**。

> 两个技能都**不含任何写死的本机路径**，可直接在他人机器上运行；依赖缺失时脚本会自己说明缺什么、怎么装。

---

## 目录约定

```
<工具名>/
├── SKILL.md               触发条件 + 工作流 + 规则速览 + 硬性约定 + 已知坑
├── workbuddy.json         侧车元数据（display_name / 一句话简介）
├── references/*.md        完整业务口径与字段说明（改规则先改这里）
└── scripts/*.py           可执行脚本
```

新增工具时做三件事，少一件都会让同事安装失败：

1. 在 README「收录工具」下补一节（怎么跑 + 口径 + 坑）
2. 在仓库根目录建同名文件夹，把技能包整个拷进去（**不要**包含 `__pycache__`、真实数据、本机绝对路径）
3. 在 `install.ps1` 的 `$Catalog` 里加一条（`Name` / `Title` / `Deps` / `RunHint`），否则一键安装装不到它
