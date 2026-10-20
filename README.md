# 商品部工具

商品部日常**货品表格**处理工具的集合仓库（公开）—— 数据核对、字段抽查、批量回写这类「拿一张表进去、拿一张表出来」的活。

每个工具是一个独立的 WorkBuddy 技能包，放在自己同名目录下，可直接复制到 `~/.workbuddy/skills/` 安装使用。

> **范围说明**：本仓库只收**货品表格类**工具。抓取类（浏览器自动化取数）、内容生成类、热点分析类工具统一放在
> [ecommerce-portfolio](https://github.com/237130014/ecommerce-portfolio)（电商工具仓库）。

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

## 安装（同事看这里）

### 方式一 · 一句话自动安装（推荐）

**不需要手动下载、解压、拷文件。** 打开 WorkBuddy 新建对话，把下面这段整句发过去，它会自己装完：

```
帮我装一个 WorkBuddy 技能：下载 https://github.com/237130014/merchandising-tools/archive/refs/heads/main.zip 到临时目录并解压，把里面的 jackyun-goods-field-audit 整个文件夹装到 C:\Users\%USERNAME%\.workbuddy\skills\ 下（如果已有同名文件夹，先改名备份成 .bak-时间戳），装好后确认 skills\jackyun-goods-field-audit\SKILL.md 这个文件存在、并确认 python 能 import openpyxl，最后告诉我装没装成功、要不要重启 WorkBuddy
```

装完**完全退出 WorkBuddy 再重开**（托盘图标也要退），新建对话拖入货品表，发：

```
用吉客云货品档案抽查技能检查这个表
```

> 为什么不用 `irm ... | iex` 那种一行安装：WorkBuddy 的安全策略会拦截 `iex`（执行任意远程代码）。上面这句走的是「下载 + 解压 + 复制文件」，不执行任何下载来的代码，能力上等价且能通过校验。

### 方式二 · 自己在终端跑脚本

仓库根目录有 `install.ps1`，走到位、校验层级、补装 `openpyxl` 都由它完成：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/237130014/merchandising-tools/main/install.ps1 | iex"
```

或下载后本地执行：`powershell -File install.ps1`。可用环境变量 `MT_SKILLS_DIR` 指定安装目录（默认 `~/.workbuddy/skills`）。

### 方式三 · 手动下载

**1. 下载**：本页右上角绿色 `Code` → `Download ZIP`。

**2. 解压**，把里面的 `jackyun-goods-field-audit` **整个文件夹**拷到：

```
C:\Users\<用户名>\.workbuddy\skills\
```

放好后层级必须是 `skills\jackyun-goods-field-audit\SKILL.md`（点进去能**直接**看到 `SKILL.md`，不能多套一层同名文件夹）。`.workbuddy` 是隐藏目录，需在资源管理器「查看」里勾选「隐藏的项目」；`skills` 不存在就手动新建。

**3. 完全退出 WorkBuddy 再重开**（托盘图标也要退），新建对话拖入货品表，发：

```
用吉客云货品档案抽查技能检查这个表
```

详细图文步骤 + 排错见 **[安装说明_同事版.md](安装说明_同事版.md)**。

**1. 下载**：本页右上角绿色 `Code` → `Download ZIP`。

**2. 解压**，把里面的 `jackyun-goods-field-audit` **整个文件夹**拷到：

```
C:\Users\<用户名>\.workbuddy\skills\
```

放好后层级必须是 `skills\jackyun-goods-field-audit\SKILL.md`（点进去能**直接**看到 `SKILL.md`，不能多套一层同名文件夹）。`.workbuddy` 是隐藏目录，需在资源管理器「查看」里勾选「隐藏的项目」；`skills` 不存在就手动新建。

**3. 完全退出 WorkBuddy 再重开**（托盘图标也要退），新建对话拖入货品表，发：

```
用吉客云货品档案抽查技能检查这个表
```

详细图文步骤 + 排错见 **[安装说明_同事版.md](安装说明_同事版.md)**。

> 技能只依赖 `openpyxl`，**不含任何写死的本机路径**，可直接在他人机器上运行；首次执行时脚本会自检依赖。

---

## 目录约定

```
<工具名>/
├── SKILL.md              触发条件 + 工作流 + 规则速览 + 硬性约定 + 已知坑
├── references/rules.md   完整业务口径（改规则先改这里）
└── scripts/*.py          可执行脚本
```

新增工具时，在 README「收录工具」下补一节，并在仓库根目录建同名文件夹。
