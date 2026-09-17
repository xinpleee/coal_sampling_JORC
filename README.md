# 煤岩地质编录自动取样

Python 3.10+ / tkinter / openpyxl implementation of the V1.0 rules in the specification.

## 使用

```bash
pip install openpyxl
python app.py
```

程序以 `data_only=False` 与 `data_only=True` 两种模式读取工作簿：公式存在缓存值时使用缓存值，没有缓存值的记录进入运行记录。原始 Excel 始终只读，结果写入用户另选的新文件。

核心逻辑在 `coal_sampling.py` 的 `SamplingEngine`，可直接用于批量处理：

```python
from coal_sampling import SamplingEngine
engine = SamplingEngine()
engine.load_xlsx("input.xlsx")
print(engine.result_rows())
```

程序输出自动取样结果、未取样煤层复核、特殊岩性未取样复核、原始编录和运行记录五个工作表。所有深度及厚度业务判断使用 `Decimal`；显示时才转换为 Excel 数值。

## 说明

- 支持 `Litho【煤岩编录】`、包含 `Litho` 的工作表和 `LOG`，多候选表会停止并提示。
- 无效深度、空岩性、深度缺口、重叠和厚度不一致不会静默丢弃。
- 默认特殊岩性和普通岩性代码可在参数窗口调整。
- 当前选择状态保留在 `SamplingEngine.selected_special` 与 `selected_thin_coal`，后续界面选择应更新这两个集合后调用完整 `calculate()`，而不是追加修改旧成果。
