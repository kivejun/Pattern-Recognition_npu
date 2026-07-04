# Pattern-Recognition_npu

模式识别课程作业仓库。原始实验文档和初始数据集位于 `experiment/`，按课程数据集说明补充后的数据集位于 `datasets/`。

## Dataset Index

| 编号 | 数据集 | 仓库位置 | 状态 |
| --- | --- | --- | --- |
| 1 | 男女身高体重数据集 | `experiment/男女data数据集/data/` | 原始数据已提供 |
| 2 | 模拟多维正态数据 | `datasets/02_simulated_gaussian/` | 已生成 2 类和 3 类数据 |
| 3 | UCI Iris 数据集 | `datasets/03_uci_iris/` | 已下载并划分 |
| 4 | MNIST 手写数字数据集 | `datasets/04_mnist/` | 已下载原始 IDX gzip 文件 |
| 5 | CIFAR-10 / CIFAR-100 | `datasets/05_cifar/` | 已下载并解压 |
| 6 | 飞机分类数据集 | `experiment/plane_dataset_4_1/` | 原始数据已提供，清单见 `datasets/07_aircraft/manifest.csv` |

## Preparation

### Python Environment

推荐使用仓库内虚拟环境：

```powershell
.\.venv\Scripts\activate
```

如果需要重新创建环境：

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\activate
python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn --upgrade pip setuptools wheel
python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn -r requirements.txt
```

启动 Jupyter：

```powershell
jupyter lab
```

数据准备脚本：

```powershell
python scripts\prepare_datasets.py
```

小型公开数据集下载：

```powershell
python scripts\prepare_datasets.py --download-small
```

CIFAR 下载：

```powershell
python scripts\prepare_datasets.py --download-cifar
```
