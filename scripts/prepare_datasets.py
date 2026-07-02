"""Prepare coursework datasets for the pattern recognition assignments.

The script intentionally uses only the Python standard library so it can run in
the current lightweight environment.
"""

from __future__ import annotations

import argparse
import csv
import gzip
import hashlib
import os
import random
import shutil
import ssl
import struct
import tarfile
import urllib.request
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple


ROOT = Path(__file__).resolve().parents[1]
DATASETS = ROOT / "datasets"

RNG_SEED = 20260615

UCI_IRIS_URL = "https://archive.ics.uci.edu/ml/machine-learning-databases/iris/iris.data"
MNIST_URLS = {
    "train-images-idx3-ubyte.gz": "http://yann.lecun.com/exdb/mnist/train-images-idx3-ubyte.gz",
    "train-labels-idx1-ubyte.gz": "http://yann.lecun.com/exdb/mnist/train-labels-idx1-ubyte.gz",
    "t10k-images-idx3-ubyte.gz": "http://yann.lecun.com/exdb/mnist/t10k-images-idx3-ubyte.gz",
    "t10k-labels-idx1-ubyte.gz": "http://yann.lecun.com/exdb/mnist/t10k-labels-idx1-ubyte.gz",
}
CIFAR10_URL = "http://cave.cs.toronto.edu/kriz/cifar-10-python.tar.gz"
CIFAR100_URL = "http://cave.cs.toronto.edu/kriz/cifar-100-python.tar.gz"


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def cholesky_2x2(cov: Sequence[Sequence[float]]) -> Tuple[float, float, float]:
    a = cov[0][0]
    b = cov[1][0]
    c = cov[1][1]
    l11 = a ** 0.5
    l21 = b / l11
    l22 = (c - l21 * l21) ** 0.5
    return l11, l21, l22


def sample_gaussian_2d(mean: Sequence[float], cov: Sequence[Sequence[float]]) -> Tuple[float, float]:
    z1 = random.gauss(0.0, 1.0)
    z2 = random.gauss(0.0, 1.0)
    l11, l21, l22 = cholesky_2x2(cov)
    return mean[0] + l11 * z1, mean[1] + l21 * z1 + l22 * z2


def write_csv(path: Path, header: Sequence[str], rows: Iterable[Sequence[object]]) -> None:
    ensure_dir(path.parent)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)


def split_rows(rows: List[Sequence[object]]) -> Tuple[List[Sequence[object]], List[Sequence[object]], List[Sequence[object]]]:
    random.shuffle(rows)
    train_end = int(len(rows) * 0.60)
    val_end = int(len(rows) * 0.80)
    return rows[:train_end], rows[train_end:val_end], rows[val_end:]


def prepare_simulated() -> None:
    random.seed(RNG_SEED)
    out = DATASETS / "02_simulated_gaussian"
    ensure_dir(out)

    specs = {
        "2class": [
            ("class_0", (0.0, 0.0), ((1.6, 0.45), (0.45, 1.0))),
            ("class_1", (3.0, 2.2), ((1.2, -0.35), (-0.35, 1.5))),
        ],
        "3class": [
            ("class_0", (0.0, 0.0), ((1.2, 0.25), (0.25, 1.0))),
            ("class_1", (3.2, 0.8), ((1.0, -0.20), (-0.20, 1.1))),
            ("class_2", (1.2, 3.4), ((1.4, 0.35), (0.35, 1.2))),
        ],
    }

    manifest_rows = []
    for name, classes in specs.items():
        all_rows: List[Sequence[object]] = []
        for label, mean, cov in classes:
            for idx in range(200):
                x1, x2 = sample_gaussian_2d(mean, cov)
                all_rows.append((f"{label}_{idx:03d}", round(x1, 6), round(x2, 6), label))
        train, val, test = split_rows(all_rows)
        header = ("sample_id", "x1", "x2", "label")
        write_csv(out / name / "train.csv", header, train)
        write_csv(out / name / "val.csv", header, val)
        write_csv(out / name / "test.csv", header, test)
        write_csv(out / name / "all.csv", header, all_rows)
        manifest_rows.append((name, len(classes), len(all_rows), len(train), len(val), len(test)))

    write_csv(
        out / "manifest.csv",
        ("dataset", "num_classes", "total", "train", "val", "test"),
        manifest_rows,
    )


def download(url: str, path: Path, insecure_ssl: bool = False) -> None:
    ensure_dir(path.parent)
    if path.exists() and path.stat().st_size > 0:
        return
    context = ssl._create_unverified_context() if insecure_ssl else None
    with urllib.request.urlopen(url, context=context) as response, path.open("wb") as f:
        shutil.copyfileobj(response, f)


def prepare_uci(download_files: bool) -> None:
    out = DATASETS / "03_uci_iris"
    ensure_dir(out / "raw")
    if download_files:
        raw = out / "raw" / "iris.data"
        download(UCI_IRIS_URL, raw)
        rows = []
        with raw.open("r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                rows.append(line.split(","))
        random.seed(RNG_SEED)
        train, val, test = split_rows(rows)
        header = ("sepal_length", "sepal_width", "petal_length", "petal_width", "label")
        write_csv(out / "train.csv", header, train)
        write_csv(out / "val.csv", header, val)
        write_csv(out / "test.csv", header, test)
        write_csv(out / "all.csv", header, rows)
    write_text(
        out / "README.md",
        """# UCI Iris Dataset

Official source: https://archive.ics.uci.edu/dataset/53/iris

This folder is prepared for the UCI dataset requirement. Run:

```powershell
python scripts/prepare_datasets.py --download-small
```

The script downloads `iris.data` and writes `train.csv`, `val.csv`, and `test.csv`.
Split ratio: 60% / 20% / 20%.
""",
    )


def parse_idx_images(path: Path, limit: int = 20) -> Tuple[int, int, int]:
    with gzip.open(path, "rb") as f:
        magic, count, rows, cols = struct.unpack(">IIII", f.read(16))
        if magic != 2051:
            raise ValueError(f"Unexpected image IDX magic number in {path}: {magic}")
        f.read(rows * cols * min(limit, count))
    return count, rows, cols


def parse_idx_labels(path: Path, limit: int = 20) -> int:
    with gzip.open(path, "rb") as f:
        magic, count = struct.unpack(">II", f.read(8))
        if magic != 2049:
            raise ValueError(f"Unexpected label IDX magic number in {path}: {magic}")
        f.read(min(limit, count))
    return count


def prepare_mnist(download_files: bool) -> None:
    out = DATASETS / "04_mnist"
    raw = out / "raw"
    ensure_dir(raw)
    if download_files:
        for filename, url in MNIST_URLS.items():
            download(url, raw / filename)
        train_count, rows, cols = parse_idx_images(raw / "train-images-idx3-ubyte.gz")
        test_count, _, _ = parse_idx_images(raw / "t10k-images-idx3-ubyte.gz")
        train_labels = parse_idx_labels(raw / "train-labels-idx1-ubyte.gz")
        test_labels = parse_idx_labels(raw / "t10k-labels-idx1-ubyte.gz")
        write_csv(
            out / "manifest.csv",
            ("split", "images", "labels", "height", "width"),
            [("train", train_count, train_labels, rows, cols), ("test", test_count, test_labels, rows, cols)],
        )
    write_text(
        out / "README.md",
        """# MNIST Dataset

Course source: http://yann.lecun.com/exdb/mnist/

This folder stores the standard IDX gzip files. Run:

```powershell
python scripts/prepare_datasets.py --download-small
```

The script downloads the four standard files and writes `manifest.csv`.
The standard split is 60,000 training images and 10,000 test images. If a validation
set is needed, use the last 10,000 training samples as validation and the first
50,000 as training.
""",
    )


def md5(path: Path) -> str:
    digest = hashlib.md5()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_extract_tar(archive: Path, destination: Path) -> None:
    dest = destination.resolve()
    with tarfile.open(archive, "r:gz") as tar:
        for member in tar.getmembers():
            target = (dest / member.name).resolve()
            if dest != target and dest not in target.parents:
                raise ValueError(f"Unsafe archive member path: {member.name}")
        tar.extractall(destination)


def prepare_cifar(download_cifar: bool) -> None:
    out = DATASETS / "05_cifar"
    raw = out / "raw"
    ensure_dir(raw)
    rows = [
        ("cifar-10-python.tar.gz", CIFAR10_URL, "c58f30108f718f92721af3b95e74349a"),
        ("cifar-100-python.tar.gz", CIFAR100_URL, "eb9058c3a382ffc7106e4002c42a8d85"),
    ]
    if download_cifar:
        for filename, url, expected_md5 in rows:
            archive = raw / filename
            download(url, archive, insecure_ssl=True)
            actual_md5 = md5(archive)
            if actual_md5 != expected_md5:
                raise ValueError(f"MD5 mismatch for {filename}: {actual_md5} != {expected_md5}")
            safe_extract_tar(archive, out)
    write_csv(out / "sources.csv", ("file", "url", "md5"), rows)
    write_text(
        out / "README.md",
        """# CIFAR-10 / CIFAR-100

Official source: https://www.cs.toronto.edu/~kriz/cifar.html

Run only when you really want the archives in this repository:

```powershell
python scripts/prepare_datasets.py --download-cifar
```

The archives are about 160 MB each. CIFAR-10 contains 60,000 32x32 color images
in 10 classes. CIFAR-100 contains 60,000 images in 100 classes.
""",
    )


def prepare_imagenet() -> None:
    out = DATASETS / "06_imagenet"
    ensure_dir(out)
    write_text(
        out / "README.md",
        """# ImageNet

Official source: https://image-net.org/update-mar-11-2021.php

ImageNet is not bundled or automatically downloaded here because it requires an
ImageNet account, agreement to the dataset terms, and substantial disk space.
After downloading manually, place the files under:

```text
datasets/06_imagenet/raw/
```

For coursework experiments, prefer a small subset with explicit `train/`,
`val/`, and `test/` folders plus a `labels.csv` file.
""",
    )


def prepare_aircraft_manifest() -> None:
    source = ROOT / "experiment" / "plane_dataset_4_1"
    out = DATASETS / "07_aircraft"
    ensure_dir(out)
    rows = []
    for split in ("train", "test"):
        split_dir = source / split
        if not split_dir.exists():
            continue
        for class_dir in sorted(p for p in split_dir.iterdir() if p.is_dir()):
            count = sum(1 for p in class_dir.iterdir() if p.is_file())
            rows.append((split, class_dir.name, count, str(class_dir.relative_to(ROOT))))
    write_csv(out / "manifest.csv", ("split", "class", "count", "source_path"), rows)
    write_text(
        out / "README.md",
        """# Aircraft Classification Dataset

The aircraft dataset already exists in:

```text
experiment/plane_dataset_4_1/
```

This folder records its train/test class counts in `manifest.csv`. Use the
original image folders directly for experiments.
""",
    )


def write_text(path: Path, text: str) -> None:
    ensure_dir(path.parent)
    path.write_text(text, encoding="utf-8")


def prepare_index() -> None:
    write_text(
        DATASETS / "README.md",
        """# Coursework Dataset Index

This directory supplements the datasets required by `experiment/0_数据集说明.docx`.

| No. | Dataset | Status |
| --- | --- | --- |
| 1 | 男女数据集 | Already available in `experiment/男女data数据集/data/` |
| 2 | 模拟高斯数据 | Generated in `02_simulated_gaussian/` |
| 3 | UCI Iris | Downloadable/generated by `--download-small` |
| 4 | MNIST | Downloadable by `--download-small` |
| 5 | CIFAR-10 / CIFAR-100 | Downloadable by `--download-cifar` |
| 6 | ImageNet | Manual download required because of account/license/size |
| 7 | 飞机分类数据集 | Already available in `experiment/plane_dataset_4_1/`; manifest in `07_aircraft/` |

Recommended coursework split for generated tabular data: 60% train, 20% val,
20% test. Existing image datasets keep their official train/test layout; add a
validation split from training images when running experiments.
""",
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--download-small", action="store_true", help="Download UCI Iris and MNIST.")
    parser.add_argument("--download-cifar", action="store_true", help="Download CIFAR-10 and CIFAR-100 archives.")
    args = parser.parse_args()

    prepare_index()
    prepare_simulated()
    prepare_uci(args.download_small)
    prepare_mnist(args.download_small)
    prepare_cifar(args.download_cifar)
    prepare_imagenet()
    prepare_aircraft_manifest()
    print(f"Dataset workspace prepared at {DATASETS}")


if __name__ == "__main__":
    main()
