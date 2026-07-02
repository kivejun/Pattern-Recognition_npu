"""Experiment 4: K-L/PCA image recognition on the aircraft dataset."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

import cv2
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from sklearn.decomposition import PCA
from sklearn.metrics import accuracy_score, confusion_matrix
from sklearn.neighbors import KNeighborsClassifier
from sklearn.neighbors import NearestCentroid
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "experiment" / "plane_dataset_4_1"
OUT_DIR = ROOT / "results" / "experiment4_kl_image"
FIG_DIR = ROOT / "reports" / "figures" / "experiment4"

IMAGE_SIZE = (32, 32)
N_COMPONENTS = 64
COMPONENT_GRID = [8, 16, 32, 64, 128]


def ensure_dirs() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    FIG_DIR.mkdir(parents=True, exist_ok=True)


def image_paths(split: str) -> List[Tuple[Path, str]]:
    split_dir = DATA_DIR / split
    items: List[Tuple[Path, str]] = []
    for class_dir in sorted(p for p in split_dir.iterdir() if p.is_dir()):
        for path in sorted(class_dir.iterdir()):
            if path.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp"}:
                items.append((path, class_dir.name))
    return items


def read_image(path: Path) -> np.ndarray:
    data = np.fromfile(str(path), dtype=np.uint8)
    image = cv2.imdecode(data, cv2.IMREAD_GRAYSCALE)
    if image is None:
        raise ValueError(f"Could not read image: {path}")
    image = cv2.resize(image, IMAGE_SIZE, interpolation=cv2.INTER_AREA)
    return image.astype(np.float32) / 255.0


def load_split(split: str) -> Tuple[np.ndarray, np.ndarray, List[Path]]:
    rows = []
    labels = []
    paths = []
    for path, label in image_paths(split):
        rows.append(read_image(path).reshape(-1))
        labels.append(label)
        paths.append(path)
    return np.vstack(rows), np.array(labels), paths


def fit_and_eval(
    x_train: np.ndarray,
    y_train: np.ndarray,
    x_test: np.ndarray,
    y_test: np.ndarray,
    n_components: int,
) -> List[Dict[str, object]]:
    n_components = min(n_components, x_train.shape[0], x_train.shape[1])
    models = {
        "PCA+NearestCentroid": make_pipeline(
            StandardScaler(),
            PCA(n_components=n_components, random_state=2026),
            NearestCentroid(),
        ),
        "PCA+1NN": make_pipeline(
            StandardScaler(),
            PCA(n_components=n_components, random_state=2026),
            KNeighborsClassifier(n_neighbors=1),
        ),
    }
    rows: List[Dict[str, object]] = []
    for name, model in models.items():
        model.fit(x_train, y_train)
        pred_train = model.predict(x_train)
        pred_test = model.predict(x_test)
        rows.append(
            {
                "method": name,
                "n_components": n_components,
                "train_accuracy": accuracy_score(y_train, pred_train),
                "test_accuracy": accuracy_score(y_test, pred_test),
            }
        )
    return rows


def write_csv(path: Path, rows: Sequence[Dict[str, object]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def format_percent(value: float) -> str:
    return f"{value * 100:.2f}\\%"


def write_summary_table(path: Path, rows: Sequence[Dict[str, object]]) -> None:
    lines = [
        "\\begin{tabular}{p{3.2cm}p{2.2cm}p{2.2cm}p{2.2cm}}",
        "\\toprule",
        "方法 & 主成分数 & 训练准确率 & 测试准确率 \\\\",
        "\\midrule",
    ]
    selected = [row for row in rows if int(row["n_components"]) == N_COMPONENTS]
    for row in selected:
        lines.append(
            f"{row['method']} & {int(row['n_components'])} & "
            f"{format_percent(float(row['train_accuracy']))} & "
            f"{format_percent(float(row['test_accuracy']))} \\\\"
        )
    lines += ["\\bottomrule", "\\end{tabular}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def plot_accuracy_curve(rows: Sequence[Dict[str, object]]) -> None:
    plt.figure(figsize=(7, 4.5))
    methods = sorted(set(row["method"] for row in rows))
    for method in methods:
        subset = [row for row in rows if row["method"] == method]
        subset = sorted(subset, key=lambda row: int(row["n_components"]))
        xs = [int(row["n_components"]) for row in subset]
        ys = [float(row["test_accuracy"]) for row in subset]
        plt.plot(xs, ys, marker="o", linewidth=1.8, label=method)
    plt.xlabel("Number of PCA components")
    plt.ylabel("Test accuracy")
    plt.title("Aircraft Recognition Accuracy with K-L/PCA Features")
    plt.grid(alpha=0.25)
    plt.legend()
    plt.tight_layout()
    plt.savefig(FIG_DIR / "accuracy_vs_components.pdf")
    plt.close()


def plot_components(pca: PCA) -> None:
    fig, axes = plt.subplots(2, 5, figsize=(8, 3.6))
    for idx, ax in enumerate(axes.ravel()):
        comp = pca.components_[idx].reshape(IMAGE_SIZE)
        ax.imshow(comp, cmap="gray")
        ax.set_title(f"PC{idx + 1}", fontsize=9)
        ax.axis("off")
    plt.suptitle("First K-L/PCA Image Components", fontsize=12)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "pca_components.pdf")
    plt.close()


def plot_samples(paths: List[Path]) -> None:
    selected = []
    seen = set()
    for path in paths:
        label = path.parent.name
        if label in seen:
            continue
        seen.add(label)
        selected.append(path)
        if len(selected) >= 10:
            break
    fig, axes = plt.subplots(2, 5, figsize=(8, 3.8))
    for ax, path in zip(axes.ravel(), selected):
        ax.imshow(read_image(path).reshape(IMAGE_SIZE), cmap="gray")
        ax.set_title(path.parent.name, fontsize=9)
        ax.axis("off")
    plt.suptitle("Aircraft Dataset Samples", fontsize=12)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "sample_images.pdf")
    plt.close()


def plot_confusion(y_true: np.ndarray, y_pred: np.ndarray, classes: Sequence[str]) -> None:
    cm = confusion_matrix(y_true, y_pred, labels=list(classes))
    cm_norm = cm / np.maximum(cm.sum(axis=1, keepdims=True), 1)
    fig, ax = plt.subplots(figsize=(8.5, 7.2))
    im = ax.imshow(cm_norm, cmap="Blues", vmin=0, vmax=1)
    ax.set_xticks(np.arange(len(classes)), classes, rotation=45, ha="right", fontsize=8)
    ax.set_yticks(np.arange(len(classes)), classes, fontsize=8)
    ax.set_xlabel("Predicted")
    ax.set_ylabel("True")
    ax.set_title("PCA+1NN Normalized Confusion Matrix")
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "confusion_matrix.pdf")
    plt.close()


def main() -> None:
    ensure_dirs()
    x_train, y_train, train_paths = load_split("train")
    x_test, y_test, _ = load_split("test")
    classes = sorted(set(y_train))

    rows: List[Dict[str, object]] = []
    for n_components in COMPONENT_GRID:
        rows.extend(fit_and_eval(x_train, y_train, x_test, y_test, n_components))

    write_csv(OUT_DIR / "metrics.csv", rows)
    write_summary_table(OUT_DIR / "summary_table.tex", rows)
    plot_accuracy_curve(rows)
    plot_samples(train_paths)

    scaler = StandardScaler()
    x_train_scaled = scaler.fit_transform(x_train)
    pca = PCA(n_components=N_COMPONENTS, random_state=2026)
    z_train = pca.fit_transform(x_train_scaled)
    z_test = pca.transform(scaler.transform(x_test))
    clf = KNeighborsClassifier(n_neighbors=1)
    clf.fit(z_train, y_train)
    y_pred = clf.predict(z_test)
    plot_components(pca)
    plot_confusion(y_test, y_pred, classes)

    explained = pca.explained_variance_ratio_
    with (OUT_DIR / "pca_params.txt").open("w", encoding="utf-8") as f:
        f.write(f"image_size: {IMAGE_SIZE[0]}x{IMAGE_SIZE[1]}\n")
        f.write(f"train_samples: {len(x_train)}\n")
        f.write(f"test_samples: {len(x_test)}\n")
        f.write(f"num_classes: {len(classes)}\n")
        f.write(f"n_components: {N_COMPONENTS}\n")
        f.write(f"explained_variance_ratio_sum: {explained.sum():.6f}\n")
        f.write(f"test_accuracy_pca_1nn: {accuracy_score(y_test, y_pred):.6f}\n")
        f.write("classes: " + ", ".join(classes) + "\n")

    print(f"Wrote metrics to {OUT_DIR / 'metrics.csv'}")
    print(f"Wrote figures to {FIG_DIR}")


if __name__ == "__main__":
    main()
