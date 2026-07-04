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


def count_by_class(labels: np.ndarray, classes: Sequence[str]) -> Dict[str, int]:
    return {name: int(np.sum(labels == name)) for name in classes}


def write_dataset_table(
    path: Path,
    train_counts: Dict[str, int],
    test_counts: Dict[str, int],
    classes: Sequence[str],
) -> None:
    lines = [
        "\\begin{tabular}{lrrr}",
        "\\toprule",
        "类别 & 训练样本数 & 测试样本数 & 合计 \\\\",
        "\\midrule",
    ]
    for name in classes:
        total = train_counts[name] + test_counts[name]
        lines.append(f"{name} & {train_counts[name]} & {test_counts[name]} & {total} \\\\")
    lines += [
        "\\midrule",
        f"合计 & {sum(train_counts.values())} & {sum(test_counts.values())} & "
        f"{sum(train_counts.values()) + sum(test_counts.values())} \\\\",
        "\\bottomrule",
        "\\end{tabular}",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


def write_summary_table(path: Path, rows: Sequence[Dict[str, object]]) -> None:
    lines = [
        "\\begin{tabular}{lccc}",
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


def write_component_table(
    path: Path,
    rows: Sequence[Dict[str, object]],
    cumulative_variance: np.ndarray,
) -> None:
    by_components: Dict[int, Dict[str, Dict[str, object]]] = {}
    for row in rows:
        by_components.setdefault(int(row["n_components"]), {})[str(row["method"])] = row

    lines = [
        "\\begin{tabular}{crrr}",
        "\\toprule",
        "主成分数 & 累计方差 & 最近质心测试准确率 & 1NN 测试准确率 \\\\",
        "\\midrule",
    ]
    for n_components in sorted(by_components):
        nearest = by_components[n_components]["PCA+NearestCentroid"]
        knn = by_components[n_components]["PCA+1NN"]
        variance = cumulative_variance[n_components - 1]
        lines.append(
            f"{n_components} & {format_percent(float(variance))} & "
            f"{format_percent(float(nearest['test_accuracy']))} & "
            f"{format_percent(float(knn['test_accuracy']))} \\\\"
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


def plot_class_distribution(
    train_counts: Dict[str, int],
    test_counts: Dict[str, int],
    classes: Sequence[str],
) -> None:
    x = np.arange(len(classes))
    width = 0.38
    fig, ax = plt.subplots(figsize=(9.6, 4.8))
    ax.bar(x - width / 2, [train_counts[name] for name in classes], width, label="Train")
    ax.bar(x + width / 2, [test_counts[name] for name in classes], width, label="Test")
    ax.set_xticks(x, classes, rotation=45, ha="right")
    ax.set_ylabel("Number of images")
    ax.set_title("Aircraft Dataset Class Distribution")
    ax.grid(axis="y", alpha=0.25)
    ax.legend()
    plt.tight_layout()
    plt.savefig(FIG_DIR / "class_distribution.pdf")
    plt.close()


def plot_explained_variance(cumulative_variance: np.ndarray) -> None:
    xs = np.arange(1, len(cumulative_variance) + 1)
    fig, ax = plt.subplots(figsize=(7.5, 4.5))
    ax.plot(xs, cumulative_variance, linewidth=2.0, color="#1f77b4")
    ax.scatter(COMPONENT_GRID, cumulative_variance[np.array(COMPONENT_GRID) - 1], color="#d62728", zorder=3)
    ax.axhline(0.9, linestyle="--", color="#555555", linewidth=1.0)
    ax.set_xlabel("Number of PCA components")
    ax.set_ylabel("Cumulative explained variance")
    ax.set_title("K-L/PCA Cumulative Explained Variance")
    ax.set_ylim(0, 1.02)
    ax.grid(alpha=0.25)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "explained_variance.pdf")
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


def plot_pca_prediction_projection(
    z_test: np.ndarray,
    y_true: np.ndarray,
    y_pred: np.ndarray,
) -> None:
    correct = y_true == y_pred
    fig, ax = plt.subplots(figsize=(7, 5.2))
    ax.scatter(
        z_test[correct, 0],
        z_test[correct, 1],
        s=18,
        alpha=0.72,
        label="Correct",
        color="#2ca02c",
    )
    ax.scatter(
        z_test[~correct, 0],
        z_test[~correct, 1],
        s=22,
        alpha=0.82,
        label="Wrong",
        color="#d62728",
        marker="x",
    )
    ax.set_xlabel("PC1")
    ax.set_ylabel("PC2")
    ax.set_title("Aircraft Test Samples in the First Two PCA Dimensions")
    ax.grid(alpha=0.25)
    ax.legend()
    plt.tight_layout()
    plt.savefig(FIG_DIR / "pca_prediction_projection.pdf")
    plt.close()


def plot_per_class_recall(y_true: np.ndarray, y_pred: np.ndarray, classes: Sequence[str]) -> None:
    cm = confusion_matrix(y_true, y_pred, labels=list(classes))
    recalls = np.diag(cm) / np.maximum(cm.sum(axis=1), 1)
    fig, ax = plt.subplots(figsize=(9.4, 4.8))
    ax.bar(np.arange(len(classes)), recalls, color="#4c78a8")
    ax.set_xticks(np.arange(len(classes)), classes, rotation=45, ha="right")
    ax.set_ylim(0, 1.02)
    ax.set_ylabel("Recall")
    ax.set_title("Per-class Recall of PCA+1NN")
    ax.grid(axis="y", alpha=0.25)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "per_class_recall.pdf")
    plt.close()


def plot_prediction_examples(
    paths: List[Path],
    y_true: np.ndarray,
    y_pred: np.ndarray,
) -> None:
    correct_indices = np.flatnonzero(y_true == y_pred)[:8]
    wrong_indices = np.flatnonzero(y_true != y_pred)[:8]
    selected = list(correct_indices) + list(wrong_indices)
    fig, axes = plt.subplots(4, 4, figsize=(8.4, 7.8))
    for ax, idx in zip(axes.ravel(), selected):
        ax.imshow(read_image(paths[int(idx)]).reshape(IMAGE_SIZE), cmap="gray")
        ok = y_true[int(idx)] == y_pred[int(idx)]
        color = "#2ca02c" if ok else "#d62728"
        ax.set_title(f"T:{y_true[int(idx)]}\nP:{y_pred[int(idx)]}", fontsize=8, color=color)
        ax.axis("off")
    for ax in axes.ravel()[len(selected) :]:
        ax.axis("off")
    plt.suptitle("PCA+1NN Prediction Examples", fontsize=12)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "prediction_examples.pdf")
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
    x_test, y_test, test_paths = load_split("test")
    classes = sorted(set(y_train))
    train_counts = count_by_class(y_train, classes)
    test_counts = count_by_class(y_test, classes)

    rows: List[Dict[str, object]] = []
    for n_components in COMPONENT_GRID:
        rows.extend(fit_and_eval(x_train, y_train, x_test, y_test, n_components))

    scaler = StandardScaler()
    x_train_scaled = scaler.fit_transform(x_train)
    max_components = min(max(COMPONENT_GRID), x_train.shape[0], x_train.shape[1])
    pca_full = PCA(n_components=max_components, random_state=2026)
    pca_full.fit(x_train_scaled)
    cumulative_variance = np.cumsum(pca_full.explained_variance_ratio_)

    write_csv(OUT_DIR / "metrics.csv", rows)
    write_dataset_table(OUT_DIR / "dataset_table.tex", train_counts, test_counts, classes)
    write_summary_table(OUT_DIR / "summary_table.tex", rows)
    write_component_table(OUT_DIR / "component_table.tex", rows, cumulative_variance)
    plot_accuracy_curve(rows)
    plot_class_distribution(train_counts, test_counts, classes)
    plot_explained_variance(cumulative_variance)
    plot_samples(train_paths)

    pca = PCA(n_components=N_COMPONENTS, random_state=2026)
    z_train = pca.fit_transform(x_train_scaled)
    z_test = pca.transform(scaler.transform(x_test))
    clf = KNeighborsClassifier(n_neighbors=1)
    clf.fit(z_train, y_train)
    y_pred = clf.predict(z_test)
    plot_components(pca)
    plot_confusion(y_test, y_pred, classes)
    plot_pca_prediction_projection(z_test, y_test, y_pred)
    plot_per_class_recall(y_test, y_pred, classes)
    plot_prediction_examples(test_paths, y_test, y_pred)

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
