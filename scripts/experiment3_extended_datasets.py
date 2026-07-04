"""Experiment 3 extension: K-L/PCA feature extraction on multiple datasets."""

from __future__ import annotations

import csv
import gzip
import pickle
import struct
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from sklearn.decomposition import PCA
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
from sklearn.metrics import accuracy_score, confusion_matrix
from sklearn.neighbors import KNeighborsClassifier, NearestCentroid
from sklearn.preprocessing import StandardScaler


ROOT = Path(__file__).resolve().parents[1]
DATASETS = ROOT / "datasets"
OUT_DIR = ROOT / "results" / "experiment3_extended"
FIG_DIR = ROOT / "reports" / "figures" / "experiment3_extended"
RNG = np.random.default_rng(20260704)


def ensure_dirs() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    FIG_DIR.mkdir(parents=True, exist_ok=True)


def read_csv_dataset(path: Path, feature_cols: Sequence[str] | None = None) -> Tuple[np.ndarray, np.ndarray]:
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        fields = reader.fieldnames or []
    if feature_cols is None:
        feature_cols = [name for name in fields if name not in {"sample_id", "label"}]
    x = np.array([[float(row[name]) for name in feature_cols] for row in rows], dtype=np.float64)
    y = np.array([row["label"] for row in rows])
    return x, y


def read_mnist_images(path: Path, limit: int | None = None) -> np.ndarray:
    with gzip.open(path, "rb") as f:
        magic, count, rows, cols = struct.unpack(">IIII", f.read(16))
        if magic != 2051:
            raise ValueError(f"Unexpected MNIST image magic number: {magic}")
        n = count if limit is None else min(limit, count)
        data = np.frombuffer(f.read(n * rows * cols), dtype=np.uint8)
    return data.reshape(n, rows * cols).astype(np.float32) / 255.0


def read_mnist_labels(path: Path, limit: int | None = None) -> np.ndarray:
    with gzip.open(path, "rb") as f:
        magic, count = struct.unpack(">II", f.read(8))
        if magic != 2049:
            raise ValueError(f"Unexpected MNIST label magic number: {magic}")
        n = count if limit is None else min(limit, count)
        labels = np.frombuffer(f.read(n), dtype=np.uint8)
    return labels.astype(str)


def load_cifar_batch(path: Path, label_key: bytes) -> Tuple[np.ndarray, np.ndarray]:
    with path.open("rb") as f:
        obj = pickle.load(f, encoding="bytes")
    x = obj[b"data"].astype(np.float32) / 255.0
    y = np.array(obj[label_key]).astype(str)
    return x, y


def load_cifar10() -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    base = DATASETS / "05_cifar" / "cifar-10-batches-py"
    xs, ys = [], []
    for idx in range(1, 6):
        x, y = load_cifar_batch(base / f"data_batch_{idx}", b"labels")
        xs.append(x)
        ys.append(y)
    test_x, test_y = load_cifar_batch(base / "test_batch", b"labels")
    return np.vstack(xs), np.concatenate(ys), test_x, test_y


def load_cifar100() -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    base = DATASETS / "05_cifar" / "cifar-100-python"
    train_x, train_y = load_cifar_batch(base / "train", b"fine_labels")
    test_x, test_y = load_cifar_batch(base / "test", b"fine_labels")
    return train_x, train_y, test_x, test_y


def stratified_subset(x: np.ndarray, y: np.ndarray, per_class: int) -> Tuple[np.ndarray, np.ndarray]:
    indices: List[int] = []
    for label in np.unique(y):
        label_idx = np.flatnonzero(y == label)
        take = min(per_class, len(label_idx))
        indices.extend(RNG.choice(label_idx, size=take, replace=False).tolist())
    idx = np.array(indices, dtype=int)
    RNG.shuffle(idx)
    return x[idx], y[idx]


def standardize(train_x: np.ndarray, test_x: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
    scaler = StandardScaler()
    return scaler.fit_transform(train_x), scaler.transform(test_x)


def fit_pca(train_x: np.ndarray, n_components: int) -> PCA:
    n_components = min(n_components, train_x.shape[1], train_x.shape[0] - 1)
    return PCA(n_components=n_components, random_state=20260704).fit(train_x)


def pca_nearest_centroid(train_z: np.ndarray, train_y: np.ndarray, test_z: np.ndarray, n_components: int) -> Tuple[np.ndarray, PCA]:
    pca = fit_pca(train_z, n_components)
    train_p = pca.transform(train_z)
    test_p = pca.transform(test_z)
    clf = NearestCentroid().fit(train_p, train_y)
    return clf.predict(test_p), pca


def lda_predict(train_z: np.ndarray, train_y: np.ndarray, test_z: np.ndarray) -> np.ndarray:
    model = LinearDiscriminantAnalysis(solver="lsqr", shrinkage="auto")
    model.fit(train_z, train_y)
    return model.predict(test_z)


def evaluate_dataset(
    name: str,
    train_x: np.ndarray,
    train_y: np.ndarray,
    test_x: np.ndarray,
    test_y: np.ndarray,
    pca_k: int,
) -> Tuple[Dict[str, object], Dict[str, object]]:
    train_z, test_z = standardize(train_x, test_x)
    pred_pc1, pca1 = pca_nearest_centroid(train_z, train_y, test_z, 1)
    pred_pca_k, pca_k_model = pca_nearest_centroid(train_z, train_y, test_z, pca_k)
    pred_lda = lda_predict(train_z, train_y, test_z)
    pred_knn = KNeighborsClassifier(n_neighbors=5).fit(pca_k_model.transform(train_z), train_y).predict(pca_k_model.transform(test_z))

    row = {
        "dataset": name,
        "pca_k": pca_k_model.n_components_,
        "pc1_variance": float(pca_k_model.explained_variance_ratio_[0]),
        "pca_k_variance": float(pca_k_model.explained_variance_ratio_.sum()),
        "pc1_accuracy": accuracy_score(test_y, pred_pc1),
        "pca_k_accuracy": accuracy_score(test_y, pred_pca_k),
        "pca_knn_accuracy": accuracy_score(test_y, pred_knn),
        "lda_accuracy": accuracy_score(test_y, pred_lda),
    }

    labels = np.unique(np.concatenate([test_y, pred_pca_k, pred_lda]))
    info = {
        "train_z": train_z,
        "test_z": test_z,
        "test_y": test_y,
        "pred_pc1": pred_pc1,
        "pred_pca_k": pred_pca_k,
        "pred_knn": pred_knn,
        "pred_lda": pred_lda,
        "pca": pca_k_model,
        "labels": labels,
        "cm_pca_k": confusion_matrix(test_y, pred_pca_k, labels=labels),
        "cm_lda": confusion_matrix(test_y, pred_lda, labels=labels),
    }
    return row, info


def write_csv(path: Path, rows: List[Dict[str, object]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def format_percent(value: object) -> str:
    return f"{float(value) * 100:.2f}\\%"


def latex_escape(value: object) -> str:
    text = str(value)
    repl = {"\\": "\\textbackslash{}", "_": "\\_", "%": "\\%", "&": "\\&", "#": "\\#", "{": "\\{", "}": "\\}"}
    return "".join(repl.get(ch, ch) for ch in text)


def write_summary_table(path: Path, rows: List[Dict[str, object]]) -> None:
    lines = [
        "\\small",
        "\\begin{tabular}{>{\\raggedright\\arraybackslash}p{2.2cm}>{\\centering\\arraybackslash}p{1.2cm}>{\\centering\\arraybackslash}p{1.6cm}>{\\centering\\arraybackslash}p{1.6cm}>{\\centering\\arraybackslash}p{1.6cm}>{\\centering\\arraybackslash}p{1.6cm}>{\\centering\\arraybackslash}p{1.6cm}}",
        "\\toprule",
        "数据集 & PCA维数 & PC1方差 & PCA累计方差 & PC1准确率 & PCA分类准确率 & LDA准确率 \\\\",
        "\\midrule",
    ]
    for row in rows:
        lines.append(
            f"{latex_escape(row['dataset'])} & {int(row['pca_k'])} & {format_percent(row['pc1_variance'])} & "
            f"{format_percent(row['pca_k_variance'])} & {format_percent(row['pc1_accuracy'])} & "
            f"{format_percent(row['pca_k_accuracy'])} & {format_percent(row['lda_accuracy'])} \\\\"
        )
    lines += ["\\bottomrule", "\\end{tabular}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def plot_pca_lda_projection(train_z: np.ndarray, train_y: np.ndarray, info: Dict[str, object], title: str, path: Path) -> None:
    pca2 = PCA(n_components=2, random_state=20260704).fit(train_z)
    pca_xy = pca2.transform(train_z)
    lda = LinearDiscriminantAnalysis(solver="svd")
    lda_xy = lda.fit_transform(train_z, train_y)
    if lda_xy.shape[1] == 1:
        lda_xy = np.c_[lda_xy[:, 0], np.zeros(len(lda_xy))]
    labels = np.unique(train_y)
    markers = ["o", "^", "s", "D", "P", "X"]
    fig, axes = plt.subplots(1, 2, figsize=(9.8, 4.1))
    for ax, xy, subtitle in [(axes[0], pca_xy, "PCA/K-L projection"), (axes[1], lda_xy[:, :2], "Supervised LDA projection")]:
        for idx, label in enumerate(labels):
            mask = train_y == label
            ax.scatter(xy[mask, 0], xy[mask, 1], s=14, marker=markers[idx % len(markers)], label=str(label), alpha=0.78)
        ax.set_title(subtitle)
        ax.set_xlabel("component 1")
        ax.set_ylabel("component 2")
        ax.grid(alpha=0.18)
    if len(labels) <= 6:
        axes[1].legend(fontsize=7, loc="best")
    fig.suptitle(title)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_pc1_distribution(train_z: np.ndarray, train_y: np.ndarray, title: str, path: Path) -> None:
    pca = PCA(n_components=1, random_state=20260704).fit(train_z)
    proj = pca.transform(train_z)[:, 0]
    labels = np.unique(train_y)
    fig, ax = plt.subplots(figsize=(7, 3.8))
    for label in labels:
        ax.hist(proj[train_y == label], bins=14, alpha=0.55, label=str(label))
    ax.set_title(title)
    ax.set_xlabel("PC1 projection")
    ax.set_ylabel("Count")
    ax.legend(fontsize=8)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_explained_variance(infos: Dict[str, Dict[str, object]], path: Path) -> None:
    fig, ax = plt.subplots(figsize=(8.6, 4.5))
    for name in ["MNIST", "CIFAR-10", "CIFAR-100"]:
        pca = infos[name]["pca"]
        ratios = np.cumsum(pca.explained_variance_ratio_)
        ax.plot(np.arange(1, len(ratios) + 1), ratios, label=name, linewidth=1.8)
    ax.set_xlabel("Number of principal components")
    ax.set_ylabel("Cumulative explained variance")
    ax.set_ylim(0, 1.02)
    ax.set_title("K-L/PCA Cumulative Explained Variance")
    ax.grid(alpha=0.22)
    ax.legend()
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_accuracy(rows: List[Dict[str, object]], path: Path) -> None:
    labels = {
        "模拟二类高斯": "Sim-2",
        "模拟三类高斯": "Sim-3",
        "UCI Iris": "Iris",
        "MNIST": "MNIST",
        "CIFAR-10": "CIFAR-10",
        "CIFAR-100": "CIFAR-100",
    }
    x = np.arange(len(rows))
    width = 0.24
    fig, ax = plt.subplots(figsize=(9.5, 4.8))
    ax.bar(x - width, [row["pc1_accuracy"] for row in rows], width, label="PC1", color="#8da0cb")
    ax.bar(x, [row["pca_k_accuracy"] for row in rows], width, label="PCA+NC", color="#66c2a5")
    ax.bar(x + width, [row["lda_accuracy"] for row in rows], width, label="LDA", color="#fc8d62")
    ax.set_xticks(x)
    ax.set_xticklabels([labels[row["dataset"]] for row in rows])
    ax.set_ylabel("Test accuracy")
    ax.set_ylim(0, 1.05)
    ax.set_title("K-L/PCA Feature Extraction Accuracy")
    ax.grid(axis="y", alpha=0.22)
    ax.legend()
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_confusion(cm: np.ndarray, labels: Sequence[str], title: str, path: Path, figsize: Tuple[float, float] = (5.4, 4.8)) -> None:
    fig, ax = plt.subplots(figsize=figsize)
    im = ax.imshow(cm, cmap="Blues")
    ax.set_title(title)
    ax.set_xlabel("Predicted")
    ax.set_ylabel("True")
    ax.set_xticks(np.arange(len(labels)))
    ax.set_yticks(np.arange(len(labels)))
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_yticklabels(labels)
    threshold = cm.max() / 2 if cm.size else 0
    for i in range(cm.shape[0]):
        for j in range(cm.shape[1]):
            ax.text(j, i, str(cm[i, j]), ha="center", va="center", color="white" if cm[i, j] > threshold else "black", fontsize=8)
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_image_predictions(
    images: np.ndarray,
    y_true: np.ndarray,
    y_pred: np.ndarray,
    path: Path,
    title: str,
    image_shape: Tuple[int, ...],
    max_images: int = 20,
) -> None:
    wrong = np.flatnonzero(y_true != y_pred)
    correct = np.flatnonzero(y_true == y_pred)
    wrong_take = wrong[: min(len(wrong), max_images // 2)]
    correct_take = correct[: max_images - len(wrong_take)]
    indices = np.concatenate([correct_take, wrong_take])
    cols = 5
    rows = int(np.ceil(len(indices) / cols))
    fig, axes = plt.subplots(rows, cols, figsize=(8.5, 1.9 * rows))
    axes = np.atleast_1d(axes).ravel()
    for ax, idx in zip(axes, indices):
        img = images[idx]
        if len(image_shape) == 2:
            ax.imshow(img.reshape(image_shape), cmap="gray")
        else:
            ax.imshow(img.reshape(3, 32, 32).transpose(1, 2, 0))
        color = "#1b9e77" if y_true[idx] == y_pred[idx] else "crimson"
        ax.set_title(f"T:{y_true[idx]}  P:{y_pred[idx]}", fontsize=8, color=color)
        ax.axis("off")
    for ax in axes[len(indices) :]:
        ax.axis("off")
    fig.suptitle(title, y=0.995)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def main() -> None:
    ensure_dirs()
    rows: List[Dict[str, object]] = []
    infos: Dict[str, Dict[str, object]] = {}

    train_x, train_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "2class" / "train.csv")
    test_x, test_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "2class" / "test.csv")
    row, info = evaluate_dataset("模拟二类高斯", train_x, train_y, test_x, test_y, pca_k=2)
    rows.append(row)
    infos[row["dataset"]] = info
    plot_pc1_distribution(info["train_z"], train_y, "Simulated 2-class PC1 distribution", FIG_DIR / "simulated_2class_pc1_distribution.pdf")
    plot_pca_lda_projection(info["train_z"], train_y, info, "Simulated 2-class PCA and LDA projections", FIG_DIR / "simulated_2class_pca_lda_projection.pdf")

    train_x, train_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "3class" / "train.csv")
    test_x, test_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "3class" / "test.csv")
    row, info = evaluate_dataset("模拟三类高斯", train_x, train_y, test_x, test_y, pca_k=2)
    rows.append(row)
    infos[row["dataset"]] = info
    plot_pca_lda_projection(info["train_z"], train_y, info, "Simulated 3-class PCA and LDA projections", FIG_DIR / "simulated_3class_pca_lda_projection.pdf")

    train_x, train_y = read_csv_dataset(DATASETS / "03_uci_iris" / "train.csv")
    test_x, test_y = read_csv_dataset(DATASETS / "03_uci_iris" / "test.csv")
    row, info = evaluate_dataset("UCI Iris", train_x, train_y, test_x, test_y, pca_k=4)
    rows.append(row)
    infos[row["dataset"]] = info
    plot_pca_lda_projection(info["train_z"], train_y, info, "UCI Iris PCA and LDA projections", FIG_DIR / "iris_pca_lda_projection.pdf")
    plot_confusion(info["cm_pca_k"], info["labels"], "Iris PCA+NearestCentroid", FIG_DIR / "iris_pca_confusion.pdf")

    mnist_raw = DATASETS / "04_mnist" / "raw"
    train_x = read_mnist_images(mnist_raw / "train-images-idx3-ubyte.gz", limit=12000)
    train_y = read_mnist_labels(mnist_raw / "train-labels-idx1-ubyte.gz", limit=12000)
    test_x = read_mnist_images(mnist_raw / "t10k-images-idx3-ubyte.gz", limit=2000)
    test_y = read_mnist_labels(mnist_raw / "t10k-labels-idx1-ubyte.gz", limit=2000)
    train_x, train_y = stratified_subset(train_x, train_y, per_class=500)
    row, info = evaluate_dataset("MNIST", train_x, train_y, test_x, test_y, pca_k=64)
    rows.append(row)
    infos[row["dataset"]] = info
    plot_confusion(info["cm_pca_k"], info["labels"], "MNIST PCA+NearestCentroid", FIG_DIR / "mnist_pca_confusion.pdf", figsize=(6.2, 5.6))
    plot_pca_lda_projection(info["train_z"], train_y, info, "MNIST PCA and LDA projections", FIG_DIR / "mnist_pca_lda_projection.pdf")
    plot_image_predictions(test_x, test_y, info["pred_pca_k"], FIG_DIR / "mnist_pca_examples.pdf", "MNIST PCA prediction examples", (28, 28))

    train_x, train_y, test_x, test_y = load_cifar10()
    train_x, train_y = stratified_subset(train_x, train_y, per_class=450)
    test_x, test_y = stratified_subset(test_x, test_y, per_class=180)
    row, info = evaluate_dataset("CIFAR-10", train_x, train_y, test_x, test_y, pca_k=64)
    rows.append(row)
    infos[row["dataset"]] = info
    plot_confusion(info["cm_pca_k"], info["labels"], "CIFAR-10 PCA+NearestCentroid", FIG_DIR / "cifar10_pca_confusion.pdf", figsize=(6.2, 5.6))
    plot_pca_lda_projection(info["train_z"], train_y, info, "CIFAR-10 PCA and LDA projections", FIG_DIR / "cifar10_pca_lda_projection.pdf")
    plot_image_predictions(test_x, test_y, info["pred_pca_k"], FIG_DIR / "cifar10_pca_examples.pdf", "CIFAR-10 PCA prediction examples", (3, 32, 32))

    train_x, train_y, test_x, test_y = load_cifar100()
    train_x, train_y = stratified_subset(train_x, train_y, per_class=70)
    test_x, test_y = stratified_subset(test_x, test_y, per_class=15)
    row, info = evaluate_dataset("CIFAR-100", train_x, train_y, test_x, test_y, pca_k=64)
    rows.append(row)
    infos[row["dataset"]] = info
    plot_pca_lda_projection(info["train_z"], train_y, info, "CIFAR-100 PCA and LDA projections", FIG_DIR / "cifar100_pca_lda_projection.pdf")
    plot_image_predictions(test_x, test_y, info["pred_pca_k"], FIG_DIR / "cifar100_pca_examples.pdf", "CIFAR-100 PCA prediction examples", (3, 32, 32))

    plot_explained_variance(infos, FIG_DIR / "image_explained_variance.pdf")
    plot_accuracy(rows, FIG_DIR / "extended_accuracy_comparison.pdf")
    write_csv(OUT_DIR / "metrics.csv", rows)
    write_summary_table(OUT_DIR / "summary_table.tex", rows)
    print(f"Wrote metrics to {OUT_DIR / 'metrics.csv'}")
    print(f"Wrote figures to {FIG_DIR}")


if __name__ == "__main__":
    main()
