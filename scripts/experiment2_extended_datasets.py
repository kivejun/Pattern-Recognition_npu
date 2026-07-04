"""Experiment 2 extension on simulated, UCI, MNIST and CIFAR datasets.

The experiment compares three routes:
1. parametric Gaussian Bayes baseline;
2. non-parametric classifier, using Parzen on low-dimensional data and kNN on
   image PCA features;
3. Fisher/LDA direct linear classifier.
"""

from __future__ import annotations

import csv
import gzip
import pickle
import struct
from pathlib import Path
from typing import Callable, Dict, List, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from sklearn.decomposition import PCA
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
from sklearn.metrics import accuracy_score, confusion_matrix
from sklearn.naive_bayes import GaussianNB
from sklearn.neighbors import KNeighborsClassifier
from sklearn.preprocessing import StandardScaler


ROOT = Path(__file__).resolve().parents[1]
DATASETS = ROOT / "datasets"
OUT_DIR = ROOT / "results" / "experiment2_extended"
FIG_DIR = ROOT / "reports" / "figures" / "experiment2_extended"
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
    train_xs, train_ys = [], []
    for idx in range(1, 6):
        x, y = load_cifar_batch(base / f"data_batch_{idx}", b"labels")
        train_xs.append(x)
        train_ys.append(y)
    test_x, test_y = load_cifar_batch(base / "test_batch", b"labels")
    return np.vstack(train_xs), np.concatenate(train_ys), test_x, test_y


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
    indices_array = np.array(indices, dtype=int)
    RNG.shuffle(indices_array)
    return x[indices_array], y[indices_array]


def preprocess(
    train_x: np.ndarray,
    test_x: np.ndarray,
    use_pca: bool,
    n_components: int = 64,
) -> Tuple[np.ndarray, np.ndarray, str, StandardScaler, PCA | None]:
    scaler = StandardScaler()
    train_scaled = scaler.fit_transform(train_x)
    test_scaled = scaler.transform(test_x)
    if not use_pca:
        return train_scaled, test_scaled, "标准化原始特征", scaler, None

    n_components = min(n_components, train_scaled.shape[1], train_scaled.shape[0] - 1)
    pca = PCA(n_components=n_components, random_state=20260704)
    train_pca = pca.fit_transform(train_scaled)
    test_pca = pca.transform(test_scaled)
    explained = pca.explained_variance_ratio_.sum() * 100
    return train_pca, test_pca, f"PCA {n_components}维，累计方差{explained:.2f}%", scaler, pca


class FullGaussianBayes:
    def __init__(self, reg: float = 1e-4) -> None:
        self.reg = reg
        self.labels: np.ndarray | None = None
        self.priors: Dict[str, float] = {}
        self.means: Dict[str, np.ndarray] = {}
        self.covs: Dict[str, np.ndarray] = {}

    def fit(self, x: np.ndarray, y: np.ndarray) -> "FullGaussianBayes":
        self.labels = np.unique(y)
        for label in self.labels:
            class_x = x[y == label]
            centered = class_x - class_x.mean(axis=0)
            cov = centered.T @ centered / len(class_x)
            cov = cov + np.eye(cov.shape[0]) * self.reg
            self.priors[str(label)] = len(class_x) / len(x)
            self.means[str(label)] = class_x.mean(axis=0)
            self.covs[str(label)] = cov
        return self

    def predict(self, x: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        scores = []
        for label in self.labels:
            mean = self.means[str(label)]
            cov = self.covs[str(label)]
            inv_cov = np.linalg.inv(cov)
            sign, logdet = np.linalg.slogdet(cov)
            diff = x - mean
            quad = np.sum((diff @ inv_cov) * diff, axis=1)
            log_density = -0.5 * (x.shape[1] * np.log(2 * np.pi) + logdet + quad)
            scores.append(log_density + np.log(self.priors[str(label)]))
        return self.labels[np.argmax(np.vstack(scores).T, axis=1)]


class ParzenBayes:
    def __init__(self, bandwidth: float = 0.8) -> None:
        self.bandwidth = bandwidth
        self.labels: np.ndarray | None = None
        self.samples: Dict[str, np.ndarray] = {}
        self.priors: Dict[str, float] = {}

    def fit(self, x: np.ndarray, y: np.ndarray) -> "ParzenBayes":
        self.labels = np.unique(y)
        for label in self.labels:
            class_x = x[y == label]
            self.samples[str(label)] = class_x
            self.priors[str(label)] = len(class_x) / len(x)
        return self

    def predict(self, x: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        scores = []
        for label in self.labels:
            samples = self.samples[str(label)]
            diff = x[:, None, :] - samples[None, :, :]
            exponent = -0.5 * np.sum((diff / self.bandwidth) ** 2, axis=2)
            density = np.exp(exponent).mean(axis=1) / (self.bandwidth ** x.shape[1])
            scores.append(np.log(density + 1e-300) + np.log(self.priors[str(label)]))
        return self.labels[np.argmax(np.vstack(scores).T, axis=1)]


def fit_lda(x: np.ndarray, y: np.ndarray) -> LinearDiscriminantAnalysis:
    # shrinkage makes the image-feature LDA stable after PCA.
    model = LinearDiscriminantAnalysis(solver="lsqr", shrinkage="auto")
    model.fit(x, y)
    return model


def loo_error(
    x: np.ndarray,
    y: np.ndarray,
    fit_predict_one: Callable[[np.ndarray, np.ndarray, np.ndarray], str],
) -> float:
    errors = 0
    for idx in range(len(x)):
        mask = np.ones(len(x), dtype=bool)
        mask[idx] = False
        pred = fit_predict_one(x[mask], y[mask], x[idx : idx + 1])
        errors += int(pred != y[idx])
    return errors / len(x)


def fit_predict_gaussian(x: np.ndarray, y: np.ndarray, query: np.ndarray, full: bool) -> str:
    if full:
        model = FullGaussianBayes().fit(x, y)
    else:
        model = GaussianNB().fit(x, y)
    return str(model.predict(query)[0])


def fit_predict_parzen(x: np.ndarray, y: np.ndarray, query: np.ndarray, bandwidth: float) -> str:
    return str(ParzenBayes(bandwidth=bandwidth).fit(x, y).predict(query)[0])


def fit_predict_knn(x: np.ndarray, y: np.ndarray, query: np.ndarray, k: int) -> str:
    k_eff = min(k, len(x))
    return str(KNeighborsClassifier(n_neighbors=k_eff).fit(x, y).predict(query)[0])


def fit_predict_lda(x: np.ndarray, y: np.ndarray, query: np.ndarray) -> str:
    return str(fit_lda(x, y).predict(query)[0])


def compute_loo_subset(
    train_x: np.ndarray,
    train_y: np.ndarray,
    use_pca: bool,
    per_class: int,
    nonparam: str,
    gaussian_full: bool,
    bandwidth: float,
    k: int,
) -> Dict[str, float]:
    sub_x, sub_y = stratified_subset(train_x, train_y, per_class)
    sub_x, _, _, _, _ = preprocess(sub_x, sub_x, use_pca=use_pca)
    return {
        "Gaussian Bayes": loo_error(sub_x, sub_y, lambda a, b, q: fit_predict_gaussian(a, b, q, gaussian_full)),
        "Non-parametric": loo_error(
            sub_x,
            sub_y,
            (lambda a, b, q: fit_predict_parzen(a, b, q, bandwidth))
            if nonparam == "parzen"
            else (lambda a, b, q: fit_predict_knn(a, b, q, k)),
        ),
        "Fisher/LDA": loo_error(sub_x, sub_y, fit_predict_lda),
        "loo_n": float(len(sub_y)),
    }


def evaluate_dataset(
    name: str,
    train_x: np.ndarray,
    train_y: np.ndarray,
    test_x: np.ndarray,
    test_y: np.ndarray,
    use_pca: bool,
    gaussian_full: bool,
    nonparam: str,
    bandwidth: float = 0.8,
    k: int = 5,
    loo_per_class: int = 0,
) -> Tuple[List[Dict[str, object]], Dict[str, object]]:
    train_z, test_z, feature_note, _, _ = preprocess(train_x, test_x, use_pca=use_pca)

    if gaussian_full:
        gaussian = FullGaussianBayes().fit(train_z, train_y)
    else:
        gaussian = GaussianNB().fit(train_z, train_y)

    if nonparam == "parzen":
        non_model = ParzenBayes(bandwidth=bandwidth).fit(train_z, train_y)
        non_note = f"Parzen h={bandwidth}"
    else:
        non_model = KNeighborsClassifier(n_neighbors=k).fit(train_z, train_y)
        non_note = f"kNN k={k}"

    lda = fit_lda(train_z, train_y)
    predictions = {
        "Gaussian Bayes": gaussian.predict(test_z),
        "Non-parametric": non_model.predict(test_z),
        "Fisher/LDA": lda.predict(test_z),
    }

    loo_values = compute_loo_subset(
        train_x,
        train_y,
        use_pca=use_pca,
        per_class=loo_per_class,
        nonparam=nonparam,
        gaussian_full=gaussian_full,
        bandwidth=bandwidth,
        k=k,
    )

    rows: List[Dict[str, object]] = []
    for method, pred in predictions.items():
        rows.append(
            {
                "dataset": name,
                "method": method,
                "feature": feature_note,
                "setting": "full Gaussian" if method == "Gaussian Bayes" and gaussian_full else (
                    "diag Gaussian" if method == "Gaussian Bayes" else (
                        non_note if method == "Non-parametric" else "multi-class LDA"
                    )
                ),
                "test_accuracy": accuracy_score(test_y, pred),
                "test_error": 1 - accuracy_score(test_y, pred),
                "loo_error": loo_values[method],
                "loo_n": int(loo_values["loo_n"]),
            }
        )

    labels = np.unique(np.concatenate([test_y, *predictions.values()]))
    return rows, {
        "train_z": train_z,
        "test_z": test_z,
        "test_y": test_y,
        "predictions": predictions,
        "labels": labels,
        "cm_lda": confusion_matrix(test_y, predictions["Fisher/LDA"], labels=labels),
        "cm_non": confusion_matrix(test_y, predictions["Non-parametric"], labels=labels),
        "cm_gaussian": confusion_matrix(test_y, predictions["Gaussian Bayes"], labels=labels),
    }


def write_csv(path: Path, rows: List[Dict[str, object]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def format_percent(value: object) -> str:
    return f"{float(value) * 100:.2f}\\%"


def latex_escape(value: object) -> str:
    text = str(value)
    repl = {
        "\\": "\\textbackslash{}",
        "_": "\\_",
        "%": "\\%",
        "&": "\\&",
        "#": "\\#",
        "{": "\\{",
        "}": "\\}",
    }
    return "".join(repl.get(ch, ch) for ch in text)


def write_summary_table(path: Path, rows: List[Dict[str, object]]) -> None:
    datasets = []
    for row in rows:
        if row["dataset"] not in datasets:
            datasets.append(row["dataset"])
    lookup = {(row["dataset"], row["method"]): row for row in rows}
    lines = [
        "\\small",
        "\\begin{tabular}{>{\\raggedright\\arraybackslash}p{2.25cm}>{\\raggedright\\arraybackslash}p{3.35cm}>{\\centering\\arraybackslash}p{1.75cm}>{\\centering\\arraybackslash}p{1.85cm}>{\\centering\\arraybackslash}p{1.75cm}>{\\centering\\arraybackslash}p{1.55cm}}",
        "\\toprule",
        "数据集 & 特征与非参数设置 & Gaussian 错误率 & 非参数错误率 & LDA 错误率 & 留一样本 \\\\",
        "\\midrule",
    ]
    for dataset in datasets:
        gaussian = lookup[(dataset, "Gaussian Bayes")]
        non = lookup[(dataset, "Non-parametric")]
        lda = lookup[(dataset, "Fisher/LDA")]
        setting = f"{gaussian['feature']}；{non['setting']}"
        lines.append(
            f"{latex_escape(dataset)} & {latex_escape(setting)} & "
            f"{format_percent(gaussian['test_error'])} & {format_percent(non['test_error'])} & "
            f"{format_percent(lda['test_error'])} & {int(lda['loo_n'])} \\\\"
        )
    lines += ["\\bottomrule", "\\end{tabular}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


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
            color = "white" if cm[i, j] > threshold else "black"
            ax.text(j, i, str(cm[i, j]), ha="center", va="center", color=color, fontsize=8)
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_lowdim_boundary(
    train_x: np.ndarray,
    train_y: np.ndarray,
    name: str,
    path: Path,
    nonparam: str,
    bandwidth: float = 0.8,
    k: int = 5,
) -> None:
    scaler = StandardScaler()
    train_z = scaler.fit_transform(train_x)
    x_min, x_max = train_z[:, 0].min() - 0.8, train_z[:, 0].max() + 0.8
    y_min, y_max = train_z[:, 1].min() - 0.8, train_z[:, 1].max() + 0.8
    xx, yy = np.meshgrid(np.linspace(x_min, x_max, 260), np.linspace(y_min, y_max, 260))
    grid = np.c_[xx.ravel(), yy.ravel()]

    lda = fit_lda(train_z, train_y)
    if nonparam == "parzen":
        non_model = ParzenBayes(bandwidth=bandwidth).fit(train_z, train_y)
        non_title = f"Parzen Bayes (h={bandwidth})"
    else:
        non_model = KNeighborsClassifier(n_neighbors=k).fit(train_z, train_y)
        non_title = f"kNN (k={k})"

    methods = [("Fisher/LDA", lda.predict(grid)), (non_title, non_model.predict(grid))]
    labels = np.unique(train_y)
    fig, axes = plt.subplots(1, 2, figsize=(10, 4.1), sharex=True, sharey=True)
    colors = ["#fde2cc", "#cfeae7", "#d8d5ee", "#f4d2df", "#d9ead3"]
    markers = ["o", "^", "s", "D", "P"]
    for ax, (title, pred) in zip(axes, methods):
        label_to_int = {label: idx for idx, label in enumerate(labels)}
        zz = np.array([label_to_int[p] for p in pred]).reshape(xx.shape)
        ax.contourf(xx, yy, zz, levels=np.arange(len(labels) + 1) - 0.5, colors=colors[: len(labels)], alpha=0.85)
        ax.contour(xx, yy, zz, levels=np.arange(0.5, len(labels) - 0.5 + 1e-9, 1), colors="black", linewidths=0.55)
        for idx, label in enumerate(labels):
            mask = train_y == label
            ax.scatter(train_z[mask, 0], train_z[mask, 1], s=16, marker=markers[idx % len(markers)], label=str(label), alpha=0.82)
        ax.set_title(title)
        ax.set_xlabel("standardized feature 1")
        ax.grid(alpha=0.18)
    axes[0].set_ylabel("standardized feature 2")
    axes[1].legend(loc="best", fontsize=8)
    fig.suptitle(name)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_pca_boundary(train_x: np.ndarray, train_y: np.ndarray, title: str, path: Path, k: int = 5) -> None:
    scaler = StandardScaler()
    z = scaler.fit_transform(train_x)
    z2 = PCA(n_components=2, random_state=20260704).fit_transform(z)
    plot_lowdim_boundary(z2, train_y, title, path, nonparam="knn", k=k)


def plot_projection(test_z: np.ndarray, test_y: np.ndarray, pred: np.ndarray, title: str, path: Path) -> None:
    z2 = PCA(n_components=2, random_state=20260704).fit_transform(test_z)
    correct = pred == test_y
    fig, ax = plt.subplots(figsize=(6.8, 4.8))
    ax.scatter(z2[correct, 0], z2[correct, 1], s=12, c="#1b9e77", alpha=0.52, label="correct")
    ax.scatter(z2[~correct, 0], z2[~correct, 1], s=18, c="#d95f02", alpha=0.82, label="wrong")
    ax.set_xlabel("PCA component 1")
    ax.set_ylabel("PCA component 2")
    ax.set_title(title)
    ax.legend()
    ax.grid(alpha=0.2)
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
    if len(indices) == 0:
        indices = np.arange(min(max_images, len(images)))

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
        ok = y_true[idx] == y_pred[idx]
        color = "#1b9e77" if ok else "crimson"
        ax.set_title(f"T:{y_true[idx]}  P:{y_pred[idx]}", fontsize=8, color=color)
        ax.axis("off")
    for ax in axes[len(indices) :]:
        ax.axis("off")
    fig.suptitle(title, y=0.995)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_error_comparison(rows: List[Dict[str, object]]) -> None:
    datasets = []
    for row in rows:
        if row["dataset"] not in datasets:
            datasets.append(row["dataset"])
    methods = ["Gaussian Bayes", "Non-parametric", "Fisher/LDA"]
    method_labels = ["Gaussian", "Nonparam", "Fisher/LDA"]
    short_names = {
        "模拟二类高斯": "Sim-2",
        "模拟三类高斯": "Sim-3",
        "UCI Iris": "Iris",
        "MNIST": "MNIST",
        "CIFAR-10": "CIFAR-10",
        "CIFAR-100": "CIFAR-100",
    }
    lookup = {(row["dataset"], row["method"]): float(row["test_error"]) for row in rows}
    x = np.arange(len(datasets))
    width = 0.24
    fig, ax = plt.subplots(figsize=(9.6, 4.8))
    colors = ["#8da0cb", "#66c2a5", "#fc8d62"]
    for idx, method in enumerate(methods):
        values = [lookup[(dataset, method)] for dataset in datasets]
        ax.bar(x + (idx - 1) * width, values, width, label=method_labels[idx], color=colors[idx])
    ax.set_xticks(x)
    ax.set_xticklabels([short_names[d] for d in datasets])
    ax.set_ylabel("Test error rate")
    ax.set_title("Experiment 2 Extended Dataset Error Comparison")
    ax.grid(axis="y", alpha=0.22)
    ax.legend()
    fig.tight_layout()
    fig.savefig(FIG_DIR / "extended_error_comparison.pdf")
    plt.close(fig)


def main() -> None:
    ensure_dirs()
    all_rows: List[Dict[str, object]] = []

    train_x, train_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "2class" / "train.csv")
    test_x, test_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "2class" / "test.csv")
    rows, info = evaluate_dataset("模拟二类高斯", train_x, train_y, test_x, test_y, False, True, "parzen", 0.8, loo_per_class=120)
    all_rows.extend(rows)
    plot_lowdim_boundary(train_x, train_y, "Simulated 2-class: Fisher/LDA and Parzen", FIG_DIR / "simulated_2class_lda_parzen_boundary.pdf", "parzen", 0.8)
    plot_confusion(info["cm_lda"], info["labels"], "Simulated 2-class LDA", FIG_DIR / "simulated_2class_lda_confusion.pdf")

    train_x, train_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "3class" / "train.csv")
    test_x, test_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "3class" / "test.csv")
    rows, info = evaluate_dataset("模拟三类高斯", train_x, train_y, test_x, test_y, False, True, "parzen", 0.8, loo_per_class=80)
    all_rows.extend(rows)
    plot_lowdim_boundary(train_x, train_y, "Simulated 3-class: LDA and Parzen", FIG_DIR / "simulated_3class_lda_parzen_boundary.pdf", "parzen", 0.8)
    plot_confusion(info["cm_lda"], info["labels"], "Simulated 3-class LDA", FIG_DIR / "simulated_3class_lda_confusion.pdf")

    train_x, train_y = read_csv_dataset(DATASETS / "03_uci_iris" / "train.csv")
    test_x, test_y = read_csv_dataset(DATASETS / "03_uci_iris" / "test.csv")
    rows, info = evaluate_dataset("UCI Iris", train_x, train_y, test_x, test_y, False, True, "parzen", 0.8, loo_per_class=30)
    all_rows.extend(rows)
    plot_pca_boundary(train_x, train_y, "UCI Iris PCA plane: LDA and kNN", FIG_DIR / "iris_lda_knn_boundary.pdf", k=5)
    plot_confusion(info["cm_lda"], info["labels"], "Iris LDA", FIG_DIR / "iris_lda_confusion.pdf")

    mnist_raw = DATASETS / "04_mnist" / "raw"
    train_x = read_mnist_images(mnist_raw / "train-images-idx3-ubyte.gz", limit=12000)
    train_y = read_mnist_labels(mnist_raw / "train-labels-idx1-ubyte.gz", limit=12000)
    test_x = read_mnist_images(mnist_raw / "t10k-images-idx3-ubyte.gz", limit=2000)
    test_y = read_mnist_labels(mnist_raw / "t10k-labels-idx1-ubyte.gz", limit=2000)
    train_x, train_y = stratified_subset(train_x, train_y, per_class=500)
    rows, info = evaluate_dataset("MNIST", train_x, train_y, test_x, test_y, True, False, "knn", k=5, loo_per_class=15)
    all_rows.extend(rows)
    plot_confusion(info["cm_lda"], info["labels"], "MNIST LDA", FIG_DIR / "mnist_lda_confusion.pdf", figsize=(6.2, 5.6))
    plot_projection(info["test_z"], info["test_y"], info["predictions"]["Fisher/LDA"], "MNIST LDA prediction on PCA features", FIG_DIR / "mnist_lda_projection.pdf")
    plot_image_predictions(test_x, test_y, info["predictions"]["Fisher/LDA"], FIG_DIR / "mnist_lda_examples.pdf", "MNIST LDA correct and wrong examples", (28, 28))

    train_x, train_y, test_x, test_y = load_cifar10()
    train_x, train_y = stratified_subset(train_x, train_y, per_class=450)
    test_x, test_y = stratified_subset(test_x, test_y, per_class=180)
    rows, info = evaluate_dataset("CIFAR-10", train_x, train_y, test_x, test_y, True, False, "knn", k=7, loo_per_class=12)
    all_rows.extend(rows)
    plot_confusion(info["cm_lda"], info["labels"], "CIFAR-10 LDA", FIG_DIR / "cifar10_lda_confusion.pdf", figsize=(6.2, 5.6))
    plot_projection(info["test_z"], info["test_y"], info["predictions"]["Fisher/LDA"], "CIFAR-10 LDA prediction on PCA features", FIG_DIR / "cifar10_lda_projection.pdf")
    plot_image_predictions(test_x, test_y, info["predictions"]["Fisher/LDA"], FIG_DIR / "cifar10_lda_examples.pdf", "CIFAR-10 LDA correct and wrong examples", (3, 32, 32))

    train_x, train_y, test_x, test_y = load_cifar100()
    train_x, train_y = stratified_subset(train_x, train_y, per_class=70)
    test_x, test_y = stratified_subset(test_x, test_y, per_class=15)
    rows, info = evaluate_dataset("CIFAR-100", train_x, train_y, test_x, test_y, True, False, "knn", k=5, loo_per_class=3)
    all_rows.extend(rows)
    plot_projection(info["test_z"], info["test_y"], info["predictions"]["Fisher/LDA"], "CIFAR-100 LDA prediction on PCA features", FIG_DIR / "cifar100_lda_projection.pdf")
    plot_image_predictions(test_x, test_y, info["predictions"]["Fisher/LDA"], FIG_DIR / "cifar100_lda_examples.pdf", "CIFAR-100 LDA correct and wrong examples", (3, 32, 32))

    write_csv(OUT_DIR / "metrics.csv", all_rows)
    write_summary_table(OUT_DIR / "summary_table.tex", all_rows)
    plot_error_comparison(all_rows)
    print(f"Wrote metrics to {OUT_DIR / 'metrics.csv'}")
    print(f"Wrote figures to {FIG_DIR}")


if __name__ == "__main__":
    main()
