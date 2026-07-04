"""Experiment 1 extension: Bayes classifiers on simulated, UCI, MNIST and CIFAR data."""

from __future__ import annotations

import csv
import gzip
import math
import pickle
import struct
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from sklearn.decomposition import PCA
from sklearn.metrics import accuracy_score, confusion_matrix, precision_score, recall_score
from sklearn.preprocessing import StandardScaler


ROOT = Path(__file__).resolve().parents[1]
DATASETS = ROOT / "datasets"
OUT_DIR = ROOT / "results" / "experiment1_extended"
FIG_DIR = ROOT / "reports" / "figures" / "experiment1_extended"
RNG = np.random.default_rng(20260703)


def ensure_dirs() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    FIG_DIR.mkdir(parents=True, exist_ok=True)


def read_csv_dataset(path: Path, feature_cols: Sequence[str] | None = None) -> Tuple[np.ndarray, np.ndarray]:
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    if feature_cols is None:
        feature_cols = [c for c in reader.fieldnames or [] if c not in {"sample_id", "label"}]
    x = np.array([[float(row[c]) for c in feature_cols] for row in rows], dtype=np.float64)
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
    indices = np.array(indices)
    RNG.shuffle(indices)
    return x[indices], y[indices]


def preprocess(
    train_x: np.ndarray,
    test_x: np.ndarray,
    use_pca: bool,
    n_components: int = 64,
) -> Tuple[np.ndarray, np.ndarray, str]:
    scaler = StandardScaler()
    train_scaled = scaler.fit_transform(train_x)
    test_scaled = scaler.transform(test_x)
    if not use_pca:
        return train_scaled, test_scaled, "标准化原始特征"
    n_components = min(n_components, train_scaled.shape[1], train_scaled.shape[0] - 1)
    pca = PCA(n_components=n_components, random_state=20260703)
    train_pca = pca.fit_transform(train_scaled)
    test_pca = pca.transform(test_scaled)
    explained = pca.explained_variance_ratio_.sum() * 100
    return train_pca, test_pca, f"PCA {n_components}维，累计方差 {explained:.2f}\\%"


def preprocess_with_pca2(train_x: np.ndarray, test_x: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
    scaler = StandardScaler()
    train_scaled = scaler.fit_transform(train_x)
    test_scaled = scaler.transform(test_x)
    pca = PCA(n_components=2, random_state=20260703)
    return pca.fit_transform(train_scaled), pca.transform(test_scaled)


class GaussianBayes:
    def __init__(self, covariance_mode: str = "full", reg: float = 1e-4) -> None:
        self.covariance_mode = covariance_mode
        self.reg = reg
        self.labels: np.ndarray | None = None
        self.priors: Dict[str, float] = {}
        self.params: Dict[str, Tuple[np.ndarray, np.ndarray]] = {}

    def fit(self, x: np.ndarray, y: np.ndarray) -> "GaussianBayes":
        self.labels = np.unique(y)
        for label in self.labels:
            class_x = x[y == label]
            mean = class_x.mean(axis=0)
            centered = class_x - mean
            cov = centered.T @ centered / len(class_x)
            if self.covariance_mode == "diag":
                cov = np.diag(np.diag(cov))
            cov = cov + np.eye(cov.shape[0]) * self.reg
            self.params[str(label)] = (mean, cov)
            self.priors[str(label)] = len(class_x) / len(x)
        return self

    def log_scores(self, x: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        scores = []
        for label in self.labels:
            mean, cov = self.params[str(label)]
            inv_cov = np.linalg.inv(cov)
            sign, logdet = np.linalg.slogdet(cov)
            if sign <= 0:
                raise ValueError("Covariance matrix is not positive definite.")
            diff = x - mean
            quad = np.sum((diff @ inv_cov) * diff, axis=1)
            log_density = -0.5 * (x.shape[1] * math.log(2 * math.pi) + logdet + quad)
            scores.append(log_density + math.log(self.priors[str(label)]))
        return np.vstack(scores).T

    def predict(self, x: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        return self.labels[np.argmax(self.log_scores(x), axis=1)]

    def predict_risk(self, x: np.ndarray, loss: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        log_scores = self.log_scores(x)
        shifted = log_scores - log_scores.max(axis=1, keepdims=True)
        posterior = np.exp(shifted)
        posterior = posterior / posterior.sum(axis=1, keepdims=True)
        risks = posterior @ loss.T
        return self.labels[np.argmin(risks, axis=1)]


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

    def log_scores(self, x: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        scores = []
        for label in self.labels:
            samples = self.samples[str(label)]
            diff = x[:, None, :] - samples[None, :, :]
            exponent = -0.5 * np.sum((diff / self.bandwidth) ** 2, axis=2)
            coef = 1.0 / ((2 * math.pi) ** (x.shape[1] / 2) * self.bandwidth**x.shape[1])
            density = coef * np.exp(exponent).mean(axis=1)
            scores.append(np.log(density + 1e-300) + math.log(self.priors[str(label)]))
        return np.vstack(scores).T

    def predict(self, x: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        return self.labels[np.argmax(self.log_scores(x), axis=1)]

    def predict_risk(self, x: np.ndarray, loss: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        log_scores = self.log_scores(x)
        shifted = log_scores - log_scores.max(axis=1, keepdims=True)
        posterior = np.exp(shifted)
        posterior = posterior / posterior.sum(axis=1, keepdims=True)
        risks = posterior @ loss.T
        return self.labels[np.argmin(risks, axis=1)]


class KNNBayes:
    def __init__(self, k: int = 5) -> None:
        self.k = k
        self.labels: np.ndarray | None = None
        self.train_x: np.ndarray | None = None
        self.train_y: np.ndarray | None = None

    def fit(self, x: np.ndarray, y: np.ndarray) -> "KNNBayes":
        self.labels = np.unique(y)
        self.train_x = x
        self.train_y = y
        return self

    def posterior(self, x: np.ndarray, chunk_size: int = 256) -> np.ndarray:
        assert self.labels is not None and self.train_x is not None and self.train_y is not None
        posts = []
        for start in range(0, len(x), chunk_size):
            query = x[start : start + chunk_size]
            d2 = (
                np.sum(query**2, axis=1, keepdims=True)
                + np.sum(self.train_x**2, axis=1)[None, :]
                - 2 * query @ self.train_x.T
            )
            nearest = np.argpartition(d2, kth=min(self.k, self.train_x.shape[0] - 1), axis=1)[:, : self.k]
            block = np.zeros((len(query), len(self.labels)))
            for row_idx, neighbors in enumerate(nearest):
                neighbor_labels = self.train_y[neighbors]
                for label_idx, label in enumerate(self.labels):
                    block[row_idx, label_idx] = np.mean(neighbor_labels == label)
            posts.append(block)
        return np.vstack(posts)

    def predict(self, x: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        return self.labels[np.argmax(self.posterior(x), axis=1)]

    def predict_risk(self, x: np.ndarray, loss: np.ndarray) -> np.ndarray:
        assert self.labels is not None
        risks = self.posterior(x) @ loss.T
        return self.labels[np.argmin(risks, axis=1)]


def loss_matrix(num_classes: int) -> np.ndarray:
    loss = np.ones((num_classes, num_classes))
    np.fill_diagonal(loss, 0.0)
    loss[:, 0] = 2.0
    loss[0, 0] = 0.0
    return loss


def metric_row(dataset: str, method: str, setting: str, y_true: np.ndarray, y_pred: np.ndarray) -> Dict[str, object]:
    labels = np.unique(y_true)
    return {
        "dataset": dataset,
        "method": method,
        "setting": setting,
        "accuracy": accuracy_score(y_true, y_pred),
        "macro_precision": precision_score(y_true, y_pred, labels=labels, average="macro", zero_division=0),
        "macro_recall": recall_score(y_true, y_pred, labels=labels, average="macro", zero_division=0),
    }


def write_csv(path: Path, rows: List[Dict[str, object]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def pct(value: float) -> str:
    return f"{value * 100:.2f}\\%"


def write_latex_table(rows: List[Dict[str, object]]) -> None:
    order = list(dict.fromkeys(row["dataset"] for row in rows))
    methods = ["参数最小错误率", "参数最小风险", "非参数最小错误率", "非参数最小风险"]
    settings = {
        "模拟二类高斯": "原始二维特征；高斯 full；Parzen $h=0.8$",
        "模拟三类高斯": "原始二维特征；高斯 full；Parzen $h=0.8$",
        "UCI Iris": "4 维花萼/花瓣特征；高斯 full；Parzen $h=0.8$",
        "MNIST": "PCA 64 维；高斯 diag；kNN $k=5$",
        "CIFAR-10": "PCA 64 维；高斯 diag；kNN $k=7$",
        "CIFAR-100": "PCA 64 维；高斯 diag；kNN $k=5$",
    }
    lookup = {(row["dataset"], row["method"]): row for row in rows}
    lines = [
        "\\small",
        "\\begin{tabular}{>{\\raggedright\\arraybackslash}p{2.4cm}>{\\raggedright\\arraybackslash}p{4.2cm}>{\\centering\\arraybackslash}p{1.8cm}>{\\centering\\arraybackslash}p{1.8cm}>{\\centering\\arraybackslash}p{1.8cm}>{\\centering\\arraybackslash}p{1.8cm}}",
        "\\toprule",
        "数据集 & 实验设置 & 参数最小错误率 & 参数最小风险 & 非参数最小错误率 & 非参数最小风险 \\\\",
        "\\midrule",
    ]
    for dataset in order:
        vals = [pct(float(lookup[(dataset, method)]["accuracy"])) for method in methods]
        lines.append(
            f"{dataset} & {settings[dataset]} & {vals[0]} & {vals[1]} & {vals[2]} & {vals[3]} \\\\"
        )
    lines.extend(["\\bottomrule", "\\end{tabular}"])
    (OUT_DIR / "summary_table.tex").write_text("\n".join(lines), encoding="utf-8")


def plot_confusion(cm: np.ndarray, labels: Sequence[str], title: str, path: Path, figsize: Tuple[float, float] = (4.8, 4.2)) -> None:
    fig, ax = plt.subplots(figsize=figsize)
    im = ax.imshow(cm, cmap="Blues")
    ax.set_title(title)
    ax.set_xlabel("Predicted")
    ax.set_ylabel("True")
    ax.set_xticks(range(len(labels)), labels, rotation=45 if len(labels) > 4 else 0, ha="right")
    ax.set_yticks(range(len(labels)), labels)
    if len(labels) <= 12:
        for i in range(cm.shape[0]):
            for j in range(cm.shape[1]):
                ax.text(j, i, str(cm[i, j]), ha="center", va="center", color="black", fontsize=8)
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_accuracy(rows: List[Dict[str, object]]) -> None:
    datasets = list(dict.fromkeys(row["dataset"] for row in rows))
    label_map = {
        "模拟二类高斯": "Sim-2",
        "模拟三类高斯": "Sim-3",
        "UCI Iris": "Iris",
        "MNIST": "MNIST",
        "CIFAR-10": "CIFAR-10",
        "CIFAR-100": "CIFAR-100",
    }
    methods = ["参数最小错误率", "参数最小风险", "非参数最小错误率", "非参数最小风险"]
    method_names = ["Param error", "Param risk", "Nonparam error", "Nonparam risk"]
    x = np.arange(len(datasets))
    width = 0.18
    fig, ax = plt.subplots(figsize=(10, 4.8))
    for idx, method in enumerate(methods):
        vals = []
        for dataset in datasets:
            match = next(row for row in rows if row["dataset"] == dataset and row["method"] == method)
            vals.append(float(match["accuracy"]) * 100)
        ax.bar(x + (idx - 1.5) * width, vals, width, label=method_names[idx])
    ax.set_xticks(x, [label_map[d] for d in datasets], rotation=0)
    ax.set_ylabel("Accuracy (%)")
    ax.set_ylim(0, 105)
    ax.legend(ncol=2, fontsize=9)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "extended_accuracy_comparison.pdf")
    plt.close(fig)


def plot_simulated_boundary(train_x: np.ndarray, train_y: np.ndarray, labels: np.ndarray) -> None:
    model = GaussianBayes("full").fit(train_x, train_y)
    knn = KNNBayes(k=7).fit(train_x, train_y)
    x_min, x_max = train_x[:, 0].min() - 1, train_x[:, 0].max() + 1
    y_min, y_max = train_x[:, 1].min() - 1, train_x[:, 1].max() + 1
    xx, yy = np.meshgrid(np.linspace(x_min, x_max, 220), np.linspace(y_min, y_max, 220))
    grid = np.c_[xx.ravel(), yy.ravel()]
    fig, axes = plt.subplots(1, 2, figsize=(9, 4))
    for ax, clf, title in [(axes[0], model, "Gaussian Bayes"), (axes[1], knn, "kNN Bayes")]:
        pred = clf.predict(grid).reshape(xx.shape)
        pred_idx = np.vectorize({label: idx for idx, label in enumerate(labels)}.get)(pred)
        ax.contourf(xx, yy, pred_idx, levels=np.arange(len(labels) + 1) - 0.5, alpha=0.18)
        for label in labels:
            pts = train_x[train_y == label]
            ax.scatter(pts[:, 0], pts[:, 1], s=14, label=label, alpha=0.85)
        ax.set_title(title)
        ax.set_xlabel("$x_1$")
        ax.set_ylabel("$x_2$")
    axes[1].legend(loc="best", fontsize=8)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "simulated_2class_boundaries.pdf")
    plt.close(fig)


def plot_lowdim_boundary(train_x: np.ndarray, train_y: np.ndarray, labels: np.ndarray, path: Path, title: str) -> None:
    model = GaussianBayes("full").fit(train_x, train_y)
    parzen = ParzenBayes(0.8).fit(train_x, train_y)
    x_min, x_max = train_x[:, 0].min() - 1, train_x[:, 0].max() + 1
    y_min, y_max = train_x[:, 1].min() - 1, train_x[:, 1].max() + 1
    xx, yy = np.meshgrid(np.linspace(x_min, x_max, 240), np.linspace(y_min, y_max, 240))
    grid = np.c_[xx.ravel(), yy.ravel()]
    fig, axes = plt.subplots(1, 2, figsize=(9.5, 4))
    label_to_idx = {label: idx for idx, label in enumerate(labels)}
    for ax, clf, subtitle in [(axes[0], model, "Gaussian Bayes"), (axes[1], parzen, "Parzen Bayes")]:
        pred = clf.predict(grid).reshape(xx.shape)
        pred_idx = np.vectorize(label_to_idx.get)(pred)
        ax.contourf(xx, yy, pred_idx, levels=np.arange(len(labels) + 1) - 0.5, alpha=0.18)
        for label in labels:
            pts = train_x[train_y == label]
            ax.scatter(pts[:, 0], pts[:, 1], s=14, label=label, alpha=0.85)
        ax.set_title(subtitle)
        ax.set_xlabel("$z_1$")
        ax.set_ylabel("$z_2$")
    axes[0].set_ylabel(f"{title}\n$z_2$")
    axes[1].legend(loc="best", fontsize=8)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_pca_prediction_projection(
    test_x: np.ndarray,
    y_true: np.ndarray,
    y_pred: np.ndarray,
    path: Path,
    title: str,
) -> None:
    correct = y_true == y_pred
    fig, ax = plt.subplots(figsize=(5.8, 4.4))
    ax.scatter(test_x[correct, 0], test_x[correct, 1], s=14, alpha=0.65, label="Correct")
    ax.scatter(test_x[~correct, 0], test_x[~correct, 1], s=26, marker="x", color="crimson", label="Wrong")
    ax.set_title(title)
    ax.set_xlabel("PCA-1")
    ax.set_ylabel("PCA-2")
    ax.legend()
    ax.grid(alpha=0.25)
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
    indices = np.arange(min(max_images, len(images)))
    cols = 5
    rows = int(math.ceil(len(indices) / cols))
    fig, axes = plt.subplots(rows, cols, figsize=(8.5, 1.9 * rows))
    axes = np.atleast_1d(axes).ravel()
    for ax, idx in zip(axes, indices):
        img = images[idx]
        if len(image_shape) == 2:
            ax.imshow(img.reshape(image_shape), cmap="gray")
        else:
            ax.imshow(img.reshape(3, 32, 32).transpose(1, 2, 0))
        color = "black" if y_true[idx] == y_pred[idx] else "crimson"
        ax.set_title(f"T:{y_true[idx]}  P:{y_pred[idx]}", fontsize=8, color=color)
        ax.axis("off")
    for ax in axes[len(indices):]:
        ax.axis("off")
    fig.suptitle(title, y=0.995)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def evaluate_dataset(
    name: str,
    train_x: np.ndarray,
    train_y: np.ndarray,
    test_x: np.ndarray,
    test_y: np.ndarray,
    use_pca: bool,
    param_cov: str,
    nonparam: str,
    k: int = 7,
    bandwidth: float = 0.8,
) -> Tuple[List[Dict[str, object]], Dict[str, np.ndarray]]:
    train_x, test_x, prep_note = preprocess(train_x, test_x, use_pca=use_pca)
    labels = np.unique(train_y)
    loss = loss_matrix(len(labels))
    rows = []

    param = GaussianBayes(param_cov).fit(train_x, train_y)
    pred_param = param.predict(test_x)
    pred_param_risk = param.predict_risk(test_x, loss)
    rows.append(metric_row(name, "参数最小错误率", f"{prep_note}；高斯{param_cov}", test_y, pred_param))
    rows.append(metric_row(name, "参数最小风险", f"{prep_note}；高斯{param_cov}；首类漏判损失2", test_y, pred_param_risk))

    if nonparam == "parzen":
        non = ParzenBayes(bandwidth=bandwidth).fit(train_x, train_y)
        non_setting = f"{prep_note}；Parzen窗 h={bandwidth}"
    else:
        non = KNNBayes(k=k).fit(train_x, train_y)
        non_setting = f"{prep_note}；kNN k={k}"
    pred_non = non.predict(test_x)
    pred_non_risk = non.predict_risk(test_x, loss)
    rows.append(metric_row(name, "非参数最小错误率", non_setting, test_y, pred_non))
    rows.append(metric_row(name, "非参数最小风险", f"{non_setting}；首类漏判损失2", test_y, pred_non_risk))

    matrices = {
        "param": confusion_matrix(test_y, pred_param, labels=labels),
        "nonparam": confusion_matrix(test_y, pred_non, labels=labels),
        "labels": labels,
        "test_y": test_y,
        "pred_param": pred_param,
        "pred_non": pred_non,
        "train_x": train_x,
        "train_y": train_y,
        "test_x": test_x,
    }
    return rows, matrices


def main() -> None:
    ensure_dirs()
    all_rows: List[Dict[str, object]] = []

    train_x, train_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "2class" / "train.csv")
    test_x, test_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "2class" / "test.csv")
    rows, matrices = evaluate_dataset("模拟二类高斯", train_x, train_y, test_x, test_y, False, "full", "parzen")
    all_rows.extend(rows)
    labels = matrices["labels"]
    plot_confusion(matrices["param"], labels, "2-class Gaussian Bayes", FIG_DIR / "simulated_2class_param_confusion.pdf")
    plot_confusion(matrices["nonparam"], labels, "2-class Parzen Bayes", FIG_DIR / "simulated_2class_nonparam_confusion.pdf")
    plot_simulated_boundary(*preprocess(train_x, train_x, False)[:1], train_y, labels)  # standardized train data

    train_x, train_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "3class" / "train.csv")
    test_x, test_y = read_csv_dataset(DATASETS / "02_simulated_gaussian" / "3class" / "test.csv")
    rows, matrices = evaluate_dataset("模拟三类高斯", train_x, train_y, test_x, test_y, False, "full", "parzen")
    all_rows.extend(rows)
    plot_lowdim_boundary(matrices["train_x"], matrices["train_y"], matrices["labels"], FIG_DIR / "simulated_3class_boundaries.pdf", "3-class Gaussian")
    plot_confusion(matrices["param"], matrices["labels"], "3-class Gaussian Bayes", FIG_DIR / "simulated_3class_param_confusion.pdf")
    plot_confusion(matrices["nonparam"], matrices["labels"], "3-class Parzen Bayes", FIG_DIR / "simulated_3class_nonparam_confusion.pdf")

    feature_cols = ["sepal_length", "sepal_width", "petal_length", "petal_width"]
    train_x, train_y = read_csv_dataset(DATASETS / "03_uci_iris" / "train.csv", feature_cols)
    test_x, test_y = read_csv_dataset(DATASETS / "03_uci_iris" / "test.csv", feature_cols)
    rows, matrices = evaluate_dataset("UCI Iris", train_x, train_y, test_x, test_y, False, "full", "parzen")
    all_rows.extend(rows)
    iris_train_2d, _ = preprocess_with_pca2(train_x, test_x)
    plot_lowdim_boundary(iris_train_2d, train_y, np.unique(train_y), FIG_DIR / "iris_pca2_boundaries.pdf", "Iris PCA")
    plot_confusion(matrices["param"], matrices["labels"], "Iris Gaussian Bayes", FIG_DIR / "iris_param_confusion.pdf")
    plot_confusion(matrices["nonparam"], matrices["labels"], "Iris Parzen Bayes", FIG_DIR / "iris_nonparam_confusion.pdf")

    mnist_raw = DATASETS / "04_mnist" / "raw"
    train_x = read_mnist_images(mnist_raw / "train-images-idx3-ubyte.gz", limit=12000)
    train_y = read_mnist_labels(mnist_raw / "train-labels-idx1-ubyte.gz", limit=12000)
    test_x = read_mnist_images(mnist_raw / "t10k-images-idx3-ubyte.gz", limit=2000)
    test_y = read_mnist_labels(mnist_raw / "t10k-labels-idx1-ubyte.gz", limit=2000)
    train_x, train_y = stratified_subset(train_x, train_y, per_class=600)
    rows, matrices = evaluate_dataset("MNIST", train_x, train_y, test_x, test_y, True, "diag", "knn", k=5)
    all_rows.extend(rows)
    plot_confusion(matrices["nonparam"], matrices["labels"], "MNIST kNN Bayes", FIG_DIR / "mnist_knn_confusion.pdf", figsize=(6.2, 5.6))
    plot_pca_prediction_projection(matrices["test_x"], matrices["test_y"], matrices["pred_non"], FIG_DIR / "mnist_pca_prediction_projection.pdf", "MNIST kNN prediction on PCA features")
    plot_image_predictions(test_x, test_y, matrices["pred_non"], FIG_DIR / "mnist_prediction_examples.pdf", "MNIST kNN prediction examples", (28, 28))

    train_x, train_y, test_x, test_y = load_cifar10()
    train_x, train_y = stratified_subset(train_x, train_y, per_class=500)
    test_x, test_y = stratified_subset(test_x, test_y, per_class=200)
    rows, matrices = evaluate_dataset("CIFAR-10", train_x, train_y, test_x, test_y, True, "diag", "knn", k=7)
    all_rows.extend(rows)
    plot_confusion(matrices["param"], matrices["labels"], "CIFAR-10 Gaussian Bayes", FIG_DIR / "cifar10_param_confusion.pdf", figsize=(6.2, 5.6))
    plot_pca_prediction_projection(matrices["test_x"], matrices["test_y"], matrices["pred_param"], FIG_DIR / "cifar10_pca_prediction_projection.pdf", "CIFAR-10 Gaussian Bayes on PCA features")
    plot_image_predictions(test_x, test_y, matrices["pred_param"], FIG_DIR / "cifar10_prediction_examples.pdf", "CIFAR-10 Gaussian Bayes prediction examples", (3, 32, 32))

    train_x, train_y, test_x, test_y = load_cifar100()
    train_x, train_y = stratified_subset(train_x, train_y, per_class=80)
    test_x, test_y = stratified_subset(test_x, test_y, per_class=20)
    rows, matrices = evaluate_dataset("CIFAR-100", train_x, train_y, test_x, test_y, True, "diag", "knn", k=5)
    all_rows.extend(rows)
    plot_pca_prediction_projection(matrices["test_x"], matrices["test_y"], matrices["pred_param"], FIG_DIR / "cifar100_pca_prediction_projection.pdf", "CIFAR-100 Gaussian Bayes on PCA features")
    plot_image_predictions(test_x, test_y, matrices["pred_param"], FIG_DIR / "cifar100_prediction_examples.pdf", "CIFAR-100 Gaussian Bayes prediction examples", (3, 32, 32))

    write_csv(OUT_DIR / "metrics.csv", all_rows)
    write_latex_table(all_rows)
    plot_accuracy(all_rows)


if __name__ == "__main__":
    main()
